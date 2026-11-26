"""Persistent, typed DuckDB store for complete daily snapshots."""

from __future__ import annotations

import csv
import gzip
import hashlib
import json
import sqlite3
import tempfile
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from contextlib import closing, contextmanager
from datetime import date, datetime
from decimal import Decimal, localcontext
from itertools import chain, groupby
from pathlib import Path

import duckdb
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.worksheet._reader import FORMULA_TAG, WorkSheetParser

import core
from core import ExportCancelled, field_choices, json_value, normalize_field, parse_date_parts, sheet_columns

# All five members are native, compressed columns. Mixed Excel columns can
# retain text codes, booleans and numbers without guessing from the first rows.
CELL = "UNION(s VARCHAR, i BIGINT, m DECIMAL(38,12), n DOUBLE, b BOOLEAN)"
FORMAT = "excel-customer-splitter/1"
DDL = """
CREATE TABLE library_format(version VARCHAR);
CREATE TABLE batches(id VARCHAR, committed_at TIMESTAMP, input_rows BIGINT, replaced_rows BIGINT, overlap_days INTEGER);
CREATE TABLE columns_catalog(id INTEGER PRIMARY KEY, field_key VARCHAR UNIQUE, display_name VARCHAR);
CREATE TABLE sources(id BIGINT PRIMARY KEY, batch VARCHAR, path VARCHAR, sheet VARCHAR, fingerprint VARCHAR, input_rows BIGINT);
CREATE TABLE details(source_id BIGINT, source_row BIGINT, business_date DATE);
CREATE TABLE calendar(day DATE PRIMARY KEY, row_count BIGINT);
CREATE TABLE migrated_tasks(task VARCHAR PRIMARY KEY);
"""


def cancelled(event):
    if event is not None and event.is_set():
        raise ExportCancelled("任务已停止")


def monetary(name):
    return any(word in name for word in ("金额", "流水", "价格", "费用", "余额", "单价", "代金券"))


def wire(value, money):
    value = json_value(value)
    if value is None:
        return ""
    if isinstance(value, bool):
        return "b" + str(value)
    if isinstance(value, int) and -(2**63) <= value < 2**63:
        return "i" + str(value)
    if isinstance(value, (float, int)):
        number = Decimal(str(value))
        if money and number.is_finite() and number.as_tuple().exponent >= -12 and number.adjusted() < 26:
            return "m" + str(value)
        if isinstance(value, int):
            raise ValueError(f"超出可无损存储范围的整数：{value}")
        return "n" + repr(value)
    return "s" + str(value)


def decode_sql(col):
    members = (("s", "VARCHAR"), ("i", "BIGINT"), ("m", "DECIMAL(38,12)"), ("n", "DOUBLE"), ("b", "BOOLEAN"))
    cases = " ".join(
        f"WHEN '{tag}' THEN union_value({tag} := substr({col},2)::{kind})::{CELL}" for tag, kind in members
    )
    return f"CASE left({col},1) {cases} END"


class CachedValueParser(WorkSheetParser):
    """Read a cached result and its formula marker in the same XML pass."""

    def parse_cell(self, element):
        cell = super().parse_cell(element)
        if cell["value"] is None and element.find(FORMULA_TAG) is not None:
            if element.get("t") == "str":
                cell["value"] = ""
            else:
                raise ValueError(
                    f"第 {cell['row']} 行第 {cell['column']} 列的公式缺少缓存结果，请用 Excel 重算并保存。"
                )
        return cell


def worksheet_rows(ws, xml):
    wb = ws.parent
    parser = CachedValueParser(
        xml,
        ws._shared_strings,
        data_only=True,
        epoch=wb.epoch,
        date_formats=wb._date_formats,
        timedelta_formats=wb._timedelta_formats,
    )
    for pos, cells in parser.parse():
        values = [None] * max((c["column"] for c in cells), default=0)
        for cell in cells:
            values[cell["column"] - 1] = cell["value"]
        yield pos, values


def excel_cells(sheet, values):
    for value in values:
        if isinstance(value, str) and value.startswith("="):
            cell = WriteOnlyCell(sheet, value)
            cell.data_type = "s"
            yield cell
        else:
            yield value


@contextmanager
def interruptible(event):
    try:
        yield
    except duckdb.InterruptException:
        cancelled(event)
        raise


def spool_rows(path, columns, rows, event, audit=False):
    keys = list(columns)
    date_positions = [keys.index(k) if k in columns else None for k in ("年", "月", "日")]
    dated = None not in date_positions
    money = [monetary(name) for name in columns.values()]
    total = 0
    audit_totals = {}
    with gzip.open(path, "wt", newline="", encoding="utf-8", compresslevel=1) as file:
        writer = csv.writer(file)
        for row_number, values in rows:
            if total % 1024 == 0:
                cancelled(event)
            values = list(values)
            if len(values) > len(keys):
                expanded = sheet_columns([*columns.values(), *([None] * (len(values) - len(keys)))])
                columns.update(expanded)
                keys = list(columns)
                money = [monetary(name) for name in columns.values()]
            values += [None] * (len(keys) - len(values))
            if not any(v is not None and str(v).strip() for v in values):
                continue
            day = parse_date_parts(*(values[i] for i in date_positions)) if dated else None
            cells = [wire(v, m) for v, m in zip(values, money)]
            writer.writerow([row_number, day, *cells])
            if audit:
                month = day.strftime("%Y-%m") if day else ""
                count, sums = audit_totals.setdefault(month, [0, [Decimal(0) for _ in keys]])
                audit_totals[month][0] = count + 1
                for i, token in enumerate(cells):
                    if money[i] and token[:1] in ("i", "m", "n"):
                        sums[i] += Decimal(token[1:])
            total += 1
    return total, audit_totals


def spool_excel(item, workdir, event):
    index, path = item
    output = Path(workdir) / f"source-{index}.csv.gz"
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            cancelled(event)
            digest.update(block)
        handle.seek(0)
        wb = load_workbook(handle, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            # Read physical cells so incorrect worksheet dimensions or unlabelled
            # trailing columns cannot truncate an otherwise valid snapshot.
            with ws._get_source() as xml:
                rows = worksheet_rows(ws, xml)
                prefix = []
                header = 1
                labels = []
                for pos, values in rows:
                    if pos == 1:
                        labels = values
                    if pos <= 30 and sum(v is not None and bool(str(v).strip()) for v in values) >= 2:
                        header, labels = pos, values
                        break
                    prefix.append((pos, values))
                    if pos >= 30:
                        break
                columns = sheet_columns(labels)
                remaining = chain(((p, v) for p, v in prefix if p > header), rows)
                total, _ = spool_rows(output, columns, remaining, event)
            return output, str(path), ws.title, digest.hexdigest(), columns, total
        except ValueError as exc:
            raise ValueError(f"{path}：{exc}") from exc
        finally:
            wb.close()


class Warehouse:
    def __init__(self, path):
        self.path = Path(path).resolve()
        self.lock = threading.RLock()
        new = not self.path.exists()
        if not new:
            with self.path.open("rb") as file:
                header = file.read(16)
            if header.startswith(b"SQLite format 3"):
                raise ValueError("这是旧 SQLite 数据库，请使用“迁移旧 SQLite”。")
            if header[8:12] != b"DUCK":
                raise ValueError("无法识别客户数据库文件。")
        self.db = duckdb.connect(str(self.path))
        try:
            self.db.execute("SET threads=4")
            self.db.execute("SET preserve_insertion_order=false")
            if new:
                with self.transaction():
                    self.db.execute(DDL)
                    self.db.execute("INSERT INTO library_format VALUES (?)", [FORMAT])
            if self.db.execute("SELECT version FROM library_format").fetchall() != [(FORMAT,)]:
                raise ValueError("不支持的客户数据库版本")
        except Exception:
            self.db.close()
            raise

    @contextmanager
    def transaction(self):
        self.db.execute("BEGIN")
        try:
            yield
            self.db.execute("COMMIT")
        except BaseException:
            self.db.execute("ROLLBACK")
            raise

    def interrupt(self):
        # DuckDB supports interrupt() from a different thread while execute() is active.
        self.db.interrupt()

    def close(self):
        with self.lock:
            self.db.close()

    def checkpoint(self):
        with self.lock:
            self.db.execute("CHECKPOINT")

    def fields(self):
        cols = self.db.execute("SELECT field_key, display_name FROM columns_catalog ORDER BY id").fetchall()
        return field_choices(dict(cols))

    def restore(self):
        with self.lock:
            last = self.db.execute("SELECT id FROM batches ORDER BY rowid DESC LIMIT 1").fetchone()
            return (last[0] if last else None), list(self.fields())

    def dates(self):
        with self.lock:
            return [
                str(day) for (day,) in self.db.execute("SELECT day FROM calendar ORDER BY day").fetchall()
            ]

    def start_batch(self):
        self.db.execute("CREATE TEMP TABLE incoming AS SELECT * FROM details LIMIT 0")
        return uuid.uuid4().hex

    def load_file(self, batch, file):
        csv_path, source, sheet, fingerprint, columns, total = file
        catalog = dict(self.db.execute("SELECT field_key,id FROM columns_catalog").fetchall())
        physical = []
        for key, display in columns.items():
            if key not in catalog:
                i = len(catalog)
                self.db.execute("INSERT INTO columns_catalog VALUES (?,?,?)", [i, key, display])
                for table in ("details", "incoming"):
                    self.db.execute(f"ALTER TABLE {table} ADD COLUMN f{i} {CELL}")
                catalog[key] = i
            physical.append(f"f{catalog[key]}")
        fid = self.db.execute("SELECT coalesce(max(id),0)+1 FROM sources").fetchone()[0]
        self.db.execute(
            "INSERT INTO sources VALUES (?,?,?,?,?,?)", [fid, batch, source, sheet, fingerprint, total]
        )
        if total:
            csv_schema = ",".join(
                ["'rownum':'BIGINT'", "'day':'DATE'"] + [f"'{c}':'VARCHAR'" for c in physical]
            )
            names = ",".join(["source_id", "source_row", "business_date", *physical])
            projection = ",".join(["?", "rownum", "day", *(decode_sql(c) for c in physical)])
            count = self.db.execute(
                f"""
                INSERT INTO incoming ({names}) SELECT {projection}
                FROM read_csv(?,columns={{{csv_schema}}},header=false,auto_detect=false,
                              delim=',',quote='"',escape='"',nullstr='',null_padding=true,parallel=false,max_line_size=134217728)
            """,
                [fid, str(csv_path)],
            ).fetchone()[0]
            if count != total:
                raise ValueError(f"{source} 入库行数不一致：{total} / {count}")
        return fid, physical

    def publish(self, batch, event):
        cancelled(event)
        self.db.execute(
            "CREATE TEMP TABLE batch_days AS SELECT business_date AS day,count(*) AS n FROM incoming WHERE business_date IS NOT NULL GROUP BY 1"
        )
        overlap, replaced = self.db.execute(
            "SELECT count(*),coalesce(sum(row_count),0) FROM calendar JOIN batch_days USING(day)"
        ).fetchone()
        if overlap:
            deleted = self.db.execute(
                "DELETE FROM details WHERE business_date IN (SELECT day FROM batch_days)"
            ).fetchone()[0]
            if deleted != replaced:
                raise ValueError(f"覆盖前旧明细与日历行数不符：{deleted} / {replaced}")
        # Undated rows cannot identify a snapshot. Replay of the same file
        # replaces its own undated rows while unrelated files remain present.
        undated_replaced = self.db.execute(
            """
            DELETE FROM details WHERE business_date IS NULL AND source_id IN
            (SELECT old.id FROM sources old JOIN sources fresh USING(fingerprint)
             WHERE fresh.batch=? AND old.batch<>?)
        """,
            [batch, batch],
        ).fetchone()[0]
        self.db.execute("INSERT OR REPLACE INTO calendar SELECT day,n FROM batch_days")
        inserted = self.db.execute(
            "INSERT INTO details SELECT * FROM incoming ORDER BY business_date,source_id,source_row"
        ).fetchone()[0]
        total = self.db.execute(
            "SELECT coalesce(sum(input_rows),0) FROM sources WHERE batch=?", [batch]
        ).fetchone()[0]
        if inserted != total:
            raise ValueError(f"批次行数不一致：{inserted} / {total}")
        self.db.execute(
            "INSERT INTO batches VALUES (?,current_timestamp,?,?,?)", [batch, total, replaced, overlap]
        )
        self.db.execute("DROP TABLE incoming")
        dated = self.db.execute("SELECT coalesce(sum(n),0) FROM batch_days").fetchone()[0]
        self.db.execute("DROP TABLE batch_days")
        self.refresh_analysis_view()
        cancelled(event)
        return (
            f"导入 {total:,} 行，覆盖 {overlap} 个已有日期的 {replaced:,} 条旧明细。"
            f"本批无有效日期 {total - dated:,} 行，替换同一文件的无日期旧明细 {undated_replaced:,} 行。"
        )

    def refresh_analysis_view(self):
        catalog = dict(self.db.execute("SELECT field_key,id FROM columns_catalog").fetchall())
        projection = []
        for label, (_, key) in self.fields().items():
            alias = label.replace('"', '""')
            projection.append(f'f{catalog[key]} AS "{alias}"')
        if projection:
            self.db.execute(
                "CREATE OR REPLACE VIEW analysis_details AS SELECT " + ", ".join(projection) + " FROM details"
            )

    def import_files(self, paths, progress=None, cancel_event=None):
        if not paths:
            raise ValueError("请选择 Excel 文件")
        with (
            self.lock,
            interruptible(cancel_event),
            tempfile.TemporaryDirectory(prefix=".split-import-", dir=self.path.parent) as tmp,
        ):
            with self.transaction(), ThreadPoolExecutor(max_workers=min(4, len(paths))) as pool:
                batch = self.start_batch()
                for file in pool.map(lambda item: spool_excel(item, tmp, cancel_event), enumerate(paths)):
                    cancelled(cancel_event)
                    self.load_file(batch, file)
                    file[0].unlink()
                    if progress:
                        progress(f"已读取 {Path(file[1]).name}，{file[-1]:,} 行")
                message = self.publish(batch, cancel_event)
            if progress:
                progress(message)
            return batch, list(self.fields())

    def migrate_sqlite(self, path, progress=None, cancel_event=None):
        uri = Path(path).resolve().as_uri() + "?mode=ro"
        with (
            self.lock,
            interruptible(cancel_event),
            closing(sqlite3.connect(uri, uri=True)) as old,
            self.transaction(),
            tempfile.TemporaryDirectory(dir=self.path.parent) as tmp,
            localcontext() as ctx,
        ):
            ctx.prec = 60
            old.execute("BEGIN")
            tasks = old.execute(
                "SELECT task_id FROM import_task WHERE status='READY' ORDER BY created_at,rowid"
            ).fetchall()
            migrated = {t for (t,) in self.db.execute("SELECT task FROM migrated_tasks").fetchall()}
            pending = [t for (t,) in tasks if t not in migrated]
            if pending and self.db.execute("SELECT count(*) FROM batches").fetchone()[0]:
                raise ValueError("请将旧 SQLite 迁入新建空库，以免旧快照覆盖已有日期。")
            for task in pending:
                cancelled(cancel_event)
                batch = self.start_batch()
                columns = dict(
                    old.execute(
                        "SELECT normalized_name,display_name FROM source_column WHERE task_id=? ORDER BY rowid",
                        (task,),
                    )
                )
                keys = list(columns)
                # Legacy duplicate keys used physical positions; new Excel imports use occurrence.
                converted = sheet_columns(columns.values())
                files = {
                    fid: (p, s)
                    for fid, p, s in old.execute(
                        "SELECT file_id,file_path,sheet_name FROM source_file WHERE task_id=?", (task,)
                    )
                }
                rows = old.execute(
                    "SELECT file_id,source_row,raw_json FROM staging_row WHERE task_id=? ORDER BY file_id,source_row,row_id",
                    (task,),
                )
                verified = 0
                for fid, group in groupby(rows, lambda row: row[0]):
                    source, sheet = files[fid]
                    spool = Path(tmp) / "migration.csv.gz"

                    def values():
                        for _, pos, raw in group:
                            row = json.loads(raw)
                            unknown = row.keys() - columns.keys()
                            if unknown:
                                raise ValueError(f"{source} 第 {pos} 行包含旧字段目录中不存在的列：{sorted(unknown)}")
                            yield pos, [row.get(k) for k in keys]

                    total, audit = spool_rows(spool, converted, values(), cancel_event, audit=True)
                    # Old caches do not contain an Excel file hash. Keep undated
                    # rows from each legacy source independently; migrated_tasks guards replays.
                    fingerprint = f"sqlite:{task}:{fid}"
                    file_id, physical = self.load_file(
                        batch, (spool, source, sheet, fingerprint, converted, total)
                    )
                    metrics = [
                        (i, c)
                        for i, (c, name) in enumerate(zip(physical, converted.values()))
                        if monetary(name)
                    ]
                    sums = [
                        f"coalesce(sum(coalesce(union_extract({c},'m'),union_extract({c},'i')::DECIMAL(38,12),try_cast(union_extract({c},'n') as DECIMAL(38,12)))),0)"
                        for _, c in metrics
                    ]
                    select = ",".join(["coalesce(strftime(business_date,'%Y-%m'),'')", "count(*)", *sums])
                    actual = self.db.execute(
                        f"SELECT {select} FROM incoming WHERE source_id=? GROUP BY 1", [file_id]
                    ).fetchall()
                    for month, count, *amounts in actual:
                        expected_n, expected_sums = audit.pop(month)
                        if count != expected_n or amounts != [expected_sums[i] for i, _ in metrics]:
                            raise ValueError(f"旧库对账失败：{source} {month}")
                    if audit:
                        raise ValueError(f"旧库存在未迁移月份：{source}")
                    verified += total
                original = old.execute(
                    "SELECT count(*) FROM staging_row WHERE task_id=?", (task,)
                ).fetchone()[0]
                if verified != original:
                    raise ValueError(f"旧批次 {task} 对账失败：{original} / {verified}")
                message = self.publish(batch, cancel_event)
                self.db.execute("INSERT INTO migrated_tasks VALUES (?)", [task])
                if progress:
                    progress(f"旧批次 {task} 行数与分月金额核对通过。{message}")
            return self.restore()

    def export_query(self, selected, customer, region, rules, start, end, names, today):
        columns = dict(self.db.execute("SELECT field_key,id FROM columns_catalog").fetchall())
        choices = self.fields()

        def column(label):
            return f"f{columns[choices[label][1]]}"

        # Use the same whitespace categories for customer selection and region matching.
        self.db.execute("CREATE OR REPLACE TEMP TABLE region_map(k VARCHAR, region VARCHAR)")
        mapping = {}
        for rule in rules:
            for value in rule.get("matchValues", []):
                mapping.setdefault(
                    normalize_field(value), str(rule.get("name") or rule.get("code") or "未分区")
                )
        if mapping:
            self.db.executemany("INSERT INTO region_map VALUES (?,?)", list(mapping.items()))

        def text_value(label):
            c = column(label)
            return f"""CASE union_tag({c}) WHEN 's' THEN union_extract({c},'s')
                WHEN 'b' THEN CASE WHEN union_extract({c},'b') THEN 'True' ELSE 'False' END
                WHEN 'm' THEN union_extract({c},'m')::DOUBLE::VARCHAR
                ELSE {c}::VARCHAR END"""

        spaces = r"[\s\p{Z}\x1c-\x1f\x{0085}]"

        def normalized(expr):
            return f"lower(regexp_replace({expr},'{spaces}+','','g'))"

        c = column(customer)
        cust_text = f"""CASE WHEN union_extract({c},'i')=0 OR union_extract({c},'n')=0
                        OR union_extract({c},'m')=0 OR union_extract({c},'b')=false THEN NULL
                        ELSE {text_value(customer)} END"""
        name_expr = (
            f"coalesce(nullif(regexp_replace({cust_text},'^{spaces}+|{spaces}+$','','g'),''),'未识别客户')"
        )
        region_expr = normalized(f"coalesce({text_value(region)},'')") if region else "NULL::VARCHAR"
        params = []
        filters = []
        if start:
            first, last = sorted([date.fromisoformat(start), date.fromisoformat(end or start)])
            filters.append("business_date BETWEEN ? AND ?")
            params.extend([first, last])
            folder = first.isoformat() if first == last else f"{first}至{last}"
        else:
            folder = None
        if names:
            filters.append(normalized(name_expr) + " IN (SELECT unnest(?))")
            params.append([normalize_field(n) for n in names])
        where = " WHERE " + " AND ".join(filters) if filters else ""
        projection = ", ".join(column(label) for label in selected)
        query = f"""
            SELECT {("?" if folder else "coalesce(business_date::VARCHAR,?)")} AS folder,
                   coalesce(mapping.region,'未分区') AS area, {name_expr} AS customer_name, {projection}
            FROM details LEFT JOIN region_map mapping ON {region_expr}=mapping.k
            {where} ORDER BY folder,area,customer_name,source_id,source_row
        """
        return query, [folder or today, *params], [choices[label][0] for label in selected]

    def export(self, output, selected, customer, region, rules, start, end, names, progress=None, event=None):
        with self.lock, interruptible(event):
            cancelled(event)
            today = date.today().isoformat()
            query, params, headers = self.export_query(
                selected, customer, region, rules, start, end, names, today
            )
            cursor = self.db.execute(query, params)

            def records():
                while batch := cursor.fetchmany(8192):
                    cancelled(event)
                    yield from batch

            root = Path(output)
            root.mkdir(parents=True, exist_ok=True)
            total = files = groups = 0
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            manifest = root / f"导出清单-{stamp}-{uuid.uuid4().hex[:8]}.csv"
            pending = manifest.with_suffix(".tmp")
            workbook = sheet = report = None
            try:
                report = pending.open("x", newline="", encoding="utf-8-sig")
                writer = csv.writer(report)
                writer.writerow(["相对路径", "明细行数"])
                for key, rows in groupby(records(), lambda row: row[:3]):
                    groups += 1
                    folder, area, _ = key
                    directory = root / core.safe_filename(folder, today) / core.safe_filename(area, "未分区")
                    directory.mkdir(parents=True, exist_ok=True)
                    part = 0
                    count = 0

                    def save():
                        nonlocal files
                        cancelled(event)
                        path = core.save_group_workbook(workbook, directory, key, part)
                        writer.writerow([str(path.relative_to(root)), count])
                        files += 1
                        if progress:
                            progress(f"已生成 {path}")

                    for row in rows:
                        if count == core.MAX_DATA_ROWS:
                            save()
                            workbook = None
                            count = 0
                        if not count:
                            part += 1
                            workbook = Workbook(write_only=True)
                            sheet = workbook.create_sheet("数据")
                            sheet.freeze_panes = "A2"
                            sheet.append(excel_cells(sheet, headers))
                        sheet.append(excel_cells(sheet, row[3:]))
                        count += 1
                        total += 1
                    save()
                    workbook = None
                cancelled(event)
                report.close()
                pending.replace(manifest)
            finally:
                if report is not None:
                    report.close()
                pending.unlink(missing_ok=True)
                if workbook is not None:
                    if sheet is not None and not sheet.closed:
                        sheet.close()
                    writer = getattr(sheet, "_writer", None)
                    if writer and Path(writer.out).exists():
                        writer.cleanup()
                    workbook.close()
            return {"rows": total, "files": files, "groups": groups, "manifest": manifest}
