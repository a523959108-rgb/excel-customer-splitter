"""Reproducible DuckDB storage/read benchmark using the production schema and export."""

import argparse
import csv
import json
import platform
import sys
import tempfile
import time
from pathlib import Path

import duckdb
from openpyxl import Workbook, load_workbook

from warehouse import CELL, Warehouse

FIELDS = ["年", "月", "日", "客户", "金额", "单号"]


def run(n, directory, *, query_probe=None):
    with tempfile.TemporaryDirectory(dir=directory) as tmp:
        root = Path(tmp)
        store = Warehouse(root / "bench.duckdb")
        db = store.db
        result = {"rows": n, "platform": platform.platform(), "duckdb": duckdb.__version__}

        def measured(key, action):
            start = time.perf_counter()
            value = action()
            result[key + "_seconds"] = round(time.perf_counter() - start, 4)
            return value

        def generate():
            with store.transaction():
                for i, field in enumerate(FIELDS):
                    db.execute("INSERT INTO columns_catalog VALUES (?,?,?)", [i, field, field])
                    db.execute(f"ALTER TABLE details ADD COLUMN f{i} {CELL}")
            # A file covers a day. Commit in batches to bound transaction/spill size.
            for day in range(0, 360, 15):
                lower = n * day // 360
                upper = n * (day + 15) // 360
                with store.transaction():
                    batch = store.start_batch()
                    db.execute(
                        """
                        INSERT INTO sources SELECT d+1, ?, 'synthetic', '数据', 'day-'||d, hi-lo
                        FROM (SELECT d, ?::BIGINT*d//360 lo, ?::BIGINT*(d+1)//360 hi
                              FROM range(?,?) r(d))
                    """,
                        [batch, n, n, day, day + 15],
                    )
                    db.execute(
                        """
                        INSERT INTO incoming
                        SELECT d+1, i-(?::BIGINT*d//360)+2, dt,
                               union_value(i:=year(dt)), union_value(i:=month(dt)), union_value(i:=day(dt)),
                               union_value(s:='客户'||(i%1000)),
                               union_value(m:=((i%99991)::BIGINT * 0.01)::DECIMAL(38,12)),
                               union_value(i:=i)
                        FROM (SELECT i,d,DATE '2025-01-01' + d::INT dt
                              FROM (SELECT i, ((i+1)*360-1)//? AS d FROM range(?,?) r(i)))
                    """,
                        [n, n, lower, upper],
                    )
                    store.publish(batch, None)
            store.checkpoint()

        try:
            measured("generate_publish", generate)
            result["database_bytes"] = store.path.stat().st_size
            dates = measured("calendar", store.dates)
            count, checksum = measured(
                "full_scan",
                lambda: db.execute("SELECT count(*), sum(union_extract(f5,'i')) FROM details").fetchone(),
            )
            if (len(dates), count, checksum) != (360, n, n * (n - 1) // 2):
                raise ValueError("Full library checksum mismatch")
            first, end = "2025-02-01", "2025-02-28"
            names = ["客户17"]
            args = (["客户", "金额", "单号"], "客户", None, [], first, end, names)
            sql, params, _ = store.export_query(*args, first)
            rows = measured("month_customer_read", lambda: db.execute(sql, params).fetchall())
            # February is day 31 to 58 inclusive. Compute expected IDs independently.
            lower, upper = n * 31 // 360, n * 59 // 360
            expected = list(range(lower + (17 - lower) % 1000, upper, 1000))
            if sorted(r[-1] for r in rows) != expected:
                raise ValueError("Month/customer query mismatch")
            result["month_customer_rows"] = len(expected)
            measured("xlsx_export", lambda: store.export(root / "out", *args))

            def read_back():
                (path,) = (root / "out").rglob("*.xlsx")
                wb = load_workbook(path, read_only=True, data_only=True)
                try:
                    data = wb.active.values
                    assert next(data) == ("客户", "金额", "单号")
                    actual = sorted((i, name, round(amount * 100)) for name, amount, i in data)
                    if actual != [(i, "客户17", i % 99991) for i in expected]:
                        raise ValueError("XLSX contents mismatch")
                finally:
                    wb.close()

            measured("xlsx_verify", read_back)
            if query_probe:
                result["query_comparison"] = query_probe(store, sql, params, rows)
            return result
        finally:
            store.close()


def run_excel(n, directory, progress=None, *, import_probe=None):
    """Generate real input files, use the normal import/export, verify every cell."""
    if n <= 0:
        raise ValueError("rows must be positive")
    with tempfile.TemporaryDirectory(dir=directory) as tmp:
        root = Path(tmp)
        store = Warehouse(root / "excel.duckdb")
        result = {"xlsx_input_rows": n, "platform": platform.platform(), "duckdb": duckdb.__version__}

        def record(i):
            return (2025, 2, 1 + i % 28, f"客户{i % 3}", (i % 99991) / 100, i)

        try:
            files = []
            start = time.perf_counter()
            for offset in range(0, n, 500000):
                wb = Workbook(write_only=True)
                ws = wb.create_sheet("数据")
                ws.append(FIELDS)
                for i in range(offset, min(n, offset + 500000)):
                    ws.append(record(i))
                path = root / f"input-{offset}.xlsx"
                wb.save(path)
                files.append(path)
                if progress:
                    progress(f"Prepared {len(files)} XLSX files ({min(n, offset + 500000):,} rows)")
            result["fixture_seconds"] = round(time.perf_counter() - start, 3)
            result["input_files"] = len(files)
            result["input_bytes"] = sum(path.stat().st_size for path in files)
            start = time.perf_counter()
            _, fields = store.import_files(files, progress=progress)
            store.checkpoint()
            result["import_seconds"] = round(time.perf_counter() - start, 3)
            result["database_bytes"] = store.path.stat().st_size
            start = time.perf_counter()
            stats = store.export(root / "out", fields, "客户", None, [], "2025-02-01", "2025-02-28", [])
            result["export_seconds"] = round(time.perf_counter() - start, 3)
            result["files"] = stats["files"]
            if progress:
                progress(f"Exported {stats['rows']:,} rows in {stats['files']} files")
            start = time.perf_counter()
            with stats["manifest"].open(encoding="utf-8-sig", newline="") as manifest:
                manifest_rows = list(csv.DictReader(manifest))
            listed = {row["相对路径"]: int(row["明细行数"]) for row in manifest_rows}
            if len(listed) != len(manifest_rows) or len(listed) != stats["files"]:
                raise ValueError("Duplicate or missing manifest entries")
            counts = {}
            output_bytes = 0
            seen = bytearray(n)
            for path in (root / "out").rglob("*.xlsx"):
                count = 0
                customer = path.stem.split("_part")[0]
                wb = load_workbook(path, read_only=True, data_only=True)
                try:
                    rows = wb.active.values
                    if tuple(FIELDS) != next(rows):
                        raise ValueError("XLSX header mismatch")
                    for row in rows:
                        i = row[-1]
                        if not 0 <= i < n or seen[i] or row != record(i) or row[3] != customer:
                            raise ValueError(f"XLSX row mismatch at {i}")
                        seen[i] = 1
                        count += 1
                finally:
                    wb.close()
                counts[str(path.relative_to(root / "out"))] = count
                output_bytes += path.stat().st_size
                if progress:
                    progress(f"Verified {path.name} ({count:,} rows)")
            if counts != listed:
                raise ValueError("Manifest and workbook row counts differ")
            if sum(seen) != n or stats["rows"] != n:
                raise ValueError("Missing XLSX rows")
            result["readback_seconds"] = round(time.perf_counter() - start, 3)
            result["output_bytes"] = output_bytes
        finally:
            store.close()
        if import_probe:
            result["import_comparison"] = import_probe(root, files, n)
        return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows", type=int, nargs="+", default=None)
    parser.add_argument("--directory", type=Path, help="temporary database/output parent directory")
    parser.add_argument(
        "--xlsx-rows", type=int, help="generate, import, export and verify this many real Excel rows"
    )
    parser.add_argument(
        "--compare-imports", action="store_true",
        help="also import the same Excel files using four processes, then reuse their parsed cache",
    )
    args = parser.parse_args()
    if args.compare_imports and args.xlsx_rows is None:
        parser.error("--compare-imports requires --xlsx-rows")
    import_probe = None
    if args.compare_imports:
        from benchmark_import import compare_same_xlsx

        import_probe = compare_same_xlsx
    if args.xlsx_rows is not None:
        result = run_excel(
            args.xlsx_rows,
            args.directory,
            lambda msg: print(msg, file=sys.stderr, flush=True),
            import_probe=import_probe,
        )
        print(json.dumps(result, ensure_ascii=False), flush=True)
    for size in args.rows or ([] if args.xlsx_rows else [10_000_000]):
        print(json.dumps(run(size, args.directory), ensure_ascii=False), flush=True)
