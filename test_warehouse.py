import csv
import json
import threading
import unittest
from collections import Counter
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from zipfile import BadZipFile

from openpyxl import Workbook

from core import ExportCancelled, ImportStore, load_region_rules
from test_integrity import read_excel, write_excel
from warehouse import Warehouse

HEADER = ["年", "月", "日", "客户", "金额", "编码"]


class WarehouseTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.path = self.root / "split.duckdb"
        self.store = Warehouse(self.path)
        self.addCleanup(lambda: self.store.close())

    def file(self, name, rows, header=HEADER):
        return write_excel(self.root / name, [header, *rows])

    def export(self, output="out", fields=None, start=None, end=None, names=None):
        return self.store.export(self.root / output, fields or HEADER, "客户", None, [], start, end, names)

    def output_rows(self, output="out"):
        return [r for p in (self.root / output).rglob("*.xlsx") for r in read_excel(p)[1:]]

    def test_multiline_first_record_round_trip(self):
        rows = [(2026, 2, 1, "甲", 2, '备注,"原文"\n下一行'), (2026, 2, 1, "乙", 3, "单行")]
        self.store.import_files([self.file("多行备注.xlsx", rows)])
        self.assertEqual(self.export()["rows"], len(rows))
        self.assertCountEqual(self.output_rows(), rows)

    def test_two_overlapping_snapshots_and_reopen(self):
        first = [(2025, 1, i, "甲", i, "001") for i in range(1, 13)]
        first += [(2025, 1, 11, "撤销", 999, "009")]
        self.store.import_files([self.file("旧.xlsx", first)])
        new = [(2025, 1, i, "甲", i + 100.25, "001") for i in range(11, 21)]
        supplements = [new[0], (2025, 1, 12, "乙", -0.16, "000")]
        files = [self.file("新.xlsx", new), self.file("同日.xlsx", supplements)]
        self.store.import_files(files)
        state = self.store.import_files(files)
        self.store.close()
        self.store = Warehouse(self.path)
        self.assertEqual(self.store.restore(), state)
        self.assertEqual(self.store.dates(), [f"2025-01-{i:02}" for i in range(1, 21)])
        self.assertEqual(self.export()["rows"], 22)
        self.assertCountEqual(self.output_rows(), first[:10] + new + supplements)
        self.export("selected", start="2025-01-11", end="2025-01-12", names=[" 甲 "])
        self.assertCountEqual(self.output_rows("selected"), [new[0], new[0], new[1]])
        with patch("core.MAX_DATA_ROWS", 1):
            stats = self.export("parts", start="2025-01-11", end="2025-01-11")
        self.assertEqual(stats["files"], 2)
        self.assertCountEqual(self.output_rows("parts"), [new[0], new[0]])

    def test_missing_dates_undated_and_changed_schema(self):
        old = [(2025, 2, day, "甲", day, "00") for day in (1, 2, 3)]
        undated = (2025, 2, 31, "旧", 7, "x")
        a = self.file("a.xlsx", [*old, undated])
        self.store.import_files([a])
        new = [(2025, 2, i, "甲", 1, "00", True) for i in (1, 3)]
        b = self.file("b.xlsx", [*new, (*undated[:-3], "新", 8, "y", False)], HEADER + ["新字段"])
        self.store.import_files([b])
        self.store.import_files([b])
        self.export()
        self.assertCountEqual(
            self.output_rows(), [old[1], undated, *(r[:-1] for r in new), (2025, 2, 31, "新", 8, "y")]
        )

    def test_failure_and_cancel_after_first_file_or_before_commit(self):
        original = [(2025, 1, 1, "甲", 1, "A")]
        self.store.import_files([self.file("old.xlsx", original)])
        before = self.store.restore()
        new = self.file("new.xlsx", [(2025, 1, 1, "乙", 2, "B", 3)], HEADER + ["额外"])
        bad = self.root / "坏.xlsx"
        bad.write_bytes(b"bad")
        with self.assertRaises(BadZipFile):
            self.store.import_files([new, bad])
        event = threading.Event()
        with self.assertRaises(ExportCancelled):
            self.store.import_files([new], lambda _: event.set(), event)
        with patch.object(self.store, "refresh_analysis_view", side_effect=RuntimeError("commit boundary")):
            with self.assertRaises(RuntimeError):
                self.store.import_files([new])
        self.assertEqual(self.store.restore(), before)
        self.assertEqual(self.store.db.execute("FROM analysis_details").fetchall(), original)
        self.assertEqual(self.store.db.execute("SELECT count(*) FROM sources").fetchone(), (1,))
        self.assertEqual(self.store.db.execute("FROM calendar").fetchall(), [(date(2025, 1, 1), 1)])
        self.store.import_files([new])
        self.assertFalse(list(self.root.glob(".split-import-*")))

    def test_duplicate_fields_types_ragged_and_literal_equal(self):
        a = self.root / "mixed.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["客户", "字段", "字段"])
        ws.append(["A/B", "001", True, "=literal"])
        ws["D2"].data_type = "s"
        ws.append(["A?B", 1.25, False, "尾"])
        wb.save(a)
        wb.close()
        _, fields = self.store.import_files([a])
        self.assertEqual(fields, ["客户", "字段", "字段（同名2）", "未命名列4"])
        self.export(fields=fields)
        self.assertCountEqual(
            self.output_rows(), [("A/B", "001", True, "=literal"), ("A?B", 1.25, False, "尾")]
        )
        self.export("subset", fields=fields[2:3])
        self.assertCountEqual(self.output_rows("subset"), [(True,), (False,)])

    def test_manifest_excludes_stale_customers_and_parts(self):
        original = [(2025, 1, 1, "甲", 1, "X")] * 2 + [(2025, 1, 1, "乙", 3, "Y")]
        self.store.import_files([self.file("old.xlsx", original)])
        with patch("core.MAX_DATA_ROWS", 1):
            self.export()
        self.store.import_files([self.file("fresh.xlsx", original[:1])])
        result = self.export()
        with result["manifest"].open(encoding="utf-8-sig", newline="") as report:
            manifest = list(csv.reader(report))
        self.assertEqual(len(manifest), 2)
        path, count = manifest[1]
        self.assertEqual(count, "1")
        self.assertEqual(read_excel(result["manifest"].parent / path)[1:], original[:1])
        self.assertEqual(len(list((self.root / "out").rglob("*.xlsx"))), 3)

    def test_invalid_date_cannot_replace_valid_day(self):
        self.store.import_files([self.file("old.xlsx", [(2025, 1, 1, "甲", 1, "A")])])
        self.store.import_files([self.file("invalid.xlsx", [(2025, 1, 1.9, "乙", 2, "B")])])
        self.assertEqual(self.store.dates(), ["2025-01-01"])
        self.assertEqual(self.export()["rows"], 2)
        self.assertEqual(self.export("one", start="2025-01-01")["rows"], 1)

    def test_uncached_formula_rejects_whole_snapshot(self):
        f = self.file("formula.xlsx", [(2026, 1, 1, "甲", "=1+2", "x")])
        with self.assertRaisesRegex(ValueError, "公式"):
            self.store.import_files([f])
        self.assertEqual(self.store.restore(), (None, []))

    def test_old_sqlite_migration_and_idempotence(self):
        path = self.root / "历史.sqlite3"
        old = ImportStore(path)
        rows = [(2025, 1, 1, "甲", 12.01, "001"), (2025, 2, 1, "乙", -0.25, "002")]
        try:
            task, _ = old.import_files([self.file("a.xlsx", rows)])
            # Pre-fix legacy column naming keyed the second header by position.
            with old.connection as db:
                db.execute("INSERT INTO source_column VALUES (?, '金额', '金额_7')", (task,))
                for rid, raw in db.execute("SELECT row_id,raw_json FROM staging_row").fetchall():
                    row = json.loads(raw)
                    row["金额_7"] = 45.67
                    db.execute("UPDATE staging_row SET raw_json=? WHERE row_id=?", (json.dumps(row), rid))
            old.import_files([self.file("later.xlsx", [(2025, 2, 1, "更新", 8.55, "003")])])
        finally:
            old.close()
        before = path.read_bytes()
        state = self.store.migrate_sqlite(path)
        self.assertEqual(state, self.store.migrate_sqlite(path))
        self.assertEqual(path.read_bytes(), before)
        self.export(fields=[*HEADER, "金额（同名2）"])
        self.assertCountEqual(
            [r + (None,) * (7 - len(r)) for r in self.output_rows()],
            [(*rows[0], 45.67), (2025, 2, 1, "更新", 8.55, "003", None)],
        )

    def test_fixtures_every_cell_all_days_and_range(self):
        sources = sorted(Path(__file__).with_name("test_samples").glob("*.xlsx"))
        _, fields = self.store.import_files(sources)
        from core import resolve_region

        rules = load_region_rules(Path(__file__).with_name("regions.json"))
        expected = []
        for f in sources:
            header, *rows = read_excel(f)
            for row in rows:
                r = dict(zip(header, row))
                expected.append(
                    (
                        str(date(r["年"], r["月"], r["日"])),
                        resolve_region(r["代表处"], rules),
                        r["客户名称"],
                        tuple(r.get(k) for k in fields),
                    )
                )
        dates = self.store.dates()
        for folder, start, end in (("daily", None, None), ("period", dates[0], dates[-1])):
            result = self.store.export(
                self.root / folder, fields, "客户名称", "代表处", rules, start, end, []
            )
            actual = Counter()
            for path in (self.root / folder).rglob("*.xlsx"):
                header, *rows = read_excel(path)
                self.assertEqual(header, tuple(fields))
                for row in rows:
                    actual[path.parent.parent.name, path.parent.name, path.stem, row] += 1
            self.assertEqual(
                actual, Counter((f"{start}至{end}" if start else d, a, c, r) for d, a, c, r in expected)
            )
            self.assertEqual(result["rows"], 96)


if __name__ == "__main__":
    unittest.main()
