import json
import unittest
from collections import Counter
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from core import ImportStore, export_task, load_region_rules, resolve_region, scan_headers


def write_excel(path, rows):
    workbook = Workbook()
    for row in rows:
        workbook.active.append(row)
    workbook.save(path)
    workbook.close()
    return path


def read_excel(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        return list(workbook.active.values)
    finally:
        workbook.close()


class ExportIntegrityTests(unittest.TestCase):
    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.output = self.root / "output"
        self.store = ImportStore(self.root / "store.sqlite3")
        self.addCleanup(self.store.close)

    def export(self, task, fields, customer="客户", **options):
        return export_task(self.store, task, self.output, fields, customer, None, None, **options)

    def test_duplicate_headers_and_literal_suffixes_across_files(self):
        a = write_excel(self.root / "a.xlsx", [
            ["客户", "金额", "金额_4", "金额", "金额（同名2）", "@2:金额", "编码"],
            ["甲", 100, 200, 300, 400, 500, "0001"],
            ["甲", 100, 200, 300, 400, 500, "0001"],
        ])
        b = write_excel(self.root / "b.xlsx", [
            ["金额", "金额", "客户", "金额_4", "@2:金额", "金额（同名2）", "编码"],
            [10, 30, "甲", 20, 50, 40, "0002"],
        ])
        task, fields = self.store.import_files([a, b])
        self.assertEqual(len(set(fields)), 7)
        self.assertEqual(fields, scan_headers([a, b]))
        self.export(task, fields)
        file, = self.output.rglob("*.xlsx")
        header, *rows = read_excel(file)
        self.assertEqual(header, read_excel(a)[0])
        expected = read_excel(a)[1:] + [("甲", 10, 20, 30, 40, 50, "0002")]
        self.assertCountEqual(rows, expected)

        # Choosing only the second duplicate must also work as a grouping key.
        self.output = self.root / "subset"
        self.export(task, [fields[3], "编码"], customer=fields[3])
        self.assertEqual({path.stem for path in self.output.rglob("*.xlsx")}, {"30", "300"})
        for path in self.output.rglob("*.xlsx"):
            header, *rows = read_excel(path)
            self.assertEqual(header, ("金额", "编码"))
            self.assertTrue(all(value == int(path.stem) for value, _ in rows))

    def test_legacy_persisted_duplicate_keys(self):
        # v1.0.3 persisted display_name twice and used the absolute column
        # position to distinguish the two keys in raw_json.
        task, _ = self.store.import_files([])
        with self.store.connection as db:
            db.executemany("INSERT INTO source_column VALUES (?, ?, ?)", [
                (task, "客户", "客户"), (task, "金额", "金额"), (task, "金额", "金额_3"),
            ])
            db.execute("INSERT INTO staging_row(task_id,file_id,source_row,raw_json) VALUES (?,1,2,?)",
                       (task, json.dumps({"客户": "甲", "金额": 100, "金额_3": 200})))
        self.export(task, ["客户", "金额", "金额"])
        path, = self.output.rglob("*.xlsx")
        self.assertEqual(read_excel(path), [("客户", "金额", "金额"), ("甲", 100, 200)])
        second = list(self.store.fields_for_task(task))[2]
        self.export(task, [second])
        self.assertEqual(read_excel(path), [("金额",), (200,)])

    def test_collisions_retries_and_multipart(self):
        rows = [("客户", "金额"), ("A/B", 0), ("A/B", 0), ("A/B", 3),
                ("A?B", 4), ("A_B", 5), ("A_B_part2", 6)]
        source = write_excel(self.root / "input.xlsx", rows)
        task, fields = self.store.import_files([source])
        with patch("core.MAX_DATA_ROWS", 2):
            result = self.export(task, fields)
            files = sorted(self.output.rglob("*.xlsx"))
            self.assertEqual((result["rows"], result["files"], len(files)), (6, 5, 5))
            contents = {p: read_excel(p)[1:] for p in files}
            self.assertCountEqual([r for block in contents.values() for r in block], rows[1:])
            self.assertTrue(all(len({r[0] for r in block}) == 1 for block in contents.values()))
            self.export(task, fields, customer_names=["A?B"])
            self.assertEqual({p: read_excel(p)[1:] for p in self.output.rglob("*.xlsx")}, contents)

    def test_existing_unmarked_file_survives_repeated_exports(self):
        source = write_excel(self.root / "input.xlsx", [["客户", "金额"], ["甲", 10]])
        task, fields = self.store.import_files([source])
        self.export(task, fields)
        prior, = self.output.rglob("*.xlsx")
        write_excel(prior, [["原文件"], ["保留"]])
        prior_bytes = prior.read_bytes()
        self.export(task, fields)
        self.export(task, fields)
        files = sorted(self.output.rglob("*.xlsx"))
        self.assertEqual(len(files), 2)
        self.assertEqual(prior.read_bytes(), prior_bytes)
        other, = [path for path in files if path != prior]
        self.assertEqual(read_excel(other), read_excel(source))

    def test_failed_publication_preserves_previous_workbook(self):
        source = write_excel(self.root / "input.xlsx", [["客户", "金额"], ["甲", 10]])
        task, fields = self.store.import_files([source])
        self.export(task, fields)
        path, = self.output.rglob("*.xlsx")
        before = path.read_bytes()
        with patch.object(Path, "replace", side_effect=PermissionError("文件正在使用")):
            with self.assertRaises(PermissionError):
                self.export(task, ["金额"])
        self.assertEqual(path.read_bytes(), before)
        self.assertEqual(list(path.parent.iterdir()), [path])
        self.export(task, ["金额"])
        self.assertEqual(read_excel(path), [("金额",), (10,)])

    def test_fixtures_all_days_and_date_range(self):
        sources = sorted((Path(__file__).parent / "test_samples").glob("*.xlsx"))
        task, fields = self.store.import_files(sources)
        rules = load_region_rules(Path(__file__).with_name("regions.json"))
        expected = {}
        for source in sources:
            header, *rows = read_excel(source)
            for row in rows:
                record = dict(zip(header, row))
                day = f"{record['年']:04}-{record['月']:02}-{record['日']:02}"
                region = resolve_region(record["代表处"], rules)
                key = (day, region, record["客户名称"])
                expected.setdefault(key, Counter())[tuple(record[f] for f in fields)] += 1
        self.assertEqual(sum(sum(block.values()) for block in expected.values()), 96)
        days = sorted(key[0] for key in expected)
        for date_range in (False, True):
            output = self.root / str(date_range)
            start, end = (days[0], days[-1]) if date_range else (None, None)
            result = export_task(self.store, task, output, fields, "客户名称", "代表处", "年", "月", "日",
                                 rules, selected_date=start, date_end=end)
            grouped = {}
            for (day, region, customer), block in expected.items():
                folder = f"{start}至{end}" if date_range else day
                grouped.setdefault((folder, region, customer), Counter()).update(block)
            actual = {}
            for path in output.rglob("*.xlsx"):
                header, *rows = read_excel(path)
                self.assertEqual(header, tuple(fields))
                actual[(path.parent.parent.name, path.parent.name, path.stem)] = Counter(rows)
            self.assertEqual(actual, grouped)
            self.assertEqual((result["rows"], result["files"]), (96, len(grouped)))


if __name__ == "__main__":
    unittest.main()
