"""Evaluate the optional Calamine reader against the production XML parser."""

import argparse
import json
import statistics
import tempfile
import time
from collections import Counter
from itertools import zip_longest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from warehouse import worksheet_rows


def current_rows(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        with ws._get_source() as source:
            yield from worksheet_rows(ws, source)
    finally:
        wb.close()


def native_rows(path):
    from python_calamine import CalamineWorkbook

    wb = CalamineWorkbook.from_path(path)
    try:
        yield from enumerate(wb.get_sheet_by_index(0).iter_rows(), 1)
    finally:
        wb.close()


def compare(files):
    from python_calamine import CalamineError

    seconds = {"current": [], "calamine": []}
    for i in range(3):
        readers = [("current", current_rows), ("calamine", native_rows)]
        for key, reader in readers if i % 2 == 0 else reversed(readers):
            start = time.perf_counter()
            for path in files:
                for _ in reader(path):
                    pass
            seconds[key].append(round(time.perf_counter() - start, 5))

    differences = Counter()
    cells = 0
    for path in files:
        reference = dict(current_rows(path))
        candidate = dict(native_rows(path))
        for r in reference.keys() | candidate.keys():
            pairs = zip_longest(reference.get(r, []), candidate.get(r, []), fillvalue=None)
            for x, y in pairs:
                cells += 1
                if type(x) is not type(y) or x != y:
                    differences[f"{type(x).__name__} -> {type(y).__name__}"] += 1

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "uncached.xlsx"
        wb = Workbook()
        wb.active.append(["客户", "金额"])
        wb.active.append(["甲", "=1+2"])
        wb.save(path)
        wb.close()
        try:
            list(current_rows(path))
        except ValueError as exc:
            if "公式缺少缓存" not in str(exc):
                raise
        else:
            raise ValueError("Production reader accepted an uncached formula")
        try:
            native_value = list(native_rows(path))[-1][1][1:]
        except CalamineError as exc:
            formula_result = {"rejected": True, "reason": str(exc)}
        else:
            formula_result = {"rejected": False, "values_after_customer": native_value}
    return {
        "input_files": [p.name for p in files],
        "physical_cells_compared": cells,
        "seconds": seconds,
        "median": {k: statistics.median(v) for k, v in seconds.items()},
        "raw_differences": differences,
        "uncached_formula_native": formula_result,
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", type=Path, help="folder with .xlsx test fixtures")
    args = parser.parse_args()
    inputs = sorted(args.path.glob("*.xlsx"))
    if not inputs:
        parser.error("no .xlsx fixtures found")
    print(json.dumps(compare(inputs), ensure_ascii=False), flush=True)
