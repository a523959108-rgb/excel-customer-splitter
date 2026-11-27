"""Check cached and multiprocessing imports against production on the bundled snapshots."""

import hashlib
import json
import shutil
from concurrent.futures import ProcessPoolExecutor
from contextlib import closing
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from benchmark_import import cached_file, spool_job
from warehouse import Warehouse


def correct_order_number(path):
    # Change one known cell without recalculating formulas or altering other worksheet data.
    with ZipFile(path) as original, ZipFile(path.with_suffix(".new"), "w") as revised:
        for item in original.infolist():
            content = original.read(item)
            if item.filename == "xl/worksheets/sheet1.xml":
                before = b"SO-20260211-0010"
                if content.count(before) != 1:
                    raise ValueError("Expected one occurrence of the order number in the fixture")
                content = content.replace(before, b"SO-CORRECTION-0010")
            revised.writestr(item, content)
    path.with_suffix(".new").replace(path)


def typed_rows(store):
    # Source order also checks duplicates, dates, and the original physical row locations.
    rows = store.db.execute("SELECT * FROM details ORDER BY source_id,source_row").fetchall()
    return [tuple((type(value), value) for value in row) for row in rows]


def replay(root):
    inputs = Path(__file__).parent / "test_samples" / "复杂快照"
    a, b, c = [Path(shutil.copy(path, root)) for path in sorted(inputs.glob("*.xlsx"))]
    cache = {}
    steps = []
    with (
        closing(Warehouse(root / "reference.duckdb")) as reference,
        closing(Warehouse(root / "candidate.duckdb")) as candidate,
        ProcessPoolExecutor(max_workers=2) as readers,
    ):
        for step, paths in (
            ("first", [a]),
            ("overlap", [b, c]),
            ("replay", [b, c]),
            ("correction", [b, c]),
        ):
            if step == "correction":
                correct_order_number(b)
            reference.import_files(paths)
            missing = []
            for path in paths:
                with path.open("rb") as file:
                    key = hashlib.file_digest(file, "sha256").hexdigest()
                if key not in cache:
                    missing.append(((len(cache) + len(missing), path), root))
            for entry in readers.map(spool_job, missing):
                cache[entry[3]] = entry
            with candidate.transaction():
                batch = candidate.start_batch()
                for path in paths:
                    candidate.load_file(batch, cached_file(path, cache))
                candidate.publish(batch, None)
            rows = typed_rows(candidate)
            undated = candidate.db.execute(
                "SELECT count(*) FROM details WHERE business_date IS NULL"
            ).fetchone()[0]
            if (
                rows != typed_rows(reference)
                or candidate.fields() != reference.fields()
                or candidate.dates() != reference.dates()
            ):
                raise ValueError(f"Snapshot mismatch during {step}")
            steps.append(
                {
                    "step": step,
                    "rows": len(rows),
                    "fields": len(candidate.fields()),
                    "undated": undated,
                    "cache_hits": len(paths) - len(missing),
                }
            )
    # A changed byte fingerprint retains the earlier file's two undated rows.
    if [(s["rows"], s["cache_hits"]) for s in steps] != [(5427, 0), (9038, 0), (9038, 2), (9040, 1)]:
        raise ValueError(f"Unexpected snapshot sizes or cache hits: {steps}")
    return steps


if __name__ == "__main__":
    with TemporaryDirectory() as directory:
        print(json.dumps(replay(Path(directory))))
