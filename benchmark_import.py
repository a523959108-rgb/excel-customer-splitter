"""Compare Excel import strategies in disposable DuckDB stores."""

import argparse
import hashlib
import json
import statistics
import tempfile
import time
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor
from pathlib import Path

from openpyxl import Workbook

from benchmark import FIELDS
from warehouse import Warehouse, spool_excel


def make_inputs(root, rows, file_count):
    files = []
    for j in range(file_count):
        wb = Workbook(write_only=True)
        sheet = wb.create_sheet("数据")
        sheet.append(FIELDS)
        for i in range(j * rows // file_count, (j + 1) * rows // file_count):
            sheet.append((2025, 2, 1 + i % 28, f"客户{i % 1000}", (i % 99991) / 100, i))
        path = root / f"input-{j}.xlsx"
        wb.save(path)
        files.append(path)
    return files


def spool_job(args):
    item, directory = args
    return spool_excel(item, directory, None)


def cached_file(path, cache):
    with path.open("rb") as f:
        digest = hashlib.file_digest(f, "sha256").hexdigest()
    spooled, _, sheet, fingerprint, columns, count = cache[digest]
    return spooled, str(path), sheet, fingerprint, columns, count


def verify(db, n, customer_count):
    # Every source identifier must occur once and its five remaining fields must
    # match the source recipe. This check runs outside all timed intervals.
    count, distinct_ids, bad = db.execute(
        """
        WITH rows AS (
            SELECT union_extract(f5,'i') id, union_extract(f0,'i') y, union_extract(f1,'i') m,
                   union_extract(f2,'i') d, union_extract(f3,'s') client,
                   coalesce(union_extract(f4,'m'), union_extract(f4,'i')) amount,
                   business_date dt FROM details
        )
        SELECT count(*), count(DISTINCT id),
               count(*) FILTER (WHERE id IS NULL OR id NOT BETWEEN 0 AND ? OR
                   y IS DISTINCT FROM 2025 OR m IS DISTINCT FROM 2 OR d IS DISTINCT FROM 1+id%28 OR
                   dt IS DISTINCT FROM DATE '2025-02-01'+(id%28)::INT OR
                   client IS DISTINCT FROM '客户'||(id%?) OR
                   amount IS DISTINCT FROM (id%99991)*0.01)
        FROM rows
    """,
        [n - 1, customer_count],
    ).fetchone()
    if (count, distinct_ids, bad) != (n, n, 0):
        raise ValueError(f"Import mismatch: {(count, distinct_ids, bad)}")


def run_case(root, paths, rows, mode, cache, customer_count=1000, cache_directory=None):
    with tempfile.TemporaryDirectory(dir=root) as directory:
        store = Warehouse(Path(directory) / "trial.duckdb")
        try:
            start = time.perf_counter()
            load_time = 0.0
            with store.transaction():
                batch = store.start_batch()
                if mode == "disk_cache":
                    # The cache is prebuilt. Rehash the actual input on every hit.
                    for path in paths:
                        entry = cached_file(path, cache)
                        load_start = time.perf_counter()
                        store.load_file(batch, entry)
                        load_time += time.perf_counter() - load_start
                else:
                    kind, workers = mode.split("_")
                    executor = ThreadPoolExecutor if kind == "thread" else ProcessPoolExecutor
                    with executor(max_workers=int(workers)) as pool:
                        jobs = ((item, cache_directory or directory) for item in enumerate(paths))
                        for entry in pool.map(spool_job, jobs):
                            load_start = time.perf_counter()
                            store.load_file(batch, entry)
                            load_time += time.perf_counter() - load_start
                            if cache_directory is None:
                                entry[0].unlink()
                            else:
                                cache[entry[3]] = entry
                staged = time.perf_counter()
                store.publish(batch, None)
            store.checkpoint()
            stop = time.perf_counter()
            verify(store.db, rows, customer_count)
            return {
                "total": round(stop - start, 4),
                "read_and_stage": round(staged - start, 4),
                "load_active": round(load_time, 4),
                "publish_commit_checkpoint": round(stop - staged, 4),
                "verified_rows": rows,
            }
        finally:
            store.close()


def compare_same_xlsx(root, paths, rows):
    """Replay benchmark.py's exact three-customer input files into two fresh databases."""
    with tempfile.TemporaryDirectory(dir=root) as cache_directory:
        cache = {}
        cold = run_case(root, paths, rows, "process_4", cache, 3, cache_directory)
        warm = run_case(root, paths, rows, "disk_cache", cache, 3)
        return {
            "process_4": cold,
            "disk_cache": warm,
            "cache_bytes": sum(entry[0].stat().st_size for entry in cache.values()),
            "files": len(paths),
        }


def compare(rows, directory, repeats):
    with tempfile.TemporaryDirectory(dir=directory) as tmp:
        root = Path(tmp)
        paths = make_inputs(root, rows, 4)
        cache_dir = root / "parsed"
        cache_dir.mkdir()
        start = time.perf_counter()
        entries = [spool_job((item, cache_dir)) for item in enumerate(paths)]
        cache_time = time.perf_counter() - start
        cache = {entry[3]: entry for entry in entries}
        cases = ["thread_1", "thread_4", "process_2", "process_4", "disk_cache"]
        timings = {key: [] for key in cases}
        for repeat in range(repeats):
            offset = (repeat * 2) % len(cases)
            for case in cases[offset:] + cases[:offset]:
                result = run_case(root, paths, rows, case, cache)
                timings[case].append(result)
                print(f"Trial {repeat + 1} {case}: {result['total']:.3f}s", flush=True)
        return {
            "rows": rows,
            "files": len(paths),
            "repeats": repeats,
            "xlsx_bytes": sum(p.stat().st_size for p in paths),
            "cache_bytes": sum(e[0].stat().st_size for e in entries),
            "single_worker_spool_seconds": round(cache_time, 4),
            "timings": timings,
            "median": {k: statistics.median(x["total"] for x in v) for k, v in timings.items()},
        }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows", type=int, default=160000)
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--directory", type=Path, default=Path(tempfile.gettempdir()))
    args = parser.parse_args()
    if not 4 <= args.rows <= 4 * 1048575 or args.repeats < 1:
        parser.error("choose 4..4194300 rows and at least one repeat")
    print(json.dumps(compare(args.rows, args.directory, args.repeats), ensure_ascii=False), flush=True)
