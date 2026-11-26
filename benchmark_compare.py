"""Standalone comparisons on synthetic data. No desktop or warehouse settings are changed."""

import argparse
import json
import statistics
import tempfile
import time
from pathlib import Path

from benchmark import run
from core import normalize_field


def compare_queries(store, sql, params, expected):
    # The six-column fixture only has text customer names. The experimental
    # dictionary resolves aliases before the fact scan; output and order stay identical.
    db = store.db
    start = time.perf_counter()
    raw_names = db.execute("SELECT DISTINCT union_extract(f3,'s') FROM details").fetchall()
    lookup = {}
    for (raw,) in raw_names:
        lookup.setdefault(normalize_field(raw), []).append(raw)
    build = time.perf_counter() - start
    prefix, separator, tail = sql.partition("business_date BETWEEN ? AND ? AND ")
    _, order, ordering = tail.partition(" ORDER BY ")
    if not separator or not order:
        raise ValueError("The benchmark no longer matches the production query")
    probe = prefix + separator + "union_extract(f3,'s') IN (SELECT unnest(?))" + order + ordering
    samples = {key: [] for key in ("current_1", "current_4", "dictionary_1", "dictionary_4")}
    for trial in range(3):
        variants = list(samples)
        # Rotate which variant sees the cache first in each round.
        for name in variants[trial:] + variants[:trial]:
            db.execute(f"SET threads={int(name[-1])}")
            start = time.perf_counter()
            if name.startswith("dictionary"):
                aliases = [raw for key in params[-1] for raw in lookup.get(key, [])]
                rows = db.execute(probe, [*params[:-1], aliases]).fetchall()
            else:
                rows = db.execute(sql, params).fetchall()
            samples[name].append(round(time.perf_counter() - start, 6))
            if rows != expected:
                raise ValueError(f"{name}: ordered projection mismatch")
    db.execute("SET threads=4")
    return {
        "raw_names": len(raw_names),
        "dictionary_build_seconds": round(build, 6),
        "seconds": samples,
        "median_seconds": {key: statistics.median(values) for key, values in samples.items()},
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows", type=int, required=True)
    parser.add_argument("--directory", type=Path, default=Path(tempfile.gettempdir()))
    args = parser.parse_args()
    print(
        json.dumps(run(args.rows, args.directory, query_probe=compare_queries), ensure_ascii=False),
        flush=True,
    )
