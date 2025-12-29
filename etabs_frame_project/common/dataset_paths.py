#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared dataset bucketing and marker configuration for large ETABS runs."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Tuple

BUCKET_SIZE = 1000
NUM_BUCKETS = 30
INPUT_BUCKET_PREFIX = "input"
OUTPUT_BUCKET_PREFIX = "output"
DONE_MARKER_FILENAME = "_DONE.flag"
GNN_INPUT_EXT = ".npz"


@dataclass(frozen=True)
class BucketInfo:
    start: int
    end: int
    suffix: str

    @property
    def range_label(self) -> str:
        return f"{self.start}-{self.end}"


def compute_bucket(case_id: int, bucket_size: int = BUCKET_SIZE, num_buckets: int = NUM_BUCKETS) -> BucketInfo:
    """
    Compute the bucket [start, end] that contains the given case_id.

    Raises:
        ValueError: if case_id is negative or exceeds the configured buckets.
    """
    if case_id < 0:
        raise ValueError("case_id must be non-negative")
    bucket_index = case_id // bucket_size
    if bucket_index >= num_buckets:
        raise ValueError(
            f"case_id {case_id} is outside the configured bucket range "
            f"({num_buckets} buckets of size {bucket_size})"
        )

    start = bucket_index * bucket_size
    end = start + bucket_size - 1
    return BucketInfo(start=start, end=end, suffix=f"{start}-{end}")


def iter_bucket_ranges(bucket_size: int = BUCKET_SIZE, num_buckets: int = NUM_BUCKETS) -> Iterable[Tuple[int, int]]:
    """Yield (start, end) pairs for each configured bucket."""
    for idx in range(num_buckets):
        start = idx * bucket_size
        yield start, start + bucket_size - 1


def build_bucket_dir(root: Path, prefix: str, bucket: BucketInfo) -> Path:
    """Return the folder path for a bucket under the given root."""
    return Path(root) / f"{prefix}{bucket.suffix}"
