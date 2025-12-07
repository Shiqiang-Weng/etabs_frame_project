#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shared helpers for exporting ETABS tables to CSV with optional filtering and
standardized logging. These helpers do NOT modify any modeling/analysis/design
logic; they only consolidate the repetitive CSV export/filter/status code used
across results extraction modules.
"""

from __future__ import annotations

import csv
import os
from typing import Callable, Iterable, Tuple


def export_table_to_csv(db, System, table_key: str, output_file: str, table_version: int = 1) -> Tuple[bool, object, int]:
    """
    Call DatabaseTables.GetTableForDisplayCSVFile and print its return value/type.

    Returns:
        (success, ret_csv, file_size_bytes)
    """
    field_key_list = System.Array.CreateInstance(System.String, 1)
    field_key_list[0] = ""

    group_name = ""
    table_version_val = System.Int32(table_version)

    ret_csv = db.GetTableForDisplayCSVFile(
        table_key,
        field_key_list,
        group_name,
        table_version_val,
        output_file,
    )

    print(f"[CSV] {table_key} -> {output_file}")
    print(f"[CSV] return: {ret_csv} (type: {type(ret_csv)})")

    success = (isinstance(ret_csv, tuple) and ret_csv[0] == 0) or ret_csv == 0
    file_size = os.path.getsize(output_file) if success and os.path.exists(output_file) else 0
    return success and file_size > 0, ret_csv, file_size


def find_component_name_column(headers: Iterable[str]) -> int | None:
    """Detect the column index that likely holds component names (Unique/Element/Label/Name)."""
    name_col_index = None
    for i, header in enumerate(headers):
        h = (header or "").lower()
        if any(kw in h for kw in ["unique", "element", "label", "name"]) and "combo" not in h:
            name_col_index = i
            break
    return name_col_index


def build_component_filter(component_names: list[str] | None) -> Callable[[list[str], list[str]], bool]:
    """
    Build a row filter based on component names. If component_names is empty/None,
    the filter always returns True.
    """
    if not component_names:
        return lambda headers, row: True

    def _filter(headers: list[str], row: list[str]) -> bool:
        idx = find_component_name_column(headers)
        if idx is None:
            return True
        return len(row) > idx and row[idx] in component_names

    return _filter


def filter_csv_rows(input_file: str, output_file: str, keep_row: Callable[[list[str], list[str]], bool]) -> Tuple[int, int]:
    """
    Filter a CSV file using keep_row(headers, row) -> bool.

    Returns:
        (total_rows_read, rows_written)
    """
    total_count = 0
    written_count = 0

    with open(input_file, "r", encoding="utf-8-sig") as infile, open(
        output_file, "w", newline="", encoding="utf-8-sig"
    ) as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        headers = next(reader, None)
        if headers is None:
            return total_count, written_count

        writer.writerow(headers)

        for row in reader:
            total_count += 1
            if keep_row(headers, row):
                writer.writerow(row)
                written_count += 1

    return total_count, written_count


def log_table_status(table_name: str, record_count: int, csv_path: str, filtered_path: str | None, file_size: int) -> None:
    """Standardized status printout for table export/filter."""
    print(f"[TABLE] {table_name}: {record_count} rows written to {csv_path} ({file_size} bytes)")
    if filtered_path:
        print(f"[TABLE] filtered -> {filtered_path}")


__all__ = [
    "export_table_to_csv",
    "find_component_name_column",
    "build_component_filter",
    "filter_csv_rows",
    "log_table_status",
]
