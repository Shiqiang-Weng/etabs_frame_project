#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Export additional ETABS 'Other Output Items' tables for downstream labeling."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Tuple

from common.etabs_api_loader import get_api_objects
from .export_utils import export_table_to_csv

TABLE_EXPORTS: Tuple[Tuple[str, str], ...] = (
    ("Centers Of Mass And Rigidity", "centers_of_mass_and_rigidity.csv"),
    ("Story Forces", "story_forces.csv"),
    ("Diaphragm Forces", "diaphragm_forces.csv"),
    ("Story Stiffness", "story_stiffness.csv"),
    ("Shear Gravity Ratios", "shear_gravity_ratios.csv"),
    ("Stiffness Gravity Ratios", "stiffness_gravity_ratios.csv"),
    ("Frame Overturning Moments in Dual Systems", "frame_overturning_moments_dual_systems.csv"),
)


def export_other_output_items_tables(sap_model, output_dir) -> Dict[str, Dict[str, object]]:
    """
    Export a curated list of "Other Output Items" tables to CSV.

    Returns:
        Mapping from table name to status dict {path, size, success, error?}.
    """
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    if sap_model is None:
        print("[WARN] sap_model is None; skip other output items export.")
        return {}

    db = getattr(sap_model, "DatabaseTables", None)
    if db is None:
        print("[WARN] sap_model.DatabaseTables not available; skip other output items export.")
        return {}

    _, System, _ = get_api_objects()
    results: Dict[str, Dict[str, object]] = {}
    print("[阶段4] 追加导出：ANALYSIS RESULTS -> Structure Output -> Other Output Items")

    for table_name, filename in TABLE_EXPORTS:
        csv_path = output_path / filename
        try:
            success, ret_csv, file_size = export_table_to_csv(db, System, table_name, str(csv_path))
            results[table_name] = {
                "path": str(csv_path),
                "size": file_size,
                "success": bool(success),
                "ret": ret_csv,
            }
            if success:
                print(f"  [OK] {table_name} -> {csv_path} ({file_size} bytes)")
            else:
                print(f"  [WARN] {table_name} 未生成有效文件 (ret={ret_csv}); 输出路径: {csv_path}")
        except Exception as exc:  # noqa: BLE001
            results[table_name] = {
                "path": str(csv_path),
                "size": 0,
                "success": False,
                "error": str(exc),
            }
            print(f"  [WARN] 导出 {table_name} 时发生异常: {exc}")

    return results


__all__ = ["export_other_output_items_tables", "TABLE_EXPORTS"]

