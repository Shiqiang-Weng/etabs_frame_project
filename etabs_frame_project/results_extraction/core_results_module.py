#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Core results exporter.
Only handles the five core files (analysis_dynamic_summary.xlsx, beam_flexure_envelope.csv,
beam_shear_envelope.csv, column_pmm_design_forces_raw.csv, column_shear_envelope.csv)
and does not change any modeling/analysis/design logic.
"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Dict, Union

from results_extraction.analysis_results_module import extract_modal_and_drift
from common.config import SCRIPT_DIRECTORY, DESIGN_DATA_DIR, ANALYSIS_DATA_DIR
from .design_forces import check_design_completion, extract_design_forces_simple

CORE_RESULT_BASENAMES = {
    "analysis_dynamic_summary.xlsx",
    "beam_flexure_envelope.csv",
    "beam_shear_envelope.csv",
    "column_pmm_design_forces_raw.csv",
    "column_shear_envelope.csv",
}
_RESULT_EXTS = {".csv", ".xls", ".xlsx", ".txt"}


def _cleanup_extra_result_files(output_dir: Path, keep_names: set[str]) -> None:
    """Delete non-core result files in the output directory (csv/xls/xlsx/txt only)."""
    output_dir.mkdir(parents=True, exist_ok=True)
    for p in output_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in _RESULT_EXTS:
            continue
        if p.name in keep_names:
            continue
        try:
            p.unlink()
            print(f" ? {p.name}")
        except Exception as e:
            print(f" :  {p.name}: {e}")


def _ensure_output_path(filename: str, output_dir: Path) -> Path:
    """
    Move an exported design file into the target output directory and return the destination path.
    Prefer design_data as source, fall back to the legacy script directory.
    """
    # 统一输出目录：优先从 data_extraction，再回退到原脚本目录
    src_candidates = [
        Path(DESIGN_DATA_DIR) / filename,
        Path(SCRIPT_DIRECTORY) / filename,
    ]
    dest = output_dir / filename
    dest.parent.mkdir(parents=True, exist_ok=True)
    for src in src_candidates:
        if src.exists():
            if dest.exists() and dest.resolve() == src.resolve():
                return dest
            try:
                shutil.move(str(src), dest)
            except Exception:
                shutil.copy2(src, dest)
            return dest
    return dest


def _log_core_file_status(name: str, dest: Path, success: bool) -> None:
    """统一打印核心结果文件的生成状态。"""
    if success and dest.exists():
        try:
            size = dest.stat().st_size
            print(f"[CORE] {name}: {dest} ({size} bytes)")
        except Exception:
            print(f"[CORE] {name}: {dest}")
    else:
        print(f"[WARN] Core result missing: {name} (expected at {dest})")


def _export_core_table(
    sap_model,
    table_key: str,
    output_filename: str,
    output_dir: Path,
    label: str,
) -> tuple[Path, bool]:
    """
    统一的核心表格导出助手，封装 CSV 导出、路径迁移与状态日志。
    仅用于核心结果文件，不修改设计逻辑。
    """
    success = False
    try:
        success = extract_design_forces_simple(
            sap_model,
            table_key,
            None,
            output_filename,
        )
    except Exception as e:
        print(f"Warning: {label} export failed: {e}")

    dest = _ensure_output_path(output_filename, output_dir)
    _log_core_file_status(label, dest, success)
    return dest, success


def _export_column_pmm_raw(sap_model, output_dir: Path) -> Path:
    """Export column P-M-M raw envelope table without generating summaries."""
    pmm_table_candidates = [
        "Concrete Column PMM Envelope - Chinese 2010",
        "Concrete Column PMM - Chinese 2010",
        "Concrete Column Envelope - Chinese 2010",
        "Concrete Column Design - P-M-M Design Forces - Chinese 2010",
        "Concrete Column Design - P-M-M Design Forces",
        "Column Design - P-M-M Design Forces",
    ]
    output_name = "column_pmm_design_forces_raw.csv"
    dest_path = _ensure_output_path(output_name, output_dir)
    for table_key in pmm_table_candidates:
        try:
            success = extract_design_forces_simple(
                sap_model,
                table_key,
                None,
                output_name,
            )
        except Exception as e:
            print(f"  {table_key}  P-M-M 导出失败: {e}")
            success = False
        if success:
            dest_path = _ensure_output_path(output_name, output_dir)
            _log_core_file_status("column_pmm_design_forces_raw", dest_path, True)
            return dest_path
    _log_core_file_status("column_pmm_design_forces_raw", dest_path, False)
    return dest_path


def export_core_results(sap_model, output_dir: Union[str, Path]) -> Dict[str, Path]:
    """
    Export core analysis/design result files and return a mapping of name to path.
    Five key outputs are always included.
    """
    output_directory = Path(output_dir)
    output_directory.mkdir(parents=True, exist_ok=True)
    analysis_directory = Path(ANALYSIS_DATA_DIR)
    analysis_directory.mkdir(parents=True, exist_ok=True)

    result: Dict[str, Path] = {}
    # Core result files are generated/recorded here
    result["analysis_dynamic_summary"] = analysis_directory / "analysis_dynamic_summary.xlsx"
    result["beam_flexure_envelope"] = output_directory / "beam_flexure_envelope.csv"
    result["beam_shear_envelope"] = output_directory / "beam_shear_envelope.csv"
    result["column_pmm_design_forces_raw"] = output_directory / "column_pmm_design_forces_raw.csv"
    result["column_shear_envelope"] = output_directory / "column_shear_envelope.csv"

    # dynamic analysis summary
    if not result["analysis_dynamic_summary"].exists():
        result["analysis_dynamic_summary"] = extract_modal_and_drift(sap_model, analysis_directory)
        _log_core_file_status("analysis_dynamic_summary", result["analysis_dynamic_summary"], True)
    else:
        _log_core_file_status("analysis_dynamic_summary", result["analysis_dynamic_summary"], True)

    # design status check (failure still tries to export)
    try:
        if not check_design_completion(sap_model):
            print("Warning: design status check failed; attempting to export core design results anyway.")
    except Exception as e:
        print(f"Warning: design status check raised an error: {e}")

    # beam flexure envelope
    result["beam_flexure_envelope"], _ = _export_core_table(
        sap_model,
        "Concrete Beam Flexure Envelope - Chinese 2010",
        "beam_flexure_envelope.csv",
        output_directory,
        "beam_flexure_envelope",
    )

    # beam shear envelope
    result["beam_shear_envelope"], _ = _export_core_table(
        sap_model,
        "Concrete Beam Shear Envelope - Chinese 2010",
        "beam_shear_envelope.csv",
        output_directory,
        "beam_shear_envelope",
    )

    # column P-M-M design forces
    try:
        result["column_pmm_design_forces_raw"] = _export_column_pmm_raw(
            sap_model, output_directory
        )
    except Exception as e:
        print(f"Warning: column P-M-M export failed: {e}")

    # column shear envelope
    result["column_shear_envelope"], _ = _export_core_table(
        sap_model,
        "Concrete Column Shear Envelope - Chinese 2010",
        "column_shear_envelope.csv",
        output_directory,
        "column_shear_envelope",
    )

    # 保留 data_extraction 中的其他结果文件，避免误删过滤/增强输出
    keep_names = {p.name for p in result.values()}

    missing_names = [name for name, path in result.items() if not path.exists()]
    if missing_names:
        print(f"[WARN] Core result files not generated: {missing_names}")
    return result


__all__ = [
    "CORE_RESULT_BASENAMES",
    "export_core_results",
]
