#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
???CSV/XLS/XLSX/TXT?"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Dict, Union

from results_extraction.analysis_results_module import extract_modal_and_drift
from common.config import SCRIPT_DIRECTORY
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
    Move an exported design file from SCRIPT_DIRECTORY into the target output
    directory and return the destination path.
    """
    src = Path(SCRIPT_DIRECTORY) / filename
    dest = output_dir / filename
    dest.parent.mkdir(parents=True, exist_ok=True)
    if src.exists():
        if dest.exists() and dest.resolve() == src.resolve():
            return dest
        try:
            shutil.move(str(src), dest)
        except Exception:
            shutil.copy2(src, dest)
    return dest


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
    for table_key in pmm_table_candidates:
        try:
            success = extract_design_forces_simple(
                sap_model,
                table_key,
                None,
                output_name,
            )
        except Exception as e:
            print(f"  {table_key}  P-M-M ? {e}")
            success = False
        if success:
            return _ensure_output_path(output_name, output_dir)
    return _ensure_output_path(output_name, output_dir)


def export_core_results(sap_model, output_dir: Union[str, Path]) -> Dict[str, Path]:
    """
    Export core analysis/design result files and return a mapping of name to path.
    Five key outputs are always included.
    """
    output_directory = Path(output_dir)
    output_directory.mkdir(parents=True, exist_ok=True)

    result: Dict[str, Path] = {}
    result["analysis_dynamic_summary"] = output_directory / "analysis_dynamic_summary.xlsx"
    result["beam_flexure_envelope"] = output_directory / "beam_flexure_envelope.csv"
    result["beam_shear_envelope"] = output_directory / "beam_shear_envelope.csv"
    result["column_pmm_design_forces_raw"] = output_directory / "column_pmm_design_forces_raw.csv"
    result["column_shear_envelope"] = output_directory / "column_shear_envelope.csv"

    # dynamic analysis summary
    if not result["analysis_dynamic_summary"].exists():
        result["analysis_dynamic_summary"] = extract_modal_and_drift(sap_model, output_directory)

    # design status check (failure still tries to export)
    try:
        if not check_design_completion(sap_model):
            print("Warning: design status check failed; attempting to export core design results anyway.")
    except Exception as e:
        print(f"Warning: design status check raised an error: {e}")

    # beam flexure envelope
    try:
        if extract_design_forces_simple(
            sap_model,
            "Concrete Beam Flexure Envelope - Chinese 2010",
            None,
            "beam_flexure_envelope.csv",
        ):
            result["beam_flexure_envelope"] = _ensure_output_path(
                "beam_flexure_envelope.csv", output_directory
            )
    except Exception as e:
        print(f"Warning: beam flexure envelope export failed: {e}")

    # beam shear envelope
    try:
        if extract_design_forces_simple(
            sap_model,
            "Concrete Beam Shear Envelope - Chinese 2010",
            None,
            "beam_shear_envelope.csv",
        ):
            result["beam_shear_envelope"] = _ensure_output_path(
                "beam_shear_envelope.csv", output_directory
            )
    except Exception as e:
        print(f"Warning: beam shear envelope export failed: {e}")

    # column P-M-M design forces
    try:
        result["column_pmm_design_forces_raw"] = _export_column_pmm_raw(
            sap_model, output_directory
        )
    except Exception as e:
        print(f"Warning: column P-M-M export failed: {e}")

    # column shear envelope
    try:
        if extract_design_forces_simple(
            sap_model,
            "Concrete Column Shear Envelope - Chinese 2010",
            None,
            "column_shear_envelope.csv",
        ):
            result["column_shear_envelope"] = _ensure_output_path(
                "column_shear_envelope.csv", output_directory
            )
    except Exception as e:
        print(f"Warning: column shear envelope export failed: {e}")

    keep_names = {p.name for p in result.values()}
    _cleanup_extra_result_files(output_directory, keep_names)
    return result

__all__ = [
    "CORE_RESULT_BASENAMES",
    "export_core_results",
]

