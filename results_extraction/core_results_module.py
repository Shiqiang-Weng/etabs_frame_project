#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
核心结果汇总模块
只导出核心分析/设计结果文件，并清理多余的结果 CSV/XLS/XLSX/TXT。
"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Dict, Union

from results_extraction.analysis_results_module import extract_modal_and_drift
from config import SCRIPT_DIRECTORY
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
    """删除输出目录中非核心的结果文件（仅限 csv/xls/xlsx/txt）。"""
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
            print(f" 已删除多余结果文件: {p.name}")
        except Exception as e:
            print(f" 警告: 无法删除文件 {p.name}: {e}")


def _ensure_output_path(filename: str, output_dir: Path) -> Path:
    """
    将设计导出的文件从 SCRIPT_DIRECTORY 移动/收拢到目标目录，返回目标路径。
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
    """仅导出柱 P-M-M 原始包络表，不生成汇总表。"""
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
            print(f"⚠️ 通过表格 {table_key} 导出 P-M-M 数据时出错: {e}")
            success = False
        if success:
            return _ensure_output_path(output_name, output_dir)
    return _ensure_output_path(output_name, output_dir)


def export_core_results(sap_model, output_dir: Union[str, Path]) -> Dict[str, Path]:
    """
    导出核心分析/设计结果文件。
    返回 dict[name, Path]，键固定为 5 个核心文件。
    """
    output_directory = Path(output_dir)
    output_directory.mkdir(parents=True, exist_ok=True)

    result: Dict[str, Path] = {}
    result["analysis_dynamic_summary"] = output_directory / "analysis_dynamic_summary.xlsx"
    result["beam_flexure_envelope"] = output_directory / "beam_flexure_envelope.csv"
    result["beam_shear_envelope"] = output_directory / "beam_shear_envelope.csv"
    result["column_pmm_design_forces_raw"] = output_directory / "column_pmm_design_forces_raw.csv"
    result["column_shear_envelope"] = output_directory / "column_shear_envelope.csv"

    # 动态分析概要
    if not result["analysis_dynamic_summary"].exists():
        result["analysis_dynamic_summary"] = extract_modal_and_drift(sap_model, output_directory)

    # 设计状态检查（失败仍继续尝试导出）
    try:
        if not check_design_completion(sap_model):
            print("⚠️ 设计状态检查未通过，仍将尝试导出核心设计结果。")
    except Exception as e:
        print(f"⚠️ 设计检查出错: {e}")

    # 梁弯矩包络
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
        print(f"⚠️ 梁弯矩包络导出异常: {e}")

    # 梁剪力包络
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
        print(f"⚠️ 梁剪力包络导出异常: {e}")

    # 柱 P-M-M 原始设计内力
    try:
        result["column_pmm_design_forces_raw"] = _export_column_pmm_raw(sap_model, output_directory)
    except Exception as e:
        print(f"⚠️ 柱 P-M-M 导出异常: {e}")

    # 柱剪力包络
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
        print(f"⚠️ 柱剪力包络导出异常: {e}")

    keep_names = {p.name for p in result.values()}
    _cleanup_extra_result_files(output_directory, keep_names)
    return result


__all__ = [
    "CORE_RESULT_BASENAMES",
    "export_core_results",
]
