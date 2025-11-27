#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
结果提取包
提供模态/层间位移提取以及核心结果导出接口。
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Union

from .analysis_results_module import (
    extract_modal_and_mass_info as extract_modal_and_mass_info_core,
    extract_modal_and_drift,
    extract_story_drifts_improved as extract_story_drifts_improved_core,
)
from .core_results_module import export_core_results
from config import SCRIPT_DIRECTORY
from etabs_setup import get_etabs_objects


def extract_modal_and_mass_info() -> None:
    """兼容旧接口，调用新的模态/质量参与提取逻辑。"""
    _, sap_model = get_etabs_objects()
    extract_modal_and_mass_info_core(sap_model)


def extract_story_drifts_improved(target_load_cases: List[str]) -> None:
    """兼容旧接口，调用新的层间位移角提取逻辑。"""
    _, sap_model = get_etabs_objects()
    extract_story_drifts_improved_core(sap_model, target_load_cases)


def extract_all_analysis_results(
    output_dir: Optional[Union[str, Path]] = None,
    sap_model=None,
):
    """提取分析结果并将输出写入 Excel 总结。"""
    output_directory = Path(output_dir) if output_dir is not None else Path(SCRIPT_DIRECTORY)
    if sap_model is None:
        _, sap_model = get_etabs_objects()

    summary_path = extract_modal_and_drift(sap_model, output_directory)
    print(f"动态分析结果概要已写入 Excel: {summary_path}")
    return summary_path


__all__ = [
    "extract_modal_and_mass_info",
    "extract_story_drifts_improved",
    "extract_all_analysis_results",
    "export_core_results",
    "extract_modal_and_drift",
]
