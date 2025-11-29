#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
结果提取/后处理阶段的规范入口。
Provides unified interfaces for analysis results, design forces, member forces,
and enhanced design result extraction.
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
from .design_forces import extract_design_forces_and_summary
from .design_results import (
    extract_and_save_beam_results,
    extract_and_save_column_results,
    extract_design_results_enhanced,
    generate_enhanced_summary_report,
    print_enhanced_validation_statistics,
    save_design_results_enhanced,
)
from .member_forces import (
    extract_and_save_frame_forces,
    extract_frame_forces,
    save_forces_to_csv,
)
from common.config import SCRIPT_DIRECTORY
from common.etabs_setup import get_etabs_objects
from .concrete_frame_detail_data import (
    extract_all_concrete_design_data,
    generate_comprehensive_summary_report,
)
from .section_diagnostic import complete_design_workflow as run_section_diagnostics


def extract_modal_and_mass_info() -> None:
    """Backward-compatible modal/mass extraction entry."""
    _, sap_model = get_etabs_objects()
    extract_modal_and_mass_info_core(sap_model)


def extract_story_drifts_improved(target_load_cases: List[str]) -> None:
    """Backward-compatible story drift extraction entry."""
    _, sap_model = get_etabs_objects()
    extract_story_drifts_improved_core(sap_model, target_load_cases)


def extract_all_analysis_results(
    output_dir: Optional[Union[str, Path]] = None,
    sap_model=None,
):
    """Extract analysis results and write the summary workbook."""
    output_directory = Path(output_dir) if output_dir is not None else Path(SCRIPT_DIRECTORY)
    if sap_model is None:
        _, sap_model = get_etabs_objects()

    summary_path = extract_modal_and_drift(sap_model, output_directory)
    print(f"Analysis summary written to Excel: {summary_path}")
    return summary_path


__all__ = [
    "extract_modal_and_mass_info",
    "extract_story_drifts_improved",
    "extract_all_analysis_results",
    "extract_modal_and_drift",
    "export_core_results",
    # design force extraction
    "extract_design_forces_and_summary",
    # member forces
    "extract_frame_forces",
    "save_forces_to_csv",
    "extract_and_save_frame_forces",
    # enhanced design results
    "extract_design_results_enhanced",
    "save_design_results_enhanced",
    "print_enhanced_validation_statistics",
    "generate_enhanced_summary_report",
    "extract_and_save_beam_results",
    "extract_and_save_column_results",
    # diagnostics/reporting
    "extract_all_concrete_design_data",
    "generate_comprehensive_summary_report",
    "run_section_diagnostics",
]
