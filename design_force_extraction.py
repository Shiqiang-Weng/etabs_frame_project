#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Compatibility shim for design force extraction.
Actual logic now lives in results_extraction.design_forces.
"""

from results_extraction.design_forces import (
    check_design_completion,
    debug_api_return_structure,
    debug_available_tables,
    debug_pmm_tables,
    extract_basic_frame_forces,
    extract_beam_design_forces,
    extract_column_design_forces,
    extract_column_pmm_design_forces,
    extract_design_forces_and_summary,
    extract_design_forces_simple,
    generate_summary_report,
    print_extraction_summary,
    test_simple_api_call,
)

__all__ = [
    "check_design_completion",
    "debug_api_return_structure",
    "debug_available_tables",
    "debug_pmm_tables",
    "extract_basic_frame_forces",
    "extract_beam_design_forces",
    "extract_column_design_forces",
    "extract_column_pmm_design_forces",
    "extract_design_forces_and_summary",
    "extract_design_forces_simple",
    "generate_summary_report",
    "print_extraction_summary",
    "test_simple_api_call",
]
