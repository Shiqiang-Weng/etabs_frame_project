#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析/设计求解阶段的规范入口。
Exposes structural analysis and concrete design workflows.
"""

from .runner import safe_run_analysis, wait_and_run_analysis
from .status import check_analysis_completion
from .design_workflow import perform_concrete_design_and_extract_results

__all__ = [
    "safe_run_analysis",
    "wait_and_run_analysis",
    "check_analysis_completion",
    "perform_concrete_design_and_extract_results",
]
