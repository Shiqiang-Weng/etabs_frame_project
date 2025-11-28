#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Compatibility proxy for legacy design entry.
Actual implementation lives in results_extraction.design_workflow.
"""

from results_extraction.design_workflow import perform_concrete_design_and_extract_results

__all__ = ["perform_concrete_design_and_extract_results"]
