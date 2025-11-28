#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Compatibility proxy for legacy imports.
Routes calls to the new analysis package.
"""

from analysis import safe_run_analysis, wait_and_run_analysis, check_analysis_completion

__all__ = ["safe_run_analysis", "wait_and_run_analysis", "check_analysis_completion"]
