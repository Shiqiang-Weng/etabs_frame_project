#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Deprecated compatibility shim for legacy imports.
Use results_extraction.design_forces instead of this file.
"""
from warnings import warn

warn(
    "design_force_extraction_old.py is deprecated; use results_extraction.design_forces",
    DeprecationWarning,
    stacklevel=2,
)

from results_extraction.design_forces import *  # noqa: F401,F403
from results_extraction.design_forces import __all__
