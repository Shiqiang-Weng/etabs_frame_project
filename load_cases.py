#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# NOTE: This module has been replaced by load_module.cases. It remains as a thin
# wrapper for backward compatibility.
"""Legacy wrapper for load case definition; please use load_module.cases instead."""

from load_module import (
    define_all_load_cases,
    define_load_cases,
    define_mass_source_simple,
    define_modal_case,
    define_response_spectrum_cases,
    define_response_spectrum_combinations,
    define_static_load_cases,
    ensure_dead_pattern,
    ensure_live_pattern,
)

__all__ = [
    "ensure_dead_pattern",
    "ensure_live_pattern",
    "define_static_load_cases",
    "define_modal_case",
    "define_response_spectrum_cases",
    "define_response_spectrum_combinations",
    "define_mass_source_simple",
    "define_all_load_cases",
    "define_load_cases",
]
