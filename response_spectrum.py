#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# NOTE: This module has been replaced by load_module.response_spectrum. It remains as a thin
# wrapper for backward compatibility.
"""Legacy wrapper for response spectrum utilities; use load_module.response_spectrum instead."""

from load_module import (
    china_response_spectrum,
    generate_response_spectrum_data,
    define_response_spectrum_functions_in_etabs,
    setup_response_spectrum,
)

__all__ = [
    "china_response_spectrum",
    "generate_response_spectrum_data",
    "define_response_spectrum_functions_in_etabs",
    "setup_response_spectrum",
]
