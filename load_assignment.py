#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# NOTE: This module has been replaced by load_module.assignment. It remains as a thin
# wrapper for backward compatibility.
"""Legacy wrapper for load assignment; please use load_module.assignment instead."""

from load_module import (
    assign_all_loads_to_frame_structure,
    assign_column_loads_fixed,
    assign_dead_and_live_loads_to_slabs,
    assign_finish_loads_to_beams,
    assign_loads_to_model,
    assign_seismic_mass_to_structure,
)

__all__ = [
    "assign_dead_and_live_loads_to_slabs",
    "assign_finish_loads_to_beams",
    "assign_column_loads_fixed",
    "assign_seismic_mass_to_structure",
    "assign_all_loads_to_frame_structure",
    "assign_loads_to_model",
]
