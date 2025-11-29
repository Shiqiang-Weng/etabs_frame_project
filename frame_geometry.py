#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Compatibility wrapper for geometry modeling.

All implementation lives in geometry_modeling/* to keep a clear package structure.
Existing imports of frame_geometry remain valid, but new code should import from
the geometry_modeling package directly.
"""

from geometry_modeling import (
    create_frame_structure,
    ensure_model_units,
    debug_joint_coordinates,
    get_all_points_reference_method,
    get_base_level_joints,
    get_base_level_joints_v2,
    get_base_level_joints_by_grid,
    get_base_level_joints_by_grid_direct,
    get_base_level_joints_by_existing_elements,
    get_base_level_joints_reference_method,
    set_rigid_base_constraints_improved,
    set_rigid_base_constraints_fixed,
    fix_base_constraints_comprehensive,
    fix_base_constraints_issue,
    _get_all_points_safe,
    _get_name_list_safe,
)

__all__ = [
    "create_frame_structure",
    "ensure_model_units",
    "debug_joint_coordinates",
    "get_all_points_reference_method",
    "get_base_level_joints",
    "get_base_level_joints_v2",
    "get_base_level_joints_by_grid",
    "get_base_level_joints_by_grid_direct",
    "get_base_level_joints_by_existing_elements",
    "get_base_level_joints_reference_method",
    "set_rigid_base_constraints_improved",
    "set_rigid_base_constraints_fixed",
    "fix_base_constraints_comprehensive",
    "fix_base_constraints_issue",
    "_get_all_points_safe",
    "_get_name_list_safe",
]
