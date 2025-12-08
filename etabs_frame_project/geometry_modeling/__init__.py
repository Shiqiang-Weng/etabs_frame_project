#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
几何建模阶段的规范入口。Provides geometry modeling APIs for ETABS frame structures.
Exports legacy-compatible symbols while pointing new code to the package namespace.
"""

from .api_compat import (
    _get_all_points_safe,
    _get_name_list_safe,
    _require_sap_model,
    debug_joint_coordinates,
    ensure_model_units,
    get_all_points_reference_method,
)
from .base_constraints import (
    BaseConstraintManager,
    BaseJointLocator,
    fix_base_constraints_comprehensive,
    fix_base_constraints_issue,
    get_base_level_joints,
    get_base_level_joints_by_existing_elements,
    get_base_level_joints_by_grid,
    get_base_level_joints_by_grid_direct,
    get_base_level_joints_reference_method,
    get_base_level_joints_v2,
    set_rigid_base_constraints_fixed,
    set_rigid_base_constraints_improved,
)
from .materials_sections import (
    define_all_materials_and_sections,
    define_diaphragms,
    define_frame_sections,
    define_materials,
    define_slab_sections,
    create_parametric_frame_sections_from_design,
)
from .layout import (
    GridConfig,
    StoryConfig,
    default_grid_config,
    default_story_config,
    grid_config_from_design,
    story_config_from_design,
)
from .model_builder import (
    ElementCreator,
    FrameGeometryWorkflow,
    create_frame_structure,
    create_frame_structure_from_design,
    fix_base_constraints_comprehensive as fix_base_constraints_comprehensive_entry,
)
from .section_assignment import assign_sections_by_design

# Keep __all__ aligned with legacy frame_geometry
__all__ = [
    "create_frame_structure",
    "ensure_model_units",
    "debug_joint_coordinates",
    "get_all_points_reference_method",
    "define_materials",
    "define_frame_sections",
    "define_slab_sections",
    "define_diaphragms",
    "define_all_materials_and_sections",
    "create_parametric_frame_sections_from_design",
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
    # Additional exports for extensibility
    "GridConfig",
    "StoryConfig",
    "default_grid_config",
    "default_story_config",
    "grid_config_from_design",
    "story_config_from_design",
    "ElementCreator",
    "FrameGeometryWorkflow",
    "create_frame_structure_from_design",
    "BaseJointLocator",
    "BaseConstraintManager",
    "assign_sections_by_design",
]

# Alias to keep the same name as in model_builder while preserving legacy __all__
fix_base_constraints_comprehensive = fix_base_constraints_comprehensive_entry
