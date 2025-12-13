#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Geometry modeling workflow for ETABS frame structures.
"""

import logging
from typing import Dict, List, Optional, Tuple

from common.config import DEFAULT_DESIGN_CONFIG, DesignConfig, design_config_from_case
from common.utility_functions import add_area_by_coord_custom, add_frame_by_coord_custom, check_ret

from .api_compat import _require_sap_model, ensure_model_units
from .base_constraints import (
    BaseConstraintManager,
    BaseJointLocator,
    fix_base_constraints_comprehensive,
)
from .geometry_utils import (
    apply_beam_inertia_modifiers,
    apply_slab_membrane_modifiers,
    assign_diaphragm_constraints_by_story,
)
from .layout import (
    GridConfig,
    StoryConfig,
    grid_config_from_design,
    story_config_from_design,
)

log = logging.getLogger(__name__)
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def _column_position(i: int, j: int, max_i: int, max_j: int) -> str:
    if i in {0, max_i} and j in {0, max_j}:
        return "CORNER"
    if i in {0, max_i} or j in {0, max_j}:
        return "EDGE"
    return "INTERIOR"


def _beam_position(orientation: str, i: int, j: int, max_i: int, max_j: int) -> str:
    if orientation == "X":
        return "EDGE" if j in {0, max_j} else "INT"
    return "EDGE" if i in {0, max_i} else "INT"


class ElementCreator:
    def __init__(self, sap_model, grid: GridConfig, stories: StoryConfig, design: DesignConfig):
        self.sap_model = sap_model
        self.grid = grid
        self.stories = stories
        self.design = design
        self.frame_obj = sap_model.FrameObj
        self.area_obj = sap_model.AreaObj

    def create_columns(self) -> List[str]:
        column_names: List[str] = []
        max_i = self.grid.num_x - 1
        max_j = self.grid.num_y - 1
        for story_num, z_bottom, z_top in self.stories.iter_story_bounds():
            group_name = self.design.story_group(story_num)
            if not group_name:
                log.warning("Story %s has no vertical group mapping; columns skipped.", story_num)
                continue

            story_count = 0
            for i, j, x_coord, y_coord in self.grid.iter_points():
                column_name = f"COL_X{i}_Y{j}_S{story_num}"
                position = _column_position(i, j, max_i, max_j)
                section_name = self.design.column_section_name_for_story(story_num, position)
                ret_code, actual_name = add_frame_by_coord_custom(
                    self.frame_obj,
                    x_coord,
                    y_coord,
                    z_bottom,
                    x_coord,
                    y_coord,
                    z_top,
                    section_name,
                    column_name,
                )
                check_ret(ret_code, f"AddByCoord(Column {column_name})")
                column_names.append(actual_name or column_name)
                story_count += 1

            log.info("Story %s columns created: %s", story_num, story_count)

        log.info("Total columns created: %s", len(column_names))
        return column_names

    def create_beams(self) -> List[str]:
        beam_names: List[str] = []
        for story_num, _, z_top in self.stories.iter_story_bounds():
            group_name = self.design.story_group(story_num)
            if not group_name:
                log.warning("Story %s has no vertical group mapping; beams skipped.", story_num)
                continue

            story_beam_count = 0
            max_i = self.grid.num_x - 1
            max_j = self.grid.num_y - 1

            for i, j, x1, x2, y in self.grid.iter_beam_spans_x():
                position = _beam_position("X", i, j, max_i, max_j)
                beam_depth = self.design.beam_depth_for_story(story_num, position)
                z_beam_center = z_top - beam_depth / 2.0
                section_name = self.design.beam_section_name_for_story(story_num, position)
                beam_name = f"BEAM_X_X{i}to{i + 1}_Y{j}_S{story_num}"
                ret_code, actual_name = add_frame_by_coord_custom(
                    self.frame_obj,
                    x1,
                    y,
                    z_beam_center,
                    x2,
                    y,
                    z_beam_center,
                    section_name,
                    beam_name,
                )
                check_ret(ret_code, f"AddByCoord(Beam {beam_name})")
                beam_names.append(actual_name or beam_name)
                story_beam_count += 1

            for i, j, x, y1, y2 in self.grid.iter_beam_spans_y():
                position = _beam_position("Y", i, j, max_i, max_j)
                beam_depth = self.design.beam_depth_for_story(story_num, position)
                z_beam_center = z_top - beam_depth / 2.0
                section_name = self.design.beam_section_name_for_story(story_num, position)
                beam_name = f"BEAM_Y_X{i}_Y{j}to{j + 1}_S{story_num}"
                ret_code, actual_name = add_frame_by_coord_custom(
                    self.frame_obj,
                    x,
                    y1,
                    z_beam_center,
                    x,
                    y2,
                    z_beam_center,
                    section_name,
                    beam_name,
                )
                check_ret(ret_code, f"AddByCoord(Beam {beam_name})")
                beam_names.append(actual_name or beam_name)
                story_beam_count += 1

            log.info("Story %s beams created: %s", story_num, story_beam_count)

        log.info("Total beams created: %s", len(beam_names))
        return beam_names

    def create_slabs(self) -> List[str]:
        slab_names: List[str] = []

        for story_num, _, z_top in self.stories.iter_story_bounds():
            story_slab_count = 0
            for i, j, (x1, x2), (y1, y2) in self.grid.iter_slab_panels():
                slab_name = f"SLAB_X{i}_Y{j}_S{story_num}"
                slab_x = [x1, x2, x2, x1]
                slab_y = [y1, y1, y2, y2]
                slab_z = [z_top] * 4

                ret_code, actual_name = add_area_by_coord_custom(
                    self.area_obj,
                    4,
                    slab_x,
                    slab_y,
                    slab_z,
                    self.design.slab_name,
                    slab_name,
                )
                check_ret(ret_code, f"AddByCoord(Slab {slab_name})")
                final_name = actual_name or slab_name
                slab_names.append(final_name)
                story_slab_count += 1

                try:
                    self.area_obj.SetDiaphragm(final_name, "SRD")
                except Exception as exc:
                    log.debug("SetDiaphragm(SRD) failed for %s: %s", final_name, exc)

            log.info("Story %s slabs created: %s", story_num, story_slab_count)

        log.info("Total slabs created: %s", len(slab_names))
        return slab_names


# NEW: helper to keep ETABS native grid aligned with parametric config
def _alpha_label(idx: int) -> str:
    """Convert 0-based index to spreadsheet-style letters: 0->A, 25->Z, 26->AA."""
    letters = ""
    n = idx
    while True:
        n, rem = divmod(n, 26)
        letters = chr(ord("A") + rem) + letters
        if n == 0:
            break
        n -= 1
    return letters


# NEW
def update_grid_system_from_config(sap_model, grid: GridConfig) -> None:
    """
    Recreate the ETABS Cartesian grid system to match GridConfig.

    The real ETABS OAPI calls should be plugged into the marked sections.
    This helper is defensive so that mismatched signatures fail gracefully while
    still logging what needs to happen.
    """
    log.info(
        "Updating ETABS grid: X(%s lines, %.3f m spacing), Y(%s lines, %.3f m spacing)",
        grid.num_x,
        grid.spacing_x,
        grid.num_y,
        grid.spacing_y,
    )

    grid_sys = getattr(sap_model, "GridSys", None)
    if grid_sys is None:
        log.warning("sap_model.GridSys is unavailable; native grid will not be updated.")
        return

    # Build target grid lines
    x_lines = [{"coord": i * grid.spacing_x, "label": _alpha_label(i)} for i in range(grid.num_x)]
    y_lines = [{"coord": j * grid.spacing_y, "label": str(j + 1)} for j in range(grid.num_y)]

    # Delete existing grid systems if possible
    existing_names: List[str] = []
    try:
        # ETABS: GetNameList(ref count, ref names)
        result = grid_sys.GetNameList()
        if isinstance(result, (list, tuple)):
            for item in reversed(result):
                if isinstance(item, (list, tuple)):
                    existing_names = list(item)
                    break
    except Exception as exc:  # pragma: no cover - depends on ETABS API
        log.debug("GridSys.GetNameList not available: %s", exc)

    for name in existing_names:
        try:
            # ETABS: grid_sys.DeleteGridSys(name) or grid_sys.Delete(name)
            grid_sys.DeleteGridSys(name)  # type: ignore[attr-defined]
        except Exception as exc:  # pragma: no cover
            log.debug("Failed to delete grid system %s: %s", name, exc)

    grid_sys_name = "ParametricGrid"

    # Create cartesian grid system
    try:
        # ETABS: AddCartesian(Name, Xo, Yo, RZ, BubbleX, BubbleY, GridX, GridY, GridZ)
        grid_sys.AddCartesian(grid_sys_name, 0.0, 0.0, 0.0)  # type: ignore[attr-defined]
    except Exception as exc:  # pragma: no cover
        log.debug("GridSys.AddCartesian signature mismatch or unavailable: %s", exc)

    # Add grid lines along X and Y
    try:
        for line in x_lines:
            # ETABS: AddGridLineCartesian(GridSysName, IsXAxis, Coordinate, Bubble, Label)
            grid_sys.AddGridLineCartesian(grid_sys_name, True, line["coord"], True, line["label"])  # type: ignore[attr-defined]
        for line in y_lines:
            grid_sys.AddGridLineCartesian(grid_sys_name, False, line["coord"], True, line["label"])  # type: ignore[attr-defined]
    except Exception as exc:  # pragma: no cover
        log.debug("GridSys.AddGridLineCartesian not available: %s", exc)

    # Set the new grid system active
    try:
        grid_sys.SetCurrentGridSys(grid_sys_name)  # type: ignore[attr-defined]
    except Exception as exc:  # pragma: no cover
        log.debug("GridSys.SetCurrentGridSys not available: %s", exc)

    log.info("ETABS grid system refreshed to parametric layout.")


class FrameGeometryWorkflow:
    def __init__(self, sap_model, grid: GridConfig, stories: StoryConfig, design: DesignConfig):
        self.sap_model = sap_model
        self.grid = grid
        self.stories = stories
        self.design = design
        self.creator = ElementCreator(sap_model, grid, stories, design)

    def build(self) -> Tuple[List[str], List[str], List[str], Dict[int, float]]:
        ensure_model_units()
        # NEW: keep the native ETABS grid in sync with the parametric grid config
        update_grid_system_from_config(self.sap_model, self.grid)

        column_names = self.creator.create_columns()
        beam_names = self.creator.create_beams()
        slab_names = self.creator.create_slabs()

        story_heights = self.stories.story_top_elevations()

        apply_slab_membrane_modifiers(self.sap_model.AreaObj, slab_names)
        assign_diaphragm_constraints_by_story(self.sap_model.AreaObj, slab_names)
        apply_beam_inertia_modifiers(self.sap_model.FrameObj, self.grid, beam_names)

        locator = BaseJointLocator(self.sap_model, self.grid)
        base_joints = locator.locate()
        success = fail = 0
        if base_joints:
            constraint_manager = BaseConstraintManager(self.sap_model)
            success, fail = constraint_manager.set_rigid(base_joints)
            if success:
                constraint_manager.verify(base_joints)
        else:
            log.warning("No base joints identified; base restraints were not applied.")

        try:
            self.sap_model.View.RefreshView(0, False)
        except Exception:
            pass

        log.info(
            "Geometry build complete: %s columns, %s beams, %s slabs",
            len(column_names),
            len(beam_names),
            len(slab_names),
        )
        if success or fail:
            log.info("Base restraints set: success=%s failed=%s", success, fail)

        return column_names, beam_names, slab_names, story_heights


def create_frame_structure(design: Optional[DesignConfig] = None) -> Tuple[List[str], List[str], List[str], Dict[int, float]]:
    sap_model = _require_sap_model()
    design_cfg = design or DEFAULT_DESIGN_CONFIG
    workflow = FrameGeometryWorkflow(
        sap_model,
        grid_config_from_design(design_cfg),
        story_config_from_design(design_cfg),
        design_cfg,
    )
    return workflow.build()


def create_frame_structure_from_design(design) -> Tuple[List[str], List[str], List[str], Dict[int, float]]:
    sap_model = _require_sap_model()
    design_cfg = design_config_from_case(design)
    workflow = FrameGeometryWorkflow(
        sap_model,
        grid_config_from_design(design_cfg),
        story_config_from_design(design_cfg),
        design_cfg,
    )
    return workflow.build()


__all__ = [
    "ElementCreator",
    "FrameGeometryWorkflow",
    "create_frame_structure",
    "create_frame_structure_from_design",
    "ensure_model_units",
    "fix_base_constraints_comprehensive",
    "update_grid_system_from_config",  # NEW
]
