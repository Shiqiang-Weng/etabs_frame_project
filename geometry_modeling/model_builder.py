#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Geometry modeling workflow for ETABS frame structures.
"""

import logging
from typing import Dict, List, Tuple

from common.config import (
    FRAME_BEAM_SECTION_NAME,
    FRAME_COLUMN_SECTION_NAME,
    SLAB_SECTION_NAME,
)
from common.etabs_setup import get_etabs_objects
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
from .layout import GridConfig, StoryConfig, default_grid_config, default_story_config

log = logging.getLogger(__name__)
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


class ElementCreator:
    def __init__(self, sap_model, grid: GridConfig, stories: StoryConfig):
        self.sap_model = sap_model
        self.grid = grid
        self.stories = stories
        self.frame_obj = sap_model.FrameObj
        self.area_obj = sap_model.AreaObj

    def create_columns(self) -> List[str]:
        column_names: List[str] = []
        for story_num, z_bottom, z_top in self.stories.iter_story_bounds():
            story_count = 0
            for i, j, x_coord, y_coord in self.grid.iter_points():
                column_name = f"COL_X{i}_Y{j}_S{story_num}"
                ret_code, actual_name = add_frame_by_coord_custom(
                    self.frame_obj,
                    x_coord,
                    y_coord,
                    z_bottom,
                    x_coord,
                    y_coord,
                    z_top,
                    FRAME_COLUMN_SECTION_NAME,
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
        beam_half_height = self.stories.beam_height / 2.0

        for story_num, _, z_top in self.stories.iter_story_bounds():
            z_beam_center = z_top - beam_half_height

            story_beam_count = 0

            for i, j, x1, x2, y in self.grid.iter_beam_spans_x():
                beam_name = f"BEAM_X_X{i}to{i + 1}_Y{j}_S{story_num}"
                ret_code, actual_name = add_frame_by_coord_custom(
                    self.frame_obj,
                    x1,
                    y,
                    z_beam_center,
                    x2,
                    y,
                    z_beam_center,
                    FRAME_BEAM_SECTION_NAME,
                    beam_name,
                )
                check_ret(ret_code, f"AddByCoord(Beam {beam_name})")
                beam_names.append(actual_name or beam_name)
                story_beam_count += 1

            for i, j, x, y1, y2 in self.grid.iter_beam_spans_y():
                beam_name = f"BEAM_Y_X{i}_Y{j}to{j + 1}_S{story_num}"
                ret_code, actual_name = add_frame_by_coord_custom(
                    self.frame_obj,
                    x,
                    y1,
                    z_beam_center,
                    x,
                    y2,
                    z_beam_center,
                    FRAME_BEAM_SECTION_NAME,
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
                    SLAB_SECTION_NAME,
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


class FrameGeometryWorkflow:
    def __init__(self, sap_model, grid: GridConfig, stories: StoryConfig):
        self.sap_model = sap_model
        self.grid = grid
        self.stories = stories
        self.creator = ElementCreator(sap_model, grid, stories)

    def build(self) -> Tuple[List[str], List[str], List[str], Dict[int, float]]:
        ensure_model_units()

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


def create_frame_structure() -> Tuple[List[str], List[str], List[str], Dict[int, float]]:
    sap_model = _require_sap_model()
    workflow = FrameGeometryWorkflow(sap_model, default_grid_config(), default_story_config())
    return workflow.build()


__all__ = [
    "ElementCreator",
    "FrameGeometryWorkflow",
    "create_frame_structure",
    "ensure_model_units",
    "fix_base_constraints_comprehensive",
]
