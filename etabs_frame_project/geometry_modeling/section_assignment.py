#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Section assignment based on parametric design mapping."""

from __future__ import annotations

import re
from typing import Any, Dict, Optional, Tuple

from common.config import design_config_from_case
from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret
from .api_compat import _get_name_list_safe


COLUMN_RE = re.compile(r"^COL_X(?P<i>\d+)_Y(?P<j>\d+)_S(?P<s>\d+)$")
BEAM_X_RE = re.compile(r"^BEAM_X_X(?P<i>\d+)to(?P<i2>\d+)_Y(?P<j>\d+)_S(?P<s>\d+)$")
BEAM_Y_RE = re.compile(r"^BEAM_Y_X(?P<i>\d+)_Y(?P<j>\d+)to(?P<j2>\d+)_S(?P<s>\d+)$")
SLAB_RE = re.compile(r"^SLAB_X(?P<i>\d+)_Y(?P<j>\d+)_S(?P<s>\d+)$")


def _parse_column(name: str) -> Optional[Tuple[int, int, int]]:
    m = COLUMN_RE.match(name)
    if not m:
        return None
    return int(m.group("i")), int(m.group("j")), int(m.group("s"))


def _parse_beam(name: str) -> Optional[Tuple[str, int, int, int]]:
    m_x = BEAM_X_RE.match(name)
    if m_x:
        return "X", int(m_x.group("i")), int(m_x.group("j")), int(m_x.group("s"))
    m_y = BEAM_Y_RE.match(name)
    if m_y:
        return "Y", int(m_y.group("i")), int(m_y.group("j")), int(m_y.group("s"))
    return None


def _parse_slab_story(name: str) -> Optional[int]:
    m = SLAB_RE.match(name)
    if not m:
        return None
    return int(m.group("s"))


def _group_id(group_name: str) -> str:
    digits = "".join(ch for ch in str(group_name) if ch.isdigit())
    return digits or str(group_name)


def _get_story_group(mapping: Dict[str, Any], story: int) -> Optional[str]:
    if not mapping:
        return None
    story_to_group = mapping.get("story_to_group") if isinstance(mapping, dict) else None
    if story_to_group:
        return story_to_group.get(story) or story_to_group.get(str(story))
    return mapping.get(story) or mapping.get(str(story))


def _column_position(i: int, j: int, topo: Dict[str, Any]) -> str:
    max_i = topo["n_x"]
    max_j = topo["n_y"]
    if i in {0, max_i} and j in {0, max_j}:
        return "CORNER"
    if i in {0, max_i} or j in {0, max_j}:
        return "EDGE"
    return "INTERIOR"


def _beam_position(orientation: str, i: int, j: int, topo: Dict[str, Any]) -> str:
    if orientation == "X":
        return "EDGE" if j in {0, topo["n_y"]} else "INT"
    return "EDGE" if i in {0, topo["n_x"]} else "INT"


def assign_sections_by_design(design, topo: Optional[Dict[str, Any]] = None) -> None:
    """Assign beam/column sections based on design mapping and topology."""
    design_cfg = design_config_from_case(design)
    topo_data = topo or design_cfg.topology

    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    frame_obj = sap_model.FrameObj
    names = _get_name_list_safe(frame_obj)
    if not names:
        raise RuntimeError("未能获取框架构件名称列表，无法进行截面分配。")

    assigned_frames = 0
    skipped_frames = 0

    for name in names:
        target_section: Optional[str] = None

        if name.startswith("COL_"):
            parsed = _parse_column(name)
            if not parsed:
                skipped_frames += 1
                continue
            i, j, story = parsed
            group_name = _get_story_group(design_cfg.group_mapping, story)
            if not group_name:
                skipped_frames += 1
                continue
            position = _column_position(i, j, topo_data)
            target_section = design_cfg.column_section_name_for_story(story, position)
        elif name.startswith("BEAM_"):
            parsed = _parse_beam(name)
            if not parsed:
                skipped_frames += 1
                continue
            orientation, i, j, story = parsed
            group_name = _get_story_group(design_cfg.group_mapping, story)
            if not group_name:
                skipped_frames += 1
                continue
            position = _beam_position(orientation, i, j, topo_data)
            target_section = design_cfg.beam_section_name_for_story(story, position)
        else:
            continue

        if not target_section:
            skipped_frames += 1
            continue

        ret = frame_obj.SetSection(name, target_section)
        check_ret(ret, f"SetSection({name}->{target_section})", (0, 1))
        assigned_frames += 1

    assigned_slabs = 0
    skipped_slabs = 0
    failed_slabs = 0

    # NEW: assign slab sections by story group (Group1=C40, Group2/3=C35 by config).
    area_obj = getattr(sap_model, "AreaObj", None)
    if area_obj is None:
        skipped_slabs = -1
    else:
        area_names = _get_name_list_safe(area_obj)
        for area_name in area_names:
            if not area_name.startswith("SLAB_"):
                continue
            story = _parse_slab_story(area_name)
            if story is None:
                skipped_slabs += 1
                continue

            target_slab_section = design_cfg.slab_section_name_for_story(story)
            if not target_slab_section:
                skipped_slabs += 1
                continue

            ret_code: Optional[int] = None
            try:
                if hasattr(area_obj, "SetSection"):
                    ret = area_obj.SetSection(area_name, target_slab_section)
                    ret_code = ret[0] if isinstance(ret, tuple) and ret else ret
                elif hasattr(area_obj, "SetProperty"):
                    ret = area_obj.SetProperty(area_name, target_slab_section)
                    ret_code = ret[0] if isinstance(ret, tuple) and ret else ret
                else:
                    skipped_slabs += 1
                    continue
            except Exception:
                failed_slabs += 1
                continue

            if ret_code in (0, 1):
                assigned_slabs += 1
            else:
                failed_slabs += 1

    slab_note = "未分配(AreaObj不可用)" if skipped_slabs == -1 else f"已分配 {assigned_slabs} 个slab，跳过 {skipped_slabs} 个，失败 {failed_slabs} 个"
    print(
        f"[分配截面] Frame: 已分配 {assigned_frames} 个，跳过 {skipped_frames} 个；"
        f" Slab: {slab_note}"
    )


__all__ = ["assign_sections_by_design"]
