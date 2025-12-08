#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Section assignment based on parametric design mapping."""

from __future__ import annotations

import re
from typing import Any, Dict, Optional, Tuple

from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret
from .api_compat import _get_name_list_safe


COLUMN_RE = re.compile(r"^COL_X(?P<i>\d+)_Y(?P<j>\d+)_S(?P<s>\d+)$")
BEAM_X_RE = re.compile(r"^BEAM_X_X(?P<i>\d+)to(?P<i2>\d+)_Y(?P<j>\d+)_S(?P<s>\d+)$")
BEAM_Y_RE = re.compile(r"^BEAM_Y_X(?P<i>\d+)_Y(?P<j>\d+)to(?P<j2>\d+)_S(?P<s>\d+)$")


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


def _column_section_name(design, group_name: str, position: str) -> Optional[str]:
    gid = _group_id(group_name)
    params = design.sizing.get(group_name, {})
    key = f"C_G{gid}_{position.title()}_b" if position != "INTERIOR" else f"C_G{gid}_Interior_b"
    if position == "EDGE":
        key = f"C_G{gid}_Edge_b"
    elif position == "CORNER":
        key = f"C_G{gid}_Corner_b"
    value = params.get(key)
    if value is None:
        return None
    return f"C_G{gid}_{position}_{int(value)}"


def _beam_section_name(design, group_name: str, position: str) -> Optional[str]:
    gid = _group_id(group_name)
    params = design.sizing.get(group_name, {})
    if position == "EDGE":
        b = params.get(f"B_G{gid}_Edge_b")
        h = params.get(f"B_G{gid}_Edge_h")
        label = "EDGE"
    else:
        b = params.get(f"B_G{gid}_Interior_b")
        h = params.get(f"B_G{gid}_Interior_h")
        label = "INT"
    if b is None or h is None:
        return None
    return f"B_G{gid}_{label}_{int(b)}x{int(h)}"


def assign_sections_by_design(design, topo: Dict[str, Any]) -> None:
    """Assign beam/column sections based on design mapping and topology."""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    frame_obj = sap_model.FrameObj
    names = _get_name_list_safe(frame_obj)
    if not names:
        raise RuntimeError("未能获取框架构件名称列表，无法进行截面分配。")

    assigned = 0
    skipped = 0

    for name in names:
        target_section: Optional[str] = None

        if name.startswith("COL_"):
            parsed = _parse_column(name)
            if not parsed:
                skipped += 1
                continue
            i, j, story = parsed
            group_name = _get_story_group(design.group_mapping, story)
            if not group_name:
                skipped += 1
                continue
            position = _column_position(i, j, topo)
            target_section = _column_section_name(design, group_name, position)
        elif name.startswith("BEAM_"):
            parsed = _parse_beam(name)
            if not parsed:
                skipped += 1
                continue
            orientation, i, j, story = parsed
            group_name = _get_story_group(design.group_mapping, story)
            if not group_name:
                skipped += 1
                continue
            position = _beam_position(orientation, i, j, topo)
            target_section = _beam_section_name(design, group_name, position)
        else:
            continue

        if not target_section:
            skipped += 1
            continue

        ret = frame_obj.SetSection(name, target_section)
        check_ret(ret, f"SetSection({name}->{target_section})", (0, 1))
        assigned += 1

    print(f"[分配截面] 已分配 {assigned} 个构件，跳过 {skipped} 个无法识别的构件")


__all__ = ["assign_sections_by_design"]
