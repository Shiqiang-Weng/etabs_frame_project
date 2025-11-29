#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Reusable helpers for geometry modifiers and simple parsing utilities.
"""

import logging
from typing import List, Optional, Sequence

from utility_functions import arr
from .layout import GridConfig

log = logging.getLogger(__name__)
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def apply_slab_membrane_modifiers(area_obj, slab_names: Sequence[str]):
    if not slab_names:
        return

    modifiers_membrane = arr(
        [
            1.0,  # f11
            1.0,  # f22
            1.0,  # f12
            0.0,  # f13
            0.0,  # f23
            0.0,  # f33
            1.0,  # m11
            1.0,  # m22
            1.0,  # m12
            1.0,  # m13
            1.0,  # m23
            1.0,  # m33
            1.0,  # weight
        ]
    )

    success = 0
    failed: List[str] = []

    for name in slab_names:
        try:
            ret = area_obj.SetModifiers(name, modifiers_membrane)
            ret_code = ret[0] if isinstance(ret, tuple) else ret
            if ret_code in (0, 1):
                success += 1
            else:
                failed.append(name)
                log.warning("SetModifiers for slab %s returned code %s", name, ret_code)
        except Exception as exc:
            failed.append(name)
            log.error("SetModifiers for slab %s failed: %s", name, exc)

    log.info("Slab membrane modifiers applied. success=%s failed=%s", success, len(failed))
    if failed:
        log.debug("Slab modifier failures (first 5): %s", failed[:5])


def assign_diaphragm_constraints_by_story(area_obj, slab_names: Sequence[str], diaphragm: str = "D1"):
    if not slab_names:
        return

    success = 0
    failed: List[str] = []

    for slab_name in slab_names:
        try:
            ret_code = area_obj.SetDiaphragm(slab_name, diaphragm)
            if ret_code in (0, 1):
                success += 1
            else:
                failed.append(slab_name)
        except Exception as exc:
            failed.append(slab_name)
            log.debug("SetDiaphragm(%s) failed for %s: %s", diaphragm, slab_name, exc)

    log.info("Diaphragm %s assigned. success=%s failed=%s", diaphragm, success, len(failed))
    if failed:
        log.debug("Diaphragm assignment failures (first 5): %s", failed[:5])


def _parse_axis_index(name: str, axis_prefix: str) -> Optional[int]:
    for part in name.split("_"):
        if part.startswith(axis_prefix) and part[1:].isdigit():
            return int(part[1:])
    return None


def _is_edge_beam(beam_name: str, grid: GridConfig) -> bool:
    if beam_name.startswith("BEAM_X_"):
        y_idx = _parse_axis_index(beam_name, "Y")
        return y_idx in {0, grid.num_y - 1}
    if beam_name.startswith("BEAM_Y_"):
        x_idx = _parse_axis_index(beam_name, "X")
        return x_idx in {0, grid.num_x - 1}
    return False


def apply_beam_inertia_modifiers(frame_obj, grid: GridConfig, beam_names: Sequence[str]):
    if not beam_names:
        return

    modifiers_edge = arr(
        [
            1.0,  # Area
            1.0,  # Shear 2
            1.0,  # Shear 3
            1.0,  # Torsion
            1.0,  # I2
            1.5,  # I3
            1.0,  # Mass
            1.0,  # Weight
        ]
    )
    modifiers_middle = arr(
        [
            1.0,  # Area
            1.0,  # Shear 2
            1.0,  # Shear 3
            1.0,  # Torsion
            1.0,  # I2
            2.0,  # I3
            1.0,  # Mass
            1.0,  # Weight
        ]
    )

    edge = middle = failed = 0
    failed_names: List[str] = []

    for name in beam_names:
        try:
            is_edge = _is_edge_beam(name, grid)
            ret = frame_obj.SetModifiers(name, modifiers_edge if is_edge else modifiers_middle)
            ret_code = ret[0] if isinstance(ret, tuple) else ret
            if ret_code in (0, 1):
                edge += int(is_edge)
                middle += int(not is_edge)
            else:
                failed += 1
                failed_names.append(name)
                log.warning("SetModifiers for beam %s returned code %s", name, ret_code)
        except Exception as exc:
            failed += 1
            failed_names.append(name)
            log.error("SetModifiers for beam %s failed: %s", name, exc)

    log.info("Beam inertia modifiers applied. edge=%s middle=%s failed=%s", edge, middle, failed)
    if failed_names:
        log.debug("Beam modifier failures (first 5): %s", failed_names[:5])


__all__ = [
    "apply_slab_membrane_modifiers",
    "assign_diaphragm_constraints_by_story",
    "apply_beam_inertia_modifiers",
    "_parse_axis_index",
    "_is_edge_beam",
]
