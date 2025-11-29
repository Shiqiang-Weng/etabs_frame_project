#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Base joint identification and restraint application utilities.
"""

import logging
from typing import Iterable, List, Sequence, Tuple

from .api_compat import (
    _get_all_points_safe,
    _get_name_list_safe,
    _require_sap_model,
    ensure_model_units,
)
from .layout import GridConfig, default_grid_config

log = logging.getLogger(__name__)
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


class BaseJointLocator:
    def __init__(self, sap_model, grid: GridConfig):
        self.sap_model = sap_model
        self.grid = grid
        self.point_obj = sap_model.PointObj
        self.frame_obj = sap_model.FrameObj

    @staticmethod
    def _dedupe(names: Iterable[str]) -> List[str]:
        seen = set()
        ordered: List[str] = []
        for name in names:
            if name and name not in seen:
                seen.add(name)
                ordered.append(name)
        return ordered

    def by_get_all_points(self, tolerance: float = 0.001) -> List[str]:
        ret, names, _, _, zs = _get_all_points_safe(self.point_obj)
        if ret != 0 or not names:
            return []
        z_min = min(zs)
        base = [name for name, z in zip(names, zs) if abs(z - z_min) <= tolerance]
        return self._dedupe(base)

    def by_existing_columns(self) -> List[str]:
        frame_names = _get_name_list_safe(self.frame_obj)
        if not frame_names:
            return []

        base: List[str] = []
        for name in frame_names:
            if "COL_" not in name or "_S1" not in name:
                continue
            pt1, pt2 = [""], [""]
            ret = self.frame_obj.GetPoints(name, pt1, pt2)
            ret_code = ret[0] if isinstance(ret, tuple) else ret
            if ret_code != 0:
                continue

            coords: List[Tuple[str, float]] = []
            for pt in (pt1[0], pt2[0]):
                if not pt:
                    continue
                x_ref, y_ref, z_ref = [0.0], [0.0], [0.0]
                ret_coord = self.point_obj.GetCoordCartesian(pt, x_ref, y_ref, z_ref)
                code = ret_coord[0] if isinstance(ret_coord, tuple) else ret_coord
                if code == 0:
                    coords.append((pt, z_ref[0]))

            if coords:
                bottom_pt = min(coords, key=lambda item: item[1])[0]
                base.append(bottom_pt)

        return self._dedupe(base)

    def by_grid_lookup(self, tolerance: float = 0.1) -> List[str]:
        base: List[str] = []
        for _, _, x_coord, y_coord in self.grid.iter_points():
            try:
                ret = self.point_obj.GetNameAtCoord(x_coord, y_coord, 0.0, tolerance)
                ret_code = ret[0] if isinstance(ret, tuple) else ret
                joint_name = ret[1] if isinstance(ret, tuple) and len(ret) > 1 else None
                if ret_code == 0 and joint_name:
                    base.append(joint_name)
            except Exception:
                continue
        return self._dedupe(base)

    def locate(self, tolerance: float = 0.001) -> List[str]:
        for strategy in (
            lambda: self.by_get_all_points(tolerance),
            self.by_existing_columns,
            lambda: self.by_grid_lookup(max(tolerance * 10, 0.1)),
        ):
            joints = strategy()
            if joints:
                return joints
        return []


class BaseConstraintManager:
    def __init__(self, sap_model):
        self.point_obj = sap_model.PointObj

    def set_rigid(self, joint_names: Sequence[str]) -> Tuple[int, int]:
        if not joint_names:
            return 0, 0

        restraint = [True, True, True, True, True, True]
        success = failed = 0
        for name in joint_names:
            try:
                ret = self.point_obj.SetRestraint(name, restraint)
                ret_code = ret[0] if isinstance(ret, tuple) else ret
                if ret_code == 0:
                    success += 1
                else:
                    failed += 1
                    log.warning("SetRestraint(%s) returned code %s", name, ret_code)
            except Exception as exc:
                failed += 1
                log.error("SetRestraint(%s) failed: %s", name, exc)

        return success, failed

    def verify(self, joint_names: Sequence[str], sample: int = 5):
        for name in list(joint_names)[:sample]:
            try:
                value = [False] * 6
                ret = self.point_obj.GetRestraint(name, value)
                ret_code = ret[0] if isinstance(ret, tuple) else ret
                restraints = list(ret[1]) if isinstance(ret, tuple) and len(ret) > 1 else value
                if ret_code == 0:
                    log.info("Joint %s restraints: %s", name, restraints)
                else:
                    log.warning("GetRestraint(%s) returned code %s", name, ret_code)
            except Exception as exc:
                log.debug("GetRestraint(%s) failed: %s", name, exc)


def get_base_level_joints(tolerance: float = 0.001) -> List[str]:
    locator = BaseJointLocator(_require_sap_model(), default_grid_config())
    return locator.locate(tolerance)


def get_base_level_joints_v2(tolerance: float = 0.001) -> List[str]:
    locator = BaseJointLocator(_require_sap_model(), default_grid_config())
    return locator.by_get_all_points(tolerance)


def get_base_level_joints_by_grid(tolerance: float = 0.001) -> List[str]:
    locator = BaseJointLocator(_require_sap_model(), default_grid_config())
    joints = locator.by_get_all_points(tolerance)
    if joints:
        return joints
    return locator.by_grid_lookup(max(tolerance * 10, 0.1))


def get_base_level_joints_by_grid_direct(tolerance: float = 0.1) -> List[str]:
    locator = BaseJointLocator(_require_sap_model(), default_grid_config())
    return locator.by_grid_lookup(tolerance)


def get_base_level_joints_by_existing_elements() -> List[str]:
    locator = BaseJointLocator(_require_sap_model(), default_grid_config())
    return locator.by_existing_columns()


def get_base_level_joints_reference_method(tolerance: float = 0.001) -> List[str]:
    return get_base_level_joints_v2(tolerance)


def set_rigid_base_constraints_improved(joint_names: Sequence[str]) -> Tuple[int, int]:
    constraint_manager = BaseConstraintManager(_require_sap_model())
    return constraint_manager.set_rigid(joint_names)


def set_rigid_base_constraints_fixed(joint_names: Sequence[str]) -> Tuple[int, int]:
    return set_rigid_base_constraints_improved(joint_names)


def fix_base_constraints_comprehensive() -> Tuple[int, int]:
    sap_model = _require_sap_model()
    ensure_model_units()

    locator = BaseJointLocator(sap_model, default_grid_config())
    base_joints = locator.locate()

    if not base_joints:
        log.error("Unable to identify base joints; skipping base restraint setup.")
        return 0, 0

    constraint_manager = BaseConstraintManager(sap_model)
    success, fail = constraint_manager.set_rigid(base_joints)
    if success:
        constraint_manager.verify(base_joints)
    return success, fail


def fix_base_constraints_issue() -> Tuple[int, int]:
    return fix_base_constraints_comprehensive()


__all__ = [
    "BaseJointLocator",
    "BaseConstraintManager",
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
]
