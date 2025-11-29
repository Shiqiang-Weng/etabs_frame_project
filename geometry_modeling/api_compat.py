#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETABS API compatibility helpers for geometry modeling.

- Ensure SapModel availability and kN-m units.
- Version-tolerant wrappers for GetAllPoints / GetNameList.
- Small debug helpers for joint coordinates.
"""

import logging
from typing import List, Tuple

from etabs_setup import get_etabs_objects

log = logging.getLogger(__name__)
if not log.handlers:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def _require_sap_model():
    _, sap_model = get_etabs_objects()
    if sap_model is None:
        raise RuntimeError("SapModel is not initialized. Run ETABS setup before modeling geometry.")
    return sap_model


def ensure_model_units() -> bool:
    sap_model = _require_sap_model()
    try:
        current_units = sap_model.GetPresentUnits()
        KNM_ENUM = 6  # ETABS eUnits.kN_m_C

        if current_units == KNM_ENUM:
            log.info("Model units already set to kN-m.")
            return True

        from etabs_api_loader import get_api_objects

        ETABSv1, _, _ = get_api_objects()
        ret_code = sap_model.SetPresentUnits(ETABSv1.eUnits.kN_m_C)

        if ret_code == 0:
            log.info("Model units set to kN-m.")
            return True

        log.warning("Setting model units returned code %s", ret_code)
        return False
    except Exception as exc:
        log.error("Failed to check or set model units: %s", exc)
        return False


def _get_all_points_safe(point_obj, csys: str = "Global") -> Tuple[int, List[str], List[float], List[float], List[float]]:
    """
    Version-tolerant wrapper for PointObj.GetAllPoints.
    Returns (ret_code, names, xs, ys, zs). ret_code==0 indicates success.
    """
    try:
        ret, names, xs, ys, zs = point_obj.GetAllPoints(csys)
        if ret == 0 and names:
            return ret, list(names), list(xs), list(ys), list(zs)
    except TypeError:
        log.debug("GetAllPoints(csys) signature mismatch, falling back to ByRef style.")
    except Exception as exc:
        log.debug("GetAllPoints(csys) call failed: %s", exc, exc_info=True)

    try:
        from etabs_api_loader import get_api_objects

        _, System, _ = get_api_objects()
        n_max = 20000
        num_dummy = System.Int32(0)
        names_arr = System.Array[System.String]([None] * n_max)
        X = System.Array[float]([0.0] * n_max)
        Y = System.Array[float]([0.0] * n_max)
        Z = System.Array[float]([0.0] * n_max)

        ret = point_obj.GetAllPoints(num_dummy, names_arr, X, Y, Z, csys)

        if isinstance(ret, tuple):
            ret_code, count = ret[0], ret[1]
            if len(ret) > 2:
                names_arr, X, Y, Z = ret[2:6]
        else:
            ret_code = ret
            count = int(num_dummy)

        if ret_code == 0 and count > 0:
            return (
                ret_code,
                list(names_arr)[:count],
                list(X)[:count],
                list(Y)[:count],
                list(Z)[:count],
            )

        log.warning("GetAllPoints (ByRef) returned code %s with count %s", ret_code, count)
    except ImportError:
        log.error("etabs_api_loader not available; cannot perform ByRef GetAllPoints.")
    except Exception as exc:
        log.error("GetAllPoints compatibility call failed: %s", exc, exc_info=True)

    return 1, [], [], [], []


def _get_name_list_safe(obj) -> List[str]:
    """
    Version-tolerant wrapper for GetNameList on ETABS objects.
    """
    try:
        ret, names = obj.GetNameList()
        if ret == 0:
            return list(names)
    except TypeError:
        log.debug("GetNameList() signature mismatch, falling back to ByRef style.")
    except Exception as exc:
        log.debug("GetNameList() call failed: %s", exc, exc_info=True)

    try:
        from etabs_api_loader import get_api_objects

        _, System, _ = get_api_objects()
        n_max = 50000
        num_dummy = System.Int32(0)
        name_arr = System.Array[System.String]([None] * n_max)

        ret = obj.GetNameList(num_dummy, name_arr)

        if isinstance(ret, tuple):
            ret_code, count = ret[0], ret[1]
            if len(ret) > 2:
                name_arr = ret[2]
        else:
            ret_code = ret
            count = int(num_dummy)

        if ret_code == 0:
            return list(name_arr)[:count]

        log.warning("GetNameList (ByRef) returned code %s", ret_code)
    except ImportError:
        log.error("etabs_api_loader not available; cannot perform ByRef GetNameList.")
    except Exception as exc:
        log.error("GetNameList compatibility call failed: %s", exc, exc_info=True)

    return []


def debug_joint_coordinates(max_joints: int = 10):
    sap_model = _require_sap_model()
    ret, pt_names, pt_x, pt_y, pt_z = _get_all_points_safe(sap_model.PointObj)
    if ret != 0 or not pt_names:
        log.warning("No joints available for debug output.")
        return

    log.info("Sample joint coordinates (up to %s joints):", max_joints)
    for idx, name in enumerate(pt_names[:max_joints]):
        log.info("  %s: (%.3f, %.3f, %.3f)", name, pt_x[idx], pt_y[idx], pt_z[idx])


def get_all_points_reference_method(include_restraints: bool = False) -> List[tuple]:
    sap_model = _require_sap_model()
    point_obj = sap_model.PointObj
    ret, pt_names, pt_x, pt_y, pt_z = _get_all_points_safe(point_obj)
    if ret != 0 or not pt_names:
        return []

    points: List[tuple] = []
    for idx, name in enumerate(pt_names):
        data: Tuple = (name, pt_x[idx], pt_y[idx], pt_z[idx])
        if include_restraints:
            restraints = [False] * 6
            try:
                ret_res = point_obj.GetRestraint(name, restraints)
                ret_code = ret_res[0] if isinstance(ret_res, tuple) else ret_res
                if isinstance(ret_res, tuple) and len(ret_res) > 1:
                    restraints = list(ret_res[1])
                if ret_code != 0:
                    restraints = [False] * 6
            except Exception:
                restraints = [False] * 6
            data = data + (restraints,)
        points.append(data)
    return points


__all__ = [
    "_require_sap_model",
    "ensure_model_units",
    "_get_all_points_safe",
    "_get_name_list_safe",
    "debug_joint_coordinates",
    "get_all_points_reference_method",
]
