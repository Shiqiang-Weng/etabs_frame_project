#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Materials and section definitions for the frame model."""

from typing import Any, Dict

from common import config
from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret
from common.config import SETTINGS


def define_materials():
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    from common.etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\nDefining concrete material...")
    pm = sap_model.PropMaterial

    check_ret(
        pm.SetMaterial(SETTINGS.sections.concrete_material_name, ETABSv1.eMatType.Concrete),
        f"SetMaterial({SETTINGS.sections.concrete_material_name})",
        (0, 1),
    )
    check_ret(
        pm.SetMPIsotropic(
            SETTINGS.sections.concrete_material_name,
            SETTINGS.sections.concrete_e_modulus,
            SETTINGS.sections.concrete_poisson,
            SETTINGS.sections.concrete_thermal_exp,
        ),
        f"SetMPIsotropic({SETTINGS.sections.concrete_material_name})",
    )
    check_ret(
        pm.SetWeightAndMass(
            SETTINGS.sections.concrete_material_name,
            1,
            SETTINGS.sections.concrete_unit_weight,
        ),
        f"SetWeightAndMass({SETTINGS.sections.concrete_material_name})",
    )
    print(f"Concrete material '{SETTINGS.sections.concrete_material_name}' defined")


def define_frame_sections():
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    print("\nDefining frame sections...")
    pf = sap_model.PropFrame

    check_ret(
        pf.SetRectangle(
            SETTINGS.sections.frame_beam_section_name,
            SETTINGS.sections.concrete_material_name,
            SETTINGS.sections.frame_beam_height,
            SETTINGS.sections.frame_beam_width,
        ),
        f"SetRectangle({SETTINGS.sections.frame_beam_section_name})",
        (0, 1),
    )
    print(
        f"Beam section '{SETTINGS.sections.frame_beam_section_name}' defined "
        f"({SETTINGS.sections.frame_beam_width:.2f}m × {SETTINGS.sections.frame_beam_height:.2f}m)"
    )

    check_ret(
        pf.SetRectangle(
            SETTINGS.sections.frame_column_section_name,
            SETTINGS.sections.concrete_material_name,
            SETTINGS.sections.frame_column_height,
            SETTINGS.sections.frame_column_width,
        ),
        f"SetRectangle({SETTINGS.sections.frame_column_section_name})",
        (0, 1),
    )
    print(
        f"Column section '{SETTINGS.sections.frame_column_section_name}' defined "
        f"({SETTINGS.sections.frame_column_width:.2f}m × {SETTINGS.sections.frame_column_height:.2f}m)"
    )


def define_slab_sections():
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    from common.etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\nDefining slab section...")
    pa = sap_model.PropArea

    check_ret(
        pa.SetSlab(
            SETTINGS.sections.slab_section_name,
            ETABSv1.eSlabType.Slab,
            ETABSv1.eShellType.Membrane,
            SETTINGS.sections.concrete_material_name,
            SETTINGS.sections.slab_thickness,
        ),
        f"SetSlab({SETTINGS.sections.slab_section_name})",
        (0, 1),
    )
    print(
        f"Slab section '{SETTINGS.sections.slab_section_name}' defined "
        f"(thickness: {SETTINGS.sections.slab_thickness:.2f}m, material: {SETTINGS.sections.concrete_material_name})"
    )


def define_diaphragms():
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    from common.etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\nDefining diaphragms...")
    diaphragm_api = sap_model.Diaphragm

    name_rigid = "RIGID"
    name_semi = "SRD"

    num_val = System.Int32(0)
    names_val = System.Array[System.String](0)
    ret_tuple = diaphragm_api.GetNameList(num_val, names_val)
    check_ret(ret_tuple[0], "Diaphragm.GetNameList")

    existing = list(ret_tuple[2]) if ret_tuple[1] > 0 and ret_tuple[2] is not None else []

    if name_rigid not in existing:
        check_ret(
            diaphragm_api.SetDiaphragm(name_rigid, False),
            f"SetDiaphragm({name_rigid})",
        )
    if name_semi not in existing:
        check_ret(
            diaphragm_api.SetDiaphragm(name_semi, True),
            f"SetDiaphragm({name_semi})",
        )

    print("Diaphragms ensured (RIGID & SRD)")


def define_all_materials_and_sections():
    define_materials()
    define_frame_sections()
    define_slab_sections()
    define_diaphragms()
    print("Materials and sections defined")


def create_parametric_frame_sections_from_design(design) -> None:
    """Create rectangle sections for all column/beam groups defined in design."""
    sample_sizing: Dict[str, Dict[str, float]] = design.sizing
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    prop_frame = sap_model.PropFrame

    for group_name, params in sample_sizing.items():
        group_id = group_name.replace("Group", "")

        col_corner = params[f"C_G{group_id}_Corner_b"]
        col_edge = params[f"C_G{group_id}_Edge_b"]
        col_int = params[f"C_G{group_id}_Interior_b"]

        b_edge = params[f"B_G{group_id}_Edge_b"]
        h_edge = params[f"B_G{group_id}_Edge_h"]
        b_int = params[f"B_G{group_id}_Interior_b"]
        h_int = params[f"B_G{group_id}_Interior_h"]

        section_specs = {
            f"C_G{group_id}_CORNER_{int(col_corner)}": (col_corner, col_corner),
            f"C_G{group_id}_EDGE_{int(col_edge)}": (col_edge, col_edge),
            f"C_G{group_id}_INTERIOR_{int(col_int)}": (col_int, col_int),
            f"B_G{group_id}_EDGE_{int(b_edge)}x{int(h_edge)}": (b_edge, h_edge),
            f"B_G{group_id}_INT_{int(b_int)}x{int(h_int)}": (b_int, h_int),
        }

        for name, (b_mm, h_mm) in section_specs.items():
            b = b_mm / 1000.0
            h = h_mm / 1000.0
            ret = prop_frame.SetRectangle(name, config.CONCRETE_MATERIAL_NAME, b, h)
            check_ret(ret, f"SetRectangle({name})")


__all__ = [
    'define_materials',
    'define_frame_sections',
    'define_slab_sections',
    'define_diaphragms',
    'define_all_materials_and_sections',
    'create_parametric_frame_sections_from_design',
]

