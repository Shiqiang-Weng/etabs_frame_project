#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Materials and section definitions for the frame model.
"""

from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret
from common.config import SETTINGS


def define_materials():
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    from etabs_api_loader import get_api_objects
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

    from etabs_api_loader import get_api_objects
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

    from etabs_api_loader import get_api_objects
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


__all__ = [
    'define_materials',
    'define_frame_sections',
    'define_slab_sections',
    'define_diaphragms',
    'define_all_materials_and_sections',
]
