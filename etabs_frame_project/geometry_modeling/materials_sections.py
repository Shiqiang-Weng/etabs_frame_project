#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Materials and section definitions for the frame model."""

from typing import Optional

from common.config import DEFAULT_DESIGN_CONFIG, SETTINGS, DesignConfig, design_config_from_case
from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret


def define_materials():
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    from common.etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\nDefining concrete material...")
    pm = sap_model.PropMaterial
    mat_cfg = SETTINGS.materials

    check_ret(
        pm.SetMaterial(mat_cfg.concrete_material_name, ETABSv1.eMatType.Concrete),
        f"SetMaterial({mat_cfg.concrete_material_name})",
        (0, 1),
    )
    check_ret(
        pm.SetMPIsotropic(
            mat_cfg.concrete_material_name,
            mat_cfg.concrete_e_modulus,
            mat_cfg.concrete_poisson,
            mat_cfg.concrete_thermal_exp,
        ),
        f"SetMPIsotropic({mat_cfg.concrete_material_name})",
    )
    check_ret(
        pm.SetWeightAndMass(
            mat_cfg.concrete_material_name,
            1,
            mat_cfg.concrete_unit_weight,
        ),
        f"SetWeightAndMass({mat_cfg.concrete_material_name})",
    )
    print(f"Concrete material '{mat_cfg.concrete_material_name}' defined")


def define_frame_sections(design: Optional[DesignConfig] = None):
    """
    Define all beam/column sections for the provided design (three vertical groups × plan groups).
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    design_cfg = design_config_from_case(design) if design is not None else DEFAULT_DESIGN_CONFIG
    mat_cfg = SETTINGS.materials

    print("\nDefining frame sections from design sampling...")
    pf = sap_model.PropFrame

    for name, width, depth in design_cfg.iter_frame_section_definitions():
        ret = pf.SetRectangle(name, mat_cfg.concrete_material_name, depth, width)
        check_ret(ret, f"SetRectangle({name})", (0, 1))
    print("Frame sections defined from design case %s" % design_cfg.case_id)


def define_slab_sections(design: Optional[DesignConfig] = None):
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    from common.etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    design_cfg = design_config_from_case(design) if design is not None else DEFAULT_DESIGN_CONFIG
    mat_cfg = SETTINGS.materials

    print("\nDefining slab section...")
    pa = sap_model.PropArea

    check_ret(
        pa.SetSlab(
            design_cfg.slab_name,
            ETABSv1.eSlabType.Slab,
            ETABSv1.eShellType.Membrane,
            mat_cfg.concrete_material_name,
            design_cfg.slab_thickness,
        ),
        f"SetSlab({design_cfg.slab_name})",
        (0, 1),
    )
    print(
        f"Slab section '{design_cfg.slab_name}' defined "
        f"(thickness: {design_cfg.slab_thickness:.2f}m, material: {mat_cfg.concrete_material_name})"
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


def define_all_materials_and_sections(design: Optional[DesignConfig] = None):
    design_cfg = design_config_from_case(design) if design is not None else DEFAULT_DESIGN_CONFIG
    define_materials()
    define_frame_sections(design_cfg)
    define_slab_sections(design_cfg)
    define_diaphragms()
    print("Materials and sections defined")


def create_parametric_frame_sections_from_design(design) -> None:
    """Backward-compatible wrapper."""
    design_cfg = design_config_from_case(design)
    define_frame_sections(design_cfg)


__all__ = [
    'define_materials',
    'define_frame_sections',
    'define_slab_sections',
    'define_diaphragms',
    'define_all_materials_and_sections',
    'create_parametric_frame_sections_from_design',
]
