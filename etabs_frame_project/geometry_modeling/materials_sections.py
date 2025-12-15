#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Materials and section definitions for the frame model."""

import re
from typing import Optional

from common.config import (
    CONCRETE_GRADE_BY_GROUP,
    CONCRETE_GRADE_PROPS,
    DEFAULT_CONCRETE_GRADE,
    DEFAULT_DESIGN_CONFIG,
    SETTINGS,
    DesignConfig,
    concrete_material_name_for_grade,
    concrete_material_name_for_group,
    design_config_from_case,
)
from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret


_SECTION_GROUP_RE = re.compile(r"_G(?P<gid>[123])_")


def _group_from_section_name(section_name: str) -> Optional[str]:
    m = _SECTION_GROUP_RE.search(section_name)
    if not m:
        return None
    return f"Group{m.group('gid')}"


def define_materials():
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    from common.etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\nDefining concrete materials...")
    pm = sap_model.PropMaterial
    mat_cfg = SETTINGS.materials

    for grade, props in CONCRETE_GRADE_PROPS.items():
        name = str(props["material_name"])
        e_modulus = float(props["E"])
        fcd = float(props["fcd"])
        ftd = float(props["ftd"])

        check_ret(
            pm.SetMaterial(name, ETABSv1.eMatType.Concrete),
            f"SetMaterial({name})",
            (0, 1),
        )
        check_ret(
            pm.SetMPIsotropic(
                name,
                e_modulus,
                mat_cfg.concrete_poisson,
                mat_cfg.concrete_thermal_exp,
            ),
            f"SetMPIsotropic({name})",
        )
        check_ret(
            pm.SetWeightAndMass(
                name,
                1,
                mat_cfg.concrete_unit_weight,
            ),
            f"SetWeightAndMass({name})",
        )

        # Optional: set concrete strength parameters if the API is available.
        # Different ETABS versions expose different signatures; keep best-effort and never fail hard.
        for api_name in ("SetOConcrete", "SetOConcrete_1", "SetOConcrete_2"):
            api = getattr(pm, api_name, None)
            if api is None:
                continue
            try:
                api(name, fcd, ftd)  # type: ignore[misc]  # COM signature may differ by ETABS version
                break
            except Exception as exc:
                print(f"[DEBUG] PropMaterial.{api_name}({name}, fcd, ftd) failed: {exc}")

        print(f"Concrete material '{name}' defined (grade={grade}, E={e_modulus}, fcd={fcd}, ftd={ftd})")


def define_frame_sections(design: Optional[DesignConfig] = None):
    """
    Define all beam/column sections for the provided design (three vertical groups × plan groups).
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    design_cfg = design_config_from_case(design) if design is not None else DEFAULT_DESIGN_CONFIG
    print("\nDefining frame sections from design sampling...")
    pf = sap_model.PropFrame

    for name, width, depth in design_cfg.iter_frame_section_definitions():
        group_name = _group_from_section_name(name)
        material_name = concrete_material_name_for_group(group_name) if group_name else SETTINGS.materials.concrete_material_name
        ret = pf.SetRectangle(name, material_name, depth, width)
        check_ret(ret, f"SetRectangle({name})", (0, 1))
    print("Frame sections defined from design case %s" % design_cfg.case_id)


def define_slab_sections(design: Optional[DesignConfig] = None):
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    from common.etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    design_cfg = design_config_from_case(design) if design is not None else DEFAULT_DESIGN_CONFIG

    print("\nDefining slab sections...")
    pa = sap_model.PropArea

    grades = {DEFAULT_CONCRETE_GRADE, *set(CONCRETE_GRADE_BY_GROUP.values())}
    for grade in sorted(grades):
        slab_name = design_cfg.slab_section_name_for_grade(grade)
        mat_name = concrete_material_name_for_grade(grade)
        check_ret(
            pa.SetSlab(
                slab_name,
                ETABSv1.eSlabType.Slab,
                ETABSv1.eShellType.Membrane,
                mat_name,
                design_cfg.slab_thickness,
            ),
            f"SetSlab({slab_name})",
            (0, 1),
        )
        print(
            f"Slab section '{slab_name}' defined "
            f"(thickness: {design_cfg.slab_thickness:.2f}m, material: {mat_name})"
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
