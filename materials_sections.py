#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
材料和截面定义模块
定义混凝土材料和框架结构截面属性
"""

from etabs_setup import get_etabs_objects
from utility_functions import check_ret
from config import (
    CONCRETE_MATERIAL_NAME, CONCRETE_E_MODULUS, CONCRETE_POISSON,
    CONCRETE_THERMAL_EXP, CONCRETE_UNIT_WEIGHT,
    FRAME_BEAM_SECTION_NAME, FRAME_COLUMN_SECTION_NAME, SLAB_SECTION_NAME,
    FRAME_BEAM_WIDTH, FRAME_BEAM_HEIGHT, FRAME_COLUMN_WIDTH, FRAME_COLUMN_HEIGHT,
    SLAB_THICKNESS
)


def define_materials():
    """定义材料属性"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    # 重新导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\n定义材料属性...")
    pm = sap_model.PropMaterial

    # 定义混凝土材料
    check_ret(
        pm.SetMaterial(CONCRETE_MATERIAL_NAME, ETABSv1.eMatType.Concrete),
        f"SetMaterial({CONCRETE_MATERIAL_NAME})", (0, 1)
    )

    # 设置各向同性材料属性
    check_ret(
        pm.SetMPIsotropic(CONCRETE_MATERIAL_NAME, CONCRETE_E_MODULUS,
                          CONCRETE_POISSON, CONCRETE_THERMAL_EXP),
        f"SetMPIsotropic({CONCRETE_MATERIAL_NAME})"
    )

    # 设置重度和质量
    check_ret(
        pm.SetWeightAndMass(CONCRETE_MATERIAL_NAME, 1, CONCRETE_UNIT_WEIGHT),
        f"SetWeightAndMass({CONCRETE_MATERIAL_NAME})"
    )

    print(f"混凝土材料 '{CONCRETE_MATERIAL_NAME}' 定义完成")


def define_frame_sections():
    """定义框架截面属性"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    print("\n定义框架截面属性...")
    pf = sap_model.PropFrame

    # 定义框架梁截面
    check_ret(
        pf.SetRectangle(FRAME_BEAM_SECTION_NAME, CONCRETE_MATERIAL_NAME,
                        FRAME_BEAM_HEIGHT, FRAME_BEAM_WIDTH),
        f"SetRectangle({FRAME_BEAM_SECTION_NAME})", (0, 1)
    )
    print(f"框架梁截面 '{FRAME_BEAM_SECTION_NAME}' 定义完成 ({FRAME_BEAM_WIDTH:.2f}m × {FRAME_BEAM_HEIGHT:.2f}m)")

    # 定义框架柱截面
    check_ret(
        pf.SetRectangle(FRAME_COLUMN_SECTION_NAME, CONCRETE_MATERIAL_NAME,
                        FRAME_COLUMN_HEIGHT, FRAME_COLUMN_WIDTH),
        f"SetRectangle({FRAME_COLUMN_SECTION_NAME})", (0, 1)
    )
    print(f"框架柱截面 '{FRAME_COLUMN_SECTION_NAME}' 定义完成 ({FRAME_COLUMN_WIDTH:.2f}m × {FRAME_COLUMN_HEIGHT:.2f}m)")


def define_slab_sections():
    """定义楼板截面属性"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    # 重新导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\n定义楼板截面属性...")
    pa = sap_model.PropArea

    # 定义楼板截面（膜单元）
    check_ret(
        pa.SetSlab(SLAB_SECTION_NAME, ETABSv1.eSlabType.Slab,
                   ETABSv1.eShellType.Membrane, CONCRETE_MATERIAL_NAME, SLAB_THICKNESS),
        f"SetSlab({SLAB_SECTION_NAME})", (0, 1)
    )
    print(f"楼板截面 '{SLAB_SECTION_NAME}' 定义完成 (厚度: {SLAB_THICKNESS:.2f}m, 膜单元)")


def define_diaphragms():
    """定义楼面约束"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    # 重新导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\n定义楼面约束...")
    diaphragm_api = sap_model.Diaphragm

    # 读取已存在的楼面名称
    name_rigid = "RIGID"
    name_semi = "SRD"

    num_val = System.Int32(0)
    names_val = System.Array[System.String](0)

    ret_tuple = diaphragm_api.GetNameList(num_val, names_val)
    check_ret(ret_tuple[0], "Diaphragm.GetNameList")

    existing = list(ret_tuple[2]) if ret_tuple[1] > 0 and ret_tuple[2] is not None else []

    # 刚性楼面
    if name_rigid not in existing:
        check_ret(
            diaphragm_api.SetDiaphragm(name_rigid, False),  # isSemiRigid = False
            f"SetDiaphragm({name_rigid})"
        )

    # 半刚性楼面
    if name_semi not in existing:
        check_ret(
            diaphragm_api.SetDiaphragm(name_semi, True),  # isSemiRigid = True
            f"SetDiaphragm({name_semi})"
        )

    print("楼面约束定义完毕：RIGID(刚性)、SRD(半刚性)")


def define_all_materials_and_sections():
    """定义所有材料和截面"""
    define_materials()
    define_frame_sections()
    define_slab_sections()
    define_diaphragms()
    print("材料和截面定义完毕。")


# 导出函数列表
__all__ = [
    'define_materials',
    'define_frame_sections',
    'define_slab_sections',
    'define_diaphragms',
    'define_all_materials_and_sections'
]