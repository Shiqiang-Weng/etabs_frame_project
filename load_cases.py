#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
荷载工况定义模块
定义静力荷载工况、模态分析工况、反应谱工况等
"""

from etabs_setup import get_etabs_objects
from utility_functions import check_ret, arr
from config import (
    MODAL_CASE_NAME, RS_FUNCTION_NAME, GRAVITY_ACCEL, RS_DAMPING_RATIO,
    GENERATE_RS_COMBOS
)


def ensure_dead_pattern():
    """确保恒荷载模式存在"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    # 动态导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\n确保DEAD荷载模式存在...")
    lp = sap_model.LoadPatterns
    check_ret(lp.Add("DEAD", ETABSv1.eLoadPatternType.Dead, 1.0, True),
              "LoadPatterns.Add(DEAD)", (0, 1))
    check_ret(lp.SetSelfWTMultiplier("DEAD", 1.0),
              "SetSelfWTMultiplier(DEAD)", (0, 1))
    print("DEAD 荷载模式确保存在。")


def ensure_live_pattern():
    """确保活荷载模式存在"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    # 动态导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\n确保LIVE荷载模式存在...")
    lp = sap_model.LoadPatterns
    ret = lp.Add("LIVE", ETABSv1.eLoadPatternType.Live, 0.0, True)
    check_ret(ret, "LoadPatterns.Add(LIVE)", (0, 1))
    if ret == 0:
        print("荷载模式 'LIVE' 已成功添加。")


def define_static_load_cases():
    """定义静力荷载工况"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    # 动态导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\n定义静力荷载工况...")
    static_lc = sap_model.LoadCases.StaticLinear

    for pattern in ["DEAD", "LIVE"]:
        check_ret(static_lc.SetCase(pattern),
                  f"StaticLinear.SetCase({pattern})", (0, 1))
        check_ret(
            static_lc.SetLoads(pattern, 1, arr(["Load"], System.String),
                               arr([pattern], System.String), arr([1.0])),
            f"StaticLinear.SetLoads({pattern})"
        )
    print("静力荷载工况定义完毕。")


def define_modal_case():
    """定义模态分析工况"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    # 动态导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print(f"\n定义模态分析工况 '{MODAL_CASE_NAME}'...")
    lc = sap_model.LoadCases
    mod_api = lc.ModalEigen

    # 检查是否已存在模态工况
    num_val = System.Int32(0)
    names_val = System.Array[System.String](0)
    ret_tuple = lc.GetNameList(num_val, names_val, ETABSv1.eLoadCaseType.Modal)
    check_ret(ret_tuple[0], "GetNameList(Modal)")
    existing_modals = list(ret_tuple[2]) if ret_tuple[1] > 0 and ret_tuple[2] is not None else []

    if MODAL_CASE_NAME not in existing_modals:
        check_ret(mod_api.SetCase(MODAL_CASE_NAME), f"ModalEigen.SetCase({MODAL_CASE_NAME})")

    # 设置特征值求解器
    if hasattr(mod_api, "SetEigenSolver"):
        check_ret(mod_api.SetEigenSolver(MODAL_CASE_NAME, 0), f"SetEigenSolver({MODAL_CASE_NAME})")
    elif hasattr(mod_api, "SetModalSolverOption"):  # 旧版API
        check_ret(mod_api.SetModalSolverOption(MODAL_CASE_NAME, 0), f"SetModalSolverOption({MODAL_CASE_NAME})")

    # 设置模态数量
    check_ret(mod_api.SetNumberModes(MODAL_CASE_NAME, 60, 1),
              f"SetNumberModes({MODAL_CASE_NAME})", (0, 1))

    print(f"模态分析工况 '{MODAL_CASE_NAME}' 定义完成（60个模态）")


def define_response_spectrum_cases():
    """定义反应谱工况"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return []

    # 动态导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\n定义反应谱工况...")
    rs_api = sap_model.LoadCases.ResponseSpectrum

    rs_cases_created = []
    for direction_label, u_dir_code in [("X", "U1"), ("Y", "U2")]:
        case_name = f"RS-{direction_label}"

        # 设置反应谱工况
        check_ret(rs_api.SetCase(case_name), f"RS.SetCase({case_name})", (0, 1))

        # 设置荷载
        check_ret(
            rs_api.SetLoads(case_name, 1, arr([u_dir_code], System.String),
                            arr([RS_FUNCTION_NAME], System.String),
                            arr([GRAVITY_ACCEL]), arr(["Global"], System.String),
                            arr([0.0])),
            f"RS.SetLoads({case_name})"
        )

        # 设置模态工况
        check_ret(rs_api.SetModalCase(case_name, MODAL_CASE_NAME),
                  f"RS.SetModalCase({case_name})")

        # 设置模态组合方法（CQC）
        if hasattr(rs_api, "SetModalComb"):
            check_ret(rs_api.SetModalComb(case_name, 0, 0.0, RS_DAMPING_RATIO, 0),
                      f"RS.SetModalComb({case_name})")

        # 设置缺失质量
        if hasattr(rs_api, "SetMissingMass"):
            check_ret(rs_api.SetMissingMass(case_name, True),
                      f"RS.SetMissingMass({case_name})")

        # 设置偶然偏心
        if hasattr(rs_api, "SetAccidentalEccen"):
            check_ret(rs_api.SetAccidentalEccen(case_name, 0.05, True, True),
                      f"RS.SetAccidentalEccen({case_name})")

        rs_cases_created.append(case_name)
        print(f"反应谱工况 '{case_name}' 定义完毕。")

    return rs_cases_created


def define_response_spectrum_combinations(rs_cases_created):
    """定义反应谱组合"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None or not GENERATE_RS_COMBOS or len(rs_cases_created) != 2:
        return

    # 动态导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\n定义地震效应组合...")
    combo_api = sap_model.RespCombo
    rs_ex, rs_ey = rs_cases_created[0], rs_cases_created[1]

    combos = [
        (f"E1_{rs_ex}_p03{rs_ey}", [(rs_ex, 1.0), (rs_ey, 0.3)]),
        (f"E2_{rs_ex}_m03{rs_ey}", [(rs_ex, 1.0), (rs_ey, -0.3)]),
        (f"E3_p03{rs_ex}_{rs_ey}", [(rs_ex, 0.3), (rs_ey, 1.0)]),
        (f"E4_m03{rs_ex}_{rs_ey}", [(rs_ex, -0.3), (rs_ey, 1.0)])
    ]

    for name, case_sfs in combos:
        check_ret(combo_api.Add(name, 0), f"RespCombo.Add({name})", (0, 1))  # 0 for SRSS
        for case, sf in case_sfs:
            check_ret(
                combo_api.SetCaseList(name, ETABSv1.eCNameType.LoadCase,
                                      System.String(case), System.Double(sf)),
                f"RespCombo.SetCaseList({name})"
            )
    print("地震效应组合定义完毕。")


def define_mass_source_simple():
    """简化版质量源定义"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        print("错误: SapModel 未初始化, 无法定义质量源。")
        return

    # 动态导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    print("\n定义质量源（简化版）...")

    # 质量源参数设置
    load_pattern_names = ["DEAD", "LIVE"]
    scale_factors = [1.0, 0.5]

    # 将Python列表转换为.NET数组
    load_pattern_names_api = arr(load_pattern_names, System.String)
    scale_factors_api = arr(scale_factors, System.Double)

    print(f"  荷载模式: {load_pattern_names}")
    print(f"  系数: {scale_factors}")

    try:
        pm = sap_model.PropMaterial

        # 使用PropMaterial.SetMassSource_1设置基本质量源
        ret = pm.SetMassSource_1(
            False,  # includeElementsMass: 不包含元素自重
            True,  # includeAdditionalMass: 包含附加质量
            True,  # includeLoads: 包含指定荷载
            len(load_pattern_names),  # 荷载模式数量
            load_pattern_names_api,  # 荷载模式名称数组
            scale_factors_api  # 荷载系数数组
        )

        check_ret(ret, f"PropMaterial.SetMassSource_1", (0, 1))
        print("✅ 质量源设置成功")

        print(f"\n--- 质量源定义完成 ---")
        print(f"DEAD荷载质量系数: 1.0, LIVE荷载质量系数: 0.5")
        print("--- 质量源定义完毕 ---")

    except Exception as e:
        print(f"❌ 质量源设置失败: {e}")
        print("建议: 请手动在ETABS界面中设置质量源")


def define_all_load_cases():
    """定义所有荷载工况"""
    ensure_dead_pattern()
    ensure_live_pattern()
    define_mass_source_simple()
    define_static_load_cases()
    define_modal_case()
    rs_cases = define_response_spectrum_cases()
    define_response_spectrum_combinations(rs_cases)
    print("荷载工况定义完毕。")


# 导出函数列表
__all__ = [
    'ensure_dead_pattern',
    'ensure_live_pattern',
    'define_static_load_cases',
    'define_modal_case',
    'define_response_spectrum_cases',
    'define_response_spectrum_combinations',
    'define_mass_source_simple',
    'define_all_load_cases'
]