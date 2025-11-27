#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析结果模块
提取模态信息、质量参与系数、层间位移角，并将输出写入 Excel。
"""

from __future__ import annotations

import io
import sys
import traceback
from contextlib import contextmanager
from pathlib import Path
from typing import List, Union

import pandas as pd

from config import MODAL_CASE_NAME
from utility_functions import check_ret


@contextmanager
def capture_stdout_to_buffer():
    """临时捕获 stdout 到内存缓冲区。"""
    old_stdout = sys.stdout
    buffer = io.StringIO()
    sys.stdout = buffer
    try:
        yield buffer
    finally:
        sys.stdout = old_stdout


def _is_important_line(line: str) -> bool:
    line = (line or "").strip()
    if not line:
        return False
    if any(ch.isdigit() for ch in line):
        return True
    keywords = [
        "开始提取模态信息和质量参与系数",
        "模态周期和频率",
        "模态参与质量系数",
        "最终累积质量参与系数",
        "开始提取相对层间位移角",
        "层间位移角提取完毕",
        "T2/T1 =",
        "T3/T2 =",
        "SumUX:",
        "SumUY:",
        "SumUZ:",
        "SumRX:",
        "SumRY:",
        "SumRZ:",
        "=== 最大层间位移角总结 ===",
        "最大位移角:",
        "位置:",
        "满足建议限值",
        "⚠️  警告",
    ]
    return any(k in line for k in keywords)


def extract_modal_and_mass_info(sap_model) -> None:
    """提取模态信息和质量参与系数 - 改进版，增强错误处理"""
    if sap_model is None or not hasattr(sap_model, "Results") or sap_model.Results is None:
        print("错误: 结果不可用，无法提取模态信息。")
        return

    # 动态导入API对象
    from etabs_api_loader import get_api_objects

    ETABSv1, System, COMException = get_api_objects()

    if System is None:
        print("错误: System 模块未正确加载，无法提取模态信息。")
        return

    print(f"\n--- 开始提取模态信息和质量参与系数 ---")
    results_api = sap_model.Results
    setup_api = results_api.Setup

    # 先检查模态工况是否存在分析结果
    print("检查模态分析结果可用性...")

    try:
        # 尝试获取当前选中的工况 - 修正参数传递方式
        num_val = System.Int32(0)
        names_val = System.Array[System.String](0)

        # 使用引用传递方式
        ret_code = setup_api.GetCaseSelectedForOutput(num_val, names_val)

        # 检查返回值 - 注意：num_val 和 names_val 是通过引用修改的
        if ret_code == 0 and num_val.Value > 0:
            # 获取实际的数组内容
            selected_cases = list(names_val) if names_val is not None else []
            print(f"当前已选择输出的工况: {selected_cases}")
        else:
            print("未检测到已选择的输出工况")

    except Exception as e:
        print(f"检查选中工况时出错: {e}")
        print("将跳过工况检查，直接设置模态工况...")

    # 重新选择模态工况
    print(f"重新选择模态工况 '{MODAL_CASE_NAME}' 进行结果输出...")
    check_ret(setup_api.DeselectAllCasesAndCombosForOutput(), "DeselectAllCasesForModal", (0, 1))
    check_ret(
        setup_api.SetCaseSelectedForOutput(MODAL_CASE_NAME),
        f"SetCaseSelectedForModal({MODAL_CASE_NAME})",
        (0, 1),
    )

    # --- 模态周期和频率 (改进错误处理) ---
    print("\n--- 模态周期和频率 ---")
    _Num_MP, _LC_MP, _ST_MP, _SN_MP, _P_MP, _F_MP, _CF_MP, _EV_MP = (
        System.Int32(0),
        System.Array[System.String](0),
        System.Array[System.String](0),
        System.Array[System.Double](0),
        System.Array[System.Double](0),
        System.Array[System.Double](0),
        System.Array[System.Double](0),
        System.Array[System.Double](0),
    )
    try:
        mp_res = results_api.ModalPeriod(_Num_MP, _LC_MP, _ST_MP, _SN_MP, _P_MP, _F_MP, _CF_MP, _EV_MP)
        ret_code = check_ret(mp_res[0], "Results.ModalPeriod", (0, 1))  # 允许返回0或1

        if ret_code == 1:
            print("  提示: 模态周期结果可能不完整或无数据，但将尝试继续处理...")

        num_m, p_val = mp_res[1], list(mp_res[5]) if mp_res[5] is not None else []

        if num_m > 0 and p_val:
            print(f"  找到 {num_m} 个模态，显示前10个:")
            print(f"{'振型号':<5} {'周期 (s)':<12} {'频率 (Hz)':<12} {'周期比':<10}")
            print("-" * 40)
            for i in range(min(num_m, 10)):  # Display first 10 modes
                T_curr = p_val[i]
                freq_curr = 1.0 / T_curr if T_curr > 0 else 0
                p_ratio_str = f"{p_val[i] / p_val[i - 1]:.3f}" if i > 0 and p_val[i - 1] != 0 else "-"
                print(f"{i + 1:<5} {T_curr:<12.4f} {freq_curr:<12.4f} {p_ratio_str:<10}")

            # 分析前几个周期的比值
            if num_m >= 2 and len(p_val) >= 2:
                t1, t2 = p_val[0], p_val[1]
                r_t21 = t2 / t1 if t1 != 0 else 0
                print(
                    f"\nT2/T1 = {t2:.4f}/{t1:.4f} = {r_t21:.3f} {'⚠️ <0.85 (扭转耦联可能显著)' if r_t21 < 0.85 and t1 != 0 else ''}"
                )
            if num_m >= 3 and len(p_val) >= 3:
                t3 = p_val[2]
                t2_for_ratio = p_val[1]
                r_t32 = t3 / t2_for_ratio if t2_for_ratio != 0 else 0
                print(
                    f"T3/T2 = {t3:.4f}/{t2_for_ratio:.4f} = {r_t32:.3f} {'⚠️ <0.85 (扭转耦联可能显著)' if r_t32 < 0.85 and t2_for_ratio != 0 else ''}"
                )
        else:
            print("  未找到模态周期结果或数据为空。")
            print("  可能原因: 1) 模态分析未完成 2) 模态工况未正确定义 3) 结构质量分布问题")

    except Exception as e:
        print(f"  提取模态周期时发生错误: {e}")
        print("  将跳过模态周期分析，继续尝试质量参与系数...")

    # --- 模态参与质量系数 (改进错误处理) ---
    print("\n--- 模态参与质量系数 ---")
    (
        _N_MPMR,
        _LC_MPMR,
        _ST_MPMR,
        _SN_MPMR,
        _P_MPMR,
        _UX,
        _UY,
        _UZ,
        _S_UX_API,
        _S_UY_API,
        _S_UZ_API,
        _RX,
        _RY,
        _RZ,
        _S_RX_API,
        _S_RY_API,
        _S_RZ_API,
    ) = (
        System.Int32(0),
        *[System.Array[System.String](0)] * 2,
        *[System.Array[System.Double](0)] * 2,
        *[System.Array[System.Double](0)] * 12,
    )  # 17 parameters total
    try:
        mpmr_res = results_api.ModalParticipatingMassRatios(
            _N_MPMR,
            _LC_MPMR,
            _ST_MPMR,
            _SN_MPMR,
            _P_MPMR,
            _UX,
            _UY,
            _UZ,
            _S_UX_API,
            _S_UY_API,
            _S_UZ_API,
            _RX,
            _RY,
            _RZ,
            _S_RX_API,
            _S_RY_API,
            _S_RZ_API,
        )
        ret_code = check_ret(mpmr_res[0], "ModalParticipatingMassRatios", (0, 1))  # 允许返回0或1

        if ret_code == 1:
            print("  提示: 质量参与系数结果可能不完整或无数据，但将尝试继续处理...")

        if len(mpmr_res) < 18:
            print(
                f"  警告: ModalParticipatingMassRatios 返回了 {len(mpmr_res)} 个值，预期为 18。API 可能已更改，请检查参数顺序！"
            )
            return

        num_m_mpmr = mpmr_res[1]
        period_val = list(mpmr_res[5]) if mpmr_res[5] is not None else []
        ux_val = list(mpmr_res[6]) if mpmr_res[6] is not None else []
        uy_val = list(mpmr_res[7]) if mpmr_res[7] is not None else []
        uz_val = list(mpmr_res[8]) if mpmr_res[8] is not None else []
        sum_ux_val = list(mpmr_res[9]) if mpmr_res[9] is not None else []
        sum_uy_val = list(mpmr_res[10]) if mpmr_res[10] is not None else []
        sum_uz_val = list(mpmr_res[11]) if mpmr_res[11] is not None else []
        rx_val = list(mpmr_res[12]) if mpmr_res[12] is not None else []
        ry_val = list(mpmr_res[13]) if mpmr_res[13] is not None else []
        rz_val = list(mpmr_res[14]) if mpmr_res[14] is not None else []
        sum_rx_val = list(mpmr_res[15]) if mpmr_res[15] is not None else []
        sum_ry_val = list(mpmr_res[16]) if mpmr_res[16] is not None else []
        sum_rz_val = list(mpmr_res[17]) if mpmr_res[17] is not None else []

        all_lists = [
            period_val,
            ux_val,
            uy_val,
            uz_val,
            sum_ux_val,
            sum_uy_val,
            sum_uz_val,
            rx_val,
            ry_val,
            rz_val,
            sum_rx_val,
            sum_ry_val,
            sum_rz_val,
        ]

        if num_m_mpmr > 0 and all(all_lists):
            print(f"  找到 {num_m_mpmr} 个模态的质量参与系数，显示前15个:")
            print(
                f"{'振型号':<5} {'周期(s)':<10} {'UX':<8} {'UY':<8} {'UZ':<8} {'RX':<8} {'RY':<8} {'RZ':<8} | {'SumUX':<8} {'SumUY':<8} {'SumUZ':<8} {'SumRX':<8} {'SumRY':<8} {'SumRZ':<8}"
            )
            print("-" * 130)

            for i in range(min(num_m_mpmr, 15)):
                T_c = period_val[i]
                print(
                    f"{i + 1:<5} {T_c:<10.4f} "
                    f"{ux_val[i]:<8.4f} {uy_val[i]:<8.4f} {uz_val[i]:<8.4f} "
                    f"{rx_val[i]:<8.4f} {ry_val[i]:<8.4f} {rz_val[i]:<8.4f} | "
                    f"{sum_ux_val[i]:<8.4f} {sum_uy_val[i]:<8.4f} {sum_uz_val[i]:<8.4f} "
                    f"{sum_rx_val[i]:<8.4f} {sum_ry_val[i]:<8.4f} {sum_rz_val[i]:<8.4f}"
                )

            final_sum_ux = sum_ux_val[-1]
            final_sum_uy = sum_uy_val[-1]
            final_sum_uz = sum_uz_val[-1]
            final_sum_rx = sum_rx_val[-1]
            final_sum_ry = sum_ry_val[-1]
            final_sum_rz = sum_rz_val[-1]

            print("\n--- 最终累积质量参与系数 ---")
            min_ratio = 0.90
            print(f"SumUX: {final_sum_ux:.3f} {'(OK)' if final_sum_ux >= min_ratio else f'(⚠️ < {min_ratio})'}")
            print(f"SumUY: {final_sum_uy:.3f} {'(OK)' if final_sum_uy >= min_ratio else f'(⚠️ < {min_ratio})'}")
            print(f"SumUZ: {final_sum_uz:.3f} {'(OK)' if final_sum_uz >= min_ratio else f'(⚠️ < {min_ratio})'}")
            print(f"SumRX: {final_sum_rx:.3f} {'(OK)' if final_sum_rx >= min_ratio else f'(⚠️ < {min_ratio})'}")
            print(f"SumRY: {final_sum_ry:.3f} {'(OK)' if final_sum_ry >= min_ratio else f'(⚠️ < {min_ratio})'}")
            print(f"SumRZ: {final_sum_rz:.3f} {'(OK)' if final_sum_rz >= min_ratio else f'(⚠️ < {min_ratio})'}")
        else:
            print("  未找到模态参与质量系数结果或数据不完整。")
            print("  可能原因: 1) 模态分析未完成 2) 质量源未正确定义 3) 结构边界条件问题")

    except Exception as e:
        print(f"  提取模态参与质量系数时发生错误: {e}")
        print("  将跳过质量参与系数分析...")
        traceback.print_exc()

    print("--- 模态信息和质量参与系数提取完毕 ---")


def extract_story_drifts_improved(sap_model, target_load_cases: List[str]) -> None:
    """提取层间位移角 - 改进版，增强错误处理和诊断"""
    if sap_model is None:
        print("错误: SapModel 未初始化, 无法提取层间位移角。")
        return
    if not hasattr(sap_model, "Results") or sap_model.Results is None:
        print("错误: 无法访问分析结果 (sap_model.Results is None)。模型可能未分析或结果不可用。")
        return

    # 动态导入API对象
    from etabs_api_loader import get_api_objects

    ETABSv1, System, COMException = get_api_objects()

    if System is None:
        print("错误: System 模块未正确加载，无法提取层间位移角。")
        return

    print(f"\n--- 开始提取相对层间位移角 ({', '.join(target_load_cases)}) ---")

    results_api = sap_model.Results
    setup_api = results_api.Setup

    # 先检查目标工况是否有结果
    print("检查目标工况的分析结果可用性...")

    # 获取所有可用的工况 - 修正参数传递方式
    try:
        num_val = System.Int32(0)
        names_val = System.Array[System.String](0)

        # 使用引用传递方式
        ret_code = sap_model.LoadCases.GetNameList(num_val, names_val)

        if ret_code == 0 and num_val.Value > 0:
            all_cases = list(names_val) if names_val is not None else []
            print(f"模型中定义的所有工况: {all_cases}")

            # 检查目标工况是否存在
            missing_cases = [case for case in target_load_cases if case not in all_cases]
            if missing_cases:
                print(f"警告: 以下工况未在模型中定义: {missing_cases}")
                target_load_cases = [case for case in target_load_cases if case in all_cases]
                if not target_load_cases:
                    print("错误: 没有有效的目标工况可以提取位移角。")
                    return

            print(f"将提取以下工况的位移角: {target_load_cases}")
        else:
            print("警告: 无法获取模型中定义的工况列表")

    except Exception as e:
        print(f"检查工况列表时出错: {e}")
        print("将跳过工况检查，直接尝试提取位移角...")

    print("重新设置输出工况选择...")
    check_ret(setup_api.DeselectAllCasesAndCombosForOutput(), "Setup.DeselectAllCasesAndCombosForOutput", (0, 1))

    selected_cases_count = 0
    for case_name in target_load_cases:
        print(f"选择工况/组合 '{case_name}' 以供输出...")
        ret_select = setup_api.SetCaseSelectedForOutput(case_name)
        ret_code = check_ret(ret_select, f"SetCaseSelectedForOutput({case_name})", (0, 1))

        if ret_code in (0, 1):
            if ret_code == 0:
                print(f"  工况 '{case_name}' 已成功选择。")
            else:
                print(f"  工况 '{case_name}' 已被选择 (状态未改变)。")
            selected_cases_count += 1
        else:
            print(f"  警告: 选择工况 '{case_name}' 失败。")

    if selected_cases_count == 0:
        print("错误: 没有成功选择任何工况进行输出。无法提取位移角。")
        return

    # 尝试设置位移角选项
    drift_option_relative = 0  # 0 for Relative drift
    print(f"尝试设置层间位移角选项为相对值...")

    drift_option_set_successfully = False
    if hasattr(setup_api, "Drift"):
        try:
            ret_drift_set = setup_api.Drift(drift_option_relative)
            check_ret(ret_drift_set, "Setup.Drift(Relative)", (0, 1))
            print("  位移角选项设置成功 (相对位移角)。")
            drift_option_set_successfully = True
        except Exception as e_drift:
            print(f"  设置位移角选项失败: {e_drift}")
    else:
        print("  当前ETABS版本可能不支持Drift选项设置，将使用默认设置。")

    print("正在调用 StoryDrifts API 获取数据...")

    # 初始化参数
    _NumberResults_ph = System.Int32(0)
    _Story_ph = System.Array[System.String](0)
    _LoadCase_ph = System.Array[System.String](0)
    _StepType_ph = System.Array[System.String](0)
    _StepNum_ph = System.Array[System.Double](0)
    _Dir_ph = System.Array[System.String](0)
    _DriftRatio_ph = System.Array[System.Double](0)
    _Label_ph = System.Array[System.String](0)
    _X_ph = System.Array[System.Double](0)
    _Y_ph = System.Array[System.Double](0)
    _Z_ph = System.Array[System.Double](0)

    try:
        # 尝试多种API调用方式
        api_call_successful = False

        # 方式1: 带Name和ItemTypeElm参数
        try:
            api_result_tuple = results_api.StoryDrifts(
                "",  # Name (空字符串表示所有楼层)
                ETABSv1.eItemTypeElm.Story,  # ItemTypeElm
                _NumberResults_ph,
                _Story_ph,
                _LoadCase_ph,
                _StepType_ph,
                _StepNum_ph,
                _Dir_ph,
                _DriftRatio_ph,
                _Label_ph,
                _X_ph,
                _Y_ph,
                _Z_ph,
            )
            api_call_successful = True
            print("  使用方式1成功调用StoryDrifts API")
        except Exception as e1:
            print(f"  方式1调用失败: {e1}")

            # 方式2: 不带前两个参数
            try:
                api_result_tuple = results_api.StoryDrifts(
                    _NumberResults_ph,
                    _Story_ph,
                    _LoadCase_ph,
                    _StepType_ph,
                    _StepNum_ph,
                    _Dir_ph,
                    _DriftRatio_ph,
                    _Label_ph,
                    _X_ph,
                    _Y_ph,
                    _Z_ph,
                )
                api_call_successful = True
                print("  使用方式2成功调用StoryDrifts API")
            except Exception as e2:
                print(f"  方式2调用失败: {e2}")
                raise Exception(f"所有StoryDrifts调用方式均失败。方式1错误: {e1}, 方式2错误: {e2}")

        if not api_call_successful:
            print("  所有StoryDrifts API调用方式均失败")
            return

        ret_code = check_ret(api_result_tuple[0], "Results.StoryDrifts", (0, 1))

        if ret_code == 1:
            print("  提示: StoryDrifts返回代码1，可能表示无数据或数据不完整，但将尝试继续处理...")

        num_res_val = api_result_tuple[1]
        story_val = list(api_result_tuple[2]) if api_result_tuple[2] is not None else []
        loadcase_val = list(api_result_tuple[3]) if api_result_tuple[3] is not None else []
        steptype_val = list(api_result_tuple[4]) if api_result_tuple[4] is not None else []
        stepnum_val = list(api_result_tuple[5]) if api_result_tuple[5] is not None else []
        dir_val = list(api_result_tuple[6]) if api_result_tuple[6] is not None else []
        drift_val = list(api_result_tuple[7]) if api_result_tuple[7] is not None else []
        label_val = list(api_result_tuple[8]) if api_result_tuple[8] is not None else []
        x_coord_val = list(api_result_tuple[9]) if api_result_tuple[9] is not None else []
        y_coord_val = list(api_result_tuple[10]) if api_result_tuple[10] is not None else []
        z_coord_val = list(api_result_tuple[11]) if api_result_tuple[11] is not None else []

        if num_res_val == 0:
            print("  未找到任何层间位移角结果。")
            print("  可能原因:")
            print("    1) 反应谱分析未完成")
            print("    2) 选择的工况没有位移结果")
            print("    3) 结构模型没有足够的层间约束")
            print("    4) 分析设置问题")
            print("  建议:")
            print("    1) 检查分析是否成功完成")
            print("    2) 在ETABS界面中手动查看Display > Show Tables > Analysis Results > Story Drift")
            print("    3) 检查工况设置和边界条件")
            return

        print(f"\n成功检索到 {num_res_val} 条层间位移角记录:")
        print("-" * 150)
        print(
            f"{'楼层名':<15} {'荷载工况/组合':<25} {'方向':<8} {'类型':<12} {'步号':<6} {'位移角 (rad)':<15} {'位移角 (‰)':<15} {'标签':<15} {'X':<10} {'Y':<10} {'Z':<10}"
        )
        print("-" * 150)

        max_drift_per_direction = {"X": 0.0, "Y": 0.0}
        max_drift_info = {"X": None, "Y": None}

        for i in range(num_res_val):
            drift_rad = drift_val[i]
            drift_permil = drift_rad * 1000.0

            direction_raw = dir_val[i].strip().upper()
            direction_key = None
            if direction_raw in ["X", "UX", "U1"]:
                direction_key = "X"
            elif direction_raw in ["Y", "UY", "U2"]:
                direction_key = "Y"

            if direction_key and direction_key in max_drift_per_direction:
                if abs(drift_permil) > abs(max_drift_per_direction[direction_key]):
                    max_drift_per_direction[direction_key] = abs(drift_permil)
                    max_drift_info[direction_key] = {
                        "story": story_val[i],
                        "load_case": loadcase_val[i],
                        "drift_permil": drift_permil,
                    }

            print(
                f"{story_val[i]:<15} {loadcase_val[i]:<25} {dir_val[i]:<8} {steptype_val[i]:<12} {stepnum_val[i]:<6.1f} {drift_rad:<15.6e} {drift_permil:<15.4f} {label_val[i]:<15} {x_coord_val[i]:<10.2f} {y_coord_val[i]:<10.2f} {z_coord_val[i]:<10.2f}"
            )

        print("-" * 150)

        print("\n=== 最大层间位移角总结 ===")
        for dir_key_summary in ["X", "Y"]:
            if max_drift_info[dir_key_summary] is not None:
                info = max_drift_info[dir_key_summary]
                print(
                    f"{dir_key_summary}方向最大位移角: {abs(info['drift_permil']):.4f}‰ (原始值: {info['drift_permil']:.4f}‰)"
                )
                print(f"  位置: {info['story']} 楼层, 工况: {info['load_case']}")
                actual_drift_limit_permil = 1.0
                if abs(info["drift_permil"]) > actual_drift_limit_permil:
                    print(
                        f"  ⚠️  警告: 超过建议限值 {actual_drift_limit_permil}‰ (1/{int(1000 / actual_drift_limit_permil)})"
                    )
                else:
                    print(f"  ✅ 满足建议限值 {actual_drift_limit_permil}‰ (1/{int(1000 / actual_drift_limit_permil)})")
            else:
                print(f"{dir_key_summary}方向: 未找到有效的位移角数据")
        print("=========================")

    except Exception as e_storydrifts_call:
        print(f"调用 StoryDrifts API 或处理其结果时发生错误: {e_storydrifts_call}")
        print("详细错误信息:")
        traceback.print_exc()
        print("\n建议:")
        print("1. 检查ETABS分析是否成功完成")
        print("2. 在ETABS界面中手动查看位移角结果: Display > Show Tables > Analysis Results > Story Drift")
        print("3. 检查API版本兼容性")
        return

    print("--- 层间位移角提取完毕 ---")


def extract_modal_and_drift(sap_model, output_dir: Union[str, Path]) -> Path:
    """
    使用现有的 ETABS 代码，提取：
      1. 模态周期和频率
      2. 模态参与质量系数
      3. 相对层间位移角 (RS-X, RS-Y)
    并把整个打印输出写入 Excel 文件：
       output_dir / 'analysis_dynamic_summary.xlsx'
    Excel 至少包含一列 'output'，每一行是一行原始文本。
    返回该 Excel 文件的 Path。
    """
    output_dir = Path(output_dir)
    summary_path = output_dir / "analysis_dynamic_summary.xlsx"

    with capture_stdout_to_buffer() as buf:
        extract_modal_and_mass_info(sap_model)
        extract_story_drifts_improved(sap_model, ["RS-X", "RS-Y"])

    text = buf.getvalue()

    # 回显到终端
    print(text, end="")

    # 写入 Excel（仅保留关键信息行）
    lines = text.splitlines()
    filtered_lines = [ln for ln in lines if _is_important_line(ln)]
    df = pd.DataFrame({"output": filtered_lines})
    summary_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(summary_path, index=False)

    return summary_path


__all__ = [
    "capture_stdout_to_buffer",
    "_is_important_line",
    "extract_modal_and_mass_info",
    "extract_story_drifts_improved",
    "extract_modal_and_drift",
]
