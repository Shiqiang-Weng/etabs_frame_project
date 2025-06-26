#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
构件内力提取模块 (修复版)
从框架构件（梁、柱）中提取内力（轴力、剪力、弯矩）
"""

import os
import csv
import traceback
from typing import List, Dict, Any

from etabs_setup import get_etabs_objects
from utility_functions import check_ret
from config import SCRIPT_DIRECTORY
from etabs_api_loader import get_api_objects


def _prepare_force_output_params():
    """
    准备用于 FrameForce API 调用的、正确类型的输出参数元组。
    """
    ETABSv1, System, COMException = get_api_objects()
    return (
        System.Int32(0),  # NumberResults
        System.Array[System.String](0),  # Obj
        System.Array[System.Double](0),  # ObjSta (Corrected to Double)
        System.Array[System.String](0),  # Elm
        System.Array[System.Double](0),  # ElmSta (Corrected to Double)
        System.Array[System.String](0),  # LoadCase
        System.Array[System.String](0),  # StepType
        System.Array[System.Double](0),  # StepNum
        System.Array[System.Double](0),  # P
        System.Array[System.Double](0),  # V2
        System.Array[System.Double](0),  # V3
        System.Array[System.Double](0),  # T
        System.Array[System.Double](0),  # M2
        System.Array[System.Double](0)  # M3
    )


def extract_frame_forces(frame_names: List[str], load_cases: List[str]) -> List[Dict[str, Any]]:
    """
    为指定的框架构件和荷载工况提取内力。

    Parameters:
    ----------
    frame_names : List[str]
        需要提取内力的框架构件名称列表。
    load_cases : List[str]
        需要提取内力的荷载工况/组合列表。

    Returns:
    -------
    List[Dict[str, Any]]
        包含内力结果的字典列表。
    """
    my_etabs, sap_model = get_etabs_objects()
    if not all([sap_model, hasattr(sap_model, "Results")]):
        print("错误: 结果不可用，无法提取构件内力。")
        return []

    ETABSv1, System, COMException = get_api_objects()
    if not all([ETABSv1, System]):
        print("错误: .NET 模块未正确加载，无法提取内力。")
        return []

    results_api = sap_model.Results
    setup_api = results_api.Setup

    print("\n--- 开始提取构件内力 ---")
    print(f"目标构件数量: {len(frame_names)}")
    print(f"目标荷载工况: {load_cases}")

    # 1. 选择要输出的荷载工况
    check_ret(setup_api.DeselectAllCasesAndCombosForOutput(), "DeselectAllCasesForForces", (0, 1))
    for case in load_cases:
        check_ret(setup_api.SetCaseSelectedForOutput(case), f"SetCaseSelectedForOutput({case})", (0, 1))

    all_forces_data = []
    processed_count = 0

    # 2. 遍历每个框架构件提取内力
    for frame_name in frame_names:
        try:
            # 准备正确类型的输出参数
            params = _prepare_force_output_params()

            # 调用API提取内力
            force_res = results_api.FrameForce(frame_name, ETABSv1.eItemTypeElm.ObjectElm, *params)

            check_ret(force_res[0], f"FrameForce({frame_name})", (0, 1))

            num_results = force_res[1]
            if num_results > 0:
                # 解包所有返回的结果
                _, _, obj_names, obj_stas, elm_names, elm_stas, res_cases, step_types, step_nums, \
                    p_forces, v2_forces, v3_forces, t_forces, m2_moments, m3_moments = force_res

                # 将结果整理成字典格式
                for i in range(num_results):
                    force_data = {
                        "FrameName": obj_names[i],
                        "Station (m)": round(obj_stas[i], 4),
                        "LoadCase": res_cases[i],
                        "P (kN)": round(p_forces[i], 3),
                        "V2 (kN)": round(v2_forces[i], 3),
                        "V3 (kN)": round(v3_forces[i], 3),
                        "T (kN-m)": round(t_forces[i], 3),
                        "M2 (kN-m)": round(m2_moments[i], 3),
                        "M3 (kN-m)": round(m3_moments[i], 3),
                    }
                    all_forces_data.append(force_data)

            processed_count += 1
            if processed_count % 100 == 0:
                print(f"  已处理 {processed_count}/{len(frame_names)} 个构件...")

        except Exception as e:
            print(f"  提取构件 '{frame_name}' 内力时发生错误: {e}")
            # traceback.print_exc() # 可取消注释以获得详细堆栈跟踪

    print(f"--- 构件内力提取完成 ---")
    print(f"共提取到 {len(all_forces_data)} 条内力记录。")
    return all_forces_data


def save_forces_to_csv(force_data: List[Dict[str, Any]], filename: str):
    """
    将提取的内力数据保存到CSV文件。

    Parameters:
    ----------
    force_data : List[Dict[str, Any]]
        内力数据列表。
    filename : str
        要保存的CSV文件名。
    """
    if not force_data:
        print("无内力数据可保存。")
        return

    filepath = os.path.join(SCRIPT_DIRECTORY, filename)
    print(f"\n正在将内力结果保存到: {filepath}")

    try:
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, 'w', newline='', encoding='utf-8-sig') as csvfile:
            fieldnames = force_data[0].keys()
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            writer.writeheader()
            writer.writerows(force_data)

        print(f"✅ 内力结果已成功保存。")
    except Exception as e:
        print(f"❌ 保存CSV文件失败: {e}")


def extract_and_save_frame_forces(all_frame_names: List[str]):
    """
    提取并保存所有指定框架构件的内力。

    Parameters:
    ----------
    all_frame_names : List[str]
        所有梁和柱的名称列表。
    """
    # 定义要为其提取内力的荷载工况
    target_cases = ["DEAD", "LIVE", "RS-X", "RS-Y"]

    # 提取内力
    force_results = extract_frame_forces(all_frame_names, target_cases)

    # 保存到CSV
    if force_results:
        save_forces_to_csv(force_results, "frame_member_forces.csv")
    else:
        print("未提取到任何构件内力，不生成CSV文件。")


# 导出函数列表
__all__ = [
    'extract_and_save_frame_forces'
]