#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析模块
运行ETABS结构分析
"""

import sys
import time
from typing import List
from etabs_setup import get_etabs_objects
from utility_functions import check_ret
from config import MODEL_PATH, MODAL_CASE_NAME, ATTACH_TO_INSTANCE


def safe_run_analysis(load_cases_to_run: List[str], delete_old_results: bool = True):
    """
    安全运行分析

    Parameters:
    ----------
    load_cases_to_run : List[str]
        要运行的荷载工况列表
    delete_old_results : bool
        是否删除旧结果
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        raise RuntimeError("SapModel 未初始化，无法运行分析。")

    # 动态导入API对象
    from etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    if System is None:
        raise RuntimeError("System 模块未正确加载，无法运行分析。")

    analyze_obj = sap_model.Analyze
    file_api = sap_model.File

    print(f"\n准备运行分析...")
    print(f"分析工况: {load_cases_to_run}")

    # 确保模型未锁定
    check_ret(sap_model.SetModelIsLocked(False), "SetModelIsLocked(False) before analysis", (0, 1))

    # 保存模型
    if file_api.Save(MODEL_PATH) != 0:
        raise RuntimeError("分析前保存模型失败。")
    print("模型已保存")

    if not load_cases_to_run:
        raise RuntimeError("未指定分析工况。")

    # 获取已定义的工况列表
    num_val = System.Int32(0)
    names_val = System.Array[System.String](0)
    ret_tuple = sap_model.LoadCases.GetNameList(num_val, names_val)
    defined_cases = list(ret_tuple[2]) if ret_tuple[0] == 0 and ret_tuple[1] > 0 and ret_tuple[2] is not None else []

    # 设置要运行的工况
    for case in load_cases_to_run:
        if defined_cases and case not in defined_cases:
            print(f"警告: 工况 '{case}' 未定义。")
            continue
        if sap_model.Analyze.SetRunCaseFlag(case, True) != 0:
            print(f"警告: 设置工况 '{case}' 运行失败。")

    # 清理旧结果
    if delete_old_results:
        try:
            check_ret(analyze_obj.DeleteResults(""), "DeleteResults", (0, 1))
            print("已清理旧的分析结果")
        except:
            print("清理旧结果失败 (可能无结果或API版本问题)。")

    # 运行分析
    print("开始运行分析...")
    if analyze_obj.RunAnalysis() != 0:
        raise RuntimeError("RunAnalysis 执行失败。")

    print("✅ 分析成功完成！")


def wait_and_run_analysis(wait_seconds: int = 5):
    """
    等待并运行分析

    Parameters:
    ----------
    wait_seconds : int
        等待时间（秒）
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        print("SapModel 未就绪，无法分析。")
        return

    # 定义要分析的工况
    load_cases = ["DEAD", "LIVE", MODAL_CASE_NAME, "RS-X", "RS-Y"]

    print(f"等待 {wait_seconds} 秒后开始分析工况: {load_cases}")
    time.sleep(wait_seconds)

    try:
        safe_run_analysis(load_cases)
    except Exception as e:
        print(f"分析执行错误: {e}")
        if not ATTACH_TO_INSTANCE:
            my_etabs, _ = get_etabs_objects()
            if my_etabs is not None:
                my_etabs.ApplicationExit(False)
        sys.exit("分析失败，脚本中止。")


def check_analysis_completion():
    """检查分析是否完成"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return False

    print("\n检查分析完成状态...")

    # 尝试获取模型信息以确认分析状态
    try:
        # 检查是否有结果可用
        if hasattr(sap_model, "Results") and sap_model.Results is not None:
            print("✅ 分析结果模块可用")
            return True
        else:
            print("❌ 分析结果模块不可用")
            return False
    except Exception as e:
        print(f"❌ 检查分析状态时出错: {e}")
        return False


# 导出函数列表
__all__ = [
    'safe_run_analysis',
    'wait_and_run_analysis',
    'check_analysis_completion'
]