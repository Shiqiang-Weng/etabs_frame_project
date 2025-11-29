#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析运行模块
负责触发 ETABS 分析流程。
"""

from __future__ import annotations

import sys
import time
from typing import Sequence

from common.config import ATTACH_TO_INSTANCE, MODEL_PATH, MODAL_CASE_NAME
from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret

# 固定分析工况顺序
DEFAULT_LOAD_CASES = ("DEAD", "LIVE", MODAL_CASE_NAME, "RS-X", "RS-Y")


def _log(message: str) -> None:
    """统一的分析阶段日志输出。"""
    print(f"[分析] {message}")


def safe_run_analysis(load_cases_to_run: Sequence[str], delete_old_results: bool = True) -> None:
    """
    安全运行 ETABS 分析：解锁模型、保存、选工况、清理旧结果并执行分析。

    Args:
        load_cases_to_run: 需要运行的工况名称列表。
        delete_old_results: 运行前是否清理旧结果。
    """
    _, sap_model = get_etabs_objects()
    if sap_model is None:
        raise RuntimeError("SapModel 未初始化，无法运行分析。")

    from common.etabs_api_loader import get_api_objects  # 动态导入以避免循环依赖

    _, System, _ = get_api_objects()
    if System is None:
        raise RuntimeError("System 模块未正确加载，无法运行分析。")

    analyze_obj = sap_model.Analyze
    file_api = sap_model.File

    _log("准备运行分析")
    _log(f"分析工况: {list(load_cases_to_run)}")

    # 确保模型未锁定
    check_ret(sap_model.SetModelIsLocked(False), "SetModelIsLocked(False) before analysis", (0, 1))

    # 保存模型
    if file_api.Save(MODEL_PATH) != 0:
        raise RuntimeError("分析前保存模型失败。")
    _log("模型已保存")

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
            _log(f"⚠️ 工况 '{case}' 未定义，跳过。")
            continue
        if sap_model.Analyze.SetRunCaseFlag(case, True) != 0:
            _log(f"⚠️ 设置工况 '{case}' 运行失败。")

    # 清理旧结果
    if delete_old_results:
        try:
            check_ret(analyze_obj.DeleteResults(""), "DeleteResults", (0, 1))
            _log("已清理旧的分析结果")
        except Exception as exc:  # noqa: BLE001
            _log(f"⚠️ 清理旧分析结果失败: {exc}（可能无历史结果或 API 限制）")

    # 运行分析
    _log("开始运行分析...")
    if analyze_obj.RunAnalysis() != 0:
        raise RuntimeError("RunAnalysis 执行失败。")

    _log("✅ 分析成功完成。")


def wait_and_run_analysis(wait_seconds: int = 5) -> None:
    """
    在指定等待后运行固定工况的分析。

    Args:
        wait_seconds: 运行前的等待时间（秒）。
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        _log("SapModel 未就绪，无法分析。")
        return

    load_cases = list(DEFAULT_LOAD_CASES)
    _log(f"等待 {wait_seconds} 秒后开始分析工况: {load_cases}")
    time.sleep(wait_seconds)

    try:
        safe_run_analysis(load_cases)
    except Exception as exc:  # noqa: BLE001
        _log(f"分析执行错误: {exc}")
        if not ATTACH_TO_INSTANCE:
            my_etabs, _ = get_etabs_objects()
            if my_etabs is not None:
                my_etabs.ApplicationExit(False)
        sys.exit("分析失败，脚本中止。")
