#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析状态检查模块
提供分析完成状态的快速检测。
"""

from __future__ import annotations

from etabs_setup import get_etabs_objects


def _log(message: str) -> None:
    """统一的分析阶段日志输出。"""
    print(f"[分析] {message}")


def check_analysis_completion() -> bool:
    """
    检查分析结果模块是否可用，用于判断分析是否完成。

    Returns:
        bool: True 表示结果模块可用，False 表示不可用或检查出错。
    """
    _, sap_model = get_etabs_objects()
    if sap_model is None:
        return False

    _log("检查分析完成状态...")

    try:
        if hasattr(sap_model, "Results") and sap_model.Results is not None:
            _log("✅ 分析结果模块可用")
            return True

        _log("❌ 分析结果模块不可用")
        return False
    except Exception as exc:  # noqa: BLE001
        _log(f"❌ 检查分析状态时出错: {exc}")
        return False
