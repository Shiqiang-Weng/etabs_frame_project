#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
File operations: model saving, output directory checks, and cleanup helpers.
"""

import os
import sys
import shutil
from pathlib import Path
from .etabs_setup import get_etabs_objects
from .utility_functions import check_ret
from .config import (
    MODEL_PATH,
    SCRIPT_DIRECTORY,
    DATA_EXTRACTION_DIR,
    ANALYSIS_DATA_DIR,
    DESIGN_DATA_DIR,
    ATTACH_TO_INSTANCE,
)


def finalize_and_save_model(create_analysis_dir: bool = True, create_design_dir: bool = True):
    """Refresh view (best effort), ensure output dir, and save the model."""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        print("Warning: SapModel not initialized; cannot save model.")
        return

    print(f"\nSaving model to {MODEL_PATH}")

    # 1) Refresh view (optional)
    try:
        from .etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()
        if ETABSv1 is not None:
            check_ret(
                ETABSv1.cView(sap_model.View).RefreshView(0, False),
                "RefreshView",
                ok_codes=(0, 1),
            )
            print("Model view refreshed.")
        else:
            print("Skip view refresh (ETABS API not available).")
    except Exception as e:
        print(f"View refresh failed (non-critical): {e}")

    # 2) Ensure output directory
    try:
        os.makedirs(SCRIPT_DIRECTORY, exist_ok=True)
        os.makedirs(DATA_EXTRACTION_DIR, exist_ok=True)
        if create_analysis_dir:
            os.makedirs(ANALYSIS_DATA_DIR, exist_ok=True)
            print(f"分析数据目录已确保存在: {ANALYSIS_DATA_DIR}")
        if create_design_dir:
            os.makedirs(DESIGN_DATA_DIR, exist_ok=True)
            print(f"设计数据目录已确保存在: {DESIGN_DATA_DIR}")
        print(f"输出目录已确保存在: {SCRIPT_DIRECTORY}")
        print(f"数据导出目录已确保存在: {DATA_EXTRACTION_DIR}")
    except Exception as e:
        sys.exit(f"创建输出目录失败: {e}")

    # 3) Save model
    try:
        from .etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()
        if ETABSv1 is None:
            sys.exit("致命错误: ETABS API 不可用，无法保存模型")

        ret_save = ETABSv1.cFile(sap_model.File).Save(MODEL_PATH)
        if ret_save != 0:
            if not ATTACH_TO_INSTANCE and my_etabs is not None:
                my_etabs.ApplicationExit(False)
            sys.exit(f"保存模型失败: {ret_save}")

        print(f"✓ 模型已保存: {MODEL_PATH}")
        if os.path.exists(MODEL_PATH):
            file_size = os.path.getsize(MODEL_PATH) / (1024 * 1024)
            print(f"文件大小: {file_size:.2f} MB")

    except Exception as e:
        print(f"保存模型时发生异常: {e}")
        if not ATTACH_TO_INSTANCE and my_etabs is not None:
            try:
                my_etabs.ApplicationExit(False)
            except Exception:
                pass
        sys.exit(f"保存模型失败: {e}")


def cleanup_etabs_on_error():
    """Close ETABS when not attaching to an existing instance."""
    if not ATTACH_TO_INSTANCE:
        try:
            my_etabs, _ = get_etabs_objects()
            if my_etabs is not None:
                my_etabs.ApplicationExit(False)
                print("ETABS 已关闭")
        except Exception as e:
            print(f"关闭 ETABS 失败: {e}")


def check_output_directory(create_analysis_dir: bool = True, create_design_dir: bool = True):
    """Ensure output directory exists (including data_extraction)."""
    try:
        os.makedirs(SCRIPT_DIRECTORY, exist_ok=True)
        os.makedirs(DATA_EXTRACTION_DIR, exist_ok=True)
        if create_analysis_dir:
            os.makedirs(ANALYSIS_DATA_DIR, exist_ok=True)
        if create_design_dir:
            os.makedirs(DESIGN_DATA_DIR, exist_ok=True)
        print(f"输出目录: {SCRIPT_DIRECTORY}")
        print(f"数据导出目录: {DATA_EXTRACTION_DIR}")
        if create_analysis_dir:
            print(f"分析数据目录: {ANALYSIS_DATA_DIR}")
        if create_design_dir:
            print(f"设计数据目录: {DESIGN_DATA_DIR}")
        return True
    except Exception as e:
        print(f"致命错误: 创建脚本输出目录 '{SCRIPT_DIRECTORY}' 失败: {e}")
        return False


def remove_pycache(root=None):
    """Best-effort removal of __pycache__ directories under root (defaults to project root)."""
    base = Path(root) if root else Path(__file__).resolve().parent.parent
    for cache_dir in base.rglob("__pycache__"):
        shutil.rmtree(cache_dir, ignore_errors=True)


__all__ = [
    "finalize_and_save_model",
    "cleanup_etabs_on_error",
    "check_output_directory",
    "remove_pycache",
]
