#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件操作模块
处理模型保存和文件管理
"""

import os
import sys
from etabs_setup import get_etabs_objects
from utility_functions import check_ret
from config import MODEL_PATH, SCRIPT_DIRECTORY, ATTACH_TO_INSTANCE


def finalize_and_save_model():
    """
    最后刷新一次视图（可选），然后保存 .EDB。
    允许 RefreshView 返回 0 或 1——1 只是"视图未变化/忽略"，不算错误。
    """
    my_etabs, sap_model = get_etabs_objects()

    # 若模型对象不存在，直接返回
    if sap_model is None:
        print("警告: SapModel 未初始化，无法保存模型。")
        return

    print(f"\n保存模型到: {MODEL_PATH}")

    # ---------- 1) 刷新视图（非关键，可删） ----------
    try:
        # 动态导入API对象
        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if ETABSv1 is not None:
            check_ret(
                ETABSv1.cView(sap_model.View).RefreshView(0, False),
                "RefreshView",
                ok_codes=(0, 1)  # ← 允许 0 或 1
            )
            print("模型视图已刷新")
        else:
            print("跳过视图刷新（ETABSv1 API对象不可用）")
    except Exception as e:
        print(f"刷新视图失败 (非关键错误): {e}")

    # ---------- 2) 确保输出目录存在 ----------
    try:
        os.makedirs(SCRIPT_DIRECTORY, exist_ok=True)
        print(f"输出目录已确保存在: {SCRIPT_DIRECTORY}")
    except Exception as e:
        sys.exit(f"创建输出目录失败: {e}")

    # ---------- 3) 保存模型 ----------
    try:
        # 动态导入API对象
        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if ETABSv1 is None:
            sys.exit("致命错误: ETABSv1 API对象不可用，无法保存模型")

        ret_save = ETABSv1.cFile(sap_model.File).Save(MODEL_PATH)
        if ret_save != 0:
            if not ATTACH_TO_INSTANCE and my_etabs is not None:
                my_etabs.ApplicationExit(False)
            sys.exit(f"保存模型失败: {ret_save}")

        print(f"✅ 模型已成功保存到: {MODEL_PATH}")

        # 显示文件信息
        if os.path.exists(MODEL_PATH):
            file_size = os.path.getsize(MODEL_PATH) / (1024 * 1024)  # MB
            print(f"文件大小: {file_size:.2f} MB")

    except Exception as e:
        print(f"保存模型时发生异常: {e}")
        if not ATTACH_TO_INSTANCE and my_etabs is not None:
            try:
                my_etabs.ApplicationExit(False)
            except:
                pass
        sys.exit(f"保存模型失败: {e}")


def cleanup_etabs_on_error():
    """错误时清理ETABS"""
    if not ATTACH_TO_INSTANCE:
        try:
            my_etabs, _ = get_etabs_objects()
            if my_etabs is not None:
                my_etabs.ApplicationExit(False)
                print("ETABS 已关闭")
        except Exception as e:
            print(f"关闭 ETABS 失败: {e}")


def check_output_directory():
    """检查输出目录"""
    try:
        os.makedirs(SCRIPT_DIRECTORY, exist_ok=True)
        print(f"输出目录: {SCRIPT_DIRECTORY}")
        return True
    except Exception as e:
        print(f"致命错误: 创建脚本输出目录 '{SCRIPT_DIRECTORY}' 失败: {e}")
        return False


# 导出函数列表
__all__ = [
    'finalize_and_save_model',
    'cleanup_etabs_on_error',
    'check_output_directory'
]