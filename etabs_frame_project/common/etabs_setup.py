#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETABS 设置模块
负责ETABS连接、模型初始化
"""

import time
import sys
from .utility_functions import check_ret
from .config import (
    ATTACH_TO_INSTANCE, REMOTE, REMOTE_COMPUTER, SPECIFY_PATH, PROGRAM_PATH,
    NUM_STORIES, TYPICAL_STORY_HEIGHT, BOTTOM_STORY_HEIGHT,
    NUM_GRID_LINES_X, NUM_GRID_LINES_Y, SPACING_X, SPACING_Y
)

# 全局变量
my_etabs = None
sap_model = None


def setup_etabs():
    """设置ETABS连接与模型初始化"""
    global my_etabs, sap_model

    # 重新导入API对象以确保它们已正确加载
    from .etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    if ETABSv1 is None:
        sys.exit("致命错误: ETABSv1 API 未正确加载")

    print("\nETABS 连接与模型初始化...")
    helper = ETABSv1.cHelper(ETABSv1.Helper())

    if ATTACH_TO_INSTANCE:
        print("正在尝试附加到已运行的ETABS 实例...")
        try:
            getter = helper.GetObjectHost if REMOTE else helper.GetObject
            my_etabs = getter(REMOTE_COMPUTER if REMOTE else "CSI.ETABS.API.ETABSObject")
            print("已成功附加到 ETABS 实例。")
        except COMException as e:
            sys.exit(f"致命错误: 附加到 ETABS 实例失败。COMException: {e}\n请确保 ETABS 正在运行。")
        except Exception as e:
            sys.exit(f"致命错误: 附加到 ETABS 实例时发生未知错误: {e}")
    else:
        print("正在启动新的 ETABS 实例...")
        try:
            creator = helper.CreateObjectHost if REMOTE and SPECIFY_PATH else \
                helper.CreateObject if SPECIFY_PATH else \
                    helper.CreateObjectProgIDHost if REMOTE else \
                        helper.CreateObjectProgID
            path_or_progid = PROGRAM_PATH if SPECIFY_PATH else "CSI.ETABS.API.ETABSObject"
            my_etabs = creator(REMOTE_COMPUTER if REMOTE else path_or_progid)
        except COMException as e:
            sys.exit(f"致命错误: 启动 ETABS实例失败。COMException: {e}\n请检查 PROGRAM_PATH 或 ProgID。")
        except Exception as e:
            sys.exit(f"致命错误: 启动 ETABS 实例时发生未知错误: {e}")

        check_ret(my_etabs.ApplicationStart(), "my_etabs.ApplicationStart")
        print("ETABS 应用程序已启动。")

    print("等待 ETABS 用户界面初始化 (大约5秒)...")
    time.sleep(5)

    sap_model = my_etabs.SapModel
    if sap_model is None:
        sys.exit("致命错误: my_etabs.SapModel 返回为 None。")

    try:
        sap_model.SetModelIsLocked(False)
        print("已尝试设置模型为未锁定状态。")
    except Exception as e_lock:
        print(f"警告: 设置模型未锁定状态失败: {e_lock}")

    check_ret(sap_model.InitializeNewModel(ETABSv1.eUnits.kN_m_C), "sap_model.InitializeNewModel")
    print(f"新模型已成功初始化, 单位设置为: kN, m, °C ")

    file_obj = ETABSv1.cFile(sap_model.File)
    check_ret(
        file_obj.NewGridOnly(NUM_STORIES, TYPICAL_STORY_HEIGHT, BOTTOM_STORY_HEIGHT,
                             NUM_GRID_LINES_X, NUM_GRID_LINES_Y, SPACING_X, SPACING_Y),
        "file_obj.NewGridOnly"
    )
    print(f"空白网格模型已创建 ({NUM_STORIES}层, X向轴线: {NUM_GRID_LINES_X}, Y向轴线: {NUM_GRID_LINES_Y})。")

    return my_etabs, sap_model


def get_etabs_objects():
    """获取ETABS对象"""
    global my_etabs, sap_model
    return my_etabs, sap_model


def get_sap_model():
    """
    获取SAP模型对象
    这是设计内力提取模块需要的函数

    Returns:
        sap_model: ETABS SAP模型对象，如果未初始化则返回None
    """
    global sap_model
    if sap_model is None:
        print("⚠️ SAP模型对象未初始化，请先运行 setup_etabs()")
        return None
    return sap_model


def set_sap_model(model):
    """
    设置SAP模型对象

    Args:
        model: ETABS SAP模型对象
    """
    global sap_model
    sap_model = model


def is_etabs_connected():
    """
    检查ETABS连接状态

    Returns:
        bool: True如果已连接，False如果未连接
    """
    global my_etabs, sap_model
    try:
        if my_etabs is None or sap_model is None:
            return False
        # 尝试执行一个简单的操作来测试连接
        _ = sap_model.GetModelFilename()
        return True
    except:
        return False


def ensure_etabs_ready():
    """
    确保ETABS已准备就绪，如果未连接则尝试重新连接

    Returns:
        bool: True如果ETABS已准备就绪，False如果失败
    """
    if is_etabs_connected():
        return True

    print("🔄 ETABS连接丢失，尝试重新连接...")
    try:
        setup_etabs()
        return is_etabs_connected()
    except Exception as e:
        print(f"❌ 重新连接ETABS失败: {e}")
        return False


# 导出函数列表
__all__ = [
    'setup_etabs',
    'get_etabs_objects',
    'get_sap_model',  # 新增
    'set_sap_model',  # 新增
    'is_etabs_connected',  # 新增
    'ensure_etabs_ready'  # 新增
]
