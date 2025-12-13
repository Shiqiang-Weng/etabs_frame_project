#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETABS 设置模块
负责 ETABS 连接、模型初始化
"""

import time
import sys
from .utility_functions import check_ret
from .config import (
    ATTACH_TO_INSTANCE,
    REMOTE,
    REMOTE_COMPUTER,
    SPECIFY_PATH,
    PROGRAM_PATH,
    DEFAULT_DESIGN_CONFIG,
    design_config_from_case,
)

# 全局变量
my_etabs = None
sap_model = None


def setup_etabs(design=None):
    """设置 ETABS 连接与模型初始化"""
    global my_etabs, sap_model
    design_cfg = design_config_from_case(design) if design is not None else DEFAULT_DESIGN_CONFIG

    # 重新导入 API 对象以确保它们已正确加载
    from .etabs_api_loader import get_api_objects

    ETABSv1, System, COMException = get_api_objects()

    if ETABSv1 is None:
        sys.exit("致命错误: ETABSv1 API 未正确加载")

    print("\nETABS 连接与模型初始化...")
    helper = ETABSv1.cHelper(ETABSv1.Helper())

    if ATTACH_TO_INSTANCE:
        print("正在尝试附加到已运行的 ETABS 实例...")
        try:
            getter = helper.GetObjectHost if REMOTE else helper.GetObject
            my_etabs = getter(REMOTE_COMPUTER if REMOTE else "CSI.ETABS.API.ETABSObject")
            print("已成功附加到 ETABS 实例。")
        except COMException as e:
            sys.exit(f"致命错误: 附加到 ETABS 实例失败。COMException: {e}\n请确认 ETABS 正在运行。")
        except Exception as e:
            sys.exit(f"致命错误: 附加到 ETABS 实例时发生未知错误 {e}")
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
            sys.exit(f"致命错误: 启动 ETABS 实例失败。COMException: {e}\n请检查 PROGRAM_PATH 或 ProgID。")
        except Exception as e:
            sys.exit(f"致命错误: 启动 ETABS 实例时发生未知错误 {e}")

        check_ret(my_etabs.ApplicationStart(), "my_etabs.ApplicationStart")
        print("ETABS 应用程序已启动。")

    print("等待 ETABS 用户界面初始化（约 5 秒）...")
    time.sleep(5)

    sap_model = my_etabs.SapModel
    if sap_model is None:
        sys.exit("致命错误: my_etabs.SapModel 返回 None。")

    try:
        sap_model.SetModelIsLocked(False)
        print("已尝试设置模型为未锁定状态。")
    except Exception as e_lock:
        print(f"警告: 设置模型未锁定状态失败: {e_lock}")

    check_ret(sap_model.InitializeNewModel(ETABSv1.eUnits.kN_m_C), "sap_model.InitializeNewModel")
    print("新模型已成功初始化，单位设置为 kN, m, °C。")

    file_obj = ETABSv1.cFile(sap_model.File)
    check_ret(
        file_obj.NewGridOnly(
            design_cfg.storeys.num_storeys,
            design_cfg.storeys.storey_height,
            design_cfg.storeys.storey_height,
            design_cfg.grid.num_x,
            design_cfg.grid.num_y,
            design_cfg.grid.spacing_x,
            design_cfg.grid.spacing_y,
        ),
        "file_obj.NewGridOnly",
    )
    print(
        f"空白网格模型已创建 ({design_cfg.storeys.num_storeys}层, X向轴线: {design_cfg.grid.num_x}, "
        f"Y向轴线: {design_cfg.grid.num_y})。"
    )

    return my_etabs, sap_model


def get_etabs_objects():
    """获取 ETABS 对象"""
    global my_etabs, sap_model
    return my_etabs, sap_model


def get_sap_model():
    """
    获取 SAP 模型对象
    这是设计内力提取模块需要的函数

    Returns:
        sap_model: ETABS SAP 模型对象，如未初始化则返回 None
    """
    global sap_model
    if sap_model is None:
        print("[警告] SAP 模型对象未初始化，请先运行 setup_etabs()")
        return None
    return sap_model


def set_sap_model(model):
    """设置 SAP 模型对象"""
    global sap_model
    sap_model = model


def is_etabs_connected():
    """
    检查 ETABS 连接状态
    Returns:
        bool: True 如已连接，False 如未连接
    """
    global my_etabs, sap_model
    try:
        if my_etabs is None or sap_model is None:
            return False
        _ = sap_model.GetModelFilename()
        return True
    except Exception:
        return False


def ensure_etabs_ready():
    """确保 ETABS 已准备就绪，如未连接则尝试重新连接"""
    if is_etabs_connected():
        return True

    print("[重连] ETABS 连接丢失，尝试重新连接...")
    try:
        setup_etabs()
        return is_etabs_connected()
    except Exception as e:
        print(f"[错误] 重新连接 ETABS 失败: {e}")
        return False


__all__ = [
    'setup_etabs',
    'get_etabs_objects',
    'get_sap_model',
    'set_sap_model',
    'is_etabs_connected',
    'ensure_etabs_ready'
]
