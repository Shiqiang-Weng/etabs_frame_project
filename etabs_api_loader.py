#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETABS API 加载器
负责加载 .NET 运行时和 ETABS API
"""

import sys
from config import ETABS_DLL_PATH


def load_dotnet_etabs_api():
    """加载.NET运行时和ETABS API"""
    global ETABSv1, System, COMException

    print("\n正在加载.NET 运行时...")
    try:
        from pythonnet import load
        load("netfx")
        print(".NET Framework 运行时 (netfx)已尝试加载。")

        import clr
        import System as Sys
        System = Sys

        from System.Runtime.InteropServices import COMException as ComExc
        COMException = ComExc
        print(".NET 运行时及 System 相关库已成功加载。")

        print(f"正在从指定路径加载 ETABS API: {ETABS_DLL_PATH}")
        clr.AddReference(ETABS_DLL_PATH)

        import ETABSv1 as EtabsApiModule
        ETABSv1 = EtabsApiModule
        print("ETABS API 引用已成功加载。现在可以通过 ETABSv1.xxx 访问其成员。")

        return ETABSv1, System, COMException

    except ImportError as exc:
        sys.exit(f"致命错误: pythonnet 库缺失或加载失败: {exc}。\n请通过 'pip install pythonnet' 命令安装。")
    except FileNotFoundError:
        sys.exit(
            f"致命错误: 在路径 {ETABS_DLL_PATH} 未找到 ETABSv1.dll 文件。\n"
            f"请确认 ETABS 安装正确, 并且 ETABS_DLL_PATH 配置无误。"
        )
    except Exception as e:
        sys.exit(f"致命错误: 加载.NET 运行时, System 库, 或 ETABS API 时发生意外错误: {e}")


def get_api_objects():
    """获取加载后的API对象"""
    global ETABSv1, System, COMException
    return ETABSv1, System, COMException


# 模块级变量
ETABSv1 = None
System = None
COMException = None