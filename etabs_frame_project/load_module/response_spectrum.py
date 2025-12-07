#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
反应谱定义模块
基于GB50011-2010规范的反应谱函数定义
"""

import sys
from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret, arr
from common.config import (
    RS_BASE_ACCEL_G, RS_CHARACTERISTIC_PERIOD, RS_DAMPING_RATIO,
    GRAVITY_ACCEL, RS_FUNCTION_NAME
)


def china_response_spectrum(T: float, zeta: float, alpha_max: float, Tg: float, g: float = GRAVITY_ACCEL) -> float:
    """
    中国规范反应谱函数 (GB50011-2010)

    Parameters:
    ----------
    T : float
        周期 (s)
    zeta : float
        阻尼比
    alpha_max : float
        最大地震影响系数
    Tg : float
        场地特征周期 (s)
    g : float
        重力加速度 (m/s²)

    Returns:
    -------
    float
        地震影响系数对应的加速度值
    """
    # 计算调整系数
    gamma = 0.9 + (0.05 - zeta) / (0.3 + 6.0 * zeta)
    eta1 = max(0.0, 0.02 + (0.05 - zeta) / (4.0 + 32.0 * zeta))
    eta2 = max(0.55, 1.0 + (0.05 - zeta) / (0.08 + 1.6 * zeta))

    # 计算地震影响系数
    current_alpha_coeff: float
    if T < 0:
        current_alpha_coeff = 0.0
    elif T == 0.0:
        current_alpha_coeff = 0.45 * alpha_max
    elif T <= 0.1:
        current_alpha_coeff = min((0.45 + 10.0 * (eta2 - 0.45) * T) * alpha_max, eta2 * alpha_max)
    elif T <= Tg:
        current_alpha_coeff = eta2 * alpha_max
    elif T <= 5.0 * Tg:
        current_alpha_coeff = ((Tg / T) ** gamma) * eta2 * alpha_max
    elif T <= 6.0:
        alpha_at_5Tg_coeff = (0.2 ** gamma) * eta2
        current_alpha_coeff = (alpha_at_5Tg_coeff - eta1 * (T - 5.0 * Tg)) * alpha_max
    else:
        alpha_at_5Tg_coeff = (0.2 ** gamma) * eta2
        current_alpha_coeff = (alpha_at_5Tg_coeff - eta1 * (6.0 - 5.0 * Tg)) * alpha_max

    # 确保不小于最小值
    current_alpha_coeff = max(current_alpha_coeff, 0.20 * alpha_max)
    current_alpha_coeff = max(0.0, current_alpha_coeff)

    return current_alpha_coeff * g


def generate_response_spectrum_data():
    """生成反应谱数据点（高密度平滑版）"""

    # 1. 分段生成周期点 (使用列表推导式)
    # 0.00 ~ 0.09 (步长 0.01)
    p1 = [round(i * 0.01, 3) for i in range(0, 10)]

    # 0.10 ~ 0.98 (步长 0.02)
    p2 = [round(i * 0.02, 3) for i in range(5, 50)]

    # 1.00 ~ 6.00 (步长 0.05，为了足够平滑建议设为0.05或0.1)
    p3 = [round(i * 0.05, 3) for i in range(20, 121)]

    # 2. 合并列表并加入特征周期 Tg
    raw_periods = p1 + p2 + p3 + [RS_CHARACTERISTIC_PERIOD]

    # 3. 去重并排序
    rs_periods = sorted(list(set(raw_periods)))

    # 4. 计算对应的反应谱值（以g为单位）
    rs_values = []
    for t in rs_periods:
        val = china_response_spectrum(
            T=t,
            zeta=RS_DAMPING_RATIO,
            alpha_max=RS_BASE_ACCEL_G,
            Tg=RS_CHARACTERISTIC_PERIOD,
            g=GRAVITY_ACCEL
        )
        # 结果保留6位小数，并转换为g
        rs_values.append(round(val / GRAVITY_ACCEL, 6))

    return rs_periods, rs_values


def define_response_spectrum_functions_in_etabs():
    """在ETABS中定义反应谱函数"""
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    # 动态导入API对象
    from common.etabs_api_loader import get_api_objects
    ETABSv1, System, COMException = get_api_objects()

    if System is None:
        sys.exit("System对象未正确加载，无法定义反应谱函数")

    print("\n定义反应谱函数...")

    # 生成反应谱数据
    rs_periods, rs_values = generate_response_spectrum_data()

    # 使用数据库表方式定义反应谱函数
    db = sap_model.DatabaseTables
    key = "Functions - Response Spectrum - User Defined"
    fields = arr(["Name", "Period", "Value", "Damping Ratio"], System.String)

    # 准备数据
    data_py = []
    for i, p in enumerate(rs_periods):
        data_py.extend([
            RS_FUNCTION_NAME,
            str(round(p, 4)),
            str(round(rs_values[i], 6)),
            str(RS_DAMPING_RATIO)
        ])

    check_ret(
        db.SetTableForEditingArray(
            key,
            System.Int32(0),
            fields,
            System.Int32(len(rs_periods)),
            arr(data_py, System.String)
        ),
        f"SetTableForEditingArray({RS_FUNCTION_NAME})"
    )

    # 应用编辑的表格
    nfe, ne, nw, ni, log = System.Int32(0), System.Int32(0), System.Int32(0), System.Int32(0), System.String("")
    ret_apply = db.ApplyEditedTables(True, nfe, ne, nw, ni, log)
    check_ret(ret_apply[0], f"ApplyEditedTables({RS_FUNCTION_NAME})")

    if ret_apply[1] > 0:
        sys.exit("反应谱函数定义失败 (致命错误)。")

    print(f"反应谱函数 '{RS_FUNCTION_NAME}' 定义成功。")
    print(f"周期范围: {min(rs_periods):.2f}s - {max(rs_periods):.2f}s，共{len(rs_periods)}个数据点")
    print(f"最大地震影响系数: {RS_BASE_ACCEL_G}")
    print(f"场地特征周期: {RS_CHARACTERISTIC_PERIOD}s")
    print(f"阻尼比: {RS_DAMPING_RATIO}")


# 导出函数列表
__all__ = [
    'china_response_spectrum',
    'generate_response_spectrum_data',
    'define_response_spectrum_functions_in_etabs'
]

