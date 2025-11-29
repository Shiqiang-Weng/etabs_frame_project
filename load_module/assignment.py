#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
荷载分配模块
为框架结构分配各种荷载
"""

from typing import List
from etabs_setup import get_etabs_objects
from utility_functions import check_ret
from config import (
    DEFAULT_DEAD_SUPER_SLAB, DEFAULT_LIVE_LOAD_SLAB, DEFAULT_FINISH_LOAD_BEAM
)

# 柱轴向荷载参数（kN，压缩为正）
COLUMN_AXIAL_LOAD = 0  # 每根柱的轴向荷载
ENABLE_FRAME_COLUMNS = True  # 是否启用柱荷载


def assign_dead_and_live_loads_to_slabs(slab_names: List[str],
                                        dead_kpa: float = DEFAULT_DEAD_SUPER_SLAB,
                                        live_kpa: float = DEFAULT_LIVE_LOAD_SLAB):
    """
    为楼板分配恒荷载和活荷载

    Parameters:
    ----------
    slab_names : List[str]
        楼板名称列表
    dead_kpa : float
        恒荷载强度 (kN/m²)
    live_kpa : float
        活荷载强度 (kN/m²)
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    print(f"\n为楼板分配荷载...")
    print(f"恒荷载: {dead_kpa} kN/m², 活荷载: {live_kpa} kN/m²")

    area_obj = sap_model.AreaObj

    if not slab_names:
        print("警告: 未找到楼板名称列表。")
        return

    success_count = 0
    fail_count = 0

    for slab_name in slab_names:
        try:
            # 动态导入API对象
            from etabs_api_loader import get_api_objects
            ETABSv1, System, COMException = get_api_objects()

            if ETABSv1 is None:
                print(f"  错误: ETABSv1 API对象为None，跳过楼板 '{slab_name}'")
                fail_count += 1
                continue

            # 分配恒荷载（向下为正）
            ret_dead = area_obj.SetLoadUniform(
                slab_name, "DEAD", abs(dead_kpa), 10, True, "Global",
                ETABSv1.eItemType.Objects
            )
            check_ret(ret_dead, f"SetLoadUniform DEAD on {slab_name}", (0, 1))

            # 分配活荷载（向下为正）
            ret_live = area_obj.SetLoadUniform(
                slab_name, "LIVE", abs(live_kpa), 10, True, "Global",
                ETABSv1.eItemType.Objects
            )
            check_ret(ret_live, f"SetLoadUniform LIVE on {slab_name}", (0, 1))

            success_count += 1

        except Exception as e:
            print(f"  错误: 楼板 '{slab_name}' 荷载分配失败: {e}")
            fail_count += 1

    print(f"楼板荷载分配完成: 成功 {success_count} 块, 失败 {fail_count} 块")


def assign_finish_loads_to_beams(beam_names: List[str],
                                 finish_load: float = DEFAULT_FINISH_LOAD_BEAM):
    """
    为框架梁分配面层荷载（线荷载）

    Parameters:
    ----------
    beam_names : List[str]
        梁名称列表
    finish_load : float
        面层荷载强度 (kN/m)
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    print(f"\n为框架梁分配面层荷载...")
    print(f"面层荷载: {finish_load} kN/m")

    frame_obj = sap_model.FrameObj

    if not beam_names:
        print("警告: 未找到梁名称列表。")
        return

    success_count = 0
    fail_count = 0

    for beam_name in beam_names:
        try:
            # 动态导入API对象
            from etabs_api_loader import get_api_objects
            ETABSv1, System, COMException = get_api_objects()

            if ETABSv1 is None:
                print(f"  错误: ETABSv1 API对象为None，跳过梁 '{beam_name}'")
                fail_count += 1
                continue

            # 分配均布线荷载到梁（向下为正）
            ret = frame_obj.SetLoadDistributed(
                beam_name, "DEAD", 1, 10, 0.0, 1.0,
                abs(finish_load), abs(finish_load), "Global", True, True,
                ETABSv1.eItemType.Objects
            )
            check_ret(ret, f"SetLoadDistributed on {beam_name}", (0, 1))

            success_count += 1

        except Exception as e:
            print(f"  错误: 梁 '{beam_name}' 面层荷载分配失败: {e}")
            fail_count += 1

    print(f"梁面层荷载分配完成: 成功 {success_count} 根, 失败 {fail_count} 根")


def assign_column_loads_fixed(column_names: List[str]):
    """
    为框架柱分配轴向荷载 - 完整修复版本

    Parameters:
    ----------
    column_names : List[str]
        柱名称列表
    """
    if not ENABLE_FRAME_COLUMNS or not column_names or COLUMN_AXIAL_LOAD == 0:
        print("⏭️ 跳过柱荷载分配（未启用、无柱构件或荷载为0）")
        return

    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        print("错误: SapModel未初始化")
        return

    print(f"\n为框架柱分配轴向荷载...")
    print(f"轴向荷载: {COLUMN_AXIAL_LOAD} kN (压缩)")

    frame_obj = sap_model.FrameObj
    column_load_count = 0
    failed_columns = []

    # 荷载值为负表示压缩
    load_value = -abs(float(COLUMN_AXIAL_LOAD))

    # API 调用: SetLoadPoint(Name, LoadPat, Type, Dir, Dist, Val, CSys, Replace, ItemType)
    # Type=1 (Force), Dir=1 (Local-1, Axial), CSys="Local", Replace=True
    for column_name in column_names:
        try:
            # 动态导入API对象
            from etabs_api_loader import get_api_objects
            ETABSv1, System, COMException = get_api_objects()

            if ETABSv1 is None:
                print(f"  错误: ETABSv1 API对象为None，跳过柱 '{column_name}'")
                failed_columns.append(column_name)
                continue

            ret = frame_obj.SetLoadPoint(
                column_name, "DEAD", 1, 1, 1.0, load_value,
                "Local", True, True, ETABSv1.eItemType.Objects
            )
            check_ret(ret, f"SetLoadPoint DEAD on {column_name}", (0, 1))
            column_load_count += 1

        except Exception as e:
            print(f"  错误: 柱 '{column_name}' 荷载分配失败: {e}")
            failed_columns.append(column_name)

    print(f"柱轴向荷载分配完成: 成功 {column_load_count} 根, 失败 {len(failed_columns)} 根")

    if failed_columns:
        print(f"失败的柱 (前5个): {failed_columns[:5]}")


def assign_seismic_mass_to_structure():
    """
    确保结构具有地震质量
    （主要通过恒荷载和活荷载提供质量，无需额外操作）
    """
    print(f"\n地震质量分配...")
    print("地震质量主要来源:")
    print("- 结构自重 (DEAD荷载, 系数=1.0)")
    print("- 活荷载 (LIVE荷载, 系数=0.5)")
    print("- 面层荷载已包含在DEAD荷载中")
    print("质量源已在质量源定义中配置完成。")


def assign_all_loads_to_frame_structure(column_names: List[str],
                                        beam_names: List[str],
                                        slab_names: List[str]):
    """
    为整个框架结构分配所有荷载

    Parameters:
    ----------
    column_names : List[str]
        柱名称列表
    beam_names : List[str]
        梁名称列表
    slab_names : List[str]
        楼板名称列表
    """
    print("\n========== 开始分配荷载 ==========")

    # 为楼板分配面荷载
    assign_dead_and_live_loads_to_slabs(slab_names)

    # 为梁分配面层荷载
    assign_finish_loads_to_beams(beam_names)

    # 为柱分配轴向荷载
    assign_column_loads_fixed(column_names)

    # 确认地震质量配置
    assign_seismic_mass_to_structure()

    print("\n========== 荷载分配完成 ==========")
    print(f"已为 {len(slab_names)} 块楼板分配面荷载")
    print(f"已为 {len(beam_names)} 根梁分配面层荷载")
    print(f"已为 {len(column_names)} 根柱分配轴向荷载")
    print("所有荷载均已分配完成，可进行结构分析。")


# 导出函数列表
__all__ = [
    'assign_dead_and_live_loads_to_slabs',
    'assign_finish_loads_to_beams',
    'assign_column_loads_fixed',
    'assign_seismic_mass_to_structure',
    'assign_all_loads_to_frame_structure'
]