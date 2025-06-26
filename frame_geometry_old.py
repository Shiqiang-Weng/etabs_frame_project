#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
框架几何创建模块（API调用签名修复版）
创建框架柱、框架梁、楼板等结构构件
包含节点隔板设置、梁惯性矩修正功能和底部约束设置
梁位置已调整到梁高中心线位置
底部约束自动从铰接修改为刚接（已完全修复所有API兼容性问题）

🔧 关键修复内容：
- 修复GetAllPoints/GetNameList API调用签名不匹配问题
- 正确处理ETABS 22及更早版本的ByRef参数要求
- 新增pythonnet System类型支持的API调用方式
- 多级备用策略确保节点获取成功
- 添加节点缓存机制和ETABS版本检测

优化特性：
- 自动单位系统设置和验证
- 健壮的底部节点获取算法
- 详细的调试和错误处理
- 多重备用方案确保可靠性
- 修复GetNameList API兼容性问题
- 修复函数名错误问题
- 正确的ByRef参数调用签名
"""

import logging
from typing import List, Tuple, Dict
from etabs_setup import get_etabs_objects
from utility_functions import check_ret, add_frame_by_coord_custom, add_area_by_coord_custom
from config import (
    NUM_GRID_LINES_X, NUM_GRID_LINES_Y, SPACING_X, SPACING_Y,
    NUM_STORIES, TYPICAL_STORY_HEIGHT, BOTTOM_STORY_HEIGHT,
    FRAME_BEAM_SECTION_NAME, FRAME_COLUMN_SECTION_NAME, SLAB_SECTION_NAME,
    FRAME_BEAM_HEIGHT
)

# 设置日志
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)


# ========== 通用API兼容性工具函数 (已按官方指南修复) ==========

def _get_all_points_safe(point_obj, csys="Global") -> Tuple[int, List, List, List, List]:
    """
    兼容所有ETABS版本的GetAllPoints方法（已按官方指南修复）。
    优先尝试新版API，失败后回退到使用 pythonnet 和 ByRef .NET 数组的旧版API调用方式。
    能够正确处理旧版API返回打包元组的情况。

    Parameters:
    ----------
    point_obj : PointObj
        ETABS点对象
    csys : str
        坐标系名称，默认"Global"

    Returns:
    -------
    tuple
        (return_code, pt_names, pt_x, pt_y, pt_z)
        return_code: 0 表示成功
    """
    # ① 优先尝试新版接口 (可能直接返回数据元组)
    try:
        ret, names, xs, ys, zs = point_obj.GetAllPoints(csys)
        if ret == 0 and names:
            log.debug("新版 GetAllPoints(csys) 接口调用成功。")
            return ret, list(names), list(xs), list(ys), list(zs)
    except TypeError:
        log.debug("新版 GetAllPoints(csys) 接口不适用 (TypeError)，回退至 ByRef .NET 数组方式。")
        pass
    except Exception as e:
        log.debug(f"新版 GetAllPoints(csys) 接口调用异常: {e}", exc_info=True)
        pass

    # ② 使用 pythonnet + ByRef .NET 数组 (针对旧版API的核心修复)
    try:
        from etabs_api_loader import get_api_objects
        _, System, _ = get_api_objects()

        n_max = 20000
        num_dummy = System.Int32(0)
        names_arr = System.Array[System.String]([None] * n_max)
        X = System.Array[float]([0.0] * n_max)
        Y = System.Array[float]([0.0] * n_max)
        Z = System.Array[float]([0.0] * n_max)

        ret = point_obj.GetAllPoints(num_dummy, names_arr, X, Y, Z, csys)

        # CRITICAL FIX: 处理旧版API返回打包元组的情况
        if isinstance(ret, tuple):
            # 返回格式: (ret_code, count, names_arr, X, Y, Z)
            ret_code, count = ret[0], ret[1]
            # 有些实现可能不返回数组，所以安全地重新赋值
            if len(ret) > 2:
                names_arr, X, Y, Z = ret[2:6]
        else:
            # 传统返回格式: 单个整数
            ret_code = ret
            count = int(num_dummy)

        if ret_code == 0 and count > 0:
            log.debug(f"ByRef .NET 数组方式成功，获取到 {count} 个节点。")
            return (
                ret_code,
                list(names_arr)[:count],
                list(X)[:count],
                list(Y)[:count],
                list(Z)[:count],
            )
        else:
            log.warning(f"ByRef .NET 数组方式调用失败或未获取到节点 (ret_code={ret_code}, count={count})")

    except ImportError:
        log.error("无法导入 etabs_api_loader 或 System 模块，无法使用 ByRef .NET 数组方式。")
    except Exception as e:
        log.error("ByRef .NET 数组方式获取节点时发生严重错误: %s", e, exc_info=True)

    # ③ 如果所有方法都失败，返回失败状态和空结果
    log.warning("所有 GetAllPoints 调用方式均失败，返回空结果。")
    return 1, [], [], [], []


def _get_name_list_safe(obj) -> List[str]:
    """
    兼容所有ETABS版本的GetNameList方法（已按官方指南修复）。
    正确处理ByRef参数调用签名和打包元组返回值。

    Parameters:
    ----------
    obj : object
        ETABS对象 (PointObj, AreaObj, FrameObj等)

    Returns:
    -------
    List[str]
        名称列表
    """
    # ① 优先尝试新版接口 (返回元组)
    try:
        ret, names = obj.GetNameList()
        if ret == 0:
            log.debug(f"新版 GetNameList() 接口调用成功 (对象: {type(obj).__name__})。")
            return list(names)
    except TypeError:
        log.debug(f"新版 GetNameList() 接口不适用 (TypeError)，回退至 ByRef .NET 数组方式。")
        pass
    except Exception as e:
        log.debug(f"新版 GetNameList() 接口调用异常: {e}", exc_info=True)
        pass

    # ② 使用 pythonnet + ByRef .NET 数组 (针对旧版API的核心修复)
    try:
        from etabs_api_loader import get_api_objects
        _, System, _ = get_api_objects()

        n_max = 50000
        num_dummy = System.Int32(0)
        MyName = System.Array[System.String]([None] * n_max)

        ret = obj.GetNameList(num_dummy, MyName)

        # CRITICAL FIX: 处理旧版API返回打包元组的情况
        if isinstance(ret, tuple):
            # 返回格式: (ret_code, count, MyName_arr)
            ret_code, count = ret[0], ret[1]
            if len(ret) > 2:
                MyName = ret[2]
        else:
            # 传统返回格式: 单个整数
            ret_code = ret
            count = int(num_dummy)

        if ret_code == 0:
            log.debug(f"ByRef .NET 数组方式 GetNameList 成功，获取到 {count} 个名称。")
            return list(MyName)[:count]
        else:
            log.warning(f"ByRef .NET 数组方式 GetNameList 调用失败 (ret_code={ret_code})")

    except ImportError:
        log.error("无法导入 etabs_api_loader 或 System 模块，无法使用 ByRef .NET 数组方式。")
    except Exception as e:
        log.error("ByRef .NET 数组方式获取名称列表时发生严重错误: %s", e, exc_info=True)

    log.warning(f"所有 GetNameList 调用方式都失败 (对象: {type(obj).__name__})")
    return []


def _get_all_point_names(point_obj) -> List[str]:
    """
    获取所有节点名称的封装函数
    """
    # 优先使用修复后的GetAllPoints方法，因为它一次性获取所有数据，效率更高
    ret, pt_names, _, _, _ = _get_all_points_safe(point_obj)
    if ret == 0 and pt_names:
        return pt_names

    # 如果GetAllPoints失败，再尝试GetNameList作为备用
    log.debug("GetAllPoints 未返回名称，尝试 _get_name_list_safe 作为备用。")
    return _get_name_list_safe(point_obj)


# ========== 单位系统管理和调试功能 ==========

def ensure_model_units():
    """
    确保模型使用正确的单位系统 (kN-m)
    使用整数比较而不是字符串比较，提高可靠性
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        log.error("SapModel 未初始化")
        return False

    try:
        # 获取当前单位
        current_units = sap_model.GetPresentUnits()
        log.info(f"当前模型单位: {current_units}")

        # 使用整数比较，避免字符串转换问题
        KNM_ENUM = 6  # ETABS eUnits.kN_m_C

        if current_units == KNM_ENUM:
            log.info("✓ 单位已是 kN-m")
            return True

        # 动态导入eUnits枚举
        try:
            from etabs_api_loader import get_api_objects
            ETABSv1, System, COMException = get_api_objects()

            # 使用枚举值设置单位
            ret_code = sap_model.SetPresentUnits(ETABSv1.eUnits.kN_m_C)
            if ret_code == 0:
                log.info("✓ 成功设置模型单位为 kN-m")
                return True
            else:
                log.warning(f"设置单位失败，返回码: {ret_code}")
                return False

        except ImportError:
            # 如果无法导入枚举，尝试直接使用整数值
            log.info("无法导入eUnits枚举，尝试使用数值...")
            try:
                ret_code = sap_model.SetPresentUnits(KNM_ENUM)
                if ret_code == 0:
                    log.info("✓ 成功设置模型单位为 kN-m (使用数值)")
                    return True
                else:
                    log.warning(f"设置单位失败，返回码: {ret_code}")
                    return False
            except Exception as e2:
                log.error(f"数值方法也失败: {e2}")
                return False

    except Exception as e:
        log.error(f"设置单位时发生异常: {e}")
        return False


def debug_joint_coordinates(max_joints=10):
    """
    调试函数：打印前几个节点的坐标信息
    使用修复后的GetAllPoints封装
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return

    log.info("=== 节点坐标调试信息 ===")

    point_obj = sap_model.PointObj

    try:
        # 获取当前单位和版本信息
        current_units = sap_model.GetPresentUnits()
        log.info(f"当前模型单位: {current_units}")

        try:
            version_info = sap_model.GetVersion()
            if len(version_info) >= 4:
                major, minor, build, rev = version_info[:4]
                log.info(f"ETABS版本: {major}.{minor} build {build}")
        except:
            log.info("无法获取ETABS版本信息")

        # 使用修复后的GetAllPoints方法
        log.debug("使用修复后的GetAllPoints方法获取节点信息...")
        ret, pt_names, pt_x, pt_y, pt_z = _get_all_points_safe(point_obj)
        number_pts = len(pt_names)

        if ret == 0 and number_pts > 0:
            log.info(f"模型中共有 {number_pts} 个节点")
            log.info(f"显示前 {min(max_joints, number_pts)} 个节点的坐标:")

            for i in range(min(max_joints, number_pts)):
                joint_name = pt_names[i]
                x, y, z = pt_x[i], pt_y[i], pt_z[i]
                log.info(f"  {joint_name}: ({x:.4f}, {y:.4f}, {z:.4f})")

            return  # 成功获取，直接返回

        else:
            log.warning(f"GetAllPoints调用失败或返回0个节点 (ret={ret})")

        # 备用方法: 使用GetNameList + GetCoordCartesian
        log.debug("尝试备用方法: GetNameList + GetCoordCartesian...")
        all_joints = _get_name_list_safe(point_obj)

        if not all_joints:
            log.warning("模型中没有节点或无法获取节点列表")
            return

        log.info(f"通过GetNameList获取到 {len(all_joints)} 个节点")
        log.info(f"显示前 {min(max_joints, len(all_joints))} 个节点的坐标:")

        for i, joint_name in enumerate(all_joints[:max_joints]):
            try:
                x_ref, y_ref, z_ref = [0.0], [0.0], [0.0]
                coord_ret = point_obj.GetCoordCartesian(joint_name, x_ref, y_ref, z_ref)

                if coord_ret[0] == 0:
                    x, y, z = x_ref[0], y_ref[0], z_ref[0]
                    log.info(f"  {joint_name}: ({x:.4f}, {y:.4f}, {z:.4f})")
                else:
                    log.warning(f"  {joint_name}: 获取坐标失败")

            except Exception as e:
                log.error(f"  {joint_name}: 异常 - {e}")

    except Exception as e:
        log.error(f"调试过程中发生异常: {e}")


# ========== 结构构件创建函数 ==========

def create_frame_columns() -> List[str]:
    """
    创建框架柱

    Returns:
    -------
    List[str]
        创建的柱名称列表
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return []

    log.info("创建框架柱...")
    frame_obj = sap_model.FrameObj

    # 计算网格坐标
    grid_x = [i * SPACING_X for i in range(NUM_GRID_LINES_X)]
    grid_y = [i * SPACING_Y for i in range(NUM_GRID_LINES_Y)]

    column_names = []
    cum_z = 0.0  # 累积高度

    for story in range(NUM_STORIES):
        story_num = story + 1
        story_height = TYPICAL_STORY_HEIGHT if story > 0 else BOTTOM_STORY_HEIGHT

        z_bottom = cum_z
        z_top = cum_z + story_height
        cum_z = z_top

        log.info(f"创建第 {story_num} 层柱 (标高: {z_bottom:.1f}m → {z_top:.1f}m)")

        story_column_count = 0

        # 在每个网格交点创建柱
        for i, x_coord in enumerate(grid_x):
            for j, y_coord in enumerate(grid_y):
                column_name = f"COL_X{i}_Y{j}_S{story_num}"

                ret_code, actual_name = add_frame_by_coord_custom(
                    frame_obj, x_coord, y_coord, z_bottom,
                    x_coord, y_coord, z_top,
                    FRAME_COLUMN_SECTION_NAME, column_name
                )

                check_ret(ret_code, f"AddByCoord(Column {column_name})")
                column_names.append(actual_name or column_name)
                story_column_count += 1

        log.info(f"第 {story_num} 层完成: {story_column_count} 根柱")

    log.info(f"框架柱创建完成，共 {len(column_names)} 根柱")
    return column_names


def create_frame_beams() -> List[str]:
    """
    创建框架梁（梁顶部与楼层顶部对齐）

    Returns:
    -------
    List[str]
        创建的梁名称列表
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return []

    log.info("创建框架梁...")
    frame_obj = sap_model.FrameObj

    # 使用配置文件中定义的梁截面高度
    beam_height = FRAME_BEAM_HEIGHT
    beam_half_height = beam_height / 2.0

    log.info(f"梁截面高度: {beam_height:.3f}m (来自配置文件)")
    log.info(f"梁顶部与楼层顶部、柱顶部对齐")
    log.info(f"梁中心线位于楼层顶部下方: {beam_half_height:.3f}m")

    # 计算网格坐标
    grid_x = [i * SPACING_X for i in range(NUM_GRID_LINES_X)]
    grid_y = [i * SPACING_Y for i in range(NUM_GRID_LINES_Y)]

    beam_names = []
    cum_z = 0.0  # 累积高度

    for story in range(NUM_STORIES):
        story_num = story + 1
        story_height = TYPICAL_STORY_HEIGHT if story > 0 else BOTTOM_STORY_HEIGHT

        z_level_top = cum_z + story_height  # 楼层顶部标高（柱顶、梁顶、板顶对齐）
        z_beam_center = z_level_top - beam_half_height  # 梁中心线标高
        cum_z = z_level_top

        log.info(f"创建第 {story_num} 层梁")
        log.debug(f"楼层顶标高: {z_level_top:.3f}m (柱顶、梁顶、板顶)")
        log.debug(f"梁中心标高: {z_beam_center:.3f}m")

        story_beam_count = 0

        # X方向梁（沿X轴方向）
        for j in range(NUM_GRID_LINES_Y):  # Y方向的每条轴线
            for i in range(NUM_GRID_LINES_X - 1):  # X方向相邻轴线间
                x1, x2 = grid_x[i], grid_x[i + 1]
                y_coord = grid_y[j]

                beam_name = f"BEAM_X_X{i}to{i + 1}_Y{j}_S{story_num}"

                ret_code, actual_name = add_frame_by_coord_custom(
                    frame_obj, x1, y_coord, z_beam_center,
                    x2, y_coord, z_beam_center,
                    FRAME_BEAM_SECTION_NAME, beam_name
                )

                check_ret(ret_code, f"AddByCoord(Beam {beam_name})")
                beam_names.append(actual_name or beam_name)
                story_beam_count += 1

        # Y方向梁（沿Y轴方向）
        for i in range(NUM_GRID_LINES_X):  # X方向的每条轴线
            for j in range(NUM_GRID_LINES_Y - 1):  # Y方向相邻轴线间
                x_coord = grid_x[i]
                y1, y2 = grid_y[j], grid_y[j + 1]

                beam_name = f"BEAM_Y_X{i}_Y{j}to{j + 1}_S{story_num}"

                ret_code, actual_name = add_frame_by_coord_custom(
                    frame_obj, x_coord, y1, z_beam_center,
                    x_coord, y2, z_beam_center,
                    FRAME_BEAM_SECTION_NAME, beam_name
                )

                check_ret(ret_code, f"AddByCoord(Beam {beam_name})")
                beam_names.append(actual_name or beam_name)
                story_beam_count += 1

        log.info(f"第 {story_num} 层完成: {story_beam_count} 根梁")

    log.info(f"框架梁创建完成，共 {len(beam_names)} 根梁")
    log.info(f"梁位置: 梁顶部与柱顶部、楼板顶部对齐在同一水平面")
    return beam_names


def create_slabs() -> List[str]:
    """
    创建楼板

    Returns:
    -------
    List[str]
        创建的楼板名称列表
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return []

    log.info("创建楼板...")
    area_obj = sap_model.AreaObj

    # 计算网格坐标
    grid_x = [i * SPACING_X for i in range(NUM_GRID_LINES_X)]
    grid_y = [i * SPACING_Y for i in range(NUM_GRID_LINES_Y)]

    slab_names = []
    cum_z = 0.0  # 累积高度

    for story in range(NUM_STORIES):
        story_num = story + 1
        story_height = TYPICAL_STORY_HEIGHT if story > 0 else BOTTOM_STORY_HEIGHT

        z_level = cum_z + story_height  # 楼板标高（层顶）
        cum_z = z_level

        log.info(f"创建第 {story_num} 层楼板 (标高: {z_level:.1f}m)")

        story_slab_count = 0

        # 在每个网格区域创建楼板
        for i in range(NUM_GRID_LINES_X - 1):
            for j in range(NUM_GRID_LINES_Y - 1):
                x1, x2 = grid_x[i], grid_x[i + 1]
                y1, y2 = grid_y[j], grid_y[j + 1]

                # 定义楼板四个角点（逆时针）
                slab_x = [x1, x2, x2, x1]
                slab_y = [y1, y1, y2, y2]
                slab_z = [z_level] * 4

                slab_name = f"SLAB_X{i}_Y{j}_S{story_num}"

                ret_code, actual_name = add_area_by_coord_custom(
                    area_obj, 4, slab_x, slab_y, slab_z,
                    SLAB_SECTION_NAME, slab_name
                )

                check_ret(ret_code, f"AddByCoord(Slab {slab_name})")
                final_name = actual_name or slab_name
                slab_names.append(final_name)

                # 为楼板分配半刚性楼面约束
                check_ret(
                    area_obj.SetDiaphragm(final_name, "SRD"),
                    f"SetDiaphragm({final_name}, SRD)"
                )

                story_slab_count += 1

        log.info(f"第 {story_num} 层完成: {story_slab_count} 块楼板")

    log.info(f"楼板创建完成，共 {len(slab_names)} 块楼板")
    return slab_names


# ========== 结构修正和约束设置函数 ==========

def apply_slab_membrane_modifiers(slab_names: List[str]):
    """
    为楼板设置膜单元修正系数，将面外刚度设为0

    Parameters:
    ----------
    slab_names : List[str]
        楼板名称列表
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        log.error("SapModel 未初始化，无法设置楼板膜单元修正。")
        return

    # 动态导入API对象
    try:
        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()
    except:
        log.warning("无法导入ETABS API对象，跳过楼板修正。")
        return

    area_obj = sap_model.AreaObj

    if not slab_names:
        log.info("未提供楼板名称列表。")
        return

    log.info("=== 楼板膜单元修正设置 ===")
    log.info(f"将为 {len(slab_names)} 块楼板设置膜单元修正")
    log.info(f"设置面外刚度为0，保持面内刚度不变")

    # 准备修正系数数组：面内刚度保持1.0，面外刚度设为0
    from utility_functions import arr
    modifiers_membrane = arr([
        1.0,  # f11 膜刚度X方向 - 保持
        1.0,  # f22 膜刚度Y方向 - 保持
        1.0,  # f12 膜剪切刚度 - 保持
        0.0,  # f13 横向剪切刚度XZ - 设为0
        0.0,  # f23 横向剪切刚度YZ - 设为0
        0.0,  # f33 弯曲刚度Z方向 - 设为0
        1.0,  # m11 质量X方向 - 保持
        1.0,  # m22 质量Y方向 - 保持
        1.0,  # m12 质量XY耦合 - 保持
        1.0,  # m13 质量XZ耦合 - 保持
        1.0,  # m23 质量YZ耦合 - 保持
        1.0,  # m33 质量Z方向 - 保持
        1.0  # weight 重量 - 保持
    ])

    successful_count = 0
    failed_count = 0
    failed_names = []

    log.info(f"正在应用膜单元修正系数...")

    for slab_name in slab_names:
        try:
            ret_tuple = area_obj.SetModifiers(slab_name, modifiers_membrane)
            ret_code = ret_tuple[0] if isinstance(ret_tuple, tuple) else ret_tuple

            if ret_code in (0, 1):
                successful_count += 1
            else:
                failed_count += 1
                failed_names.append(slab_name)
                log.warning(f"楼板 '{slab_name}' 设置失败，返回码: {ret_code}")

        except Exception as e:
            failed_count += 1
            failed_names.append(slab_name)
            log.error(f"楼板 '{slab_name}' 设置异常: {e}")

    # 强制刷新模型视图
    try:
        sap_model.View.RefreshView(0, False)
        log.info("模型视图已刷新")
    except Exception as e:
        log.error(f"刷新视图失败: {e}")

    # 输出结果统计
    log.info(f"楼板膜单元修正完成:")
    log.info(f"  成功处理: {successful_count} 块楼板")
    log.info(f"  处理失败: {failed_count} 块楼板")
    log.info(f"  面内刚度: f11 = f22 = f12 = 1.0 (保持)")
    log.info(f"  面外刚度: f13 = f23 = f33 = 0.0 (清零)")
    log.info(f"  工程意义: 楼板仅传递面内力，不传递弯矩")

    if failed_names:
        log.warning(f"失败的楼板 (前5个): {failed_names[:5]}")


def assign_diaphragm_constraints_by_story(column_names: List[str], beam_names: List[str], slab_names: List[str]):
    """
    为每层楼板设置隔板约束D1

    Parameters:
    ----------
    column_names : List[str]
        柱名称列表
    beam_names : List[str]
        梁名称列表
    slab_names : List[str]
        楼板名称列表
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        log.error("SapModel 未初始化，无法设置隔板约束。")
        return

    log.info("=== 设置楼板隔板约束 D1 ===")

    # 按楼层分组楼板
    story_slabs = {}

    # 处理楼板
    for slab_name in slab_names:
        if "_S" in slab_name:
            story_num = int(slab_name.split("_S")[-1])
            if story_num not in story_slabs:
                story_slabs[story_num] = []
            story_slabs[story_num].append(slab_name)

    successful_count = 0
    failed_count = 0
    failed_names = []

    area_obj = sap_model.AreaObj

    # 为每层楼板设置隔板约束
    for story_num in sorted(story_slabs.keys()):
        slabs = story_slabs[story_num]
        log.info(f"第 {story_num} 层: {len(slabs)} 块楼板")

        story_success = 0
        story_failed = 0

        # 为每块楼板设置隔板约束
        for slab_name in slabs:
            try:
                # 尝试使用SetDiaphragm方法设置隔板约束
                ret_code = area_obj.SetDiaphragm(slab_name, "D1")

                if ret_code in (0, 1):
                    successful_count += 1
                    story_success += 1
                else:
                    failed_count += 1
                    story_failed += 1
                    failed_names.append(slab_name)

            except Exception as e:
                failed_count += 1
                story_failed += 1
                failed_names.append(slab_name)
                if story_failed == 1:  # 只在第一次失败时打印错误详情
                    log.error(f"楼板设置异常: {e}")

        if story_success > 0:
            log.info(f"成功设置 {story_success} 块楼板隔板约束D1")
        if story_failed > 0:
            log.warning(f"失败设置 {story_failed} 块楼板")

    # 强制刷新模型视图
    try:
        sap_model.View.RefreshView(0, False)
        log.info("模型视图已刷新")
    except Exception as e:
        log.error(f"刷新视图失败: {e}")

    log.info(f"楼板隔板约束设置完成:")
    log.info(f"  成功处理: {successful_count} 块楼板")
    log.info(f"  处理失败: {failed_count} 块楼板")
    log.info(f"  约束类型: D1 (刚性隔板)")
    log.info(f"  工程意义: 确保楼层内刚体位移协调")

    if failed_names:
        log.warning(f"失败的楼板 (前5个): {failed_names[:5]}")


def apply_beam_inertia_modifiers(beam_names: List[str]):
    """
    为梁设置惯性矩修正系数
    边梁（x和y方向第一轴和最后一个轴）：3轴惯性矩放大1.5倍
    中梁：3轴惯性矩放大2倍

    Parameters:
    ----------
    beam_names : List[str]
        梁名称列表
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        log.error("SapModel 未初始化，无法设置梁惯性矩修正。")
        return

    from utility_functions import arr
    frame_obj = sap_model.FrameObj

    log.info("=== 梁惯性矩修正设置 ===")
    log.info(f"边梁（首末轴线）：3轴惯性矩 × 1.5")
    log.info(f"中梁（其他轴线）：3轴惯性矩 × 2.0")

    # 确定边界轴线索引
    first_x_axis = 0
    last_x_axis = NUM_GRID_LINES_X - 1
    first_y_axis = 0
    last_y_axis = NUM_GRID_LINES_Y - 1

    log.info(f"X方向边轴线: {first_x_axis}, {last_x_axis}")
    log.info(f"Y方向边轴线: {first_y_axis}, {last_y_axis}")

    # 准备修正系数数组
    # 边梁修正系数（3轴惯性矩×1.5）
    modifiers_edge = arr([
        1.0,  # Cross sectional area
        1.0,  # Shear area in direction 2
        1.0,  # Shear area in direction 3
        1.0,  # Torsional constant
        1.0,  # Moment of inertia about 2-axis
        1.5,  # Moment of inertia about 3-axis (放大1.5倍)
        1.0,  # Mass per unit length
        1.0  # Weight per unit length
    ])

    # 中梁修正系数（3轴惯性矩×2）
    modifiers_middle = arr([
        1.0,  # Cross sectional area
        1.0,  # Shear area in direction 2
        1.0,  # Shear area in direction 3
        1.0,  # Torsional constant
        1.0,  # Moment of inertia about 2-axis
        2.0,  # Moment of inertia about 3-axis (放大2倍)
        1.0,  # Mass per unit length
        1.0  # Weight per unit length
    ])

    edge_beam_count = 0
    middle_beam_count = 0
    failed_count = 0
    failed_names = []

    log.info(f"正在分析并设置 {len(beam_names)} 根梁的惯性矩修正...")

    for beam_name in beam_names:
        try:
            is_edge_beam = False

            # 判断是否为边梁
            if "BEAM_X_" in beam_name:
                # X方向梁，检查Y轴坐标
                parts = beam_name.split("_")
                for part in parts:
                    if part.startswith("Y") and part[1:].isdigit():
                        y_index = int(part[1:])
                        if y_index == first_y_axis or y_index == last_y_axis:
                            is_edge_beam = True
                        break

            elif "BEAM_Y_" in beam_name:
                # Y方向梁，检查X轴坐标
                parts = beam_name.split("_")
                for part in parts:
                    if part.startswith("X") and part[1:].isdigit():
                        x_index = int(part[1:])
                        if x_index == first_x_axis or x_index == last_x_axis:
                            is_edge_beam = True
                        break

            # 应用相应的修正系数
            if is_edge_beam:
                ret_tuple = frame_obj.SetModifiers(beam_name, modifiers_edge)
                edge_beam_count += 1
            else:
                ret_tuple = frame_obj.SetModifiers(beam_name, modifiers_middle)
                middle_beam_count += 1

            ret_code = ret_tuple[0] if isinstance(ret_tuple, tuple) else ret_tuple

            if ret_code not in (0, 1):
                failed_count += 1
                failed_names.append(beam_name)
                log.warning(f"梁 '{beam_name}' 设置失败，返回码: {ret_code}")

        except Exception as e:
            failed_count += 1
            failed_names.append(beam_name)
            log.error(f"梁 '{beam_name}' 设置异常: {e}")

    # 强制刷新模型视图
    try:
        sap_model.View.RefreshView(0, False)
        log.info("模型视图已刷新")
    except Exception as e:
        log.error(f"刷新视图失败: {e}")

    # 输出结果统计
    log.info(f"梁惯性矩修正完成:")
    log.info(f"  边梁处理: {edge_beam_count} 根 (3轴惯性矩 × 1.5)")
    log.info(f"  中梁处理: {middle_beam_count} 根 (3轴惯性矩 × 2.0)")
    log.info(f"  处理失败: {failed_count} 根")
    log.info(f"  工程意义: 考虑楼板对梁刚度的贡献")

    if failed_names:
        log.warning(f"失败的梁 (前5个): {failed_names[:5]}")


# ========== 底部约束相关函数（完全修复版本） ==========

def get_base_level_joints_v2(tolerance=0.001) -> List[str]:
    """
    改进的底部节点获取方法（版本2）
    使用修复后的GetAllPoints方法，正确处理ByRef参数

    Parameters:
    ----------
    tolerance : float
        Z坐标容差（米）

    Returns:
    -------
    List[str]
        底部节点名称列表
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        log.error("SapModel 未初始化")
        return []

    log.info("=== 改进的底部节点获取方法 V2 ===")
    log.info(f"容差设置: {tolerance}m")

    point_obj = sap_model.PointObj
    joint_coords = {}
    z_coordinates = []

    # 方法1: 使用修复后的GetAllPoints方法
    log.debug("尝试使用修复后的GetAllPoints方法...")
    ret, pt_names, pt_x, pt_y, pt_z = _get_all_points_safe(point_obj)
    number_pts = len(pt_names)

    if ret == 0 and number_pts > 0:
        log.info(f"通过GetAllPoints获取到 {number_pts} 个节点")
        for i in range(number_pts):
            joint_coords[pt_names[i]] = (pt_x[i], pt_y[i], pt_z[i])
            z_coordinates.append(pt_z[i])
    else:
        if ret != 0:
            log.warning(f"GetAllPoints 调用失败，返回码: {ret}。尝试备用方法...")
        else:
            log.warning("GetAllPoints返回0个节点，尝试备用方法...")

    # 方法2: 如果GetAllPoints失败，尝试GetNameList + GetCoordCartesian
    if not joint_coords:
        log.debug("尝试GetNameList + GetCoordCartesian方法...")
        all_joint_names = _get_name_list_safe(point_obj)

        if not all_joint_names:
            log.warning("无法通过任何方法获取节点列表")
            return []

        log.info(f"通过GetNameList成功获取 {len(all_joint_names)} 个节点名称")
        for joint_name in all_joint_names:
            try:
                x_ref, y_ref, z_ref = [0.0], [0.0], [0.0]
                coord_ret = point_obj.GetCoordCartesian(joint_name, x_ref, y_ref, z_ref)
                if coord_ret[0] == 0:
                    x, y, z = x_ref[0], y_ref[0], z_ref[0]
                    joint_coords[joint_name] = (x, y, z)
                    z_coordinates.append(z)
            except Exception as e:
                log.error(f"获取节点 {joint_name} 坐标失败: {e}")
                continue

    if not z_coordinates:
        log.warning("无法获取任何节点的坐标")
        return []

    # 找到最小Z坐标（底部标高）
    z_min = min(z_coordinates)
    log.info(f"找到最低标高: {z_min:.4f}m")

    # 筛选底部节点
    base_joints = [
        name for name, (_, _, z) in joint_coords.items() if abs(z - z_min) <= tolerance
    ]
    log.info(f"识别到 {len(base_joints)} 个底部节点")
    for joint in base_joints[:5]:
        log.debug(f"  - 底部节点: {joint} at Z={joint_coords[joint][2]:.4f}")

    return base_joints


def get_base_level_joints_by_grid_direct() -> List[str]:
    """
    通过已知网格坐标直接查找底部节点（最可靠的备用方法）
    不依赖任何GetNameList API调用

    Returns:
    -------
    List[str]
        底部节点名称列表
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return []

    log.info("=== 通过网格坐标直接查找底部节点 ===")

    # 计算预期的网格坐标
    grid_x = [i * SPACING_X for i in range(NUM_GRID_LINES_X)]
    grid_y = [i * SPACING_Y for i in range(NUM_GRID_LINES_Y)]

    log.info(f"预期网格坐标:")
    log.info(f"  X: {grid_x}")
    log.info(f"  Y: {grid_y}")

    point_obj = sap_model.PointObj
    base_joints = []
    tolerance = 0.1  # 10cm容差

    # 在每个预期的网格交点查找节点
    for i, x_coord in enumerate(grid_x):
        for j, y_coord in enumerate(grid_y):
            try:
                # 使用GetNameAtCoord方法查找节点
                ret_tuple = point_obj.GetNameAtCoord(x_coord, y_coord, 0.0, tolerance)

                if ret_tuple[0] == 0:  # 找到节点
                    joint_name = ret_tuple[1]
                    if joint_name and joint_name not in base_joints:
                        base_joints.append(joint_name)
                        log.debug(f"找到节点: {joint_name} at grid ({i}, {j}) -> ({x_coord:.1f}, {y_coord:.1f}, 0.0)")

            except Exception as e:
                # 如果GetNameAtCoord也失败，尝试其他方法
                try:
                    # 尝试使用更宽松的容差
                    ret_tuple = point_obj.GetNameAtCoord(x_coord, y_coord, 0.0, tolerance * 5)
                    if ret_tuple[0] == 0:
                        joint_name = ret_tuple[1]
                        if joint_name and joint_name not in base_joints:
                            base_joints.append(joint_name)
                            log.debug(f"找到节点(宽松): {joint_name} at grid ({i}, {j})")
                except:
                    continue

    log.info(f"通过网格坐标直接查找到 {len(base_joints)} 个底部节点")
    return base_joints


def get_base_level_joints_by_existing_elements() -> List[str]:
    """
    通过已创建的结构构件获取底部节点（终极备用方案）
    利用已知的柱名称模式来推断底部节点
    根据官方文档使用正确的ByRef调用方式

    Returns:
    -------
    List[str]
        底部节点名称列表
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return []

    log.info("=== 通过结构构件获取底部节点 ===")

    frame_obj = sap_model.FrameObj
    base_joints = []

    try:
        # 方法1: 获取所有框架单元名称，然后筛选第一层柱
        log.debug("尝试获取所有框架单元名称...")
        frame_names = _get_name_list_safe(frame_obj)

        if frame_names:
            log.info(f"成功获取 {len(frame_names)} 个框架单元")
            first_story_columns = [name for name in frame_names if "COL_" in name and "_S1" in name]
            log.info(f"找到 {len(first_story_columns)} 根第一层柱")

            for column_name in first_story_columns:
                try:
                    pt1, pt2 = [""], [""]
                    ret_code = frame_obj.GetPoints(column_name, pt1, pt2)
                    if ret_code[0] == 0 and pt1[0] and pt2[0]:
                        point_obj = sap_model.PointObj
                        x1_ref, y1_ref, z1_ref = [0.0], [0.0], [0.0]
                        coord1_ret = point_obj.GetCoordCartesian(pt1[0], x1_ref, y1_ref, z1_ref)
                        x2_ref, y2_ref, z2_ref = [0.0], [0.0], [0.0]
                        coord2_ret = point_obj.GetCoordCartesian(pt2[0], x2_ref, y2_ref, z2_ref)

                        if coord1_ret[0] == 0 and coord2_ret[0] == 0:
                            bottom_joint = pt1[0] if z1_ref[0] <= z2_ref[0] else pt2[0]
                            if bottom_joint not in base_joints:
                                base_joints.append(bottom_joint)
                except Exception as e:
                    log.debug(f"处理柱 {column_name} 失败: {e}")
                    continue
        else:
            log.warning("获取框架单元列表失败")

        log.info(f"通过结构构件获取到 {len(base_joints)} 个底部节点")
        return base_joints

    except Exception as e:
        log.error(f"通过结构构件获取节点失败: {e}")
        return []


def get_all_points_reference_method(include_restraints=False) -> List[tuple]:
    """
    基于参考代码的get_all_points函数实现
    使用修复后的GetAllPoints方法获取所有节点信息

    Parameters:
    ----------
    include_restraints : bool
        是否包含约束信息

    Returns:
    -------
    List[tuple]
        节点信息列表，每个元素为 (节点名, x, y, z, [约束信息])
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        return []

    try:
        point_obj = sap_model.PointObj
        ret, pt_names, pt_x, pt_y, pt_z = _get_all_points_safe(point_obj)
        number_pts = len(pt_names)

        if ret != 0 or number_pts == 0:
            log.warning(f"GetAllPoints调用失败或返回0个节点 (ret={ret})")
            return []

        log.info(f"通过GetAllPoints获取到 {number_pts} 个节点")

        points = []
        for i in range(number_pts):
            point_data = (pt_names[i], pt_x[i], pt_y[i], pt_z[i])
            if include_restraints:
                try:
                    restraint_data = point_obj.GetRestraint(pt_names[i])
                    restraints = restraint_data[1] if restraint_data[0] == 0 else [False] * 6
                    point_data += (restraints,)
                except:
                    point_data += ([False] * 6,)
            points.append(point_data)

        return points

    except Exception as e:
        log.error(f"GetAllPoints参考方法失败: {e}")
        return []


def get_base_level_joints_reference_method(tolerance=0.001) -> List[str]:
    """
    基于参考代码实现的底部节点获取方法
    使用修复后的GetAllPoints作为主要获取方式

    Parameters:
    ----------
    tolerance : float
        Z坐标容差（米）

    Returns:
    -------
    List[str]
        底部节点名称列表
    """
    log.info("=== 基于参考代码的底部节点获取方法 ===")
    all_points = get_all_points_reference_method(include_restraints=False)

    if not all_points:
        log.warning("无法获取任何节点信息")
        return []

    log.info(f"获取到 {len(all_points)} 个节点")
    z_coordinates = [point[3] for point in all_points]
    if not z_coordinates:
        return []
    z_min = min(z_coordinates)
    log.info(f"找到最低标高: {z_min:.4f}m")

    base_joints = [p[0] for p in all_points if abs(p[3] - z_min) <= tolerance]
    log.info(f"识别到 {len(base_joints)} 个底部节点")
    return base_joints


def get_base_level_joints_by_grid() -> List[str]:
    """
    通过改进方法获取底部节点（兼容原函数名）
    使用多级备用策略确保获取成功，优先使用最可靠的方法。
    """
    ensure_model_units()

    # 策略1: 基于修复后的GetAllPoints和坐标分析 (最通用和可靠)
    log.info("尝试策略1: 基于修复后的GetAllPoints和坐标分析...")
    base_joints = get_base_level_joints_reference_method(0.001)

    # 策略2: 通过已创建的结构构件获取
    if not base_joints:
        log.warning("策略1失败，尝试策略2: 通过结构构件获取...")
        base_joints = get_base_level_joints_by_existing_elements()

    # 策略3: 网格直接查找方法
    if not base_joints:
        log.warning("策略2失败，尝试策略3: 网格直接查找...")
        base_joints = get_base_level_joints_by_grid_direct()

    if base_joints:
        log.info(f"✓ 成功获取到 {len(base_joints)} 个底部节点")
        expected_count = NUM_GRID_LINES_X * NUM_GRID_LINES_Y
        if len(base_joints) == expected_count:
            log.info(f"✓ 节点数量正确，预期 {expected_count} 个，实际 {len(base_joints)} 个")
        else:
            log.warning(f"⚠ 节点数量异常，预期 {expected_count} 个，实际 {len(base_joints)} 个")
    else:
        log.error("✗ 所有方法都失败，无法获取底部节点")

    return base_joints


def get_base_level_joints() -> List[str]:
    """
    获取底部基础层的所有节点名称（优化版本）

    Returns:
    -------
    List[str]
        底部节点名称列表
    """
    log.info("获取底部基础层节点...")
    log.info("  使用优化的多策略底部节点识别方法...")
    return get_base_level_joints_by_grid()


def set_rigid_base_constraints_improved(joint_names: List[str]) -> Tuple[int, int]:
    """
    改进的底部刚接约束设置

    Parameters:
    ----------
    joint_names : List[str]
        需要设置约束的节点名称列表

    Returns:
    -------
    Tuple[int, int]
        (成功数量, 失败数量)
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None:
        log.error("SapModel 未初始化")
        return 0, 0

    if not joint_names:
        log.warning("未提供节点列表用于设置约束")
        return 0, 0

    log.info("=== 改进的底部刚接约束设置 ===")
    log.info(f"将为 {len(joint_names)} 个节点设置刚接约束")
    log.info(f"约束: UX=UY=UZ=RX=RY=RZ=True")

    point_obj = sap_model.PointObj
    restraint_rigid = [True, True, True, True, True, True]
    successful_count, failed_count = 0, 0
    failed_details = []

    for joint_name in joint_names:
        try:
            # CRITICAL FIX: 准备接收可能是元组的返回值
            ret = point_obj.SetRestraint(joint_name, restraint_rigid)

            # 解包返回值
            if isinstance(ret, tuple):
                ret_code = ret[0]
            else:
                ret_code = ret

            if ret_code == 0:
                successful_count += 1
                log.debug(f"✓ {joint_name}: 约束设置成功")
            else:
                failed_count += 1
                failed_details.append(f"{joint_name}(返回码:{ret_code})")
                log.warning(f"✗ {joint_name}: 设置失败，返回码: {ret_code}")
        except Exception as e:
            failed_count += 1
            failed_details.append(f"{joint_name}(异常:{str(e)[:50]})")
            log.error(f"✗ {joint_name}: 发生异常: {e}")

    try:
        sap_model.View.RefreshView(0, False)
        log.info("模型视图已刷新")
    except Exception as e:
        log.error(f"刷新视图失败: {e}")

    log.info(f"约束设置结果:")
    log.info(f"  成功: {successful_count}/{len(joint_names)} 个节点")
    log.info(f"  失败: {failed_count}/{len(joint_names)} 个节点")

    if failed_details:
        log.warning(f"失败详情: {'; '.join(failed_details[:3])}...")

    return successful_count, failed_count


def set_rigid_base_constraints_fixed(joint_names: List[str]) -> Tuple[int, int]:
    """
    为指定节点设置刚接约束（兼容原函数名）
    """
    return set_rigid_base_constraints_improved(joint_names)


def verify_constraints_with_getrestraint(joint_names: List[str]):
    """
    使用GetRestraint方法验证约束设置
    """
    my_etabs, sap_model = get_etabs_objects()
    if sap_model is None: return

    log.info("=== 使用GetRestraint验证约束设置 (抽查前5个) ===")
    point_obj = sap_model.PointObj

    for joint_name in joint_names[:5]:
        try:
            value = [False] * 6
            # GetRestraint 也可能返回元组
            ret = point_obj.GetRestraint(joint_name, value)

            if isinstance(ret, tuple):
                ret_code = ret[0]
                # 如果元组包含更新后的值，可以这样获取
                if len(ret) > 1:
                    value = list(ret[1])
            else:
                ret_code = ret

            if ret_code == 0:
                status = "固定" if all(value) else "部分或无约束"
                log.info(f"  节点 {joint_name}: {status} - {value}")
            else:
                log.warning(f"  节点 {joint_name}: 获取约束失败，返回码: {ret_code}")
        except Exception as e:
            log.error(f"  节点 {joint_name}: 验证异常: {e}")


def fix_base_constraints_comprehensive() -> Tuple[int, int]:
    """
    综合修复底部约束问题的主函数

    Returns:
    -------
    Tuple[int, int]
        (成功数量, 失败数量)
    """
    log.info("=" * 60)
    log.info("综合修复底部约束问题")
    log.info("=" * 60)

    log.info("步骤1: 检查并设置模型单位...")
    if not ensure_model_units():
        log.warning("⚠ 单位设置可能有问题，继续尝试...")

    log.info("步骤2: 调试当前节点信息...")
    debug_joint_coordinates(5)

    log.info("步骤3: 获取底部节点 (采用多策略方法)...")
    base_joints = get_base_level_joints()

    if not base_joints:
        log.error("✗ 无法获取任何底部节点，约束设置中止。")
        return 0, 0

    log.info(f"✓ 成功获取 {len(base_joints)} 个底部节点")

    log.info("步骤4: 设置底部约束...")
    success_count, fail_count = set_rigid_base_constraints_improved(base_joints)

    if success_count > 0:
        log.info("步骤5: 验证约束设置...")
        verify_constraints_with_getrestraint(base_joints)

    log.info("=" * 60)
    log.info(f"修复完成 - 成功设置 {success_count} 个刚接，失败 {fail_count} 个。")
    log.info("=" * 60)

    return success_count, fail_count


def fix_base_constraints_issue() -> Tuple[int, int]:
    """
    修复底部约束设置问题的主函数（兼容原函数名）
    """
    return fix_base_constraints_comprehensive()


# ========== 主要结构创建函数 ==========

def create_frame_structure() -> Tuple[List[str], List[str], List[str], Dict[int, float]]:
    """
    创建完整的框架结构（包含底部约束设置） - 完全优化版本

    Returns:
    -------
    Tuple[List[str], List[str], List[str], Dict[int, float]]
        (柱名称列表, 梁名称列表, 楼板名称列表, 楼层高度字典)
    """
    log.info("=" * 60)
    log.info("开始创建框架结构 - 完全优化版本")
    log.info("=" * 60)

    log.info("步骤0: 设置模型单位系统...")
    if not ensure_model_units():
        log.warning("⚠ 单位设置失败，但继续执行...")

    log.info("步骤1: 创建结构构件...")
    column_names = create_frame_columns()
    beam_names = create_frame_beams()
    slab_names = create_slabs()

    log.info("步骤2: 计算楼层高度...")
    story_heights = {}
    cum_height = 0
    for story in range(NUM_STORIES):
        story_num = story + 1
        height = TYPICAL_STORY_HEIGHT if story > 0 else BOTTOM_STORY_HEIGHT
        cum_height += height
        story_heights[story_num] = cum_height
    log.info(f"楼层高度配置: {story_heights}")

    log.info("步骤3: 应用结构修正...")
    apply_slab_membrane_modifiers(slab_names)
    assign_diaphragm_constraints_by_story(column_names, beam_names, slab_names)
    apply_beam_inertia_modifiers(beam_names)

    log.info("步骤4: 设置底部约束 (使用完全修复的方法)...")
    success_count, fail_count = fix_base_constraints_comprehensive()

    if success_count > 0:
        log.info(f"✓ 底部约束设置成功: {success_count} 个节点")
    else:
        log.error(f"✗ 底部约束设置失败: {fail_count} 个节点")

    log.info("=" * 60)
    log.info("框架结构创建完成")
    log.info(f"构件统计: {len(column_names)} 根柱, {len(beam_names)} 根梁, {len(slab_names)} 块楼板")
    log.info(f"约束统计: {success_count} 个底部约束成功, {fail_count} 个失败")
    log.info("=" * 60)

    return column_names, beam_names, slab_names, story_heights


# ========== 导出函数列表 ==========

__all__ = [
    'create_frame_structure',
    'ensure_model_units',
    'get_base_level_joints_v2',
    'get_base_level_joints_by_grid_direct',
    'get_base_level_joints_by_existing_elements',
    'get_base_level_joints_reference_method',
    'get_all_points_reference_method',
    'fix_base_constraints_comprehensive',
    'debug_joint_coordinates',
    '_get_all_points_safe',
    '_get_name_list_safe',
]