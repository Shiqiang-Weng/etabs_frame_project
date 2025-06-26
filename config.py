#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETABS 框架结构建模配置文件
Configuration settings for ETABS frame structure modeling
"""

import os

# ========== ETABS 路径配置 ==========
USE_NET_CORE = True
PROGRAM_PATH = r"C:\Program Files\Computers and Structures\ETABS 22\ETABS.exe"
ETABS_DLL_PATH = r"C:\Program Files\Computers and Structures\ETABS 22\ETABSv1.dll"
SCRIPT_DIRECTORY = r"C:\Users\Shiqi\Desktop\etabs_script_output_frame"

# ========== 模型文件配置 ==========
MODEL_NAME = "Frame_Model_10Story_v6_0_1.edb"
MODEL_PATH = os.path.join(SCRIPT_DIRECTORY, MODEL_NAME)

# ========== ETABS 连接配置 ==========
ATTACH_TO_INSTANCE = False
SPECIFY_PATH = False
REMOTE = False
REMOTE_COMPUTER = "YourRemoteComputerName"

# ========== 结构网格与楼层参数 ==========
NUM_GRID_LINES_X = 5  # X方向轴线数量
NUM_GRID_LINES_Y = 3  # Y方向轴线数量
SPACING_X = 6.0       # X方向轴线间距 (米)
SPACING_Y = 6.0       # Y方向轴线间距 (米)
NUM_STORIES = 10      # 楼层数
TYPICAL_STORY_HEIGHT = 3.0    # 标准层层高 (米)
BOTTOM_STORY_HEIGHT = 3.0     # 底层层高 (米)

# ========== 框架结构参数 ==========
# 框架梁参数
FRAME_BEAM_WIDTH = 0.4         # 框架梁宽度 (米)
FRAME_BEAM_HEIGHT = 0.7        # 框架梁高度 (米)

# 框架柱参数
FRAME_COLUMN_WIDTH = 0.6       # 框架柱宽度 (米)
FRAME_COLUMN_HEIGHT = 0.6      # 框架柱高度 (米)

# 楼板参数
SLAB_THICKNESS = 0.15          # 楼板厚度 (米)

# ========== 材料属性 ==========
CONCRETE_MATERIAL_NAME = "C30/37"
CONCRETE_E_MODULUS = 30000000      # 弹性模量 (kN/m²)
CONCRETE_POISSON = 0.2             # 泊松比
CONCRETE_THERMAL_EXP = 1.0e-5      # 热膨胀系数
CONCRETE_UNIT_WEIGHT = 26.0        # 容重 (kN/m³)

# ========== 截面属性名称 ==========
FRAME_BEAM_SECTION_NAME = "FB400X700"     # 框架梁截面
FRAME_COLUMN_SECTION_NAME = "FC600X600"   # 框架柱截面
SLAB_SECTION_NAME = "Slab-150"            # 楼板截面

# ========== 地震参数 (GB50011-2010) ==========
# 地震设防烈度：7 度（0.1 g）
RS_DESIGN_INTENSITY = 7
RS_BASE_ACCEL_G = 0.08        # 最大地震影响系数 αmax：0.0800
RS_SEISMIC_GROUP = 3          # 地震分组：第三组
RS_SITE_CLASS = "III"         # 场地类别：III 类
RS_CHARACTERISTIC_PERIOD = 0.65  # 场地特征周期 Tg：0.65 s

# 反应谱相关参数
RS_FUNCTION_NAME = f"UserRSFunc_GB50011_G{RS_SEISMIC_GROUP}_{RS_SITE_CLASS}_{RS_DESIGN_INTENSITY}deg"
MODAL_CASE_NAME = "MODAL_RS"
RS_DAMPING_RATIO = 0.05
GRAVITY_ACCEL = 9.80665

# 是否生成反应谱组合
GENERATE_RS_COMBOS = True

# ========== 荷载参数 ==========
DEFAULT_DEAD_SUPER_SLAB = 5    # 楼板恒荷载 (kN/m²)
DEFAULT_LIVE_LOAD_SLAB = 2.0      # 楼板活荷载 (kN/m²)
DEFAULT_FINISH_LOAD_BEAM = 0.5    # 梁面层荷载 (kN/m)

# ========== 全局变量初始化 ==========
ETABSv1 = None
System = None
COMException = None

# ========== Concrete Design Parameters (GB50010-2010) ==========
PERFORM_CONCRETE_DESIGN = True
CONCRETE_DESIGN_CODE = "GB 50010-2010(2015)" # Chinese Concrete Design Code