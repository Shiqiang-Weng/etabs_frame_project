#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Legacy configuration globals plus structured `SETTINGS` for ETABS frame workflows.
Existing code continues to rely on module-level constants; new code should prefer `SETTINGS`.
"""

import os
from dataclasses import dataclass
from typing import Optional

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

# ========== 结构网格与层高参数 ==========
NUM_GRID_LINES_X = 5
NUM_GRID_LINES_Y = 3
SPACING_X = 6.0
SPACING_Y = 6.0
NUM_STORIES = 10
TYPICAL_STORY_HEIGHT = 3.0
BOTTOM_STORY_HEIGHT = 3.0

# ========== 构件参数 ==========
FRAME_BEAM_WIDTH = 0.4
FRAME_BEAM_HEIGHT = 0.7
FRAME_COLUMN_WIDTH = 0.6
FRAME_COLUMN_HEIGHT = 0.6
SLAB_THICKNESS = 0.15

# ========== 材料与截面名称 ==========
CONCRETE_MATERIAL_NAME = "C30/37"
CONCRETE_E_MODULUS = 30000000
CONCRETE_POISSON = 0.2
CONCRETE_THERMAL_EXP = 1.0e-5
CONCRETE_UNIT_WEIGHT = 26.0

FRAME_BEAM_SECTION_NAME = "FB400X700"
FRAME_COLUMN_SECTION_NAME = "FC600X600"
SLAB_SECTION_NAME = "Slab-150"

# ========== 地震参数 (GB50011-2010) ==========
RS_DESIGN_INTENSITY = 7
RS_BASE_ACCEL_G = 0.08
RS_SEISMIC_GROUP = 3
RS_SITE_CLASS = "III"
RS_CHARACTERISTIC_PERIOD = 0.65
RS_FUNCTION_NAME = f"UserRSFunc_GB50011_G{RS_SEISMIC_GROUP}_{RS_SITE_CLASS}_{RS_DESIGN_INTENSITY}deg"
MODAL_CASE_NAME = "MODAL_RS"
RS_DAMPING_RATIO = 0.05
GRAVITY_ACCEL = 9.80665
GENERATE_RS_COMBOS = True

# ========== 荷载参数 ==========
DEFAULT_DEAD_SUPER_SLAB = 5
DEFAULT_LIVE_LOAD_SLAB = 2.0
DEFAULT_FINISH_LOAD_BEAM = 0.5

# ========== 全局变量初始化 ==========
ETABSv1 = None
System = None
COMException = None

# ========== Concrete Design Parameters (GB50010-2010) ==========
PERFORM_CONCRETE_DESIGN = True
CONCRETE_DESIGN_CODE = "GB 50010-2010(2015)"
EXPORT_ALL_DESIGN_FILES = False

# ---------------------------------------------------------------------------
# Optional structured settings (non-breaking): exposes the above constants via
# typed dataclasses. Existing call sites can continue using globals.
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class PathsConfig:
    use_net_core: bool
    program_path: str
    dll_path: str
    script_directory: str
    model_path: str


@dataclass(frozen=True)
class GridConfig:
    num_grid_lines_x: int
    num_grid_lines_y: int
    spacing_x: float
    spacing_y: float
    num_stories: int
    typical_story_height: float
    bottom_story_height: float


@dataclass(frozen=True)
class SectionsConfig:
    frame_beam_width: float
    frame_beam_height: float
    frame_beam_section_name: str
    frame_column_width: float
    frame_column_height: float
    frame_column_section_name: str
    slab_thickness: float
    slab_section_name: str
    concrete_material_name: str
    concrete_e_modulus: float
    concrete_poisson: float
    concrete_thermal_exp: float
    concrete_unit_weight: float


@dataclass(frozen=True)
class LoadsConfig:
    default_dead_super_slab: float
    default_live_load_slab: float
    default_finish_load_beam: float


@dataclass(frozen=True)
class ResponseSpectrumConfig:
    modal_case_name: str
    rs_function_name: str
    rs_damping_ratio: float
    rs_base_accel_g: float
    rs_site_class: str
    rs_seismic_group: int
    rs_characteristic_period: float
    generate_rs_combos: bool
    gravity_accel: float


@dataclass(frozen=True)
class DesignConfig:
    perform_concrete_design: bool
    export_all_design_files: bool


@dataclass(frozen=True)
class Settings:
    paths: PathsConfig
    grid: GridConfig
    sections: SectionsConfig
    loads: LoadsConfig
    response_spectrum: ResponseSpectrumConfig
    design: DesignConfig


SETTINGS = Settings(
    paths=PathsConfig(
        use_net_core=USE_NET_CORE,
        program_path=PROGRAM_PATH,
        dll_path=ETABS_DLL_PATH,
        script_directory=SCRIPT_DIRECTORY,
        model_path=MODEL_PATH,
    ),
    grid=GridConfig(
        num_grid_lines_x=NUM_GRID_LINES_X,
        num_grid_lines_y=NUM_GRID_LINES_Y,
        spacing_x=SPACING_X,
        spacing_y=SPACING_Y,
        num_stories=NUM_STORIES,
        typical_story_height=TYPICAL_STORY_HEIGHT,
        bottom_story_height=BOTTOM_STORY_HEIGHT,
    ),
    sections=SectionsConfig(
        frame_beam_width=FRAME_BEAM_WIDTH,
        frame_beam_height=FRAME_BEAM_HEIGHT,
        frame_beam_section_name=FRAME_BEAM_SECTION_NAME,
        frame_column_width=FRAME_COLUMN_WIDTH,
        frame_column_height=FRAME_COLUMN_HEIGHT,
        frame_column_section_name=FRAME_COLUMN_SECTION_NAME,
        slab_thickness=SLAB_THICKNESS,
        slab_section_name=SLAB_SECTION_NAME,
        concrete_material_name=CONCRETE_MATERIAL_NAME,
        concrete_e_modulus=CONCRETE_E_MODULUS,
        concrete_poisson=CONCRETE_POISSON,
        concrete_thermal_exp=CONCRETE_THERMAL_EXP,
        concrete_unit_weight=CONCRETE_UNIT_WEIGHT,
    ),
    loads=LoadsConfig(
        default_dead_super_slab=DEFAULT_DEAD_SUPER_SLAB,
        default_live_load_slab=DEFAULT_LIVE_LOAD_SLAB,
        default_finish_load_beam=DEFAULT_FINISH_LOAD_BEAM,
    ),
    response_spectrum=ResponseSpectrumConfig(
        modal_case_name=MODAL_CASE_NAME,
        rs_function_name=RS_FUNCTION_NAME,
        rs_damping_ratio=RS_DAMPING_RATIO,
        rs_base_accel_g=RS_BASE_ACCEL_G,
        rs_site_class=RS_SITE_CLASS,
        rs_seismic_group=RS_SEISMIC_GROUP,
        rs_characteristic_period=RS_CHARACTERISTIC_PERIOD,
        generate_rs_combos=GENERATE_RS_COMBOS,
        gravity_accel=GRAVITY_ACCEL,
    ),
    design=DesignConfig(
        perform_concrete_design=PERFORM_CONCRETE_DESIGN,
        export_all_design_files=EXPORT_ALL_DESIGN_FILES,
    ),
)
