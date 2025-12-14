#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Configuration and design data structures for ETABS frame workflows."""

import os
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Mapping, Optional, Tuple

# ---------------------------- unit helpers ---------------------------------

MM_TO_M = 1e-3
DEFAULT_STORY_HEIGHT = 3.0

# ---------------------------- ETABS paths -----------------------------------

USE_NET_CORE = True
PROGRAM_PATH = r"C:\Program Files\Computers and Structures\ETABS 22\ETABS.exe"
ETABS_DLL_PATH = r"C:\Program Files\Computers and Structures\ETABS 22\ETABSv1.dll"
SCRIPT_DIRECTORY = r"C:\Users\Shiqi\Desktop\etabs_script_output_frame"
# 统一结果输出子目录（除模型文件外的所有导出文件）
DATA_EXTRACTION_DIR = os.path.join(SCRIPT_DIRECTORY, "data_extraction")
# 子目录：分析/设计结果
ANALYSIS_DATA_DIR = os.path.join(DATA_EXTRACTION_DIR, "analysis_data")
DESIGN_DATA_DIR = os.path.join(DATA_EXTRACTION_DIR, "design_data")

# ---------------------------- 模型文件配置 -----------------------------------

MODEL_NAME = "Frame_Model_10Story_v6_0_1.edb"
MODEL_PATH = os.path.join(SCRIPT_DIRECTORY, MODEL_NAME)

# ---------------------------- ETABS 连接配置 ---------------------------------

ATTACH_TO_INSTANCE = False
SPECIFY_PATH = False
REMOTE = False
REMOTE_COMPUTER = "YourRemoteComputerName"

# ---------------------------- 材料与截面名称 ---------------------------------

CONCRETE_MATERIAL_NAME = "C30/37"
CONCRETE_E_MODULUS = 30000000
CONCRETE_POISSON = 0.2
CONCRETE_THERMAL_EXP = 1.0e-5
CONCRETE_UNIT_WEIGHT = 26.0
# C30 concrete axial compressive strength design value (fc) per GB 50010, MPa (= N/mm^2)
CONCRETE_FC_MPA = 14.3

SLAB_THICKNESS = 0.10
SLAB_SECTION_NAME = "Slab-100"

# Legacy defaults kept for compatibility with diagnostics; new flow uses design-driven names.
FRAME_BEAM_SECTION_NAME = "FB250X500"
FRAME_COLUMN_SECTION_NAME = "FC600X600"

# ---------------------------- 地震参数 (GB50011-2010) ------------------------

RS_DESIGN_INTENSITY = 8
RS_BASE_ACCEL_G = 0.16
RS_SEISMIC_GROUP = 1
RS_SITE_CLASS = "III"
RS_CHARACTERISTIC_PERIOD = 0.45
RS_FUNCTION_NAME = f"UserRSFunc_GB50011_G{RS_SEISMIC_GROUP}_{RS_SITE_CLASS}_{RS_DESIGN_INTENSITY}deg"
MODAL_CASE_NAME = "MODAL_RS"
RS_DAMPING_RATIO = 0.05
GRAVITY_ACCEL = 9.80665
GENERATE_RS_COMBOS = True

# ---------------------------- 荷载参数 ---------------------------------------

DEFAULT_DEAD_SUPER_SLAB = 5
DEFAULT_LIVE_LOAD_SLAB = 2.0
DEFAULT_FINISH_LOAD_BEAM = 5

# ---------------------------- 设计参数 ---------------------------------------

PERFORM_CONCRETE_DESIGN = True
CONCRETE_DESIGN_CODE = "GB 50010-2010(2015)"
EXPORT_ALL_DESIGN_FILES = False
REANALYZE_BEFORE_DESIGN = False
ENABLE_LEGACY_DESIGN_EXPORT = False
DESIGN_DEBUG_LOGS = False

# ---------------------------- 基础配置 dataclasses ---------------------------


@dataclass(frozen=True)
class PathsConfig:
    use_net_core: bool
    program_path: str
    dll_path: str
    script_directory: str
    data_extraction_dir: str
    analysis_data_dir: str
    design_data_dir: str
    model_path: str


@dataclass(frozen=True)
class MaterialConfig:
    concrete_material_name: str
    concrete_e_modulus: float
    concrete_poisson: float
    concrete_thermal_exp: float
    concrete_unit_weight: float
    slab_thickness: float
    slab_section_name: str
    concrete_fc: Optional[float] = None


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
class DesignOptions:
    perform_concrete_design: bool
    export_all_design_files: bool
    reanalyze_before_design: bool
    enable_legacy_design_export: bool
    design_debug_logs: bool


@dataclass(frozen=True)
class Settings:
    paths: PathsConfig
    materials: MaterialConfig
    loads: LoadsConfig
    response_spectrum: ResponseSpectrumConfig
    design: DesignOptions


SETTINGS = Settings(
    paths=PathsConfig(
        use_net_core=USE_NET_CORE,
        program_path=PROGRAM_PATH,
        dll_path=ETABS_DLL_PATH,
        script_directory=SCRIPT_DIRECTORY,
        data_extraction_dir=DATA_EXTRACTION_DIR,
        analysis_data_dir=ANALYSIS_DATA_DIR,
        design_data_dir=DESIGN_DATA_DIR,
        model_path=MODEL_PATH,
    ),
    materials=MaterialConfig(
        concrete_material_name=CONCRETE_MATERIAL_NAME,
        concrete_e_modulus=CONCRETE_E_MODULUS,
        concrete_poisson=CONCRETE_POISSON,
        concrete_thermal_exp=CONCRETE_THERMAL_EXP,
        concrete_unit_weight=CONCRETE_UNIT_WEIGHT,
        slab_thickness=SLAB_THICKNESS,
        slab_section_name=SLAB_SECTION_NAME,
        concrete_fc=CONCRETE_FC_MPA,
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
    design=DesignOptions(
        perform_concrete_design=PERFORM_CONCRETE_DESIGN,
        export_all_design_files=EXPORT_ALL_DESIGN_FILES,
        reanalyze_before_design=REANALYZE_BEFORE_DESIGN,
        enable_legacy_design_export=ENABLE_LEGACY_DESIGN_EXPORT,
        design_debug_logs=DESIGN_DEBUG_LOGS,
    ),
)

# ---------------------------- 设计采样驱动的网格/截面 -------------------------


@dataclass(frozen=True)
class GridConfig:
    """平面网格配置：由拓扑参数 (n_x, n_y, l_x_mm, l_y_mm) 推导。"""

    num_x: int  # 网格线数量（x 方向）
    num_y: int  # 网格线数量（y 方向）
    spacing_x: float  # 相邻网格线间距 (m)
    spacing_y: float  # 相邻网格线间距 (m)

    @property
    def x_coords(self) -> List[float]:
        return [i * self.spacing_x for i in range(self.num_x)]

    @property
    def y_coords(self) -> List[float]:
        return [j * self.spacing_y for j in range(self.num_y)]

    def iter_points(self) -> Iterable[Tuple[int, int, float, float]]:
        """遍历所有网格交点 (i, j, x, y)。"""
        for i, x in enumerate(self.x_coords):
            for j, y in enumerate(self.y_coords):
                yield i, j, x, y


@dataclass(frozen=True)
class StoreyConfig:
    """楼层数量与层高；所有楼层层高一致。"""

    num_storeys: int  # N_st
    storey_height: float  # m

    @property
    def elevations(self) -> List[float]:
        # 0 层为基础，依次向上
        return [k * self.storey_height for k in range(self.num_storeys + 1)]


@dataclass(frozen=True)
class ColumnSectionGroup:
    corner_b: float  # m
    edge_b: float  # m
    interior_b: float  # m


@dataclass(frozen=True)
class BeamSectionGroup:
    edge_b: float  # m
    edge_h: float  # m
    interior_b: float  # m
    interior_h: float  # m


@dataclass(frozen=True)
class VerticalGroupSections:
    columns: ColumnSectionGroup
    beams: BeamSectionGroup


@dataclass(frozen=True)
class SectionsConfig:
    g1: VerticalGroupSections  # bottom floors
    g2: VerticalGroupSections  # middle floors
    g3: VerticalGroupSections  # top floors


def _default_group_mapping(num_storeys: int) -> Dict[str, Any]:
    """Bottom=1层，剩余楼层按中/顶部均分（向上取整给中部）。"""
    if num_storeys <= 0:
        return {"groups": {}, "story_to_group": {}}
    bottom = 1
    remaining = max(num_storeys - 1, 0)
    middle = remaining // 2 + (1 if remaining % 2 else 0)
    top = remaining - middle
    counts = [bottom, middle, top]
    labels = ["bottom", "middle", "top"]
    start = 1
    groups: Dict[str, Dict[str, Any]] = {}
    story_to_group: Dict[int, str] = {}
    for idx, count in enumerate(counts):
        group_name = f"Group{idx + 1}"
        stories = list(range(start, start + count))
        groups[group_name] = {"label": labels[idx], "stories": stories}
        for story in stories:
            story_to_group[story] = group_name
        start += count
    return {"groups": groups, "story_to_group": story_to_group}


def _normalize_group_name(name: str) -> str:
    digits = "".join(ch for ch in str(name) if ch.isdigit())
    return f"Group{digits}" if digits else str(name)


@dataclass(frozen=True)
class DesignConfig:
    """完整的单个采样 case 的建模配置。"""

    case_id: int
    grid: GridConfig
    storeys: StoreyConfig
    sections: SectionsConfig
    slab_thickness: float = 0.10  # m，保持原来的默认 0.10，可按需调整
    slab_section_name: Optional[str] = None
    group_mapping: Optional[Mapping[str, Any]] = None

    @property
    def slab_name(self) -> str:
        if self.slab_section_name:
            return self.slab_section_name
        return f"SLAB_{int(round(self.slab_thickness / MM_TO_M))}"

    @property
    def topology(self) -> Dict[str, int]:
        spans_x = self.grid.num_x - 1
        spans_y = self.grid.num_y - 1
        lx_mm = int(round(self.grid.spacing_x / MM_TO_M))
        ly_mm = int(round(self.grid.spacing_y / MM_TO_M))
        return {
            "N_st": self.storeys.num_storeys,
            "n_x": spans_x,
            "n_y": spans_y,
            "l_x": lx_mm,
            "l_y": ly_mm,
            "l_x_mm": lx_mm,
            "l_y_mm": ly_mm,
        }

    def story_group(self, story: int) -> Optional[str]:
        mapping = self.group_mapping or _default_group_mapping(self.storeys.num_storeys)
        story_to_group = mapping.get("story_to_group") if isinstance(mapping, dict) else None
        if story_to_group:
            if story in story_to_group:
                return story_to_group[story]
            if str(story) in story_to_group:
                return story_to_group[str(story)]
        if isinstance(mapping, dict) and story in mapping:
            return str(mapping[story])
        if isinstance(mapping, dict) and str(story) in mapping:
            return str(mapping[str(story)])
        # fallback by range
        defaults = _default_group_mapping(self.storeys.num_storeys)
        return defaults["story_to_group"].get(story)

    def _group_sections(self, group_name: str) -> VerticalGroupSections:
        normalized = _normalize_group_name(group_name)
        group_map = {
            "Group1": self.sections.g1,
            "Group2": self.sections.g2,
            "Group3": self.sections.g3,
        }
        if normalized not in group_map:
            raise KeyError(f"Unknown group name: {group_name}")
        return group_map[normalized]

    def _group_id(self, group_name: str) -> str:
        normalized = _normalize_group_name(group_name)
        digits = "".join(ch for ch in normalized if ch.isdigit())
        return digits or normalized

    def column_dims_and_name(self, group_name: str, position: str) -> Tuple[float, float, str]:
        group = self._group_sections(group_name)
        gid = self._group_id(group_name)
        position_upper = position.upper()
        if position_upper == "CORNER":
            width = group.columns.corner_b
        elif position_upper == "EDGE":
            width = group.columns.edge_b
        else:
            width = group.columns.interior_b
        name = f"C_G{gid}_{position_upper}_{int(round(width / MM_TO_M))}"
        return width, width, name

    def beam_dims_and_name(self, group_name: str, position: str) -> Tuple[float, float, str]:
        group = self._group_sections(group_name)
        gid = self._group_id(group_name)
        label = "EDGE" if position.upper() == "EDGE" else "INT"
        if label == "EDGE":
            width = group.beams.edge_b
            depth = group.beams.edge_h
        else:
            width = group.beams.interior_b
            depth = group.beams.interior_h
        name = f"B_G{gid}_{label}_{int(round(width / MM_TO_M))}x{int(round(depth / MM_TO_M))}"
        return width, depth, name

    def beam_depth_for_story(self, story: int, position: str) -> float:
        group = self.story_group(story) or "Group1"
        _, depth, _ = self.beam_dims_and_name(group, position)
        return depth

    def beam_section_name_for_story(self, story: int, position: str) -> str:
        group = self.story_group(story) or "Group1"
        _, _, name = self.beam_dims_and_name(group, position)
        return name

    def column_section_name_for_story(self, story: int, position: str) -> str:
        group = self.story_group(story) or "Group1"
        _, _, name = self.column_dims_and_name(group, position)
        return name

    def iter_frame_section_definitions(self) -> Iterable[Tuple[str, float, float]]:
        for gid, group in zip(["1", "2", "3"], [self.sections.g1, self.sections.g2, self.sections.g3]):
            for position, width in [
                ("CORNER", group.columns.corner_b),
                ("EDGE", group.columns.edge_b),
                ("INTERIOR", group.columns.interior_b),
            ]:
                name = f"C_G{gid}_{position}_{int(round(width / MM_TO_M))}"
                yield name, width, width

            for label, (width, depth) in [
                ("EDGE", (group.beams.edge_b, group.beams.edge_h)),
                ("INT", (group.beams.interior_b, group.beams.interior_h)),
            ]:
                name = f"B_G{gid}_{label}_{int(round(width / MM_TO_M))}x{int(round(depth / MM_TO_M))}"
                yield name, width, depth

    @classmethod
    def from_sample(cls, sample: Mapping[str, Any], story_height: float = DEFAULT_STORY_HEIGHT) -> "DesignConfig":
        """
        从一行采样数据 (dict / pandas Series) 构造 DesignConfig。
        支持 flat CSV（包含 N_st, n_x, n_y, l_x_mm, l_y_mm, C_G*, B_G* 字段）或结构化 Topology/Sizing。
        """

        def _get(mapping: Mapping[str, Any], keys: List[str]) -> Any:
            for key in keys:
                if key in mapping:
                    val = mapping[key]
                    if val is not None:
                        return val
                key_lower = str(key).lower()
                for candidate_key, candidate_val in mapping.items():
                    if str(candidate_key).lower() == key_lower and candidate_val is not None:
                        return candidate_val
            raise KeyError(keys[0])

        sample_dict = dict(sample)
        topo_src = sample_dict.get("Topology") if isinstance(sample_dict.get("Topology"), Mapping) else sample_dict
        sizing_src = sample_dict.get("Sizing") if isinstance(sample_dict.get("Sizing"), Mapping) else sample_dict

        n_st = int(_get(topo_src, ["N_st", "n_st", "N_ST"]))
        n_x = int(_get(topo_src, ["n_x", "Nx", "nX", "grid_x"]))
        n_y = int(_get(topo_src, ["n_y", "Ny", "nY", "grid_y"]))
        lx_mm = float(_get(topo_src, ["l_x_mm", "l_x", "Lx"]))
        ly_mm = float(_get(topo_src, ["l_y_mm", "l_y", "Ly"]))

        grid = GridConfig(
            num_x=n_x + 1,
            num_y=n_y + 1,
            spacing_x=lx_mm * MM_TO_M,
            spacing_y=ly_mm * MM_TO_M,
        )

        storeys = StoreyConfig(num_storeys=n_st, storey_height=story_height)

        def _col_group(prefix: str) -> ColumnSectionGroup:
            if isinstance(sizing_src, Mapping) and f"Group{prefix[-1]}" in sizing_src:
                group_data = sizing_src[f"Group{prefix[-1]}"]
                return ColumnSectionGroup(
                    corner_b=float(_get(group_data, [f"C_{prefix}_Corner_b", f"C_G{prefix[-1]}_Corner_b"])) * MM_TO_M,
                    edge_b=float(_get(group_data, [f"C_{prefix}_Edge_b", f"C_G{prefix[-1]}_Edge_b"])) * MM_TO_M,
                    interior_b=float(_get(group_data, [f"C_{prefix}_Interior_b", f"C_G{prefix[-1]}_Interior_b"])) * MM_TO_M,
                )
            return ColumnSectionGroup(
                corner_b=float(_get(sizing_src, [f"C_{prefix}_Corner_b"])) * MM_TO_M,
                edge_b=float(_get(sizing_src, [f"C_{prefix}_Edge_b"])) * MM_TO_M,
                interior_b=float(_get(sizing_src, [f"C_{prefix}_Interior_b"])) * MM_TO_M,
            )

        def _beam_group(prefix: str) -> BeamSectionGroup:
            if isinstance(sizing_src, Mapping) and f"Group{prefix[-1]}" in sizing_src:
                group_data = sizing_src[f"Group{prefix[-1]}"]
                return BeamSectionGroup(
                    edge_b=float(_get(group_data, [f"B_{prefix}_Edge_b", f"B_G{prefix[-1]}_Edge_b"])) * MM_TO_M,
                    edge_h=float(_get(group_data, [f"B_{prefix}_Edge_h", f"B_G{prefix[-1]}_Edge_h"])) * MM_TO_M,
                    interior_b=float(_get(group_data, [f"B_{prefix}_Interior_b", f"B_G{prefix[-1]}_Interior_b"])) * MM_TO_M,
                    interior_h=float(_get(group_data, [f"B_{prefix}_Interior_h", f"B_G{prefix[-1]}_Interior_h"])) * MM_TO_M,
                )
            return BeamSectionGroup(
                edge_b=float(_get(sizing_src, [f"B_{prefix}_Edge_b"])) * MM_TO_M,
                edge_h=float(_get(sizing_src, [f"B_{prefix}_Edge_h"])) * MM_TO_M,
                interior_b=float(_get(sizing_src, [f"B_{prefix}_Interior_b"])) * MM_TO_M,
                interior_h=float(_get(sizing_src, [f"B_{prefix}_Interior_h"])) * MM_TO_M,
            )

        sections = SectionsConfig(
            g1=VerticalGroupSections(columns=_col_group("G1"), beams=_beam_group("G1")),
            g2=VerticalGroupSections(columns=_col_group("G2"), beams=_beam_group("G2")),
            g3=VerticalGroupSections(columns=_col_group("G3"), beams=_beam_group("G3")),
        )

        case_id_raw = sample_dict.get("case_id") or sample_dict.get("num") or sample_dict.get("id") or 0
        try:
            case_id_val = int(case_id_raw)
        except Exception:
            case_id_val = 0

        mapping = sample_dict.get("GroupMapping") or _default_group_mapping(n_st)

        slab_name = sample_dict.get("slab_section_name") or sample_dict.get("SlabSectionName")
        slab_thickness = float(sample_dict.get("slab_thickness", sample_dict.get("SlabThickness", SLAB_THICKNESS)))

        return cls(
            case_id=case_id_val,
            grid=grid,
            storeys=storeys,
            sections=sections,
            slab_thickness=slab_thickness,
            slab_section_name=slab_name,
            group_mapping=mapping,
        )


def design_config_from_case(design_case: Any, story_height: float = DEFAULT_STORY_HEIGHT) -> DesignConfig:
    """
    统一入口：支持 param_sampling.DesignCaseConfig、平面 dict 或已构造的 DesignConfig。
    """
    if isinstance(design_case, DesignConfig):
        return design_case

    if hasattr(design_case, "topology") and hasattr(design_case, "sizing"):
        topo = getattr(design_case, "topology")
        sizing = getattr(design_case, "sizing")
        sample: Dict[str, Any] = {
            "case_id": getattr(design_case, "case_id", 0),
            "Topology": topo,
            "Sizing": sizing,
            "GroupMapping": getattr(design_case, "group_mapping", None),
        }
        return DesignConfig.from_sample(sample, story_height=story_height)

    if isinstance(design_case, Mapping):
        return DesignConfig.from_sample(design_case, story_height=story_height)

    raise TypeError(f"Unsupported design case type: {type(design_case)}")


# 默认设计（兼容旧单案例工作流）：5x3 网格、10 层、6000 mm 跨距、统一截面 250x500/600x600
DEFAULT_DESIGN_SAMPLE: Dict[str, Any] = {
    "case_id": 0,
    "N_st": 10,
    "n_x": 4,
    "n_y": 2,
    "l_x_mm": 6000,
    "l_y_mm": 6000,
    "C_G1_Corner_b": 600,
    "C_G1_Edge_b": 600,
    "C_G1_Interior_b": 600,
    "B_G1_Edge_b": 250,
    "B_G1_Edge_h": 500,
    "B_G1_Interior_b": 250,
    "B_G1_Interior_h": 500,
    "C_G2_Corner_b": 600,
    "C_G2_Edge_b": 600,
    "C_G2_Interior_b": 600,
    "B_G2_Edge_b": 250,
    "B_G2_Edge_h": 500,
    "B_G2_Interior_b": 250,
    "B_G2_Interior_h": 500,
    "C_G3_Corner_b": 600,
    "C_G3_Edge_b": 600,
    "C_G3_Interior_b": 600,
    "B_G3_Edge_b": 250,
    "B_G3_Edge_h": 500,
    "B_G3_Interior_b": 250,
    "B_G3_Interior_h": 500,
}

DEFAULT_DESIGN_CONFIG = DesignConfig.from_sample(DEFAULT_DESIGN_SAMPLE)

# 全局变量初始化（ETABS API 占位）
ETABSv1 = None
System = None
COMException = None
