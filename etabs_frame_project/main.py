#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Canonical ETABS frame pipeline runner with four stages:
1) 几何建模 (geometry_modeling)
2) 荷载定义与施加 (load_module)
3) 分析 / 设计求解 (analysis)
4) 结果提取 / 后处理 (results_extraction)
"""

import sys
import time
import traceback
from pathlib import Path
from typing import Optional

from common import config
from common.etabs_api_loader import load_dotnet_etabs_api
from common import etabs_setup, file_operations

import geometry_modeling
import load_module
import analysis
import results_extraction
from parametric_model.param_sampling import (
    DesignCaseConfig,
    generate_param_plan,
    load_param_plan,
)

PROJECT_ROOT = Path(__file__).resolve().parent
PLAN_DIR = PROJECT_ROOT / "parametric_model"
PLAN_PATH = PLAN_DIR / "param_plan.jsonl"
PLAN_SEED = 42
PLAN_CASES = 10


def prepare_design() -> DesignCaseConfig:
    """Generate/load plan (10 samples) and return case 0 design config."""
    PLAN_DIR.mkdir(parents=True, exist_ok=True)
    if not PLAN_PATH.exists():
        print(f"[参数化模式] 未找到方案文件，自动生成: {PLAN_PATH} (cases={PLAN_CASES})")
        generate_param_plan(num_cases=PLAN_CASES, out_path=PLAN_PATH, seed=PLAN_SEED)
    else:
        print(f"[参数化模式] 使用已存在的方案文件: {PLAN_PATH}")

    if not PLAN_PATH.exists():
        raise FileNotFoundError(f"期望的方案文件不存在: {PLAN_PATH}")

    samples = load_param_plan(PLAN_PATH)
    print(f"[参数化模式] 从 {PLAN_PATH} 读取到 {len(samples)} 个样本")
    if not samples:
        raise RuntimeError(f"方案文件 {PLAN_PATH} 中没有任何样本")

    case_index = 0
    sample = samples[case_index]
    case_id = sample.get("case_id", case_index)
    design_case = DesignCaseConfig.from_sample(case_id, sample)
    print(f"[参数化模式] 将运行 case_index={case_index}, case_id={design_case.case_id}")
    return design_case


def print_project_info(design: Optional[DesignCaseConfig] = None) -> None:
    """打印项目和脚本配置信息"""
    num_stories = design.topology["N_st"] if design else config.NUM_STORIES
    total_height = config.BOTTOM_STORY_HEIGHT + (num_stories - 1) * config.TYPICAL_STORY_HEIGHT
    print("=" * 80)
    print("ETABS 框架结构自动建模脚本 (规范四阶段流水线)")
    print("=" * 80)
    print("关键参数:")
    print(f"- 楼层数 {num_stories}, 总高: {total_height:.1f} m")
    print(f"- 执行设计: {'是' if config.PERFORM_CONCRETE_DESIGN else '否'}")
    print(f"- 提取设计内力: {'是' if config.PERFORM_CONCRETE_DESIGN and config.EXPORT_ALL_DESIGN_FILES else '否'}")
    if design:
        topo = design.topology
        print(f"- 参数化方案 case_id={design.case_id}: n_x={topo['n_x']}, n_y={topo['n_y']}, "
              f"l_x={topo['l_x']}mm, l_y={topo['l_y']}mm")
    print("=" * 80)


def stage_setup_and_init():
    """阶段 0：输出目录检查 + API 加载 + 连接 ETABS"""
    print("\n[阶段0] 系统初始化与 ETABS 连接")
    if not file_operations.check_output_directory():
        sys.exit("输出目录检查失败，脚本终止")
    load_dotnet_etabs_api()
    _, sap_model = etabs_setup.setup_etabs()
    return sap_model


def stage_geometry(sap_model, design: Optional[DesignCaseConfig] = None):
    """阶段 1：几何建模（含材料/截面定义）"""
    print("\n[阶段1] 几何建模")
    geometry_modeling.define_all_materials_and_sections()
    if design is None:
        column_names, beam_names, slab_names, _ = geometry_modeling.create_frame_structure()
    else:
        column_names, beam_names, slab_names, _ = geometry_modeling.create_frame_structure_from_design(design)
        geometry_modeling.create_parametric_frame_sections_from_design(design)
        geometry_modeling.assign_sections_by_design(design, design.topology)
    return column_names, beam_names, slab_names


def stage_loads(column_names, beam_names, slab_names):
    """阶段 2：荷载定义与施加"""
    print("\n[阶段2] 荷载定义与施加")
    load_module.setup_response_spectrum()
    load_module.define_all_load_cases()
    load_module.assign_all_loads_to_frame_structure(column_names, beam_names, slab_names)
    file_operations.finalize_and_save_model()


def stage_analysis(sap_model):
    """阶段 3：结构分析"""
    print("\n[阶段3] 结构分析")
    analysis.wait_and_run_analysis(5)
    if not analysis.check_analysis_completion():
        print("[提醒] 分析状态检查异常，但继续尝试提取结果")


def stage_results(sap_model, frame_element_names, output_dir: Path, workflow_state):
    """阶段 4：结果提取 / 后处理"""
    print("\n[阶段4] 结果提取与报表")
    summary_path = results_extraction.extract_modal_and_drift(sap_model, output_dir)
    print(f"动态分析结果概要已写入 Excel: {summary_path}")
    results_extraction.export_beam_and_column_element_forces(output_dir)
    results_extraction.export_frame_member_forces(output_dir)
    workflow_state["analysis_completed"] = True
    return summary_path


def stage_design_and_forces(sap_model, column_names, beam_names, output_dir: Path, workflow_state):
    """阶段 5：构件设计 + 设计内力提取（按配置可选）"""
    analysis_output_dir = Path(config.ANALYSIS_DATA_DIR)
    design_output_dir = Path(config.DESIGN_DATA_DIR)
    if not config.PERFORM_CONCRETE_DESIGN:
        print("\n[阶段5] 跳过构件设计与设计内力提取（配置关闭）")
        return

    print("\n[阶段5] 混凝土构件设计与设计结果提取")
    try:
        design_ok = analysis.perform_concrete_design_and_extract_results()
        workflow_state["design_completed"] = bool(design_ok)
        if design_ok:
            print("[完成] 设计和结果提取验证通过")
        else:
            print("[警告] 设计和结果提取失败，请检查日志")
    except Exception as exc:  # noqa: BLE001
        workflow_state["design_completed"] = False
        print(f"[错误] 构件设计模块发生严重错误: {exc}")
        traceback.print_exc()
        return

    if not workflow_state["design_completed"]:
        print("因设计阶段未成功，跳过设计内力提取")
        return

    core_files = results_extraction.export_core_results(sap_model, design_output_dir)
    if core_files:
        print("\n核心结果文件:")
        for name, path in core_files.items():
            print(f"  - {name}: {path}")
    missing_keys = {name for name, path in core_files.items() if not Path(path).exists()}
    workflow_state["force_extraction_completed"] = not missing_keys
    if missing_keys:
        print(f"[警告] 核心结果缺少: {sorted(missing_keys)}")

    if not config.EXPORT_ALL_DESIGN_FILES:
        print("已生成核心结果文件，跳过全量设计 CSV 导出")
        return

    try:
        if results_extraction.extract_design_forces_and_summary(column_names, beam_names):
            workflow_state["force_extraction_completed"] = True
            print("构件设计内力提取成功（全量导出）")
        else:
            print("构件设计内力提取失败，请检查日志")
    except Exception as exc:  # noqa: BLE001
        print(f"设计内力提取模块发生严重错误: {exc}")
        traceback.print_exc()
    try:
        print("\n[阶段5] 设计完成后刷新梁/柱/框架内力（含设计组合）")
        analysis_output_dir.mkdir(parents=True, exist_ok=True)
        results_extraction.export_beam_and_column_element_forces(analysis_output_dir)
        results_extraction.export_frame_member_forces(analysis_output_dir)
    except Exception as exc:  # noqa: BLE001
        print(f"[WARN] 设计后刷新梁/柱内力失败: {exc}")


def generate_final_report(start_time, workflow_state):
    """生成并打印最终的执行总结报告"""
    elapsed_time = time.time() - start_time
    design_dir = Path(config.DESIGN_DATA_DIR)
    output_dir = Path(config.DATA_EXTRACTION_DIR)
    analysis_dir = Path(config.ANALYSIS_DATA_DIR)
    print("\n" + "=" * 80)
    print("框架结构建模与分析流程完成")
    print(f"总执行时间 {elapsed_time:.2f} 秒")
    print("=" * 80)
    status_map = {True: "成功", False: "失败", None: "跳过"}
    print("执行状态总结:")
    print(f"   - 结构建模与分析: {status_map[workflow_state.get('analysis_completed', False)]}")
    if config.PERFORM_CONCRETE_DESIGN:
        design_status = status_map[workflow_state.get('design_completed', False)]
        print(f"   - 构件设计: {design_status}")
        if workflow_state.get("design_completed"):
            force_status = status_map[workflow_state.get('force_extraction_completed', False)]
        else:
            force_status = "跳过 (设计未成功)"
        print(f"   - 设计内力提取: {force_status}")
    else:
        print(f"   - 构件设计: {status_map[None]}")
        print(f"   - 设计内力提取: {status_map[None]}")
    print("\n主要输出文件位于 data_extraction 子目录:")
    print(f"   - 模型文件: {config.MODEL_PATH}")
    print(f"   - 分析内力: {analysis_dir / 'frame_member_forces.csv'}")
    if workflow_state.get("design_completed"):
        print(f"   - 配筋结果: {design_dir / 'concrete_design_results_enhanced.csv'}")
        print(f"   - 设计报告: {design_dir / 'design_summary_report.txt'}")
    print(f"   - 梁分析内力表: {analysis_dir / 'beam_element_forces.csv'}")
    print(f"   - 柱分析内力表: {analysis_dir / 'column_element_forces.csv'}")

    core_paths = [
        ("analysis_dynamic_summary.xlsx", analysis_dir / "analysis_dynamic_summary.xlsx"),
        ("beam_flexure_envelope.csv", design_dir / "beam_flexure_envelope.csv"),
        ("beam_shear_envelope.csv", design_dir / "beam_shear_envelope.csv"),
        ("column_pmm_design_forces_raw.csv", design_dir / "column_pmm_design_forces_raw.csv"),
        ("column_shear_envelope.csv", design_dir / "column_shear_envelope.csv"),
        ("beam_element_forces.csv", analysis_dir / "beam_element_forces.csv"),
        ("column_element_forces.csv", analysis_dir / "column_element_forces.csv"),
    ]

    if workflow_state.get("force_extraction_completed"):
        print("   - 动态分析概要:", core_paths[0][1])
        print("   - 梁弯矩包络:", core_paths[1][1])
        print("   - 梁剪力包络:", core_paths[2][1])
        print("   - 柱P-M-M 原始:", core_paths[3][1])
        print("   - 柱剪力包络:", core_paths[4][1])
        print("   - 梁分析内力表:", core_paths[5][1])
        print("   - 柱分析内力表:", core_paths[6][1])
        if config.EXPORT_ALL_DESIGN_FILES:
            print("   - 其他设计输出：已启用全量导出，请查看目录。")

    if config.PERFORM_CONCRETE_DESIGN and workflow_state.get("design_completed"):
        missing_core = [name for name, path in core_paths if not path.exists()]
        if missing_core:
            print(f"[警告] 核心结果文件未生成: {missing_core}")
    print("=" * 80)


def run_pipeline():
    """Run the full four-stage pipeline."""
    design_case = prepare_design()
    script_start_time = time.time()
    workflow_state = {
        "analysis_completed": False,
        "design_completed": False,
        "force_extraction_completed": False,
    }

    try:
        print_project_info(design_case)
        sap_model = stage_setup_and_init()
        column_names, beam_names, slab_names = stage_geometry(sap_model, design_case)
        stage_loads(column_names, beam_names, slab_names)
        stage_analysis(sap_model)
        output_dir = Path(config.DATA_EXTRACTION_DIR)
        analysis_output_dir = Path(config.ANALYSIS_DATA_DIR)
        design_output_dir = Path(config.DESIGN_DATA_DIR)
        output_dir.mkdir(parents=True, exist_ok=True)
        analysis_output_dir.mkdir(parents=True, exist_ok=True)
        design_output_dir.mkdir(parents=True, exist_ok=True)
        stage_results(sap_model, column_names + beam_names, analysis_output_dir, workflow_state)
        stage_design_and_forces(sap_model, column_names, beam_names, design_output_dir, workflow_state)
    except SystemExit as exc:
        print("\n--- 脚本已结束 ---")
        if exc.code != 0:
            print(f"退出代码 {exc.code}")
    except Exception as exc:  # noqa: BLE001
        print("\n--- 未预料的运行时错误 ---")
        print(f"错误类型: {type(exc).__name__}: {exc}")
        traceback.print_exc()
        file_operations.cleanup_etabs_on_error()
        sys.exit(1)
    finally:
        generate_final_report(script_start_time, workflow_state)
        if not config.ATTACH_TO_INSTANCE:
            print("脚本执行完毕，ETABS 将保持打开状态。")
        file_operations.remove_pycache()


if __name__ == "__main__":
    run_pipeline()
