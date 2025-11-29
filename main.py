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

from common import config
from common.etabs_api_loader import load_dotnet_etabs_api
from common import etabs_setup, file_operations

import geometry_modeling
import load_module
import analysis
import results_extraction


def print_project_info() -> None:
    """打印项目和脚本配置信息"""
    total_height = config.BOTTOM_STORY_HEIGHT + (config.NUM_STORIES - 1) * config.TYPICAL_STORY_HEIGHT
    print("=" * 80)
    print("ETABS 框架结构自动建模脚本 (规范四阶段流水线)")
    print("=" * 80)
    print("关键参数:")
    print(f"- 楼层数 {config.NUM_STORIES}, 总高: {total_height:.1f} m")
    print(f"- 执行设计: {'是' if config.PERFORM_CONCRETE_DESIGN else '否'}")
    print(f"- 提取设计内力: {'是' if config.PERFORM_CONCRETE_DESIGN and config.EXPORT_ALL_DESIGN_FILES else '否'}")
    print("=" * 80)


def stage_setup_and_init():
    """阶段 0：输出目录检查 + API 加载 + 连接 ETABS"""
    print("\n[阶段0] 系统初始化与 ETABS 连接")
    if not file_operations.check_output_directory():
        sys.exit("输出目录检查失败，脚本终止")
    load_dotnet_etabs_api()
    _, sap_model = etabs_setup.setup_etabs()
    return sap_model


def stage_geometry(sap_model):
    """阶段 1：几何建模（含材料/截面定义）"""
    print("\n[阶段1] 几何建模")
    geometry_modeling.define_all_materials_and_sections()
    column_names, beam_names, slab_names, _ = geometry_modeling.create_frame_structure()
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
    results_extraction.extract_and_save_frame_forces(frame_element_names)
    workflow_state["analysis_completed"] = True
    return summary_path


def stage_design_and_forces(sap_model, column_names, beam_names, output_dir: Path, workflow_state):
    """阶段 5：构件设计 + 设计内力提取（按配置可选）"""
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

    core_files = results_extraction.export_core_results(sap_model, output_dir)
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


def generate_final_report(start_time, workflow_state):
    """生成并打印最终的执行总结报告"""
    elapsed_time = time.time() - start_time
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
    print("\n主要输出文件位于脚本目录:")
    print(f"   - 模型文件: {config.MODEL_PATH}")
    print("   - 分析内力: frame_member_forces.csv")
    if workflow_state.get("design_completed"):
        print("   - 配筋结果: concrete_design_results.csv")
        print("   - 设计报告: design_summary_report.txt")
    if workflow_state.get("force_extraction_completed"):
        print("   - 动态分析概要: analysis_dynamic_summary.xlsx")
        print("   - 梁弯矩包络: beam_flexure_envelope.csv")
        print("   - 梁剪力包络: beam_shear_envelope.csv")
        print("   - 柱P-M-M 原始: column_pmm_design_forces_raw.csv")
        print("   - 柱剪力包络: column_shear_envelope.csv")
        if config.EXPORT_ALL_DESIGN_FILES:
            print("   - 其他设计输出：已启用全量导出，请查看目录。")
    print("=" * 80)


def run_pipeline():
    """Run the full four-stage pipeline."""
    script_start_time = time.time()
    workflow_state = {
        "analysis_completed": False,
        "design_completed": False,
        "force_extraction_completed": False,
    }

    try:
        print_project_info()
        sap_model = stage_setup_and_init()
        column_names, beam_names, slab_names = stage_geometry(sap_model)
        stage_loads(column_names, beam_names, slab_names)
        stage_analysis(sap_model)
        stage_results(sap_model, column_names + beam_names, Path(config.SCRIPT_DIRECTORY), workflow_state)
        stage_design_and_forces(sap_model, column_names, beam_names, Path(config.SCRIPT_DIRECTORY), workflow_state)
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


if __name__ == "__main__":
    run_pipeline()
