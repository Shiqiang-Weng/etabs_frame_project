#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Canonical ETABS frame pipeline runner with five toggleable stages + final report:
1) 几何建模 (geometry_modeling)
2) 荷载定义与施加 (load_module)
3) 结构分析 (analysis)
4) 分析结果提取 / 后处理 (results_extraction)
5) 构件设计 + 设计内力提取（可选）
"""

import json
import re
import sys
import time
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from common import config
from common.config import design_config_from_case
from common.dataset_paths import (
    BUCKET_SIZE,
    DONE_MARKER_FILENAME,
    INPUT_BUCKET_PREFIX,
    NUM_BUCKETS,
    OUTPUT_BUCKET_PREFIX,
    build_bucket_dir,
    compute_bucket,
    iter_bucket_ranges,
)
from common.etabs_api_loader import load_dotnet_etabs_api
from common import etabs_setup, file_operations

import geometry_modeling
import load_module
import analysis
import results_extraction
from results_extraction.other_output_items import export_other_output_items_tables
from parametric_model.param_sampling import (
    DesignCaseConfig,
    extend_param_plan_auto_generated,
    generate_param_plan_multi_files,
    find_param_plan_file,
    load_param_plan,
)
from rerun_missing_cases import rerun_missing_cases
from gnn_dataset import extract_gnn_features

PROJECT_ROOT = Path(__file__).resolve().parent
PLAN_DIR = PROJECT_ROOT / "parametric_model"
PLAN_PREFIX = "param_plan"
PLAN_PATH = PLAN_DIR / f"{PLAN_PREFIX}.jsonl"
PLAN_AUTO_DIR = PLAN_DIR / "param_plan_auto_generated"
PLAN_AUTO_BASE_NAME = "param_plan_auto_generated"
PLAN_SEED = 42
PLAN_AUTO_TARGET_CASES = 30000
PLAN_AUTO_PER_FILE = 1000
PLAN_AUTO_NUM_FILES = (PLAN_AUTO_TARGET_CASES + PLAN_AUTO_PER_FILE - 1) // PLAN_AUTO_PER_FILE
PLAN_AUTO_MAX_ATTEMPTS = 3000000

ORIGINAL_SCRIPT_DIRECTORY = Path(config.SCRIPT_DIRECTORY)
ORIGINAL_MODEL_NAME = config.MODEL_NAME


@dataclass
class PipelineOptions:
    run_stage1_geometry: bool = True
    run_stage2_loads: bool = True
    run_stage3_analysis: bool = True
    run_stage4_results: bool = True
    run_stage5_design: bool = False
    run_final_report: bool = True


def get_case_bucket(case_id: int, bucket_size: int = BUCKET_SIZE, num_buckets: int = NUM_BUCKETS):
    """Return (start, end, suffix) bucket info for a case id."""
    bucket = compute_bucket(case_id, bucket_size=bucket_size, num_buckets=num_buckets)
    return bucket.start, bucket.end, bucket.suffix


def ensure_bucket_directories(root: Path, prefix: str) -> None:
    """Ensure all bucket folders (0-999, 1000-1999, ...) exist under root."""
    root.mkdir(parents=True, exist_ok=True)
    for start, end in iter_bucket_ranges(bucket_size=BUCKET_SIZE, num_buckets=NUM_BUCKETS):
        bucket_dir = root / f"{prefix}{start}-{end}"
        bucket_dir.mkdir(parents=True, exist_ok=True)


def get_output_bucket_dir(case_id: int) -> Path:
    """Bucket directory under SCRIPT_DIRECTORY for a case."""
    bucket = compute_bucket(case_id, bucket_size=BUCKET_SIZE, num_buckets=NUM_BUCKETS)
    return build_bucket_dir(ORIGINAL_SCRIPT_DIRECTORY, OUTPUT_BUCKET_PREFIX, bucket)


def get_input_bucket_dir(case_id: int) -> Path:
    """Bucket directory under PLAN_AUTO_DIR for graph inputs."""
    bucket = compute_bucket(case_id, bucket_size=BUCKET_SIZE, num_buckets=NUM_BUCKETS)
    return build_bucket_dir(PLAN_AUTO_DIR, INPUT_BUCKET_PREFIX, bucket)


def get_case_output_dir(case_id: int) -> Path:
    """Full case directory (bucketed) for outputs."""
    return get_output_bucket_dir(case_id) / f"case_{case_id}"


def get_done_flag_path(case_id: int) -> Path:
    """Path to the DONE marker file for a case."""
    return get_case_output_dir(case_id) / DONE_MARKER_FILENAME


def _resolve_case_id(sample: dict, fallback_id: int) -> int:
    for key in ("case_id", "num", "case", "id", "case_no"):
        raw = sample.get(key)
        if raw is None or (isinstance(raw, str) and not raw.strip()):
            continue
        try:
            return int(raw)
        except (TypeError, ValueError):
            continue
    return fallback_id


def _refresh_settings_paths() -> None:
    """刷新 config.SETTINGS 中的路径信息，保持新旧接口一致。"""
    config.SETTINGS = config.Settings(
        paths=config.PathsConfig(
            use_net_core=config.USE_NET_CORE,
            program_path=config.PROGRAM_PATH,
            dll_path=config.ETABS_DLL_PATH,
            script_directory=config.SCRIPT_DIRECTORY,
            data_extraction_dir=config.DATA_EXTRACTION_DIR,
            analysis_data_dir=config.ANALYSIS_DATA_DIR,
            design_data_dir=config.DESIGN_DATA_DIR,
            model_path=config.MODEL_PATH,
        ),
        materials=config.SETTINGS.materials,
        loads=config.SETTINGS.loads,
        response_spectrum=config.SETTINGS.response_spectrum,
        design=config.DesignOptions(
            perform_concrete_design=config.PERFORM_CONCRETE_DESIGN,
            export_all_design_files=config.EXPORT_ALL_DESIGN_FILES,
            reanalyze_before_design=config.REANALYZE_BEFORE_DESIGN,
            enable_legacy_design_export=config.ENABLE_LEGACY_DESIGN_EXPORT,
            design_debug_logs=config.DESIGN_DEBUG_LOGS,
        ),
    )


def _patch_module_paths_for_case() -> None:
    """同步其他模块中缓存的路径常量，确保输出按案例分目录。"""
    import common.file_operations as file_ops
    file_ops.MODEL_PATH = config.MODEL_PATH
    file_ops.SCRIPT_DIRECTORY = config.SCRIPT_DIRECTORY
    file_ops.DATA_EXTRACTION_DIR = config.DATA_EXTRACTION_DIR
    file_ops.ANALYSIS_DATA_DIR = config.ANALYSIS_DATA_DIR
    file_ops.DESIGN_DATA_DIR = config.DESIGN_DATA_DIR

    import analysis.runner as runner
    runner.MODEL_PATH = config.MODEL_PATH

    import analysis.design_workflow as design_workflow
    design_workflow.DESIGN_DATA_DIR = config.DESIGN_DATA_DIR
    design_workflow.SCRIPT_DIRECTORY = config.SCRIPT_DIRECTORY
    design_workflow.PERFORM_CONCRETE_DESIGN = config.PERFORM_CONCRETE_DESIGN

    import results_extraction.member_forces as member_forces
    member_forces.ANALYSIS_DATA_DIR = config.ANALYSIS_DATA_DIR

    import results_extraction.core_results_module as core_results_module
    core_results_module.SCRIPT_DIRECTORY = config.SCRIPT_DIRECTORY
    core_results_module.DESIGN_DATA_DIR = config.DESIGN_DATA_DIR
    core_results_module.ANALYSIS_DATA_DIR = config.ANALYSIS_DATA_DIR

    import results_extraction.design_forces as design_forces
    design_forces.DESIGN_DATA_DIR = config.DESIGN_DATA_DIR

    import results_extraction.concrete_frame_detail_data as cfdd
    cfdd.DESIGN_DATA_DIR = config.DESIGN_DATA_DIR

    import results_extraction as re_init
    re_init.SCRIPT_DIRECTORY = config.SCRIPT_DIRECTORY
    re_init.DATA_EXTRACTION_DIR = config.DATA_EXTRACTION_DIR
    re_init.ANALYSIS_DATA_DIR = config.ANALYSIS_DATA_DIR
    re_init.DESIGN_DATA_DIR = config.DESIGN_DATA_DIR


def set_case_output_paths(case_id: int) -> Path:
    """
    根据案例编号更新路径到 case_{id} 子目录，所有输出（模型/分析/设计）均落在该目录。
    """
    bucket_dir = get_output_bucket_dir(case_id)
    case_dir = bucket_dir / f"case_{case_id}"
    data_dir = case_dir / "data_extraction"
    analysis_dir = data_dir / "analysis_data"
    design_dir = data_dir / "design_data"

    bucket_dir.mkdir(parents=True, exist_ok=True)
    config.SCRIPT_DIRECTORY = str(case_dir)
    config.DATA_EXTRACTION_DIR = str(data_dir)
    config.ANALYSIS_DATA_DIR = str(analysis_dir)
    config.DESIGN_DATA_DIR = str(design_dir)
    config.MODEL_PATH = str(case_dir / ORIGINAL_MODEL_NAME)

    _refresh_settings_paths()
    _patch_module_paths_for_case()
    return case_dir


def _collect_auto_plan_files(auto_dir: Path, base_name: str) -> List[Path]:
    """收集自动生成的多文件方案列表（按数字后缀排序）。"""
    if not auto_dir.exists():
        return []
    pattern = re.compile(rf"^{re.escape(base_name)}_(\d+)$")
    matches: List[Tuple[int, Path]] = []
    for path in auto_dir.glob(f"{base_name}_*.csv"):
        if not path.is_file():
            continue
        match = pattern.match(path.stem)
        if not match:
            continue
        matches.append((int(match.group(1)), path))
    matches.sort(key=lambda item: item[0])
    return [path for _, path in matches]


def _load_plan_files(plan_files: List[Path]) -> List[dict]:
    """加载多个方案文件并拼接样本列表。"""
    samples: List[dict] = []
    for path in plan_files:
        part = load_param_plan(path)
        print(f"[参数化模式] 从 {path} 读取到 {len(part)} 个样本")
        samples.extend(part)
    return samples


def _count_samples_in_files(plan_files: List[Path]) -> int:
    """逐文件统计样本数量（使用 load_param_plan 保持口径一致）。"""
    total = 0
    for path in plan_files:
        total += len(load_param_plan(path))
    return total


def prepare_design_cases() -> Tuple[Path, List[DesignCaseConfig]]:
    """Locate or generate a param plan and normalize all cases."""
    PLAN_DIR.mkdir(parents=True, exist_ok=True)
    ensure_bucket_directories(PLAN_AUTO_DIR, INPUT_BUCKET_PREFIX)
    manual_plan = find_param_plan_file(PLAN_DIR, PLAN_PREFIX)
    plan_files: List[Path] = []

    if manual_plan:
        print(f"[参数化模式] 检测到方案文件: {manual_plan}")
        plan_files = [manual_plan]
    else:
        auto_files = _collect_auto_plan_files(PLAN_AUTO_DIR, PLAN_AUTO_BASE_NAME)
        if auto_files:
            print(f"[参数化模式] 使用已存在的自动方案目录: {PLAN_AUTO_DIR}")
            existing_count = _count_samples_in_files(auto_files)
            remaining = PLAN_AUTO_TARGET_CASES - existing_count
            print(
                f"[参数化模式] 自动方案 existing_count={existing_count}, "
                f"target={PLAN_AUTO_TARGET_CASES}, remaining={remaining}"
            )
            if existing_count < PLAN_AUTO_TARGET_CASES:
                print(
                    f"[参数化模式] 检测到已有 {existing_count}，将扩展到 {PLAN_AUTO_TARGET_CASES}"
                )
                extend_param_plan_auto_generated(
                    out_dir=PLAN_AUTO_DIR,
                    target_total_cases=PLAN_AUTO_TARGET_CASES,
                    per_file=PLAN_AUTO_PER_FILE,
                    seed=PLAN_SEED,
                    batch_flush_size=1000,
                    sleep_seconds_between_flush=3,
                    max_attempts=PLAN_AUTO_MAX_ATTEMPTS,
                    base_name=PLAN_AUTO_BASE_NAME,
                )
                auto_files = _collect_auto_plan_files(PLAN_AUTO_DIR, PLAN_AUTO_BASE_NAME)
                expanded_count = _count_samples_in_files(auto_files)
                print(
                    f"[参数化模式] 扩展完成: files={len(auto_files)}, total_samples={expanded_count}"
                )
            else:
                print("[参数化模式] 无需扩展，已满足 30000")
            plan_files = auto_files
        else:
            pattern_path = PLAN_AUTO_DIR / f"{PLAN_AUTO_BASE_NAME}_*.csv"
            print(
                f"[参数化模式] 未找到方案文件，自动生成: "
                f"{pattern_path} (cases={PLAN_AUTO_TARGET_CASES})"
            )
            plan_files = generate_param_plan_multi_files(
                total_cases=PLAN_AUTO_TARGET_CASES,
                out_dir=PLAN_AUTO_DIR,
                num_files=PLAN_AUTO_NUM_FILES,
                seed=PLAN_SEED,
                batch_flush_size=1000,
                sleep_seconds_between_flush=3,
                max_attempts=PLAN_AUTO_MAX_ATTEMPTS,
            )

    if not plan_files:
        raise FileNotFoundError("未找到任何参数方案文件，自动生成也失败。")

    samples = _load_plan_files(plan_files)
    if not samples:
        raise RuntimeError(f"方案文件 {plan_files} 中没有任何样本")

    design_cases: List[DesignCaseConfig] = []
    for sample in samples:
        fallback = len(design_cases)
        case_id = _resolve_case_id(sample, fallback)
        design_cases.append(DesignCaseConfig.from_sample(case_id, sample))
    return plan_files[0], design_cases


def prepare_design():
    """保持兼容：仅返回首个方案的 DesignConfig。"""
    _, design_cases = prepare_design_cases()
    design_case = design_cases[0]
    design_cfg = design_config_from_case(design_case)
    print(f"[参数化模式] 将运行 case_index=0, case_id={design_cfg.case_id}")
    return design_cfg


def print_project_info(design_cfg=None) -> None:
    """打印项目和脚本配置信息"""
    from common.config import DEFAULT_DESIGN_CONFIG

    cfg = design_cfg or DEFAULT_DESIGN_CONFIG
    num_stories = cfg.storeys.num_storeys
    total_height = num_stories * cfg.storeys.storey_height
    print("=" * 80)
    print("ETABS 框架结构自动建模脚本 (分阶段可开关流水线)")
    print("=" * 80)
    print("关键参数:")
    print(f"- 楼层数 {num_stories}, 总高: {total_height:.1f} m")
    print(f"- 执行设计: {'是' if config.PERFORM_CONCRETE_DESIGN else '否'}")
    print(f"- 提取设计内力: {'是' if config.PERFORM_CONCRETE_DESIGN and config.EXPORT_ALL_DESIGN_FILES else '否'}")
    if design_cfg:
        topo = cfg.topology
        print(
            f"- 参数化方案 case_id={cfg.case_id}: n_x={topo['n_x']}, n_y={topo['n_y']}, "
            f"l_x={topo['l_x_mm']}mm, l_y={topo['l_y_mm']}mm"
        )
    print("=" * 80)


def stage_setup_and_init(design_cfg, options: PipelineOptions):
    """阶段 0：输出目录检查 + API 加载 + 连接 ETABS"""
    print("\n[阶段0] 系统初始化与 ETABS 连接")
    if not file_operations.check_output_directory(
        create_analysis_dir=options.run_stage4_results, create_design_dir=options.run_stage5_design
    ):
        sys.exit("输出目录检查失败，脚本终止")
    load_dotnet_etabs_api()
    _, sap_model = etabs_setup.setup_etabs(design_cfg)
    return sap_model


def stage_geometry(sap_model, design_cfg, execute: bool, workflow_state: dict):
    """阶段 1：几何建模（含材料/截面定义）"""
    if not execute:
        print("\n[阶段1] 跳过几何建模（execute=False）")
        workflow_state["stage1_geometry"] = None
        return None, None, None
    print("\n[阶段1] 几何建模")
    try:
        geometry_modeling.define_all_materials_and_sections(design_cfg)
        column_names, beam_names, slab_names, _ = geometry_modeling.create_frame_structure(design_cfg)
        geometry_modeling.assign_sections_by_design(design_cfg)
        workflow_state["stage1_geometry"] = True
        return column_names, beam_names, slab_names
    except Exception:
        workflow_state["stage1_geometry"] = False
        raise


def stage_loads(column_names, beam_names, slab_names, execute: bool, workflow_state: dict, analysis_output_needed: bool):
    """阶段 2：荷载定义与施加"""
    if not execute:
        print("\n[阶段2] 跳过荷载定义与施加（execute=False）")
        workflow_state["stage2_loads"] = None
        return
    if not column_names or not beam_names or not slab_names:
        raise RuntimeError(
            "阶段1被跳过或未生成构件时无法执行阶段2。请先运行几何建模，或实现从既有模型中读取构件列表。"
        )
    print("\n[阶段2] 荷载定义与施加")
    try:
        load_module.setup_response_spectrum()
        load_module.define_all_load_cases()
        load_module.assign_all_loads_to_frame_structure(column_names, beam_names, slab_names)
        file_operations.finalize_and_save_model(
            create_analysis_dir=analysis_output_needed,
            create_design_dir=config.PERFORM_CONCRETE_DESIGN,
        )
        workflow_state["stage2_loads"] = True
    except Exception:
        workflow_state["stage2_loads"] = False
        raise


def stage_analysis(sap_model, execute: bool, workflow_state: dict):
    """阶段 3：结构分析"""
    if not execute:
        print("\n[阶段3] 跳过结构分析（execute=False）")
        workflow_state["stage3_analysis"] = None
        return
    print("\n[阶段3] 结构分析")
    try:
        analysis.wait_and_run_analysis(5)
        analysis_ok = analysis.check_analysis_completion()
        if not analysis_ok:
            print("[提醒] 分析状态检查异常，但继续尝试提取结果")
        workflow_state["stage3_analysis"] = analysis_ok
    except Exception:
        workflow_state["stage3_analysis"] = False
        raise


def stage_results(sap_model, frame_element_names, output_dir: Path, workflow_state: dict, execute: bool):
    """阶段 4：结果提取 / 后处理（仅分析结果）"""
    if not execute:
        print("\n[阶段4] 跳过结果提取与报表（execute=False）")
        workflow_state["stage4_results"] = None
        workflow_state["analysis_completed"] = None
        return None
    if not frame_element_names:
        raise RuntimeError(
            "阶段1被跳过时无法执行阶段4（缺少框架构件列表）。请先运行几何建模，或实现从既有模型中读取构件列表。"
        )

    print("\n[阶段4] 结果提取与报表")
    summary_path = None
    success = True

    try:
        summary_path = results_extraction.extract_modal_and_drift(sap_model, output_dir)
        print(f"动态分析结果概要已写入 Excel: {summary_path}")
    except Exception as exc:  # noqa: BLE001
        success = False
        print(f"[WARN] 动态分析结果提取失败: {exc}")
        traceback.print_exc()

    try:
        results_extraction.export_beam_and_column_element_forces(output_dir)
    except Exception as exc:  # noqa: BLE001
        success = False
        print(f"[WARN] 梁柱内力导出失败: {exc}")
        traceback.print_exc()

    try:
        results_extraction.export_frame_member_forces(output_dir)
    except Exception as exc:  # noqa: BLE001
        success = False
        print(f"[WARN] 框架构件内力导出失败: {exc}")
        traceback.print_exc()

    try:
        export_other_output_items_tables(sap_model, output_dir)
    except Exception as exc:  # noqa: BLE001
        success = False
        print(f"[WARN] Other Output Items 导出失败: {exc}")
        traceback.print_exc()

    workflow_state["stage4_results"] = success
    workflow_state["analysis_completed"] = success
    return summary_path


def stage_design_and_forces(
    sap_model,
    column_names,
    beam_names,
    output_dir: Path,
    workflow_state: dict,
    design_cfg,
    execute: bool,
):
    """阶段 5：构件设计 + 设计内力提取（按配置可选）"""
    analysis_output_dir = Path(config.ANALYSIS_DATA_DIR)
    design_output_dir = Path(output_dir)
    if not execute:
        print("\n[阶段5] 跳过构件设计与设计内力提取（execute=False）")
        workflow_state["stage5_design"] = None
        workflow_state["design_completed"] = None
        workflow_state["force_extraction_completed"] = None
        return
    if not column_names or not beam_names:
        raise RuntimeError(
            "阶段1被跳过时无法执行阶段5（缺少梁/柱列表）。请先运行几何建模，或实现从既有模型中读取构件列表。"
        )

    print("\n[阶段5] 混凝土构件设计与设计结果提取")
    design_output_dir.mkdir(parents=True, exist_ok=True)
    try:
        design_ok = analysis.perform_concrete_design_and_extract_results(design_cfg)
        workflow_state["design_completed"] = bool(design_ok)
        if design_ok:
            print("[完成] 设计和结果提取验证通过")
        else:
            print("[警告] 设计和结果提取失败，请检查日志")
    except Exception as exc:  # noqa: BLE001
        workflow_state["design_completed"] = False
        workflow_state["stage5_design"] = False
        print(f"[错误] 构件设计模块发生严重错误: {exc}")
        traceback.print_exc()
        return

    if not workflow_state["design_completed"]:
        workflow_state["stage5_design"] = False
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
        workflow_state["stage5_design"] = workflow_state["design_completed"] and workflow_state.get(
            "force_extraction_completed", True
        )
        print("已生成核心结果文件，跳过全量设计 CSV 导出")
        return

    try:
        if results_extraction.extract_design_forces_and_summary(column_names, beam_names):
            workflow_state["force_extraction_completed"] = True
            print("构件设计内力提取成功（全量导出）")
        else:
            print("构件设计内力提取失败，请检查日志")
    except Exception as exc:  # noqa: BLE001
        workflow_state["force_extraction_completed"] = False
        print(f"设计内力提取模块发生严重错误: {exc}")
        traceback.print_exc()
    try:
        print("\n[阶段5] 设计完成后刷新梁/柱/框架内力（含设计组合）")
        analysis_output_dir.mkdir(parents=True, exist_ok=True)
        results_extraction.export_beam_and_column_element_forces(analysis_output_dir)
        results_extraction.export_frame_member_forces(analysis_output_dir)
    except Exception as exc:  # noqa: BLE001
        print(f"[WARN] 设计后刷新梁/柱内力失败: {exc}")
    workflow_state["stage5_design"] = bool(
        workflow_state.get("design_completed")
        and (workflow_state.get("force_extraction_completed") is not False)
    )


def generate_final_report(start_time, workflow_state, execute: bool, options: PipelineOptions):
    """生成并打印最终的执行总结报告"""
    if not execute:
        print("\n[最终报告] 跳过报告生成（execute=False）")
        return
    elapsed_time = time.time() - start_time
    design_dir = Path(config.DESIGN_DATA_DIR)
    analysis_dir = Path(config.ANALYSIS_DATA_DIR)
    print("\n" + "=" * 80)
    print("框架结构建模与分析流程完成")
    print(f"总执行时间 {elapsed_time:.2f} 秒")
    print("=" * 80)
    status_map = {True: "成功", False: "失败", None: "跳过"}
    print("执行状态总结:")
    stage_statuses = [
        ("阶段1 几何建模", workflow_state.get("stage1_geometry")),
        ("阶段2 荷载施加", workflow_state.get("stage2_loads")),
        ("阶段3 结构分析", workflow_state.get("stage3_analysis")),
        ("阶段4 结果提取", workflow_state.get("stage4_results")),
        ("阶段5 构件设计", workflow_state.get("stage5_design")),
    ]
    for label, state in stage_statuses:
        print(f"   - {label}: {status_map[state]}")

    print("\n主要输出文件位于 data_extraction 子目录:")
    print(f"   - 模型文件: {config.MODEL_PATH}")
    if options.run_stage4_results:
        print(f"   - 分析内力汇总: {analysis_dir / 'frame_member_forces.csv'}")
        print(f"   - 梁分析内力表: {analysis_dir / 'beam_element_forces.csv'}")
        print(f"   - 柱分析内力表: {analysis_dir / 'column_element_forces.csv'}")
        print(f"   - 动态分析概要: {analysis_dir / 'analysis_dynamic_summary.xlsx'}")
    else:
        print("   - 分析结果导出: 已跳过")

    if options.run_stage5_design:
        if workflow_state.get("design_completed"):
            print(f"   - 配筋结果: {design_dir / 'concrete_design_results_enhanced.csv'}")
            print(f"   - 设计报告: {design_dir / 'design_summary_report.txt'}")
            if workflow_state.get("force_extraction_completed") and config.EXPORT_ALL_DESIGN_FILES:
                print(f"   - 梁弯矩包络: {design_dir / 'beam_flexure_envelope.csv'}")
                print(f"   - 梁剪力包络: {design_dir / 'beam_shear_envelope.csv'}")
                print(f"   - 柱P-M-M 原始: {design_dir / 'column_pmm_design_forces_raw.csv'}")
                print(f"   - 柱剪力包络: {design_dir / 'column_shear_envelope.csv'}")
        else:
            print("   - 设计输出: 设计阶段未成功，未生成设计文件")
    else:
        print("   - 设计输出: 已跳过（设计阶段关闭）")

    if options.run_stage5_design and workflow_state.get("design_completed"):
        core_paths = [
            design_dir / "concrete_design_results_enhanced.csv",
            design_dir / "design_summary_report.txt",
        ]
        missing_core = [str(path.name) for path in core_paths if not path.exists()]
        if missing_core:
            print(f"[警告] 核心设计结果文件未生成: {missing_core}")
    print("=" * 80)


def is_case_completed(case_id: int) -> bool:
    return get_done_flag_path(case_id).exists()


def mark_case_completed(case_id: int, workflow_state: dict, extra_meta: Optional[dict] = None) -> Path:
    """Write a DONE marker after key exports complete."""
    marker_path = get_done_flag_path(case_id)
    marker_path.parent.mkdir(parents=True, exist_ok=True)
    bucket = compute_bucket(case_id, bucket_size=BUCKET_SIZE, num_buckets=NUM_BUCKETS)
    payload = {
        "case_id": case_id,
        "bucket": bucket.suffix,
        "analysis_completed": workflow_state.get("analysis_completed"),
        "design_completed": workflow_state.get("design_completed"),
        "force_extraction_completed": workflow_state.get("force_extraction_completed"),
        "timestamp": time.time(),
    }
    if extra_meta:
        payload.update(extra_meta)
    marker_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[断点续跑] DONE 标记写入 {marker_path}")
    return marker_path


def summarize_existing_progress(design_cases: List[DesignCaseConfig]) -> set[int]:
    completed_ids = {case.case_id for case in design_cases if is_case_completed(case.case_id)}
    next_case = next((case.case_id for case in design_cases if case.case_id not in completed_ids), None)
    print(f"[断点续跑] 已完成 {len(completed_ids)}/{len(design_cases)} 个案例；下一个待跑: {next_case}")
    return completed_ids


def run_single_case(design_case: DesignCaseConfig, case_index: int, total_cases: int, options: PipelineOptions) -> None:
    """Execute the full pipeline for one design case."""
    config.PERFORM_CONCRETE_DESIGN = options.run_stage5_design
    design_cfg = design_config_from_case(design_case)
    print("\n" + "=" * 80)
    print(f"[参数化模式] 开始案例 {case_index}/{total_cases} (case_id={design_cfg.case_id})")
    print("=" * 80)
    done_flag = get_done_flag_path(design_cfg.case_id)
    if done_flag.exists():
        print(f"[断点续跑] case_id={design_cfg.case_id} 已完成，标记文件: {done_flag}，跳过执行。")
        return

    # 统一将所有输出指向 case_{id} 子目录，避免多案例结果互相覆盖
    case_dir = set_case_output_paths(design_cfg.case_id)
    print(f"[参数化模式] 当前案例输出目录: {case_dir}")
    script_start_time = time.time()
    workflow_state = {
        "stage1_geometry": None,
        "stage2_loads": None,
        "stage3_analysis": None,
        "stage4_results": None,
        "stage5_design": None,
        "analysis_completed": None,
        "design_completed": None,
        "force_extraction_completed": None,
    }
    gnn_input_path: Optional[Path] = None
    analysis_output_needed = options.run_stage4_results or options.run_stage5_design

    try:
        print_project_info(design_cfg)
        sap_model = stage_setup_and_init(design_cfg, options)
        column_names, beam_names, slab_names = stage_geometry(
            sap_model, design_cfg, execute=options.run_stage1_geometry, workflow_state=workflow_state
        )
        # ---- GNN 输入特征导出（几何建模完成后立即执行） ----
        if options.run_stage1_geometry and column_names and beam_names:
            try:
                gnn_input_path = extract_gnn_features(
                    sap_model=sap_model,
                    design_cfg=design_cfg,
                    frame_element_names=column_names + beam_names,
                    input_root=PLAN_AUTO_DIR,
                    bucket_size=BUCKET_SIZE,
                    num_buckets=NUM_BUCKETS,
                )
            except Exception as exc:  # noqa: BLE001
                print(f"[GNN][WARN] 图输入导出失败（将继续后续流程）: {exc}")
                traceback.print_exc()
        elif not options.run_stage1_geometry:
            print("[GNN] 跳过图输入导出（几何建模未执行）")
        stage_loads(
            column_names,
            beam_names,
            slab_names,
            execute=options.run_stage2_loads,
            workflow_state=workflow_state,
            analysis_output_needed=analysis_output_needed,
        )
        stage_analysis(sap_model, execute=options.run_stage3_analysis, workflow_state=workflow_state)
        output_dir = Path(config.DATA_EXTRACTION_DIR)
        analysis_output_dir = Path(config.ANALYSIS_DATA_DIR)
        design_output_dir = Path(config.DESIGN_DATA_DIR)

        if options.run_stage4_results or options.run_stage5_design:
            output_dir.mkdir(parents=True, exist_ok=True)
        if options.run_stage4_results:
            analysis_output_dir.mkdir(parents=True, exist_ok=True)
        if options.run_stage5_design:
            design_output_dir.mkdir(parents=True, exist_ok=True)

        stage_results(
            sap_model,
            (column_names or []) + (beam_names or []),
            analysis_output_dir,
            workflow_state,
            execute=options.run_stage4_results,
        )
        stage_design_and_forces(
            sap_model,
            column_names,
            beam_names,
            design_output_dir,
            workflow_state,
            design_cfg,
            execute=options.run_stage5_design,
        )

        analysis_ok = options.run_stage4_results is False or workflow_state.get("analysis_completed") is True
        design_ok = options.run_stage5_design is False or workflow_state.get("design_completed") is True
        if analysis_ok and design_ok:
            mark_case_completed(
                design_cfg.case_id,
                workflow_state,
                extra_meta={
                    "output_dir": str(case_dir),
                    "gnn_input": str(gnn_input_path) if gnn_input_path else None,
                },
            )
        else:
            print("[断点续跑] 未写 DONE 标记，因为分析/设计阶段未完成。")
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
        generate_final_report(script_start_time, workflow_state, execute=options.run_final_report, options=options)
        if not config.ATTACH_TO_INSTANCE:
            # 每个案例结束后等待 10 秒再关闭 ETABS，保证模型文件写入完成
            print(f"[参数化模式] 案例 {design_cfg.case_id} 执行完毕，等待 10 秒后关闭 ETABS ...")
            time.sleep(10)
            try:
                my_etabs, _ = etabs_setup.get_etabs_objects()
                if my_etabs is not None:
                    my_etabs.ApplicationExit(False)
            except Exception as exc:  # noqa: BLE001
                print(f"[WARN] 关闭 ETABS 失败: {exc}")
            finally:
                try:
                    etabs_setup.set_sap_model(None)
                    etabs_setup.my_etabs = None  # type: ignore[attr-defined]
                except Exception:
                    pass
        file_operations.remove_pycache()


def run_pipeline(options: Optional[PipelineOptions] = None):
    """Run the full pipeline for all parametric cases."""
    options = options or PipelineOptions()
    plan_path, design_cases = prepare_design_cases()
    if not design_cases:
        print("[参数化模式] 未获取到任何案例，终止执行。")
        return

    ensure_bucket_directories(ORIGINAL_SCRIPT_DIRECTORY, OUTPUT_BUCKET_PREFIX)
    ensure_bucket_directories(PLAN_AUTO_DIR, INPUT_BUCKET_PREFIX)

    # 断点续跑（目录模式）：只要 case_{id} 输出目录存在，则认为该 case 已经“开始/有产出”，默认从最大已存在 case_id + 1 开始
    existing_dir_case_ids = [
        case.case_id for case in design_cases if get_case_output_dir(case.case_id).exists()
    ]
    resume_from_case_id = (max(existing_dir_case_ids) + 1) if existing_dir_case_ids else None
    if resume_from_case_id is not None:
        print(
            f"[断点续跑][目录模式] 检测到最大已存在输出目录 case_id={resume_from_case_id - 1}，"
            f"将从 case_id={resume_from_case_id} 开始运行。"
        )
    else:
        print("[断点续跑][目录模式] 未检测到任何已存在的 case 输出目录，将从第一个 case 开始运行。")

    total_cases = len(design_cases)
    completed_cases = summarize_existing_progress(design_cases)

    candidate_cases = [
        case for case in design_cases if resume_from_case_id is None or case.case_id >= resume_from_case_id
    ]
    pending_cases = sum(1 for case in candidate_cases if case.case_id not in completed_cases)

    def _maybe_rerun_missing_cases() -> None:
        target_case_id = PLAN_AUTO_TARGET_CASES - 1
        if not any(case.case_id == target_case_id for case in design_cases):
            print(f"[缺失重跑] 未找到 case_id={target_case_id}，跳过缺失重跑。")
            return
        if is_case_completed(target_case_id) or get_case_output_dir(target_case_id).exists():
            rerun_missing_cases(
                design_cases=design_cases,
                run_case_fn=run_single_case,
                options=options,
                project_root=PROJECT_ROOT,
                output_root=ORIGINAL_SCRIPT_DIRECTORY,
                get_done_flag_path=get_done_flag_path,
                get_case_output_dir=get_case_output_dir,
            )
        else:
            print(f"[缺失重跑] case_{target_case_id} 尚未完成，跳过缺失重跑。")

    print(f"[参数化模式] 计划运行 {pending_cases}/{total_cases} 个案例 (来源: {plan_path.name})")
    if pending_cases <= 0:
        print("[参数化模式] 所有案例均已完成，退出。")
        _maybe_rerun_missing_cases()
        return

    executed_index = 0
    for idx, design_case in enumerate(design_cases, start=1):
        if resume_from_case_id is not None and design_case.case_id < resume_from_case_id:
            continue
        if design_case.case_id in completed_cases:
            print(f"[参数化模式] 跳过已完成案例 case_id={design_case.case_id}")
            continue
        executed_index += 1
        run_single_case(design_case, executed_index, total_cases, options)

    _maybe_rerun_missing_cases()


if __name__ == "__main__":
    run_pipeline()
