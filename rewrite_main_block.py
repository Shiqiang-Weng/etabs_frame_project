from pathlib import Path

path = Path('main.py')
lines = path.read_text(encoding='utf-8').splitlines()

# Update header comment above optional module loader
comment_idx = next(i for i, line in enumerate(lines) if line.startswith('# ---'))
lines[comment_idx] = '# --- 可选模块动态导入 ---'

optional_block = [
    'def _import_optional_module(module_names, function_name):',
    '    """尝试从多个可能的模块名中导入一个函数。返回 (函数对象, 是否成功)。"""',
    '    for module_name in module_names:',
    '        try:',
    '            module = __import__(module_name, fromlist=[function_name])',
    '            func = getattr(module, function_name)',
    '            print(f"[可选模块] {module_name!r} 导入成功")',
    '            return func, True',
    '        except ImportError:',
    '            print(f"[可选模块] 未找到模块 {module_name!r}，尝试下一个...")',
    '            continue',
    '    print("[可选模块] 所有候选模块({})均导入失败".format(", ".join(module_names)))',
    '    return None, False',
    '',
]

project_info_block = [
    'def print_project_info():',
    '    """打印项目和脚本配置信息"""',
    '    print("=" * 80)',
    '    print("ETABS 框架结构自动建模脚本 v7.0 (优化版)")',
    '    print("=" * 80)',
    '    print("模块状态:")',
    '    print(f"- 设计模块: {\'可用\' if design_module_available else \'不可用\'}")',
    '    print(f"- 设计内力提取模块: {\'可用\' if design_force_extraction_available else \'不可用\'}")',
    '    print("\\n关键参数:")',
    '    print(f"- 楼层数 {NUM_STORIES}, 总高: {BOTTOM_STORY_HEIGHT + (NUM_STORIES - 1) * TYPICAL_STORY_HEIGHT:.1f}m")',
    '    print(f"- 执行设计: {\'是\' if PERFORM_CONCRETE_DESIGN else \'否\'}")',
    '    print(f"- 提取设计内力: {\'是\' if PERFORM_CONCRETE_DESIGN and design_force_extraction_available else \'否\'}")',
    '    print("=" * 80)',
    '',
]

analysis_block = [
    'def run_analysis_and_results_extraction(sap_model, frame_element_names):',
    '    """阶段五和六：结构分析与结果提取"""',
    '    print("\\n[阶段五/六] 结构分析(analysis) 与结果提取(results_extraction)")',
    '    wait_and_run_analysis(5)',
    '    if not check_analysis_completion():',
    '        print("[提醒] 分析状态检查异常，但继续尝试提取结果")',
    '    dynamic_summary_path = extract_modal_and_drift(sap_model, SCRIPT_DIRECTORY)',
    '    print(f"动态分析结果概要已写入 Excel: {dynamic_summary_path}")',
    '    extract_and_save_frame_forces(frame_element_names)',
    '',
]

design_block = [
    'def run_design_and_force_extraction(workflow_state, sap_model, column_names, beam_names):',
    '    """阶段八和九：构件设计与设计内力提取"""',
    '    if not PERFORM_CONCRETE_DESIGN:',
    '        print("\\n[跳过] 阶段八&九：根据配置跳过构件设计和内力提取")',
    '        return',
    '',
    '    # --- 阶段八：构件设计 ---',
    '    print("\\n[阶段八] 混凝土构件配筋设计")',
    '    if not design_module_available:',
    '        print("[错误] 设计模块不可用，无法执行设计")',
    '        return',
    '',
    '    try:',
    '        if perform_concrete_design_and_extract_results():',
    '            print("[完成] 设计和结果提取验证通过")',
    '            workflow_state[\'design_completed\'] = True',
    '        else:',
    '            print("[警告] 设计和结果提取失败，请检查 design_module 日志")',
    '    except Exception as e:',
    '        print(f"[错误] 构件设计模块发生严重错误: {e}")',
    '        traceback.print_exc()',
    '',
    '    # --- 阶段九：设计内力提取 ---',
    '    print("\\n[阶段九] 构件设计内力提取")',
    '    if not workflow_state[\'design_completed\']:',
    '        print("因设计阶段未成功，跳过设计内力提取")',
    '        return',
    '',
    '    core_files = export_core_results(sap_model, SCRIPT_DIRECTORY)',
    '    expected_core_keys = {',
    '        "analysis_dynamic_summary",',
    '        "beam_flexure_envelope",',
    '        "beam_shear_envelope",',
    '        "column_pmm_design_forces_raw",',
    '        "column_shear_envelope",',
    '    }',
    '    if core_files:',
    '        print("\\n核心结果文件:")',
    '        for name, path in core_files.items():',
    '            print(f"  - {name}: {path}")',
    '    missing_keys = {name for name, path in core_files.items() if not Path(path).exists()}',
    '    workflow_state[\'force_extraction_completed\'] = not missing_keys',
    '    if missing_keys:',
    '        print(f"[警告] 核心结果缺少: {sorted(missing_keys)}")',
    '',
    '    if not EXPORT_ALL_DESIGN_FILES:',
    '        print("已生成核心结果文件，跳过全量设计 CSV 导出")',
    '        return',
    '    if not design_force_extraction_available:',
    '        print("设计内力提取模块不可用，跳过")',
    '        return',
    '',
    '    try:',
    '        if extract_design_forces_and_summary(column_names, beam_names):',
    '            print("构件设计内力提取成功（全量导出）")',
    '            workflow_state[\'force_extraction_completed\'] = True',
    '        else:',
    '            print("构件设计内力提取失败，请检查日志")',
    '    except Exception as e:',
    '        print(f"设计内力提取模块发生严重错误: {e}")',
    '        traceback.print_exc()',
    '',
]

final_report_block = [
    'def generate_final_report(start_time, workflow_state):',
    '    """生成并打印最终的执行总结报告"""',
    '    elapsed_time = time.time() - start_time',
    '    print("\\n" + "=" * 80)',
    '    print("框架结构建模与分析流程完成")',
    '    print(f"总执行时间 {elapsed_time:.2f} 秒")',
    '    print("=" * 80)',
    '    print("执行状态总结:")',
    '    status_map = {True: "成功", False: "失败", None: "跳过"}',
    '    print(f"   - 结构建模与分析: {status_map[True]}")',
    '    if PERFORM_CONCRETE_DESIGN:',
    '        design_status = status_map[workflow_state["design_completed"]] if design_module_available else "跳过 (模块不可用)"',
    '        print(f"   - 构件设计: {design_status}")',
    '        if workflow_state["design_completed"]:',
    '            if design_force_extraction_available:',
    '                force_status = status_map[workflow_state["force_extraction_completed"]]',
    '            else:',
    '                force_status = "跳过 (模块不可用)"',
    '        else:',
    '            force_status = "跳过 (设计未成功)"',
    '        print(f"   - 设计内力提取: {force_status}")',
    '    else:',
    '        print(f"   - 构件设计: {status_map[None]}")',
    '        print(f"   - 设计内力提取: {status_map[None]}")',
    '    print("\\n主要输出文件位于脚本目录:")',
    '    print(f"   - 模型文件: {MODEL_PATH}")',
    '    print("   - 分析内力: frame_member_forces.csv")',
    '    if workflow_state.get("design_completed"):',
    '        print("   - 配筋结果: concrete_design_results.csv")',
    '        print("   - 设计报告: design_summary_report.txt")',
    '    if workflow_state.get("force_extraction_completed"):',
    '        print("   - 动态分析概要: analysis_dynamic_summary.xlsx")',
    '        print("   - 梁弯矩包络: beam_flexure_envelope.csv")',
    '        print("   - 梁剪力包络: beam_shear_envelope.csv")',
    '        print("   - 柱P-M-M 原始: column_pmm_design_forces_raw.csv")',
    '        print("   - 柱剪力包络: column_shear_envelope.csv")',
    '        if EXPORT_ALL_DESIGN_FILES:',
    '            print("   - 其他设计输出：已启用全量导出，请查看目录。")',
    '    print("=" * 80)',
    '',
]

main_block = [
    'def main():',
    '    """主函数：协调所有建模、分析和设计流程"""',
    '    script_start_time = time.time()',
    '    workflow_state = {',
    '        "design_completed": False,',
    '        "force_extraction_completed": False,',
    '    }',
    '',
    '    try:',
    '        print_project_info()',
    '        sap_model = run_setup_and_initialization()',
    '        run_model_definition(sap_model)',
    '        column_names, beam_names = run_geometry_and_loading(sap_model)',
    '        run_analysis_and_results_extraction(sap_model, column_names + beam_names)',
    '        run_design_and_force_extraction(workflow_state, sap_model, column_names, beam_names)',
    '    except SystemExit as e:',
    '        print("\\n--- 脚本已结束 ---")',
    '        if e.code != 0:',
    '            print(f"退出代码 {e.code}")',
    '    except Exception as e:',
    '        print("\\n--- 未预料的运行时错误 ---")',
    '        print(f"错误类型: {type(e).__name__}: {e}")',
    '        traceback.print_exc()',
    '        cleanup_etabs_on_error()',
    '        sys.exit(1)',
    '    finally:',
    '        generate_final_report(script_start_time, workflow_state)',
    '        if not ATTACH_TO_INSTANCE:',
    '            print("脚本执行完毕，ETABS 将保持打开状态。")',
    '',
]
setup_block = [
    'def run_setup_and_initialization():',
    '    """阶段一：系统初始化和 ETABS 连接"""',
    '    print("\\n[阶段一] 系统初始化")',
    '    if not check_output_directory():',
    '        sys.exit("输出目录检查失败，脚本终止")',
    '    load_dotnet_etabs_api()',
    '    _, sap_model = setup_etabs()',
    '    return sap_model',
    '',
]


def replace_block(start_def: str, end_def: str, new_block: list[str]):
    start = next(i for i, line in enumerate(lines) if line.startswith(f'def {start_def}'))
    end = next(i for i, line in enumerate(lines) if line.startswith(f'def {end_def}'))
    lines[start:end] = new_block

# Replace optional import helper
start_opt = next(i for i, line in enumerate(lines) if line.startswith('def _import_optional_module'))
end_opt = next(i for i, line in enumerate(lines) if line.startswith('def print_project_info'))
lines[start_opt:end_opt] = optional_block

# Replace project info block
replace_block('print_project_info', 'run_setup_and_initialization', project_info_block + [''])

# Replace setup/init block
replace_block('run_setup_and_initialization', 'run_model_definition', setup_block + [''])

# Replace analysis and design blocks
replace_block('run_analysis_and_results_extraction', 'run_design_and_force_extraction', analysis_block + [''])
replace_block('run_design_and_force_extraction', 'generate_final_report', design_block + [''])

# Replace final report block
start_final = next(i for i, line in enumerate(lines) if line.startswith('def generate_final_report'))
end_final = next(i for i, line in enumerate(lines) if line.startswith('def main'))
lines[start_final:end_final] = final_report_block

# Replace main block
start_main = next(i for i, line in enumerate(lines) if line.startswith('def main'))
end_main = next(i for i, line in enumerate(lines) if line.startswith('if __name__'))
lines[start_main:end_main] = main_block

path.write_text('\n'.join(lines) + '\n', encoding='utf-8')
