# main.py
import sys
import time
import traceback
import os

# 导入配置和核心模块
from config import *
from etabs_api_loader import load_dotnet_etabs_api
from etabs_setup import setup_etabs
from materials_sections import define_all_materials_and_sections
from response_spectrum import define_response_spectrum_functions_in_etabs
from load_cases import define_all_load_cases
from frame_geometry import create_frame_structure
from load_assignment import assign_all_loads_to_frame_structure
from analysis_module import wait_and_run_analysis, check_analysis_completion
from results_extraction import extract_all_analysis_results
from file_operations import finalize_and_save_model, cleanup_etabs_on_error, check_output_directory
from member_force_extraction import extract_and_save_frame_forces


# --- 模块动态导入 ---
def _import_optional_module(module_names, function_name):
    """
    尝试从多个可能的模块名中导入一个函数。
    返回 (函数对象, 是否成功)
    """
    for module_name in module_names:
        try:
            module = __import__(module_name, fromlist=[function_name])
            func = getattr(module, function_name)
            print(f"✅ 模块 '{module_name}' 导入成功。")
            return func, True
        except ImportError:
            print(f"ℹ️ 未找到模块 '{module_name}'，尝试下一个...")
            continue
    print(f"⚠️ 所有可选模块 ({', '.join(module_names)}) 均导入失败。")
    return None, False


# 动态导入设计模块
perform_concrete_design_and_extract_results, design_module_available = _import_optional_module(
    ['design_module'], 'perform_concrete_design_and_extract_results'
)

# 动态导入设计内力提取模块 (支持多个备选名称)
extract_design_forces_and_summary, design_force_extraction_available = _import_optional_module(
    ['design_force_extraction', 'design_force_extraction_fixed', 'design_force_extraction_improved'],
    'extract_design_forces_and_summary'
)


def print_project_info():
    """打印项目和脚本配置信息"""
    print("=" * 80)
    print("ETABS 框架结构自动建模脚本 v7.0 (优化版)")
    print("=" * 80)
    print("模块状态:")
    print(f"- 设计模块: {'✅ 可用' if design_module_available else '❌ 不可用'}")
    print(f"- 设计内力提取模块: {'✅ 可用' if design_force_extraction_available else '❌ 不可用'}")
    print("\n关键参数:")
    print(f"- 楼层数: {NUM_STORIES}, 总高: {BOTTOM_STORY_HEIGHT + (NUM_STORIES - 1) * TYPICAL_STORY_HEIGHT:.1f}m")
    print(f"- 执行设计: {'是' if PERFORM_CONCRETE_DESIGN else '否'}")
    print(f"- 提取设计内力: {'是' if PERFORM_CONCRETE_DESIGN and design_force_extraction_available else '否'}")
    print("=" * 80)


def run_setup_and_initialization():
    """阶段一：系统初始化和ETABS连接"""
    print("\n🚀 阶段一：系统初始化")
    if not check_output_directory():
        sys.exit("❌ 输出目录检查失败，脚本中止。")
    load_dotnet_etabs_api()
    _, sap_model = setup_etabs()
    return sap_model


def run_model_definition(sap_model):
    """阶段二：定义材料、截面和工况"""
    print("\n🏗️ 阶段二：模型定义")
    define_all_materials_and_sections()
    define_response_spectrum_functions_in_etabs()
    define_all_load_cases()


def run_geometry_and_loading(sap_model):
    """阶段三和四：几何建模与荷载分配"""
    print("\n🏢 阶段三 & 四：几何建模与荷载分配")
    column_names, beam_names, slab_names, _ = create_frame_structure()
    assign_all_loads_to_frame_structure(column_names, beam_names, slab_names)
    finalize_and_save_model()
    return column_names, beam_names


def run_analysis_and_results_extraction(sap_model, frame_element_names):
    """阶段六和七：结构分析与结果提取"""
    print("\n🔍 阶段六 & 七：结构分析与结果提取")
    wait_and_run_analysis(5)
    if not check_analysis_completion():
        print("⚠️ 分析状态检查异常，但继续尝试提取结果。")
    extract_all_analysis_results()
    extract_and_save_frame_forces(frame_element_names)


def run_design_and_force_extraction(workflow_state, column_names, beam_names):
    """阶段八和九：构件设计与设计内力提取"""
    if not PERFORM_CONCRETE_DESIGN:
        print("\n⏭️ 阶段八 & 九：根据配置跳过构件设计和内力提取。")
        return

    # --- 阶段八：构件设计 ---
    print("\n🏗️ 阶段八：混凝土构件配筋设计")
    if not design_module_available:
        print("❌ 设计模块不可用，无法执行设计。")
        return

    try:
        if perform_concrete_design_and_extract_results():
            print("✅ 设计和结果提取验证通过。")
            workflow_state['design_completed'] = True
        else:
            print("⚠️ 设计和结果提取失败，请检查 design_module 日志。")
    except Exception as e:
        print(f"❌ 构件设计模块发生严重错误: {e}")
        traceback.print_exc()

    # --- 阶段九：设计内力提取 ---
    print("\n🔬 阶段九：构件设计内力提取")
    if not workflow_state['design_completed']:
        print("⏭️ 因设计阶段未成功，跳过设计内力提取。")
        return
    if not design_force_extraction_available:
        print("⏭️ 设计内力提取模块不可用，跳过。")
        return

    try:
        if extract_design_forces_and_summary(column_names, beam_names):
            print("✅ 构件设计内力提取成功。")
            workflow_state['force_extraction_completed'] = True
        else:
            print("⚠️ 构件设计内力提取失败，请检查日志。")
    except Exception as e:
        print(f"❌ 设计内力提取模块发生严重错误: {e}")
        traceback.print_exc()


def generate_final_report(start_time, workflow_state):
    """生成并打印最终的执行总结报告"""
    elapsed_time = time.time() - start_time
    print("\n" + "=" * 80)
    print("🎉 框架结构建模与分析全部流程完成！")
    print(f"⏱️ 总执行时间: {elapsed_time:.2f} 秒")
    print("=" * 80)

    print("📋 执行状态总结:")
    status_map = {True: '✅ 成功', False: '❌ 失败', None: '⏭️ 跳过'}

    print(f"   - 结构建模与分析: {status_map[True]}")

    if PERFORM_CONCRETE_DESIGN:
        design_status = status_map[workflow_state['design_completed']] if design_module_available else "⏭️ 模块不可用"
        print(f"   - 构件设计: {design_status}")

        force_status = "⏭️ 跳过 (设计未成功)"
        if workflow_state['design_completed']:
            if design_force_extraction_available:
                force_status = status_map[workflow_state['force_extraction_completed']]
            else:
                force_status = "⏭️ 模块不可用"
        print(f"   - 设计内力提取: {force_status}")
    else:
        print(f"   - 构件设计: {status_map[None]}")
        print(f"   - 设计内力提取: {status_map[None]}")

    print("\n📁 主要输出文件位于脚本目录:")
    print(f"   - 模型文件: {MODEL_PATH}")
    print(f"   - 分析内力: frame_member_forces.csv")
    if workflow_state.get('design_completed'):
        print(f"   - 配筋结果: concrete_design_results.csv")
        print(f"   - 设计报告: design_summary_report.txt")
    if workflow_state.get('force_extraction_completed'):
        print(f"   - 柱设计内力: column_design_forces.csv")
        print(f"   - 梁设计结果: beam_flexure_envelope.csv 或 beam_design_forces.csv")
        print(f"   - 柱剪力包络: column_shear_envelope.csv (若存在)")
        print(f"   - 节点包络: joint_envelope.csv (若存在)")
        print(f"   - 内力汇总: design_forces_summary_report.txt")

    print("=" * 80)


def main():
    """主函数 - 协调所有建模、分析和设计流程"""
    script_start_time = time.time()

    # 初始化工作流状态
    workflow_state = {
        'design_completed': False,
        'force_extraction_completed': False
    }

    try:
        print_project_info()

        # 执行核心流程
        sap_model = run_setup_and_initialization()
        run_model_definition(sap_model)
        column_names, beam_names = run_geometry_and_loading(sap_model)
        run_analysis_and_results_extraction(sap_model, column_names + beam_names)

        # 执行可选的设计和内力提取流程
        run_design_and_force_extraction(workflow_state, column_names, beam_names)

    except SystemExit as e:
        print("\n--- 脚本已中止 ---")
        if e.code != 0:
            print(f"退出代码: {e.code}")
    except Exception as e:
        print("\n--- 未预料的运行时错误 ---")
        print(f"错误类型: {type(e).__name__}: {e}")
        traceback.print_exc()
        cleanup_etabs_on_error()
        sys.exit(1)
    finally:
        generate_final_report(script_start_time, workflow_state)
        if not ATTACH_TO_INSTANCE:
            print("脚本执行完毕，ETABS 将保持打开状态。")


if __name__ == "__main__":
    main()
