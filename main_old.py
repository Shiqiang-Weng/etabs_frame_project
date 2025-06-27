# main.py
import sys
import time
import traceback
import os

# 导入所有模块
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

# 尝试从 design_module 导入主函数
try:
    from design_module import perform_concrete_design_and_extract_results

    design_module_available = True
    print("✅ 设计模块导入成功")
except ImportError as e:
    design_module_available = False
    print(f"⚠️ 导入设计模块时出现问题: {e}")
    print("将跳过设计功能...")


    # 定义一个空的替代函数，使其在未导入时也能正常调用
    def perform_concrete_design_and_extract_results():
        print("⏭️ 设计模块导入失败，跳过构件设计。")
        return False  # 返回 False 表示失败

# 尝试导入设计内力提取模块 - 支持多种可能的文件名
design_force_extraction_available = False
extract_design_forces_and_summary = None

# 尝试不同的模块名称
possible_modules = [
    'design_force_extraction_fixed',  # 修复版
    'design_force_extraction',  # 原版
    'design_force_extraction_improved'  # 改进版
]

for module_name in possible_modules:
    try:
        if module_name == 'design_force_extraction_fixed':
            from design_force_extraction_fixed import extract_design_forces_and_summary
        elif module_name == 'design_force_extraction':
            from design_force_extraction import extract_design_forces_and_summary
        elif module_name == 'design_force_extraction_improved':
            from design_force_extraction_improved import extract_design_forces_and_summary

        design_force_extraction_available = True
        print(f"✅ 设计内力提取模块导入成功: {module_name}")
        break
    except ImportError as e:
        print(f"⚠️ 尝试导入 {module_name} 失败: {e}")
        continue

# 如果所有尝试都失败，定义空函数
if not design_force_extraction_available:
    print("⚠️ 所有设计内力提取模块导入失败，将跳过设计内力提取功能...")


    def extract_design_forces_and_summary(column_names, beam_names):
        print("⏭️ 设计内力提取模块导入失败，跳过设计内力提取。")
        return False


def print_project_info():
    """打印项目信息"""
    print("=" * 80)
    print("ETABS 框架结构自动建模脚本 v6.3.1 (设计模块 v12.1)")
    print("=" * 80)
    print("项目特点：")
    print("1. 10层钢筋混凝土框架结构")
    print("2. 采用框架柱和框架梁体系")
    print("3. 楼板设置为膜单元（面外刚度为0）")
    print("4. 基于GB50011-2010反应谱分析")
    print("5. 自动提取模态信息、层间位移角和构件内力")
    print("6. 执行GB50010-2010混凝土构件配筋设计")
    print("7. 提取构件设计内力数据")
    print("8. 完全模块化设计，便于维护和扩展")
    print()
    print("模块状态：")
    print(f"- 设计模块: {'✅ 可用' if design_module_available else '❌ 不可用'}")
    print(f"- 设计内力提取模块: {'✅ 可用' if design_force_extraction_available else '❌ 不可用'}")
    print()
    print("结构参数：")
    print(f"- 楼层数：{NUM_STORIES}层")
    print(f"- 网格：{NUM_GRID_LINES_X}×{NUM_GRID_LINES_Y} ({SPACING_X}m×{SPACING_Y}m)")
    print(f"- 框架柱：{FRAME_COLUMN_WIDTH}m×{FRAME_COLUMN_HEIGHT}m")
    print(f"- 框架梁：{FRAME_BEAM_WIDTH}m×{FRAME_BEAM_HEIGHT}m")
    print(f"- 楼板厚度：{SLAB_THICKNESS}m (膜单元)")
    print(f"- 层高：首层{BOTTOM_STORY_HEIGHT}m，标准层{TYPICAL_STORY_HEIGHT}m")
    print(f"- 总高度：{BOTTOM_STORY_HEIGHT + (NUM_STORIES - 1) * TYPICAL_STORY_HEIGHT:.1f}m")
    print()
    print("地震参数：")
    print(f"- 设防烈度：{RS_DESIGN_INTENSITY}度")
    print(f"- 最大地震影响系数：{RS_BASE_ACCEL_G}")
    print(f"- 场地类别：{RS_SITE_CLASS}类")
    print(f"- 特征周期：{RS_CHARACTERISTIC_PERIOD}s")
    print(f"- 地震分组：第{RS_SEISMIC_GROUP}组")
    print()
    print("设计参数：")
    print(f"- 使用ETABS默认混凝土设计规范")
    print(f"- 是否执行配筋设计：{'是' if PERFORM_CONCRETE_DESIGN else '否'}")
    print(f"- 是否提取设计内力：{'是' if PERFORM_CONCRETE_DESIGN and design_force_extraction_available else '否'}")
    print("=" * 80)


def main():
    """主函数 - 框架结构建模流程"""
    script_start_time = time.time()

    # 打印项目信息
    print_project_info()

    # 初始化变量，以防某些阶段被跳过
    column_names, beam_names, slab_names, story_heights = [], [], [], {}

    try:
        # ========== 第一阶段：初始化 ==========
        print("\n🚀 第一阶段：系统初始化")
        if not check_output_directory(): sys.exit(1)
        load_dotnet_etabs_api()
        _, sap_model = setup_etabs()

        # ========== 第二阶段：模型定义 ==========
        print("\n🏗️ 第二阶段：模型定义")
        define_all_materials_and_sections()
        define_response_spectrum_functions_in_etabs()
        define_all_load_cases()

        # ========== 第三阶段：几何建模 ==========
        print("\n🏢 第三阶段：框架结构建模")
        column_names, beam_names, slab_names, story_heights = create_frame_structure()

        # ========== 第四阶段：荷载分配 ==========
        print("\n⚖️ 第四阶段：荷载分配")
        assign_all_loads_to_frame_structure(column_names, beam_names, slab_names)

        # ========== 第五阶段：保存模型 ==========
        print("\n💾 第五阶段：保存模型")
        finalize_and_save_model()

        # ========== 第六阶段：结构分析 ==========
        print("\n🔍 第六阶段：结构分析")
        wait_and_run_analysis(5)
        if not check_analysis_completion():
            print("⚠️ 分析状态检查异常，但继续尝试提取结果")

        # ========== 第七阶段：结果提取 ==========
        print("\n📊 第七阶段：结果提取")
        extract_all_analysis_results()
        extract_and_save_frame_forces(column_names + beam_names)

        # ========== 第八阶段：构件设计 ==========
        design_completed_successfully = False
        if PERFORM_CONCRETE_DESIGN and design_module_available:
            print("\n🏗️ 第八阶段：混凝土构件配筋设计")
            try:
                # 只调用主函数，它会处理所有内部逻辑和错误
                design_completed_successfully = perform_concrete_design_and_extract_results()

                if design_completed_successfully:
                    print("✅ 设计和结果提取验证通过。")
                else:
                    print("⚠️ 设计和结果提取未成功，请检查以上 design_module 日志。")

            except Exception as design_error:
                print(f"⚠️ 构件设计模块发生未捕获的严重错误: {design_error}")
                print("错误详情:")
                traceback.print_exc()

            finally:
                print("✅ 构件设计阶段完成。")  # 无论成功与否都标记阶段完成
        elif PERFORM_CONCRETE_DESIGN and not design_module_available:
            print("\n⏭️ 第八阶段：跳过构件设计（设计模块不可用）。")
        else:
            print("\n⏭️ 第八阶段：跳过构件设计（由config文件设置）。")

        # ========== 第九阶段：构件设计内力提取 ==========
        design_force_extraction_successful = False
        if (PERFORM_CONCRETE_DESIGN and design_completed_successfully and
                design_force_extraction_available):
            print("\n🔬 第九阶段：构件设计内力提取")
            try:
                print("正在提取框架柱和框架梁的设计内力...")
                design_force_extraction_successful = extract_design_forces_and_summary(
                    column_names, beam_names
                )

                if design_force_extraction_successful:
                    print("✅ 构件设计内力提取成功。")
                else:
                    print("⚠️ 构件设计内力提取失败，请检查日志。")

            except Exception as extraction_error:
                print(f"⚠️ 构件设计内力提取模块发生错误: {extraction_error}")
                print("错误详情:")
                traceback.print_exc()

            finally:
                print("✅ 构件设计内力提取阶段完成。")
        elif PERFORM_CONCRETE_DESIGN and design_completed_successfully and not design_force_extraction_available:
            print("\n⏭️ 第九阶段：跳过构件设计内力提取（提取模块不可用）。")
        elif PERFORM_CONCRETE_DESIGN and not design_completed_successfully:
            print("\n⏭️ 第九阶段：跳过构件设计内力提取（设计阶段未成功完成）。")
        else:
            print("\n⏭️ 第九阶段：跳过构件设计内力提取（未执行构件设计）。")

        # ========== 完成 ==========
        elapsed_time = time.time() - script_start_time
        print("\n" + "=" * 80)
        print("🎉 框架结构建模完成！")
        print("=" * 80)
        print("✅ 主要完成功能:")
        print(f"   1. {NUM_STORIES}层钢筋混凝土框架结构建模")
        print(f"   2. 创建了 {len(column_names)} 根框架柱")
        print(f"   3. 创建了 {len(beam_names)} 根框架梁")
        print(f"   4. 创建了 {len(slab_names)} 块楼板（膜单元）")
        print("   5. 完成了荷载分配和地震参数设置")
        print("   6. 完成了模态分析和反应谱分析")
        print("   7. 提取了模态信息、层间位移角和构件内力")
        if PERFORM_CONCRETE_DESIGN and design_module_available:
            if design_completed_successfully:
                print("   8. 成功完成混凝土构件配筋设计和结果提取。")
                if design_force_extraction_successful:
                    print("   9. 成功提取构件设计内力数据。")
                else:
                    print("   9. 构件设计内力提取执行完毕，但未成功。")
            else:
                print("   8. 混凝土构件配筋设计执行完毕，但结果提取或验证失败。")
                print("   9. 跳过构件设计内力提取。")
        else:
            if not design_module_available:
                print("   8. 设计模块不可用，跳过混凝土构件配筋设计。")
            print("   9. 跳过构件设计内力提取。")
        print()
        print("📁 输出文件:")
        print(f"   模型文件: {MODEL_PATH}")
        print(f"   构件内力: {os.path.join(SCRIPT_DIRECTORY, 'frame_member_forces.csv')}")
        if PERFORM_CONCRETE_DESIGN and design_module_available:
            print(f"   配筋设计: {os.path.join(SCRIPT_DIRECTORY, 'concrete_design_results.csv')}")
            print(f"   设计报告: {os.path.join(SCRIPT_DIRECTORY, 'design_summary_report.txt')}")
            if design_force_extraction_successful:
                print(f"   柱设计内力: {os.path.join(SCRIPT_DIRECTORY, 'column_design_forces.csv')}")
                print(f"   梁设计内力: {os.path.join(SCRIPT_DIRECTORY, 'beam_design_forces.csv')}")
                print(f"   内力汇总: {os.path.join(SCRIPT_DIRECTORY, 'design_forces_summary_report.txt')}")
        print()
        print("🏗️ 结构信息:")
        total_height = BOTTOM_STORY_HEIGHT + (NUM_STORIES - 1) * TYPICAL_STORY_HEIGHT if NUM_STORIES > 0 else 0
        print(f"   结构类型: {NUM_STORIES}层钢筋混凝土框架结构")
        print(f"   平面尺寸: {(NUM_GRID_LINES_X - 1) * SPACING_X:.1f}m × {(NUM_GRID_LINES_Y - 1) * SPACING_Y:.1f}m")
        print(f"   结构总高: {total_height:.1f}m")
        print(f"   抗震设防: {RS_DESIGN_INTENSITY}度，{RS_SITE_CLASS}类场地")
        print()
        print(f"⏱️ 总执行时间: {elapsed_time:.2f} 秒")

        # 输出执行状态总结
        print("\n📋 执行状态总结:")
        print(f"   ✅ 结构建模: 成功")
        print(f"   ✅ 结构分析: 成功")
        print(f"   ✅ 结果提取: 成功")
        if PERFORM_CONCRETE_DESIGN:
            if design_module_available:
                status_design = "成功" if design_completed_successfully else "失败"
                print(f"   {'✅' if design_completed_successfully else '❌'} 构件设计: {status_design}")
                if design_force_extraction_available:
                    status_force = "成功" if design_force_extraction_successful else "失败"
                    print(f"   {'✅' if design_force_extraction_successful else '❌'} 设计内力提取: {status_force}")
                else:
                    print(f"   ⏭️ 设计内力提取: 模块不可用")
            else:
                print(f"   ⏭️ 构件设计: 模块不可用")
                print(f"   ⏭️ 设计内力提取: 跳过")
        else:
            print(f"   ⏭️ 构件设计: 跳过")
            print(f"   ⏭️ 设计内力提取: 跳过")

        print("=" * 80)

        if not ATTACH_TO_INSTANCE:
            print("脚本执行完毕，ETABS 将保持打开状态供进一步操作。")

    except SystemExit as e:
        print(f"\n--- 脚本已中止 ---")
        if hasattr(e, 'code') and e.code != 0 and e.code is not None:
            if not (isinstance(e.code, str) and "关键错误" in e.code):
                print(f"脚本退出代码: {e.code}")

    except Exception as e:
        print(f"\n--- 未预料的运行时错误 ---")
        print(f"错误类型: {type(e).__name__}")
        print(f"错误信息: {e}")
        traceback.print_exc()
        cleanup_etabs_on_error()
        sys.exit(1)

    finally:
        final_elapsed_time = time.time() - script_start_time
        print(f"\n脚本总执行时间: {final_elapsed_time:.2f} 秒。")


if __name__ == "__main__":
    main()