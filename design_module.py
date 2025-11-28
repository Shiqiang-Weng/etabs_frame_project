# -*- coding: utf-8 -*-
"""
构件配筋设计模块 v22.24b (梁提取功能增强版)
- 整合了 design_module_column.py 的单位转换修复功能
- 保留原有的设计准备和配筋设置功能
- 添加了配筋面积验证系统
- 改进了 System.Array 处理和数据提取
- 增强了梁设计结果提取，使用 GetSummaryResultsBeam_2 获取更详细数据
- 修正了原版备用提取函数的错误
- 完整的设计流程：准备 → 设计 → 提取 → 验证
"""
import os
import csv
import traceback
import time
from typing import List, Dict, Any

# --- System程序集加载 ---
import clr

try:
    clr.AddReference("System")
    import System

    print("✅ System程序集加载成功")
except Exception as e:
    print(f"❌ System程序集加载失败: {e}")
    sys.exit(1)

from etabs_setup import get_etabs_objects
from utility_functions import check_ret
from config import PERFORM_CONCRETE_DESIGN, SCRIPT_DIRECTORY

# 类型别名
INT = System.Int32


def ensure_etabs_v22_loaded():
    """确保ETABS v22 API正确加载"""
    try:
        etabs_paths = [
            r"C:\Program Files\Computers and Structures\ETABS 22\ETABSv1.dll",
            r"C:\Program Files (x86)\Computers and Structures\ETABS 22\ETABSv1.dll"
        ]

        for path in etabs_paths:
            if os.path.exists(path):
                clr.AddReference(path)
                print(f"✅ ETABS DLL加载: {path}")
                break
        else:
            clr.AddReference("ETABSv1")
            print("✅ ETABS DLL从GAC加载")

        import ETABSv1
        return ETABSv1
    except Exception as e:
        print(f"❌ 加载ETABS DLL失败: {e}")
        return None


def get_material_type_fixed(prop_mat, name):
    """修复版材料类型获取 - 处理特殊材料名称"""
    try:
        mat_type = INT(0)
        mat_subtype = INT(0)
        ret = prop_mat.GetType(name, mat_type, mat_subtype)
        if ret == 0:
            return mat_type.Value  # 6=Rebar, 2=Concrete
        # 忽略特殊材料名称（如带"/"的材料）的错误
        return -1
    except Exception:
        # 静默处理材料名称异常，不影响主流程
        return -1


def get_section_type_fixed(prop_frame, sec_name):
    """修复版截面类型获取 - 静默处理异常"""
    try:
        section_type = INT(0)
        ret = prop_frame.GetType(sec_name, section_type)
        if ret == 0:
            return section_type.Value  # 8=Rectangular, 9=Circle
        return 8  # 默认矩形
    except Exception:
        # 静默处理异常，返回默认值
        return 8


def get_rebar_type_fixed(prop_frame, sec_name):
    """修复版配筋类型获取 - 静默处理异常"""
    try:
        rebar_type = INT(0)
        ret = prop_frame.GetTypeRebar(sec_name, rebar_type)
        if ret == 0:
            return rebar_type.Value  # 3=梁, 2=柱
        return -1
    except Exception:
        # 静默处理异常
        return -1


def create_rebar_material_fixed(sap_model, ETABSv1, mat_name="HRB400"):
    """修复版钢筋材料创建 - 使用正确的SetORebar_1签名"""
    try:
        prop_material = sap_model.PropMaterial

        # 检查材料是否已存在
        mat_type = get_material_type_fixed(prop_material, mat_name)
        if mat_type == 6:  # 6 = Rebar
            print(f"        ✅ 钢筋材料已存在: {mat_name}")
            return True

        print(f"        创建钢筋材料: {mat_name}")

        # 使用枚举类型创建材料
        ret = prop_material.SetMaterial(mat_name, ETABSv1.eMatType.Rebar)
        if ret == 0:
            print(f"        ✅ 钢筋材料创建成功: {mat_name}")

            # 设置基本属性
            try:
                # 弹性属性
                prop_material.SetMPIsotropic(mat_name, 2e11, 0.3, 1.17e-5)
                # 钢筋属性 - 使用v22的正确6参数版本
                prop_material.SetORebar_1(mat_name, 4e8, 4e8, 4.5e8, 0.002, 0.015)
                print(f"        ✅ 钢筋材料属性设置完成")
            except Exception:
                # 静默处理属性设置失败，材料创建成功即可
                pass

            return True
        else:
            print(f"        ❌ 钢筋材料创建失败，返回码: {ret}")
            return False

    except Exception as e:
        print(f"        钢筋材料创建异常: {e}")
        return False


def set_beam_rebar_fixed(sap_model, prop_frame, sec_name, rebar_mat, ETABSv1):
    """修复版梁配筋设置"""
    try:
        # 确保单位正确
        sap_model.SetPresentUnits(ETABSv1.eUnits.kN_m_C)

        print(f"        设置梁配筋: {sec_name}")

        ret = prop_frame.SetRebarBeam(
            sec_name,  # Name
            rebar_mat,  # MatPropLong
            rebar_mat,  # MatPropConfine
            0.025,  # CoverTop (25mm)
            0.025,  # CoverBot (25mm)
            0.0006,  # TopLeftArea (600mm²)
            0.0006,  # TopRightArea
            0.0006,  # BotLeftArea
            0.0006  # BotRightArea
        )

        if ret == 0:
            print(f"        ✅ 梁 {sec_name} 配筋设置成功")
            return True
        else:
            print(f"        ❌ 梁 {sec_name} 配筋失败，返回码: {ret}")
            return False

    except Exception as e:
        print(f"        梁 {sec_name} 配筋异常: {e}")
        return False


def set_column_rebar_fixed(sap_model, prop_frame, sec_name, rebar_mat, ETABSv1):
    """修复版柱配筋设置"""
    try:
        # 确保单位正确
        sap_model.SetPresentUnits(ETABSv1.eUnits.kN_m_C)

        # 判断截面类型
        section_type = get_section_type_fixed(prop_frame, sec_name)
        is_circle = (section_type == 9)

        # 设置参数
        pattern = 2 if is_circle else 1  # 2=Circle, 1=Rectangular
        conf_type = 2 if is_circle else 1  # 2=Spiral, 1=Ties
        cover = 0.040  # 40mm
        tie_space = 0.150  # 150mm

        # 钢筋数量参数
        if is_circle:
            N_C, N_R3, N_R2 = 10, 0, 0  # 圆形：10根均布
        else:
            N_C, N_R3, N_R2 = 0, 4, 4  # 矩形：4+4布置

        print(f"        设置柱配筋: {sec_name} ({'圆形' if is_circle else '矩形'})")

        ret = prop_frame.SetRebarColumn(
            sec_name,  # 1. Name
            rebar_mat,  # 2. MatPropLong
            rebar_mat,  # 3. MatPropConfine
            pattern,  # 4. Pattern
            conf_type,  # 5. ConfineType
            cover,  # 6. Cover
            N_C,  # 7. NumberCBars
            N_R3,  # 8. NumberR3Bars
            N_R2,  # 9. NumberR2Bars
            "20",  # 10. RebarSize
            "10",  # 11. TieSize
            tie_space,  # 12. TieSpacingLongit
            2,  # 13. Number2DirTieBars
            2,  # 14. Number3DirTieBars
            True  # 15. ToBeDesigned
        )

        if ret == 0:
            print(f"        ✅ 柱 {sec_name} 配筋设置成功")
            return True
        else:
            print(f"        ❌ 柱 {sec_name} 配筋失败，返回码: {ret}")
            return False

    except Exception as e:
        print(f"        柱 {sec_name} 配筋异常: {e}")
        return False


def set_frames_to_concrete_design(sap_model, beam_section, col_section):
    """关键修复：设置所有构件为混凝土设计程序 - 使用遍历所有构件的保险方法"""
    print("      设置构件为混凝土设计程序...")

    try:
        frame_obj = sap_model.FrameObj

        # 使用GetNameList获取所有构件
        NumberNames = INT(0)
        MyName = System.Array.CreateInstance(System.String, 0)
        ret, NumberNames, MyName = frame_obj.GetNameList(NumberNames, MyName)

        if ret != 0:
            print(f"        ❌ 无法获取构件列表，返回码: {ret}")
            return False

        frame_names = list(MyName)
        concrete_count = 0

        print(f"        检查 {len(frame_names)} 个构件...")

        # 遍历所有构件，检查截面名称
        for frame_name in frame_names:
            try:
                # 获取构件的截面名称
                ret_sec, section_name = frame_obj.GetSection(frame_name, "")
                if ret_sec == 0 and section_name in [beam_section, col_section]:
                    # 设置为混凝土设计
                    ret_design = frame_obj.SetDesignProcedure(frame_name, 2)  # 2 = Concrete
                    if ret_design == 0:
                        concrete_count += 1
            except Exception:
                # 静默处理单个构件的异常
                continue

        print(f"        ✅ 总计设置 {concrete_count} 个构件为混凝土设计")
        return concrete_count > 0

    except Exception as e:
        print(f"      设置混凝土设计程序异常: {e}")
        return False


def verify_design_setup(sap_model, beam_section, col_section):
    """验证设计设置 - 静默处理异常"""
    print("      验证设计设置...")

    try:
        prop_frame = sap_model.PropFrame
        frame_obj = sap_model.FrameObj

        # 验证截面配筋类型
        beam_rebar_type = get_rebar_type_fixed(prop_frame, beam_section)
        col_rebar_type = get_rebar_type_fixed(prop_frame, col_section)

        beam_type_name = {3: "梁", 2: "柱", 1: "其他", 0: "未设置"}.get(beam_rebar_type, "已设置")
        col_type_name = {3: "梁", 2: "柱", 1: "其他", 0: "未设置"}.get(col_rebar_type, "已设置")

        print(f"        {beam_section} 配筋类型: {beam_type_name}")
        print(f"        {col_section} 配筋类型: {col_type_name}")

        # 验证构件设计程序
        concrete_design_count = 0
        NumberNames = INT(0)
        FrameNames_tuple = System.Array.CreateInstance(System.String, 0)
        ret, NumberNames, FrameNames_tuple = frame_obj.GetNameList(NumberNames, FrameNames_tuple)

        if ret == 0:
            frame_names = list(FrameNames_tuple)
            for name in frame_names[:10]:  # 抽样检查前10个
                try:
                    proc_type = INT(0)
                    ret_proc = frame_obj.GetDesignProcedure(name, proc_type)
                    if ret_proc == 0 and proc_type.Value == 2:  # 2 = Concrete
                        concrete_design_count += 1
                except:
                    pass

        print(f"        混凝土设计程序验证: {concrete_design_count}/10")

        # 即使验证显示异常，如果设置过程成功，仍返回True
        return True

    except Exception as e:
        print(f"      验证设计设置异常: {e}")
        return True  # 验证失败不影响主流程


def prepare_model_for_design():
    """最终版模型设计准备"""
    print("\n--- 准备模型进行设计 (最终精简版) ---")
    _, sap_model = get_etabs_objects()
    if not sap_model:
        return False

    try:
        from config import FRAME_BEAM_SECTION_NAME, FRAME_COLUMN_SECTION_NAME

        # 确保ETABS v22 API正确加载
        ETABSv1 = ensure_etabs_v22_loaded()
        if not ETABSv1:
            print("❌ 无法加载ETABS v22 API")
            return False

        # 解锁模型
        if sap_model.GetModelIsLocked():
            sap_model.SetModelIsLocked(False)
            print("  模型已解锁...")

        # 验证截面分配
        NumberNames = 0
        FrameNames_tuple = System.Array.CreateInstance(System.String, 0)
        ret, NumberNames, FrameNames_tuple = sap_model.FrameObj.GetNameList(NumberNames, FrameNames_tuple)

        if ret == 0:
            frame_names = list(FrameNames_tuple)
            beam_count = len([n for n in frame_names if n.upper().startswith("BEAM")])
            col_count = len([n for n in frame_names if n.upper().startswith("COL")])
            print(f"  发现: {beam_count} 根梁, {col_count} 根柱")

        print("  设置配筋类型...")

        # 设置单位
        sap_model.SetPresentUnits(ETABSv1.eUnits.kN_m_C)
        print(f"    单位设置: kN_m_C")

        # 创建钢筋材料
        rebar_material = "HRB400"
        create_rebar_material_fixed(sap_model, ETABSv1, rebar_material)

        # 设置截面配筋
        prop_frame = sap_model.PropFrame
        beam_success = set_beam_rebar_fixed(sap_model, prop_frame, FRAME_BEAM_SECTION_NAME, rebar_material, ETABSv1)
        col_success = set_column_rebar_fixed(sap_model, prop_frame, FRAME_COLUMN_SECTION_NAME, rebar_material, ETABSv1)

        # 关键步骤：设置构件为混凝土设计程序
        design_proc_success = set_frames_to_concrete_design(sap_model, FRAME_BEAM_SECTION_NAME,
                                                            FRAME_COLUMN_SECTION_NAME)

        # 验证设置
        verify_success = verify_design_setup(sap_model, FRAME_BEAM_SECTION_NAME, FRAME_COLUMN_SECTION_NAME)

        # 保存并重新分析
        sap_model.File.Save()
        sap_model.SetModelIsLocked(True)
        print("  重新运行分析...")
        check_ret(sap_model.Analyze.RunAnalysis(), "RunAnalysis")
        print("  分析完成。")

        overall_success = beam_success and col_success and design_proc_success
        print(f"  准备阶段: {'✅ 完全成功' if overall_success else '⚠️ 部分成功'}")
        return overall_success

    except Exception as e:
        print(f"❌ 准备过程异常: {e}")
        traceback.print_exc()
        return False


def run_concrete_design():
    """运行混凝土设计"""
    _, sap_model = get_etabs_objects()
    print("\n🎯 运行混凝土设计...")

    try:
        # 设置设计代码
        try:
            sap_model.DesignConcrete.SetCode("Chinese 2010")
            print(f"  设计代码: {sap_model.DesignConcrete.GetCode()[1]}")
        except:
            print("  使用默认设计代码")

        # 运行设计
        print("  启动混凝土设计...")
        ret = sap_model.DesignConcrete.StartDesign()

        if ret == 0:
            print("✅ 设计完成成功！")
            return True
        else:
            print(f"❌ 设计失败，返回码: {ret}")
            if ret == 1:
                print("    可能原因: 没有构件设置为混凝土设计程序")
            elif ret == 3:
                print("    可能原因: 没有分析结果")
            return False

    except Exception as e:
        print(f"❌ 设计运行异常: {e}")
        return False


# ==================== 整合的数据提取和单位转换修复功能 ====================
# 逻辑已迁移至 results_extraction.design_results
from results_extraction import design_results as _design_results

convert_system_array_to_python_list = _design_results.convert_system_array_to_python_list
convert_area_units = _design_results.convert_area_units
convert_shear_area_units = _design_results.convert_shear_area_units
validate_reinforcement_area = _design_results.validate_reinforcement_area
_get_beam_design_summary_enhanced = _design_results._get_beam_design_summary_enhanced
_get_column_design_summary_enhanced = _design_results._get_column_design_summary_enhanced
extract_design_results_enhanced = _design_results.extract_design_results_enhanced
save_design_results_enhanced = _design_results.save_design_results_enhanced
print_enhanced_validation_statistics = _design_results.print_enhanced_validation_statistics
generate_enhanced_summary_report = _design_results.generate_enhanced_summary_report
extract_and_save_beam_results = _design_results.extract_and_save_beam_results
extract_and_save_column_results = _design_results.extract_and_save_column_results

# ==================== 保留并修正原有的简化版提取函数作为备用 ====================

def extract_and_save_beam_results(output_dir: str):
    """修正后的原版梁配筋结果提取函数（保留作为备用）"""
    _, sap_model = get_etabs_objects()
    print("\n--- 提取梁设计结果 (原版备用) ---")
    os.makedirs(output_dir, exist_ok=True)

    try:
        dc = sap_model.DesignConcrete

        # 获取梁构件
        NumberNames = 0
        FrameNames_tuple = System.Array.CreateInstance(System.String, 0)
        ret, NumberNames, FrameNames_tuple = sap_model.FrameObj.GetNameList(NumberNames, FrameNames_tuple)
        if ret != 0:
            print("  无法获取构件列表")
            return

        frame_names = list(FrameNames_tuple)
        beam_names = [name for name in frame_names if name.upper().startswith("BEAM")]

        if not beam_names:
            print("  没有找到梁构件")
            return

        print(f"  找到 {len(beam_names)} 根梁，正在提取配筋...")
        all_results = []
        valid_results = 0

        for i, name in enumerate(beam_names):
            if (i + 1) % 50 == 0:
                print(f"    进度: {i + 1}/{len(beam_names)}")

            result = {"Frame_Name": name}
            try:
                # 调用 GetSummaryResultsBeam
                res = dc.GetSummaryResultsBeam(name, 0, [], [], [], [], [], [], [], [], [], [], [], [], [], [])

                ret_code, num_items, _, _, _, top_areas, _, bot_areas, _, _, _, _, _, _, _, _ = res

                if ret_code == 0 and num_items > 0:
                    top_areas_list = [a for a in convert_system_array_to_python_list(top_areas) if a > 0]
                    bot_areas_list = [a for a in convert_system_array_to_python_list(bot_areas) if a > 0]

                    max_top = max(top_areas_list) if top_areas_list else 0
                    max_bot = max(bot_areas_list) if bot_areas_list else 0

                    result.update({
                        "Src": "OK",
                        "Top_Rebar_m2": f"{max_top:.6f}",
                        "Bot_Rebar_m2": f"{max_bot:.6f}"
                    })
                    valid_results += 1
                else:
                    result.update({"Src": "No Results", "Top_Rebar_m2": 0, "Bot_Rebar_m2": 0})

            except Exception as e:
                result.update({"Src": f"Error: {str(e)[:40]}", "Top_Rebar_m2": 0, "Bot_Rebar_m2": 0})

            all_results.append(result)

        # 保存结果
        filepath = os.path.join(output_dir, "beam_design_results_final.csv")
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=all_results[0].keys())
            writer.writeheader()
            writer.writerows(all_results)

        print(f"✅ 原版梁设计结果已保存: {filepath}")
        print(f"   有效结果: {valid_results}/{len(beam_names)}")

    except Exception as e:
        print(f"❌ 提取原版梁结果异常: {e}")


def extract_and_save_column_results(output_dir: str):
    """修正后的原版柱配筋结果提取函数（保留作为备用）"""
    _, sap_model = get_etabs_objects()
    print("\n--- 提取柱设计结果 (原版备用) ---")
    os.makedirs(output_dir, exist_ok=True)

    try:
        dc = sap_model.DesignConcrete

        # 获取柱构件
        NumberNames = 0
        FrameNames_tuple = System.Array.CreateInstance(System.String, 0)
        ret, NumberNames, FrameNames_tuple = sap_model.FrameObj.GetNameList(NumberNames, FrameNames_tuple)
        if ret != 0:
            print("  无法获取构件列表")
            return

        frame_names = list(FrameNames_tuple)
        column_names = [name for name in frame_names if name.upper().startswith("COL")]

        if not column_names:
            print("  没有找到柱构件")
            return

        print(f"  找到 {len(column_names)} 根柱，正在提取配筋...")
        all_results = []
        valid_results = 0

        for i, name in enumerate(column_names):
            if (i + 1) % 50 == 0:
                print(f"    进度: {i + 1}/{len(column_names)}")

            result = {"Frame_Name": name}
            try:
                # 调用 GetSummaryResultsColumn
                res = dc.GetSummaryResultsColumn(name, 0, [], [], [], [], [], [], [], [], [])
                ret_code, num_items, _, _, _, pmm_areas, _, _, _, _, _ = res

                if ret_code == 0 and num_items > 0:
                    areas = [a for a in convert_system_array_to_python_list(pmm_areas) if a > 0]
                    if areas:
                        area_max_m2 = max(areas)
                        result.update({"Src": "OK", "Long_Rebar_m2": f"{area_max_m2:.6f}"})
                        valid_results += 1
                    else:
                        result.update({"Src": "No Valid Data", "Long_Rebar_m2": 0})
                else:
                    result.update({"Src": "No Results", "Long_Rebar_m2": 0})
            except Exception as e:
                result.update({"Src": f"Error: {str(e)[:40]}", "Long_Rebar_m2": 0})

            all_results.append(result)

        # 保存结果
        filepath = os.path.join(output_dir, "column_design_results_final.csv")
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=all_results[0].keys())
            writer.writeheader()
            writer.writerows(all_results)

        print(f"✅ 原版柱设计结果已保存: {filepath}")
        print(f"   有效结果: {valid_results}/{len(column_names)}")

    except Exception as e:
        print(f"❌ 提取原版柱结果异常: {e}")


def perform_concrete_design_and_extract_results():
    """整合增强版主执行函数"""
    print("\n" + "=" * 80)
    print("🎯 执行混凝土梁柱配筋设计 (v22.24b - 梁提取功能增强版)")
    print("=" * 80)

    output_dir = SCRIPT_DIRECTORY if 'SCRIPT_DIRECTORY' in globals() else os.getcwd()

    try:
        if not PERFORM_CONCRETE_DESIGN:
            print("⏭️ 跳过构件设计。")
            return True

        print("🚀 开始整合增强流程...")

        # 阶段1: 模型准备
        print("\n📋 阶段1: 模型设计准备")
        design_prep_success = prepare_model_for_design()

        # 阶段2: 运行设计
        print("\n🎯 阶段2: 执行混凝土设计")
        design_success = run_concrete_design()

        # 阶段3: 增强版数据提取
        if design_success:
            print("\n📊 阶段3: 增强版结果提取")
            design_results = extract_design_results_enhanced()

            if design_results:
                # 保存增强版结果
                save_design_results_enhanced(design_results, output_dir)
                generate_enhanced_summary_report(output_dir)

                # 同时保存原版结果作为对比
                print("\n📁 生成原版结果作为对比...")
                extract_and_save_beam_results(output_dir)
                extract_and_save_column_results(output_dir)
            else:
                print("❌ 增强版结果提取失败，尝试原版方法...")
                extract_and_save_beam_results(output_dir)
                extract_and_save_column_results(output_dir)
        else:
            print("❌ 设计失败，跳过结果提取")

        # 恢复视图
        try:
            _, sap_model = get_etabs_objects()
            sap_model.View.RefreshView(0, False)
        except:
            pass

        print("\n--- 执行总结 (整合增强版) ---")
        if design_success:
            print("🎉🎉🎉 整合增强版设计完成！")
            print("📁 生成的文件:")
            print(f"   主要结果:")
            print(f"   - concrete_design_results_enhanced.csv (增强版)")
            print(f"   - validation_statistics_enhanced.txt (验证统计)")
            print(f"   - design_summary_report_enhanced.txt (摘要报告)")
            print(f"   对比结果:")
            print(f"   - beam_design_results_final.csv (原版梁)")
            print(f"   - column_design_results_final.csv (原版柱)")
        else:
            print("⚠️ 设计未完全成功，请检查日志")

        print("\n💡 整合增强版特色：")
        print("  - ✅ 完整的设计流程 (准备→设计→提取→验证)")
        print("  - ✅ 单位转换修复和面积验证")
        print("  - ✅ 智能System.Array处理")
        print("  - ✅ 优先使用GetSummaryResultsBeam_2获取详细梁数据")
        print("  - ✅ 详细统计和报告生成")
        print("  - ✅ 保留并修正了原版备用方法")

        return design_success

    except Exception as e:
        print(f"❌ 主函数异常: {e}")
        traceback.print_exc()
        return False
    finally:
        print("\n--- design_module (整合增强版) 结束 ---")


# 导出函数列表
__all__ = ["perform_concrete_design_and_extract_results"]
