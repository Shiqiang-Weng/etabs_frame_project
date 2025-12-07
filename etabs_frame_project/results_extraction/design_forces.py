# design_forces.py (migrated from design_force_extraction_fixed.py)
"""
构件设计内力提取模块（已迁移至 results_extraction 包）

用于提取混凝土构件设计后的控制内力、配筋信息和设计总结。

功能概览：
  - 提取 Design Forces - Columns (柱设计内力)
  - 提取 Concrete Beam Flexure Envelope - Chinese 2010（梁弯矩包络）
  - 提取 Concrete Beam Shear Envelope - Chinese 2010（梁剪力包络）
  - 提取 Concrete Column Shear Envelope - Chinese 2010（柱剪力包络）
  - 提取 Concrete Joint Envelope - Chinese 2010（节点包络）
  - 提取混凝土柱 P-M-M 设计内力：
      * 直接导出 Concrete Column PMM Envelope - Chinese 2010（或兼容表名）完整表
      * 通过 DesignConcrete.GetSummaryResultsColumn 生成汇总表
  - 改进 CSV 过滤逻辑，支持“不过滤，整表导出”

说明：
  - 对柱 P-M-M 原始表中的 P、M Major、M Minor、At Major、At Minor 等字段，
    本脚本仅做“原样导出”，不进行任何单位换算，保证与 ETABS 交互界面显示完全一致。
"""

import os
import csv
import traceback
from datetime import datetime

from common.config import DESIGN_DATA_DIR
from common.config import *  # noqa: F401,F403
from common.etabs_setup import get_sap_model, ensure_etabs_ready
from common.utility_functions import check_ret, arr
from results_extraction.export_utils import (
    build_component_filter,
    export_table_to_csv,
    filter_csv_rows,
    log_table_status,
)


# =============================================================================
# 顶层入口函数
# =============================================================================
def extract_design_forces_and_summary(column_names, beam_names):
    """
    提取构件设计内力的主函数

    Args:
        column_names (list): 框架柱名称列表（通常为柱的 UniqueName/或 Label）
        beam_names (list): 框架梁名称列表

    Returns:
        bool: 提取是否成功
    """
    print("=" * 60)
    print("🔬 开始构件设计内力提取")
    print("=" * 60)

    try:
        # ------------------------------------------------------------------ #
        # 0) 确保 ETABS 连接正常
        # ------------------------------------------------------------------ #
        if not ensure_etabs_ready():
            print("❌ 无法建立ETABS连接，请确保ETABS已打开并已加载模型。")
            return False

        sap_model = get_sap_model()
        if sap_model is None:
            print("❌ 无法获取ETABS模型对象。")
            return False

        print("✅ ETABS连接正常，模型对象获取成功")

        # ------------------------------------------------------------------ #
        # 1) 检查设计是否完成 & 关键设计表是否可用
        # ------------------------------------------------------------------ #
        if not check_design_completion(sap_model):
            print("❌ 设计未完成或设计表格不可用，无法提取设计内力")
            return False

        # ------------------------------------------------------------------ #
        # 2) 做一些简单的 API 调试输出（可选）
        # ------------------------------------------------------------------ #
        print("🔍 开始API调试分析...")
        test_simple_api_call(sap_model, "Design Forces - Columns")
        test_simple_api_call(sap_model, "Concrete Beam Flexure Envelope - Chinese 2010")
        test_simple_api_call(sap_model, "Concrete Column Shear Envelope - Chinese 2010")
        test_simple_api_call(sap_model, "Concrete Joint Envelope - Chinese 2010")

        # ------------------------------------------------------------------ #
        # 3) 提取框架柱设计内力 (Design Forces - Columns)
        # ------------------------------------------------------------------ #
        print("📊 正在提取框架柱设计内力...")
        column_design_success = extract_design_forces_simple(
            sap_model,
            "Design Forces - Columns",
            column_names,
            "column_design_forces.csv",
        )

        if not column_design_success:
            print("🔄 简化方法失败，尝试备用柱设计内力提取方法...")
            column_design_success = extract_column_design_forces(
                sap_model, column_names
            )

        # ------------------------------------------------------------------ #
        # 3.5) 提取混凝土柱 P-M-M 设计内力
        # ------------------------------------------------------------------ #
        print("📊 正在提取混凝土柱 P-M-M 设计内力 (Concrete Column PMM / Summary)...")
        column_pmm_success = extract_column_pmm_design_forces(sap_model, column_names)
        if column_pmm_success:
            print(
                "✅ 混凝土柱 P-M-M 设计内力提取成功: "
                "column_pmm_design_forces_raw.csv / column_pmm_design_summary.csv"
            )
        else:
            print("⚠️ 未能提取柱 P-M-M 设计内力表 (Concrete Column PMM / Summary)。")

        # ------------------------------------------------------------------ #
        # 4) 提取框架梁弯矩包络 (Concrete Beam Flexure Envelope - Chinese 2010)
        # ------------------------------------------------------------------ #
        print("📊 正在提取框架梁设计包络...")
        beam_table_to_extract = "Concrete Beam Flexure Envelope - Chinese 2010"
        beam_output_filename = "beam_flexure_envelope.csv"
        print(f"🎯 目标表格: {beam_table_to_extract}")

        # 不按构件名过滤，整表导出
        beam_design_success = extract_design_forces_simple(
            sap_model, beam_table_to_extract, None, beam_output_filename
        )

        # 如果简化方法失败，尝试旧版表格
        if not beam_design_success:
            print("🔄 简化方法失败，尝试提取旧版内力表 Design Forces - Beams ...")
            beam_design_success = extract_design_forces_simple(
                sap_model, "Design Forces - Beams", beam_names, "beam_design_forces.csv"
            )
            if not beam_design_success:
                print("🔄 再次失败，尝试备用梁设计内力提取方法...")
                beam_design_success = extract_beam_design_forces(
                    sap_model, beam_names
                )

        # ------------------------------------------------------------------ #
        # 5) 提取混凝土梁剪力包络 (Concrete Beam Shear Envelope - Chinese 2010)
        # ------------------------------------------------------------------ #
        print("📊 正在提取混凝土梁剪力包络 (Concrete Beam Shear Envelope - Chinese 2010)...")
        beam_shear_success = extract_design_forces_simple(
            sap_model,
            "Concrete Beam Shear Envelope - Chinese 2010",
            None,
            "beam_shear_envelope.csv",
        )
        if beam_shear_success:
            print("✅ 梁剪力包络提取成功: beam_shear_envelope.csv")
        else:
            print("⚠️ 梁剪力包络提取失败 (表格可能不存在或无数据)")

        # ------------------------------------------------------------------ #
        # 6) 提取混凝土柱剪力包络 (Concrete Column Shear Envelope - Chinese 2010)
        # ------------------------------------------------------------------ #
        print("📊 正在提取混凝土柱剪力包络 (Concrete Column Shear Envelope - Chinese 2010)...")
        column_shear_success = extract_design_forces_simple(
            sap_model,
            "Concrete Column Shear Envelope - Chinese 2010",
            None,
            "column_shear_envelope.csv",
        )
        if column_shear_success:
            print("✅ 柱剪力包络提取成功: column_shear_envelope.csv")
        else:
            print("⚠️ 柱剪力包络提取失败 (表格可能不存在或无数据)")

        # ------------------------------------------------------------------ #
        # 7) 提取混凝土节点包络 (Concrete Joint Envelope - Chinese 2010)
        # ------------------------------------------------------------------ #
        print("📊 正在提取混凝土节点包络 (Concrete Joint Envelope - Chinese 2010)...")
        joint_envelope_success = extract_design_forces_simple(
            sap_model,
            "Concrete Joint Envelope - Chinese 2010",
            None,
            "joint_envelope.csv",
        )
        if joint_envelope_success:
            print("✅ 节点包络提取成功: joint_envelope.csv")
        else:
            print("⚠️ 节点包络提取失败 (表格可能不存在或无数据)")

        # ------------------------------------------------------------------ #
        # 8) 根据提取结果生成汇总报告
        # ------------------------------------------------------------------ #
        # 这里仍然以“柱设计内力 + 梁弯矩包络是否成功”为主条件，
        # 梁剪力 / 柱剪力 / 节点包络均视为增强信息，不影响报告生成。
        csv_extraction_success = column_design_success and beam_design_success
        summary_success = False

        if csv_extraction_success:
            print("✅ CSV数据提取完成，正在生成汇总报告...")
            summary_success = generate_summary_report(column_names, beam_names)
            print_extraction_summary()
        else:
            print("⚠️ 部分或全部CSV设计内力提取失败，不生成汇总报告。")

        overall_success = csv_extraction_success and summary_success

        if overall_success:
            print("\n✅ 所有构件设计内力提取任务成功完成。")
        else:
            print("\n⚠️ 部分设计内力提取任务失败，请检查以上日志。")

        return overall_success

    except Exception as e:
        print(f"❌ 构件设计内力提取过程中发生严重错误: {e}")
        traceback.print_exc()
        return False


# =============================================================================
# 设计完成状态检查（已加入 PMM Envelope + 梁剪力表）
# =============================================================================
def check_design_completion(sap_model):
    """
    检查设计是否已完成。
    使用数据库表方式检查常见设计结果表是否可用。
    """
    try:
        print("🔍 正在检查设计完成状态...")

        from common.etabs_api_loader import get_api_objects

        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载，无法检查设计状态")
            return False

        db = sap_model.DatabaseTables

        # 要检查的设计表格列表（含新表 + 兼容旧表名）
        design_tables_to_check = [
            "Design Forces - Beams",
            "Design Forces - Columns",
            "Concrete Beam Flexure Envelope - Chinese 2010",
            "Concrete Beam Shear Envelope - Chinese 2010",
            "Concrete Column Shear Envelope - Chinese 2010",
            "Concrete Joint Envelope - Chinese 2010",

            # ★ 关键：真正的柱 PMM 包络表，一般与交互界面一致
            "Concrete Column PMM Envelope - Chinese 2010",

            # 兼容旧名/其它版本：
            "Concrete Column PMM - Chinese 2010",
            "Concrete Column Envelope - Chinese 2010",

            "Concrete Column Design - P-M-M Design Forces - Chinese 2010",
            "Concrete Column Design - P-M-M Design Forces",
            "Concrete Beam Design - Flexural & Shear Forces",
        ]

        # 这些重要表如果不可用，要给出明显提示
        important_tables_for_warning = [
            "Concrete Column PMM Envelope - Chinese 2010",
            "Concrete Column PMM - Chinese 2010",
            "Concrete Column Design - P-M-M Design Forces - Chinese 2010",
            "Concrete Column Design - P-M-M Design Forces",
            "Concrete Beam Design - Flexural & Shear Forces",
            "Concrete Beam Flexure Envelope - Chinese 2010",
            "Concrete Beam Shear Envelope - Chinese 2010",
            "Concrete Column Shear Envelope - Chinese 2010",
            "Concrete Joint Envelope - Chinese 2010",
        ]

        found_tables = []

        for table_key in design_tables_to_check:
            try:
                field_key_list = System.Array.CreateInstance(System.String, 1)
                field_key_list[0] = ""

                group_name = ""
                table_version = System.Int32(0)
                fields_keys_included = System.Array.CreateInstance(System.String, 0)
                number_records = System.Int32(0)
                table_data = System.Array.CreateInstance(System.String, 0)

                ret = db.GetTableForDisplayArray(
                    table_key,
                    field_key_list,
                    group_name,
                    table_version,
                    fields_keys_included,
                    number_records,
                    table_data,
                )

                if isinstance(ret, tuple):
                    error_code = ret[0]
                    if error_code == 0:
                        found_tables.append(table_key)
                        print(f"✅ 找到设计表格: {table_key}")
                        if len(ret) > 5:
                            try:
                                record_array = ret[5]
                                record_count = (
                                    len(record_array)
                                    if hasattr(record_array, "__len__")
                                    else 0
                                )
                                print(f"   📊 记录数组长度(元素数): {record_count}")
                            except Exception:
                                pass
                    else:
                        if table_key in important_tables_for_warning:
                            print(
                                f"ℹ️ 表格当前不可用: {table_key} (错误码: {error_code})"
                            )
                elif ret == 0:
                    found_tables.append(table_key)
                    print(f"✅ 找到设计表格: {table_key}")
                else:
                    if table_key in important_tables_for_warning:
                        print(f"ℹ️ 表格当前不可用: {table_key} (返回码: {ret})")

            except Exception as e:
                print(f"⚠️ 检查表格 {table_key} 时出错: {str(e)}")
                continue

        if len(found_tables) >= 2:
            print(f"✅ 成功找到 {len(found_tables)} 个设计表格，可以继续提取。")
            return True
        elif len(found_tables) > 0:
            print(
                f"⚠️ 只找到 {len(found_tables)} 个设计表格，可能设计未完全完成，但仍尝试继续。"
            )
            return True
        else:
            print("❌ 未找到任何设计表格")
            print("💡 请确保已完成混凝土设计计算:")
            print("   1. Design → Concrete Frame Design → Start Design/Check of Structure")
            print("   2. 等待设计计算完成")
            print("   3. 检查是否有设计错误或警告")
            return False

    except Exception as e:
        print(f"❌ 检查设计完成状态时发生严重错误: {e}")
        traceback.print_exc()
        return False


# =============================================================================
# 通用的简化 CSV 导出方法
# =============================================================================
def extract_design_forces_simple(sap_model, table_key, component_names, output_filename):
    """
    简化的设计内力提取方法（DatabaseTables.GetTableForDisplayCSVFile）

    Args:
        sap_model: ETABS SapModel
        table_key (str): 数据库表键，例如 "Design Forces - Columns"
        component_names (list|None): 需要过滤的构件名称（UniqueName/Label），None 表示整表导出
        output_filename (str): 输出 CSV 文件名（不含路径，脚本自动拼 DESIGN_DATA_DIR）

    Returns:
        bool: 是否导出成功（以及是否至少写出了一条记录）
    """
    try:
        print(f"?? 简化提取方法 - 表格: {table_key}")

        from common.etabs_api_loader import get_api_objects

        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("? System对象未正确加载")
            return False

        db = sap_model.DatabaseTables

        filter_by_names = component_names is not None and len(component_names) > 0
        if not filter_by_names:
            print("?? 当前不按构件名称过滤，将导出整张表。")

        # 统一所有输出文件到 design_data 子目录
        output_dir = _ensure_design_output_dir()
        output_file = os.path.join(output_dir, output_filename)
        print("?? 尝试CSV导出方法...")
        csv_success, ret_csv, file_size = export_table_to_csv(
            db, System, table_key, output_file, table_version=1
        )

        if not csv_success:
            print(f"? CSV导出失败，返回码: {ret_csv}")
            return False

        if file_size < 10:
            print("?? CSV文件大小异常，可能未包含有效数据。")
            return False

        filtered_file = output_file.replace(".csv", "_filtered.csv")
        keep_row = build_component_filter(component_names)

        try:
            total_count, written_count = filter_csv_rows(output_file, filtered_file, keep_row)
        except Exception as e:
            print(f"? CSV过滤失败: {e}")
            return False

        log_table_status(
            table_key,
            written_count if filter_by_names else total_count,
            output_file,
            filtered_file,
            file_size,
        )

        if filter_by_names and written_count == 0:
            print("?? 未匹配到指定构件，过滤结果为空。")

        return True

    except Exception as e:
        print(f"? CSV导出过程中发生错误: {e}")
        traceback.print_exc()
        return False

# =============================================================================
# 备用：梁设计内力提取（通过 Array 方式）
# =============================================================================
def extract_beam_design_forces(sap_model, beam_names):
    """
    提取框架梁设计内力（备用方法）
    """
    try:
        from common.etabs_api_loader import get_api_objects

        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载，无法提取梁设计内力")
            return False

        output_file = os.path.join(DESIGN_DATA_DIR, "beam_flexure_envelope.csv")

        possible_table_keys = [
            "Concrete Beam Flexure Envelope - Chinese 2010",
            "Design Forces - Beams",
            "Concrete Beam Design - Flexural & Shear Forces",
            "Beam Design Forces",
        ]

        db = sap_model.DatabaseTables
        table_key = None
        final_result = None

        for key in possible_table_keys:
            try:
                print(f"🔍 尝试访问表格: {key}")

                field_key_list = System.Array.CreateInstance(System.String, 1)
                field_key_list[0] = ""

                group_name = ""
                table_version = System.Int32(0)
                fields_keys_included = System.Array.CreateInstance(System.String, 0)
                number_records = System.Int32(0)
                table_data = System.Array.CreateInstance(System.String, 0)

                test_result = db.GetTableForDisplayArray(
                    key,
                    field_key_list,
                    group_name,
                    table_version,
                    fields_keys_included,
                    number_records,
                    table_data,
                )

                success = False
                if isinstance(test_result, tuple):
                    if test_result[0] == 0:
                        success = True
                        final_result = test_result
                elif test_result == 0:
                    success = True

                if success:
                    table_key = key
                    print(f"✅ 成功访问表格: {key}")
                    break

            except Exception as e:
                print(f"⚠️ 测试表格 {key} 时出错: {e}")
                continue

        if table_key is None or final_result is None:
            print("❌ 无法找到任何可用的框架梁设计内力表格")
            return False

        try:
            if isinstance(final_result, tuple):
                fields_keys_included = final_result[3] if len(final_result) > 3 else None
                number_records = final_result[4] if len(final_result) > 4 else None
                table_data = final_result[5] if len(final_result) > 5 else None

                if hasattr(fields_keys_included, "__len__") and hasattr(
                    fields_keys_included, "__getitem__"
                ):
                    field_keys_list = [
                        str(fields_keys_included[i])
                        for i in range(len(fields_keys_included))
                    ]
                else:
                    field_keys_list = []

                if isinstance(number_records, (int, float)):
                    num_records = int(number_records)
                else:
                    num_records = 0

                if hasattr(table_data, "__len__") and hasattr(
                    table_data, "__getitem__"
                ):
                    table_data_list = [
                        str(table_data[i]) for i in range(len(table_data))
                    ]
                else:
                    table_data_list = []
            else:
                return False

            if num_records == 0:
                print(f"⚠️ 表格 '{table_key}' 中没有数据记录")
                return False

            print(f"📋 成功获取 {num_records} 条记录")

            with open(output_file, "w", newline="", encoding="utf-8-sig") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(field_keys_list)

                num_fields = len(field_keys_list)
                if num_fields > 0:
                    data_rows = [
                        table_data_list[i : i + num_fields]
                        for i in range(0, len(table_data_list), num_fields)
                    ]
                else:
                    data_rows = []

                unique_name_index = None
                for i, field in enumerate(field_keys_list):
                    if "unique" in field.lower() and "name" in field.lower():
                        unique_name_index = i
                        break

                written_count = 0
                if unique_name_index is None:
                    for row in data_rows:
                        writer.writerow(row)
                    written_count = len(data_rows)
                else:
                    for row in data_rows:
                        if (
                            len(row) > unique_name_index
                            and row[unique_name_index] in beam_names
                        ):
                            writer.writerow(row)
                            written_count += 1

                print(f"✅ 成功保存 {written_count} 条框架梁设计数据")
                print(f"📄 文件已保存至: {output_file}")

            return written_count > 0

        except Exception as e:
            print(f"❌ 解析API结果时出错: {e}")
            traceback.print_exc()
            return False

    except Exception as e:
        print(f"❌ 提取框架梁设计数据失败: {e}")
        traceback.print_exc()
        return False


# =============================================================================
# 汇总报告生成
# =============================================================================
def generate_summary_report(column_names, beam_names):
    """
    生成设计内力提取的汇总报告
    """
    try:
        output_dir = _ensure_design_output_dir()
        output_file = os.path.join(output_dir, "design_forces_summary_report.txt")

        column_csv = os.path.join(output_dir, "column_design_forces.csv")
        column_pmm_raw_csv = os.path.join(
            output_dir, "column_pmm_design_forces_raw.csv"
        )
        column_pmm_csv = os.path.join(output_dir, "column_pmm_design_summary.csv")
        beam_envelope_csv = os.path.join(output_dir, "beam_flexure_envelope.csv")
        beam_forces_csv = os.path.join(output_dir, "beam_design_forces.csv")
        beam_shear_csv = os.path.join(output_dir, "beam_shear_envelope.csv")
        column_shear_csv = os.path.join(output_dir, "column_shear_envelope.csv")
        joint_envelope_csv = os.path.join(output_dir, "joint_envelope.csv")

        column_records = 0
        column_pmm_raw_records = 0
        column_pmm_records = 0
        beam_records = 0
        beam_shear_records = 0
        column_shear_records = 0
        joint_records = 0
        beam_file_used = "N/A"
        is_envelope_data = False

        if os.path.exists(column_csv):
            with open(column_csv, "r", encoding="utf-8-sig") as f:
                column_records = max(sum(1 for _ in f) - 1, 0)

        if os.path.exists(column_pmm_raw_csv):
            with open(column_pmm_raw_csv, "r", encoding="utf-8-sig") as f:
                column_pmm_raw_records = max(sum(1 for _ in f) - 1, 0)

        if os.path.exists(column_pmm_csv):
            with open(column_pmm_csv, "r", encoding="utf-8-sig") as f:
                column_pmm_records = max(sum(1 for _ in f) - 1, 0)

        if os.path.exists(beam_envelope_csv):
            with open(beam_envelope_csv, "r", encoding="utf-8-sig") as f:
                beam_records = max(sum(1 for _ in f) - 1, 0)
                beam_file_used = "beam_flexure_envelope.csv"
                is_envelope_data = True
        elif os.path.exists(beam_forces_csv):
            with open(beam_forces_csv, "r", encoding="utf-8-sig") as f:
                beam_records = max(sum(1 for _ in f) - 1, 0)
                beam_file_used = "beam_design_forces.csv"
                is_envelope_data = False

        if os.path.exists(beam_shear_csv):
            with open(beam_shear_csv, "r", encoding="utf-8-sig") as f:
                beam_shear_records = max(sum(1 for _ in f) - 1, 0)

        if os.path.exists(column_shear_csv):
            with open(column_shear_csv, "r", encoding="utf-8-sig") as f:
                column_shear_records = max(sum(1 for _ in f) - 1, 0)

        if os.path.exists(joint_envelope_csv):
            with open(joint_envelope_csv, "r", encoding="utf-8-sig") as f:
                joint_records = max(sum(1 for _ in f) - 1, 0)

        with open(output_file, "w", encoding="utf-8") as f:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write("=" * 80 + "\n")
            f.write("构件设计结果提取汇总报告\n")
            f.write(f"报告生成时间: {now}\n")
            f.write("=" * 80 + "\n\n")

            f.write("📄 提取文件列表\n")
            f.write("-" * 40 + "\n")
            f.write(
                "1. column_design_forces.csv             - 框架柱设计内力详细数据\n"
            )
            f.write(
                "2. column_pmm_design_forces_raw.csv     - 柱 P-M-M 设计内力原始表（Concrete Column PMM Envelope - Chinese 2010）\n"
            )
            f.write(
                "3. column_pmm_design_summary.csv        - 柱 P-M-M 设计汇总结果（GetSummaryResultsColumn）\n"
            )
            f.write(
                f"4. {beam_file_used} - 框架梁设计结果详细数据（弯矩 / 综合设计内力）\n"
            )
            f.write(
                "5. beam_shear_envelope.csv              - 混凝土梁剪力包络 (若成功提取)\n"
            )
            f.write(
                "6. column_shear_envelope.csv            - 混凝土柱剪力包络 (若成功提取)\n"
            )
            f.write(
                "7. joint_envelope.csv                   - 混凝土节点包络 (若成功提取)\n"
            )
            f.write(
                "8. design_forces_summary_report.txt     - 本汇总报告\n"
            )
            f.write("\n")

            f.write("📊 提取构件范围与结果\n")
            f.write("-" * 40 + "\n")
            f.write(f"请求提取的框架柱数量: {len(column_names)}\n")
            f.write(f"实际提取的框架柱记录数: {column_records}\n")
            f.write(f"柱 P-M-M 原始记录数: {column_pmm_raw_records}\n")
            f.write(f"柱 P-M-M 设计汇总记录数: {column_pmm_records}\n")
            f.write(f"梁剪力包络记录数: {beam_shear_records}\n")
            f.write(f"柱剪力包络记录数: {column_shear_records}\n")
            f.write(f"请求提取的框架梁数量: {len(beam_names)}\n")
            f.write(f"实际提取的框架梁记录数: {beam_records}\n")
            f.write(f"节点包络记录数: {joint_records}\n\n")

            f.write("📋 数据字段说明 (根据提取的表格)\n")
            f.write("-" * 40 + "\n")
            if is_envelope_data:
                f.write(
                    "梁数据来自 'Concrete Beam Flexure Envelope - Chinese 2010' 表格，典型字段包括:\n"
                )
                f.write(
                    "-ve Moment / +ve Moment   - 端截面负/正弯矩包络 (kN·m)\n"
                )
                f.write(
                    "As Top / As Bottom        - 顶/底部配筋面积 (mm^2)\n"
                )
                f.write("Section / Location        - 截面号与位置\n")
            else:
                f.write("梁数据来自 'Design Forces - Beams' 表格:\n")
                f.write("P    - 轴力 (kN)\n")
                f.write("V2   - 局部2方向剪力 (kN)\n")
                f.write("V3   - 局部3方向剪力 (kN)\n")
                f.write("T    - 扭矩 (kN·m)\n")
                f.write("M2   - 局部2轴弯矩 (kN·m)\n")
                f.write("M3   - 局部3轴弯矩 (kN·m)\n")

            f.write(
                "\n梁剪力包络表（beam_shear_envelope.csv）通常来自 "
                "'Concrete Beam Shear Envelope - Chinese 2010' 表，"
                "提供在控制组合下的剪力包络值及对应荷载组合名称，同样保持 ETABS 原始单位。\n"
            )

            f.write("\n柱数据字段通常包括 P, V2, V3, M2, M3 等；\n")
            f.write(
                "柱 P-M-M 原始表（column_pmm_design_forces_raw.csv）直接对应 "
                "'Concrete Column PMM Envelope - Chinese 2010' 或兼容表，"
                "包括 Story, Label, UniqueName, Section, Location, "
                "P, M Major, M Minor, At Major, At Minor, PMM Combo, PMM Ratio 或配筋率, Status 等字段。\n"
            )
            f.write(
                "其中 At Major / At Minor 等配筋面积类字段，本脚本一律按 ETABS 原始数值写入，"
                "不做任何单位转换，保证与图形界面显示一致。\n"
            )
            f.write(
                "柱 P-M-M 设计汇总文件（column_pmm_design_summary.csv）给出按中国规范组合后的控制弯矩 / 轴力设计结果，"
                "包括 PMM 组合名、配筋面积或应力比、剪力控制组合及箍筋面积等。\n"
            )
            f.write("柱剪力包络表通常提供各楼层柱在控制组合下的剪力包络及相关组合信息。\n")
            f.write("节点包络表通常提供节点弯矩、剪力或 D/C 比等控制指标的包络值。\n\n")

            f.write("⚠️ 重要说明\n")
            f.write("-" * 40 + "\n")
            f.write("1. 本脚本提取的是设计结果或设计内力，请注意区分。\n")
            f.write("2. 包络(Envelope)数据通常包含最终配筋或控制内力，更具参考价值。\n")
            f.write("3. P-M-M 汇总结果直接来源于 ETABS 的 DesignConcrete.GetSummaryResultsColumn 或相应设计表。\n")
            f.write("4. 所有面积类字段（如 As、At、Av、PMMArea 等）均保持 ETABS 原始单位，不做单位换算。\n")
            f.write("5. 请结合 ETABS 设计结果和相关规范，对数据进行核对与使用。\n")
            f.write("6. 建议进行人工复核重要构件和关键节点的设计结果。\n")
            f.write("7. 本报告仅供参考，最终设计以正式图纸及审图意见为准。\n")
            f.write("8. 如果提取记录数为 0，请检查构件设计是否完成且目标表格存在。\n")
            f.write("\n")

            f.write("=" * 80 + "\n")
            f.write("报告生成完成\n")
            f.write("=" * 80 + "\n")

        print(f"✅ 设计结果汇总报告已保存至: {output_file}")
        return True

    except Exception as e:
        print(f"❌ 生成设计内力汇总报告失败: {e}")
        traceback.print_exc()
        return False


def print_extraction_summary():
    """在控制台打印提取结果汇总（简版）"""
    print("\n" + "=" * 60)
    print("📋 构件设计结果提取完成汇总")
    print("=" * 60)
    print("✅ 已生成的文件(若对应步骤成功):")
    print("   1. column_design_forces.csv                  - 框架柱设计内力/结果")
    print("   2. column_pmm_design_forces_raw.csv          - 柱 P-M-M 设计内力原始表 (Concrete Column PMM Envelope)")
    print("   3. column_pmm_design_summary.csv             - 柱 P-M-M 设计内力汇总")
    print("   4. beam_flexure_envelope.csv (或 beam_design_forces.csv) - 框架梁弯矩/设计结果")
    print("   5. beam_shear_envelope.csv                   - 梁剪力包络 (Concrete Beam Shear Envelope)")
    print("   6. column_shear_envelope.csv                 - 柱剪力包络 (Concrete Column Shear Envelope)")
    print("   7. joint_envelope.csv                        - 节点包络 (Concrete Joint Envelope)")
    print("   8. design_forces_summary_report.txt          - 提取任务汇总报告")
    print()
    print("📊 内容包括:")
    print("   • 各构件在不同荷载组合下的设计内力或包络值")
    print("   • 可能包括轴力(P)、剪力(V)、弯矩(M)、扭矩(T)、配筋面积(As/At/Av)、P-M-M 配筋面积/应力比、D/C 比等")
    print("   • 构件位置信息(Story, Station/Location)")
    print("   • 荷载组合名称(Combo / OutputCase / PMMCombo / VMajorCombo / VMinorCombo)")
    print("=" * 60)


# =============================================================================
# 若干调试函数
# =============================================================================
def test_simple_api_call(sap_model, table_key):
    """
    简单的API调用测试，用于验证数据结构
    """
    try:
        print(f"🧪 测试简单API调用 - 表格: {table_key}")

        from common.etabs_api_loader import get_api_objects

        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return None

        db = sap_model.DatabaseTables

        try:
            field_key_list = System.Array.CreateInstance(System.String, 3)
            field_key_list[0] = "Story"
            field_key_list[1] = (
                "Column"
                if "Column" in table_key
                else "Beam"
                if "Beam" in table_key
                else "Label"
            )
            field_key_list[2] = "UniqueName"
        except Exception:
            field_key_list = System.Array.CreateInstance(System.String, 1)
            field_key_list[0] = ""

        group_name = ""
        table_version = System.Int32(0)
        fields_keys_included = System.Array.CreateInstance(System.String, 0)
        number_records = System.Int32(0)
        table_data = System.Array.CreateInstance(System.String, 0)

        ret = db.GetTableForDisplayArray(
            table_key,
            field_key_list,
            group_name,
            table_version,
            fields_keys_included,
            number_records,
            table_data,
        )

        print(f"🔍 简单调用返回: {ret}")

        if isinstance(ret, tuple) and len(ret) >= 6:
            error_code = ret[0]
            if error_code == 0:
                fields_included = ret[3]
                num_records = ret[4]
                data_array = ret[5]

                print("✅ 成功调用，解析结果:")
                print(f"   记录数: {num_records}")

                if hasattr(fields_included, "__len__"):
                    field_list = [
                        str(fields_included[i]) for i in range(len(fields_included))
                    ]
                    print(f"   字段列表: {field_list}")

                if hasattr(data_array, "__len__") and len(data_array) > 0:
                    sample_size = min(15, len(data_array))
                    sample_data = [str(data_array[i]) for i in range(sample_size)]
                    print(f"   数据样本: {sample_data}")

                return ret
            else:
                print(f"❌ API调用失败，错误码: {error_code}")
                return None
        else:
            print(f"❌ 返回结构异常: {ret}")
            return None

    except Exception as e:
        print(f"❌ 简单API测试失败: {e}")
        return None


def debug_api_return_structure(sap_model, table_key):
    """
    调试函数：分析API返回的数据结构
    """
    try:
        print(f"🔍 调试API返回结构 - 表格: {table_key}")

        from common.etabs_api_loader import get_api_objects

        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return

        db = sap_model.DatabaseTables

        field_key_list = System.Array.CreateInstance(System.String, 1)
        field_key_list[0] = ""

        group_name = ""
        table_version = System.Int32(0)
        fields_keys_included = System.Array.CreateInstance(System.String, 0)
        number_records = System.Int32(0)
        table_data = System.Array.CreateInstance(System.String, 0)

        ret = db.GetTableForDisplayArray(
            table_key,
            field_key_list,
            group_name,
            table_version,
            fields_keys_included,
            number_records,
            table_data,
        )

        print(f"📊 API返回值类型: {type(ret)}")
        print(f"📊 API返回值: {ret}")

        if isinstance(ret, tuple):
            print(f"📊 元组长度: {len(ret)}")
            for i, item in enumerate(ret):
                print(f"   [{i}] 类型: {type(item)}, 值: {item}")
                if hasattr(item, "__len__") and not isinstance(
                    item, (str, int, float)
                ):
                    try:
                        print(f"       长度: {len(item)}")
                        if 0 < len(item) < 20:
                            print(
                                f"       内容: {[str(item[j]) for j in range(min(5, len(item)))]}"
                            )
                    except Exception:
                        pass
    except Exception as e:
        print(f"❌ 调试API结构时出错: {e}")
        traceback.print_exc()


def debug_available_tables(sap_model):
    """
    调试函数：列出部分常见可用的数据库表格
    """
    try:
        print("🔍 调试：列出常见可用的数据库表格...")

        from common.etabs_api_loader import get_api_objects

        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return

        db = sap_model.DatabaseTables

        common_tables = [
            "Analysis Results",
            "Design Results",
            "Element Forces - Frames",
            "Modal Information",
            "Story Drifts",
            "Joint Reactions",
            "Design Forces - Beams",
            "Design Forces - Columns",
            "Concrete Column PMM Envelope - Chinese 2010",
            "Concrete Column PMM - Chinese 2010",
            "Concrete Column Design - P-M-M Design Forces",
            "Concrete Beam Design - Flexural & Shear Forces",
            "Concrete Beam Flexure Envelope - Chinese 2010",
            "Concrete Beam Shear Envelope - Chinese 2010",
            "Concrete Column Shear Envelope - Chinese 2010",
            "Concrete Joint Envelope - Chinese 2010",
            "Concrete Column Envelope - Chinese 2010",
        ]

        available_tables = []

        for table in common_tables:
            try:
                field_key_list = System.Array.CreateInstance(System.String, 1)
                field_key_list[0] = ""

                group_name = ""
                table_version = System.Int32(0)
                fields_keys_included = System.Array.CreateInstance(System.String, 0)
                number_records = System.Int32(0)
                table_data = System.Array.CreateInstance(System.String, 0)

                ret = db.GetTableForDisplayArray(
                    table,
                    field_key_list,
                    group_name,
                    table_version,
                    fields_keys_included,
                    number_records,
                    table_data,
                )

                if (isinstance(ret, tuple) and ret[0] == 0) or ret == 0:
                    available_tables.append(table)

            except Exception:
                continue

        print(f"✅ 找到 {len(available_tables)} 个可用表格(在预设列表中):")
        for table in available_tables:
            print(f"   • {table}")

        if not available_tables:
            print("❌ 预设列表中的表格均不可用")

        return available_tables

    except Exception as e:
        print(f"❌ 调试表格列表时出错: {e}")
        return []


def debug_pmm_tables(sap_model):
    """
    调试函数：列出所有名字里包含 'Concrete Column PMM' 的数据库表格，
    用来确认正确的 TableKey（不同版本/语言的 ETABS 表名可能略有差异）。
    """
    try:
        print("🔍 调试：搜索包含 'Concrete Column PMM' 的表格...")

        from common.etabs_api_loader import get_api_objects

        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return

        db = sap_model.DatabaseTables

        NumberTables = System.Int32(0)
        table_keys = System.Array.CreateInstance(System.String, 0)
        table_names = System.Array.CreateInstance(System.String, 0)
        import_type = System.Array.CreateInstance(System.Int32, 0)
        is_empty = System.Array.CreateInstance(System.Boolean, 0)

        ret = db.GetAllTables(
            NumberTables,
            table_keys,
            table_names,
            import_type,
            is_empty,
        )

        if isinstance(ret, tuple):
            err = ret[0]
            if err != 0:
                print(f"❌ GetAllTables 调用失败，错误码: {err}")
                return
            NumberTables = int(ret[1])
            table_keys = ret[2]
            table_names = ret[3]
            import_type = ret[4]
            is_empty = ret[5]
        else:
            if ret != 0:
                print(f"❌ GetAllTables 调用失败，错误码: {ret}")
                return
            NumberTables = int(NumberTables)

        matches = []
        for i in range(NumberTables):
            try:
                key = str(table_keys[i])
                name = str(table_names[i])
                if "Concrete Column PMM" in key:
                    empty_flag = False
                    if hasattr(is_empty, "__len__") and len(is_empty) > i:
                        empty_flag = bool(is_empty[i])
                    matches.append((key, name, empty_flag))
            except Exception:
                continue

        if not matches:
            print("⚠️ 没有找到包含 'Concrete Column PMM' 的表格。")
            return

        print(f"✅ 找到 {len(matches)} 个相关表格:")
        for key, name, empty_flag in matches:
            empty_str = "空表" if empty_flag else "有数据"
            print(f"   • {key}  |  {name}  |  {empty_str}")

    except Exception as e:
        print(f"❌ 调试 PMM 表格列表时出错: {e}")
        traceback.print_exc()


# =============================================================================
# 备用：基本分析内力提取
# =============================================================================
def extract_basic_frame_forces(sap_model, column_names, beam_names):
    """
    备用方法：提取基本的构件分析内力（非设计内力）
    """
    try:
        print("?? 尝试提取基本构件分析内力...")

        from common.etabs_api_loader import get_api_objects

        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("? System对象未正确加载")
            return False

        db = sap_model.DatabaseTables
        table_key = "Element Forces - Frames"
        print(f"?? 尝试访问表格: {table_key}")

        field_key_list = System.Array.CreateInstance(System.String, 1)
        field_key_list[0] = ""

        group_name = ""
        table_version = System.Int32(0)
        fields_keys_included = System.Array.CreateInstance(System.String, 0)
        number_records = System.Int32(0)
        table_data = System.Array.CreateInstance(System.String, 0)

        ret = db.GetTableForDisplayArray(
            table_key,
            field_key_list,
            group_name,
            table_version,
            fields_keys_included,
            number_records,
            table_data,
        )

        success = (isinstance(ret, tuple) and ret[0] == 0) or (ret == 0)

        if not success:
            print("? 无法访问基本内力表格")
            return False

        if isinstance(ret, tuple) and len(ret) >= 6:
            fields_keys_included = ret[3]
            number_records = ret[4]
            table_data = ret[5]

            field_keys_list = (
                [str(field) for field in fields_keys_included]
                if fields_keys_included
                else []
            )
            num_records = (
                int(number_records) if hasattr(number_records, "__int__") else 0
            )

            if hasattr(table_data, "__len__") and hasattr(table_data, "__getitem__"):
                table_data_list = [
                    str(table_data[i]) for i in range(len(table_data))
                ]
            else:
                table_data_list = []

            if num_records == 0:
                print("? 基本内力表格中没有数据")
                return False

        output_dir = _ensure_design_output_dir()
        output_file = os.path.join(output_dir, "basic_frame_forces.csv")
        with open(output_file, "w", newline="", encoding="utf-8-sig") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(field_keys_list)
            num_fields = len(field_keys_list)
            if num_fields > 0:
                data_rows = [
                    table_data_list[i: i + num_fields]
                    for i in range(0, len(table_data_list), num_fields)
                ]
                for row in data_rows:
                    writer.writerow(row)
        print(f"✅ 基本构件内力数据已保存至: {output_file}")
        return True
    except Exception as e:
        print(f"? 提取基本构件内力失败: {e}")
        traceback.print_exc()
        return False

# =============================================================================
# 导出符号清单（供外部兼容导入）
# =============================================================================
__all__ = [
    "check_design_completion",
    "debug_api_return_structure",
    "debug_available_tables",
    "debug_pmm_tables",
    "extract_basic_frame_forces",
    "extract_beam_design_forces",
    "extract_column_design_forces",
    "extract_column_pmm_design_forces",
    "extract_design_forces_and_summary",
    "extract_design_forces_simple",
    "generate_summary_report",
    "print_extraction_summary",
    "test_simple_api_call",
]

# =============================================================================
# 脚本独立运行调试入口
# =============================================================================
if __name__ == "__main__":
    print("此模块是ETABS自动化项目的一部分，应在主程序 main.py 中调用。")
    print("直接运行此文件不会执行任何ETABS操作。")
    print("请运行 main.py 来执行完整的建模和设计流程。")
    print("\n如果需要单独测试此模块，请确保:")
    print("1. ETABS已打开并加载了完成设计的模型")
    print("2. 已运行 setup_etabs() 初始化连接")
    print("3. 已完成混凝土构件设计计算")

    try:
        from common.etabs_setup import get_sap_model, ensure_etabs_ready

        if ensure_etabs_ready():
            sap_model = get_sap_model()
            if sap_model:
                print("\n🔍 调试模式：列出常见可用表格...")
                debug_available_tables(sap_model)

                print("\n🔍 调试模式：搜索 Concrete Column PMM 相关表格...")
                debug_pmm_tables(sap_model)
    except Exception:
        print("\n⚠️ 无法连接到ETABS进行调试")

from pathlib import Path


def _ensure_design_output_dir() -> str:
    """确保设计结果输出目录存在，返回绝对路径。"""
    Path(DESIGN_DATA_DIR).mkdir(parents=True, exist_ok=True)
    return DESIGN_DATA_DIR
