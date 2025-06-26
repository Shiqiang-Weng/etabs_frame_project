# design_force_extraction_fixed.py
"""
构件设计内力提取模块 - 修复版
用于提取混凝土构件设计后的控制内力、配筋信息和设计总结
修复了GetTableForDisplayArray方法参数处理问题
"""

import os
import csv
import traceback
import sys
from datetime import datetime
from config import *
from etabs_setup import get_sap_model, ensure_etabs_ready
from utility_functions import check_ret, arr


def extract_design_forces_and_summary(column_names, beam_names):
    """
    提取构件设计内力的主函数

    Args:
        column_names (list): 框架柱名称列表
        beam_names (list): 框架梁名称列表

    Returns:
        bool: 提取是否成功
    """
    print("=" * 60)
    print("🔬 开始构件设计内力提取")
    print("=" * 60)

    try:
        # 确保ETABS连接正常
        if not ensure_etabs_ready():
            print("❌ 无法建立ETABS连接，请确保ETABS已打开并已加载模型。")
            return False

        # 获取SAP模型对象
        sap_model = get_sap_model()
        if sap_model is None:
            print("❌ 无法获取ETABS模型对象。")
            return False

        print("✅ ETABS连接正常，模型对象获取成功")

        # 检查设计是否完成
        if not check_design_completion(sap_model):
            print("❌ 设计未完成或设计表格不可用，无法提取设计内力")
            return False

        # 添加调试信息
        print("🔍 开始API调试分析...")
        test_simple_api_call(sap_model, "Design Forces - Columns")
        test_simple_api_call(sap_model, "Design Forces - Beams")

        # 提取框架柱设计内力
        print("📊 正在提取框架柱设计内力...")
        # 先尝试简化方法
        column_design_success = extract_design_forces_simple(
            sap_model, "Design Forces - Columns", column_names, "column_design_forces.csv"
        )

        # 如果简化方法失败，尝试原方法
        if not column_design_success:
            print("🔄 简化方法失败，尝试原方法...")
            column_design_success = extract_column_design_forces(sap_model, column_names)

        # 提取框架梁设计内力
        print("📊 正在提取框架梁设计内力...")
        # 先尝试简化方法
        beam_design_success = extract_design_forces_simple(
            sap_model, "Design Forces - Beams", beam_names, "beam_design_forces.csv"
        )

        # 如果简化方法失败，尝试原方法
        if not beam_design_success:
            print("🔄 简化方法失败，尝试原方法...")
            beam_design_success = extract_beam_design_forces(sap_model, beam_names)

        # 检查CSV提取是否成功
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


def check_design_completion(sap_model):
    """
    检查设计是否已完成
    使用修复后的数据库表方式检查可用表格

    Args:
        sap_model: ETABS模型对象

    Returns:
        bool: 设计是否完成
    """
    try:
        print("🔍 正在检查设计完成状态...")

        # 动态导入API对象
        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载，无法检查设计状态")
            return False

        # 使用数据库表方式获取可用表格
        db = sap_model.DatabaseTables

        # 要检查的设计表格
        design_tables_to_check = [
            "Design Forces - Beams",
            "Design Forces - Columns",
            "Concrete Column Design - P-M-M Design Forces",
            "Concrete Beam Design - Flexural & Shear Forces"
        ]

        found_tables = []

        for table_key in design_tables_to_check:
            try:
                # 创建空的字段列表 - 这是关键修复点
                field_key_list = System.Array.CreateInstance(System.String, 1)
                field_key_list[0] = ""

                group_name = ""

                # 正确初始化输出参数
                table_version = System.Int32(0)
                fields_keys_included = System.Array.CreateInstance(System.String, 0)
                number_records = System.Int32(0)
                table_data = System.Array.CreateInstance(System.String, 0)

                # 使用正确的参数调用API - 关键是要传递引用
                ret = db.GetTableForDisplayArray(
                    table_key,
                    field_key_list,  # ref parameter
                    group_name,
                    table_version,  # ref parameter
                    fields_keys_included,  # ref parameter
                    number_records,  # ref parameter
                    table_data  # ref parameter
                )

                # 检查返回值
                if isinstance(ret, tuple):
                    # 如果返回元组，第一个元素是错误码
                    error_code = ret[0]
                    if error_code == 0:
                        found_tables.append(table_key)
                        print(f"✅ 找到设计表格: {table_key}")

                        # 如果有数据，显示记录数
                        if len(ret) > 5:
                            try:
                                record_count = ret[5] if hasattr(ret[5], '__len__') else 0
                                if hasattr(record_count, '__len__'):
                                    record_count = len(record_count)
                                print(f"   📊 包含 {record_count} 条记录")
                            except:
                                pass
                    else:
                        print(f"⚠️ 表格不可用: {table_key} (错误码: {error_code})")
                elif ret == 0:
                    found_tables.append(table_key)
                    print(f"✅ 找到设计表格: {table_key}")
                else:
                    print(f"⚠️ 表格不可用: {table_key} (返回码: {ret})")

            except Exception as e:
                print(f"⚠️ 检查表格 {table_key} 时出错: {str(e)}")
                continue

        if len(found_tables) >= 2:  # 至少要有梁和柱的设计表格
            print(f"✅ 成功找到 {len(found_tables)} 个设计表格，设计已完成")
            return True
        elif len(found_tables) > 0:
            print(f"⚠️ 只找到 {len(found_tables)} 个设计表格，可能设计未完全完成")
            return True  # 部分完成也允许继续
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


def extract_design_forces_simple(sap_model, table_key, component_names, output_filename):
    """
    简化的设计内力提取方法

    Args:
        sap_model: ETABS模型对象
        table_key: 表格键名
        component_names: 构件名称列表
        output_filename: 输出文件名

    Returns:
        bool: 提取是否成功
    """
    try:
        print(f"🔍 简化提取方法 - 表格: {table_key}")

        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return False

        db = sap_model.DatabaseTables

        # 使用CSV导出方法作为备选
        print("🔄 尝试CSV导出方法...")

        output_file = os.path.join(SCRIPT_DIRECTORY, output_filename)

        # 创建空字段列表以获取所有字段
        field_key_list = System.Array.CreateInstance(System.String, 1)
        field_key_list[0] = ""

        group_name = ""
        table_version = System.Int32(1)

        # 尝试CSV导出
        ret_csv = db.GetTableForDisplayCSVFile(
            table_key,
            field_key_list,
            group_name,
            table_version,
            output_file
        )

        print(f"🔍 CSV导出返回值: {ret_csv}")
        print(f"🔍 CSV导出返回类型: {type(ret_csv)}")

        # 检查CSV导出结果
        csv_success = False
        if isinstance(ret_csv, tuple):
            error_code = ret_csv[0]
            if error_code == 0:
                csv_success = True
        elif ret_csv == 0:
            csv_success = True

        if csv_success and os.path.exists(output_file):
            print(f"✅ CSV导出成功: {output_file}")

            # 检查文件大小
            file_size = os.path.getsize(output_file)
            print(f"📊 CSV文件大小: {file_size} 字节")

            if file_size > 0:
                print(f"✅ CSV导出成功: {output_file}")

            # 读取并过滤CSV文件
            filtered_file = output_file.replace('.csv', '_filtered.csv')

            try:
                with open(output_file, 'r', encoding='utf-8-sig') as infile:
                    with open(filtered_file, 'w', newline='', encoding='utf-8-sig') as outfile:
                        reader = csv.reader(infile)
                        writer = csv.writer(outfile)

                        headers = next(reader)
                        writer.writerow(headers)

                        # 找到构件名称列
                        name_col_index = None
                        for i, header in enumerate(headers):
                            if any(keyword in header.lower() for keyword in ['unique', 'element', 'label', 'name']):
                                if 'combo' not in header.lower():
                                    name_col_index = i
                                    break

                        written_count = 0
                        total_count = 0

                        for row in reader:
                            total_count += 1
                            if name_col_index is not None and len(row) > name_col_index:
                                if row[name_col_index] in component_names:
                                    writer.writerow(row)
                                    written_count += 1
                            elif name_col_index is None:
                                # 如果找不到名称列，保存所有数据
                                writer.writerow(row)
                                written_count += 1

                        print(f"✅ 过滤完成: {written_count}/{total_count} 条记录")
                        print(f"📄 过滤后文件: {filtered_file}")

                        return written_count > 0

            except Exception as e:
                print(f"⚠️ CSV过滤失败: {e}")
                print(f"💡 原始CSV文件仍可用: {output_file}")
                return True

        else:
            print(f"❌ CSV导出失败，返回码: {ret_csv}")
            return False

    except Exception as e:
        print(f"❌ 简化提取方法失败: {e}")
        traceback.print_exc()
        return False


def extract_column_design_forces(sap_model, column_names):
    """
    提取框架柱设计内力
    使用修复后的数据库表方式提取数据

    Args:
        sap_model: ETABS模型对象
        column_names (list): 框架柱名称列表

    Returns:
        bool: 提取是否成功
    """
    try:
        # 动态导入API对象
        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载，无法提取柱设计内力")
            return False

        output_file = os.path.join(SCRIPT_DIRECTORY, 'column_design_forces.csv')

        # 尝试多个可能的表格名称
        possible_table_keys = [
            "Design Forces - Columns",
            "Concrete Column Design - P-M-M Design Forces",
            "Column Design Forces"
        ]

        db = sap_model.DatabaseTables
        table_key = None
        successful_result = None

        for key in possible_table_keys:
            try:
                print(f"🔍 尝试访问表格: {key}")

                # 创建空的字段列表来测试表格存在性
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
                    table_data
                )

                # 检查结果
                success = False
                if isinstance(test_result, tuple):
                    error_code = test_result[0]
                    if error_code == 0:
                        success = True
                        successful_result = test_result
                elif test_result == 0:
                    success = True

                if success:
                    table_key = key
                    print(f"✅ 成功访问表格: {key}")
                    break
                else:
                    print(f"⚠️ 表格不可用: {key}")

            except Exception as e:
                print(f"⚠️ 测试表格 {key} 时出错: {e}")
                continue

        if table_key is None:
            print("❌ 无法找到任何可用的框架柱设计内力表格")
            return False

        print(f"🔍 正在从表格 '{table_key}' 提取详细数据...")

        # 定义要提取的字段 - 使用更通用的字段名
        possible_field_sets = [
            ['Story', 'Column', 'UniqueName', 'Combo', 'StationLoc', 'P', 'V2', 'V3', 'T', 'M2', 'M3'],
            ['Story', 'Column', 'Unique Name', 'Combo', 'Station Loc', 'P', 'V2', 'V3', 'T', 'M2', 'M3'],
            ['Story', 'Element', 'UniqueName', 'LoadCase', 'Station', 'P', 'V2', 'V3', 'T', 'M2', 'M3'],
            ['Story', 'Label', 'UniqueName', 'OutputCase', 'Location', 'P', 'V2', 'V3', 'T', 'M2', 'M3']
        ]

        extraction_successful = False
        final_result = None

        for field_set in possible_field_sets:
            try:
                print(f"🔄 尝试字段集: {field_set}")

                # 创建字段列表
                field_key_list = System.Array.CreateInstance(System.String, len(field_set))
                for i, field in enumerate(field_set):
                    field_key_list[i] = field

                group_name = ""
                table_version = System.Int32(0)
                fields_keys_included = System.Array.CreateInstance(System.String, 0)
                number_records = System.Int32(0)
                table_data = System.Array.CreateInstance(System.String, 0)

                # 调用API
                ret = db.GetTableForDisplayArray(
                    table_key,
                    field_key_list,
                    group_name,
                    table_version,
                    fields_keys_included,
                    number_records,
                    table_data
                )

                # 检查结果
                success = False
                if isinstance(ret, tuple):
                    error_code = ret[0]
                    if error_code == 0:
                        success = True
                        final_result = ret
                elif ret == 0:
                    success = True

                if success:
                    print(f"✅ 成功使用字段集提取数据")
                    extraction_successful = True
                    break
                else:
                    print(f"⚠️ 字段集不适用")

            except Exception as e:
                print(f"⚠️ 使用字段集 {field_set} 时出错: {e}")
                continue

        if not extraction_successful or final_result is None:
            print("❌ 无法使用任何字段集提取数据")
            return False

        # 解析结果
        try:
            print(f"🔍 调试：API返回结果类型: {type(final_result)}")
            print(f"🔍 调试：API返回结果长度: {len(final_result) if hasattr(final_result, '__len__') else 'N/A'}")

            if isinstance(final_result, tuple):
                print(f"🔍 调试：元组内容类型: {[type(item) for item in final_result]}")

                # 根据调试信息，正确的元组结构是：
                # [0] error_code (int)
                # [1] updated_field_list (System.String[])
                # [2] group_name_out (int) - 似乎是版本号
                # [3] fields_keys_included (System.String[]) - 实际的字段列表
                # [4] number_records (int) - 记录数
                # [5] table_data (System.String[]) - 表格数据

                error_code = final_result[0]
                updated_field_list = final_result[1] if len(final_result) > 1 else None
                version_out = final_result[2] if len(final_result) > 2 else None
                fields_keys_included = final_result[3] if len(final_result) > 3 else None  # 这是字段列表
                number_records = final_result[4] if len(final_result) > 4 else None  # 这是记录数
                table_data = final_result[5] if len(final_result) > 5 else None  # 这是数据

                print(f"🔍 调试：错误码: {error_code}")
                print(f"🔍 调试：fields_keys_included类型: {type(fields_keys_included)}")
                print(f"🔍 调试：number_records类型: {type(number_records)}")
                print(f"🔍 调试：table_data类型: {type(table_data)}")

                # 处理字段列表 - 应该在索引3位置
                if hasattr(fields_keys_included, '__len__') and hasattr(fields_keys_included, '__getitem__'):
                    # 如果是数组类型
                    field_keys_list = [str(fields_keys_included[i]) for i in range(len(fields_keys_included))]
                    print(f"🔍 解析出的字段列表: {field_keys_list}")
                else:
                    # 使用原始请求的字段列表
                    field_keys_list = field_set
                    print("⚠️ 使用原始字段列表，因为API未返回正确的字段信息")

                # 处理记录数 - 应该在索引4位置
                if isinstance(number_records, (int, float)):
                    num_records = int(number_records)
                    print(f"🔍 解析出的记录数: {num_records}")
                else:
                    print(f"⚠️ 无法解析记录数，类型: {type(number_records)}")
                    num_records = 0

                # 处理表格数据 - 应该在索引5位置
                if hasattr(table_data, '__len__') and hasattr(table_data, '__getitem__'):
                    table_data_list = [str(table_data[i]) for i in range(len(table_data))]
                    print(f"🔍 解析出的数据长度: {len(table_data_list)}")
                elif table_data is None:
                    table_data_list = []
                    print("⚠️ 表格数据为空")
                else:
                    print(f"⚠️ 无法解析表格数据类型: {type(table_data)}")
                    table_data_list = []

            else:
                print("❌ API返回结果不是元组格式")
                return False

            if num_records == 0:
                print(f"⚠️ 表格 '{table_key}' 中没有数据记录")
                print("💡 提示: 请确保已完成混凝土柱设计计算")
                return False

            print(f"📋 成功获取 {num_records} 条记录")
            print(f"📝 可用字段: {field_keys_list}")

            # 保存到CSV文件
            with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(field_keys_list)

                # 将一维数组转换为二维数组
                num_fields = len(field_keys_list)
                if num_fields > 0:
                    data_rows = [table_data_list[i:i + num_fields] for i in
                                 range(0, len(table_data_list), num_fields)]
                else:
                    data_rows = []

                # 寻找构件名称字段
                unique_name_index = None
                for i, field in enumerate(field_keys_list):
                    field_lower = field.lower()
                    if ('unique' in field_lower and 'name' in field_lower) or \
                            ('element' in field_lower) or \
                            ('label' in field_lower):
                        unique_name_index = i
                        break

                if unique_name_index is None:
                    print("⚠️ 无法确定构件名称字段，保存所有数据")
                    # 如果找不到名称字段，保存所有数据
                    for row in data_rows:
                        writer.writerow(row)
                    written_count = len(data_rows)
                else:
                    # 筛选指定构件的数据
                    written_count = 0
                    if data_rows:
                        unique_names_sample = list(set([row[unique_name_index] for row in data_rows[:10]]))
                        print(f"📋 数据中构件名称示例: {unique_names_sample[:5]}")

                    for row in data_rows:
                        if len(row) > unique_name_index and row[unique_name_index] in column_names:
                            writer.writerow(row)
                            written_count += 1

                print(f"✅ 成功保存 {written_count} 条框架柱设计内力数据")
                print(f"📄 文件已保存至: {output_file}")

            return written_count > 0

        except Exception as e:
            print(f"❌ 解析API结果时出错: {e}")
            traceback.print_exc()
            return False

    except Exception as e:
        print(f"❌ 提取框架柱设计内力失败: {e}")
        traceback.print_exc()
        return False


def extract_beam_design_forces(sap_model, beam_names):
    """
    提取框架梁设计内力
    使用修复后的数据库表方式提取数据

    Args:
        sap_model: ETABS模型对象
        beam_names (list): 框架梁名称列表

    Returns:
        bool: 提取是否成功
    """
    try:
        # 动态导入API对象
        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载，无法提取梁设计内力")
            return False

        output_file = os.path.join(SCRIPT_DIRECTORY, 'beam_design_forces.csv')

        # 尝试多个可能的表格名称
        possible_table_keys = [
            "Design Forces - Beams",
            "Concrete Beam Design - Flexural & Shear Forces",
            "Beam Design Forces"
        ]

        db = sap_model.DatabaseTables
        table_key = None

        for key in possible_table_keys:
            try:
                print(f"🔍 尝试访问表格: {key}")

                # 创建空的字段列表来测试表格存在性
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
                    table_data
                )

                # 检查结果
                success = False
                if isinstance(test_result, tuple):
                    error_code = test_result[0]
                    if error_code == 0:
                        success = True
                elif test_result == 0:
                    success = True

                if success:
                    table_key = key
                    print(f"✅ 成功访问表格: {key}")
                    break

            except Exception as e:
                print(f"⚠️ 测试表格 {key} 时出错: {e}")
                continue

        if table_key is None:
            print("❌ 无法找到任何可用的框架梁设计内力表格")
            return False

        print(f"🔍 正在从表格 '{table_key}' 提取详细数据...")

        # 定义要提取的字段
        possible_field_sets = [
            ['Story', 'Beam', 'UniqueName', 'Combo', 'Station', 'P', 'V2', 'V3', 'T', 'M2', 'M3'],
            ['Story', 'Beam', 'Unique Name', 'Combo', 'Station Loc', 'P', 'V2', 'V3', 'T', 'M2', 'M3'],
            ['Story', 'Element', 'UniqueName', 'LoadCase', 'Location', 'P', 'V2', 'V3', 'T', 'M2', 'M3'],
            ['Story', 'Label', 'UniqueName', 'OutputCase', 'Station', 'P', 'V2', 'V3', 'T', 'M2']
        ]

        extraction_successful = False
        final_result = None

        for field_set in possible_field_sets:
            try:
                print(f"🔄 尝试字段集: {field_set}")

                # 创建字段列表
                field_key_list = System.Array.CreateInstance(System.String, len(field_set))
                for i, field in enumerate(field_set):
                    field_key_list[i] = field

                group_name = ""
                table_version = System.Int32(0)
                fields_keys_included = System.Array.CreateInstance(System.String, 0)
                number_records = System.Int32(0)
                table_data = System.Array.CreateInstance(System.String, 0)

                # 调用API
                ret = db.GetTableForDisplayArray(
                    table_key,
                    field_key_list,
                    group_name,
                    table_version,
                    fields_keys_included,
                    number_records,
                    table_data
                )

                # 检查结果
                success = False
                if isinstance(ret, tuple):
                    error_code = ret[0]
                    if error_code == 0:
                        success = True
                        final_result = ret
                elif ret == 0:
                    success = True

                if success:
                    print(f"✅ 成功使用字段集提取数据")
                    extraction_successful = True
                    break

            except Exception as e:
                print(f"⚠️ 使用字段集 {field_set} 时出错: {e}")
                continue

        if not extraction_successful or final_result is None:
            print("❌ 无法使用任何字段集提取数据")
            return False

        # 解析结果 - 使用正确的元组结构
        try:
            print(f"🔍 调试：API返回结果类型: {type(final_result)}")
            print(f"🔍 调试：API返回结果长度: {len(final_result) if hasattr(final_result, '__len__') else 'N/A'}")

            if isinstance(final_result, tuple):
                print(f"🔍 调试：元组内容类型: {[type(item) for item in final_result]}")

                # 根据调试信息，正确的元组结构是：
                # [0] error_code (int)
                # [1] updated_field_list (System.String[])
                # [2] version_out (int)
                # [3] fields_keys_included (System.String[]) - 实际的字段列表
                # [4] number_records (int) - 记录数
                # [5] table_data (System.String[]) - 表格数据

                error_code = final_result[0]
                updated_field_list = final_result[1] if len(final_result) > 1 else None
                version_out = final_result[2] if len(final_result) > 2 else None
                fields_keys_included = final_result[3] if len(final_result) > 3 else None  # 这是字段列表
                number_records = final_result[4] if len(final_result) > 4 else None  # 这是记录数
                table_data = final_result[5] if len(final_result) > 5 else None  # 这是数据

                print(f"🔍 调试：错误码: {error_code}")
                print(f"🔍 调试：fields_keys_included类型: {type(fields_keys_included)}")
                print(f"🔍 调试：number_records类型: {type(number_records)}")
                print(f"🔍 调试：table_data类型: {type(table_data)}")

                # 处理字段列表 - 应该在索引3位置
                if hasattr(fields_keys_included, '__len__') and hasattr(fields_keys_included, '__getitem__'):
                    # 如果是数组类型
                    field_keys_list = [str(fields_keys_included[i]) for i in range(len(fields_keys_included))]
                    print(f"🔍 解析出的字段列表: {field_keys_list}")
                else:
                    # 使用原始请求的字段列表
                    field_keys_list = field_set
                    print("⚠️ 使用原始字段列表，因为API未返回正确的字段信息")

                # 处理记录数 - 应该在索引4位置
                if isinstance(number_records, (int, float)):
                    num_records = int(number_records)
                    print(f"🔍 解析出的记录数: {num_records}")
                else:
                    print(f"⚠️ 无法解析记录数，类型: {type(number_records)}")
                    num_records = 0

                # 处理表格数据 - 应该在索引5位置
                if hasattr(table_data, '__len__') and hasattr(table_data, '__getitem__'):
                    table_data_list = [str(table_data[i]) for i in range(len(table_data))]
                    print(f"🔍 解析出的数据长度: {len(table_data_list)}")
                elif table_data is None:
                    table_data_list = []
                    print("⚠️ 表格数据为空")
                else:
                    print(f"⚠️ 无法解析表格数据类型: {type(table_data)}")
                    table_data_list = []

            else:
                print("❌ API返回结果不是元组格式")
                return False

            if num_records == 0:
                print(f"⚠️ 表格 '{table_key}' 中没有数据记录")
                print("💡 提示: 请确保已完成混凝土梁设计计算")
                return False

            print(f"📋 成功获取 {num_records} 条记录")

            # 保存到CSV文件
            with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(field_keys_list)

                num_fields = len(field_keys_list)
                if num_fields > 0:
                    data_rows = [table_data_list[i:i + num_fields] for i in
                                 range(0, len(table_data_list), num_fields)]
                else:
                    data_rows = []

                # 寻找构件名称字段
                unique_name_index = None
                for i, field in enumerate(field_keys_list):
                    field_lower = field.lower()
                    if ('unique' in field_lower and 'name' in field_lower) or \
                            ('element' in field_lower) or \
                            ('label' in field_lower):
                        unique_name_index = i
                        break

                written_count = 0
                if unique_name_index is None:
                    print("⚠️ 无法确定构件名称字段，保存所有数据")
                    for row in data_rows:
                        writer.writerow(row)
                    written_count = len(data_rows)
                else:
                    for row in data_rows:
                        if len(row) > unique_name_index and row[unique_name_index] in beam_names:
                            writer.writerow(row)
                            written_count += 1

                print(f"✅ 成功保存 {written_count} 条框架梁设计内力数据")
                print(f"📄 文件已保存至: {output_file}")

            return written_count > 0

        except Exception as e:
            print(f"❌ 解析API结果时出错: {e}")
            traceback.print_exc()
            return False

    except Exception as e:
        print(f"❌ 提取框架梁设计内力失败: {e}")
        traceback.print_exc()
        return False


def generate_summary_report(column_names, beam_names):
    """
    生成设计内力提取的汇总报告

    Args:
        column_names (list): 框架柱名称列表
        beam_names (list): 框架梁名称列表

    Returns:
        bool: 报告生成是否成功
    """
    try:
        output_file = os.path.join(SCRIPT_DIRECTORY, 'design_forces_summary_report.txt')

        # 检查CSV文件是否存在并统计记录数
        column_csv = os.path.join(SCRIPT_DIRECTORY, 'column_design_forces.csv')
        beam_csv = os.path.join(SCRIPT_DIRECTORY, 'beam_design_forces.csv')

        column_records = 0
        beam_records = 0

        if os.path.exists(column_csv):
            with open(column_csv, 'r', encoding='utf-8-sig') as f:
                column_records = sum(1 for line in f) - 1  # 减去表头

        if os.path.exists(beam_csv):
            with open(beam_csv, 'r', encoding='utf-8-sig') as f:
                beam_records = sum(1 for line in f) - 1  # 减去表头

        with open(output_file, 'w', encoding='utf-8') as f:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write("=" * 80 + "\n")
            f.write("构件设计内力提取汇总报告\n")
            f.write(f"报告生成时间: {now}\n")
            f.write("=" * 80 + "\n\n")

            f.write("📄 提取文件列表\n")
            f.write("-" * 40 + "\n")
            f.write("1. column_design_forces.csv - 框架柱设计内力详细数据\n")
            f.write("2. beam_design_forces.csv - 框架梁设计内力详细数据\n")
            f.write("3. design_forces_summary_report.txt - 本汇总报告\n")
            f.write("\n")

            f.write(f"📊 提取构件范围与结果\n")
            f.write("-" * 40 + "\n")
            f.write(f"请求提取的框架柱数量: {len(column_names)}\n")
            f.write(f"实际提取的框架柱记录数: {column_records}\n")
            f.write(f"请求提取的框架梁数量: {len(beam_names)}\n")
            f.write(f"实际提取的框架梁记录数: {beam_records}\n\n")

            f.write("📋 数据字段说明\n")
            f.write("-" * 40 + "\n")
            f.write("P    - 轴力 (kN)\n")
            f.write("V2   - Y方向剪力 (kN)\n")
            f.write("V3   - Z方向剪力 (kN)\n")
            f.write("T    - 扭矩 (kN·m)\n")
            f.write("M2   - Y轴弯矩 (kN·m)\n")
            f.write("M3   - Z轴弯矩 (kN·m)\n")
            f.write("Combo - 荷载组合名称\n")
            f.write("Station/Location - 构件位置坐标\n\n")

            f.write("⚠️ 重要说明\n")
            f.write("-" * 40 + "\n")
            f.write("1. 设计内力为各荷载组合下的包络设计内力值。\n")
            f.write("2. 本脚本提取的是设计内力，而非分析内力。\n")
            f.write("3. 请结合ETABS设计结果和相关规范，对数据进行核对与使用。\n")
            f.write("4. 建议进行人工复核重要构件的设计结果。\n")
            f.write("5. 本报告仅供参考，最终设计以正式图纸为准。\n")
            f.write("6. 如果提取记录数为0，请检查构件设计是否完成。\n")
            f.write("\n")

            f.write("=" * 80 + "\n")
            f.write("报告生成完成\n")
            f.write("=" * 80 + "\n")

        print(f"✅ 设计内力汇总报告已保存至: {output_file}")
        return True

    except Exception as e:
        print(f"❌ 生成设计内力汇总报告失败: {e}")
        traceback.print_exc()
        return False


def print_extraction_summary():
    """打印提取结果汇总"""
    print("\n" + "=" * 60)
    print("📋 构件设计内力提取完成汇总")
    print("=" * 60)
    print("✅ 已生成的文件:")
    print("   1. column_design_forces.csv - 框架柱设计内力")
    print("   2. beam_design_forces.csv - 框架梁设计内力")
    print("   3. design_forces_summary_report.txt - 提取任务汇总报告")
    print()
    print("📊 内容包括:")
    print("   • 各构件在不同荷载组合下的设计内力值")
    print("   • 轴力(P)、剪力(V2,V3)、弯矩(M2,M3)、扭矩(T)")
    print("   • 构件位置信息(Story, Station/Location)")
    print("   • 荷载组合名称(Combo)")
    print("=" * 60)


def test_simple_api_call(sap_model, table_key):
    """
    简单的API调用测试，用于验证数据结构

    Args:
        sap_model: ETABS模型对象
        table_key: 表格键名
    """
    try:
        print(f"🧪 测试简单API调用 - 表格: {table_key}")

        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return None

        db = sap_model.DatabaseTables

        # 只请求3个简单字段
        field_key_list = System.Array.CreateInstance(System.String, 3)
        field_key_list[0] = "Story"
        field_key_list[1] = "Column" if "Column" in table_key else "Beam"
        field_key_list[2] = "UniqueName"

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
            table_data
        )

        print(f"🔍 简单调用返回: {ret}")

        if isinstance(ret, tuple) and len(ret) >= 6:
            error_code = ret[0]
            if error_code == 0:
                # 按照新理解的结构解析
                fields_included = ret[3]  # 字段列表
                num_records = ret[4]  # 记录数
                data_array = ret[5]  # 数据数组

                print(f"✅ 成功调用，解析结果:")
                print(f"   字段数组类型: {type(fields_included)}")
                print(f"   记录数类型: {type(num_records)}, 值: {num_records}")
                print(f"   数据数组类型: {type(data_array)}")

                if hasattr(fields_included, '__len__'):
                    print(f"   字段数组长度: {len(fields_included)}")
                    field_list = [str(fields_included[i]) for i in range(len(fields_included))]
                    print(f"   字段列表: {field_list}")

                if hasattr(data_array, '__len__'):
                    print(f"   数据数组长度: {len(data_array)}")
                    # 显示前几条数据
                    if len(data_array) > 0:
                        sample_size = min(15, len(data_array))  # 显示前5行数据 (3字段 x 5行 = 15)
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
        traceback.print_exc()
        return None


def debug_api_return_structure(sap_model, table_key):
    """
    调试函数：分析API返回的数据结构

    Args:
        sap_model: ETABS模型对象
        table_key: 表格键名
    """
    try:
        print(f"🔍 调试API返回结构 - 表格: {table_key}")

        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return

        db = sap_model.DatabaseTables

        # 创建简单的字段列表
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
            table_data
        )

        print(f"📊 API返回值类型: {type(ret)}")
        print(f"📊 API返回值: {ret}")

        if isinstance(ret, tuple):
            print(f"📊 元组长度: {len(ret)}")
            for i, item in enumerate(ret):
                print(f"   [{i}] 类型: {type(item)}, 值: {item}")
                if hasattr(item, '__len__') and not isinstance(item, (str, int, float)):
                    try:
                        print(f"       长度: {len(item)}")
                        if len(item) > 0 and len(item) < 20:  # 只显示小数组的内容
                            print(f"       内容: {[str(item[j]) for j in range(min(5, len(item)))]}")
                    except:
                        pass

        # 尝试使用具体字段
        print(f"\n🔍 尝试使用具体字段...")
        field_key_list2 = System.Array.CreateInstance(System.String, 3)
        field_key_list2[0] = "Story"
        field_key_list2[1] = "Column" if "Column" in table_key else "Beam"
        field_key_list2[2] = "P"

        ret2 = db.GetTableForDisplayArray(
            table_key,
            field_key_list2,
            group_name,
            table_version,
            fields_keys_included,
            number_records,
            table_data
        )

        print(f"📊 具体字段API返回值类型: {type(ret2)}")
        if isinstance(ret2, tuple):
            print(f"📊 具体字段元组长度: {len(ret2)}")
            for i, item in enumerate(ret2):
                print(f"   [{i}] 类型: {type(item)}")
                if hasattr(item, '__len__') and not isinstance(item, (str, int, float)):
                    try:
                        print(f"       长度: {len(item)}")
                    except:
                        pass

    except Exception as e:
        print(f"❌ 调试API结构时出错: {e}")
        traceback.print_exc()


def debug_available_tables(sap_model):
    """
    调试函数：列出所有可用的数据库表格
    用于排查表格名称问题

    Args:
        sap_model: ETABS模型对象
    """
    try:
        print("🔍 调试：列出所有可用的数据库表格...")

        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return

        db = sap_model.DatabaseTables

        # 尝试获取表格列表的常见方法
        common_tables = [
            "Analysis Results", "Design Results", "Element Forces - Frames",
            "Modal Information", "Story Drifts", "Joint Reactions",
            "Design Forces - Beams", "Design Forces - Columns",
            "Concrete Column Design", "Concrete Beam Design",
            "Steel Design", "Composite Beam Design"
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
                    table_data
                )

                if isinstance(ret, tuple):
                    error_code = ret[0]
                    if error_code == 0:
                        available_tables.append(table)
                elif ret == 0:
                    available_tables.append(table)

            except Exception as e:
                continue

        print(f"✅ 找到 {len(available_tables)} 个可用表格:")
        for table in available_tables:
            print(f"   • {table}")

        if not available_tables:
            print("❌ 未找到任何可用表格")
            print("💡 可能的原因:")
            print("   1. 模型未完成分析")
            print("   2. 模型未完成设计")
            print("   3. API连接问题")

        return available_tables

    except Exception as e:
        print(f"❌ 调试表格列表时出错: {e}")
        return []


def extract_basic_frame_forces(sap_model, column_names, beam_names):
    """
    备用方法：提取基本的构件分析内力（非设计内力）
    当设计表格不可用时使用

    Args:
        sap_model: ETABS模型对象
        column_names (list): 框架柱名称列表
        beam_names (list): 框架梁名称列表

    Returns:
        bool: 提取是否成功
    """
    try:
        print("🔧 尝试提取基本构件分析内力...")

        from etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return False

        db = sap_model.DatabaseTables

        # 尝试提取基本的构件内力表格
        table_key = "Element Forces - Frames"

        print(f"🔍 尝试访问表格: {table_key}")

        # 创建空字段列表来获取所有字段
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
            table_data
        )

        success = False
        if isinstance(ret, tuple):
            error_code = ret[0]
            if error_code == 0:
                success = True
        elif ret == 0:
            success = True

        if not success:
            print(f"❌ 无法访问基本内力表格")
            return False

        # 解析结果
        if isinstance(ret, tuple) and len(ret) >= 6:
            fields_keys_included = ret[4]
            number_records = ret[5]
            table_data = ret[6] if len(ret) > 6 else ret[5]

            field_keys_list = [str(field) for field in fields_keys_included] if fields_keys_included else []
            num_records = int(number_records) if hasattr(number_records, '__int__') else 0

            if hasattr(table_data, '__len__') and hasattr(table_data, '__getitem__'):
                table_data_list = [str(table_data[i]) for i in range(len(table_data))]
            else:
                table_data_list = []

        if num_records == 0:
            print("❌ 基本内力表格中没有数据")
            return False

        print(f"📋 基本内力表格包含 {num_records} 条记录")
        print(f"📝 可用字段: {field_keys_list}")

        # 保存基本内力数据
        output_file = os.path.join(SCRIPT_DIRECTORY, 'basic_frame_forces.csv')

        with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(field_keys_list)

            num_fields = len(field_keys_list)
            if num_fields > 0:
                data_rows = [table_data_list[i:i + num_fields] for i in
                             range(0, len(table_data_list), num_fields)]

                # 保存所有数据（因为我们无法准确区分设计构件）
                for row in data_rows:
                    writer.writerow(row)

        print(f"✅ 基本构件内力数据已保存至: {output_file}")
        print("💡 注意: 这是分析内力，不是设计内力")

        return True

    except Exception as e:
        print(f"❌ 提取基本构件内力失败: {e}")
        traceback.print_exc()
        return False


if __name__ == "__main__":
    # 测试代码
    print("此模块是ETABS自动化项目的一部分，应在主程序 main.py 中调用。")
    print("直接运行此文件不会执行任何ETABS操作。")
    print("请运行 main.py 来执行完整的建模和设计流程。")
    print("\n如果需要单独测试此模块，请确保:")
    print("1. ETABS已打开并加载了完成设计的模型")
    print("2. 已运行 setup_etabs() 初始化连接")
    print("3. 已完成混凝土构件设计计算")

    # 可以添加简单的调试测试
    try:
        from etabs_setup import get_sap_model, ensure_etabs_ready

        if ensure_etabs_ready():
            sap_model = get_sap_model()
            if sap_model:
                print("\n🔍 调试模式：列出可用表格...")
                debug_available_tables(sap_model)
    except:
        print("\n⚠️ 无法连接到ETABS进行调试")