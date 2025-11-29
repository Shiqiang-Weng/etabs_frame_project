# design_force_extraction_extended.py
"""
扩展的构件设计内力提取模块
用于提取混凝土构件设计后的所有设计表格数据
包括弯曲包络、剪切包络、PMM包络和节点包络等
"""

import os
import csv
import traceback
import sys
from datetime import datetime
from common.config import *  # noqa: F401,F403
from common.etabs_setup import get_sap_model, ensure_etabs_ready
from common.utility_functions import check_ret, arr


def extract_all_concrete_design_data(column_names, beam_names):
    """
    提取所有混凝土设计数据的主函数

    Args:
        column_names (list): 框架柱名称列表
        beam_names (list): 框架梁名称列表

    Returns:
        bool: 提取是否成功
    """
    print("=" * 80)
    print("🔬 开始提取所有混凝土设计数据")
    print("=" * 80)

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
        if not check_design_completion_extended(sap_model):
            print("❌ 设计未完成或设计表格不可用，无法提取设计数据")
            return False

        # 定义要提取的表格配置
        concrete_tables = {
            # 梁相关表格
            'concrete_beam_flexure_envelope': {
                'table_key': 'Concrete Beam Flexure Envelope',
                'alternative_keys': [
                    'Table: Concrete Beam Flexure Envelope - Chinese 2010',
                    'Concrete Beam Flexure Envelope - Chinese 2010',
                    'Concrete Frame Design 2 - Beam Flexure Envelope'
                ],
                'filename': 'concrete_beam_flexure_envelope.csv',
                'description': '混凝土梁弯曲包络数据',
                'component_names': beam_names
            },
            'concrete_beam_shear_envelope': {
                'table_key': 'Concrete Beam Shear Envelope',
                'alternative_keys': [
                    'Table: Concrete Beam Shear Envelope - Chinese 2010',
                    'Concrete Beam Shear Envelope - Chinese 2010',
                    'Concrete Frame Design 2 - Beam Shear Envelope'
                ],
                'filename': 'concrete_beam_shear_envelope.csv',
                'description': '混凝土梁剪切包络数据',
                'component_names': beam_names
            },

            # 柱相关表格
            'concrete_column_pmm_envelope': {
                'table_key': 'Concrete Column PMM Envelope',
                'alternative_keys': [
                    'Table: Concrete Column PMM Envelope - Chinese 2010',
                    'Concrete Column PMM Envelope - Chinese 2010',
                    'Concrete Frame Design 2 - Column PMM Envelope'
                ],
                'filename': 'concrete_column_pmm_envelope.csv',
                'description': '混凝土柱PMM包络数据',
                'component_names': column_names
            },
            'concrete_column_shear_envelope': {
                'table_key': 'Concrete Column Shear Envelope',
                'alternative_keys': [
                    'Table: Concrete Column Shear Envelope - Chinese 2010',
                    'Concrete Column Shear Envelope - Chinese 2010',
                    'Concrete Frame Design 2 - Column Shear Envelope'
                ],
                'filename': 'concrete_column_shear_envelope.csv',
                'description': '混凝土柱剪切包络数据',
                'component_names': column_names
            },

            # 节点相关表格
            'concrete_joint_envelope': {
                'table_key': 'Concrete Joint Envelope',
                'alternative_keys': [
                    'Table: Concrete Joint Envelope - Chinese 2010',
                    'Concrete Joint Envelope - Chinese 2010',
                    'Concrete Frame Design 2 - Joint Envelope'
                ],
                'filename': 'concrete_joint_envelope.csv',
                'description': '混凝土节点包络数据',
                'component_names': None  # 节点数据不需要过滤特定构件
            }
        }

        # 提取每个表格的数据
        extraction_results = {}
        successful_extractions = 0

        for table_id, table_config in concrete_tables.items():
            print(f"\n📊 正在提取 {table_config['description']}...")

            success = extract_concrete_design_table(
                sap_model,
                table_config['table_key'],
                table_config['alternative_keys'],
                table_config['filename'],
                table_config['component_names'],
                table_config['description']
            )

            extraction_results[table_id] = success
            if success:
                successful_extractions += 1

        # 生成综合汇总报告
        print(f"\n📋 正在生成综合汇总报告...")
        summary_success = generate_comprehensive_summary_report(
            column_names, beam_names, concrete_tables, extraction_results
        )

        # 输出提取结果统计
        total_tables = len(concrete_tables)
        print(f"\n{'=' * 60}")
        print(f"📊 混凝土设计数据提取完成")
        print(f"{'=' * 60}")
        print(f"✅ 成功提取: {successful_extractions}/{total_tables} 个表格")
        print(f"📄 汇总报告生成: {'成功' if summary_success else '失败'}")

        if successful_extractions == total_tables:
            print("🎉 所有表格提取成功！")
            return True
        elif successful_extractions > 0:
            print("⚠️ 部分表格提取成功，请检查失败的表格")
            return True
        else:
            print("❌ 所有表格提取失败，请检查设计状态")
            return False

    except Exception as e:
        print(f"❌ 混凝土设计数据提取过程中发生严重错误: {e}")
        traceback.print_exc()
        return False


def check_design_completion_extended(sap_model):
    """
    检查混凝土设计是否已完成（扩展版本）
    检查更多的设计相关表格

    Args:
        sap_model: ETABS模型对象

    Returns:
        bool: 设计是否完成
    """
    try:
        print("🔍 正在检查混凝土设计完成状态...")

        from common.etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载，无法检查设计状态")
            return False

        db = sap_model.DatabaseTables

        # 要检查的混凝土设计表格
        design_tables_to_check = [
            "Concrete Beam Flexure Envelope",
            "Concrete Beam Shear Envelope",
            "Concrete Column PMM Envelope",
            "Concrete Column Shear Envelope",
            "Concrete Joint Envelope",
            # 备选表格名称
            "Table: Concrete Beam Flexure Envelope - Chinese 2010",
            "Table: Concrete Beam Shear Envelope - Chinese 2010",
            "Table: Concrete Column PMM Envelope - Chinese 2010",
            "Table: Concrete Column Shear Envelope - Chinese 2010",
            "Table: Concrete Joint Envelope - Chinese 2010"
        ]

        found_tables = []

        for table_key in design_tables_to_check:
            try:
                # 创建空的字段列表
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

                # 检查返回值
                if isinstance(ret, tuple):
                    error_code = ret[0]
                    if error_code == 0:
                        found_tables.append(table_key)
                        print(f"✅ 找到设计表格: {table_key}")

                        # 显示记录数
                        if len(ret) > 4:
                            try:
                                record_count = ret[4] if hasattr(ret[4], '__int__') else 0
                                print(f"   📊 包含 {record_count} 条记录")
                            except:
                                pass
                    else:
                        print(f"⚠️ 表格不可用: {table_key} (错误码: {error_code})")
                elif ret == 0:
                    found_tables.append(table_key)
                    print(f"✅ 找到设计表格: {table_key}")

            except Exception as e:
                print(f"⚠️ 检查表格 {table_key} 时出错: {str(e)[:100]}")
                continue

        if len(found_tables) >= 3:  # 至少要有3个设计表格
            print(f"✅ 成功找到 {len(found_tables)} 个混凝土设计表格，设计已完成")
            return True
        elif len(found_tables) > 0:
            print(f"⚠️ 只找到 {len(found_tables)} 个设计表格，可能设计未完全完成")
            print("💡 建议检查以下设计状态：")
            print("   1. 混凝土梁弯曲设计是否完成")
            print("   2. 混凝土梁剪切设计是否完成")
            print("   3. 混凝土柱PMM设计是否完成")
            print("   4. 混凝土柱剪切设计是否完成")
            print("   5. 混凝土节点设计是否完成")
            return True  # 部分完成也允许继续
        else:
            print("❌ 未找到任何混凝土设计表格")
            print("💡 请确保已完成混凝土设计计算:")
            print("   1. Design → Concrete Frame Design → Start Design/Check of Structure")
            print("   2. 等待设计计算完成")
            print("   3. 检查是否有设计错误或警告")
            print("   4. 确认选择了中国规范（Chinese 2010）")
            return False

    except Exception as e:
        print(f"❌ 检查混凝土设计完成状态时发生严重错误: {e}")
        traceback.print_exc()
        return False


def extract_concrete_design_table(sap_model, table_key, alternative_keys, filename, component_names, description):
    """
    提取单个混凝土设计表格的通用函数

    Args:
        sap_model: ETABS模型对象
        table_key: 主要表格键名
        alternative_keys: 备选表格键名列表
        filename: 输出文件名
        component_names: 构件名称列表（用于过滤，None表示不过滤）
        description: 表格描述

    Returns:
        bool: 提取是否成功
    """
    try:
        from common.etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print(f"❌ System对象未正确加载，无法提取{description}")
            return False

        output_file = os.path.join(SCRIPT_DIRECTORY, filename)
        db = sap_model.DatabaseTables

        # 尝试所有可能的表格名称
        all_possible_keys = [table_key] + alternative_keys
        successful_table_key = None
        extraction_result = None

        for key in all_possible_keys:
            try:
                print(f"🔍 尝试访问表格: {key}")

                # 先测试表格是否可用
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
                        extraction_result = test_result
                elif test_result == 0:
                    success = True
                    extraction_result = test_result

                if success:
                    successful_table_key = key
                    print(f"✅ 成功访问表格: {key}")
                    break
                else:
                    print(f"⚠️ 表格不可用: {key}")

            except Exception as e:
                print(f"⚠️ 测试表格 {key} 时出错: {str(e)[:100]}")
                continue

        if successful_table_key is None:
            print(f"❌ 无法找到任何可用的表格用于提取{description}")
            return False

        # 尝试使用CSV导出方法
        print(f"🔄 尝试CSV导出方法...")

        try:
            # 创建空字段列表以获取所有字段
            field_key_list = System.Array.CreateInstance(System.String, 1)
            field_key_list[0] = ""

            group_name = ""
            table_version = System.Int32(1)

            # CSV导出
            ret_csv = db.GetTableForDisplayCSVFile(
                successful_table_key,
                field_key_list,
                group_name,
                table_version,
                output_file
            )

            csv_success = False
            if isinstance(ret_csv, tuple):
                error_code = ret_csv[0]
                if error_code == 0:
                    csv_success = True
            elif ret_csv == 0:
                csv_success = True

            if csv_success and os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"✅ CSV导出成功: {output_file} (大小: {file_size} 字节)")

                if file_size > 0:
                    # 如果需要过滤特定构件，则进行过滤
                    if component_names:
                        filtered_file = output_file.replace('.csv', '_filtered.csv')
                        filter_success = filter_csv_by_components(
                            output_file, filtered_file, component_names
                        )
                        if filter_success:
                            print(f"✅ 数据过滤完成: {filtered_file}")
                    else:
                        print(f"✅ 完整数据已保存（未过滤）: {output_file}")

                    return True
                else:
                    print(f"⚠️ CSV文件为空")
                    return False
            else:
                print(f"❌ CSV导出失败")

        except Exception as csv_error:
            print(f"⚠️ CSV导出方法失败: {csv_error}")

        # 如果CSV导出失败，尝试数组方法
        print(f"🔄 尝试数组方法...")
        return extract_table_using_array_method(
            sap_model, successful_table_key, output_file, component_names, description
        )

    except Exception as e:
        print(f"❌ 提取{description}失败: {e}")
        traceback.print_exc()
        return False


def extract_table_using_array_method(sap_model, table_key, output_file, component_names, description):
    """
    使用数组方法提取表格数据

    Args:
        sap_model: ETABS模型对象
        table_key: 表格键名
        output_file: 输出文件路径
        component_names: 构件名称列表（用于过滤）
        description: 表格描述

    Returns:
        bool: 提取是否成功
    """
    try:
        from common.etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        db = sap_model.DatabaseTables

        # 先获取所有可用字段
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

        if not isinstance(ret, tuple) or len(ret) < 6:
            print(f"❌ API返回结果格式异常")
            return False

        error_code = ret[0]
        if error_code != 0:
            print(f"❌ API调用失败，错误码: {error_code}")
            return False

        # 解析结果
        fields_keys_included = ret[3] if len(ret) > 3 else None
        number_records = ret[4] if len(ret) > 4 else None
        table_data = ret[5] if len(ret) > 5 else None

        # 处理字段列表
        if hasattr(fields_keys_included, '__len__') and hasattr(fields_keys_included, '__getitem__'):
            field_keys_list = [str(fields_keys_included[i]) for i in range(len(fields_keys_included))]
        else:
            print(f"⚠️ 无法获取字段列表")
            return False

        # 处理记录数
        if isinstance(number_records, (int, float)):
            num_records = int(number_records)
        else:
            print(f"⚠️ 无法解析记录数")
            num_records = 0

        # 处理数据
        if hasattr(table_data, '__len__') and hasattr(table_data, '__getitem__'):
            table_data_list = [str(table_data[i]) for i in range(len(table_data))]
        else:
            table_data_list = []

        if num_records == 0 or len(table_data_list) == 0:
            print(f"⚠️ 表格中没有数据记录")
            return False

        print(f"📋 成功获取 {num_records} 条记录")
        print(f"📝 字段: {field_keys_list}")

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

            written_count = 0

            if component_names:
                # 需要过滤特定构件
                unique_name_index = find_component_name_column(field_keys_list)

                if unique_name_index is not None:
                    for row in data_rows:
                        if len(row) > unique_name_index and row[unique_name_index] in component_names:
                            writer.writerow(row)
                            written_count += 1
                else:
                    print("⚠️ 无法确定构件名称字段，保存所有数据")
                    for row in data_rows:
                        writer.writerow(row)
                    written_count = len(data_rows)
            else:
                # 不需要过滤，保存所有数据
                for row in data_rows:
                    writer.writerow(row)
                written_count = len(data_rows)

            print(f"✅ 成功保存 {written_count} 条{description}数据")
            print(f"📄 文件已保存至: {output_file}")

        return written_count > 0

    except Exception as e:
        print(f"❌ 数组方法提取失败: {e}")
        traceback.print_exc()
        return False


def filter_csv_by_components(input_file, output_file, component_names):
    """
    按构件名称过滤CSV文件

    Args:
        input_file: 输入CSV文件路径
        output_file: 输出CSV文件路径
        component_names: 构件名称列表

    Returns:
        bool: 过滤是否成功
    """
    try:
        with open(input_file, 'r', encoding='utf-8-sig') as infile:
            with open(output_file, 'w', newline='', encoding='utf-8-sig') as outfile:
                reader = csv.reader(infile)
                writer = csv.writer(outfile)

                headers = next(reader)
                writer.writerow(headers)

                # 找到构件名称列
                name_col_index = find_component_name_column(headers)

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
                return written_count > 0

    except Exception as e:
        print(f"❌ CSV过滤失败: {e}")
        return False


def find_component_name_column(headers):
    """
    在表头中查找构件名称列的索引

    Args:
        headers: 表头列表

    Returns:
        int: 构件名称列索引，如果找不到返回None
    """
    name_keywords = [
        'unique', 'uniquename', 'element', 'label', 'name', 'beam', 'column'
    ]

    for i, header in enumerate(headers):
        header_lower = header.lower().replace(' ', '').replace('_', '')
        for keyword in name_keywords:
            if keyword in header_lower and 'combo' not in header_lower:
                return i

    return None


def generate_comprehensive_summary_report(column_names, beam_names, concrete_tables, extraction_results):
    """
    生成综合的混凝土设计数据汇总报告

    Args:
        column_names (list): 框架柱名称列表
        beam_names (list): 框架梁名称列表
        concrete_tables (dict): 表格配置字典
        extraction_results (dict): 提取结果字典

    Returns:
        bool: 报告生成是否成功
    """
    try:
        output_file = os.path.join(SCRIPT_DIRECTORY, 'concrete_design_comprehensive_report.txt')

        with open(output_file, 'w', encoding='utf-8') as f:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write("=" * 100 + "\n")
            f.write("混凝土设计数据综合提取报告\n")
            f.write(f"报告生成时间: {now}\n")
            f.write("=" * 100 + "\n\n")

            # 提取概况
            total_tables = len(concrete_tables)
            successful_tables = sum(1 for success in extraction_results.values() if success)

            f.write("📊 提取概况\n")
            f.write("-" * 50 + "\n")
            f.write(f"计划提取表格数量: {total_tables}\n")
            f.write(f"成功提取表格数量: {successful_tables}\n")
            f.write(f"提取成功率: {successful_tables / total_tables * 100:.1f}%\n\n")

            # 构件信息
            f.write("🏗️ 构件信息\n")
            f.write("-" * 50 + "\n")
            f.write(f"框架柱数量: {len(column_names)}\n")
            f.write(f"框架梁数量: {len(beam_names)}\n\n")

            # 详细提取结果
            f.write("📋 详细提取结果\n")
            f.write("-" * 50 + "\n")

            for table_id, table_config in concrete_tables.items():
                success = extraction_results.get(table_id, False)
                status = "✅ 成功" if success else "❌ 失败"
                f.write(f"{status} {table_config['description']}\n")
                f.write(f"     文件名: {table_config['filename']}\n")
                f.write(f"     表格键: {table_config['table_key']}\n")

                # 检查文件是否存在并统计记录数
                file_path = os.path.join(SCRIPT_DIRECTORY, table_config['filename'])
                if success and os.path.exists(file_path):
                    try:
                        with open(file_path, 'r', encoding='utf-8-sig') as csv_file:
                            record_count = sum(1 for line in csv_file) - 1  # 减去表头
                        f.write(f"     记录数: {record_count} 条\n")
                    except:
                        f.write(f"     记录数: 无法读取\n")
                f.write("\n")

            # 生成的文件列表
            f.write("📄 生成的文件列表\n")
            f.write("-" * 50 + "\n")
            file_index = 1
            for table_id, table_config in concrete_tables.items():
                if extraction_results.get(table_id, False):
                    f.write(f"{file_index}. {table_config['filename']} - {table_config['description']}\n")
                    file_index += 1
            f.write(f"{file_index}. concrete_design_comprehensive_report.txt - 本综合报告\n\n")

            # 数据字段说明
            f.write("📝 数据字段说明\n")
            f.write("-" * 50 + "\n")
            f.write("通用字段:\n")
            f.write("  Story          - 楼层\n")
            f.write("  Label/Element  - 构件标签\n")
            f.write("  UniqueName     - 构件唯一名称\n")
            f.write("  Section        - 截面名称\n")
            f.write("  Location       - 位置/测点\n")
            f.write("  Combo          - 荷载组合\n\n")

            f.write("梁弯曲包络字段:\n")
            f.write("  -ve Moment     - 负弯矩 (kN·m)\n")
            f.write("  +ve Moment     - 正弯矩 (kN·m)\n")
            f.write("  As Top         - 顶部钢筋面积 (mm²)\n")
            f.write("  As Bot         - 底部钢筋面积 (mm²)\n\n")

            f.write("梁剪切包络字段:\n")
            f.write("  V2             - 剪力 (kN)\n")
            f.write("  Av/s           - 箍筋配筋率 (mm²/mm)\n")
            f.write("  VRebar         - 箍筋承担剪力 (kN)\n\n")

            f.write("柱PMM包络字段:\n")
            f.write("  P              - 轴力 (kN)\n")
            f.write("  M2             - 2轴弯矩 (kN·m)\n")
            f.write("  M3             - 3轴弯矩 (kN·m)\n")
            f.write("  AsReqd         - 所需钢筋面积 (mm²)\n")
            f.write("  AsProv         - 提供钢筋面积 (mm²)\n\n")

            f.write("柱剪切包络字段:\n")
            f.write("  V2             - 2方向剪力 (kN)\n")
            f.write("  V3             - 3方向剪力 (kN)\n")
            f.write("  Av2/s          - 2方向箍筋配筋率 (mm²/mm)\n")
            f.write("  Av3/s          - 3方向箍筋配筋率 (mm²/mm)\n\n")

            f.write("节点包络字段:\n")
            f.write("  Joint          - 节点名称\n")
            f.write("  VRatio         - 剪力比\n")
            f.write("  BCCRatio       - 梁柱接触比\n")
            f.write("  Status         - 设计状态\n\n")

            # 使用说明
            f.write("📖 使用说明\n")
            f.write("-" * 50 + "\n")
            f.write("1. 包络数据说明:\n")
            f.write("   • 包络数据为所有荷载组合下的最不利值\n")
            f.write("   • 正负弯矩分别对应受拉区在不同侧的情况\n")
            f.write("   • 配筋面积为满足承载力要求的最小配筋\n\n")

            f.write("2. 数据验证建议:\n")
            f.write("   • 对比ETABS界面显示的设计结果\n")
            f.write("   • 检查关键构件的配筋是否合理\n")
            f.write("   • 验证荷载组合的完整性\n")
            f.write("   • 确认设计规范参数设置正确\n\n")

            f.write("3. 注意事项:\n")
            f.write("   • 本数据仅供设计参考，不能直接用于施工\n")
            f.write("   • 需要结合构造要求进行配筋调整\n")
            f.write("   • 建议进行人工复核重要构件\n")
            f.write("   • 最终设计以正式图纸为准\n\n")

            # 故障排除
            f.write("🔧 故障排除\n")
            f.write("-" * 50 + "\n")
            f.write("如果某些表格提取失败，可能的原因:\n")
            f.write("1. 对应的设计类型未完成计算\n")
            f.write("2. 设计规范选择不正确\n")
            f.write("3. 模型中没有对应类型的构件\n")
            f.write("4. ETABS版本与API兼容性问题\n\n")

            f.write("解决建议:\n")
            f.write("1. 重新运行混凝土设计计算\n")
            f.write("2. 检查设计偏好设置\n")
            f.write("3. 确认模型包含待设计构件\n")
            f.write("4. 尝试手动导出对应表格\n\n")

            f.write("=" * 100 + "\n")
            f.write("报告生成完成\n")
            f.write("如有问题请检查ETABS设计状态或联系技术支持\n")
            f.write("=" * 100 + "\n")

        print(f"✅ 综合汇总报告已保存至: {output_file}")
        return True

    except Exception as e:
        print(f"❌ 生成综合汇总报告失败: {e}")
        traceback.print_exc()
        return False


def debug_concrete_design_tables(sap_model):
    """
    调试函数：列出所有可用的混凝土设计相关表格
    用于排查表格名称和可用性

    Args:
        sap_model: ETABS模型对象

    Returns:
        list: 可用表格列表
    """
    try:
        print("🔍 调试：列出所有可用的混凝土设计表格...")

        from common.etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return []

        db = sap_model.DatabaseTables

        # 扩展的混凝土设计表格候选列表
        concrete_design_tables = [
            # 基础表格名称
            "Concrete Beam Flexure Envelope",
            "Concrete Beam Shear Envelope",
            "Concrete Column PMM Envelope",
            "Concrete Column Shear Envelope",
            "Concrete Joint Envelope",

            # 带规范后缀的表格名称
            "Concrete Beam Flexure Envelope - Chinese 2010",
            "Concrete Beam Shear Envelope - Chinese 2010",
            "Concrete Column PMM Envelope - Chinese 2010",
            "Concrete Column Shear Envelope - Chinese 2010",
            "Concrete Joint Envelope - Chinese 2010",

            # Table前缀的表格名称
            "Table: Concrete Beam Flexure Envelope",
            "Table: Concrete Beam Shear Envelope",
            "Table: Concrete Column PMM Envelope",
            "Table: Concrete Column Shear Envelope",
            "Table: Concrete Joint Envelope",

            # Table前缀带规范的表格名称
            "Table: Concrete Beam Flexure Envelope - Chinese 2010",
            "Table: Concrete Beam Shear Envelope - Chinese 2010",
            "Table: Concrete Column PMM Envelope - Chinese 2010",
            "Table: Concrete Column Shear Envelope - Chinese 2010",
            "Table: Concrete Joint Envelope - Chinese 2010",

            # 其他可能的命名格式
            "Concrete Frame Design 2 - Beam Flexure Envelope",
            "Concrete Frame Design 2 - Beam Shear Envelope",
            "Concrete Frame Design 2 - Column PMM Envelope",
            "Concrete Frame Design 2 - Column Shear Envelope",
            "Concrete Frame Design 2 - Joint Envelope",

            # 详细设计表格
            "Concrete Beam Detail Data",
            "Concrete Column Detail Data",
            "Concrete Frame Summary Data"
        ]

        available_tables = []
        table_info = {}

        for table in concrete_design_tables:
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

                success = False
                record_count = 0

                if isinstance(ret, tuple):
                    error_code = ret[0]
                    if error_code == 0:
                        success = True
                        # 尝试获取记录数
                        if len(ret) > 4:
                            try:
                                record_count = int(ret[4]) if hasattr(ret[4], '__int__') else 0
                            except:
                                record_count = 0
                elif ret == 0:
                    success = True

                if success:
                    available_tables.append(table)
                    table_info[table] = record_count
                    print(f"✅ {table} (记录数: {record_count})")

            except Exception as e:
                continue

        print(f"\n📊 总结:")
        print(f"✅ 找到 {len(available_tables)} 个可用的混凝土设计表格")

        if available_tables:
            print(f"\n📋 表格详情:")
            for table in available_tables:
                count = table_info.get(table, 0)
                status = "有数据" if count > 0 else "无数据"
                print(f"   • {table} - {status} ({count} 条记录)")
        else:
            print("❌ 未找到任何可用的混凝土设计表格")
            print("💡 可能的原因:")
            print("   1. 混凝土设计未完成")
            print("   2. 设计规范选择问题")
            print("   3. 模型中没有混凝土构件")

        return available_tables

    except Exception as e:
        print(f"❌ 调试混凝土设计表格时出错: {e}")
        traceback.print_exc()
        return []


def export_table_definitions(sap_model):
    """
    导出表格定义，帮助理解表格结构

    Args:
        sap_model: ETABS模型对象
    """
    try:
        print("📋 正在导出表格字段定义...")

        from common.etabs_api_loader import get_api_objects
        ETABSv1, System, COMException = get_api_objects()

        if System is None:
            print("❌ System对象未正确加载")
            return

        db = sap_model.DatabaseTables

        # 获取可用表格
        available_tables = debug_concrete_design_tables(sap_model)

        if not available_tables:
            print("❌ 没有可用表格，无法导出定义")
            return

        output_file = os.path.join(SCRIPT_DIRECTORY, 'table_field_definitions.txt')

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("ETABS混凝土设计表格字段定义\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")

            for table_name in available_tables[:5]:  # 只处理前5个表格以避免过长
                try:
                    f.write(f"表格: {table_name}\n")
                    f.write("-" * 50 + "\n")

                    # 获取字段定义
                    field_key_list = System.Array.CreateInstance(System.String, 1)
                    field_key_list[0] = ""

                    group_name = ""
                    table_version = System.Int32(0)
                    fields_keys_included = System.Array.CreateInstance(System.String, 0)
                    number_records = System.Int32(0)
                    table_data = System.Array.CreateInstance(System.String, 0)

                    ret = db.GetTableForDisplayArray(
                        table_name,
                        field_key_list,
                        group_name,
                        table_version,
                        fields_keys_included,
                        number_records,
                        table_data
                    )

                    if isinstance(ret, tuple) and len(ret) > 3:
                        fields_included = ret[3]
                        if hasattr(fields_included, '__len__') and hasattr(fields_included, '__getitem__'):
                            field_list = [str(fields_included[i]) for i in range(len(fields_included))]
                            f.write(f"字段数量: {len(field_list)}\n")
                            f.write("字段列表:\n")
                            for i, field in enumerate(field_list):
                                f.write(f"  {i + 1:2d}. {field}\n")
                        else:
                            f.write("无法获取字段信息\n")
                    else:
                        f.write("表格访问失败\n")

                    f.write("\n")

                except Exception as e:
                    f.write(f"处理表格时出错: {e}\n\n")

        print(f"✅ 表格字段定义已保存至: {output_file}")

    except Exception as e:
        print(f"❌ 导出表格定义失败: {e}")


# 主程序入口点更新
def main():
    """
    主程序入口，用于测试扩展的提取功能
    """
    print("=" * 80)
    print("混凝土设计数据提取模块 - 扩展版本")
    print("=" * 80)

    # 示例构件名称（实际使用时应从模型中获取）
    example_columns = ['C1', 'C2', 'C3', 'C4']
    example_beams = ['B1', 'B2', 'B3', 'B4']

    try:
        # 确保ETABS连接
        if not ensure_etabs_ready():
            print("❌ 无法连接ETABS，请确保ETABS已打开")
            return False

        sap_model = get_sap_model()
        if sap_model is None:
            print("❌ 无法获取ETABS模型")
            return False

        print("✅ ETABS连接成功")

        # 调试模式：列出可用表格
        print("\n🔍 调试模式：检查可用表格...")
        debug_concrete_design_tables(sap_model)

        # 导出表格定义
        print("\n📋 导出表格字段定义...")
        export_table_definitions(sap_model)

        # 提取所有混凝土设计数据
        print("\n🚀 开始提取混凝土设计数据...")
        success = extract_all_concrete_design_data(example_columns, example_beams)

        if success:
            print("\n🎉 混凝土设计数据提取完成！")
        else:
            print("\n⚠️ 混凝土设计数据提取部分完成或失败")

        return success

    except Exception as e:
        print(f"❌ 主程序执行失败: {e}")
        traceback.print_exc()
        return False


if __name__ == "__main__":
    print("此模块是ETABS自动化项目的扩展部分，应在主程序中调用。")
    print("直接运行此文件将进入测试模式。")
    print("\n如果需要测试此模块，请确保:")
    print("1. ETABS已打开并加载了完成混凝土设计的模型")
    print("2. 已运行混凝土构件设计计算")
    print("3. 设计规范设置正确（如Chinese 2010）")

    # 询问是否运行测试
    response = input("\n是否运行测试模式？(y/n): ").lower().strip()
    if response in ['y', 'yes', '是']:
        main()
    else:
        print("测试已取消。")
