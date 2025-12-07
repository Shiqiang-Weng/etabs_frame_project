# -*- coding: utf-8 -*-
from pathlib import Path
path = Path('etabs_frame_project/results_extraction/design_forces.py')
text = path.read_text(encoding='utf-8')
start = text.find('def extract_basic_frame_forces(')
end = text.find('\n# =============================================================================\n# 导出符号清单', start)
if start == -1 or end == -1:
    raise SystemExit(f'block not found start={start} end={end}')
new_block = '''def extract_basic_frame_forces(sap_model, column_names, beam_names):
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

        output_file = os.path.join(DATA_EXTRACTION_DIR, "basic_frame_forces.csv")
        with open(output_file, "w", newline="", encoding="utf-8-sig") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(field_keys_list)
            num_fields = len(field_keys_list)
            if num_fields > 0:
                data_rows = [
                    table_data_list[i : i + num_fields]
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
'''
new_text = text[:start] + new_block + text[end:]
path.write_text(new_text, encoding='utf-8')
print('rewrote extract_basic_frame_forces')
