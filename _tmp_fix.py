# -*- coding: utf-8 -*-
from pathlib import Path
path = Path('etabs_frame_project/results_extraction/design_forces.py')
lines = path.read_text(encoding='utf-8').splitlines()
idx = None
for i, line in enumerate(lines):
    if 'basic_frame_forces.csv' in line and 'output_file' in line:
        idx = i
        break
if idx is None:
    raise SystemExit('not found')
# rewrite block lines idx .. idx+11 expected structure
new_block = [
    '        output_file = os.path.join(DATA_EXTRACTION_DIR, "basic_frame_forces.csv")',
    '        with open(output_file, "w", newline="", encoding="utf-8-sig") as csvfile:',
    '            writer = csv.writer(csvfile)',
    '            writer.writerow(field_keys_list)',
    '            num_fields = len(field_keys_list)',
    '            if num_fields > 0:',
    '                data_rows = [',
    '                    table_data_list[i : i + num_fields]',
    '                    for i in range(0, len(table_data_list), num_fields)',
    '                ]',
    '                for row in data_rows:',
    '                    writer.writerow(row)',
    '        print(f"? 基本构件内力数据已保存至: {output_file}")',
    '        return True',
]
lines[idx:idx+12] = new_block
path.write_text("\n".join(lines) + "\n", encoding='utf-8')
print('patched indentation')
