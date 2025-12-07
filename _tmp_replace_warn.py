from pathlib import Path
p=Path('etabs_frame_project/results_extraction/member_forces.py')
text=p.read_text(encoding='utf-8')
text=text.replace('?? {table_name} 表导出失败，错误码: {err_code}','⚠️ {table_name} 表导出失败，错误码: {err_code}')
p.write_text(text, encoding='utf-8')
print('final replace')
