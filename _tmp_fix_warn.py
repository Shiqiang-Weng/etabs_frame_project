from pathlib import Path
text = Path('etabs_frame_project/results_extraction/member_forces.py').read_text(encoding='utf-8')
text = text.replace("?? {table_name} 表导出失败，错误码: {err_code}", "⚠️ {table_name} 表导出失败，错误码: {err_code}")
Path('etabs_frame_project/results_extraction/member_forces.py').write_text(text, encoding='utf-8')
print('warn strings replaced')
