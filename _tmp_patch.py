from pathlib import Path
path = Path('etabs_frame_project/results_extraction/concrete_frame_detail_data.py')
text = path.read_text(encoding='utf-8')
start = text.find('def find_component_name_column')
end = text.find('def generate_comprehensive_summary_report', start)
if start == -1 or end == -1:
    raise SystemExit(f'start {start} end {end}')
text = text[:start] + text[end:]
path.write_text(text, encoding='utf-8')
print('removed local find_component_name_column')
