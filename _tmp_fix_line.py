from pathlib import Path
path=Path('etabs_frame_project/results_extraction/design_forces.py')
text=path.read_text(encoding='utf-8')
old='        print(f"? 基本构件内力数据已保存至: {output_file}")'
new='        print(f"✅ 基本构件内力数据已保存至: {output_file}")'
if old not in text:
    raise SystemExit('target line not found')
text=text.replace(old,new,1)
path.write_text(text,encoding='utf-8')
print('replaced print line')
