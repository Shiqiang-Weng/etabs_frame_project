from pathlib import Path
text = Path('etabs_frame_project/results_extraction/member_forces.py').read_text(encoding='utf-8')
start = text.find('def export_element_forces_table')
end = text.find('\ndef export_beam_and_column_element_forces', start)
if start == -1 or end == -1:
    raise SystemExit(f'block not found start={start} end={end}')
new_block = '''def export_element_forces_table(db, System, table_name: str, out_csv_path: Path) -> bool:
    """
    从 ETABS 数据库表导出梁/柱单元内力表到 CSV。
    table_name 例如 'Element Forces - Beams' 或 'Element Forces - Columns'
    """
    try:
        out_csv_path.parent.mkdir(parents=True, exist_ok=True)
        print(f"正在导出表: {table_name}")

        success, ret_csv, file_size = export_table_to_csv(
            db, System, table_name, str(out_csv_path), table_version=1
        )

        err_code = ret_csv[0] if isinstance(ret_csv, tuple) else ret_csv
        if err_code != 0 or not success:
            print(f"⚠️ {table_name} 表导出失败，错误码: {err_code}")
            return False

        record_count = 0
        try:
            with open(out_csv_path, "r", encoding="utf-8-sig") as f:
                record_count = max(sum(1 for _ in f) - 1, 0)
        except Exception:
            record_count = 0

        print(f"错误码: {err_code} | 记录数: {record_count} | CSV: {out_csv_path}")
        return True
    except Exception as e:
        print(f"[ERROR] 导出 {table_name} 失败: {e}")
        traceback.print_exc()
        return False
'''
text = text[:start] + new_block + text[end:]
Path('etabs_frame_project/results_extraction/member_forces.py').write_text(text, encoding='utf-8')
print('rewrote export_element_forces_table block')
