from pathlib import Path
path = Path('results_extraction/design_results.py')
text = path.read_text(encoding='utf-8')
for func in ['extract_and_save_beam_results', 'extract_and_save_column_results']:
    target = f"def {func}(output_dir: str) -> None:\n    \"\"\"\n    备用版"
    idx = text.find(target)
    if idx == -1:
        raise SystemExit(f'target missing {func}')
    insert_pos = idx + target.find('\n') + len('\n')*0
text = text.replace("    _, sap_model = get_etabs_objects()\n", "    _ensure_api_objects()\n    _, sap_model = get_etabs_objects()\n", 2)
path.write_text(text, encoding='utf-8')
