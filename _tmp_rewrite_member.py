from pathlib import Path
p = Path('etabs_frame_project/results_extraction/member_forces.py')
text = p.read_text(encoding='utf-8')
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


def _select_all_cases_and_combos(sap_model, System) -> None:
    """选择所有工况与组合后再导出，确保包含组合行。"""
    results_setup = sap_model.Results.Setup
    try:
        check_ret(results_setup.DeselectAllCasesAndCombosForOutput(), "DeselectAllCasesAndCombosForOutput", (0, 1))
    except Exception as exc:
        print(f"[WARN] 清空已有结果选择失败: {exc}")

    if hasattr(results_setup, "SetAllCasesAndCombosSelectedForOutput"):
        try:
            check_ret(results_setup.SetAllCasesAndCombosSelectedForOutput(), "SetAllCasesAndCombosSelectedForOutput", (0, 1))
            print("已选择所有工况与组合用于内力导出。")
            return
        except Exception as exc:
            print(f"[WARN] SetAllCasesAndCombosSelectedForOutput 失败，改为逐个选择: {exc}")

    try:
        num_case = System.Int32(0)
        case_names = System.Array[System.String](0)
        ret = sap_model.LoadCases.GetNameList(num_case, case_names)
        if isinstance(ret, tuple) and ret[0] == 0 and ret[1] > 0:
            for name in list(ret[2]):
                try:
                    check_ret(results_setup.SetCaseSelectedForOutput(name), f"SetCaseSelectedForOutput({name})", (0, 1))
                except Exception:
                    pass
    except Exception as exc:
        print(f"[WARN] 选择工况失败: {exc}")

    try:
        num_combo = System.Int32(0)
        combo_names = System.Array[System.String](0)
        ret_c = sap_model.RespCombo.GetNameList(num_combo, combo_names)
        if isinstance(ret_c, tuple) and ret_c[0] == 0 and ret_c[1] > 0:
            for name in list(ret_c[2]):
                try:
                    check_ret(results_setup.SetComboSelectedForOutput(name), f"SetComboSelectedForOutput({name})", (0, 1))
                except Exception:
                    pass
    except Exception as exc:
        print(f"[WARN] 选择组合失败: {exc}")
'''
text = text[:start] + new_block + text[end:]
p.write_text(text, encoding='utf-8')
print('rewrote export_element_forces_table + selector')
