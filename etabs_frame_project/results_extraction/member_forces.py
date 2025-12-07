#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
 results_extraction 
?"""

from __future__ import annotations

import os
import csv
import traceback
from typing import List, Dict, Any, Optional

from pathlib import Path
from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret
from common.config import ANALYSIS_DATA_DIR
from common.etabs_api_loader import get_api_objects
from results_extraction.export_utils import export_table_to_csv


def export_element_forces_table(db, System, table_name: str, out_csv_path: Path) -> bool:
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
        # 不再过滤 Case Type，仅依赖 ETABS 结果选择（全工况+组合）

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

        if record_count <= 0:
            print(f"⚠️ {table_name} 表导出失败，未获取到记录（错误码: {err_code}）")
            return False

        case_types_summary, sample_output_cases, has_combo = _log_case_type_overview(out_csv_path)
        combo_rows_case = sum(v for k, v in case_types_summary.items() if k and (k.lower().startswith("comb") or k == "Combination"))
        combo_rows_exact, combo_samples = _log_combination_samples(out_csv_path)
        # 控制台明确打印组合统计与示例，便于确认已包含组合内力
        print(f"错误码: {err_code} | 记录数: {record_count} | CSV: {out_csv_path}")
        print(f"Case Type 统计: {case_types_summary}")
        print(f"Output Case 示例: {sample_output_cases}")
        print(f"   - 组合行数(按 Case Type 统计): {combo_rows_case}")
        print(f"   - 组合行数(按行过滤统计): {combo_rows_exact}")
        print(f"   - 组合 Output Case 示例: {combo_samples}")
        if has_combo or combo_rows_exact > 0:
            print("✅ 梁/柱 Element Forces 已包含工况 + 组合内力")
        else:
            print(f"⚠️ {table_name} 导出未检测到组合行，请检查结果选择。")
        return True
    except Exception as e:
        print(f"[ERROR] 导出 {table_name} 失败: {e}")
        traceback.print_exc()
        return False


def _select_all_cases_and_combos(sap_model, System, db=None) -> None:
    """选择所有工况与组合后再导出，确保包含组合行。"""
    results_setup = sap_model.Results.Setup
    selected_cases = 0
    selected_load_combos = 0  # 设计/荷载组合（如 DConS1）
    selected_design_combos = 0  # 设计组合（若 API 提供 DesignCombo/DesignComb）
    selected_resp_combos = 0  # 反应谱等响应组合
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
                    selected_cases += 1
                except Exception:
                    pass
    except Exception as exc:
        print(f"[WARN] 选择工况失败: {exc}")

    # 先尝试选择荷载组合（设计组合，例如 DConS1 等），兼容不同 API 名称
    try:
        for attr_name, label in [("LoadCombo", "LoadCombo"), ("LoadCombos", "LoadCombos"), ("Combo", "Combo")]:
            if hasattr(sap_model, attr_name):
                num_lc = System.Int32(0)
                lc_names = System.Array[System.String](0)
                ret_lc = getattr(sap_model, attr_name).GetNameList(num_lc, lc_names)
                if isinstance(ret_lc, tuple) and ret_lc[0] == 0 and ret_lc[1] > 0:
                    lc_list = list(ret_lc[2])
                    print(f"[INFO] {label} 可用数量: {len(lc_list)}, 示例: {lc_list[:5]}")
                    for name in lc_list:
                        try:
                            check_ret(results_setup.SetComboSelectedForOutput(name), f"SetComboSelectedForOutput({name})", (0, 1))
                            selected_load_combos += 1
                        except Exception:
                            pass
                break  # 找到一个可用接口即可
    except Exception as exc:
        print(f"[WARN] 选择荷载组合失败: {exc}")

    # 尝试选择设计组合（如 API 提供 DesignCombo/DesignComb，可能包含 DConS1 等）
    try:
        for attr_name, label in [("DesignCombo", "DesignCombo"), ("DesignComb", "DesignComb"), ("DesignCombinations", "DesignCombinations")]:
            if hasattr(sap_model, attr_name):
                num_dc = System.Int32(0)
                dc_names = System.Array[System.String](0)
                ret_dc = getattr(sap_model, attr_name).GetNameList(num_dc, dc_names)
                if isinstance(ret_dc, tuple) and ret_dc[0] == 0 and ret_dc[1] > 0:
                    dc_list = list(ret_dc[2])
                    print(f"[INFO] {label} 可用数量: {len(dc_list)}, 示例: {dc_list[:5]}")
                    for name in dc_list:
                        try:
                            check_ret(results_setup.SetComboSelectedForOutput(name), f"SetComboSelectedForOutput({name})", (0, 1))
                            selected_design_combos += 1
                        except Exception:
                            pass
                break
    except Exception as exc:
        print(f"[WARN] 选择设计组合失败: {exc}")

    # 进一步尝试通过 ConcreteDesign API 获取设计组合（如 DConS1）
    try:
        design_combo_names = _get_concrete_design_combos(sap_model, System)
        if design_combo_names:
            for name in design_combo_names:
                try:
                    check_ret(results_setup.SetComboSelectedForOutput(name), f"SetComboSelectedForOutput({name})", (0, 1))
                    selected_design_combos += 1
                except Exception:
                    pass
    except Exception as exc:
        print(f"[WARN] ConcreteDesign 设计组合选择失败: {exc}")

    # 再尝试选择响应组合（如 RS 相关）
    try:
        num_combo = System.Int32(0)
        combo_names = System.Array[System.String](0)
        ret_c = sap_model.RespCombo.GetNameList(num_combo, combo_names)
        if isinstance(ret_c, tuple) and ret_c[0] == 0 and ret_c[1] > 0:
            rc_list = list(ret_c[2])
            print(f"[INFO] RespCombos 可用数量: {len(rc_list)}, 示例: {rc_list[:5]}")
            for name in rc_list:
                try:
                    check_ret(results_setup.SetComboSelectedForOutput(name), f"SetComboSelectedForOutput({name})", (0, 1))
                    selected_resp_combos += 1
                except Exception:
                    pass
    except Exception as exc:
        print(f"[WARN] 选择组合失败: {exc}")

    selected_combos = selected_load_combos + selected_design_combos + selected_resp_combos
    if selected_cases or selected_combos:
        print(
            f"已手动选择 {selected_cases} 个工况、"
            f"{selected_load_combos} 个荷载组合、{selected_design_combos} 个设计组合、"
            f"{selected_resp_combos} 个响应组合用于内力导出。"
        )
    # 如仍未找到组合，尝试从组合定义表格中读取名称再选中
    if selected_combos == 0 and db is not None:
        _select_combos_from_tables(db, System, results_setup)
    if selected_combos == 0:
        print("[WARN] 当前未找到任何组合，请确认已生成荷载/设计组合，或在设计完成后重试导出。")


def _log_case_type_overview(csv_path: Path):
    """简单统计 Case Type/Output Case，确认包含组合行。"""
    case_type_counts = {}
    output_cases_sample = set()
    has_combo = False
    try:
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                ct = row.get("Case Type") or row.get("CaseType")
                oc = row.get("Output Case") or row.get("OutputCase")
                if ct:
                    case_type_counts[ct] = case_type_counts.get(ct, 0) + 1
                    if not has_combo and (ct.lower().startswith("comb") or ct == "Combination"):
                        has_combo = True
                if oc and len(output_cases_sample) < 6:
                    output_cases_sample.add(oc)
    except Exception as exc:
        print(f"[WARN] 读取 {csv_path} 统计 Case Type 失败: {exc}")
    return case_type_counts, sorted(output_cases_sample), has_combo


def _log_combination_samples(csv_path: Path):
    """额外按行扫描 Combination 行，便于确认是否含设计组合（如 DConS1）。"""
    combo_rows = 0
    combo_samples = set()
    try:
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                ct = (row.get("Case Type") or row.get("CaseType") or "").strip()
                oc = (row.get("Output Case") or row.get("OutputCase") or "").strip()
                if ct.lower().startswith("comb") or ct == "Combination":
                    combo_rows += 1
                    if oc and len(combo_samples) < 6:
                        combo_samples.add(oc)
    except Exception as exc:
        print(f"[WARN] 读取 {csv_path} 统计 Combination 行失败: {exc}")
    return combo_rows, sorted(combo_samples)


def _get_concrete_design_combos(sap_model, System) -> List[str]:
    """尝试通过 ConcreteDesign API 获取设计组合名称（如 DConS1）。"""
    names: List[str] = []
    if not hasattr(sap_model, "DesignConcrete"):
        return names
    dc_api = sap_model.DesignConcrete
    for method in ["GetComboList", "GetCombo", "GetComboDef"]:
        if hasattr(dc_api, method):
            try:
                num = System.Int32(0)
                arr = System.Array[System.String](0)
                ret = getattr(dc_api, method)(num, arr)
                if isinstance(ret, tuple) and ret[0] == 0 and ret[1] > 0:
                    names.extend(list(ret[2]))
            except Exception as exc:
                print(f"[WARN] ConcreteDesign.{method} 获取设计组合失败: {exc}")
    if names:
        uniq = sorted(set(names))
        print(f"[INFO] ConcreteDesign 设计组合数量: {len(uniq)}, 示例: {uniq[:5]}")
    else:
        print("[WARN] ConcreteDesign API 未返回任何设计组合（如 DConS1）。")
    return names


def _select_combos_from_tables(db, System, results_setup) -> None:
    """从组合定义表格读取组合名并尝试选中，作为 API 获取失败时的兜底。"""
    tables = [
        "Load Combination Definitions",
        "Design Combination Definitions",
        "Load Combination Definitions - User",
        "Concrete Frame Design - Combinations",
        "Concrete Design Combination Definitions",
    ]
    for table_key in tables:
        try:
            tmp_path = Path(ANALYSIS_DATA_DIR) / f"__tmp_{table_key.replace(' ', '_').lower()}.csv"
            ok, ret_csv, size = export_table_to_csv(db, System, table_key, str(tmp_path), table_version=1)
            if not ok or size <= 0:
                continue
            names = []
            with open(tmp_path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    name = row.get("Combo Name") or row.get("Combination") or row.get("Name") or row.get("Combo") or ""
                    if name:
                        names.append(name)
            if names:
                print(f"[INFO] 通过表格 {table_key} 获取组合数量: {len(names)}, 示例: {names[:5]}")
                for name in names:
                    try:
                        check_ret(results_setup.SetComboSelectedForOutput(name), f"SetComboSelectedForOutput({name})", (0, 1))
                    except Exception:
                        pass
                break
        except Exception as exc:
            print(f"[WARN] 从表格 {table_key} 读取组合失败: {exc}")


def _load_element_csv(csv_path: Path, member_type: str) -> List[Dict[str, Any]]:
    """读取梁/柱 Element Forces CSV，附加 MemberType 并保留 Story/工况等字段。"""
    records: List[Dict[str, Any]] = []
    if not csv_path.exists():
        print(f"[WARN] 找不到文件: {csv_path}")
        return records
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            record = {
                "Story": row.get("Story", ""),
                "MemberType": member_type,
                "Label": row.get("Beam/Column") or row.get("Label") or "",
                "Unique Name": row.get("Unique Name") or row.get("UniqueName") or "",
                "Output Case": row.get("Output Case") or row.get("OutputCase") or "",
                "Case Type": row.get("Case Type") or row.get("CaseType") or "",
                "Step Type": row.get("Step Type") or row.get("StepType") or "",
                "Station": row.get("Station") or row.get("Sta") or "",
                "P": row.get("P", ""),
                "V2": row.get("V2", ""),
                "V3": row.get("V3", ""),
                "T": row.get("T", ""),
                "M2": row.get("M2", ""),
                "M3": row.get("M3", ""),
                "Element": row.get("Element", ""),
                "Elem Station": row.get("Elem Station") or row.get("ElemStation") or "",
                "Location": row.get("Location", ""),
            }
            records.append(record)
    return records


def _get_story_for_frame(sap_model, System, frame_name: str) -> str:
    """尝试获取构件所在楼层，失败则返回空字符串。"""
    try:
        story = System.String("")
        label = System.String("")
        # GetLabel 返回 Label 和 Story
        ret = sap_model.FrameObj.GetLabel(frame_name, label, story)
        if isinstance(ret, tuple):
            # 可能返回 (ret, label, story)
            if len(ret) >= 3 and ret[2]:
                return str(ret[2])
            if len(ret) >= 2 and ret[1]:
                return str(ret[1])
        if story:
            return str(story)
    except Exception:
        pass
    return ""


def export_beam_and_column_element_forces(output_dir: Path) -> Dict[str, Path]:
    """
    导出梁/柱分析内力表到指定目录（data_extraction），不修改分析/设计逻辑。
    """
    _, sap_model = get_etabs_objects()
    if sap_model is None or not hasattr(sap_model, "DatabaseTables"):
        print("[WARN] SapModel 未就绪，跳过梁/柱内力表导出。")
        return {}

    ETABSv1, System, COMException = get_api_objects()
    if System is None:
        print("[WARN] System 未加载，跳过梁/柱内力表导出。")
        return {}

    db = sap_model.DatabaseTables

    results: Dict[str, Path] = {}
    beam_csv = output_dir / "beam_element_forces.csv"
    col_csv = output_dir / "column_element_forces.csv"

    print("\n--- 导出梁/柱分析内力表 (Element Forces) ---")
    # 先选中所有工况与组合，确保导出的表包含荷载组合
    _select_all_cases_and_combos(sap_model, System, db)
    if export_element_forces_table(db, System, "Element Forces - Beams", beam_csv):
        results["beam_element_forces"] = beam_csv
    if export_element_forces_table(db, System, "Element Forces - Columns", col_csv):
        results["column_element_forces"] = col_csv

    return results


def export_frame_member_forces(output_dir: Path) -> Optional[Path]:
    """
    使用梁/柱 Element Forces 汇总生成 frame_member_forces.csv（包含所有楼层与组合）。
    """
    beam_csv = output_dir / "beam_element_forces.csv"
    col_csv = output_dir / "column_element_forces.csv"

    # 若源文件不存在，先尝试导出一次
    if not beam_csv.exists() or not col_csv.exists():
        print("[INFO] 未找到梁/柱 Element Forces CSV，尝试先导出...")
        export_beam_and_column_element_forces(output_dir)

    beam_records = _load_element_csv(beam_csv, "Beam")
    col_records = _load_element_csv(col_csv, "Column")
    all_records = beam_records + col_records
    if not all_records:
        print("⚠️ 未能生成 frame_member_forces.csv，源数据为空。")
        return None

    output_path = output_dir / "frame_member_forces.csv"
    fieldnames = [
        "Story",
        "MemberType",
        "Label",
        "Unique Name",
        "Output Case",
        "Case Type",
        "Step Type",
        "Station",
        "P",
        "V2",
        "V3",
        "T",
        "M2",
        "M3",
        "Element",
        "Elem Station",
        "Location",
    ]

    print("\n--- 正在生成 frame_member_forces.csv（汇总所有楼层梁柱内力） ---")
    output_dir.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for rec in all_records:
            writer.writerow(rec)

    story_count = len({r.get("Story", "") for r in all_records})
    print("✅ frame_member_forces.csv 已生成：")
    print(f"   - 构件数: {len(all_records)}")
    print(f"   - 楼层数: {story_count}")
    print("   - 已包含工况 + 荷载组合")
    return output_path


def _prepare_force_output_params():
    """
     FrameForce API ?    """
    ETABSv1, System, COMException = get_api_objects()
    return (
        System.Int32(0),  # NumberResults
        System.Array[System.String](0),  # Obj
        System.Array[System.Double](0),  # ObjSta (Corrected to Double)
        System.Array[System.String](0),  # Elm
        System.Array[System.Double](0),  # ElmSta (Corrected to Double)
        System.Array[System.String](0),  # LoadCase
        System.Array[System.String](0),  # StepType
        System.Array[System.Double](0),  # StepNum
        System.Array[System.Double](0),  # P
        System.Array[System.Double](0),  # V2
        System.Array[System.Double](0),  # V3
        System.Array[System.Double](0),  # T
        System.Array[System.Double](0),  # M2
        System.Array[System.Double](0),  # M3
    )


def extract_frame_forces(frame_names: List[str], load_cases: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """提取框架构件内力，默认选中所有工况+组合，支持传入特定工况。"""
    my_etabs, sap_model = get_etabs_objects()
    if not all([sap_model, hasattr(sap_model, "Results")]):
        print("SAP model not initialized; cannot extract frame forces.")
        return []

    ETABSv1, System, COMException = get_api_objects()
    if not all([ETABSv1, System]):
        print("ETABS API not loaded; cannot extract frame forces.")
        return []

    results_api = sap_model.Results
    setup_api = results_api.Setup

    print("\n--- 提取框架构件内力 ---")
    print(f"构件数量: {len(frame_names)}")
    print(f"指定工况: {load_cases if load_cases else '全部工况+组合'}")

    # 1) 结果选择：优先使用“全选工况+组合”，否则按传入列表
    if load_cases:
        check_ret(setup_api.DeselectAllCasesAndCombosForOutput(), "DeselectAllCasesForForces", (0, 1))
        for case in load_cases:
            check_ret(
                setup_api.SetCaseSelectedForOutput(case),
                f"SetCaseSelectedForOutput({case})",
                (0, 1),
            )
    else:
        # 选中所有工况与组合，确保含组合结果
        _select_all_cases_and_combos(sap_model, System)

    all_forces_data: List[Dict[str, Any]] = []
    processed_count = 0

    # 2) 逐构件提取内力
    for frame_name in frame_names:
        try:
            params = _prepare_force_output_params()
            force_res = results_api.FrameForce(frame_name, ETABSv1.eItemTypeElm.ObjectElm, *params)
            check_ret(force_res[0], f"FrameForce({frame_name})", (0, 1))

            num_results = force_res[1]
            # 尝试获取楼层名，便于 CSV 标识楼层
            story_name = _get_story_for_frame(sap_model, System, frame_name)
            if num_results > 0:
                (
                    _,
                    _,
                    obj_names,
                    obj_stas,
                    elm_names,
                    elm_stas,
                    res_cases,
                    step_types,
                    step_nums,
                    p_forces,
                    v2_forces,
                    v3_forces,
                    t_forces,
                    m2_moments,
                    m3_moments,
                ) = force_res

                for i in range(num_results):
                    force_data = {
                        "Story": story_name,
                        "FrameName": obj_names[i],
                        "Station (m)": round(obj_stas[i], 4),
                        "LoadCase": res_cases[i],
                        "P (kN)": round(p_forces[i], 3),
                        "V2 (kN)": round(v2_forces[i], 3),
                        "V3 (kN)": round(v3_forces[i], 3),
                        "T (kN-m)": round(t_forces[i], 3),
                        "M2 (kN-m)": round(m2_moments[i], 3),
                        "M3 (kN-m)": round(m3_moments[i], 3),
                    }
                    all_forces_data.append(force_data)

            processed_count += 1
            if processed_count % 100 == 0:
                print(f"  Progress {processed_count}/{len(frame_names)} ...")

        except Exception as e:
            print(f"   Error retrieving '{frame_name}': {e}")
            # traceback.print_exc()  # 

    print("--- Frame force extraction complete ---")
    print(f" {len(all_forces_data)} records collected")
    if all_forces_data:
        sample_cases = sorted({row.get("LoadCase", "") for row in all_forces_data})[:6]
        story_counts: Dict[str, int] = {}
        for row in all_forces_data:
            story = row.get("Story", "")
            story_counts[story] = story_counts.get(story, 0) + 1
        print(f"   - 输出工况示例: {sample_cases}")
        print(f"   - Story 分布: {story_counts}")
    return all_forces_data


def save_forces_to_csv(force_data: List[Dict[str, Any]], filename: str):
    """Save force data to CSV."""
    if not force_data:
        print("No force data to save.")
        return

    # 统一所有输出文件到分析子目录
    filepath = os.path.join(ANALYSIS_DATA_DIR, filename)
    print(f"\nSaving frame forces to: {filepath}")

    try:
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, "w", newline="", encoding="utf-8-sig") as csvfile:
            fieldnames = force_data[0].keys()
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(force_data)
        print("Frame forces CSV written.")
    except Exception as e:
        print(f"Failed to write frame forces CSV: {e}")


def extract_and_save_frame_forces(all_frame_names: List[str]):
    """
    提取并保存框架构件内力（含全部工况与组合），统一输出到 data_extraction。
    改为基于梁/柱 Element Forces 汇总生成。
    """
    export_frame_member_forces(Path(ANALYSIS_DATA_DIR))


__all__ = [
    "extract_and_save_frame_forces",
    "extract_frame_forces",
    "save_forces_to_csv",
    "export_beam_and_column_element_forces",
    "export_frame_member_forces",
]




