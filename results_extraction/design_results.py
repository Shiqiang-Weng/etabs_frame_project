#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Design result extraction utilities (migrated from design_module).
"""

from __future__ import annotations

import os
import csv
import time
import traceback
from typing import Any, Dict, List

from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret
from common.etabs_api_loader import get_api_objects

ETABSv1, System, COMException = get_api_objects()


def _ensure_api_objects():
    """Lazy-refresh ETABS API objects to avoid None during design extraction."""
    global ETABSv1, System, COMException
    if System is None or ETABSv1 is None:
        ETABSv1, System, COMException = get_api_objects()
    return ETABSv1, System, COMException

# ====================  ====================

def convert_system_array_to_python_list(system_array):
    """Convert System.Array (or Python iterable) to a Python list safely."""
    _ensure_api_objects()
    if system_array is None:
        return []

    try:
        # System.Array?
        if hasattr(system_array, 'Length'):
            result = []
            for i in range(system_array.Length):
                result.append(system_array[i])
            return result
        elif hasattr(system_array, '__len__'):
            return list(system_array)
        else:
            return [system_array] if system_array is not None else []
    except Exception as e:
        print(f"     System.Array: {e}")
        return []


def convert_area_units(area_in_m2: float) -> float:
    """Convert area from m^2 to mm^2 (with legacy correction factor)."""
    if area_in_m2 is None or area_in_m2 == 0:
        return 0.0
    # ?m ?mm
    #   1,000,000 ?
    # ?
    corrected_area_mm2 = (area_in_m2 * 1000000) / 1000
    return corrected_area_mm2


def convert_shear_area_units(shear_area_in_m2_per_m: float) -> float:
    """Convert shear reinforcement area from m^2/m to mm^2/m."""
    if shear_area_in_m2_per_m is None or shear_area_in_m2_per_m == 0:
        return 0.0
    # : m/m * (1000mm/m) = mm/m
    return shear_area_in_m2_per_m * 1000000


def validate_reinforcement_area(area_mm2: float, element_type: str = "column") -> Dict[str, Any]:
    """Quick reasonableness check for reinforcement area to flag unit mistakes."""
    area_value = area_mm2 if area_mm2 is not None else 0.0
    validation_result = {
        "is_valid": False,
        "area_mm2": area_value,
        "area_cm2": area_value / 100,
        "warnings": [],
        "suggestions": [],
    }

    element = (element_type or "").lower()
    if element == "column":
        if area_value < 1000:  # < 10 cm^2
            validation_result["warnings"].append("Column reinforcement area appears too small; check units")
        elif area_value > 50000:  # > 500 cm^2
            validation_result["warnings"].append("Column reinforcement area appears too large; check units")
        elif 1000 <= area_value <= 20000:
            validation_result["is_valid"] = True
        if area_value > 100000:
            validation_result["suggestions"].append("Verify unit conversion; value is unusually large")

    elif element == "beam":
        if area_value < 500:  # < 5 cm^2
            validation_result["warnings"].append("Beam reinforcement area appears too small")
        elif area_value > 30000:  # > 300 cm^2
            validation_result["warnings"].append("Beam reinforcement area appears too large; check units")
        elif 500 <= area_value <= 15000:
            validation_result["is_valid"] = True

    else:
        validation_result["warnings"].append(f"Unknown element type: {element_type}")

    return validation_result
def _get_beam_design_summary_enhanced(design_concrete, beam_name: str) -> Dict[str, Any]:
    """Enhanced beam design summary using ETABS API."""
    try:
        # ?
        error_code, number_results = 1, 0
        top_areas, bot_areas, vmajor_areas = [], [], []
        source = "API-"

        # PI
        if hasattr(design_concrete, 'GetSummaryResultsBeam_2'):
            try:
                #  GetSummaryResultsBeam_2 (26 parameters)
                # We pass placeholders for the 'ref' parameters
                result = design_concrete.GetSummaryResultsBeam_2(
                    beam_name, 0, [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [],
                    [], []
                )

                if isinstance(result, tuple) and len(result) == 25:
                    source = "API-2-"
                    # Unpack all 25 results
                    (error_code, number_results, _, _, _, top_areas, _, _, _,
                     _, bot_areas, _, _, _, _, vmajor_areas, _, _, _,
                     _, _, _, _, _, _) = result
                else:
                    return {"Source": "API-2-", "Error": f": {type(result)}, : {len(result)}"}
            except Exception as e_2:
                # If GetSummaryResultsBeam_2 fails, log it and fallback
                print(f"     GetSummaryResultsBeam_2  ({beam_name}): {e_2}, API...")
                pass  # Fallback will be attempted below

        # APIPI
        if source != "API-2-":
            result = design_concrete.GetSummaryResultsBeam(
                beam_name, 0, [], [], [], [], [], [], [], [], [], [], [], [], [], []
            )

            if not isinstance(result, tuple) or len(result) != 16:
                return {"Source": "API-1-", "Error": f": {type(result)}, : {len(result)}"}

            # API?
            (error_code, number_results, _, _, _, top_areas, _, bot_areas,
             _, vmajor_areas, _, _, _, _, _, _) = result
            source = "API-1-"

        # PI
        if error_code != 0:
            return {"Source": source.replace("", ""), "Error": f"API: {error_code}"}

        # 
        if number_results == 0:
            return {"Source": source.replace("", "unknown"), "Warning": ""}

        # System.Arrayython?
        try:
            top_areas_list = convert_system_array_to_python_list(top_areas)
            bot_areas_list = convert_system_array_to_python_list(bot_areas)
            vmajor_areas_list = convert_system_array_to_python_list(vmajor_areas)

            # ?
            top_areas_mm2 = [convert_area_units(float(x)) for x in top_areas_list if x is not None and x > 0]
            bot_areas_mm2 = [convert_area_units(float(x)) for x in bot_areas_list if x is not None and x > 0]
            # ?
            vmajor_areas_mm2_per_m = [convert_shear_area_units(float(x)) for x in vmajor_areas_list if
                                      x is not None and x > 0]

            max_top = max(top_areas_mm2) if top_areas_mm2 else 0.0
            max_bot = max(bot_areas_mm2) if bot_areas_mm2 else 0.0
            max_vmajor = max(vmajor_areas_mm2_per_m) if vmajor_areas_mm2_per_m else 0.0

            # ?
            top_validation = validate_reinforcement_area(max_top, "unknown")
            bot_validation = validate_reinforcement_area(max_bot, "unknown")

            result_dict = {
                "Source": source,
                "Top_As_mm2": round(max_top, 2),
                "Bot_As_mm2": round(max_bot, 2),
                "V_Major_As_mm2_per_m": round(max_vmajor, 2),  # 
                "Top_As_cm2": round(max_top / 100, 2),
                "Bot_As_cm2": round(max_bot / 100, 2),
                "Num_Results": number_results,
                "Top_Validation": "" if top_validation["is_valid"] else "unknown",
                "Bot_Validation": "" if bot_validation["is_valid"] else "unknown"
            }

            # 
            warnings = []
            if top_validation["warnings"]:
                warnings.extend(top_validation["warnings"])
            if bot_validation["warnings"]:
                warnings.extend(bot_validation["warnings"])

            if warnings:
                result_dict["Warnings"] = "; ".join(warnings)

            return result_dict

        except Exception as parse_error:
            return {"Source": source.replace("", ""), "Error": f": {str(parse_error)}"}

    except Exception as e:
        return {"Source": "API-", "Error": str(e)}


def _get_column_design_summary_enhanced(design_concrete, col_name: str) -> Dict[str, Any]:
    """unknown"""
    try:
        if not hasattr(design_concrete, 'GetSummaryResultsColumn'):
            return {"Source": "API-unknown", "Error": "GetSummaryResultsColumn not available"}

        # PI
        try:
            result = design_concrete.GetSummaryResultsColumn(
                col_name,  # column name
                0,  # NumberItems
                [],  # FrameName
                [],  # Location
                [],  # PMMCombo
                [],  # PMMArea
                [],  # PMMRatio
                [],  # VMajorCombo
                [],  # VMinorCombo
                [],  # ErrorSummary
                [],  # WarningSummary
            )
        except Exception as api_error:
            # 11
            parameter_counts = [9, 10, 12, 13, 14, 15, 16]
            for param_count in parameter_counts:
                try:
                    params = [col_name, 0] + [[] for _ in range(param_count - 2)]
                    result = design_concrete.GetSummaryResultsColumn(*params)
                    break
                except:
                    continue
            else:
                return {"Source": "API-", "Error": f": {str(api_error)}"}

        # ?
        if not isinstance(result, tuple) or len(result) < 2:
            return {"Source": "API-", "Error": f""}

        # 
        error_code = result[0] if len(result) > 0 else 1
        number_results = result[1] if len(result) > 1 else 0

        # PI
        if error_code != 0:
            return {"Source": "API-", "Error": f"API: {error_code}"}

        # no data returned
        if number_results == 0:
            return {"Source": "API-no-data", "Warning": "Element has no design results"}

        # 
        try:
            pmm_areas = None
            pmm_ratios = None

            # System.Double[]?
            for i in range(2, len(result)):
                item = result[i]
                if str(type(item)) == "<class 'System.Double[]'>":
                    if pmm_areas is None:
                        pmm_areas = item
                    elif pmm_ratios is None:
                        pmm_ratios = item
                        break

            if pmm_areas is not None:
                pmm_areas_list = convert_system_array_to_python_list(pmm_areas)
                # ?
                pmm_areas_mm2 = [convert_area_units(float(x)) for x in pmm_areas_list if x is not None and x != 0]
                max_area = max(pmm_areas_mm2) if pmm_areas_mm2 else 0.0
            else:
                max_area = 0.0
                pmm_areas_list = []

            if pmm_ratios is not None:
                pmm_ratios_list = convert_system_array_to_python_list(pmm_ratios)
                pmm_ratios_float = [float(x) for x in pmm_ratios_list if x is not None]
                avg_ratio = sum(pmm_ratios_float) / len(pmm_ratios_float) if pmm_ratios_float else 0.0
            else:
                avg_ratio = 0.0

            # ?
            area_validation = validate_reinforcement_area(max_area, "unknown")

            result_dict = {
                "Source": "API-",
                "Total_As_mm2": round(max_area, 2),
                "Total_As_cm2": round(max_area / 100, 2),
                "PMM_Ratio": round(avg_ratio, 6),
                "PMM_Combo": "",
                "Num_Results": number_results,
                "Raw_PMM_Count": len(pmm_areas_list) if pmm_areas else 0,
                "Error_Code": error_code,
                "Area_Validation": "" if area_validation["is_valid"] else "unknown"
            }

            # 
            if area_validation["warnings"]:
                result_dict["Validation_Warnings"] = "; ".join(area_validation["warnings"])

            if area_validation["suggestions"]:
                result_dict["Validation_Suggestions"] = "; ".join(area_validation["suggestions"])

            return result_dict

        except Exception as parse_error:
            return {
                "Source": "API-",
                "Total_As_mm2": 0.0,
                "Total_As_cm2": 0.0,
                "PMM_Ratio": 0.0,
                "PMM_Combo": "",
                "Num_Results": number_results,
                "Error_Code": error_code,
                "Parse_Error": str(parse_error)
            }

    except Exception as e:
        return {"Source": "API-", "Error": str(e)}


def extract_design_results_enhanced() -> List[Dict[str, Any]]:
    """Enhanced beam/column design extraction with basic validation and summaries."""
    _ensure_api_objects()
    _ensure_api_objects()
    _, sap_model = get_etabs_objects()
    print("\n--- Enhanced design extraction ---")

    try:
        print("   Preparing story and frame lists...")

        NumberStories, StoryNamesArr = 0, System.Array.CreateInstance(System.String, 0)
        ret, number_stories, story_names_tuple = sap_model.Story.GetNameList(NumberStories, StoryNamesArr)
        story_names = list(story_names_tuple)
        check_ret(ret, "Story.GetNameList")

        print(f"  Stories detected: {number_stories}")

        all_frame_names = []
        for story in story_names:
            NumberItemsOnStory, StoryFrameNamesArr = 0, System.Array.CreateInstance(System.String, 0)
            ret, count, story_frames_tuple = sap_model.FrameObj.GetNameListOnStory(story, NumberItemsOnStory,
                                                                                   StoryFrameNamesArr)
            if ret == 0 and count > 0:
                all_frame_names.extend(list(story_frames_tuple))

        frame_names = sorted(list(set(all_frame_names)))
        if not frame_names:
            print("No frame names found; skipping design results extraction.")
            return []

        # simple name heuristics
        beam_names = [n for n in frame_names if any(kw in n.upper() for kw in ['BEAM', 'B_', 'B-'])]
        column_names = [n for n in frame_names if
                        any(kw in n.upper() for kw in ['COL_', 'COL-', 'C_', 'C-', 'COLUMN'])]

        print(f"  Frames detected: beams={len(beam_names)}, columns={len(column_names)}")

        design_concrete = sap_model.DesignConcrete
        all_results = []

        print(f"\n    Processing beams...")
        beam_success_count = 0
        beam_no_data_count = 0
        beam_warning_count = 0

        for i, name in enumerate(beam_names):
            if (i + 1) % 50 == 0 or i == len(beam_names) - 1:
                print(f"    Beam progress: {i + 1}/{len(beam_names)}")

            result = _get_beam_design_summary_enhanced(design_concrete, name)
            if "" in result.get("Source", ""):
                beam_success_count += 1
                if result.get("Warnings"):
                    beam_warning_count += 1
            elif "unknown" in result.get("Source", ""):
                beam_no_data_count += 1
            all_results.append({"Frame_Name": name, "Element_Type": "unknown", **result})

        print(
            f"  Beams - success: {beam_success_count}, no data: {beam_no_data_count}, warnings: {beam_warning_count}"
        )

        print(f"\n    Processing columns...")
        col_success_count = 0
        col_partial_count = 0
        col_no_data_count = 0
        col_validation_warning_count = 0

        for i, name in enumerate(column_names):
            if (i + 1) % 30 == 0 or i == len(column_names) - 1:
                print(
                    f"    Column progress ({i + 1}/{len(column_names)}) - success: {col_success_count}, partial: {col_partial_count}, warnings: {col_validation_warning_count}"
                )

            result = _get_column_design_summary_enhanced(design_concrete, name)
            if result.get("Source") == "API-":
                col_success_count += 1
                if result.get("Area_Validation") == "unknown":
                    col_validation_warning_count += 1
            elif result.get("Source") == "API-":
                col_partial_count += 1
            elif result.get("Source") == "API-unknown":
                col_no_data_count += 1
            all_results.append({"Frame_Name": name, "Element_Type": "unknown", **result})

        print(
            f"  Columns - success: {col_success_count}, partial: {col_partial_count}, warnings: {col_validation_warning_count} "
        )

        total_success = beam_success_count + col_success_count + col_partial_count
        print(f"\n   Total processed: {total_success}/{len(all_results)}")

        successful_columns = [r for r in all_results if r.get("Element_Type") == "unknown" and r.get("Source") == "API-"]
        if successful_columns:
            areas_mm2 = [float(r.get("Total_As_mm2", 0)) for r in successful_columns if r.get("Total_As_mm2")]
            areas_cm2 = [a / 100 for a in areas_mm2]

            if areas_mm2:
                print(f"\n   Column reinforcement statistics:")
                print(
                    f"    Range: {min(areas_mm2):.0f} - {max(areas_mm2):.0f} mm^2 ({min(areas_cm2):.1f} - {max(areas_cm2):.1f} cm^2)")
                print(
                    f"    Average: {sum(areas_mm2) / len(areas_mm2):.0f} mm^2 ({sum(areas_cm2) / len(areas_cm2):.1f} cm^2)")

                # ?
                reasonable_count = sum(1 for r in successful_columns if r.get("Area_Validation") == "")
                print(
                    f"    : {reasonable_count}/{len(successful_columns)} ({reasonable_count / len(successful_columns) * 100:.1f}%)")

        return all_results

    except Exception as e:
        print(f"Warning: failed to extract design results: {e}")
        traceback.print_exc()
        return []


def save_design_results_enhanced(design_data: List[Dict[str, Any]], output_dir: str):
    """CSV"""
    if not design_data:
        print(" ")
        return

    filepath = os.path.join(output_dir, "concrete_design_results_enhanced.csv")
    print(f"\n : {filepath}")

    try:
        all_keys = set().union(*(d.keys() for d in design_data))

        # ?
        fieldnames = [
            'Frame_Name', 'Element_Type', 'Source',
            # ?
            'Top_As_mm2', 'Bot_As_mm2', 'V_Major_As_mm2_per_m',
            'Top_As_cm2', 'Bot_As_cm2',
            'Top_Validation', 'Bot_Validation',
            # ?
            'Total_As_mm2', 'Total_As_cm2', 'PMM_Ratio', 'PMM_Combo',
            'Area_Validation', 'Validation_Warnings', 'Validation_Suggestions',
            # ?
            'Num_Results', 'Raw_PMM_Count', 'Error_Code',
            'Parse_Error', 'Warning', 'Error', 'Warnings'
        ]

        final_fieldnames = [k for k in fieldnames if k in all_keys] + sorted(
            [k for k in all_keys if k not in fieldnames])

        with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=final_fieldnames, extrasaction='ignore')
            writer.writeheader()
            writer.writerows(design_data)

        print(f"Total design records: {len(design_data)}")

        # 
        print_enhanced_validation_statistics(design_data, output_dir)

    except Exception as e:
        print(f"Failed to write combined CSV: {e}")


def print_enhanced_validation_statistics(design_data: List[Dict[str, Any]], output_dir: str):
    """Print and write a simple validation summary for design results."""
    if not design_data:
        print("No design data to summarize.")
        return

    successful_columns = [r for r in design_data if r.get("Element_Type") == "column" and "API-" in r.get("Source", "")]
    successful_beams = [r for r in design_data if r.get("Element_Type") == "beam" and "API-" in r.get("Source", "")]

    stats_lines = [
        "=== Validation Summary ===",
        f"Time: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        f"Total entries: {len(design_data)}",
        f"With design results: {len(successful_columns) + len(successful_beams)}",
        f"Beams: {len(successful_beams)}",
        f"Columns: {len(successful_columns)}",
    ]

    if successful_columns:
        reasonable_count = sum(1 for r in successful_columns if r.get("Area_Validation") == "")
        stats_lines.append("")
        stats_lines.append("Column area validation:")
        stats_lines.append(
            f"  OK: {reasonable_count}/{len(successful_columns)} ({reasonable_count / len(successful_columns) * 100:.1f}%)"
        )
        stats_lines.append(
            f"  Needs review: {len(successful_columns) - reasonable_count}/{len(successful_columns)} "
            f"({(len(successful_columns) - reasonable_count) / len(successful_columns) * 100:.1f}%)"
        )

    if successful_beams:
        beam_reasonable_top = sum(1 for r in successful_beams if r.get("Top_Validation") == "")
        beam_reasonable_bot = sum(1 for r in successful_beams if r.get("Bot_Validation") == "")
        stats_lines.append("")
        stats_lines.append("Beam validation:")
        stats_lines.append(
            f"  OK (top): {beam_reasonable_top}/{len(successful_beams)} ({beam_reasonable_top / len(successful_beams) * 100:.1f}%)"
        )
        stats_lines.append(
            f"  OK (bottom): {beam_reasonable_bot}/{len(successful_beams)} ({beam_reasonable_bot / len(successful_beams) * 100:.1f}%)"
        )

    stats_text = "`n".join(stats_lines) + "`n"
    print(stats_text)

    stats_file = os.path.join(output_dir, "validation_statistics_enhanced.txt")
    try:
        with open(stats_file, 'w', encoding='utf-8') as f:
            f.write(stats_text)
        print(f"Saved validation statistics to {stats_file}")
    except Exception as e:
        print(f"Failed to write validation statistics: {e}")
def generate_enhanced_summary_report(output_dir: str):
    """Write a lightweight design summary report pointing to key outputs."""
    os.makedirs(output_dir, exist_ok=True)
    beam_path = os.path.join(output_dir, "beam_design_results_final.csv")
    column_path = os.path.join(output_dir, "column_design_results_final.csv")
    enhanced_path = os.path.join(output_dir, "concrete_design_results_enhanced.csv")
    report_path = os.path.join(output_dir, "design_summary_report.txt")

    lines = [
        "Design Result Summary",
        "---------------------",
        f"Beam results: {beam_path}",
        f"Column results: {column_path}",
        f"Enhanced combined results: {enhanced_path}",
    ]

    try:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        print(f"Summary report written to {report_path}")
    except Exception as e:
        print(f"Failed to write summary report: {e}")
def extract_and_save_beam_results(output_dir: str) -> None:
    """
    ?
    """
    _ensure_api_objects()
    _, sap_model = get_etabs_objects()
    print("\n--- Beam design results ---")
    os.makedirs(output_dir, exist_ok=True)

    try:
        dc = sap_model.DesignConcrete

        number_names = 0
        frame_names_tuple = System.Array.CreateInstance(System.String, 0)
        ret, number_names, frame_names_tuple = sap_model.FrameObj.GetNameList(number_names, frame_names_tuple)
        if ret != 0:
            print("  Failed to get frame name list.")
            return

        frame_names = list(frame_names_tuple)
        beam_names = [name for name in frame_names if name.upper().startswith("BEAM")]

        if not beam_names:
            print("  No beams found.")
            return

        print(f"   {len(beam_names)} beams to process...")
        all_results = []
        valid_results = 0

        for i, name in enumerate(beam_names):
            if (i + 1) % 50 == 0:
                print(f"    Progress: {i + 1}/{len(beam_names)}")

            result = {"Frame_Name": name}
            try:
                res = dc.GetSummaryResultsBeam(name, 0, [], [], [], [], [], [], [], [], [], [], [], [], [], [])
                ret_code, num_items, _, _, _, top_areas, _, bot_areas, *_ = res

                if ret_code == 0 and num_items > 0:
                    top_areas_list = [a for a in convert_system_array_to_python_list(top_areas) if a > 0]
                    bot_areas_list = [a for a in convert_system_array_to_python_list(bot_areas) if a > 0]

                    max_top = max(top_areas_list) if top_areas_list else 0
                    max_bot = max(bot_areas_list) if bot_areas_list else 0

                    result.update({"Src": "OK", "Top_Rebar_m2": f"{max_top:.6f}", "Bot_Rebar_m2": f"{max_bot:.6f}"})
                    valid_results += 1
                else:
                    result.update({"Src": "No Results", "Top_Rebar_m2": 0, "Bot_Rebar_m2": 0})

            except Exception as exc:  # noqa: BLE001
                result.update({"Src": f"Error: {str(exc)[:40]}", "Top_Rebar_m2": 0, "Bot_Rebar_m2": 0})

            all_results.append(result)

        filepath = os.path.join(output_dir, "beam_design_results_final.csv")
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=all_results[0].keys())
            writer.writeheader()
            writer.writerows(all_results)

        print(f"Beam results saved to {filepath}")
        print(f"   Completed: {valid_results}/{len(beam_names)}")

    except Exception as exc:  # noqa: BLE001
        print(f"Failed to save beam results: {exc}")


def extract_and_save_column_results(output_dir: str) -> None:
    """Extract and save column design summaries (original format)."""
    _, sap_model = get_etabs_objects()
    print("\n--- Column design results ---")
    os.makedirs(output_dir, exist_ok=True)

    try:
        dc = sap_model.DesignConcrete

        number_names = 0
        frame_names_tuple = System.Array.CreateInstance(System.String, 0)
        ret, number_names, frame_names_tuple = sap_model.FrameObj.GetNameList(number_names, frame_names_tuple)
        if ret != 0:
            print("  Failed to get frame name list.")
            return

        frame_names = list(frame_names_tuple)
        column_names = [name for name in frame_names if name.upper().startswith("COL")]

        if not column_names:
            print("  No columns found.")
            return

        print(f"   {len(column_names)} columns to process...")
        all_results = []
        valid_results = 0

        for i, name in enumerate(column_names):
            if (i + 1) % 50 == 0:
                print(f"    Progress: {i + 1}/{len(column_names)}")

            result = {"Frame_Name": name}
            try:
                res = dc.GetSummaryResultsColumn(name, 0, [], [], [], [], [], [], [], [], [], [], [], [])
                ret_code, num_items, pmm_areas, *_ = res

                if ret_code == 0 and num_items > 0:
                    areas = [a for a in convert_system_array_to_python_list(pmm_areas) if a > 0]
                    max_area = max(areas) if areas else 0
                    result.update({"Src": "OK", "Long_Rebar_m2": f"{max_area:.6f}"})
                    valid_results += 1
                else:
                    result.update({"Src": "No Results", "Long_Rebar_m2": 0})

            except Exception as exc:  # noqa: BLE001
                result.update({"Src": f"Error: {str(exc)[:40]}", "Long_Rebar_m2": 0})

            all_results.append(result)

        filepath = os.path.join(output_dir, "column_design_results_final.csv")
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=all_results[0].keys())
            writer.writeheader()
            writer.writerows(all_results)

        print(f"Column results saved to {filepath}")
        print(f"   Completed: {valid_results}/{len(column_names)}")

    except Exception as exc:  # noqa: BLE001
        print(f"Failed to save column results: {exc}")

__all__ = [
    'convert_system_array_to_python_list',
    'convert_area_units',
    'convert_shear_area_units',
    'validate_reinforcement_area',
    '_get_beam_design_summary_enhanced',
    '_get_column_design_summary_enhanced',
    'extract_design_results_enhanced',
    'save_design_results_enhanced',
    'print_enhanced_validation_statistics',
    'generate_enhanced_summary_report',
    'extract_and_save_beam_results',
    'extract_and_save_column_results',
]


