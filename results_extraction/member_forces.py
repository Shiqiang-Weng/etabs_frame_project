#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
 results_extraction 
?"""

from __future__ import annotations

import os
import csv
import traceback
from typing import List, Dict, Any

from common.etabs_setup import get_etabs_objects
from common.utility_functions import check_ret
from common.config import SCRIPT_DIRECTORY
from common.etabs_api_loader import get_api_objects


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


def extract_frame_forces(frame_names: List[str], load_cases: List[str]) -> List[Dict[str, Any]]:
    """Extract frame forces for the given load cases."""
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

    print("\n--- ?---")
    print(f": {len(frame_names)}")
    print(f": {load_cases}")

    # 1. 
    check_ret(setup_api.DeselectAllCasesAndCombosForOutput(), "DeselectAllCasesForForces", (0, 1))
    for case in load_cases:
        check_ret(
            setup_api.SetCaseSelectedForOutput(case),
            f"SetCaseSelectedForOutput({case})",
            (0, 1),
        )

    all_forces_data = []
    processed_count = 0

    # 2. 
    for frame_name in frame_names:
        try:
            params = _prepare_force_output_params()

            force_res = results_api.FrameForce(frame_name, ETABSv1.eItemTypeElm.ObjectElm, *params)

            check_ret(force_res[0], f"FrameForce({frame_name})", (0, 1))

            num_results = force_res[1]
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
    return all_forces_data


def save_forces_to_csv(force_data: List[Dict[str, Any]], filename: str):
    """Save force data to CSV."""
    if not force_data:
        print("No force data to save.")
        return

    filepath = os.path.join(SCRIPT_DIRECTORY, filename)
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
    ?    """
    target_cases = ["DEAD", "LIVE", "RS-X", "RS-Y"]

    force_results = extract_frame_forces(all_frame_names, target_cases)

    if force_results:
        save_forces_to_csv(force_results, "frame_member_forces.csv")
    else:
        print("No frame forces to write.")


__all__ = [
    "extract_and_save_frame_forces",
    "extract_frame_forces",
    "save_forces_to_csv",
]

