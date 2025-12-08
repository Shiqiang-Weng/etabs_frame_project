#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Parameter sampling and config helpers for parametric frame modeling."""

from __future__ import annotations

import json
import math
import random
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

_HERE = Path(__file__).resolve().parent
_PROJECT_ROOT = _HERE.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from common import config

# ---- Design space definitions -------------------------------------------------

TOPOLOGY_BOUNDS = {
    "N_st": (6, 10, 1),
    "n_x": (3, 5, 1),
    "n_y": (5, 10, 1),
    "l_x": (3600, 7200, 600),
    "l_y": (3600, 7200, 600),
}

BEAM_HEIGHT_RANGE_LOCAL = (400, 800)
BEAM_HEIGHT_RANGE_GLOBAL = (400, 1000)
BEAM_WIDTH_RANGE_LOCAL = (150, 400)
BEAM_WIDTH_RANGE_GLOBAL = (200, 550)
BEAM_STEP = 50

COLUMN_SIZE_RANGE = (400, 800)
COLUMN_STEP = 100

# Empirical column sizing coefficients
COLUMN_LAMBDA_FACTOR = 1.35
COLUMN_FC = 20.0  # MPa

# Unified surface load (kN/m2) used for axial load estimation
UNIFORM_SURFACE_LOAD = config.DEFAULT_DEAD_SUPER_SLAB + config.DEFAULT_LIVE_LOAD_SLAB


@dataclass
class DesignCaseConfig:
    topology: Dict[str, Any]
    sizing: Dict[str, Any]
    group_mapping: Dict[str, Any]
    case_id: int

    @classmethod
    def from_sample(cls, case_id: int, sample: Dict[str, Any]) -> "DesignCaseConfig":
        return cls(
            topology=sample["Topology"],
            sizing=sample["Sizing"],
            group_mapping=sample["GroupMapping"],
            case_id=case_id,
        )


# ---- Core sampling helpers ----------------------------------------------------

def _round_to_step(value: float, step: int) -> int:
    return int(step * round(float(value) / step))


def _clamp(value: float, min_value: float, max_value: float) -> float:
    return max(min_value, min(max_value, value))


def sample_topology(rng: random.Random) -> Dict[str, int]:
    topo = {}
    for key, (v_min, v_max, step) in TOPOLOGY_BOUNDS.items():
        values = list(range(v_min, v_max + step, step))
        topo[key] = rng.choice(values)
    return topo


def split_story_groups(num_stories: int) -> Dict[str, Any]:
    base = num_stories // 3
    remainder = num_stories % 3
    counts = [
        base + (1 if remainder > 0 else 0),
        base + (1 if remainder > 1 else 0),
        base,
    ]
    groups: Dict[str, Dict[str, List[int]]] = {}
    story_to_group: Dict[int, str] = {}
    start = 1
    labels = ["bottom", "middle", "top"]

    for idx, count in enumerate(counts):
        group_name = f"Group{idx + 1}"
        stories = list(range(start, start + count))
        groups[group_name] = {"label": labels[idx], "stories": stories}
        for story in stories:
            story_to_group[story] = group_name
        start += count

    return {"groups": groups, "story_to_group": story_to_group}


def sample_beam_section(span_length_mm: float, rng: random.Random) -> Tuple[int, int]:
    """Sample (b, h) for a beam span length using empirical formulas (mm)."""
    mu_h = span_length_mm / 48.0
    sigma_h = span_length_mm / 240.0

    h = None
    for _ in range(5000):
        h = _round_to_step(rng.gauss(mu_h, sigma_h), BEAM_STEP)
        if BEAM_HEIGHT_RANGE_LOCAL[0] <= h <= BEAM_HEIGHT_RANGE_LOCAL[1]:
            break
    else:
        # Fallback: clamp mean into local range if repeated resampling fails
        h = _round_to_step(_clamp(mu_h, *BEAM_HEIGHT_RANGE_LOCAL), BEAM_STEP)

    mu_b = 5 * h / 12.0
    sigma_b = h / 24.0
    b = None
    for _ in range(5000):
        b = _round_to_step(rng.gauss(mu_b, sigma_b), BEAM_STEP)
        if BEAM_WIDTH_RANGE_LOCAL[0] <= b <= BEAM_WIDTH_RANGE_LOCAL[1]:
            break
    else:
        b = _round_to_step(_clamp(mu_b, *BEAM_WIDTH_RANGE_LOCAL), BEAM_STEP)

    if b > h:
        b = h

    h = _clamp(h, *BEAM_HEIGHT_RANGE_GLOBAL)
    b = _clamp(b, *BEAM_WIDTH_RANGE_GLOBAL)

    return int(b), int(h)


def calculate_column_baseline(N_v_kN: float, lambda_factor: float, fc: float) -> int:
    """Estimate square column width (mm) from axial load (kN)."""
    area_mm2 = lambda_factor * (N_v_kN * 1e3) / (0.75 * fc)
    b_mm = math.sqrt(area_mm2)
    b_mm = _round_to_step(b_mm, COLUMN_STEP)
    b_mm = _clamp(b_mm, *COLUMN_SIZE_RANGE)
    return int(b_mm)


def _derive_column_sections(base_b: int) -> Tuple[int, int, int]:
    corner = _clamp(_round_to_step(base_b * 0.9, BEAM_STEP), *COLUMN_SIZE_RANGE)
    edge = _clamp(_round_to_step(base_b, BEAM_STEP), *COLUMN_SIZE_RANGE)
    interior = _clamp(_round_to_step(base_b * 1.1, BEAM_STEP), *COLUMN_SIZE_RANGE)
    ordered = sorted([corner, edge, interior])
    # enforce corner <= edge <= interior ordering after rounding
    return int(ordered[0]), int(ordered[1]), int(ordered[2])


def generate_structural_sample(rng: Optional[random.Random] = None) -> Dict[str, Any]:
    rng = rng or random.Random()
    topo = sample_topology(rng)
    group_mapping = split_story_groups(topo["N_st"])

    span_max = max(topo["l_x"], topo["l_y"])
    footprint_x_m = topo["n_x"] * topo["l_x"] / 1000.0
    footprint_y_m = topo["n_y"] * topo["l_y"] / 1000.0
    floor_area = footprint_x_m * footprint_y_m
    num_columns = (topo["n_x"] + 1) * (topo["n_y"] + 1)
    floor_axial_per_column = (floor_area * UNIFORM_SURFACE_LOAD) / num_columns

    sizing: Dict[str, Dict[str, float]] = {}
    prev_columns: Optional[Dict[str, int]] = None
    prev_beams: Optional[Dict[str, int]] = None

    for idx, group_name in enumerate(["Group1", "Group2", "Group3"], start=1):
        stories = group_mapping["groups"][group_name]["stories"]
        if not stories:
            continue

        carried_stories = topo["N_st"] - min(stories) + 1
        N_v = floor_axial_per_column * carried_stories
        base_b = calculate_column_baseline(N_v, COLUMN_LAMBDA_FACTOR, COLUMN_FC)
        col_corner, col_edge, col_interior = _derive_column_sections(base_b)

        b_edge, h_edge = sample_beam_section(span_max, rng)
        b_int, h_int = sample_beam_section(span_max, rng)

        if prev_columns:
            col_corner = min(col_corner, prev_columns["corner"])
            col_edge = min(col_edge, prev_columns["edge"])
            col_interior = min(col_interior, prev_columns["interior"])
        if prev_beams:
            b_edge = min(b_edge, prev_beams["edge_b"])
            h_edge = min(h_edge, prev_beams["edge_h"])
            b_int = min(b_int, prev_beams["int_b"])
            h_int = min(h_int, prev_beams["int_h"])

        sizing[group_name] = {
            f"C_G{idx}_Corner_b": col_corner,
            f"C_G{idx}_Edge_b": col_edge,
            f"C_G{idx}_Interior_b": col_interior,
            f"B_G{idx}_Edge_b": b_edge,
            f"B_G{idx}_Edge_h": h_edge,
            f"B_G{idx}_Interior_b": b_int,
            f"B_G{idx}_Interior_h": h_int,
        }

        prev_columns = {"corner": col_corner, "edge": col_edge, "interior": col_interior}
        prev_beams = {"edge_b": b_edge, "edge_h": h_edge, "int_b": b_int, "int_h": h_int}

    return {"Topology": topo, "Sizing": sizing, "GroupMapping": group_mapping}


# ---- Plan generation and IO ---------------------------------------------------

def flatten_sample_for_key(sample: Dict[str, Any]) -> Tuple:
    topo = sample["Topology"]
    sizing = sample["Sizing"]

    key_parts: List[Any] = [
        topo["N_st"],
        topo["n_x"],
        topo["n_y"],
        topo["l_x"],
        topo["l_y"],
    ]

    for gid in range(1, 4):
        group_name = f"Group{gid}"
        params = sizing[group_name]
        key_parts.extend(
            [
                params[f"C_G{gid}_Corner_b"],
                params[f"C_G{gid}_Edge_b"],
                params[f"C_G{gid}_Interior_b"],
                params[f"B_G{gid}_Edge_b"],
                params[f"B_G{gid}_Edge_h"],
                params[f"B_G{gid}_Interior_b"],
                params[f"B_G{gid}_Interior_h"],
            ]
        )

    return tuple(key_parts)


def _write_plan_jsonl(samples: List[Dict[str, Any]], out_path: Path) -> None:
    with out_path.open("w", encoding="utf-8") as f:
        for sample in samples:
            f.write(json.dumps(sample, ensure_ascii=False) + "\n")


def _write_plan_csv(samples: List[Dict[str, Any]], out_path: Path) -> Path:
    import csv

    csv_path = out_path.with_suffix(".csv")
    headers = [
        "case_id",
        "N_st",
        "n_x",
        "n_y",
        "l_x_mm",
        "l_y_mm",
    ]
    for gid in range(1, 4):
        headers.extend(
            [
                f"C_G{gid}_Corner_b",
                f"C_G{gid}_Edge_b",
                f"C_G{gid}_Interior_b",
                f"B_G{gid}_Edge_b",
                f"B_G{gid}_Edge_h",
                f"B_G{gid}_Interior_b",
                f"B_G{gid}_Interior_h",
            ]
        )

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for sample in samples:
            topo = sample["Topology"]
            sizing = sample["Sizing"]
            row: List[Any] = [
                sample["case_id"],
                topo["N_st"],
                topo["n_x"],
                topo["n_y"],
                topo["l_x"],
                topo["l_y"],
            ]
            for gid in range(1, 4):
                params = sizing[f"Group{gid}"]
                row.extend(
                    [
                        params[f"C_G{gid}_Corner_b"],
                        params[f"C_G{gid}_Edge_b"],
                        params[f"C_G{gid}_Interior_b"],
                        params[f"B_G{gid}_Edge_b"],
                        params[f"B_G{gid}_Edge_h"],
                        params[f"B_G{gid}_Interior_b"],
                        params[f"B_G{gid}_Interior_h"],
                    ]
                )
            writer.writerow(row)

    return csv_path


def generate_param_plan(num_cases: int, out_path: Path, seed: Optional[int] = None) -> None:
    rng = random.Random(seed)
    seen: set[Tuple] = set()
    samples: List[Dict[str, Any]] = []

    max_attempts = max(num_cases * 20, num_cases + 10)
    attempts = 0
    out_path.parent.mkdir(parents=True, exist_ok=True)
    print(f"[param_sampling] target={num_cases}, max_attempts={max_attempts}, seed={seed}")
    last_log_attempt = 0
    last_log_unique = 0

    while len(samples) < num_cases and attempts < max_attempts:
        attempts += 1
        sample = generate_structural_sample(rng)
        key = flatten_sample_for_key(sample)
        if key in seen:
            if attempts - last_log_attempt >= 20000 and len(samples) == last_log_unique:
                print(
                    f"[param_sampling] still deduping... attempts={attempts}, unique={len(samples)}",
                    flush=True,
                )
                last_log_attempt = attempts
                last_log_unique = len(samples)
            continue

        case_id = len(samples)
        sample_out = {
            "case_id": case_id,
            "Topology": sample["Topology"],
            "Sizing": sample["Sizing"],
            "GroupMapping": sample["GroupMapping"],
        }
        samples.append(sample_out)
        seen.add(key)
        if case_id > 0 and case_id % 1000 == 0:
            print(f"[param_sampling] progress: {case_id}/{num_cases} unique samples (attempts={attempts})", flush=True)
            last_log_attempt = attempts
            last_log_unique = len(samples)

    if len(samples) < num_cases:
        print(f"[警告] 仅生成 {len(samples)} / {num_cases} 个唯一样本（尝试 {attempts} 次）")
    else:
        print(f"[完成] 生成 {len(samples)} 个唯一样本，重复跳过 {attempts - len(samples)} 次")

    _write_plan_jsonl(samples, out_path)
    csv_path = _write_plan_csv(samples, out_path)
    print(f"[param_sampling] 写入完成: JSONL -> {out_path}, CSV -> {csv_path}", flush=True)


def load_param_plan(path: Path) -> List[Dict[str, Any]]:
    samples: List[Dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            text = line.strip()
            if not text:
                continue
            samples.append(json.loads(text))
    return samples


__all__ = [
    "DesignCaseConfig",
    "calculate_column_baseline",
    "flatten_sample_for_key",
    "generate_param_plan",
    "generate_structural_sample",
    "load_param_plan",
    "sample_beam_section",
    "sample_topology",
    "split_story_groups",
]


if __name__ == "__main__":
    num_cases = 10
    out_path = _HERE / "param_plan.jsonl"
    print(f"[param_sampling] 开始生成 {num_cases} 个唯一样本 → {out_path}")
    generate_param_plan(num_cases=num_cases, out_path=out_path, seed=42)
    print(
        f"[param_sampling] 生成完成，JSONL 写入: {out_path}, "
        f"CSV 写入: {out_path.with_suffix('.csv')}"
    )
