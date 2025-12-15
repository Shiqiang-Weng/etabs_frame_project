#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Parameter sampling and config helpers for parametric frame modeling."""

from __future__ import annotations

import csv
import json
import random
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

_HERE = Path(__file__).resolve().parent
_PROJECT_ROOT = _HERE.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

# ---- Design space definitions -------------------------------------------------

TOPOLOGY_BOUNDS = {
    # 楼层数与网格/跨数范围
    "N_st": (5, 8, 1),
    "n_x": (5, 10, 1),
    "n_y": (3, 5, 1),
    # 跨度范围 (mm)
    "l_x": (3600, 7200, 600),
    "l_y": (3600, 7200, 600),
}

# 梁截面尺寸范围 (mm)
BEAM_HEIGHT_RANGE = (400, 1000)
BEAM_WIDTH_RANGE = (200, 550)
BEAM_STEP = 50

# 柱截面范围 (mm)
COLUMN_SIZE_RANGE = (400, 800)
COLUMN_STEP = 50
MIN_COL_SIZE = 400

_COLUMN_GRID = list(range(COLUMN_SIZE_RANGE[0], COLUMN_SIZE_RANGE[1] + COLUMN_STEP, COLUMN_STEP))
_BEAM_WIDTH_GRID = list(range(BEAM_WIDTH_RANGE[0], BEAM_WIDTH_RANGE[1] + BEAM_STEP, BEAM_STEP))
_BEAM_HEIGHT_GRID = list(range(BEAM_HEIGHT_RANGE[0], BEAM_HEIGHT_RANGE[1] + BEAM_STEP, BEAM_STEP))
_BEAM_PAIR_GRID = [
    (b, h)
    for b in _BEAM_WIDTH_GRID
    for h in _BEAM_HEIGHT_GRID
    if b <= h and 1.5 <= (h / b) <= 3.0
]

PLAN_FILE_PREFIX = "param_plan"
SUPPORTED_PLAN_SUFFIXES = (".jsonl", ".json", ".csv", ".xlsx", ".xls")
CSV_HEADERS = [
    "case_id",
    "N_st",
    "n_x",
    "n_y",
    "l_x_mm",
    "l_y_mm",
    "C_G1_Corner_b",
    "C_G1_Edge_b",
    "C_G1_Interior_b",
    "B_G1_Edge_b",
    "B_G1_Edge_h",
    "B_G1_Interior_b",
    "B_G1_Interior_h",
    "C_G2_Corner_b",
    "C_G2_Edge_b",
    "C_G2_Interior_b",
    "B_G2_Edge_b",
    "B_G2_Edge_h",
    "B_G2_Interior_b",
    "B_G2_Interior_h",
    "C_G3_Corner_b",
    "C_G3_Edge_b",
    "C_G3_Interior_b",
    "B_G3_Edge_b",
    "B_G3_Edge_h",
    "B_G3_Interior_b",
    "B_G3_Interior_h",
]


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


def sample_topology(rng: random.Random) -> Dict[str, int]:
    topo = {}
    for key, (v_min, v_max, step) in TOPOLOGY_BOUNDS.items():
        values = list(range(v_min, v_max + step, step))
        topo[key] = rng.choice(values)
    return topo


def split_story_groups(num_stories: int) -> Dict[str, Any]:
    if num_stories <= 0:
        counts = [0, 0, 0]
    else:
        bottom = 1
        remaining = max(num_stories - 1, 0)
        middle = remaining // 2 + (1 if remaining % 2 else 0)
        top = remaining - middle
        counts = [bottom, middle, top]
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


def sample_column_section(min_value: int, rng: random.Random) -> int:
    candidates = [v for v in _COLUMN_GRID if v >= min_value]
    if not candidates:
        raise ValueError("No valid column size candidates above minimum constraint")
    return int(rng.choice(candidates))


def _pick_beam_pair(min_b: int, min_h: int, rng: random.Random) -> Tuple[int, int]:
    candidates = [(b, h) for (b, h) in _BEAM_PAIR_GRID if b >= min_b and h >= min_h]
    if not candidates:
        raise ValueError("No valid beam pairs above minimum constraint")
    return tuple(rng.choice(candidates))  # type: ignore[return-value]


def sample_downwards(max_bound: int, min_global_bound: int, rng: random.Random, candidates: Optional[List[int]] = None) -> int:
    """
    Sticky downward-biased sampler.
    - 50%: keep max_bound.
    - 50%: pick a strictly smaller value within [min_global_bound, max_bound).
    """
    if max_bound < min_global_bound:
        raise ValueError("max_bound is smaller than the allowed minimum")

    pool = candidates if candidates is not None else list(range(min_global_bound, max_bound + 1))
    pool = [v for v in pool if min_global_bound <= v <= max_bound]
    if not pool:
        raise ValueError("No valid candidates within bounds")

    if max_bound == min_global_bound:
        return min_global_bound

    if rng.random() < 0.5 and max_bound in pool:
        return max_bound

    smaller = [v for v in pool if v < max_bound]
    if not smaller:
        return max_bound
    return int(rng.choice(smaller))


def _pick_beam_pair_with_upper_bounds(max_b: int, max_h: int, rng: random.Random) -> Tuple[int, int]:
    """
    Pick a beam pair not exceeding the provided upper bounds with downward bias.
    Uses sample_downwards for each dimension, then selects any valid pair under the capped sizes.
    """
    width_cap = sample_downwards(max_b, BEAM_WIDTH_RANGE[0], rng, _BEAM_WIDTH_GRID)
    height_cap = sample_downwards(max_h, BEAM_HEIGHT_RANGE[0], rng, _BEAM_HEIGHT_GRID)

    candidates = [(b, h) for (b, h) in _BEAM_PAIR_GRID if b <= width_cap and h <= height_cap]
    if not candidates:
        raise ValueError("No valid beam pairs under the provided upper bounds")

    # If both caps stayed at the ceiling, keep the current size to honor the sticky rule.
    if (width_cap, height_cap) in candidates and width_cap == max_b and height_cap == max_h:
        return width_cap, height_cap

    return tuple(rng.choice(candidates))


def _sample_column_groups(rng: random.Random) -> Dict[str, Dict[str, int]]:
    for _ in range(5000):
        try:
            interior1 = int(rng.choice(_COLUMN_GRID))
            edge1 = sample_downwards(interior1, MIN_COL_SIZE, rng, _COLUMN_GRID)
            corner1 = sample_downwards(edge1, MIN_COL_SIZE, rng, _COLUMN_GRID)

            interior2 = sample_downwards(interior1, MIN_COL_SIZE, rng, _COLUMN_GRID)
            edge2 = sample_downwards(min(edge1, interior2), MIN_COL_SIZE, rng, _COLUMN_GRID)
            corner2 = sample_downwards(min(corner1, edge2), MIN_COL_SIZE, rng, _COLUMN_GRID)

            interior3 = sample_downwards(interior2, MIN_COL_SIZE, rng, _COLUMN_GRID)
            edge3 = sample_downwards(min(edge2, interior3), MIN_COL_SIZE, rng, _COLUMN_GRID)
            corner3 = sample_downwards(min(corner2, edge3), MIN_COL_SIZE, rng, _COLUMN_GRID)

            return {
                "Group1": {
                    "corner": corner1,
                    "edge": edge1,
                    "interior": interior1,
                },
                "Group2": {
                    "corner": corner2,
                    "edge": edge2,
                    "interior": interior2,
                },
                "Group3": {
                    "corner": corner3,
                    "edge": edge3,
                    "interior": interior3,
                },
            }
        except ValueError:
            continue
    raise RuntimeError("Failed to sample monotonic column sections after multiple attempts")


def _sample_beam_groups(rng: random.Random) -> Dict[str, Dict[str, int]]:
    for _ in range(5000):
        try:
            int1_b, int1_h = rng.choice(_BEAM_PAIR_GRID)
            edge1_b, edge1_h = _pick_beam_pair_with_upper_bounds(int1_b, int1_h, rng)

            int2_b, int2_h = _pick_beam_pair_with_upper_bounds(int1_b, int1_h, rng)
            edge2_b, edge2_h = _pick_beam_pair_with_upper_bounds(min(edge1_b, int2_b), min(edge1_h, int2_h), rng)

            int3_b, int3_h = _pick_beam_pair_with_upper_bounds(int2_b, int2_h, rng)
            edge3_b, edge3_h = _pick_beam_pair_with_upper_bounds(min(edge2_b, int3_b), min(edge2_h, int3_h), rng)

            return {
                "Group1": {
                    "edge_b": edge1_b,
                    "edge_h": edge1_h,
                    "int_b": int1_b,
                    "int_h": int1_h,
                },
                "Group2": {
                    "edge_b": edge2_b,
                    "edge_h": edge2_h,
                    "int_b": int2_b,
                    "int_h": int2_h,
                },
                "Group3": {
                    "edge_b": edge3_b,
                    "edge_h": edge3_h,
                    "int_b": int3_b,
                    "int_h": int3_h,
                },
            }
        except ValueError:
            continue
    raise RuntimeError("Failed to sample monotonic beam sections after multiple attempts")


def _assert_monotonic_constraints(sizing: Dict[str, Dict[str, int]]) -> None:
    for gid in range(1, 4):
        params = sizing[f"Group{gid}"]
        corner = params[f"C_G{gid}_Corner_b"]
        edge = params[f"C_G{gid}_Edge_b"]
        interior = params[f"C_G{gid}_Interior_b"]
        edge_b = params[f"B_G{gid}_Edge_b"]
        edge_h = params[f"B_G{gid}_Edge_h"]
        int_b = params[f"B_G{gid}_Interior_b"]
        int_h = params[f"B_G{gid}_Interior_h"]

        assert corner <= edge <= interior, "Column ordering within group violated"
        assert edge_b <= int_b and edge_h <= int_h, "Beam ordering within group violated"
        for b_val, h_val in [(edge_b, edge_h), (int_b, int_h)]:
            ratio = h_val / b_val
            assert b_val <= h_val, "Beam depth must not be smaller than width"
            assert 1.5 <= ratio <= 3.0, "Beam aspect ratio out of bounds"

    for name in ["Corner_b", "Edge_b", "Interior_b"]:
        g1 = sizing["Group1"][f"C_G1_{name}"]
        g2 = sizing["Group2"][f"C_G2_{name}"]
        g3 = sizing["Group3"][f"C_G3_{name}"]
        assert g1 >= g2 >= g3, f"Column monotonicity violated for {name}"

    for beam_key in ["Edge_b", "Edge_h", "Interior_b", "Interior_h"]:
        g1 = sizing["Group1"][f"B_G1_{beam_key}"]
        g2 = sizing["Group2"][f"B_G2_{beam_key}"]
        g3 = sizing["Group3"][f"B_G3_{beam_key}"]
        assert g1 >= g2 >= g3, f"Beam monotonicity violated for {beam_key}"


def generate_structural_sample(rng: Optional[random.Random] = None) -> Dict[str, Any]:
    rng = rng or random.Random()
    topo = sample_topology(rng)
    group_mapping = split_story_groups(topo["N_st"])

    for _ in range(5000):
        columns = _sample_column_groups(rng)
        beams = _sample_beam_groups(rng)
        sizing: Dict[str, Dict[str, int]] = {}

        for idx in range(1, 4):
            sizing[f"Group{idx}"] = {
                f"C_G{idx}_Corner_b": columns[f"Group{idx}"]["corner"],
                f"C_G{idx}_Edge_b": columns[f"Group{idx}"]["edge"],
                f"C_G{idx}_Interior_b": columns[f"Group{idx}"]["interior"],
                f"B_G{idx}_Edge_b": beams[f"Group{idx}"]["edge_b"],
                f"B_G{idx}_Edge_h": beams[f"Group{idx}"]["edge_h"],
                f"B_G{idx}_Interior_b": beams[f"Group{idx}"]["int_b"],
                f"B_G{idx}_Interior_h": beams[f"Group{idx}"]["int_h"],
            }

        try:
            _assert_monotonic_constraints(sizing)
        except AssertionError:
            continue

        return {"Topology": topo, "Sizing": sizing, "GroupMapping": group_mapping}

    raise RuntimeError("Failed to generate structural sample satisfying monotonic constraints")


# ---- Plan generation and IO ---------------------------------------------------

def _resolve_output_paths(out_path: Path) -> Tuple[Path, Path]:
    """
    根据传入路径推导 JSONL/CSV 输出路径。
    - 若指定 .csv，则 CSV 使用该路径，JSONL 同名改后缀。
    - 若指定 .jsonl，则 JSONL 使用该路径，CSV 同名改后缀。
    - 其他情况均按同名生成 jsonl/csv。
    """
    suffix = out_path.suffix.lower()
    if suffix == ".csv":
        csv_path = out_path
        jsonl_path = out_path.with_suffix(".jsonl")
    elif suffix == ".jsonl":
        jsonl_path = out_path
        csv_path = out_path.with_suffix(".csv")
    else:
        jsonl_path = out_path.with_suffix(".jsonl")
        csv_path = out_path.with_suffix(".csv")
    return jsonl_path, csv_path


def _flatten_sample_row(sample: Dict[str, Any]) -> List[Any]:
    """将结构化样本展开为 CSV 行，顺序与 CSV_HEADERS 一致。"""
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
    return row


def _flush_buffer_to_files(buffer: List[Dict[str, Any]], csv_path: Path, jsonl_path: Path, header_written: bool) -> bool:
    """将缓冲区样本追加写入 CSV/JSONL 文件。"""
    if not buffer:
        return header_written

    csv_path.parent.mkdir(parents=True, exist_ok=True)
    jsonl_path.parent.mkdir(parents=True, exist_ok=True)

    with csv_path.open("a", newline="", encoding="utf-8") as f_csv:
        writer = csv.writer(f_csv)
        if not header_written:
            writer.writerow(CSV_HEADERS)
            header_written = True
        for sample in buffer:
            writer.writerow(_flatten_sample_row(sample))

    with jsonl_path.open("a", encoding="utf-8") as f_jsonl:
        for sample in buffer:
            f_jsonl.write(json.dumps(sample, ensure_ascii=False) + "\n")

    return header_written


def generate_param_plan(
    num_cases: int,
    out_path: Path,
    seed: Optional[int] = None,
    batch_flush_size: int = 1000,
    sleep_seconds_between_flush: int = 3,
    max_attempts: Optional[int] = None,
    case_id_offset: int = 0,
    external_seen: Optional[set] = None,
) -> None:
    """
    自动采样生成参数方案：
    - 目标 num_cases 个 unique 样本；
    - 每 batch_flush_size 条写盘并 sleep，避免长时间无响应；
    - 支持外部去重集合与 case_id 偏移，便于跨文件全局唯一与连续编号。
    """
    rng = random.Random(seed)
    seen: set[Tuple] = external_seen if external_seen is not None else set()
    buffer: List[Dict[str, Any]] = []
    samples_count = 0

    max_attempts_val = max_attempts or max(num_cases * 100, num_cases + 10)
    attempts = 0

    jsonl_path, csv_path = _resolve_output_paths(out_path)
    # 清理旧文件，避免追加到历史数据
    if csv_path.exists():
        csv_path.unlink()
    if jsonl_path.exists():
        jsonl_path.unlink()

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prefix = "[param_sampling]"
    print(f"{prefix} target={num_cases}, max_attempts={max_attempts_val}, seed={seed}")
    header_written = False

    while samples_count < num_cases and attempts < max_attempts_val:
        attempts += 1
        sample = generate_structural_sample(rng)

        key = flatten_sample_for_key(sample)
        if key in seen:
            continue

        case_id = case_id_offset + samples_count  # 全局连续的 case_id
        sample_out = {
            "case_id": case_id,
            "Topology": sample["Topology"],
            "Sizing": sample["Sizing"],
            "GroupMapping": sample["GroupMapping"],
        }
        seen.add(key)
        buffer.append(sample_out)
        samples_count += 1

        if samples_count % 1000 == 0 or samples_count == num_cases:
            print(f"{prefix} progress: {samples_count}/{num_cases} unique samples (attempts={attempts})", flush=True)

        # 每批写盘并休息，提升可感知性
        if len(buffer) >= batch_flush_size or samples_count == num_cases:
            header_written = _flush_buffer_to_files(buffer, csv_path, jsonl_path, header_written)
            print(
                f"{prefix} flush: {samples_count}/{num_cases} written to {csv_path.name} (attempts={attempts})",
                flush=True,
            )
            if sleep_seconds_between_flush > 0:
                print(f"{prefix} sleep {sleep_seconds_between_flush} seconds before continuing...", flush=True)
                time.sleep(sleep_seconds_between_flush)
            buffer.clear()

    if buffer:
        header_written = _flush_buffer_to_files(buffer, csv_path, jsonl_path, header_written)
        print(f"{prefix} flush: {samples_count}/{num_cases} written to {csv_path.name} (final buffer)", flush=True)

    if attempts >= max_attempts_val and samples_count < num_cases:
        print(
            f"[警告]{prefix} 达到最大尝试次数 max_attempts={max_attempts_val}，"
            f"仅生成 {samples_count}/{num_cases} 个唯一样本（attempts={attempts}）",
            flush=True,
        )
    else:
        print(f"{prefix} DONE: {samples_count}/{num_cases} unique samples, attempts={attempts}, file={csv_path}")


def generate_param_plan_multi_files(
    total_cases: int,
    out_dir: Path,
    num_files: int = 10,
    seed: Optional[int] = 42,
    batch_flush_size: int = 1000,
    sleep_seconds_between_flush: int = 3,
    max_attempts: Optional[int] = None,
) -> List[Path]:
    """
    生成多个方案文件，总数 total_cases，平均分配到 num_files 个 CSV。
    - 保证所有文件间 case_id 全局连续且唯一。
    - 共享去重集合，确保全局 unique。
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    paths: List[Path] = []
    base_name = "param_plan_auto_generated"

    # 计算分配：前 remainder 个文件多 1 个
    global_seen: set = set()
    current_offset = 0

    for idx in range(num_files):
        remaining_files = num_files - idx
        remaining_cases = max(total_cases - current_offset, 0)
        base = remaining_cases // remaining_files if remaining_files else 0
        remainder = remaining_cases % remaining_files if remaining_files else 0
        per_file = base + (1 if remainder > 0 else 0)
        if per_file <= 0:
            continue
        file_seed = seed + idx if seed is not None else None
        suffix = f"{idx + 1:02d}"
        csv_path = out_dir / f"{base_name}_{suffix}.csv"
        before = len(global_seen)

        generate_param_plan(
            num_cases=per_file,
            out_path=csv_path,
            seed=file_seed,
            batch_flush_size=batch_flush_size,
            sleep_seconds_between_flush=sleep_seconds_between_flush,
            max_attempts=max_attempts,
            case_id_offset=current_offset,
            external_seen=global_seen,
        )

        paths.append(csv_path)
        produced = len(global_seen) - before
        current_offset += produced

    return paths


def _get_value(flat: Dict[str, Any], keys: List[str]) -> Any:
    """Return the first non-empty value for candidate keys (case-insensitive)."""
    for key in keys:
        if key in flat and flat[key] not in ("", None):
            return flat[key]
        key_lower = str(key).lower()
        for candidate_key, candidate_val in flat.items():
            if str(candidate_key).lower() == key_lower and candidate_val not in ("", None):
                return candidate_val
    return None


def _to_int(value: Any, field_name: str, allow_none: bool = False, default: Optional[int] = None) -> int:
    if value is None or (isinstance(value, str) and not value.strip()):
        if allow_none:
            return default  # type: ignore[return-value]
        raise ValueError(f"[param_sampling] 字段 {field_name} 缺失或为空，无法解析参数方案")
    try:
        return int(float(value))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"[param_sampling] 字段 {field_name} 的值无法转换为整数: {value!r}") from exc


def _normalize_case_id(raw_case_id: Any, fallback_case_id: int) -> int:
    if raw_case_id is None or (isinstance(raw_case_id, str) and not raw_case_id.strip()):
        return fallback_case_id
    try:
        return int(raw_case_id)
    except (TypeError, ValueError):
        try:
            return int(float(raw_case_id))
        except (TypeError, ValueError):
            return fallback_case_id


def _normalize_structured_sample(sample: Dict[str, Any], fallback_case_id: int) -> Dict[str, Any]:
    """Normalize a structured sample that already contains Topology/Sizing."""
    if "Topology" not in sample or "Sizing" not in sample:
        raise ValueError("[param_sampling] 结构化方案缺少 Topology 或 Sizing 字段")

    topo_raw = sample["Topology"] or {}
    sizing_raw = sample["Sizing"] or {}

    topo = {
        "N_st": _to_int(_get_value(topo_raw, ["N_st", "n_st", "num_stories"]), "Topology.N_st"),
        "n_x": _to_int(_get_value(topo_raw, ["n_x", "Nx", "grid_x", "num_grid_x"]), "Topology.n_x"),
        "n_y": _to_int(_get_value(topo_raw, ["n_y", "Ny", "grid_y", "num_grid_y"]), "Topology.n_y"),
        "l_x": _to_int(_get_value(topo_raw, ["l_x", "l_x_mm", "Lx", "span_x"]), "Topology.l_x"),
        "l_y": _to_int(_get_value(topo_raw, ["l_y", "l_y_mm", "Ly", "span_y"]), "Topology.l_y"),
    }

    normalized_sizing: Dict[str, Dict[str, int]] = {}
    for gid in range(1, 4):
        group_key = f"Group{gid}"
        group_data = sizing_raw.get(group_key) or sizing_raw.get(group_key.lower()) if hasattr(sizing_raw, "get") else None
        if not isinstance(group_data, dict):
            raise ValueError(f"[param_sampling] 方案缺少或损坏的截面参数: {group_key}")
        normalized_sizing[group_key] = {
            f"C_G{gid}_Corner_b": _to_int(_get_value(group_data, [f"C_G{gid}_Corner_b"]), f"{group_key}.C_G{gid}_Corner_b"),
            f"C_G{gid}_Edge_b": _to_int(_get_value(group_data, [f"C_G{gid}_Edge_b"]), f"{group_key}.C_G{gid}_Edge_b"),
            f"C_G{gid}_Interior_b": _to_int(_get_value(group_data, [f"C_G{gid}_Interior_b"]), f"{group_key}.C_G{gid}_Interior_b"),
            f"B_G{gid}_Edge_b": _to_int(_get_value(group_data, [f"B_G{gid}_Edge_b"]), f"{group_key}.B_G{gid}_Edge_b"),
            f"B_G{gid}_Edge_h": _to_int(_get_value(group_data, [f"B_G{gid}_Edge_h"]), f"{group_key}.B_G{gid}_Edge_h"),
            f"B_G{gid}_Interior_b": _to_int(_get_value(group_data, [f"B_G{gid}_Interior_b"]), f"{group_key}.B_G{gid}_Interior_b"),
            f"B_G{gid}_Interior_h": _to_int(_get_value(group_data, [f"B_G{gid}_Interior_h"]), f"{group_key}.B_G{gid}_Interior_h"),
        }

    group_mapping = sample.get("GroupMapping") or split_story_groups(topo["N_st"])
    case_id = _normalize_case_id(sample.get("case_id") or sample.get("num"), fallback_case_id)
    return {"case_id": case_id, "Topology": topo, "Sizing": normalized_sizing, "GroupMapping": group_mapping}


def _sample_from_flat(flat: Dict[str, Any], fallback_case_id: int) -> Dict[str, Any]:
    """Build a sample from flat table-like data."""
    topo = {
        "N_st": _to_int(_get_value(flat, ["N_st", "n_st", "stories", "num_stories"]), "N_st"),
        "n_x": _to_int(_get_value(flat, ["n_x", "grid_x", "Nx"]), "n_x"),
        "n_y": _to_int(_get_value(flat, ["n_y", "grid_y", "Ny"]), "n_y"),
        "l_x": _to_int(_get_value(flat, ["l_x_mm", "l_x", "span_x", "Lx"]), "l_x"),
        "l_y": _to_int(_get_value(flat, ["l_y_mm", "l_y", "span_y", "Ly"]), "l_y"),
    }

    sizing: Dict[str, Dict[str, int]] = {}
    for gid in range(1, 4):
        sizing[f"Group{gid}"] = {
            f"C_G{gid}_Corner_b": _to_int(_get_value(flat, [f"C_G{gid}_Corner_b"]), f"C_G{gid}_Corner_b"),
            f"C_G{gid}_Edge_b": _to_int(_get_value(flat, [f"C_G{gid}_Edge_b"]), f"C_G{gid}_Edge_b"),
            f"C_G{gid}_Interior_b": _to_int(_get_value(flat, [f"C_G{gid}_Interior_b"]), f"C_G{gid}_Interior_b"),
            f"B_G{gid}_Edge_b": _to_int(_get_value(flat, [f"B_G{gid}_Edge_b"]), f"B_G{gid}_Edge_b"),
            f"B_G{gid}_Edge_h": _to_int(_get_value(flat, [f"B_G{gid}_Edge_h"]), f"B_G{gid}_Edge_h"),
            f"B_G{gid}_Interior_b": _to_int(_get_value(flat, [f"B_G{gid}_Interior_b"]), f"B_G{gid}_Interior_b"),
            f"B_G{gid}_Interior_h": _to_int(_get_value(flat, [f"B_G{gid}_Interior_h"]), f"B_G{gid}_Interior_h"),
        }

    case_id = _normalize_case_id(
        _get_value(flat, ["case_id", "num", "case", "case_no", "id"]),
        fallback_case_id,
    )
    return {"case_id": case_id, "Topology": topo, "Sizing": sizing, "GroupMapping": split_story_groups(topo["N_st"])}


def find_param_plan_file(plan_dir: Optional[Path] = None, prefix: str = PLAN_FILE_PREFIX) -> Optional[Path]:
    """Find a param_plan file in plan_dir with supported suffixes."""
    directory = plan_dir or _HERE
    if not directory.exists():
        return None

    for suffix in SUPPORTED_PLAN_SUFFIXES:
        for path in sorted(directory.glob(f"{prefix}*{suffix}")):
            if path.is_file():
                return path

    fallback = sorted(p for p in directory.glob(f"{prefix}*") if p.is_file())
    return fallback[0] if fallback else None


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
    csv_path = out_path
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADERS)
        for sample in samples:
            writer.writerow(_flatten_sample_row(sample))
    return csv_path


def load_param_plan(path: Path) -> List[Dict[str, Any]]:
    samples: List[Dict[str, Any]] = []

    def _append_normalized(raw_sample: Dict[str, Any], index_one_based: int) -> None:
        try:
            normalized = _normalize_structured_sample(raw_sample, index_one_based)
        except Exception:
            normalized = _sample_from_flat(raw_sample, index_one_based)
        samples.append(normalized)

    suffix = path.suffix.lower()
    if suffix in (".jsonl", ".ndjson", ".txt", ""):
        with path.open("r", encoding="utf-8") as f:
            for idx, line in enumerate(f, start=1):
                text = line.strip()
                if not text:
                    continue
                _append_normalized(json.loads(text), idx)
    elif suffix == ".json":
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            for idx, sample in enumerate(data, start=1):
                _append_normalized(sample, idx)
        elif isinstance(data, dict) and "cases" in data and isinstance(data["cases"], list):
            for idx, sample in enumerate(data["cases"], start=1):
                _append_normalized(sample, idx)
        elif isinstance(data, dict):
            _append_normalized(data, 1)
        else:
            raise ValueError(f"[param_sampling] 无法从 JSON 方案解析样本: {path}")
    elif suffix == ".csv":
        with path.open("r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for idx, row in enumerate(reader, start=1):
                if not any(row.values()):
                    continue
                _append_normalized(row, idx)
    elif suffix in (".xlsx", ".xls"):
        try:
            import pandas as pd
        except ImportError as exc:  # pragma: no cover - depends on env
            raise ImportError("[param_sampling] 读取 Excel 方案需要安装 pandas 和 openpyxl") from exc
        df = pd.read_excel(path)
        for idx, (_, row) in enumerate(df.iterrows(), start=1):
            data = {k: row[k] for k in df.columns}
            if not any(value not in (None, "") for value in data.values()):
                continue
            _append_normalized(data, idx)
    else:
        raise ValueError(f"[param_sampling] 不支持的方案文件类型: {path.suffix}")

    return samples


__all__ = [
    "DesignCaseConfig",
    "flatten_sample_for_key",
    "generate_param_plan",
    "generate_param_plan_multi_files",
    "generate_structural_sample",
    "load_param_plan",
    "find_param_plan_file",
    "sample_column_section",
    "sample_topology",
    "split_story_groups",
]


if __name__ == "__main__":
    example_cases = 100
    out_dir = _HERE / "param_plan_auto_generated"
    print(f"[param_sampling] 生成 {example_cases} 个唯一样本 → {out_dir}")
    generate_param_plan_multi_files(
        total_cases=example_cases,
        out_dir=out_dir,
        num_files=2,
        seed=42,
        batch_flush_size=200,
        sleep_seconds_between_flush=0,
        max_attempts=50000,
    )
    print(f"[param_sampling] 生成完成，目录: {out_dir}")
