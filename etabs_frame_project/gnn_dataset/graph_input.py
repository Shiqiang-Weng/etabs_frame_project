#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Graph input export (nodes/edges) aligned with ETABS frame objects."""

from __future__ import annotations

import csv
import json
import math
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np

from common.config import MM_TO_M, SETTINGS
from common.dataset_paths import (
    BUCKET_SIZE,
    GNN_INPUT_EXT,
    INPUT_BUCKET_PREFIX,
    NUM_BUCKETS,
    build_bucket_dir,
    compute_bucket,
)

def _infer_story_from_z(z: float, story_height: float) -> int:
    if story_height <= 0:
        return 0
    return int(round(z / story_height))


def _distance(p1: Tuple[float, float, float], p2: Tuple[float, float, float]) -> float:
    return float(math.sqrt(sum((a - b) ** 2 for a, b in zip(p1, p2))))


def _section_dims(section_name: Optional[str], section_dims: Dict[str, Tuple[float, float]]) -> Tuple[float, float]:
    if section_name and section_name in section_dims:
        return section_dims[section_name]

    # fallback: parse like 300x600 from property name (interpreted as mm)
    if section_name:
        match = re.search(r"(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)", section_name)
        if match:
            b_mm, h_mm = float(match.group(1)), float(match.group(2))
            return b_mm * MM_TO_M, h_mm * MM_TO_M
    return 0.0, 0.0


def _section_map_from_design(design_cfg) -> Dict[str, Tuple[float, float]]:
    mapping: Dict[str, Tuple[float, float]] = {}
    for name, width, depth in design_cfg.iter_frame_section_definitions():
        mapping[name] = (float(width), float(depth))
    return mapping


def _safe_get_points(frame_obj, name: str) -> Optional[Tuple[str, str]]:
    try:
        ret = frame_obj.GetPoints(name)
        if isinstance(ret, tuple) and len(ret) >= 3:
            return str(ret[1]), str(ret[2])
    except Exception:
        return None
    return None


def _safe_get_section(frame_obj, name: str) -> Optional[str]:
    try:
        ret = frame_obj.GetSection(name)
        if isinstance(ret, tuple) and len(ret) >= 2:
            return str(ret[1])
    except Exception:
        return None
    return None


def _safe_get_point_coord(point_obj, point_name: str) -> Optional[Tuple[float, float, float]]:
    try:
        ret = point_obj.GetCoordCartesian(point_name)
        if isinstance(ret, tuple) and len(ret) >= 4 and ret[0] in (0, 1):
            return float(ret[1]), float(ret[2]), float(ret[3])
    except Exception:
        return None
    return None


def _infer_member_type(frame_name: str, p1: Tuple[float, float, float], p2: Tuple[float, float, float], section_name: Optional[str]) -> int:
    """Return 1 for column, 0 for beam."""
    name_upper = (frame_name or "").upper()
    if name_upper.startswith("COL"):
        return 1
    if name_upper.startswith("BEAM"):
        return 0
    if section_name:
        sec_upper = section_name.upper()
        if sec_upper.startswith("C_") or "COLUMN" in sec_upper:
            return 1
        if sec_upper.startswith("B_"):
            return 0
    dz = abs(p1[2] - p2[2])
    dx = abs(p1[0] - p2[0])
    dy = abs(p1[1] - p2[1])
    if dz > max(dx, dy) * 0.1:
        return 1
    return 0


def _calculate_inertia(width: float, depth: float) -> Tuple[float, float]:
    """Return (Ix, Iy) assuming rectangle width=t2, depth=t3 (m^4)."""
    if width <= 0 or depth <= 0:
        return 0.0, 0.0
    i22 = width * depth**3 / 12.0
    i33 = depth * width**3 / 12.0
    return i22, i33


def _node_key_from_coord(x: float, y: float, z: float) -> Tuple[int, int, int]:
    """Stable integer key (mm) to deduplicate nodes."""
    return (int(round(x / MM_TO_M)), int(round(y / MM_TO_M)), int(round(z / MM_TO_M)))


def _node_name_from_coord(x: float, y: float, z: float) -> str:
    kx, ky, kz = _node_key_from_coord(x, y, z)
    return f"P_X{kx}_Y{ky}_Z{kz}"


def _normalize_group_name_local(name: object) -> str:
    """Normalize vertical group name to Group1/2/3; fallback to Group3."""
    if name is None:
        return "Group3"
    text = str(name)
    digits = "".join(ch for ch in text if ch.isdigit())
    if digits:
        return f"Group{digits}"
    upper = text.upper()
    if upper.startswith("GROUP"):
        return text
    return "Group3"


def _column_position(i: int, j: int, max_i: int, max_j: int) -> str:
    if i in {0, max_i} and j in {0, max_j}:
        return "CORNER"
    if i in {0, max_i} or j in {0, max_j}:
        return "EDGE"
    return "INTERIOR"


def _beam_position(orientation: str, i: int, j: int, max_i: int, max_j: int) -> str:
    if orientation == "X":
        return "EDGE" if j in {0, max_j} else "INT"
    return "EDGE" if i in {0, max_i} else "INT"


def _iter_points_from_grid(grid) -> Iterable[Tuple[int, int, float, float]]:
    if hasattr(grid, "iter_points"):
        yield from grid.iter_points()
        return
    xs = getattr(grid, "x_coords", [])
    ys = getattr(grid, "y_coords", [])
    for i, x in enumerate(xs):
        for j, y in enumerate(ys):
            yield i, j, x, y


def _iter_beam_spans_x_from_grid(grid) -> Iterable[Tuple[int, int, float, float, float]]:
    xs = getattr(grid, "x_coords", [])
    ys = getattr(grid, "y_coords", [])
    for j, y in enumerate(ys):
        for i in range(max(len(xs) - 1, 0)):
            yield i, j, xs[i], xs[i + 1], y


def _iter_beam_spans_y_from_grid(grid) -> Iterable[Tuple[int, int, float, float, float]]:
    xs = getattr(grid, "x_coords", [])
    ys = getattr(grid, "y_coords", [])
    for i, x in enumerate(xs):
        for j in range(max(len(ys) - 1, 0)):
            yield i, j, x, ys[j], ys[j + 1]


def _story_group_safe(design_cfg, story_num: int) -> str:
    return _normalize_group_name_local(design_cfg.story_group(story_num))


def _build_design_frames(design_cfg) -> Tuple[List[Dict[str, object]], Dict[Tuple[int, int, int], Tuple[float, float, float]]]:
    """
    Build frame definitions directly from design configuration (no ETABS OAPI dependency).

    Returns:
        frames: list of {name, p1, p2, section, type, width, depth}
        nodes: dict key(mm tuple) -> coord(m)
    """
    frames: List[Dict[str, object]] = []
    nodes: Dict[Tuple[int, int, int], Tuple[float, float, float]] = {}

    grid = design_cfg.grid
    stories = design_cfg.storeys
    max_i = grid.num_x - 1
    max_j = grid.num_y - 1

    # Columns
    for story_idx in range(stories.num_storeys):
        story_num = story_idx + 1
        z_bottom = story_idx * stories.storey_height
        z_top = z_bottom + stories.storey_height

        for i, j, x_coord, y_coord in _iter_points_from_grid(design_cfg.grid):
            position = _column_position(i, j, max_i, max_j)
            group_name = _story_group_safe(design_cfg, story_num)
            section_name = design_cfg.column_section_name_for_story(story_num, position)
            width, depth, _ = design_cfg.column_dims_and_name(group_name, position)
            name = f"COL_X{i}_Y{j}_S{story_num}"
            p1 = (x_coord, y_coord, z_bottom)
            p2 = (x_coord, y_coord, z_top)
            k1 = _node_key_from_coord(*p1)
            k2 = _node_key_from_coord(*p2)
            nodes.setdefault(k1, p1)
            nodes.setdefault(k2, p2)
            frames.append(
                {
                    "name": name,
                    "p1": p1,
                    "p2": p2,
                    "section": section_name,
                    "type": 1,  # column
                    "width": width,
                    "depth": depth,
                }
            )

    # Beams
    for story_idx in range(stories.num_storeys):
        story_num = story_idx + 1
        z_top = (story_idx + 1) * stories.storey_height
        max_i = grid.num_x - 1
        max_j = grid.num_y - 1

        # X direction
        for i, j, x1, x2, y in _iter_beam_spans_x_from_grid(design_cfg.grid):
            position = _beam_position("X", i, j, max_i, max_j)
            group_name = _story_group_safe(design_cfg, story_num)
            beam_depth = design_cfg.beam_depth_for_story(story_num, position)
            z_beam_center = z_top - beam_depth / 2.0
            section_name = design_cfg.beam_section_name_for_story(story_num, position)
            width, depth, _ = design_cfg.beam_dims_and_name(group_name, position)
            name = f"BEAM_X_X{i}to{i + 1}_Y{j}_S{story_num}"
            p1 = (x1, y, z_beam_center)
            p2 = (x2, y, z_beam_center)
            k1 = _node_key_from_coord(*p1)
            k2 = _node_key_from_coord(*p2)
            nodes.setdefault(k1, p1)
            nodes.setdefault(k2, p2)
            frames.append(
                {
                    "name": name,
                    "p1": p1,
                    "p2": p2,
                    "section": section_name,
                    "type": 0,  # beam
                    "width": width,
                    "depth": depth,
                }
            )

        # Y direction
        for i, j, x, y1, y2 in _iter_beam_spans_y_from_grid(design_cfg.grid):
            position = _beam_position("Y", i, j, max_i, max_j)
            group_name = _story_group_safe(design_cfg, story_num)
            beam_depth = design_cfg.beam_depth_for_story(story_num, position)
            z_beam_center = z_top - beam_depth / 2.0
            section_name = design_cfg.beam_section_name_for_story(story_num, position)
            width, depth, _ = design_cfg.beam_dims_and_name(group_name, position)
            name = f"BEAM_Y_X{i}_Y{j}to{j + 1}_S{story_num}"
            p1 = (x, y1, z_beam_center)
            p2 = (x, y2, z_beam_center)
            k1 = _node_key_from_coord(*p1)
            k2 = _node_key_from_coord(*p2)
            nodes.setdefault(k1, p1)
            nodes.setdefault(k2, p2)
            frames.append(
                {
                    "name": name,
                    "p1": p1,
                    "p2": p2,
                    "section": section_name,
                    "type": 0,  # beam
                    "width": width,
                    "depth": depth,
                }
            )

    return frames, nodes


def export_case_graph_input(
    sap_model,
    design_cfg,
    frame_element_names: List[str],
    input_root: Path,
    bucket_size: int = BUCKET_SIZE,
    num_buckets: int = NUM_BUCKETS,
) -> Optional[Path]:
    """
    Export a single case graph (node/edge features) to an .npz file.
    Prefers deterministic geometry from design config; ETABS OAPI is not required.

    Returns:
        Path to the saved .npz file, or None if export was skipped/failed.
    """
    bucket = compute_bucket(design_cfg.case_id, bucket_size=bucket_size, num_buckets=num_buckets)
    bucket_dir = build_bucket_dir(Path(input_root), INPUT_BUCKET_PREFIX, bucket)
    bucket_dir.mkdir(parents=True, exist_ok=True)
    out_path = bucket_dir / f"case_{design_cfg.case_id}{GNN_INPUT_EXT}"
    if out_path.exists():
        print(f"[GNN] 图输入已存在，跳过重写: {out_path}")
        return out_path

    # Build geometry deterministically from design (no dependency on FrameObj.GetPoints)
    frames, node_map = _build_design_frames(design_cfg)

    pga_g = SETTINGS.response_spectrum.rs_base_accel_g
    pga_ms2 = pga_g * SETTINGS.response_spectrum.gravity_accel
    alpha_max = pga_g
    tg = SETTINGS.response_spectrum.rs_characteristic_period
    damping_ratio = SETTINGS.response_spectrum.rs_damping_ratio
    youngs_modulus = SETTINGS.materials.concrete_e_modulus
    fc_val = getattr(SETTINGS.materials, "concrete_fc", None)
    fc_mpa = float(fc_val) if fc_val is not None else 0.0

    node_names: List[str] = []
    node_coords: List[Tuple[float, float, float]] = []
    for key, coord in sorted(node_map.items()):
        node_names.append(_node_name_from_coord(*coord))
        node_coords.append(coord)
    node_index = {name: idx for idx, name in enumerate(node_names)}

    edges: List[Dict[str, object]] = []
    for frame in frames:
        p1 = frame["p1"]  # type: ignore[index]
        p2 = frame["p2"]  # type: ignore[index]
        length = _distance(p1, p2)
        width = float(frame["width"])  # type: ignore[call-overload]
        depth = float(frame["depth"])  # type: ignore[call-overload]
        area = width * depth if width > 0 and depth > 0 else 0.0
        ix, iy = _calculate_inertia(width, depth)
        edges.append(
            {
                "name": frame["name"],
                "points": (p1, p2),
                "section": frame["section"],
                "type": frame["type"],
                "width": width,
                "depth": depth,
                "area": area,
                "length": length,
                "ix": ix,
                "iy": iy,
            }
        )

    node_features: List[List[float]] = []
    for coord in node_coords:
        x, y, z = coord
        floor_num = _infer_story_from_z(z, design_cfg.storeys.storey_height)
        node_features.append([float(floor_num), float(x), float(y), float(z), pga_ms2, alpha_max, tg, damping_ratio])

    edge_index_pairs: List[Tuple[int, int]] = []
    edge_features: List[List[float]] = []
    edge_meta: List[Dict[str, object]] = []
    for edge in edges:
        p1, p2 = edge["points"]  # type: ignore[index]
        n1 = _node_name_from_coord(*p1)
        n2 = _node_name_from_coord(*p2)
        if n1 not in node_index or n2 not in node_index:
            continue
        start_idx = node_index[n1]
        end_idx = node_index[n2]
        edge_index_pairs.append((start_idx, end_idx))
        edge_features.append(
            [
                float(edge["type"]),  # T: 0 beam, 1 column
                float(edge["width"]),
                float(edge["depth"]),
                float(edge["area"]),
                float(edge["length"]),
                youngs_modulus,
                float(edge["ix"]),
                float(edge["iy"]),
                fc_mpa,
            ]
        )
        edge_meta.append(
            {
                "name": edge["name"],
                "section": edge["section"],
                "points": [n1, n2],
                "length": edge["length"],
                "type": edge["type"],
            }
        )

    node_features_arr = np.asarray(node_features, dtype=np.float32)
    if edge_features:
        edge_features_arr = np.asarray(edge_features, dtype=np.float32)
    else:
        edge_features_arr = np.zeros((0, 9), dtype=np.float32)
    edge_index_arr = np.asarray(edge_index_pairs, dtype=np.int64)
    if edge_index_arr.size > 0:
        edge_index_arr = edge_index_arr.T  # shape [2, num_edges]
    else:
        edge_index_arr = np.zeros((2, 0), dtype=np.int64)

    meta = {
        "case_id": design_cfg.case_id,
        "bucket": bucket.range_label,
        "materials": {
            "concrete_material_name": SETTINGS.materials.concrete_material_name,
            "E": youngs_modulus,
            "fc_MPa": fc_mpa,
        },
        "seismic": {
            "pga_g": pga_g,
            "pga_ms2": pga_ms2,
            "alpha_max_g": alpha_max,
            "tg_s": tg,
            "xi": damping_ratio,
        },
        "loads": {
            "finish_line_load_beam": SETTINGS.loads.default_finish_load_beam,
        },
        "node_name_to_index": node_index,
        "edge_names": [e["name"] for e in edges],
        "edge_sections": [e["section"] for e in edges],
        "feature_fields": {
            "node": ["F", "X", "Y", "Z", "PGA_ms2", "alpha_max_g", "Tg_s", "xi"],
            "edge": ["T", "b_m", "h_m", "A_m2", "L_m", "E", "Ix_m4", "Iy_m4", "fc_MPa"],
        },
        "units": {
            "length": "m",
            "PGA_ms2": "m/s^2",
            "alpha_max_g": "g",
            "fc": "MPa",
        },
        "edge_meta": edge_meta,
        "source": "design-derived",
    }

    np.savez_compressed(
        out_path,
        node_features=node_features_arr,
        edge_index=edge_index_arr,
        edge_features=edge_features_arr,
        x=node_features_arr,  # alias for GNN loaders
        edge_attr=edge_features_arr,  # alias for GNN loaders
        meta=json.dumps(meta, ensure_ascii=False),
    )
    file_size = out_path.stat().st_size if out_path.exists() else 0
    # Also dump CSVs for quick inspection
    nodes_csv = out_path.with_name(f"{out_path.stem}_nodes.csv")
    edges_csv = out_path.with_name(f"{out_path.stem}_edges.csv")
    with nodes_csv.open("w", newline="", encoding="utf-8-sig") as f_nodes:
        writer = csv.writer(f_nodes)
        writer.writerow(["node_id", "name", "F", "X", "Y", "Z", "PGA_ms2", "alpha_max_g", "Tg_s", "xi"])
        for idx, (name, feats) in enumerate(zip(node_names, node_features), start=0):
            writer.writerow([idx, name, *[f"{v:.6g}" for v in feats]])
    with edges_csv.open("w", newline="", encoding="utf-8-sig") as f_edges:
        writer = csv.writer(f_edges)
        writer.writerow(
            ["edge_id", "name", "type", "section", "start_node", "end_node", "b_m", "h_m", "A_m2", "L_m", "E", "Ix_m4", "Iy_m4", "fc_MPa"]
        )
        for idx, (edge, feats, (s, t)) in enumerate(zip(edge_meta, edge_features, edge_index_pairs), start=0):
            writer.writerow(
                [
                    idx,
                    edge["name"],
                    edge["type"],
                    edge["section"],
                    s,
                    t,
                    *[f"{v:.6g}" for v in feats],
                ]
            )
    print(f"[GNN] CSV 导出: {nodes_csv.name}, {edges_csv.name}")
    print(
        f"[GNN] case {design_cfg.case_id} 图输入已保存: {out_path} "
        f"(nodes={len(node_names)}, edges={len(edge_index_pairs)}, size={file_size} bytes)"
    )
    return out_path


def extract_gnn_features(
    sap_model,
    design_cfg,
    frame_element_names: List[str],
    input_root: Path,
    bucket_size: int = BUCKET_SIZE,
    num_buckets: int = NUM_BUCKETS,
) -> Optional[Path]:
    """
    Thin wrapper kept for clarity: export graph input for a case into bucketed input folders.
    """
    return export_case_graph_input(
        sap_model=sap_model,
        design_cfg=design_cfg,
        frame_element_names=frame_element_names,
        input_root=input_root,
        bucket_size=bucket_size,
        num_buckets=num_buckets,
    )


__all__ = ["export_case_graph_input"]
