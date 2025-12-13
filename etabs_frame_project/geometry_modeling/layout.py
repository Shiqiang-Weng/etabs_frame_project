#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Grid and story layout definitions for geometry modeling."""

from dataclasses import dataclass
from typing import Dict, Iterable, List, Tuple, TYPE_CHECKING

from common import config
from common.config import DEFAULT_DESIGN_CONFIG, DesignConfig, design_config_from_case

if TYPE_CHECKING:
    from parametric_model.param_sampling import DesignCaseConfig


@dataclass(frozen=True)
class GridConfig:
    num_x: int
    num_y: int
    spacing_x: float
    spacing_y: float

    @property
    def x_coords(self) -> List[float]:
        return [i * self.spacing_x for i in range(self.num_x)]

    @property
    def y_coords(self) -> List[float]:
        return [j * self.spacing_y for j in range(self.num_y)]

    def iter_points(self) -> Iterable[Tuple[int, int, float, float]]:
        for i, x in enumerate(self.x_coords):
            for j, y in enumerate(self.y_coords):
                yield i, j, x, y

    def iter_beam_spans_x(self) -> Iterable[Tuple[int, int, float, float, float]]:
        xs = self.x_coords
        ys = self.y_coords
        for j, y in enumerate(ys):
            for i in range(len(xs) - 1):
                yield i, j, xs[i], xs[i + 1], y

    def iter_beam_spans_y(self) -> Iterable[Tuple[int, int, float, float, float]]:
        xs = self.x_coords
        ys = self.y_coords
        for i, x in enumerate(xs):
            for j in range(len(ys) - 1):
                yield i, j, x, ys[j], ys[j + 1]

    def iter_slab_panels(self) -> Iterable[Tuple[int, int, Tuple[float, float], Tuple[float, float]]]:
        xs = self.x_coords
        ys = self.y_coords
        for i in range(len(xs) - 1):
            for j in range(len(ys) - 1):
                yield i, j, (xs[i], xs[i + 1]), (ys[j], ys[j + 1])


@dataclass(frozen=True)
class StoryConfig:
    num_stories: int
    story_height: float

    def iter_story_bounds(self) -> Iterable[Tuple[int, float, float]]:
        for idx in range(self.num_stories):
            z_bottom = idx * self.story_height
            z_top = z_bottom + self.story_height
            yield idx + 1, z_bottom, z_top

    def story_top_elevations(self) -> Dict[int, float]:
        tops: Dict[int, float] = {}
        for story_num, _, z_top in self.iter_story_bounds():
            tops[story_num] = z_top
        return tops


def default_grid_config() -> GridConfig:
    grid = DEFAULT_DESIGN_CONFIG.grid
    return GridConfig(grid.num_x, grid.num_y, grid.spacing_x, grid.spacing_y)


def default_story_config() -> StoryConfig:
    storeys = DEFAULT_DESIGN_CONFIG.storeys
    return StoryConfig(storeys.num_storeys, storeys.storey_height)


def grid_config_from_design(design: "DesignCaseConfig") -> GridConfig:
    design_cfg: DesignConfig = design_config_from_case(design)
    grid = design_cfg.grid
    return GridConfig(
        num_x=grid.num_x,
        num_y=grid.num_y,
        spacing_x=grid.spacing_x,
        spacing_y=grid.spacing_y,
    )


def story_config_from_design(design: "DesignCaseConfig") -> StoryConfig:
    design_cfg: DesignConfig = design_config_from_case(design)
    storeys = design_cfg.storeys
    return StoryConfig(num_stories=storeys.num_storeys, story_height=storeys.storey_height)


__all__ = [
    "GridConfig",
    "StoryConfig",
    "default_grid_config",
    "default_story_config",
    "grid_config_from_design",
    "story_config_from_design",
]
