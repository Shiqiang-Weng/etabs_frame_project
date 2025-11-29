#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Grid and story layout definitions for geometry modeling.
"""

from dataclasses import dataclass
from typing import Dict, Iterable, List, Tuple

from config import (
    NUM_GRID_LINES_X,
    NUM_GRID_LINES_Y,
    SPACING_X,
    SPACING_Y,
    NUM_STORIES,
    TYPICAL_STORY_HEIGHT,
    BOTTOM_STORY_HEIGHT,
    FRAME_BEAM_HEIGHT,
)


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
    typical_height: float
    bottom_height: float
    beam_height: float

    def iter_story_bounds(self) -> Iterable[Tuple[int, float, float]]:
        z = 0.0
        for idx in range(self.num_stories):
            height = self.typical_height if idx > 0 else self.bottom_height
            z_bottom = z
            z_top = z + height
            z = z_top
            yield idx + 1, z_bottom, z_top

    def story_top_elevations(self) -> Dict[int, float]:
        tops: Dict[int, float] = {}
        for story_num, _, z_top in self.iter_story_bounds():
            tops[story_num] = z_top
        return tops


def default_grid_config() -> GridConfig:
    return GridConfig(NUM_GRID_LINES_X, NUM_GRID_LINES_Y, SPACING_X, SPACING_Y)


def default_story_config() -> StoryConfig:
    return StoryConfig(NUM_STORIES, TYPICAL_STORY_HEIGHT, BOTTOM_STORY_HEIGHT, FRAME_BEAM_HEIGHT)


__all__ = ["GridConfig", "StoryConfig", "default_grid_config", "default_story_config"]
