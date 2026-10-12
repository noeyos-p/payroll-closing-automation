"""Format-agnostic 2D grid so the same block-walking/value-matching engine
(core/parsing.py, core/heuristic_match.py) works over an Excel worksheet or
an HWP table without caring which. An openpyxl Worksheet already exposes
`.max_row`/`.max_column`/`.cell(row, column).value`, so it satisfies this
shape directly — `Grid` below is only needed for sources (like HWP) that
don't come from openpyxl."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class GridCell:
    value: object


class Grid:
    """A plain in-memory 2D grid, 1-indexed like openpyxl (row=1, column=1
    is the top-left cell)."""

    def __init__(self, rows: list[list[object]]):
        self._rows = rows
        self.max_row = len(rows)
        self.max_column = max((len(r) for r in rows), default=0)

    def cell(self, row: int, column: int) -> GridCell:
        r, c = row - 1, column - 1
        if 0 <= r < len(self._rows) and 0 <= c < len(self._rows[r]):
            return GridCell(self._rows[r][c])
        return GridCell(None)
