from __future__ import annotations
from typing import Optional, List, Tuple

from ..core.settings import BOARD_W, BOARD_H, HIDDEN_ROWS


class Board:
    """
    Stores locked blocks in a grid: (BOARD_H + HIDDEN_ROWS) rows by BOARD_W cols.
    Each cell holds a piece kind string (e.g. "T") or None.
    """
    def __init__(self):
        self.rows = BOARD_H + HIDDEN_ROWS
        self.cols = BOARD_W
        self.grid: List[List[Optional[str]]] = [[None for _ in range(self.cols)] for _ in range(self.rows)]

    def in_bounds(self, x: int, y: int) -> bool:
        return 0 <= x < self.cols and 0 <= y < self.rows

    def cell(self, x: int, y: int) -> Optional[str]:
        if not self.in_bounds(x, y):
            return None
        return self.grid[y][x]

    def is_empty(self, x: int, y: int) -> bool:
        return self.in_bounds(x, y) and self.grid[y][x] is None

    def can_place(self, cells: List[Tuple[int, int]]) -> bool:
        for x, y in cells:
            # allow y < 0? In our model y starts in hidden buffer, so no need.
            if not self.in_bounds(x, y):
                return False
            if self.grid[y][x] is not None:
                return False
        return True

    def lock(self, kind: str, cells: List[Tuple[int, int]]):
        for x, y in cells:
            if self.in_bounds(x, y):
                self.grid[y][x] = kind

    def clear_lines(self) -> int:
        """Clears full rows (including hidden rows if somehow filled). Returns number cleared."""
        new_grid = []
        cleared = 0
        for y in range(self.rows):
            if all(self.grid[y][x] is not None for x in range(self.cols)):
                cleared += 1
            else:
                new_grid.append(self.grid[y])

        # add empty rows at top
        while len(new_grid) < self.rows:
            new_grid.insert(0, [None for _ in range(self.cols)])

        self.grid = new_grid
        return cleared

    def is_game_over(self) -> bool:
        """Game over if any blocks exist in the hidden spawn area."""
        for y in range(HIDDEN_ROWS):
            for x in range(self.cols):
                if self.grid[y][x] is not None:
                    return True
        return False
