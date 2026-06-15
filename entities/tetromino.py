from __future__ import annotations
from dataclasses import dataclass
from typing import List, Tuple, Dict

# 4x4 matrices for each piece orientation (0, R, 2, L)
# We'll use 4x4 for all pieces, including O, to simplify rotation math.

SHAPES: Dict[str, List[List[str]]] = {
    "I": [
        [
            "....",
            "IIII",
            "....",
            "....",
        ],
        [
            "..I.",
            "..I.",
            "..I.",
            "..I.",
        ],
        [
            "....",
            "....",
            "IIII",
            "....",
        ],
        [
            ".I..",
            ".I..",
            ".I..",
            ".I..",
        ],
    ],
    "O": [
        [
            "....",
            ".OO.",
            ".OO.",
            "....",
        ],
        [
            "....",
            ".OO.",
            ".OO.",
            "....",
        ],
        [
            "....",
            ".OO.",
            ".OO.",
            "....",
        ],
        [
            "....",
            ".OO.",
            ".OO.",
            "....",
        ],
    ],
    "T": [
        [
            "....",
            ".TTT",
            "..T.",
            "....",
        ],
        [
            "....",
            "..T.",
            ".TT.",
            "..T.",
        ],
        [
            "....",
            "..T.",
            ".TTT",
            "....",
        ],
        [
            "....",
            ".T..",
            ".TT.",
            ".T..",
        ],
    ],
    "S": [
        [
            "....",
            "..SS",
            ".SS.",
            "....",
        ],
        [
            "....",
            ".S..",
            ".SS.",
            "..S.",
        ],
        [
            "....",
            "..SS",
            ".SS.",
            "....",
        ],
        [
            "....",
            ".S..",
            ".SS.",
            "..S.",
        ],
    ],
    "Z": [
        [
            "....",
            ".ZZ.",
            "..ZZ",
            "....",
        ],
        [
            "....",
            "..Z.",
            ".ZZ.",
            ".Z..",
        ],
        [
            "....",
            ".ZZ.",
            "..ZZ",
            "....",
        ],
        [
            "....",
            "..Z.",
            ".ZZ.",
            ".Z..",
        ],
    ],
    "J": [
        [
            "....",
            "JJJ.",
            "...J",
            "....",
        ],
        [
            "..J.",
            "..J.",
            "..J.",
            ".J..",
        ],
        [
            "....",
            "J...",
            ".JJJ",
            "....",
        ],
        [
            "..J.",
            ".J..",
            ".J..",
            ".J..",
        ],
    ],
    "L": [
        [
            "....",
            ".LLL",
            ".L..",
            "....",
        ],
        [
            "....",
            ".LL.",
            "..L.",
            "..L.",
        ],
        [
            "....",
            "...L",
            ".LLL",
            "....",
        ],
        [
            "....",
            ".L..",
            ".L..",
            ".LL.",
        ],
    ],
}

# Basic SRS-ish wall kicks for JLSTZ and I pieces.
# Each entry is a list of (dx, dy) tests applied in order.
# Keys: (from_rot, to_rot)
JLSTZ_KICKS = {
    (0, 1): [(0, 0), (-1, 0), (-1, +1), (0, -2), (-1, -2)],
    (1, 0): [(0, 0), (+1, 0), (+1, -1), (0, +2), (+1, +2)],
    (1, 2): [(0, 0), (+1, 0), (+1, -1), (0, +2), (+1, +2)],
    (2, 1): [(0, 0), (-1, 0), (-1, +1), (0, -2), (-1, -2)],
    (2, 3): [(0, 0), (+1, 0), (+1, +1), (0, -2), (+1, -2)],
    (3, 2): [(0, 0), (-1, 0), (-1, -1), (0, +2), (-1, +2)],
    (3, 0): [(0, 0), (-1, 0), (-1, -1), (0, +2), (-1, +2)],
    (0, 3): [(0, 0), (+1, 0), (+1, +1), (0, -2), (+1, -2)],
}

I_KICKS = {
    (0, 1): [(0, 0), (-2, 0), (+1, 0), (-2, -1), (+1, +2)],
    (1, 0): [(0, 0), (+2, 0), (-1, 0), (+2, +1), (-1, -2)],
    (1, 2): [(0, 0), (-1, 0), (+2, 0), (-1, +2), (+2, -1)],
    (2, 1): [(0, 0), (+1, 0), (-2, 0), (+1, -2), (-2, +1)],
    (2, 3): [(0, 0), (+2, 0), (-1, 0), (+2, +1), (-1, -2)],
    (3, 2): [(0, 0), (-2, 0), (+1, 0), (-2, -1), (+1, +2)],
    (3, 0): [(0, 0), (+1, 0), (-2, 0), (+1, -2), (-2, +1)],
    (0, 3): [(0, 0), (-1, 0), (+2, 0), (-1, +2), (+2, -1)],
}


@dataclass
class Tetromino:
    kind: str
    x: int
    y: int
    rot: int = 0  # 0..3

    def cells(self, rot_override: int | None = None, pos_override: Tuple[int, int] | None = None):
        r = self.rot if rot_override is None else rot_override
        px, py = (self.x, self.y) if pos_override is None else pos_override
        shape = SHAPES[self.kind][r]
        out = []
        for row in range(4):
            for col in range(4):
                if shape[row][col] != ".":
                    out.append((px + col, py + row))
        return out

    def rotated(self, direction: int) -> int:
        # direction: +1 (CW), -1 (CCW)
        return (self.rot + direction) % 4

    def kick_tests(self, to_rot: int) -> List[Tuple[int, int]]:
        if self.kind == "O":
            return [(0, 0)]
        if self.kind == "I":
            return I_KICKS.get((self.rot, to_rot), [(0, 0)])
        return JLSTZ_KICKS.get((self.rot, to_rot), [(0, 0)])
