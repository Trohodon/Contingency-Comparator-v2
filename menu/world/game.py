from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional, Tuple

import pygame

from ..core.settings import (
    BOARD_W, BOARD_H, HIDDEN_ROWS,
    GRAVITY_START, GRAVITY_FACTOR, GRAVITY_MIN,
    LOCK_DELAY,
    SCORE_LINE_CLEAR,
    SOFT_DROP_SCORE_PER_CELL, HARD_DROP_SCORE_PER_CELL,
    LINES_PER_LEVEL,
    MOVE_REPEAT_DELAY, MOVE_REPEAT_RATE
)
from ..core.timing import Timer
from ..entities.tetromino import Tetromino
from .board import Board
from .bag import Bag7


@dataclass
class InputState:
    left: bool = False
    right: bool = False
    down: bool = False


class TetrisGame:
    def __init__(self):
        self.reset()

    def reset(self):
        self.board = Board()
        self.bag = Bag7()
        self.queue: List[str] = [self.bag.next() for _ in range(5)]
        self.hold: Optional[str] = None
        self.hold_used = False

        self.score = 0
        self.lines = 0
        self.level = 0

        self.state = "play"  # play, pause, gameover

        self.active = self._spawn_piece()
        self.lock_timer = 0.0

        # input repeat timers (for left/right held)
        self.input_state = InputState()
        self.lr_repeat_dir = 0  # -1 left, +1 right
        self.lr_delay = 0.0
        self.lr_rate = 0.0

        self.gravity_timer = Timer(self.gravity_seconds())

    def gravity_seconds(self) -> float:
        g = GRAVITY_START * (GRAVITY_FACTOR ** self.level)
        return max(GRAVITY_MIN, g)

    def _spawn_piece(self) -> Tetromino:
        kind = self.queue.pop(0)
        self.queue.append(self.bag.next())

        # Spawn near top center in hidden area
        # 4x4 piece -> x around center - 2
        x = (BOARD_W // 2) - 2
        y = 0  # inside hidden rows buffer
        piece = Tetromino(kind=kind, x=x, y=y, rot=0)

        # If can't place at spawn => game over
        if not self.board.can_place(piece.cells()):
            self.state = "gameover"
        return piece

    def _try_move(self, dx: int, dy: int) -> bool:
        if self.state != "play":
            return False
        new_pos = (self.active.x + dx, self.active.y + dy)
        cells = self.active.cells(pos_override=new_pos)
        if self.board.can_place(cells):
            self.active.x += dx
            self.active.y += dy
            return True
        return False

    def _try_rotate(self, direction: int) -> bool:
        if self.state != "play":
            return False
        to_rot = self.active.rotated(direction)
        for kx, ky in self.active.kick_tests(to_rot):
            new_pos = (self.active.x + kx, self.active.y + ky)
            cells = self.active.cells(rot_override=to_rot, pos_override=new_pos)
            if self.board.can_place(cells):
                self.active.rot = to_rot
                self.active.x += kx
                self.active.y += ky
                return True
        return False

    def _hard_drop(self):
        if self.state != "play":
            return
        dropped = 0
        while self._try_move(0, 1):
            dropped += 1
        if dropped > 0:
            self.score += dropped * HARD_DROP_SCORE_PER_CELL
        self._lock_piece()

    def _soft_drop_step(self) -> bool:
        if self._try_move(0, 1):
            self.score += SOFT_DROP_SCORE_PER_CELL
            return True
        return False

    def _lock_piece(self):
        # lock into board
        self.board.lock(self.active.kind, self.active.cells())
        cleared = self.board.clear_lines()

        if cleared > 0:
            base = SCORE_LINE_CLEAR.get(cleared, 0)
            self.score += base * (self.level + 1)
            self.lines += cleared
            self.level = self.lines // LINES_PER_LEVEL
            self.gravity_timer.set_threshold(self.gravity_seconds())

        # spawn next
        self.hold_used = False
        self.lock_timer = 0.0
        self.active = self._spawn_piece()

        if self.board.is_game_over():
            self.state = "gameover"

    def hold_piece(self):
        if self.state != "play":
            return
        if self.hold_used:
            return
        self.hold_used = True

        cur = self.active.kind
        if self.hold is None:
            self.hold = cur
            self.active = self._spawn_piece()
        else:
            swap = self.hold
            self.hold = cur
            # respawn swap as new active at spawn location
            x = (BOARD_W // 2) - 2
            y = 0
            self.active = Tetromino(kind=swap, x=x, y=y, rot=0)
            if not self.board.can_place(self.active.cells()):
                self.state = "gameover"

    def ghost_cells(self) -> List[Tuple[int, int]]:
        # copy position and drop until collision
        x, y = self.active.x, self.active.y
        rot = self.active.rot
        while True:
            test_cells = self.active.cells(rot_override=rot, pos_override=(x, y + 1))
            if self.board.can_place(test_cells):
                y += 1
            else:
                break
        return self.active.cells(rot_override=rot, pos_override=(x, y))

    def set_pause(self, paused: bool):
        if self.state == "gameover":
            return
        self.state = "pause" if paused else "play"

    def toggle_pause(self):
        if self.state == "play":
            self.set_pause(True)
        elif self.state == "pause":
            self.set_pause(False)

    # ---------------- Input Handling ----------------

    def handle_keydown(self, key: int):
        if key == pygame.K_p:
            self.toggle_pause()
            return

        if self.state != "play":
            return

        if key in (pygame.K_LEFT, pygame.K_a):
            self.input_state.left = True
            self._try_move(-1, 0)
            self.lr_repeat_dir = -1
            self.lr_delay = 0.0
            self.lr_rate = 0.0

        elif key in (pygame.K_RIGHT, pygame.K_d):
            self.input_state.right = True
            self._try_move(+1, 0)
            self.lr_repeat_dir = +1
            self.lr_delay = 0.0
            self.lr_rate = 0.0

        elif key in (pygame.K_DOWN, pygame.K_s):
            self.input_state.down = True
            # no immediate move here; update loop will do faster gravity

        elif key in (pygame.K_UP, pygame.K_x):
            self._try_rotate(+1)

        elif key == pygame.K_z:
            self._try_rotate(-1)

        elif key == pygame.K_SPACE:
            self._hard_drop()

        elif key == pygame.K_c:
            self.hold_piece()

    def handle_keyup(self, key: int):
        if key in (pygame.K_LEFT, pygame.K_a):
            self.input_state.left = False
            if self.lr_repeat_dir == -1:
                self.lr_repeat_dir = 0

        elif key in (pygame.K_RIGHT, pygame.K_d):
            self.input_state.right = False
            if self.lr_repeat_dir == +1:
                self.lr_repeat_dir = 0

        elif key in (pygame.K_DOWN, pygame.K_s):
            self.input_state.down = False

    # ---------------- Update Loop ----------------

    def update(self, dt: float):
        if self.state != "play":
            return

        # handle left/right held repeat
        if self.lr_repeat_dir != 0:
            self.lr_delay += dt
            if self.lr_delay >= MOVE_REPEAT_DELAY:
                self.lr_rate += dt
                while self.lr_rate >= MOVE_REPEAT_RATE:
                    self.lr_rate -= MOVE_REPEAT_RATE
                    moved = self._try_move(self.lr_repeat_dir, 0)
                    if not moved:
                        break

        # gravity timing (soft drop speeds it up)
        gravity_threshold = self.gravity_seconds()
        if self.input_state.down:
            gravity_threshold = max(0.01, gravity_threshold / 15.0)
        self.gravity_timer.set_threshold(gravity_threshold)

        fell = False
        if self.gravity_timer.tick(dt):
            fell = self._try_move(0, 1)

        # lock delay logic: if on ground, start timer, lock when exceeded
        if not self._can_fall_one():
            self.lock_timer += dt
            if self.lock_timer >= LOCK_DELAY:
                self._lock_piece()
        else:
            self.lock_timer = 0.0

    def _can_fall_one(self) -> bool:
        test = self.active.cells(pos_override=(self.active.x, self.active.y + 1))
        return self.board.can_place(test)
