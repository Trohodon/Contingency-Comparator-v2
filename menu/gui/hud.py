import pygame

from ..core.settings import (
    BOARD_H, CELL,
    BOARD_Y, SIDE_PANEL_X, SIDE_PANEL_W,
    PANEL_BG, GRID_COLOR, TEXT_COLOR, MUTED_TEXT, PIECE_COLORS
)
from ..core.utils import draw_text, draw_panel, lighten, darken


class HUD:
    def __init__(self):
        self.panel_rect = pygame.Rect(SIDE_PANEL_X, BOARD_Y, SIDE_PANEL_W, BOARD_H * CELL)

        # Mini preview sizing
        self.mini = 18
        self.box_size = 4 * self.mini
        self.box_pad = 10

        # Controls text spacing (matches utils font sizes)
        self.ctrl_line_h = 18

    def draw(self, screen, game):
        draw_panel(screen, self.panel_rect, PANEL_BG)

        x0 = self.panel_rect.x + 20
        y_top = self.panel_rect.y + 20
        y_bottom = self.panel_rect.bottom - 20

        # ---------- Controls area (reserve space at bottom) ----------
        ctrl_lines = [
            "←/→ : Move",
            "↓ : Soft drop",
            "Space : Hard drop",
            "Up/X : Rotate CW",
            "Z : Rotate CCW",
            "C : Hold",
            "P : Pause",
            "R : Restart (Game Over)",
        ]
        ctrl_title_h = 24
        ctrl_block_h = ctrl_title_h + len(ctrl_lines) * self.ctrl_line_h

        ctrl_y = y_bottom - ctrl_block_h
        draw_text(screen, "Controls", x0, ctrl_y, kind="med", color=MUTED_TEXT)
        cy = ctrl_y + ctrl_title_h
        for s in ctrl_lines:
            draw_text(screen, s, x0, cy, kind="small", color=MUTED_TEXT)
            cy += self.ctrl_line_h

        # Everything above this must fit into [y_top .. ctrl_y - gap]
        safe_bottom = ctrl_y - 16  # gap above controls

        # ---------- Header ----------
        y = y_top
        draw_text(screen, "TETRIS", x0, y, kind="big")
        y += 56

        draw_text(screen, f"Score: {game.score}", x0, y, kind="med", color=TEXT_COLOR); y += 28
        draw_text(screen, f"Level: {game.level}", x0, y, kind="med", color=TEXT_COLOR); y += 28
        draw_text(screen, f"Lines: {game.lines}", x0, y, kind="med", color=TEXT_COLOR); y += 36

        # ---------- Next + Hold (adaptive) ----------
        # We try to show up to 5 next pieces, but shrink the count until it fits.
        def height_for_next(n_next: int) -> int:
            next_title = 26
            next_boxes = n_next * self.box_size + max(0, n_next - 1) * self.box_pad
            hold_title = 26
            hold_box = self.box_size
            # plus section gaps
            return next_title + next_boxes + 26 + hold_title + hold_box

        max_next = 5
        while max_next > 2 and (y + height_for_next(max_next) > safe_bottom):
            max_next -= 1

        # Next section
        draw_text(screen, "Next", x0, y, kind="med", color=MUTED_TEXT)
        y += 26
        self._draw_queue(screen, game.queue, x0, y, count=max_next)
        y += max_next * self.box_size + max(0, max_next - 1) * self.box_pad
        y += 26

        # Hold section
        draw_text(screen, "Hold", x0, y, kind="med", color=MUTED_TEXT)
        y += 26
        self._draw_hold(screen, game.hold, x0, y)

    def _draw_queue(self, screen, queue, x, y, count=5):
        for i, kind in enumerate(queue[:count]):
            rect = pygame.Rect(x, y + i * (self.box_size + self.box_pad), self.box_size, self.box_size)
            pygame.draw.rect(screen, darken(PANEL_BG, 10), rect, border_radius=10)
            pygame.draw.rect(screen, GRID_COLOR, rect, width=2, border_radius=10)
            self._draw_mini_piece(screen, kind, rect)

    def _draw_hold(self, screen, hold_kind, x, y):
        rect = pygame.Rect(x, y, self.box_size, self.box_size)
        pygame.draw.rect(screen, darken(PANEL_BG, 10), rect, border_radius=10)
        pygame.draw.rect(screen, GRID_COLOR, rect, width=2, border_radius=10)
        if hold_kind:
            self._draw_mini_piece(screen, hold_kind, rect)

    def _draw_mini_piece(self, screen, kind, rect):
        layouts = {
            "I": [(0, 1), (1, 1), (2, 1), (3, 1)],
            "O": [(1, 1), (2, 1), (1, 2), (2, 2)],
            "T": [(1, 1), (0, 2), (1, 2), (2, 2)],
            "S": [(1, 1), (2, 1), (0, 2), (1, 2)],
            "Z": [(0, 1), (1, 1), (1, 2), (2, 2)],
            "J": [(0, 1), (0, 2), (1, 2), (2, 2)],
            "L": [(2, 1), (0, 2), (1, 2), (2, 2)],
        }
        cells = layouts.get(kind, [])
        color = PIECE_COLORS.get(kind, (200, 200, 200))

        ox = rect.x + 6
        oy = rect.y + 6
        for cx, cy in cells:
            r = pygame.Rect(ox + cx * self.mini, oy + cy * self.mini, self.mini - 2, self.mini - 2)
            pygame.draw.rect(screen, color, r, border_radius=4)
            pygame.draw.rect(screen, lighten(color, 30), r, width=2, border_radius=4)
