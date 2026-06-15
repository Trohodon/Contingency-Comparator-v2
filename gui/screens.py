import pygame

from ..core.settings import (
    BOARD_W, BOARD_H, HIDDEN_ROWS, CELL,
    BOARD_X, BOARD_Y, PANEL_BG, GRID_COLOR, TEXT_COLOR, MUTED_TEXT, PIECE_COLORS
)
from ..core.utils import draw_text, draw_panel, lighten, darken
from .hud import HUD


class ScreenManager:
    def __init__(self, game):
        self.game = game
        self.hud = HUD()

    def handle_event(self, event):
        if event.type == pygame.KEYDOWN:
            if self.game.state == "gameover" and event.key == pygame.K_r:
                self.game.reset()
                return

            self.game.handle_keydown(event.key)

        elif event.type == pygame.KEYUP:
            self.game.handle_keyup(event.key)

    def update(self, dt: float):
        self.game.update(dt)

    def draw(self, screen):
        self._draw_board_area(screen)
        self.hud.draw(screen, self.game)

        if self.game.state == "pause":
            self._draw_overlay(screen, "PAUSED", "Press P to resume")
        elif self.game.state == "gameover":
            self._draw_overlay(screen, "GAME OVER", "Press R to restart")

    def _draw_board_area(self, screen):
        # board panel background
        board_px_w = BOARD_W * CELL
        board_px_h = BOARD_H * CELL
        board_rect = pygame.Rect(BOARD_X, BOARD_Y, board_px_w, board_px_h)
        draw_panel(screen, board_rect.inflate(24, 24), PANEL_BG)
        pygame.draw.rect(screen, darken(PANEL_BG, 10), board_rect)

        # grid
        for x in range(BOARD_W + 1):
            px = BOARD_X + x * CELL
            pygame.draw.line(screen, GRID_COLOR, (px, BOARD_Y), (px, BOARD_Y + board_px_h), 1)
        for y in range(BOARD_H + 1):
            py = BOARD_Y + y * CELL
            pygame.draw.line(screen, GRID_COLOR, (BOARD_X, py), (BOARD_X + board_px_w, py), 1)

        # locked blocks (skip hidden rows)
        for gy in range(HIDDEN_ROWS, BOARD_H + HIDDEN_ROWS):
            for gx in range(BOARD_W):
                kind = self.game.board.grid[gy][gx]
                if kind:
                    self._draw_cell(screen, gx, gy - HIDDEN_ROWS, kind, solid=True)

        # ghost
        if self.game.state == "play":
            ghost = self.game.ghost_cells()
            for (gx, gy) in ghost:
                vy = gy - HIDDEN_ROWS
                if vy >= 0:
                    self._draw_cell(screen, gx, vy, self.game.active.kind, solid=False)

        # active piece
        for (gx, gy) in self.game.active.cells():
            vy = gy - HIDDEN_ROWS
            if vy >= 0:
                self._draw_cell(screen, gx, vy, self.game.active.kind, solid=True)

    def _draw_cell(self, screen, gx, gy, kind, solid=True):
        color = PIECE_COLORS.get(kind, (200, 200, 200))
        x = BOARD_X + gx * CELL
        y = BOARD_Y + gy * CELL
        r = pygame.Rect(x + 2, y + 2, CELL - 4, CELL - 4)

        if solid:
            pygame.draw.rect(screen, color, r, border_radius=6)
            pygame.draw.rect(screen, lighten(color, 30), r, width=2, border_radius=6)
        else:
            # ghost outline
            pygame.draw.rect(screen, color, r, width=2, border_radius=6)

    def _draw_overlay(self, screen, title, subtitle):
        w, h = screen.get_size()
        overlay = pygame.Surface((w, h), pygame.SRCALPHA)
        overlay.fill((0, 0, 0, 140))
        screen.blit(overlay, (0, 0))

        rect = pygame.Rect(0, 0, 520, 220)
        rect.center = (w // 2, h // 2)
        draw_panel(screen, rect, (20, 24, 38), border=(60, 70, 100))

        draw_text(screen, title, rect.centerx, rect.y + 55, kind="big", align="center")
        draw_text(screen, subtitle, rect.centerx, rect.y + 120, kind="med", color=MUTED_TEXT, align="center")
