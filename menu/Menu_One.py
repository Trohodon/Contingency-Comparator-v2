# menu/Menu_One.py
from __future__ import annotations

import os
import tempfile

# Optional: hide pygame support prompt
os.environ.setdefault("SUPPORT_PROMPT", "1")

import pygame  # noqa: E402


def _lock_path() -> str:
    return os.path.join(tempfile.gettempdir(), "menu_one.lock")


def _release_lock():
    try:
        os.remove(_lock_path())
    except Exception:
        pass


def main():
    # If launcher already wrote the lock, we overwrite with our PID
    try:
        with open(_lock_path(), "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
    except Exception:
        pass

    try:
        # --- your original program, just package imports ---
        from menu.core.settings import WINDOW_W, WINDOW_H, FPS, CAPTION, BG_COLOR
        from menu.core.utils import init_fonts
        from menu.world.game import TetrisGame
        from menu.gui.screens import ScreenManager

        pygame.init()
        init_fonts()  # IMPORTANT: prevents "font not initialized"

        pygame.display.set_caption(CAPTION)
        screen = pygame.display.set_mode((WINDOW_W, WINDOW_H))
        clock = pygame.time.Clock()

        game = TetrisGame()
        screens = ScreenManager(game)

        running = True
        while running:
            dt = clock.tick(FPS) / 1000.0

            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    running = False
                    break
                screens.handle_event(event)

            screens.update(dt)

            screen.fill(BG_COLOR)
            screens.draw(screen)
            pygame.display.flip()

        pygame.quit()

    finally:
        _release_lock()


if __name__ == "__main__":
    main()