# menu/Menu_One.py

import pygame

from menu.core.settings import WINDOW_W, WINDOW_H, FPS, CAPTION, BG_COLOR
from menu.core.utils import init_fonts
from menu.world.game import TetrisGame
from menu.gui.screens import ScreenManager


def main():
    pygame.init()
    init_fonts()  # prevents "font not initialized"

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


if __name__ == "__main__":
    main()