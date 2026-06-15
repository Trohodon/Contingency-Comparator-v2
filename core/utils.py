import pygame
from .settings import TEXT_COLOR, FONT_NAME, FONT_BIG, FONT_MED, FONT_SMALL

_FONT_CACHE = None


def init_fonts():
    """
    Call after pygame.init().
    Safe to call multiple times.
    """
    global _FONT_CACHE

    if not pygame.font.get_init():
        pygame.font.init()

    if _FONT_CACHE is None:
        _FONT_CACHE = {
            "big": pygame.font.Font(FONT_NAME, FONT_BIG),
            "med": pygame.font.Font(FONT_NAME, FONT_MED),
            "small": pygame.font.Font(FONT_NAME, FONT_SMALL),
        }


def get_font(kind="med") -> pygame.font.Font:
    global _FONT_CACHE
    if _FONT_CACHE is None:
        init_fonts()
    return _FONT_CACHE.get(kind, _FONT_CACHE["med"])


def draw_text(surface, text, x, y, kind="med", color=TEXT_COLOR, align="topleft"):
    font = get_font(kind)
    img = font.render(str(text), True, color)
    rect = img.get_rect()
    setattr(rect, align, (x, y))
    surface.blit(img, rect)
    return rect


def draw_panel(surface, rect, bg, border=None):
    pygame.draw.rect(surface, bg, rect, border_radius=12)
    if border:
        pygame.draw.rect(surface, border, rect, width=2, border_radius=12)


def lighten(color, amt=30):
    r, g, b = color
    return (min(255, r + amt), min(255, g + amt), min(255, b + amt))


def darken(color, amt=30):
    r, g, b = color
    return (max(0, r - amt), max(0, g - amt), max(0, b - amt))
