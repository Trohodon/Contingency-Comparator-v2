# -------- Window / timing --------
FPS = 60
CAPTION = "Tetris (Modular Python)"
WINDOW_W = 980
WINDOW_H = 720

# -------- Board geometry --------
BOARD_W = 10
BOARD_H = 20
HIDDEN_ROWS = 2
CELL = 30

# Layout positions (pixel)
BOARD_X = 280
BOARD_Y = 60

# UI panels
SIDE_PANEL_W = 320
SIDE_PANEL_X = BOARD_X + BOARD_W * CELL + 40

# Colors
BG_COLOR = (12, 14, 22)
PANEL_BG = (18, 21, 33)
GRID_COLOR = (30, 36, 56)
TEXT_COLOR = (235, 238, 245)
MUTED_TEXT = (170, 178, 200)

# Piece colors
PIECE_COLORS = {
    "I": (70, 220, 255),
    "O": (255, 220, 70),
    "T": (180, 95, 255),
    "S": (90, 235, 120),
    "Z": (255, 90, 120),
    "J": (80, 140, 255),
    "L": (255, 160, 80),
}

# -------- Gameplay tuning --------
GRAVITY_START = 0.75
GRAVITY_FACTOR = 0.86
GRAVITY_MIN = 0.05

SOFT_DROP_MULT = 15.0
LOCK_DELAY = 0.5

MOVE_REPEAT_DELAY = 0.12
MOVE_REPEAT_RATE = 0.035

SCORE_LINE_CLEAR = {1: 100, 2: 300, 3: 500, 4: 800}
SOFT_DROP_SCORE_PER_CELL = 1
HARD_DROP_SCORE_PER_CELL = 2
LINES_PER_LEVEL = 10

# Fonts
FONT_NAME = None  # default pygame font
FONT_BIG = 40
FONT_MED = 24
FONT_SMALL = 18
