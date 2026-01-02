# tools/icon_builder.py
# Utility-pole icon generator for Dominion Contingency Comparator
# Outputs:
#   - assets/app.ico     (Windows icon, multi-size)
#   - assets/app_256.png (preview)
#
# Requirements:
#   py -m pip install pillow --user

import os
import math
from PIL import Image, ImageDraw, ImageFilter

SIZES = [16, 24, 32, 48, 64, 128, 256]

# Professional palette (no neon)
SKY_TOP = (18, 46, 92, 255)
SKY_MID = (32, 78, 140, 255)
SKY_BOT = (110, 165, 210, 255)

POLE = (32, 36, 44, 255)         # dark charcoal
POLE_EDGE = (255, 255, 255, 55)  # subtle highlight stroke
WIRE = (28, 30, 36, 230)
WIRE_HI = (255, 255, 255, 35)
INSULATOR = (70, 78, 92, 255)
INSULATOR_HI = (255, 255, 255, 60)

def lerp(a, b, t: float):
    return int(a + (b - a) * t)

def lerp_rgba(c1, c2, t: float):
    return (
        lerp(c1[0], c2[0], t),
        lerp(c1[1], c2[1], t),
        lerp(c1[2], c2[2], t),
        lerp(c1[3], c2[3], t),
    )

def vertical_gradient(w: int, h: int, top, mid, bottom):
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    px = img.load()
    for y in range(h):
        t = y / max(1, (h - 1))
        if t < 0.60:
            tt = t / 0.60
            c = lerp_rgba(top, mid, tt)
        else:
            tt = (t - 0.60) / 0.40
            c = lerp_rgba(mid, bottom, tt)
        for x in range(w):
            px[x, y] = c
    return img

def rounded_tile(size: int, pad: int, radius: int, bg: Image.Image):
    tile = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    mask = Image.new("L", (size, size), 0)
    md = ImageDraw.Draw(mask)
    md.rounded_rectangle((pad, pad, size - pad, size - pad), radius=radius, fill=255)
    tile.paste(bg, (0, 0), mask)

    td = ImageDraw.Draw(tile)
    td.rounded_rectangle(
        (pad, pad, size - pad, size - pad),
        radius=radius,
        outline=(255, 255, 255, 45),
        width=max(1, size // 96),
    )
    return tile

def sag_y(x, x0, x1, y0, sag):
    """
    Simple parabola sag: y = y0 + sag * (1 - ((2t-1)^2))
    so it sags most at center and is y0 at ends.
    """
    if x1 == x0:
        return y0
    t = (x - x0) / (x1 - x0)
    p = (2 * t - 1)
    return y0 + sag * (1 - p * p)

def draw_wire(d: ImageDraw.ImageDraw, x0, x1, y0, sag, width, color, hi=False):
    pts = []
    steps = max(16, int((x1 - x0) / 6))
    for i in range(steps + 1):
        x = x0 + (x1 - x0) * (i / steps)
        y = sag_y(x, x0, x1, y0, sag)
        pts.append((int(x), int(y)))
    d.line(pts, fill=color, width=width)
    if hi and width >= 2:
        d.line([(px, py - 1) for (px, py) in pts], fill=WIRE_HI, width=max(1, width // 2))

def draw_pole_icon(size: int) -> Image.Image:
    pad = max(2, size // 18)
    radius = max(5, size // 5)

    # Background
    bg = vertical_gradient(size, size, SKY_TOP, SKY_MID, SKY_BOT)

    # subtle vignette for contrast
    vig = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    vd = ImageDraw.Draw(vig)
    r = int(size * 0.72)
    cx, cy = int(size * 0.50), int(size * 0.45)
    vd.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(0, 0, 0, 55))
    vig = vig.filter(ImageFilter.GaussianBlur(max(1, size // 16)))
    bg = Image.alpha_composite(bg, vig)

    tile = rounded_tile(size, pad, radius, bg)
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    img = Image.alpha_composite(img, tile)

    d = ImageDraw.Draw(img)

    # Detail level
    tiny = size <= 24

    # Geometry
    x_center = int(size * 0.50)
    pole_top = int(size * 0.18)
    pole_bot = int(size * 0.88)
    pole_w = max(2, int(size * (0.10 if not tiny else 0.14)))
    half_w = pole_w // 2

    # Crossarm
    arm_y = int(size * 0.32)
    arm_len = int(size * (0.62 if not tiny else 0.70))
    arm_h = max(2, size // 28)

    arm_x0 = x_center - arm_len // 2
    arm_x1 = x_center + arm_len // 2

    # Wires
    wire_y_base = arm_y - int(size * 0.02)
    sag = max(1, int(size * (0.035 if not tiny else 0.025)))
    wire_w = max(1, size // 56)

    # draw wires behind pole
    draw_wire(d, pad, size - pad, wire_y_base, sag, wire_w, WIRE, hi=not tiny)
    draw_wire(d, pad, size - pad, wire_y_base + int(size * 0.06), sag + int(size * 0.01), wire_w, WIRE, hi=not tiny)
    draw_wire(d, pad, size - pad, wire_y_base + int(size * 0.12), sag + int(size * 0.02), wire_w, WIRE, hi=not tiny)

    # Pole (main shaft) with slight taper
    pole_poly = [
        (x_center - half_w - int(size * 0.01), pole_bot),
        (x_center + half_w + int(size * 0.01), pole_bot),
        (x_center + half_w, pole_top),
        (x_center - half_w, pole_top),
    ]
    d.polygon(pole_poly, fill=POLE)

    # Highlight edge
    if size >= 32:
        d.line([(x_center - half_w, pole_top), (x_center - half_w - 1, pole_bot)], fill=POLE_EDGE, width=max(1, size // 128))

    # Crossarm
    d.rounded_rectangle((arm_x0, arm_y - arm_h // 2, arm_x1, arm_y + arm_h // 2),
                        radius=max(1, arm_h // 2),
                        fill=POLE)

    # Insulators (3)
    if not tiny:
        ins_r = max(2, size // 40)
        xs = [arm_x0 + int(arm_len * 0.18), x_center, arm_x1 - int(arm_len * 0.18)]
        for ix in xs:
            # base
            d.ellipse((ix - ins_r, arm_y - ins_r, ix + ins_r, arm_y + ins_r), fill=INSULATOR)
            # top highlight
            d.ellipse((ix - ins_r + 1, arm_y - ins_r + 1, ix + ins_r - 1, arm_y + ins_r - 1),
                      outline=INSULATOR_HI, width=1)

        # small “hardware” bolts
        bolt_r = max(1, size // 110)
        d.ellipse((x_center - bolt_r, arm_y + arm_h // 2 + 2 - bolt_r,
                   x_center + bolt_r, arm_y + arm_h // 2 + 2 + bolt_r), fill=(255, 255, 255, 80))

    # Simple transformer box (optional, only when large enough)
    if size >= 64:
        box_w = int(size * 0.14)
        box_h = int(size * 0.12)
        bx0 = x_center + half_w + int(size * 0.03)
        by0 = int(size * 0.52)
        d.rounded_rectangle((bx0, by0, bx0 + box_w, by0 + box_h),
                            radius=max(2, size // 64),
                            fill=(42, 46, 56, 255),
                            outline=(255, 255, 255, 45),
                            width=max(1, size // 160))
        # strap
        d.line([(bx0, by0 + box_h // 2), (bx0 + box_w, by0 + box_h // 2)], fill=(255, 255, 255, 35), width=max(1, size // 200))

    # Slight soften at tiny sizes to reduce jaggies
    if size <= 32:
        img = img.filter(ImageFilter.GaussianBlur(0.25))
    # Sharpen for big sizes
    if size >= 64:
        img = img.filter(ImageFilter.UnsharpMask(radius=1, percent=120, threshold=2))

    return img

def main():
    here = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(here, ".."))
    out_dir = os.path.join(project_root, "assets")
    os.makedirs(out_dir, exist_ok=True)

    ico_path = os.path.join(out_dir, "app.ico")
    png_path = os.path.join(out_dir, "app_256.png")

    images = [draw_pole_icon(s) for s in SIZES]

    # preview
    images[-1].save(png_path, format="PNG")

    # multi-size ico
    images[0].save(ico_path, format="ICO", sizes=[(s, s) for s in SIZES])

    print("Utility-pole icon generated:")
    print(" -", ico_path)
    print(" -", png_path)

if __name__ == "__main__":
    main()