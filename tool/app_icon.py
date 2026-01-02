# tools/icon_builder.py
# Utility-pole icon generator (more aesthetic, smoother, with depth)
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

# ---------- Modern, clean palette ----------
SKY_TOP = (18, 55, 110, 255)
SKY_MID = (42, 110, 185, 255)
SKY_BOT = (165, 215, 245, 255)

POLE_DARK = (34, 38, 46, 255)
POLE_MID  = (52, 58, 70, 255)
POLE_HI   = (255, 255, 255, 60)

WIRE = (25, 27, 32, 240)
WIRE_HI = (255, 255, 255, 35)

HARDWARE = (95, 105, 120, 255)
HARDWARE_HI = (255, 255, 255, 55)

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

def add_gloss(img: Image.Image, strength: int):
    """Subtle top gloss for a more app-icon look."""
    w, h = img.size
    gloss = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(gloss)
    # A soft white ellipse at the top
    r = int(min(w, h) * 0.85)
    cx, cy = int(w * 0.5), int(h * 0.15)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(255, 255, 255, strength))
    gloss = gloss.filter(ImageFilter.GaussianBlur(max(1, w // 18)))
    return Image.alpha_composite(img, gloss)

def add_vignette(img: Image.Image, strength: int):
    """Darken edges a bit so the foreground pops."""
    w, h = img.size
    vig = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(vig)
    r = int(min(w, h) * 0.80)
    cx, cy = int(w * 0.5), int(h * 0.55)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(0, 0, 0, strength))
    vig = vig.filter(ImageFilter.GaussianBlur(max(1, w // 14)))
    return Image.alpha_composite(img, vig)

def sag_y(x, x0, x1, y0, sag):
    # Parabola sag. y0 at ends, max sag at center.
    if x1 == x0:
        return y0
    t = (x - x0) / (x1 - x0)
    p = 2 * t - 1
    return y0 + sag * (1 - p * p)

def draw_wire(draw: ImageDraw.ImageDraw, x0, x1, y0, sag, width):
    pts = []
    steps = max(20, int((x1 - x0) / 10))
    for i in range(steps + 1):
        x = x0 + (x1 - x0) * (i / steps)
        y = sag_y(x, x0, x1, y0, sag)
        pts.append((int(x), int(y)))

    # main wire
    draw.line(pts, fill=WIRE, width=width)
    # highlight
    if width >= 2:
        draw.line([(x, y - 1) for x, y in pts], fill=WIRE_HI, width=max(1, width // 2))

def draw_pole_layer(size: int) -> Image.Image:
    """
    Draw the pole + wires at 'size' with nice shading.
    This will be called at oversampled resolution then downscaled.
    """
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    pad = max(2, size // 18)
    tiny = size <= 24 * 4  # because we render at 4x

    # Layout
    x_center = int(size * 0.52)
    pole_top = int(size * 0.16)
    pole_bot = int(size * 0.90)

    pole_w = max(10, int(size * (0.09 if not tiny else 0.11)))
    half_w = pole_w // 2

    arm_y = int(size * 0.30)
    arm_len = int(size * 0.66)
    arm_h = max(10, size // 28)

    arm_x0 = x_center - arm_len // 2
    arm_x1 = x_center + arm_len // 2

    # --- Wires behind pole ---
    wire_w = max(4, size // 70)
    y0 = arm_y - int(size * 0.02)
    sag = max(4, int(size * 0.030))
    draw_wire(d, pad, size - pad, y0, sag, wire_w)
    draw_wire(d, pad, size - pad, y0 + int(size * 0.06), sag + int(size * 0.01), wire_w)
    draw_wire(d, pad, size - pad, y0 + int(size * 0.12), sag + int(size * 0.02), wire_w)

    # --- Pole shadow (soft) ---
    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    sd.rounded_rectangle(
        (x_center - half_w + int(size * 0.015), pole_top + int(size * 0.01),
         x_center + half_w + int(size * 0.015), pole_bot + int(size * 0.01)),
        radius=max(6, pole_w // 3),
        fill=(0, 0, 0, 110),
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(max(2, size // 80)))
    img = Image.alpha_composite(img, shadow)
    d = ImageDraw.Draw(img)

    # --- Pole body with subtle vertical gradient feel ---
    pole = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    pd = ImageDraw.Draw(pole)
    # main
    pd.rounded_rectangle(
        (x_center - half_w, pole_top, x_center + half_w, pole_bot),
        radius=max(6, pole_w // 3),
        fill=POLE_MID
    )
    # darker right edge
    pd.rounded_rectangle(
        (x_center + int(half_w * 0.30), pole_top, x_center + half_w, pole_bot),
        radius=max(6, pole_w // 3),
        fill=POLE_DARK
    )
    # highlight left edge
    pd.line([(x_center - half_w + 2, pole_top + 6), (x_center - half_w + 2, pole_bot - 6)],
            fill=POLE_HI, width=max(2, size // 140))

    pole = pole.filter(ImageFilter.GaussianBlur(max(0, size // 500)))
    img = Image.alpha_composite(img, pole)
    d = ImageDraw.Draw(img)

    # --- Crossarm ---
    d.rounded_rectangle(
        (arm_x0, arm_y - arm_h // 2, arm_x1, arm_y + arm_h // 2),
        radius=max(6, arm_h // 2),
        fill=POLE_MID
    )
    # crossarm shadow line
    d.line([(arm_x0, arm_y + arm_h // 2 - 2), (arm_x1, arm_y + arm_h // 2 - 2)],
           fill=(0, 0, 0, 80), width=max(2, size // 180))
    # crossarm highlight
    d.line([(arm_x0, arm_y - arm_h // 2 + 2), (arm_x1, arm_y - arm_h // 2 + 2)],
           fill=(255, 255, 255, 35), width=max(2, size // 200))

    # Center bolt
    bolt_r = max(8, size // 90)
    d.ellipse((x_center - bolt_r, arm_y - bolt_r, x_center + bolt_r, arm_y + bolt_r), fill=HARDWARE)
    d.ellipse((x_center - bolt_r + 3, arm_y - bolt_r + 3, x_center + bolt_r - 3, arm_y + bolt_r - 3),
              outline=HARDWARE_HI, width=max(2, size // 300))

    # --- Insulators / attachments ---
    # Keep them simple at tiny sizes
    if not tiny:
        ins_r = max(10, size // 55)
        xs = [arm_x0 + int(arm_len * 0.18), x_center, arm_x1 - int(arm_len * 0.18)]
        for ix in xs:
            # base
            d.rounded_rectangle((ix - ins_r, arm_y - ins_r, ix + ins_r, arm_y + ins_r),
                                radius=ins_r // 2, fill=HARDWARE)
            # cap highlight
            d.line([(ix - ins_r + 3, arm_y - ins_r + 3), (ix + ins_r - 3, arm_y - ins_r + 3)],
                   fill=HARDWARE_HI, width=max(2, size // 260))
            # tiny hook to wire
            d.line([(ix, arm_y + ins_r), (ix, arm_y + ins_r + int(size * 0.018))],
                   fill=(0, 0, 0, 120), width=max(2, size // 260))

    # --- Foreground: re-draw short wire segments in front of pole for depth ---
    # This makes the pole feel "in" the scene.
    front = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    fd = ImageDraw.Draw(front)
    clip_x0 = x_center - half_w - int(size * 0.03)
    clip_x1 = x_center + half_w + int(size * 0.03)

    # small segments near the pole area
    for i, dy in enumerate([0, int(size * 0.06), int(size * 0.12)]):
        y = y0 + dy
        sag_i = sag + int(size * 0.01) * i
        # left segment
        draw_wire(fd, pad, clip_x0, y, sag_i, wire_w)
        # right segment
        draw_wire(fd, clip_x1, size - pad, y, sag_i, wire_w)

    img = Image.alpha_composite(img, front)
    return img

def make_icon(target_size: int) -> Image.Image:
    # Oversample for smoothness
    scale = 4
    size = target_size * scale

    pad = max(2, size // 18)
    radius = max(10, size // 5)

    # Background tile
    bg = vertical_gradient(size, size, SKY_TOP, SKY_MID, SKY_BOT)
    bg = add_gloss(bg, strength=70)
    bg = add_vignette(bg, strength=55)

    tile = rounded_tile(size, pad, radius, bg)

    # Foreground pole+wires
    pole_layer = draw_pole_layer(size)

    img = Image.alpha_composite(tile, pole_layer)

    # Subtle final polish (very light)
    if size >= 64 * scale:
        img = img.filter(ImageFilter.UnsharpMask(radius=2, percent=120, threshold=3))

    # Downsample to target size
    img = img.resize((target_size, target_size), resample=Image.Resampling.LANCZOS)

    # Tiny-size cleanup: slightly crisp edges
    if target_size <= 32:
        img = img.filter(ImageFilter.UnsharpMask(radius=1, percent=140, threshold=2))

    return img

def main():
    here = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(here, ".."))
    out_dir = os.path.join(project_root, "assets")
    os.makedirs(out_dir, exist_ok=True)

    ico_path = os.path.join(out_dir, "app.ico")
    png_path = os.path.join(out_dir, "app_256.png")

    images = [make_icon(s) for s in SIZES]

    images[-1].save(png_path, format="PNG")
    images[0].save(ico_path, format="ICO", sizes=[(s, s) for s in SIZES])

    print("Utility-pole icon generated:")
    print(" -", ico_path)
    print(" -", png_path)

if __name__ == "__main__":
    main()