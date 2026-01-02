# tools/icon_builder.py
# Illustrated sunset powerline icon generator (vector-ish, aesthetic)
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

# --- Sunset palette (inspired by your reference image) ---
SKY_TOP   = (180, 60, 70, 255)     # dusky red
SKY_MID   = (235, 130, 60, 255)    # orange
SKY_BOT   = (255, 225, 145, 255)   # warm yellow

CLOUD_1   = (245, 150, 90, 255)    # darker cloud
CLOUD_2   = (255, 190, 115, 255)   # lighter cloud
CLOUD_3   = (255, 215, 150, 255)   # highlight cloud

FIELD_DK  = (70, 150, 90, 255)     # green
FIELD_LT  = (115, 195, 115, 255)   # bright green stripe
FIELD_MID = (95, 175, 105, 255)

POLE_1    = (60, 80, 80, 255)      # teal-ish pole
POLE_2    = (45, 55, 60, 255)      # shadow
POLE_WOOD = (85, 55, 40, 255)      # wood (for distant poles)

WIRE      = (25, 25, 28, 235)
WIRE_HI   = (255, 255, 255, 25)

def lerp(a, b, t: float) -> int:
    return int(a + (b - a) * t)

def lerp_rgba(c1, c2, t: float):
    return (lerp(c1[0], c2[0], t), lerp(c1[1], c2[1], t), lerp(c1[2], c2[2], t), lerp(c1[3], c2[3], t))

def vertical_gradient(w: int, h: int, top, mid, bottom):
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    px = img.load()
    for y in range(h):
        t = y / max(1, h - 1)
        if t < 0.55:
            tt = t / 0.55
            c = lerp_rgba(top, mid, tt)
        else:
            tt = (t - 0.55) / 0.45
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
        outline=(255, 255, 255, 50),
        width=max(1, size // 96),
    )
    return tile

def add_vignette(img: Image.Image, strength: int):
    w, h = img.size
    v = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(v)
    r = int(min(w, h) * 0.88)
    cx, cy = int(w * 0.5), int(h * 0.55)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(0, 0, 0, strength))
    v = v.filter(ImageFilter.GaussianBlur(max(1, w // 14)))
    return Image.alpha_composite(img, v)

def sag_curve_points(x0, x1, y0, sag, steps):
    pts = []
    for i in range(steps + 1):
        t = i / steps
        x = x0 + (x1 - x0) * t
        p = 2 * t - 1
        y = y0 + sag * (1 - p * p)
        pts.append((int(x), int(y)))
    return pts

def draw_wire(draw: ImageDraw.ImageDraw, x0, x1, y0, sag, width, hi=True):
    steps = max(20, int((x1 - x0) / 18))
    pts = sag_curve_points(x0, x1, y0, sag, steps)
    draw.line(pts, fill=WIRE, width=width)
    if hi and width >= 2:
        draw.line([(x, y - 1) for (x, y) in pts], fill=WIRE_HI, width=max(1, width // 2))

def blob_cloud(draw: ImageDraw.ImageDraw, x, y, w, h, color):
    # simple “cloud” blob with 3-4 circles
    r1 = int(h * 0.55)
    r2 = int(h * 0.70)
    r3 = int(h * 0.60)
    draw.ellipse((x + int(w*0.00), y + int(h*0.25), x + int(w*0.35), y + int(h*0.95)), fill=color)
    draw.ellipse((x + int(w*0.20), y + int(h*0.05), x + int(w*0.60), y + int(h*0.85)), fill=color)
    draw.ellipse((x + int(w*0.45), y + int(h*0.20), x + int(w*0.85), y + int(h*0.95)), fill=color)
    # base band
    draw.rounded_rectangle((x + int(w*0.05), y + int(h*0.45), x + int(w*0.85), y + int(h*0.98)),
                           radius=int(h*0.25), fill=color)

def draw_pole(draw: ImageDraw.ImageDraw, x, y_top, y_bot, w, cross_y, cross_len, style="teal", detail=True):
    # pole body (slight taper)
    taper = max(1, int(w * 0.18))
    x0 = x - w//2
    x1 = x + w//2
    poly = [(x0, y_bot), (x1, y_bot), (x1 - taper, y_top), (x0 + taper, y_top)]
    if style == "wood":
        base = POLE_WOOD
        shadow = (60, 35, 25, 255)
        highlight = (255, 255, 255, 25)
    else:
        base = POLE_1
        shadow = POLE_2
        highlight = (255, 255, 255, 30)

    # shadow underlay
    draw.polygon([(p[0] + 2, p[1] + 2) for p in poly], fill=(0, 0, 0, 70))
    # main
    draw.polygon(poly, fill=base)
    # right shadow strip
    draw.polygon([(x1 - int(w*0.15), y_bot), (x1, y_bot), (x1 - taper, y_top), (x1 - taper - int(w*0.12), y_top)], fill=shadow)
    # subtle highlight
    draw.line([(x0 + int(w*0.12), y_top + 3), (x0 + int(w*0.12), y_bot - 3)], fill=highlight, width=max(1, w//6))

    # crossarm
    arm_h = max(2, w//3)
    ax0 = x - cross_len//2
    ax1 = x + cross_len//2
    draw.rounded_rectangle((ax0, cross_y - arm_h//2, ax1, cross_y + arm_h//2),
                           radius=arm_h//2, fill=base)
    draw.line([(ax0, cross_y + arm_h//2 - 1), (ax1, cross_y + arm_h//2 - 1)], fill=(0, 0, 0, 80), width=1)

    if detail:
        # insulators (small caps)
        ins_r = max(2, w//3)
        for ix in [ax0 + int(cross_len*0.2), x, ax1 - int(cross_len*0.2)]:
            draw.ellipse((ix - ins_r, cross_y - ins_r - arm_h//2,
                          ix + ins_r, cross_y + ins_r - arm_h//2),
                         fill=(230, 235, 240, 255))
            draw.ellipse((ix - ins_r + 1, cross_y - ins_r - arm_h//2 + 1,
                          ix + ins_r - 1, cross_y + ins_r - arm_h//2 - 1),
                         outline=(0, 0, 0, 60), width=1)

def make_icon(target: int) -> Image.Image:
    # Oversample for clean vector edges
    scale = 4
    size = target * scale
    pad = max(6, size // 22)
    radius = max(18, size // 5)

    # Background
    bg = vertical_gradient(size, size, SKY_TOP, SKY_MID, SKY_BOT)

    # Clouds (layered)
    clouds = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    cd = ImageDraw.Draw(clouds)

    # cloud band positions tuned for the reference vibe
    blob_cloud(cd, int(size*0.05), int(size*0.18), int(size*0.55), int(size*0.18), CLOUD_1)
    blob_cloud(cd, int(size*0.35), int(size*0.22), int(size*0.65), int(size*0.20), CLOUD_2)
    blob_cloud(cd, int(size*0.10), int(size*0.30), int(size*0.70), int(size*0.22), CLOUD_2)
    blob_cloud(cd, int(size*0.40), int(size*0.34), int(size*0.58), int(size*0.20), CLOUD_3)

    clouds = clouds.filter(ImageFilter.GaussianBlur(max(1, size // 260)))
    bg = Image.alpha_composite(bg, clouds)

    # Field
    field = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    fd = ImageDraw.Draw(field)
    horizon = int(size * 0.62)
    fd.rectangle((0, horizon, size, size), fill=FIELD_DK)

    # angled stripes
    stripe_w = int(size * 0.10)
    for i in range(-2, 14):
        x0 = int(i * stripe_w)
        fd.polygon([(x0, size), (x0 + stripe_w, size), (x0 + int(stripe_w*2.0), horizon), (x0 + int(stripe_w*1.0), horizon)],
                   fill=FIELD_MID if i % 2 == 0 else FIELD_LT)
    # soften a bit
    field = field.filter(ImageFilter.GaussianBlur(max(0, size // 700)))
    bg = Image.alpha_composite(bg, field)

    # Clip to rounded app tile + subtle vignette
    tile = rounded_tile(size, pad, radius, bg)
    tile = add_vignette(tile, strength=50)

    # Foreground drawing
    fg = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(fg)

    # Complexity decisions
    tiny = target <= 24
    small = target <= 32

    # Big pole (foreground)
    pole_x = int(size * 0.60)
    pole_top = int(size * 0.16)
    pole_bot = int(size * 0.93)
    pole_w = int(size * (0.09 if not tiny else 0.12))
    cross_y = int(size * 0.30)
    cross_len = int(size * 0.55)

    # Distant poles (depth) — only when we have room
    if target >= 64:
        # a few small poles receding left
        for k in range(4):
            t = k / 3
            x = int(lerp(int(size*0.12), int(size*0.42), t))
            yb = int(lerp(int(size*0.90), int(size*0.75), t))
            yt = int(lerp(int(size*0.55), int(size*0.40), t))
            w = int(lerp(int(size*0.022), int(size*0.045), t))
            cy = int(lerp(int(size*0.50), int(size*0.36), t))
            clen = int(lerp(int(size*0.16), int(size*0.30), t))
            draw_pole(d, x, yt, yb, w, cy, clen, style="wood", detail=False)

    # Wires (perspective sweep)
    wire_w = max(3, size // 160)
    x0 = -int(size * 0.10)
    x1 = int(size * 1.10)
    base_y = int(size * 0.10)
    for i in range(5):
        y = base_y + int(size * 0.06) * i
        sag = int(size * 0.015) + i * int(size * 0.004)
        draw_wire(d, x0, x1, y, sag, wire_w, hi=not tiny)

    # Big pole on top (so it intersects wires cleanly)
    draw_pole(d, pole_x, pole_top, pole_bot, pole_w, cross_y, cross_len, style="teal", detail=(not tiny))

    # slight polish blur to avoid jaggies, then sharpen on downscale
    if target <= 32:
        fg = fg.filter(ImageFilter.GaussianBlur(0.45))
    else:
        fg = fg.filter(ImageFilter.GaussianBlur(0.25))

    img = Image.alpha_composite(tile, fg)

    # Final: downsample to target
    img = img.resize((target, target), resample=Image.Resampling.LANCZOS)
    if target <= 32:
        img = img.filter(ImageFilter.UnsharpMask(radius=1, percent=150, threshold=2))
    elif target >= 64:
        img = img.filter(ImageFilter.UnsharpMask(radius=1, percent=120, threshold=2))

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

    print("Illustrated sunset pole icon generated:")
    print(" -", ico_path)
    print(" -", png_path)

if __name__ == "__main__":
    main()