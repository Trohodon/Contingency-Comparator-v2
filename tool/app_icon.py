# tools/icon_builder.py
# Modern icon generator for: Dominion Contingency Comparator (DCC)
# Outputs:
#   - assets/app.ico     (Windows icon, multi-size)
#   - assets/app_256.png (preview)
#
# Requirements:
#   py -m pip install pillow --user

import os
from PIL import Image, ImageDraw, ImageFont, ImageFilter

SIZES = [16, 24, 32, 48, 64, 128, 256]

# ---------- Professional "grid + compare" palette ----------
NAVY0 = (8, 18, 38, 255)         # deep navy
NAVY1 = (14, 42, 92, 255)        # dominion-like blue
NAVY2 = (24, 74, 140, 255)       # lighter blue
STEEL = (210, 220, 235, 255)     # light steel
WHITE = (245, 248, 255, 255)
ACCENT = (255, 170, 40, 255)     # warm accent (warning/contingency)
ACCENT2 = (255, 110, 80, 255)    # secondary accent
INK = (0, 0, 0, 255)

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

def soft_radial_glow(size: int, center, radius, color_rgba):
    glow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(glow)
    cx, cy = center
    r = radius
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=color_rgba)
    glow = glow.filter(ImageFilter.GaussianBlur(max(1, size // 18)))
    return glow

def rounded_tile_mask(size: int, pad: int, radius: int):
    m = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(m)
    d.rounded_rectangle((pad, pad, size - pad, size - pad), radius=radius, fill=255)
    return m

def load_font(size: int):
    # Prefer Windows UI fonts; fall back safely.
    candidates = [
        ("segoeuib.ttf", int(size * 0.30)),
        ("seguisb.ttf", int(size * 0.30)),
        ("arialbd.ttf", int(size * 0.30)),
        ("arial.ttf", int(size * 0.30)),
    ]
    for name, pt in candidates:
        try:
            return ImageFont.truetype(name, max(10, pt))
        except Exception:
            pass
    return ImageFont.load_default()

def text_size(draw: ImageDraw.ImageDraw, text: str, font):
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        return bbox[2] - bbox[0], bbox[3] - bbox[1]
    except Exception:
        try:
            return font.getsize(text)
        except Exception:
            return (len(text) * 8, 12)

def draw_shadowed_line(d: ImageDraw.ImageDraw, p1, p2, width, color, shadow_alpha=70, shadow_offset=1):
    # subtle shadow for contrast on small sizes
    sx, sy = shadow_offset, shadow_offset
    d.line([(p1[0] + sx, p1[1] + sy), (p2[0] + sx, p2[1] + sy)], fill=(0, 0, 0, shadow_alpha), width=width)
    d.line([p1, p2], fill=color, width=width)

def draw_node(d: ImageDraw.ImageDraw, x, y, r, fill, outline=None, ow=1, shadow=True):
    if shadow:
        d.ellipse((x - r + 1, y - r + 1, x + r + 1, y + r + 1), fill=(0, 0, 0, 80))
    d.ellipse((x - r, y - r, x + r, y + r), fill=fill, outline=outline, width=ow)

def make_icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))

    pad = max(2, size // 18)
    radius = max(5, size // 5)

    # ----- Background tile -----
    bg = vertical_gradient(size, size, NAVY0, NAVY1, NAVY2)

    # soft highlight glow (top-left) + contingency glow (bottom-right)
    bg = Image.alpha_composite(bg, soft_radial_glow(size, (int(size * 0.30), int(size * 0.28)), int(size * 0.42), (255, 255, 255, 35)))
    bg = Image.alpha_composite(bg, soft_radial_glow(size, (int(size * 0.80), int(size * 0.80)), int(size * 0.45), (255, 170, 40, 18)))

    # clip to rounded tile
    tile = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    mask = rounded_tile_mask(size, pad, radius)
    tile.paste(bg, (0, 0), mask)

    # subtle border
    td = ImageDraw.Draw(tile)
    td.rounded_rectangle(
        (pad, pad, size - pad, size - pad),
        radius=radius,
        outline=(255, 255, 255, 45),
        width=max(1, size // 96),
    )

    img = Image.alpha_composite(img, tile)

    # ----- Foreground: "compare panels" + "grid one-line" -----
    fg = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(fg)

    # layout box
    left = int(size * 0.14)
    top = int(size * 0.18)
    right = int(size * 0.86)
    bottom = int(size * 0.82)

    # panel background
    panel_r = max(3, size // 14)
    d.rounded_rectangle((left, top, right, bottom), radius=panel_r, fill=(255, 255, 255, 18), outline=(255, 255, 255, 30), width=max(1, size // 128))

    # split line
    split_x = int(size * 0.51)
    d.line([(split_x, top + int(size * 0.03)), (split_x, bottom - int(size * 0.03))],
           fill=(255, 255, 255, 28), width=max(1, size // 128))

    # Decide detail level
    tiny = size <= 24
    small = size <= 32

    # --- Compare arrows + Δ (core "Comparator" idea) ---
    arrow_w = max(1, size // 64) if (tiny or small) else max(2, size // 72)
    mid_y = int(size * 0.50)

    # left->right arrow
    ax1 = int(size * 0.32)
    ax2 = int(size * 0.70)
    draw_shadowed_line(d, (ax1, mid_y - int(size * 0.11)), (ax2, mid_y - int(size * 0.11)), arrow_w, (255, 255, 255, 190), shadow_alpha=60, shadow_offset=1)
    d.polygon([(ax2, mid_y - int(size * 0.11)),
               (ax2 - int(size * 0.05), mid_y - int(size * 0.14)),
               (ax2 - int(size * 0.05), mid_y - int(size * 0.08))],
              fill=(255, 255, 255, 200))

    # right->left arrow (slightly lower)
    draw_shadowed_line(d, (ax2, mid_y + int(size * 0.11)), (ax1, mid_y + int(size * 0.11)), arrow_w, (255, 255, 255, 150), shadow_alpha=45, shadow_offset=1)
    d.polygon([(ax1, mid_y + int(size * 0.11)),
               (ax1 + int(size * 0.05), mid_y + int(size * 0.08)),
               (ax1 + int(size * 0.05), mid_y + int(size * 0.14))],
              fill=(255, 255, 255, 160))

    # Δ mark (difference)
    if not tiny:
        font = load_font(size)
        delta = "Δ"
        tw, th = text_size(d, delta, font)
        dx = int(size * 0.51 - tw / 2)
        dy = int(size * 0.50 - th / 2)
        # shadow + accent
        d.text((dx + 1, dy + 1), delta, font=font, fill=(0, 0, 0, 120))
        d.text((dx, dy), delta, font=font, fill=ACCENT)

    # --- One-line grid motif (left panel) ---
    # Keep it readable: 4 buses + lines
    line_w = max(1, size // 48) if small else max(2, size // 56)

    # coordinates in left panel
    lx0, lx1 = left + int(size * 0.06), split_x - int(size * 0.06)
    ly0, ly1 = top + int(size * 0.08), bottom - int(size * 0.08)

    # nodes
    n1 = (int(lx0), int(ly0 + (ly1 - ly0) * 0.20))
    n2 = (int(lx1), int(ly0 + (ly1 - ly0) * 0.28))
    n3 = (int(lx0 + (lx1 - lx0) * 0.35), int(ly0 + (ly1 - ly0) * 0.60))
    n4 = (int(lx1), int(ly0 + (ly1 - ly0) * 0.72))

    # lines
    if not tiny:
        draw_shadowed_line(d, n1, n2, line_w, (210, 220, 235, 200))
        draw_shadowed_line(d, n1, n3, line_w, (210, 220, 235, 180))
        draw_shadowed_line(d, n3, n4, line_w, (210, 220, 235, 180))
        draw_shadowed_line(d, n2, n4, line_w, (210, 220, 235, 140))

    # nodes
    r_node = max(2, size // 22)
    for (x, y) in [n1, n2, n3, n4]:
        draw_node(d, x, y, r_node, fill=WHITE, outline=(255, 255, 255, 120), ow=max(1, size // 160), shadow=True)

    # --- Contingency indicator (right panel): warning spark + result bars ---
    rx0, rx1 = split_x + int(size * 0.06), right - int(size * 0.06)
    ry0, ry1 = top + int(size * 0.10), bottom - int(size * 0.10)

    # simplified for tiny sizes: just a bolt mark
    bolt = [
        (int(rx0 + (rx1 - rx0) * 0.40), int(ry0 + (ry1 - ry0) * 0.05)),
        (int(rx0 + (rx1 - rx0) * 0.58), int(ry0 + (ry1 - ry0) * 0.38)),
        (int(rx0 + (rx1 - rx0) * 0.50), int(ry0 + (ry1 - ry0) * 0.38)),
        (int(rx0 + (rx1 - rx0) * 0.66), int(ry0 + (ry1 - ry0) * 0.85)),
        (int(rx0 + (rx1 - rx0) * 0.42), int(ry0 + (ry1 - ry0) * 0.52)),
        (int(rx0 + (rx1 - rx0) * 0.52), int(ry0 + (ry1 - ry0) * 0.52)),
    ]

    # shadow + bolt
    d.polygon([(x + 1, y + 1) for x, y in bolt], fill=(0, 0, 0, 90))
    d.polygon(bolt, fill=ACCENT)

    if not small:
        # "result bars" (like severity list)
        bar_h = max(2, size // 44)
        gap = max(2, size // 56)
        bx0 = int(rx0)
        by = int(ry0 + (ry1 - ry0) * 0.55)
        widths = [0.90, 0.70, 0.82]
        cols = [(255, 255, 255, 160), (255, 255, 255, 120), (ACCENT2[0], ACCENT2[1], ACCENT2[2], 150)]
        for i in range(3):
            bw = int((rx1 - rx0) * widths[i])
            d.rounded_rectangle((bx0, by + i * (bar_h + gap), bx0 + bw, by + i * (bar_h + gap) + bar_h),
                                radius=bar_h // 2,
                                fill=cols[i])

    # Composite foreground with a tiny blur to reduce jaggies at small sizes
    if size <= 32:
        fg = fg.filter(ImageFilter.GaussianBlur(0.3))

    img = Image.alpha_composite(img, fg)

    # ----- Optional small "DCC" mark (only when it won't clutter) -----
    if size >= 48:
        mark = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        md = ImageDraw.Draw(mark)
        font = load_font(size)
        label = "DCC"
        tw, th = text_size(md, label, font)
        mx = int(size * 0.16)
        my = int(size * 0.10)
        # badge behind
        bh = int(th * 1.20)
        bw = int(tw * 1.25)
        md.rounded_rectangle((mx - int(size * 0.02), my - int(size * 0.02), mx - int(size * 0.02) + bw, my - int(size * 0.02) + bh),
                             radius=max(3, bh // 3),
                             fill=(0, 0, 0, 60),
                             outline=(255, 255, 255, 35),
                             width=max(1, size // 128))
        # text
        md.text((mx, my), label, font=font, fill=(255, 255, 255, 220))
        # tiny accent underline
        ux0 = mx
        uy = my + th + max(1, size // 96)
        md.line([(ux0, uy), (ux0 + int(tw * 0.65), uy)], fill=(ACCENT[0], ACCENT[1], ACCENT[2], 220), width=max(1, size // 128))
        mark = mark.filter(ImageFilter.GaussianBlur(max(0, size // 512)))
        img = Image.alpha_composite(img, mark)

    # sharpen for larger sizes
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

    images = [make_icon(s) for s in SIZES]

    # preview
    images[-1].save(png_path, format="PNG")

    # multi-size .ico
    images[0].save(ico_path, format="ICO", sizes=[(s, s) for s in SIZES])

    print("Modern DCC icon generated successfully:")
    print(" -", ico_path)
    print(" -", png_path)

if __name__ == "__main__":
    main()