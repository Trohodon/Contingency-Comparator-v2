# tools/icon_builder.py
# Dominion Contingency Comparator (DCC) - dynamic retro-futuristic icon generator
#
# Outputs:
#   - assets/app.ico     (Windows icon, multi-size)
#   - assets/app_256.png (preview)
#
# Install:
#   py -m pip install pillow --user

import os
import random
import math
from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageChops, ImageEnhance

SIZES = [16, 24, 32, 48, 64, 128, 256]

# ---------- Theme ----------
DEEP    = (10, 12, 28, 255)      # near-black navy
PURPLE  = (130, 58, 240, 255)    # neon purple
PINK    = (255, 64, 180, 255)    # hot pink
CYAN    = (64, 220, 255, 255)    # neon cyan
TEAL    = (40, 255, 200, 255)
WHITE   = (245, 248, 255, 255)

APP_MARK = "DCC"     # top-left text
SEED = 1337          # change to get a different “dynamic” pattern, deterministic


# ----------------- helpers -----------------
def clamp255(v: int) -> int:
    return 0 if v < 0 else 255 if v > 255 else v

def lerp(a, b, t: float):
    return int(a + (b - a) * t)

def lerp_rgba(c1, c2, t: float):
    return (
        lerp(c1[0], c2[0], t),
        lerp(c1[1], c2[1], t),
        lerp(c1[2], c2[2], t),
        lerp(c1[3], c2[3], t),
    )

def add_glow_from_alpha(layer: Image.Image, glow_color, blur_radius: int, intensity: float = 1.0):
    """Create glow from layer alpha and return RGBA glow image."""
    a = layer.split()[-1]
    glow = Image.new("RGBA", layer.size, glow_color)
    glow.putalpha(a)
    glow = glow.filter(ImageFilter.GaussianBlur(blur_radius))
    if intensity != 1.0:
        ga = glow.split()[-1]
        ga = ga.point(lambda p: clamp255(int(p * intensity)))
        glow.putalpha(ga)
    return glow

def safe_round_rect(draw: ImageDraw.ImageDraw, xy, r: int, fill=None, outline=None, width=1):
    draw.rounded_rectangle(xy, radius=r, fill=fill, outline=outline, width=width)

def load_best_font(size: int):
    # Try Windows fonts; fallback is fine.
    candidates = [
        ("segoeuib.ttf", int(size * 0.34)),
        ("seguisb.ttf",  int(size * 0.34)),
        ("arialbd.ttf",  int(size * 0.34)),
        ("arial.ttf",    int(size * 0.34)),
    ]
    for name, pt in candidates:
        try:
            return ImageFont.truetype(name, max(10, pt))
        except Exception:
            pass
    return ImageFont.load_default()

def measure_text(draw: ImageDraw.ImageDraw, text: str, font):
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        return (bbox[2] - bbox[0], bbox[3] - bbox[1])
    except Exception:
        try:
            return font.getsize(text)
        except Exception:
            return (len(text) * 8, 12)

def make_vertical_gradient(w: int, h: int, top, mid, bottom):
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    px = img.load()
    for y in range(h):
        t = y / max(1, (h - 1))
        if t < 0.55:
            tt = t / 0.55
            c = lerp_rgba(top, mid, tt)
        else:
            tt = (t - 0.55) / 0.45
            c = lerp_rgba(mid, bottom, tt)
        for x in range(w):
            px[x, y] = c
    return img

def make_radial_glow(size: int, center, radius: int, inner_rgba, outer_rgba):
    """Cheap radial gradient via concentric circles. Looks good at icon sizes."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    cx, cy = center
    steps = max(24, size // 2)
    for i in range(steps, 0, -1):
        t = i / steps
        # invert for inner->outer
        c = lerp_rgba(inner_rgba, outer_rgba, 1.0 - t)
        r = int(radius * t)
        d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=c)
    return img.filter(ImageFilter.GaussianBlur(max(1, size // 40)))

def add_scanlines(img: Image.Image, strength: float):
    """Subtle horizontal scanlines."""
    w, h = img.size
    lines = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(lines)
    alpha = int(30 * strength)
    for y in range(0, h, 2):
        d.line([(0, y), (w, y)], fill=(0, 0, 0, alpha), width=1)
    return Image.alpha_composite(img, lines)

def add_subtle_noise(img: Image.Image, amount: int, seed: int):
    """Add low-amplitude monochrome noise to break banding."""
    random.seed(seed)
    w, h = img.size
    noise = Image.new("L", (w, h), 0)
    px = noise.load()
    for y in range(h):
        for x in range(w):
            px[x, y] = random.randint(0, amount)
    noise = noise.filter(ImageFilter.GaussianBlur(1))
    noise_rgba = Image.merge("RGBA", (noise, noise, noise, noise.point(lambda p: int(p * 0.35))))
    return ImageChops.screen(img, noise_rgba)

def vignette(img: Image.Image, strength: float):
    w, h = img.size
    mask = Image.new("L", (w, h), 0)
    d = ImageDraw.Draw(mask)
    # strong edges
    d.ellipse((-w * 0.2, -h * 0.2, w * 1.2, h * 1.2), fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(int(min(w, h) * 0.18)))
    # invert-ish by subtracting from white
    inv = ImageChops.invert(mask)
    inv = inv.point(lambda p: int(p * 0.55 * strength))
    shade = Image.new("RGBA", (w, h), (0, 0, 0, 255))
    shade.putalpha(inv)
    return Image.alpha_composite(img, shade)

def polygon_points(center, radius, sides, rotation=0.0):
    cx, cy = center
    pts = []
    for i in range(sides):
        a = rotation + (2 * math.pi * i / sides)
        pts.append((cx + radius * math.cos(a), cy + radius * math.sin(a)))
    return pts

def draw_energy_arcs(draw: ImageDraw.ImageDraw, center, r, width, color1, color2, seed: int):
    """Arc fragments to suggest motion/energy."""
    random.seed(seed)
    cx, cy = center
    for i in range(5):
        start = random.randint(0, 300)
        span = random.randint(30, 90)
        col = lerp_rgba(color1, color2, i / 4)
        bbox = (cx - r, cy - r, cx + r, cy + r)
        draw.arc(bbox, start=start, end=start + span, fill=col, width=width)

def draw_power_glyph(layer: Image.Image, size: int, seed: int):
    """
    Abstract transmission/power glyph:
    - central “tower” + two side nodes + connectors
    Designed to stay readable even when scaled down.
    """
    random.seed(seed)
    d = ImageDraw.Draw(layer)

    w = h = size
    cx = int(w * 0.52)
    cy = int(h * 0.58)

    # scale knobs
    lw = max(2, size // 48)
    node_r = max(3, size // 22)

    # main tower spine
    top = (cx, int(h * 0.30))
    mid = (cx, int(h * 0.52))
    bot = (cx, int(h * 0.80))
    d.line([top, bot], fill=(255, 255, 255, 180), width=lw)

    # crossbars
    bar1_y = int(h * 0.42)
    bar2_y = int(h * 0.58)
    bar_w1 = int(w * 0.22)
    bar_w2 = int(w * 0.30)
    d.line([(cx - bar_w1, bar1_y), (cx + bar_w1, bar1_y)], fill=(64, 220, 255, 170), width=lw)
    d.line([(cx - bar_w2, bar2_y), (cx + bar_w2, bar2_y)], fill=(255, 64, 180, 155), width=lw)

    # side nodes
    left = (int(w * 0.22), int(h * 0.62))
    right = (int(w * 0.82), int(h * 0.50))
    d.line([left, (cx, bar2_y)], fill=(255, 64, 180, 140), width=lw)
    d.line([right, (cx, bar1_y)], fill=(64, 220, 255, 140), width=lw)

    # nodes themselves
    for (x, y) in [left, right, (cx, bar1_y), (cx, bar2_y)]:
        d.ellipse((x - node_r, y - node_r, x + node_r, y + node_r),
                  fill=WHITE, outline=(64, 220, 255, 200), width=max(1, lw // 2))

    # little “spark” tick
    if size >= 48:
        sx = int(w * 0.74)
        sy = int(h * 0.32)
        d.line([(sx, sy), (sx + int(w * 0.06), sy + int(h * 0.06))], fill=(255, 64, 180, 190), width=lw)
        d.line([(sx + int(w * 0.06), sy), (sx, sy + int(h * 0.06))], fill=(64, 220, 255, 190), width=lw)

def draw_radar_sweep(layer: Image.Image, size: int, center, radius: int, angle_deg: float, color):
    """A wedge with gradient alpha, blurred, gives a ‘sweep’ feel."""
    w = h = size
    cx, cy = center

    sweep = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(sweep)

    # Sweep wedge polygons at a few alpha levels (stacked)
    for i in range(6):
        t = i / 5
        ang = math.radians(angle_deg + t * 18.0)  # wedge width
        ang0 = math.radians(angle_deg)

        r = radius
        p0 = (cx, cy)
        p1 = (cx + r * math.cos(ang0), cy + r * math.sin(ang0))
        p2 = (cx + r * math.cos(ang),  cy + r * math.sin(ang))

        a = int(90 * (1.0 - t))
        c = (color[0], color[1], color[2], a)
        d.polygon([p0, p1, p2], fill=c)

    sweep = sweep.filter(ImageFilter.GaussianBlur(max(1, size // 40)))
    layer.alpha_composite(sweep, (0, 0))

def make_icon(size: int) -> Image.Image:
    random.seed(SEED + size)

    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    pad = max(2, size // 18)
    radius = max(6, size // 5)

    # ----- base tile -----
    bg = make_vertical_gradient(
        size, size,
        top=DEEP,
        mid=(38, 18, 72, 255),
        bottom=(12, 8, 30, 255),
    )

    # subtle “aurora” glow blobs
    bg = Image.alpha_composite(bg, make_radial_glow(
        size, center=(int(size * 0.30), int(size * 0.28)),
        radius=int(size * 0.55),
        inner_rgba=(64, 220, 255, 110),
        outer_rgba=(0, 0, 0, 0),
    ))
    bg = Image.alpha_composite(bg, make_radial_glow(
        size, center=(int(size * 0.78), int(size * 0.78)),
        radius=int(size * 0.60),
        inner_rgba=(255, 64, 180, 95),
        outer_rgba=(0, 0, 0, 0),
    ))

    # clip to rounded square
    tile = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    mask = Image.new("L", (size, size), 0)
    md = ImageDraw.Draw(mask)
    md.rounded_rectangle((pad, pad, size - pad, size - pad), radius=radius, fill=255)
    tile.paste(bg, (0, 0), mask)

    # border stroke
    td = ImageDraw.Draw(tile)
    stroke_w = max(1, size // 64)
    td.rounded_rectangle(
        (pad, pad, size - pad, size - pad),
        radius=radius,
        outline=(120, 170, 255, 85),
        width=stroke_w,
        fill=None
    )

    img = Image.alpha_composite(img, tile)

    # ----- energy core (top-right) -----
    core = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    cd = ImageDraw.Draw(core)

    core_c = (int(size * 0.72), int(size * 0.32))
    core_r = int(size * 0.26)
    ring_w = max(2, size // 28)

    # outer ring
    cd.ellipse(
        (core_c[0] - core_r, core_c[1] - core_r, core_c[0] + core_r, core_c[1] + core_r),
        outline=(64, 220, 255, 210),
        width=ring_w
    )
    # inner ring
    cd.ellipse(
        (core_c[0] - int(core_r * 0.62), core_c[1] - int(core_r * 0.62),
         core_c[0] + int(core_r * 0.62), core_c[1] + int(core_r * 0.62)),
        outline=(255, 64, 180, 170),
        width=max(1, ring_w - 1)
    )

    # arc fragments (motion)
    if size >= 32:
        draw_energy_arcs(cd, core_c, int(core_r * 0.92), max(1, ring_w - 1),
                         color1=(64, 220, 255, 220), color2=(255, 64, 180, 220),
                         seed=SEED + 91 + size)

    # sweep wedge overlay
    if size >= 48:
        draw_radar_sweep(core, size, core_c, int(core_r * 1.05), angle_deg=-55.0, color=(64, 220, 255, 255))

    # glow + composite
    core_glow = add_glow_from_alpha(core, (64, 220, 255, 255), blur_radius=max(2, size // 14), intensity=1.15)
    img = Image.alpha_composite(img, core_glow)
    img = Image.alpha_composite(img, core)

    # ----- center glyph (power grid) -----
    glyph = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw_power_glyph(glyph, size, seed=SEED + 222)

    # Hex shield behind glyph (adds “badge” depth)
    if size >= 24:
        badge = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        bd = ImageDraw.Draw(badge)
        c = (int(size * 0.50), int(size * 0.62))
        rr = int(size * 0.28)
        pts = polygon_points(c, rr, sides=6, rotation=math.radians(30))
        bd.polygon(pts, fill=(20, 24, 50, 170), outline=(130, 58, 240, 130))
        badge = badge.filter(ImageFilter.GaussianBlur(max(0, size // 200)))
        img = Image.alpha_composite(img, badge)

    glyph_glow_p = add_glow_from_alpha(glyph, (255, 64, 180, 255), blur_radius=max(2, size // 16), intensity=1.0)
    glyph_glow_c = add_glow_from_alpha(glyph, (64, 220, 255, 255), blur_radius=max(2, size // 18), intensity=0.9)
    img = Image.alpha_composite(img, glyph_glow_p)
    img = Image.alpha_composite(img, glyph_glow_c)
    img = Image.alpha_composite(img, glyph)

    # ----- top-left mark (DCC) -----
    # For tiny sizes, don’t force text (it becomes mush).
    if size >= 32:
        txt = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        td = ImageDraw.Draw(txt)
        font = load_best_font(size)

        tw, th = measure_text(td, APP_MARK, font)
        tx = int(size * 0.12)
        ty = int(size * 0.10)

        # underline neon pill
        bar_h = max(6, int(th * 0.65))
        bar_w = max(28, int(tw * 1.15))
        bar = Image.new("RGBA", (bar_w, bar_h), (0, 0, 0, 0))
        bd = ImageDraw.Draw(bar)
        safe_round_rect(
            bd, (0, 0, bar_w, bar_h),
            r=max(4, bar_h // 2),
            fill=(255, 64, 180, 140),
            outline=(64, 220, 255, 170),
            width=max(1, size // 140)
        )
        bar = bar.filter(ImageFilter.GaussianBlur(max(0, size // 220)))
        txt.alpha_composite(bar, (tx - int(size * 0.01), ty + int(th * 0.55)))

        # text
        td.text((tx + 1, ty + 1), APP_MARK, font=font, fill=(0, 0, 0, 120))
        td.text((tx, ty), APP_MARK, font=font, fill=WHITE)

        txt_glow = add_glow_from_alpha(txt, (255, 64, 180, 255), blur_radius=max(2, size // 18), intensity=1.05)
        img = Image.alpha_composite(img, txt_glow)
        img = Image.alpha_composite(img, txt)

    # ----- polish pass -----
    img = add_subtle_noise(img, amount=max(8, size // 18), seed=SEED + 999 + size)
    img = add_scanlines(img, strength=0.75 if size >= 64 else 0.45)
    img = vignette(img, strength=0.85)

    # Slight contrast bump (keeps neon pop without crushing)
    img = ImageEnhance.Contrast(img).enhance(1.06)

    # Sharpen for bigger sizes
    if size >= 64:
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

    # preview png
    images[-1].save(png_path, format="PNG")

    # multi-size .ico
    images[0].save(ico_path, format="ICO", sizes=[(s, s) for s in SIZES])

    print("Dynamic icon generated successfully:")
    print(" -", ico_path)
    print(" -", png_path)


if __name__ == "__main__":
    main()