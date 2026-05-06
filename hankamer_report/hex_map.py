"""
hex_map.py

Generates a US contiguous-48-state hex grid heat map as a PNG image using PIL.
No external SVG renderer required.

Usage:
    from hex_map import generate_hex_map_png
    png_bytes = generate_hex_map_png(state_counts)   # dict {abbr: count}
"""

import io
import math
from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------------------
# Hex grid layout: (col, row)
# Flat-top hexagons. Odd cols shift down by half a row.
# Matches standard NPR-style US hex map grid.
# ---------------------------------------------------------------------------
HEX_GRID = {
    # Row 0
    "WA": (0, 0), "MT": (1, 0), "ND": (2, 0), "MN": (3, 0), "WI": (4, 0),
    "MI": (6, 0), "NY": (7, 0), "VT": (8, 0), "ME": (9, 0),

    # Row 1
    "OR": (0, 1), "ID": (1, 1), "WY": (2, 1), "SD": (3, 1), "IA": (4, 1),
    "IL": (5, 1), "IN": (6, 1), "OH": (7, 1), "PA": (8, 1), "NH": (9, 1),
    "MA": (8, 0),  # MA overlaps with NY row — adjust
    "RI": (9, 2), "CT": (8, 2), "NJ": (9, 3),

    # Row 2
    "CA": (0, 2), "NV": (1, 2), "CO": (2, 2), "NE": (3, 2), "MO": (4, 2),
    "KY": (5, 2), "WV": (6, 2), "VA": (7, 2), "MD": (8, 3), "DE": (9, 4),
    "DC": (8, 4),

    # Row 3
    "AZ": (1, 3), "UT": (1, 2),  # UT same slot as NV — fix below
    "NM": (2, 3), "KS": (3, 3), "AR": (4, 3), "TN": (5, 3),
    "NC": (7, 3), "SC": (8, 5),

    # Row 4
    "OK": (3, 4), "LA": (4, 4), "MS": (5, 4), "AL": (6, 4), "GA": (7, 4),
    "FL": (8, 6),

    # Row 5
    "TX": (3, 5),

    # Non-contiguous
    "AK": (0, 6), "HI": (1, 6),
}

# Use a clean, well-tested grid (no duplicate slots)
HEX_GRID = {
    "WA": (0, 0),  "MT": (1, 0),  "ND": (2, 0),  "MN": (3, 0),
    "WI": (4, 0),  "MI": (6, 0),  "NY": (7, 0),  "VT": (8, 0),  "ME": (9, 0),

    "OR": (0, 1),  "ID": (1, 1),  "WY": (2, 1),  "SD": (3, 1),
    "IA": (4, 1),  "IL": (5, 1),  "IN": (6, 1),  "OH": (7, 1),
    "PA": (8, 1),  "NH": (9, 1),

    "CA": (0, 2),  "NV": (1, 2),  "UT": (2, 2),  "CO": (3, 2),  "NE": (4, 2),
    "MO": (5, 2),  "KY": (6, 2),  "WV": (7, 2),  "VA": (8, 2),  "MA": (9, 2),

    "AZ": (1, 3),  "NM": (2, 3),  "KS": (3, 3),  "AR": (4, 3),
    "TN": (5, 3),  "NC": (6, 3),  "SC": (7, 3),  "MD": (8, 3),  "CT": (9, 3),

    "OK": (3, 4),  "LA": (4, 4),  "MS": (5, 4),  "AL": (6, 4),
    "GA": (7, 4),  "DE": (8, 4),  "NJ": (9, 4),

    "TX": (3, 5),  "FL": (7, 5),  "RI": (9, 5),

    "AK": (0, 6),  "HI": (1, 6),
}


def _lerp_color(c1, c2, t):
    return tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))


def _count_to_color(count, max_count):
    if count == 0:
        return (217, 217, 217, 255)  # gray for zero
    light = (234, 244, 238)   # #EAF4EE
    dark  = (26,  71,  49)    # #1A4731
    t = math.sqrt(count / max_count) if max_count > 0 else 0
    t = min(1.0, max(0.05, t))
    rgb = _lerp_color(light, dark, t)
    return (*rgb, 255)


def _label_color(count, max_count):
    if count == 0:
        return (100, 100, 100, 255)
    t = math.sqrt(count / max_count) if max_count > 0 else 0
    return (255, 255, 255, 255) if t > 0.3 else (26, 71, 49, 255)


def _flat_hex_points(cx, cy, size):
    """6 vertices of a flat-top hexagon."""
    pts = []
    for i in range(6):
        a = math.radians(60 * i)
        pts.append((cx + size * math.cos(a), cy + size * math.sin(a)))
    return pts


def generate_hex_map_png(state_counts: dict, width: int = 1180, height: int = 860) -> bytes:
    """
    Returns PNG bytes of a US hex grid heat map colored by student count.

    state_counts: dict of {"TX": 414, "CA": 48, ...}
    """
    img = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    counts = {k.upper(): int(v) for k, v in state_counts.items()}
    max_count = max((v for v in counts.values() if v > 0), default=1)

    # Grid dimensions: 10 cols (0-9), 7 rows (0-6)
    n_cols, n_rows = 10, 7
    margin_x, margin_y = 40, 30

    avail_w = width - 2 * margin_x
    avail_h = height - 2 * margin_y

    # Flat-top hex: col spacing = size*1.5, row spacing = size*sqrt(3)
    size_w = avail_w / (1.5 * n_cols + 0.5)
    size_h = avail_h / (math.sqrt(3) * n_rows)
    size = min(size_w, size_h) * 0.93

    sp_x = size * 1.5
    sp_y = size * math.sqrt(3)

    grid_w = sp_x * (n_cols - 1) + size * 2
    grid_h = sp_y * (n_rows - 1)
    ox = (width - grid_w) / 2 + size
    oy = (height - grid_h) / 2 + sp_y / 2

    def to_px(col, row):
        x = ox + col * sp_x
        y = oy + row * sp_y + (sp_y / 2 if col % 2 == 1 else 0)
        return x, y

    # Fonts
    try:
        font_ab = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", int(size * 0.36))
        font_ct = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", int(size * 0.27))
    except Exception:
        font_ab = ImageFont.load_default()
        font_ct = font_ab

    for state, (col, row) in HEX_GRID.items():
        cx, cy = to_px(col, row)
        count = counts.get(state, 0)
        fill = _count_to_color(count, max_count)
        txt_col = _label_color(count, max_count)

        pts = _flat_hex_points(cx, cy, size * 0.96)
        draw.polygon(pts, fill=fill)
        draw.polygon(pts, outline=(180, 180, 180, 180))

        # Abbreviation label
        bb = draw.textbbox((0, 0), state, font=font_ab)
        tw, th = bb[2]-bb[0], bb[3]-bb[1]
        draw.text((cx - tw/2, cy - th/2 - size*0.07), state, font=font_ab, fill=txt_col)

        # Count label (non-zero only)
        if count > 0:
            cs = str(count)
            bb2 = draw.textbbox((0, 0), cs, font=font_ct)
            tw2, th2 = bb2[2]-bb2[0], bb2[3]-bb2[1]
            draw.text((cx - tw2/2, cy + size*0.1), cs, font=font_ct, fill=txt_col)

    # Flatten to white background
    bg = Image.new("RGB", img.size, (255, 255, 255))
    bg.paste(img, mask=img.split()[3])

    buf = io.BytesIO()
    bg.save(buf, format="PNG", dpi=(150, 150))
    return buf.getvalue()


if __name__ == "__main__":
    sample = {
        "TX": 414, "CA": 48, "CO": 19, "IL": 15, "AZ": 14,
        "WA": 10, "KS": 9, "MN": 8, "NY": 8, "MO": 7,
        "FL": 6, "GA": 5, "NC": 5, "VA": 4, "OH": 4,
        "PA": 3, "NJ": 3, "TN": 2, "OR": 2, "NE": 2, "ID": 1,
    }
    data = generate_hex_map_png(sample)
    with open("/home/claude/hex_map_test.png", "wb") as f:
        f.write(data)
    print(f"Saved hex_map_test.png ({len(data):,} bytes)")
