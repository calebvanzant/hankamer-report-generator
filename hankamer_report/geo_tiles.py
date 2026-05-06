"""
geo_tiles.py

Generates the Option D "state tile grid" PNG for slide 3.
Dark green background, TX hero tile, then all other states as small tiles.

Usage:
    from geo_tiles import generate_geo_tiles_png
    png_bytes = generate_geo_tiles_png(state_rows, regional_rows, config)
"""

import io
import math
from PIL import Image, ImageDraw, ImageFont

DARK_GREEN  = (26,  71, 49)    # #1A4731
MID_GREEN   = (45, 100, 70)    # tile backgrounds
LIGHT_TILE  = (60, 120, 88)    # smaller tiles
GOLD        = (245, 184, 0)    # #F5B800
WHITE       = (255, 255, 255)
DIM_WHITE   = (200, 220, 208)
VERY_DIM    = (140, 175, 155)


def _font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()


BOLD_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
REG_PATH  = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"


def _text_size(draw, text, font):
    bb = draw.textbbox((0, 0), text, font=font)
    return bb[2] - bb[0], bb[3] - bb[1]


def _rounded_rect(draw, x0, y0, x1, y1, radius, fill):
    draw.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=fill)


def generate_geo_tiles_png(
    state_rows: list,   # list of (abbr, count) sorted largest first
    width: int = 1120,
    height: int = 840,
    grand_total: int = 0,  # pass attendance grand total for accurate TX%
) -> bytes:
    """
    state_rows: list of (abbr, count) — all states with students, TX first.
    grand_total: if provided, used for TX% calculation instead of sum of state_rows.
    Returns PNG bytes.
    """
    img = Image.new("RGB", (width, height), DARK_GREEN)
    draw = ImageDraw.Draw(img)

    if not state_rows:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    # Split TX from the rest
    tx_count = 0
    others = []
    total = grand_total if grand_total > 0 else sum(c for _, c in state_rows if c)
    for abbr, count in state_rows:
        if not abbr or not count:
            continue
        if abbr.upper() == "TX":
            tx_count = count
        else:
            others.append((abbr.upper(), int(count)))

    tx_pct = round(tx_count / total * 100) if total else 0

    pad = 28

    # --- TX hero tile ---
    tx_h = int(height * 0.22)
    tx_w = width - pad * 2
    tx_y0 = pad
    tx_y1 = tx_y0 + tx_h
    _rounded_rect(draw, pad, tx_y0, pad + tx_w, tx_y1, 10, MID_GREEN)

    f_tx_label = _font(REG_PATH,  int(tx_h * 0.22))
    f_tx_num   = _font(BOLD_PATH, int(tx_h * 0.52))
    f_tx_pct   = _font(REG_PATH,  int(tx_h * 0.18))

    # "Texas" label top-left
    draw.text((pad + 18, tx_y0 + 12), "Texas", font=f_tx_label, fill=DIM_WHITE)
    # Big count in gold, vertically centered
    num_str = f"{tx_count:,}"
    nw, nh = _text_size(draw, num_str, f_tx_num)
    draw.text((pad + 18, tx_y0 + tx_h//2 - nh//2), num_str, font=f_tx_num, fill=GOLD)
    # Pct label right side
    pct_str = f"{tx_pct}% of all students"
    pw, ph = _text_size(draw, pct_str, f_tx_pct)
    draw.text((pad + tx_w - pw - 18, tx_y1 - ph - 14), pct_str, font=f_tx_pct, fill=VERY_DIM)

    # --- Other state tiles ---
    grid_y0 = tx_y1 + 16
    grid_h  = height - grid_y0 - pad
    grid_w  = width - pad * 2

    n = len(others)
    if n == 0:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    # Choose cols to fit nicely — aim for tiles ~proportional
    cols = min(n, max(4, math.ceil(math.sqrt(n * 1.6))))
    rows_needed = math.ceil(n / cols)

    gap = 6
    tile_w = (grid_w - gap * (cols - 1)) // cols
    tile_h = (grid_h - gap * (rows_needed - 1)) // rows_needed
    tile_h = min(tile_h, int(tile_w * 0.9))  # keep roughly square-ish

    f_abbr  = _font(BOLD_PATH, max(9, int(tile_h * 0.30)))
    f_count = _font(REG_PATH,  max(8, int(tile_h * 0.26)))

    max_others = max(c for _, c in others) if others else 1

    for i, (abbr, count) in enumerate(others):
        col = i % cols
        row = i // cols
        x0 = pad + col * (tile_w + gap)
        y0 = grid_y0 + row * (tile_h + gap)
        x1 = x0 + tile_w
        y1 = y0 + tile_h

        # Slightly lighter tile for higher counts
        t = math.sqrt(count / max_others)
        r = int(LIGHT_TILE[0] + (MID_GREEN[0] - LIGHT_TILE[0]) * (1 - t))
        g = int(LIGHT_TILE[1] + (MID_GREEN[1] - LIGHT_TILE[1]) * (1 - t))
        b = int(LIGHT_TILE[2] + (MID_GREEN[2] - LIGHT_TILE[2]) * (1 - t))
        _rounded_rect(draw, x0, y0, x1, y1, 5, (r, g, b))

        # Abbr centered top
        aw, ah = _text_size(draw, abbr, f_abbr)
        draw.text((x0 + (tile_w - aw) // 2, y0 + int(tile_h * 0.12)), abbr,
                  font=f_abbr, fill=WHITE)

        # Count centered bottom in gold
        cs = str(count)
        cw, ch = _text_size(draw, cs, f_count)
        draw.text((x0 + (tile_w - cw) // 2, y1 - ch - int(tile_h * 0.12)), cs,
                  font=f_count, fill=GOLD)

    buf = io.BytesIO()
    img.save(buf, format="PNG", dpi=(150, 150))
    return buf.getvalue()


if __name__ == "__main__":
    sample = [
        ("TX", 414), ("CA", 48), ("CO", 19), ("IL", 15), ("AZ", 14),
        ("WA", 10), ("KS", 9), ("MN", 8), ("NY", 8), ("MO", 7),
        ("FL", 6), ("GA", 5), ("NC", 5), ("VA", 4), ("OH", 4),
        ("PA", 3), ("NJ", 3), ("TN", 2), ("OR", 2), ("NE", 2),
        ("ID", 1), ("VT", 2), ("NH", 2), ("MA", 3), ("CT", 2),
        ("AR", 7), ("OK", 6), ("SC", 3), ("AL", 2), ("IN", 6),
        ("WI", 3), ("IA", 2), ("SD", 1), ("WY", 1), ("NM", 4),
        ("UT", 2), ("NV", 1), ("MT", 1),
    ]
    data = generate_geo_tiles_png(sample)
    with open("/home/claude/geo_tiles_test.png", "wb") as f:
        f.write(data)
    print(f"Saved geo_tiles_test.png ({len(data):,} bytes)")
