"""
slide_visuals.py

Generates PNG visuals for slides 3 and 4:
  - top10_bars_png()      : horizontal bar chart of top 10 states (slide 3)
  - companion_dots_png()  : dot-scale companion frequency chart (slide 4)
  - large_events_png()    : large events summary panel (slide 4)
"""

import io
import math
from PIL import Image, ImageDraw, ImageFont

DARK_GREEN  = (26,  71,  49)
MID_GREEN   = (45, 100,  70)
LIGHT_GREEN = (234, 244, 238)
GOLD        = (245, 184,   0)
WHITE       = (255, 255, 255)
GRAY_BG     = (245, 247, 245)
GRAY_TEXT   = (100, 110, 105)
DARK_TEXT   = ( 30,  40,  35)

BOLD_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
REG_PATH  = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"

FREQ_ORDER = {
    "very common": 5,
    "common":      4,
    "occasional":  3,
    "rare":        2,
    "uncommon":    1,
}
FREQ_LABEL = {5:"Very common", 4:"Common", 3:"Occasional", 2:"Rare", 1:"Uncommon"}


def _font(path, size):
    try:
        return ImageFont.truetype(path, max(1, int(size)))
    except Exception:
        return ImageFont.load_default()


def _tw(draw, text, font):
    bb = draw.textbbox((0, 0), text, font=font)
    return bb[2] - bb[0], bb[3] - bb[1]


def _rr(draw, x0, y0, x1, y1, r, fill, outline=None):
    draw.rounded_rectangle([x0, y0, x1, y1], radius=max(0,r), fill=fill, outline=outline)


# ── TOP 10 STATES BAR CHART ─────────────────────────────────────────────────

def top10_bars_png(state_rows, width=1100, height=860):
    """
    state_rows: list of (abbr, count) sorted largest→smallest.
    Returns PNG bytes. Dark green bars on white background.
    """
    img = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(img)

    rows = [(a, int(c)) for a, c in state_rows if a and c][:10]
    if not rows:
        buf = io.BytesIO(); img.save(buf, "PNG"); return buf.getvalue()

    max_cnt = rows[0][1]
    n = len(rows)

    pad_l = int(width * 0.10)   # left for state label
    pad_r = int(width * 0.12)   # right for count label
    pad_t = int(height * 0.06)
    pad_b = int(height * 0.06)

    bar_area_w = width - pad_l - pad_r
    total_h    = height - pad_t - pad_b
    bar_h      = int(total_h / n * 0.58)
    gap        = int(total_h / n * 0.42)
    bar_h      = min(bar_h, 72)

    f_label = _font(BOLD_PATH, bar_h * 0.52)
    f_count = _font(REG_PATH,  bar_h * 0.42)
    f_title = _font(BOLD_PATH, int(height * 0.040))

    # Title
    title = "Top States by Student Count"
    tw, th = _tw(draw, title, f_title)
    draw.text(((width - tw) // 2, int(pad_t * 0.2)), title, font=f_title, fill=DARK_GREEN)

    for i, (abbr, cnt) in enumerate(rows):
        y0 = pad_t + i * (bar_h + gap)
        y1 = y0 + bar_h
        cx = (y0 + y1) // 2

        # State label (right-aligned)
        lw, lh = _tw(draw, abbr, f_label)
        draw.text((pad_l - lw - 10, cx - lh // 2), abbr, font=f_label, fill=DARK_TEXT)

        # Bar
        bar_w = int(bar_area_w * cnt / max_cnt)
        # gradient: TX full dark, others proportionally lighter
        t = math.sqrt(cnt / max_cnt)
        r = int(LIGHT_GREEN[0] + (DARK_GREEN[0] - LIGHT_GREEN[0]) * t)
        g = int(LIGHT_GREEN[1] + (DARK_GREEN[1] - LIGHT_GREEN[1]) * t)
        b = int(LIGHT_GREEN[2] + (DARK_GREEN[2] - LIGHT_GREEN[2]) * t)
        _rr(draw, pad_l, y0, pad_l + bar_w, y1, 4, (r, g, b))

        # Count label
        cs = f"{cnt:,}"
        cw, ch = _tw(draw, cs, f_count)
        cx2 = pad_l + bar_w + 8
        draw.text((cx2, cx - ch // 2), cs, font=f_count, fill=GRAY_TEXT)

    buf = io.BytesIO()
    img.save(buf, "PNG", dpi=(150,150))
    return buf.getvalue()


# ── COMPANION DOT-SCALE CHART ────────────────────────────────────────────────

def companion_dots_png(comp_rows, width=900, height=760):
    """
    comp_rows: list of (companion_type, frequency_label).
    Draws a dot-scale chart — 5 filled dots = Very common, 1 = Uncommon.
    """
    img = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(img)

    rows = [(ct, fl) for ct, fl in comp_rows if ct and fl]
    if not rows:
        buf = io.BytesIO(); img.save(buf, "PNG"); return buf.getvalue()

    n = len(rows)
    pad_t = int(height * 0.08)
    pad_b = int(height * 0.05)
    pad_l = int(width  * 0.04)
    pad_r = int(width  * 0.04)

    row_h    = (height - pad_t - pad_b) // n
    dot_area = int(width * 0.38)   # right portion for dots
    label_w  = width - pad_l - pad_r - dot_area

    DOT_COUNT = 5
    dot_r     = int(min(row_h * 0.28, dot_area / (DOT_COUNT * 2.8)))
    dot_gap   = int(dot_area / DOT_COUNT)

    f_label = _font(BOLD_PATH, int(row_h * 0.34))
    f_freq  = _font(REG_PATH,  int(row_h * 0.26))
    f_title = _font(BOLD_PATH, int(height * 0.050))

    # Column headers
    ty = int(pad_t * 0.15)
    draw.text((pad_l, ty), "Who Joins Students", font=f_title, fill=DARK_GREEN)

    dot_x0 = pad_l + label_w

    # Header dots label
    hw, hh = _tw(draw, "Frequency", f_freq)
    draw.text((dot_x0 + (dot_area - hw) // 2, ty), "Frequency", font=f_freq, fill=GRAY_TEXT)

    # Alternating row backgrounds
    for i, (ctype, freq_label) in enumerate(rows):
        y0 = pad_t + i * row_h
        y1 = y0 + row_h
        cy = (y0 + y1) // 2

        # Alternating stripe
        if i % 2 == 0:
            draw.rectangle([0, y0, width, y1], fill=(248, 251, 249))

        # Companion label
        lw, lh = _tw(draw, ctype, f_label)
        draw.text((pad_l, cy - lh // 2), ctype, font=f_label, fill=DARK_TEXT)

        # Frequency text (dimmed)
        level = FREQ_ORDER.get(freq_label.lower().strip(), 0)
        freq_disp = FREQ_LABEL.get(level, freq_label)
        fw, fh = _tw(draw, freq_disp, f_freq)
        # draw.text((pad_l + label_w - fw - 12, cy - fh//2), freq_disp, font=f_freq, fill=GRAY_TEXT)

        # Dots
        for d in range(DOT_COUNT):
            dx = dot_x0 + d * dot_gap + dot_gap // 2
            filled = d < level
            fill_c = DARK_GREEN if filled else LIGHT_GREEN
            draw.ellipse([dx - dot_r, cy - dot_r, dx + dot_r, cy + dot_r],
                         fill=fill_c, outline=(180, 200, 190))

        # Divider line
        draw.line([(0, y1), (width, y1)], fill=(220, 230, 225), width=1)

    buf = io.BytesIO()
    img.save(buf, "PNG", dpi=(150, 150))
    return buf.getvalue()


# ── LARGE EVENTS PANEL ───────────────────────────────────────────────────────

def large_events_png(events, width=860, height=760):
    """
    events: list of dicts with keys: name, students, total, note.
    Returns PNG bytes.
    """
    img = Image.new("RGB", (width, height), DARK_GREEN)
    draw = ImageDraw.Draw(img)

    if not events:
        buf = io.BytesIO(); img.save(buf, "PNG"); return buf.getvalue()

    pad = 22
    n   = len(events)

    f_title    = _font(BOLD_PATH, int(height * 0.050))
    f_stat_num = _font(BOLD_PATH, int(height * 0.068))
    f_stat_lbl = _font(REG_PATH,  int(height * 0.030))
    f_note     = _font(REG_PATH,  int(height * 0.032))

    # Title
    title = "LARGE-SCALE EVENTS"
    tw, th = _tw(draw, title, f_title)
    draw.text(((width - tw) // 2, pad), title, font=f_title, fill=GOLD)

    card_h  = int((height - pad * 2 - th - 16) / n) - 8
    card_w  = width - pad * 2
    card_y0 = pad + th + 14

    # Layout constants: left 48% = name zone, right 52% = stats
    name_zone_w = int(card_w * 0.46)
    stat_zone_x = pad + name_zone_w + 10
    stat_zone_w = card_w - name_zone_w - 10
    stat_col_w  = stat_zone_w // 2

    for i, ev in enumerate(events):
        y0 = card_y0 + i * (card_h + 8)
        y1 = y0 + card_h
        _rr(draw, pad, y0, pad + card_w, y1, 8, MID_GREEN)

        name     = ev.get("name", "")
        note     = ev.get("note", "")
        students = ev.get("students")
        total    = ev.get("total")

        # Name — auto-shrink to fit name_zone_w
        f_name_base = int(card_h * 0.38)
        fn = _font(BOLD_PATH, f_name_base)
        nw, nh = _tw(draw, name, fn)
        while nw > name_zone_w - 16 and f_name_base > 10:
            f_name_base = int(f_name_base * 0.88)
            fn = _font(BOLD_PATH, f_name_base)
            nw, nh = _tw(draw, name, fn)

        name_y = y0 + int(card_h * 0.14)
        draw.text((pad + 14, name_y), name, font=fn, fill=WHITE)

        if note:
            nrow, noh = _tw(draw, note, f_note)
            draw.text((pad + 14, name_y + nh + 5), note, font=f_note, fill=(170, 205, 185))

        # Stats — right zone, vertically centered
        stat_cols = [(students, "students"), (total, "total visitors")]
        stat_cols = [(v, l) for v, l in stat_cols if v is not None]
        for j, (val, lbl) in enumerate(stat_cols):
            sx     = stat_zone_x + j * stat_col_w
            vs     = f"{val:,}"
            vw, vh = _tw(draw, vs, f_stat_num)
            lw, lh = _tw(draw, lbl, f_stat_lbl)
            sy_num = y0 + int(card_h * 0.14)
            sy_lbl = y0 + int(card_h * 0.60)
            draw.text((sx + (stat_col_w - vw) // 2, sy_num), vs, font=f_stat_num, fill=GOLD)
            draw.text((sx + (stat_col_w - lw)  // 2, sy_lbl), lbl, font=f_stat_lbl, fill=(175, 208, 188))

    buf = io.BytesIO()
    img.save(buf, "PNG", dpi=(150, 150))
    return buf.getvalue()


# ── test ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Top 10 bars
    states = [("TX",414),("CA",48),("CO",19),("IL",15),("AZ",14),
              ("WA",10),("KS",9),("MN",8),("NY",8),("MO",7)]
    open("/home/claude/test_top10.png","wb").write(top10_bars_png(states))
    print("Saved test_top10.png")

    # Companion dots
    comp = [("Mother (solo)","Very common"),("Father (solo)","Common"),
            ("Both parents","Very common"),("Sibling(s)","Occasional"),
            ("Extended family","Rare"),("Friends / classmates","Rare"),
            ("Student alone","Uncommon")]
    open("/home/claude/test_companion.png","wb").write(companion_dots_png(comp))
    print("Saved test_companion.png")

    # Large events
    evts = [
        {"name":"Fall Premiere",         "students":445,"total":820},
        {"name":"Spring Premiere",        "students":445,"total":820},
        {"name":"Baylor Scholars Days",   "students":100,"total":None,"note":"2 events, ~100 students each"},
        {"name":"Invitation 2 Excellence","students":204,"total":None,"note":"2 events, 204 total students"},
    ]
    open("/home/claude/test_events.png","wb").write(large_events_png(evts))
    print("Saved test_events.png")
