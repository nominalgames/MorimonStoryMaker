"""Builds the Morimon adventure PDF straight from Python (ReportLab).

The layout reproduces the pages of the old Adventure_Template.docx (its cover page is no longer used): 5.5" x 8.5"
pages, 0.2" margins, Georgia / Cambria / Consolas type, the same colours, boxes,
rules and running header/footer. Nothing here needs Word or LibreOffice.

All positions below are written as distances from the TOP of the page (in points)
and converted by `_y()`, because that's how they were measured from the docx render.
"""
import os

from reportlab.lib.colors import HexColor, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

from location_reveal import get_location_reveal

PAGE_W, PAGE_H = 396.0, 612.0  # 5.5" x 8.5"
MARGIN = 14.4  # 0.2"
CONTENT_W = PAGE_W - 2 * MARGIN
RIGHT = PAGE_W - MARGIN
MAX_IMAGE_H = 4.5 * 72

# ---- colours (same values as the docx template) ----
DARK = HexColor("#2A221C")
BODY = HexColor("#4A4642")
LABEL = HexColor("#8A8682")
BLUE = HexColor("#1B5E8C")
LIGHT_ON_BLUE = HexColor("#CFE3EF")
BOX_BG = HexColor("#F2F1EF")
BOX_DASH = HexColor("#B5B0AA")
BOX_TEXT = HexColor("#A8A49E")
GREEN, GREEN_BG = HexColor("#4C7A4C"), HexColor("#EAF4EA")
AMBER, AMBER_BG = HexColor("#C98A2E"), HexColor("#FBF2E0")
RED, RED_BG = HexColor("#C0532D"), HexColor("#FBE9E3")
ROW_BG, ROW_LINE = HexColor("#F7F6F4"), HexColor("#D8D4CE")
CELL_LINE = HexColor("#E2DFDA")

# ---- fonts ----
_FONT_FILES = {
    "Georgia": ("georgia.ttf", "Times-Roman"),
    "Georgia-Bold": ("georgiab.ttf", "Times-Bold"),
    "Cambria": ("cambria.ttc", "Times-Roman"),
    "Cambria-Bold": ("cambriab.ttf", "Times-Bold"),
    "Cambria-Italic": ("cambriai.ttf", "Times-Italic"),
    "Consolas": ("consola.ttf", "Courier"),
    "Consolas-Bold": ("consolab.ttf", "Courier-Bold"),
}
_fonts_ready = False
_font_map = {}


def _register_fonts():
    """Use the real Windows fonts when present; otherwise fall back to the
    built-in PDF fonts (Times / Courier) so the app still works anywhere."""
    global _fonts_ready
    if _fonts_ready:
        return
    windir = os.environ.get("WINDIR", r"C:\Windows")
    dirs = [os.path.join(windir, "Fonts"), os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")]
    for name, (filename, fallback) in _FONT_FILES.items():
        _font_map[name] = fallback
        for d in dirs:
            path = os.path.join(d, filename)
            if os.path.exists(path):
                try:
                    pdfmetrics.registerFont(TTFont(name, path, subfontIndex=0))
                    _font_map[name] = name
                    break
                except Exception:
                    continue
    _fonts_ready = True


def _f(name):
    return _font_map[name]


def _y(top):
    return PAGE_H - top


def _tracked(text):
    """The template writes small labels as spaced-out capitals ("P A R T   O N E")."""
    return " ".join(text.upper())


def _width(text, font, size):
    return pdfmetrics.stringWidth(text, _f(font), size)


def _text(c, x, base, s, font, size, color, align="l"):
    c.setFillColor(color)
    c.setFont(_f(font), size)
    if align == "l":
        c.drawString(x, _y(base), s)
    elif align == "r":
        c.drawRightString(x, _y(base), s)
    else:
        c.drawCentredString(x, _y(base), s)


def _hline(c, x0, x1, top, color, width):
    c.setStrokeColor(color)
    c.setLineWidth(width)
    c.line(x0, _y(top), x1, _y(top))


def _rect(c, x, top, w, h, fill=None, stroke=None, line=0.5, dash=None):
    c.saveState()
    if fill is not None:
        c.setFillColor(fill)
    if stroke is not None:
        c.setStrokeColor(stroke)
        c.setLineWidth(line)
    if dash:
        c.setDash(*dash)
    c.rect(x, _y(top + h), w, h, fill=1 if fill is not None else 0, stroke=1 if stroke is not None else 0)
    c.restoreState()


def _diamond(c, cx, base, color, half=2.1):
    cy = _y(base) + 2.5
    p = c.beginPath()
    p.moveTo(cx, cy + half)
    p.lineTo(cx + half, cy)
    p.lineTo(cx, cy - half)
    p.lineTo(cx - half, cy)
    p.close()
    c.setFillColor(color)
    c.drawPath(p, fill=1, stroke=0)


# ---- running header / footer ----

def _header(c, label, dark=True, title="Template"):
    left_color = DARK if dark else white
    label_color = LABEL if dark else LIGHT_ON_BLUE
    rule = DARK if dark else LIGHT_ON_BLUE
    right = _tracked(label)
    left = _tracked(title)
    room = CONTENT_W - _width(right, "Consolas", 7) - 12
    if _width(left, "Consolas-Bold", 7) > room:
        left = title.upper()  # too long to space out, so drop the letter spacing
    while len(left) > 1 and _width(left, "Consolas-Bold", 7) > room:
        left = left[:-2] + "…"  # still too long (very long name): shorten it
    _text(c, MARGIN, 20.8, left, "Consolas-Bold", 7, left_color)
    _text(c, RIGHT, 20.8, right, "Consolas", 7, label_color, "r")
    _hline(c, MARGIN, RIGHT, 27.0, rule, 0.75)


def _footer(c, label, page_number=None, dark=True, base=595.8):
    color = LABEL if dark else LIGHT_ON_BLUE
    text = _tracked(label)
    gap = _width("  ", "Consolas", 7)
    dia = 4.2
    num = str(c.getPageNumber()) if page_number else ""
    total = _width(text, "Consolas", 7) + gap + dia + (gap + _width(num, "Consolas", 7) if num else 0)
    x = PAGE_W / 2 - total / 2
    _text(c, x, base, text, "Consolas", 7, color)
    x += _width(text, "Consolas", 7) + gap
    _diamond(c, x + dia / 2, base, color if dark else LIGHT_ON_BLUE)
    x += dia + gap
    if num:
        _text(c, x, base, num, "Consolas", 7, color)


def _eyebrow(c, text, base=34.2, size=7.5):
    _text(c, MARGIN, base, text.upper(), "Consolas-Bold", size, BLUE)


def _heading(c, text, base=56.6, size=19):
    """Draws the page heading and returns how far everything below it must move down.
    Long text (e.g. a very long player name) shrinks first, then wraps to more lines."""
    while size > 14 and _width(text, "Georgia-Bold", size) > CONTENT_W:
        size -= 0.5
    if _width(text, "Georgia-Bold", size) <= CONTENT_W:
        _text(c, MARGIN, base, text, "Georgia-Bold", size, DARK)
        return 0.0
    lines = _wrap(text, "Georgia-Bold", size, CONTENT_W, CONTENT_W)
    leading = size * 1.2
    for i, line in enumerate(lines):
        _text(c, MARGIN, base + i * leading, line, "Georgia-Bold", size, DARK)
    return leading * (len(lines) - 1)


def _wrap(text, font, size, first_width, width):
    """Greedy word wrap. The first line may be narrower (room for a drop cap)."""
    lines, current, limit = [], "", first_width
    for word in text.split():
        trial = word if not current else current + " " + word
        if _width(trial, font, size) <= limit or not current:
            current = trial
        else:
            lines.append(current)
            current, limit = word, width
    if current:
        lines.append(current)
    return lines


def _placeholder_box(c, top, height, label="[ image placeholder ]"):
    _rect(c, MARGIN + 0.3, top, CONTENT_W - 0.6, height, fill=BOX_BG, stroke=BOX_DASH, line=0.75, dash=(3, 2))
    _text(c, PAGE_W / 2, top + 11.6, label, "Consolas", 7.5, BOX_TEXT, "c")


# ---- pages ----

# ---- location reveal (heading + text + icon bullets) ----
# Bullet design follows the zine's "icon-list": a blue circular badge with a white
# icon, then a bold serif title. Sizes are the zine's CSS pixels scaled to this page.
BADGE_D = 19.0  # badge diameter
BULLET_INDENT = 28.0  # badge + gap before the text
BULLET_GAP = 8.0  # space between bullets


def _pin_badge(c, x, top, d=BADGE_D):
    """Blue circle with the zine's white map-pin icon (24x24 SVG path, y-down)."""
    cx, cy = x + d / 2, _y(top + d / 2)
    c.setFillColor(BLUE)
    c.circle(cx, cy, d / 2, stroke=0, fill=1)

    s = (d * 0.57) / 24.0  # icon is ~57% of the badge, like 16px in 28px

    def px(sx):
        return cx + (sx - 12) * s

    def py(sy):
        return cy - (sy - 12) * s

    k = 0.5523 * 7
    p = c.beginPath()
    p.moveTo(px(12), py(21))
    p.curveTo(px(12), py(21), px(5), py(16.4), px(5), py(11))
    p.curveTo(px(5), py(11 - k), px(12 - k), py(4), px(12), py(4))
    p.curveTo(px(12 + k), py(4), px(19), py(11 - k), px(19), py(11))
    p.curveTo(px(19), py(16.4), px(12), py(21), px(12), py(21))
    p.close()
    c.setFillColor(white)
    c.drawPath(p, fill=1, stroke=0)
    c.setFillColor(BLUE)
    c.circle(px(12), py(10.6), 2.6 * s, stroke=0, fill=1)


def draw_location_reveal(c, top, reveal):
    """Draw one location reveal starting at `top` (distance from the top of the page):
    a heading with the location's name, its reveal text, then one icon bullet per
    sub-location. `reveal` is a location_reveal.LocationReveal.

    Returns the y of the bottom of what it drew, so the caller can keep flowing content
    below it, or call this again to stack another reveal.
    """
    base = top + 14
    extra = _heading(c, reveal.name, base=base)  # same style as the page heading
    y = base + extra + 16

    if reveal.reveal_text:
        for line in _wrap(reveal.reveal_text, "Cambria", 8.5, CONTENT_W, CONTENT_W):
            _text(c, MARGIN, y, line, "Cambria", 8.5, BODY)
            y += 9.94
        y = y - 9.94 + 3  # bottom edge of the paragraph
    else:
        y = base + extra + 3

    bottom = y
    y += 12
    text_w = CONTENT_W - BULLET_INDENT
    for name in reveal.sub_locations:
        lines = _wrap(name, "Georgia-Bold", 10.5, text_w, text_w)
        _pin_badge(c, MARGIN, y)
        for i, line in enumerate(lines):
            _text(c, MARGIN + BULLET_INDENT, y + 13.1 + i * 12.6, line, "Georgia-Bold", 10.5, DARK)
        row_h = max(BADGE_D, 13.1 + (len(lines) - 1) * 12.6 + 4)
        bottom = y + row_h
        y = bottom + BULLET_GAP
    return bottom


def _page_content(c, opening_text, player_name, location_id, image_path, title):
    reveal = get_location_reveal(location_id)
    location_name = reveal.name
    _header(c, location_name, title=title)
    _eyebrow(c, "Section label")
    dy = _heading(c, f"{player_name}'s Journey Begins")

    # opening paragraph with a blue drop cap
    opening = opening_text.replace("[Player_Name]", player_name).strip()
    cap, rest = opening[0], opening[1:]
    cap_w = _width(cap, "Georgia-Bold", 22)
    _text(c, MARGIN, 87.0 + dy, cap, "Georgia-Bold", 22, BLUE)
    lines = _wrap(rest, "Cambria", 8.5, CONTENT_W - cap_w, CONTENT_W)
    base = 87.0 + dy
    for i, line in enumerate(lines):
        _text(c, MARGIN + cap_w if i == 0 else MARGIN, base, line, "Cambria", 8.5, BODY)
        base += 12.9 if i == 0 else 9.94

    # location picture: full text width, centred, never taller than 4.5"
    img_top = base - 9.94 + 9.8
    reader = ImageReader(image_path)
    iw, ih = reader.getSize()
    w = CONTENT_W
    h = w * ih / iw
    if h > MAX_IMAGE_H:
        h = MAX_IMAGE_H
        w = h * iw / ih
    c.drawImage(reader, PAGE_W / 2 - w / 2, _y(img_top + h), width=w, height=h, mask="auto")
    _text(c, RIGHT, img_top + h + 8.0, location_name.upper(), "Consolas", 6.5, LABEL, "r")

    # reveal the starting location: heading, reveal text and sub-location bullets
    caption_base = img_top + h + 8.0
    reveal_bottom = draw_location_reveal(c, caption_base + 14, reveal)

    para = _wrap("A second paragraph can follow the image placeholder, continuing the "
                 "section's explanation or adding a supporting detail.", "Cambria", 8.5, CONTENT_W, CONTENT_W)
    base = reveal_bottom + 18
    for line in para:
        _text(c, MARGIN, base, line, "Cambria", 8.5, BODY)
        base += 9.94
    _footer(c, location_name, True)


def _page_divider(c, title):
    c.setFillColor(BLUE)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)  # full-bleed
    _header(c, "Part One", dark=False, title=title)
    _text(c, MARGIN, 224.7, _tracked("Part One"), "Consolas-Bold", 8, LIGHT_ON_BLUE)
    _text(c, MARGIN, 254.6, "Section Title", "Georgia-Bold", 26, white)
    _text(c, MARGIN, 276.4, "A short paragraph introducing this part of the document sits below the big title.",
          "Cambria", 8.5, LIGHT_ON_BLUE)
    for i, label in enumerate(["First Subsection", "Second Subsection", "Third Subsection"]):
        base = 294.6 + i * 9.2
        _text(c, MARGIN, base, f"0{i + 1}", "Consolas-Bold", 7, white)
        _text(c, 29.9, base, _tracked(label), "Consolas", 7, LIGHT_ON_BLUE)
    _footer(c, "Part One", None, dark=False, base=536.2)


def _page_steps(c, title):
    _header(c, "How It Works", title=title)
    _eyebrow(c, "How it works")
    _heading(c, "Numbered Steps")
    _text(c, MARGIN, 74.9, "Use this layout for any process the reader should follow in order.", "Cambria", 8.5, BODY)

    circled = ["①", "②", "③", "④"]
    names = ["First Step", "Second Step", "Third Step", "Fourth Step"]
    for i, name in enumerate(names):
        base = 93.7 + i * 28.7
        glyph_font = "Cambria-Bold"
        _text(c, MARGIN, base, circled[i], glyph_font, 10.5, BLUE)
        _text(c, 31.7, base, name, "Georgia-Bold", 10.5, DARK)
        _text(c, 30.4, base + 10.9, "One or two sentences describing this step.", "Cambria", 8, BODY)

    top = 198.5
    _rect(c, MARGIN, top, CONTENT_W, 33.0, fill=DARK)
    _text(c, MARGIN + 7, 213.1, "Info Box Title", "Georgia-Bold", 11, white)
    rows = [("OPTION ONE", 231.5, 242.1), ("OPTION TWO", 258.9, 270.0), ("OPTION THREE", 286.8, 297.9)]
    for label, row_top, base in rows:
        _rect(c, MARGIN, row_top, CONTENT_W, 27.8, fill=ROW_BG)
        _hline(c, MARGIN, RIGHT, row_top + 27.8, ROW_LINE, 0.5)
        lab = label + " — "
        _text(c, MARGIN + 7, base, lab, "Consolas-Bold", 7, DARK)
        _text(c, MARGIN + 7 + _width(lab, "Consolas-Bold", 7), base, "Short description of this option.",
              "Cambria", 7.5, BODY)
    _footer(c, "How It Works", True)


def _page_callouts(c, title):
    _header(c, "Tiered Callouts", title=title)
    _eyebrow(c, "Reading the result")
    _heading(c, "Tiered Callouts")
    _text(c, MARGIN, 74.9, "Use color to separate outcomes into best / okay / worst tiers.", "Cambria", 8.5, BODY)

    tiers = [
        ("BEST", "Best Outcome", "Describe the most favorable result here.", GREEN, GREEN_BG, 83.7),
        ("OKAY", "Middle Outcome", "Describe the in-between result here.", AMBER, AMBER_BG, 123.9),
        ("WORST", "Worst Outcome", "Describe the least favorable result here.", RED, RED_BG, 164.1),
    ]
    tag_w = 50.4
    for tag, title, desc, color, bg, top in tiers:
        h = 28.8
        _rect(c, MARGIN, top, tag_w, h, fill=bg)
        _rect(c, MARGIN + tag_w, top, CONTENT_W - tag_w, h, fill=white)
        _rect(c, MARGIN, top, CONTENT_W, h, stroke=color, line=0.5)
        _text(c, MARGIN + tag_w / 2, top + 11.1, tag, "Consolas-Bold", 6.5, color, "c")
        base = top + 12.2
        _text(c, MARGIN + tag_w + 7, base, title + "  ", "Georgia-Bold", 9.5, DARK)
        _text(c, MARGIN + tag_w + 7 + _width(title + "  ", "Georgia-Bold", 9.5), base, desc, "Cambria", 7.5, BODY)

    _text(c, MARGIN, 223.6, "FOUR ATTRIBUTES", "Consolas-Bold", 7, BLUE)
    half = CONTENT_W / 2
    attrs = [("Attribute One", 0, 240.1), ("Attribute Two", 1, 240.1),
             ("Attribute Three", 0, 288.8), ("Attribute Four", 1, 288.8)]
    for title, col, base in attrs:
        x = MARGIN + col * half
        _text(c, x + 7, base, title, "Georgia-Bold", 9.5, DARK)
        _text(c, x + 7, base + 20.9, "One-line description.", "Cambria", 7.5, BODY)
    for col in (0, 1):
        x = MARGIN + col * half + 0.75
        c.setStrokeColor(BLUE)
        c.setLineWidth(1.5)
        c.line(x, _y(228.5), x, _y(325.8))
    _footer(c, "Tiered Callouts", True)


def _page_table(c, title):
    _header(c, "Table & Badges", title=title)
    _eyebrow(c, "Reference")
    _heading(c, "Table & Badges")
    _text(c, MARGIN, 74.9, "A small reference grid, plus inline colored status badges.", "Cambria", 8.5, BODY)

    col_w = CONTENT_W / 4
    cols = ["ROW", "COL A", "COL B", "COL C"]
    _rect(c, MARGIN, 83.7, CONTENT_W, 24.1, fill=DARK)
    for j, name in enumerate(cols):
        _text(c, MARGIN + j * col_w + 4, 92.1, name, "Consolas-Bold", 6.5, white)

    data = [["Item 1", "—", "2x", "1x"], ["Item 2", "2x", "—", "1/2"], ["Item 3", "1x", "1/2", "—"]]
    tops = [107.8, 133.6, 159.9]
    heights = [25.8, 26.3, 26.3]
    for row, top, h in zip(data, tops, heights):
        for j, val in enumerate(row):
            x = MARGIN + j * col_w
            if val == "2x":
                _rect(c, x, top, col_w, h, fill=GREEN_BG)
            elif val == "1/2":
                _rect(c, x, top, col_w, h, fill=RED_BG)
            _text(c, x + 4, top + 10.1, val, "Cambria-Bold" if j == 0 else "Cambria", 8, DARK if j == 0 else BODY)
        _hline(c, MARGIN, RIGHT, top + h, CELL_LINE, 0.5)

    _text(c, MARGIN, 213.1, "BADGE PILLS", "Consolas-Bold", 7, BLUE)
    badge_w = CONTENT_W / 3
    for i, (label, color) in enumerate([("STATUS ONE", BLUE), ("STATUS TWO", RED), ("STATUS THREE", GREEN)]):
        x = MARGIN + i * badge_w
        _rect(c, x, 218.0, badge_w, 22.9, fill=color)
        _text(c, x + badge_w / 2, 225.9, label, "Consolas-Bold", 6.5, white, "c")
    _footer(c, "Table & Badges", True)


def build_adventure_pdf(opening_text, player_name, location_id, image_path, output_path):
    """Write the finished five-page adventure PDF to `output_path`."""
    _register_fonts()
    c = canvas.Canvas(output_path, pagesize=(PAGE_W, PAGE_H))
    header_title = f"{player_name}'s Adventure"  # left side of every page header
    c.setTitle(f"{player_name}'s Journey Begins")
    c.setAuthor("Morimon Story Maker")

    _page_content(c, opening_text, player_name, location_id, image_path, header_title)
    c.showPage()
    _page_divider(c, header_title)
    c.showPage()
    _page_steps(c, header_title)
    c.showPage()
    _page_callouts(c, header_title)
    c.showPage()
    _page_table(c, header_title)
    c.showPage()
    c.save()
