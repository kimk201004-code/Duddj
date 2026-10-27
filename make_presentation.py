#!/usr/bin/env python3
"""Build the FIVE GUYS review deck (5 slides, 16:9) with python-pptx.

    python make_presentation.py      ->  FiveGuys_Presentation.pptx

Presentation rule: no readable text on the slides, only photos that help the
audience follow the talk. Design: Five Guys red/white/black with a red-white
checker band. Photos: Wikimedia Commons (credits in assets/CREDITS.md and in
the file's properties, never on a slide).
"""
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu

ROOT = Path(__file__).resolve().parent
PHOTOS = ROOT / "assets" / "photos"
OUTPUT = ROOT / "FiveGuys_Presentation.pptx"

# --- Brand palette: Five Guys red, white, black
RED = RGBColor(0xED, 0x1C, 0x24)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
TILE_DARK = RGBColor(0x1E, 0x1E, 0x1E)   # panel behind the map on black slides

# --- Geometry (inches unless noted)
EMU_IN = 914400
SLIDE_W_EMU, SLIDE_H_EMU = 12192000, 6858000      # 13.333 x 7.5 in
SLIDE_W, SLIDE_H = SLIDE_W_EMU / EMU_IN, SLIDE_H_EMU / EMU_IN
CHECK_COLS, CHECK_ROWS = 60, 2                    # red/white checker band
CHECK_EMU = SLIDE_W_EMU // CHECK_COLS
BAND_Y = SLIDE_H - CHECK_ROWS * CHECK_EMU / EMU_IN
M = 0.5                                           # outer margin
TOP = 0.45
BOTTOM = BAND_Y - 0.4                             # photos end above the band
GAP = 0.2                                         # gutter between photos
RADIUS = 0.14                                     # photo corner radius

PHOTO_CREDITS = (
    "Photos via Wikimedia Commons. "
    "Slide 1: Andy Li (CC0). "
    "Slide 2: Tjeerd Wiersma (CC BY 2.0); Danielstats (CC0). "
    "Slide 3: Sarah Stierch (CC0); Bobak Ha'Eri (CC BY 3.0). "
    "Slide 4: Bobak Ha'Eri (CC BY 3.0); US Embassy Canada (CC BY 2.0); "
    "ccbarr (CC BY-SA 2.0); Roger469 (CC BY-SA 3.0). "
    "Slide 5: Simon Doggett (CC BY 2.0); Kalsoum17 (CC BY-SA 4.0); "
    "Five Guys country map, Canuckguy and others (CC0)."
)


def E(inches):
    return Emu(int(round(inches * EMU_IN)))


# --------------------------------------------------------------- shape helpers
def _flat(shape, color):
    """Solid fill, no outline, no theme style (so no inherited shadow/bevel)."""
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)


def _round_corners(shape, w, h):
    geom = shape._element.spPr.find(qn("a:prstGeom"))
    av = geom.find(qn("a:avLst"))
    if av is None:
        av = etree.SubElement(geom, qn("a:avLst"))
    for gd in list(av):
        av.remove(gd)
    gd = etree.SubElement(av, qn("a:gd"))
    gd.set("name", "adj")
    gd.set("fmla", f"val {int(min(0.5, RADIUS / min(w, h)) * 100000)}")


def add_panel(slide, x, y, w, h, color, name):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(x), E(y), E(w), E(h))
    _flat(s, color)
    _round_corners(s, w, h)
    s.name = name
    return s


def add_photo(slide, file, x, y, w, h, *, focus=(0.5, 0.5), alt="", fit=False):
    """Photo with rounded corners that covers the box (cropped around `focus`).

    fit=True keeps the whole image instead (letterboxed inside the box)."""
    path = PHOTOS / file
    with Image.open(path) as im:
        iw, ih = im.size
    box, img = w / h, iw / ih
    if fit:
        if img > box:
            nh = w / img
            y, h = y + (h - nh) / 2, nh
        else:
            nw = h * img
            x, w = x + (w - nw) / 2, nw
    pic = slide.shapes.add_picture(str(path), E(x), E(y), E(w), E(h))
    if not fit:
        if img > box:
            keep = box / img
            left = min(max(focus[0] - keep / 2, 0), 1 - keep)
            pic.crop_left, pic.crop_right = left, 1 - keep - left
        else:
            keep = img / box
            top = min(max(focus[1] - keep / 2, 0), 1 - keep)
            pic.crop_top, pic.crop_bottom = top, 1 - keep - top
        pic.auto_shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
        _round_corners(pic, w, h)
    pic.name = f"Photo {Path(file).stem}"
    pic._element.nvPicPr.cNvPr.set("descr", alt)
    return pic


def checker_band(slide, on=RED, off=WHITE):
    """Five Guys red/white checker pattern along the bottom edge (one group)."""
    grp = slide.shapes.add_group_shape()
    y0 = SLIDE_H_EMU - CHECK_ROWS * CHECK_EMU
    _flat(grp.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, y0, SLIDE_W_EMU, CHECK_ROWS * CHECK_EMU), off)
    for r in range(CHECK_ROWS):
        for c in range(CHECK_COLS):
            if (r + c) % 2 == 0:
                _flat(grp.shapes.add_shape(MSO_SHAPE.RECTANGLE, c * CHECK_EMU, y0 + r * CHECK_EMU,
                                           CHECK_EMU, CHECK_EMU), on)
    grp.name = "Checker band"


def new_slide(prs, color):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # "Blank": no placeholders, no text
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = color
    return s


# ---------------------------------------------------------------------- slides
W = SLIDE_W - 2 * M        # usable width
H = BOTTOM - TOP           # usable height


def slide_title(prs):
    """Intro: the restaurant itself (its red neon sign carries the name)."""
    s = new_slide(prs, RED)
    add_photo(s, "store_brighton_night.jpg", M, TOP, W, H, focus=(0.5, 0.5),
              alt="Five Guys restaurant at night with its red neon sign")
    checker_band(s)


def slide_marketing_1(prs):
    """$0 advertising & open kitchen: cooks in view, fresh potatoes stacked in the store."""
    s = new_slide(prs, WHITE)
    lw = 7.6
    add_photo(s, "open_kitchen_covent_garden.jpg", M, TOP, lw, H, focus=(0.42, 0.5),
              alt="Five Guys cooks making burgers in an open kitchen")
    add_photo(s, "fresh_potatoes_peanuts.jpg", M + lw + GAP, TOP, W - lw - GAP, H, focus=(0.5, 0.5),
              alt="Sacks of fresh potatoes and boxes of peanuts inside a Five Guys")
    checker_band(s)


def slide_marketing_2(prs):
    """Mass customization (toppings) & brand experience (free peanuts)."""
    s = new_slide(prs, WHITE)
    half = (W - GAP) / 2
    add_photo(s, "toppings_burger.jpg", M, TOP, half, H, focus=(0.5, 0.5),
              alt="Five Guys burger stacked with toppings")
    add_photo(s, "free_peanuts.jpg", M + half + GAP, TOP, half, H, focus=(0.45, 0.55),
              alt="A box of free in-shell peanuts at a Five Guys")
    checker_band(s)


def slide_burger(prs):
    """Bacon cheeseburger + the three techniques: sear, caramelize, saute."""
    s = new_slide(prs, BLACK)
    hw = 7.9
    add_photo(s, "bacon_cheeseburger.jpg", M, TOP, hw, H, focus=(0.5, 0.5),
              alt="Five Guys bacon cheeseburger")
    cx, cw = M + hw + GAP, W - hw - GAP
    ch = (H - 2 * GAP) / 3
    steps = [
        ("sear_patties_grill.jpg", (0.5, 0.85), "Burger patties searing on a flat grill"),
        ("caramelized_onions.jpg", (0.5, 0.5), "Onions caramelizing in a pan"),
        ("sauteed_mushrooms.jpg", (0.5, 0.5), "Sauteed mushrooms"),
    ]
    for i, (file, focus, alt) in enumerate(steps):
        add_photo(s, file, cx, TOP + i * (ch + GAP), cw, ch, focus=focus, alt=alt)
    checker_band(s)


def slide_fries(prs):
    """Cajun fries: frying in oil, the finished fries, and the global map (conclusion)."""
    s = new_slide(prs, BLACK)
    lw = 5.3
    lh = (H - GAP) / 2
    add_photo(s, "fries_frying.jpg", M, TOP, lw, lh, focus=(0.5, 0.5),
              alt="Hand-cut fries frying in hot oil")
    add_panel(s, M, TOP + lh + GAP, lw, lh, TILE_DARK, name="Map panel")
    add_photo(s, "five_guys_world_map.png", M + 0.2, TOP + lh + GAP + 0.2, lw - 0.4, lh - 0.4,
              fit=True, alt="World map with the countries that have Five Guys restaurants in red")
    add_photo(s, "cajun_fries.jpg", M + lw + GAP, TOP, W - lw - GAP, H, focus=(0.42, 0.5),
              alt="Five Guys Cajun fries in a white cup")
    checker_band(s)


# ---------------------------------------------------------------------- theme
def apply_brand_theme(prs):
    """Brand colors + Arial in the theme, 16:9-scaled master/layout placeholders."""
    theme = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme.blob)
    a = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = root.find(".//a:clrScheme", a)
    scheme.set("name", "Five Guys")
    colors = {"dk1": "000000", "lt1": "FFFFFF", "dk2": "1A1A1A", "lt2": "F2F2F2",
              "accent1": "ED1C24", "accent2": "000000", "accent3": "8C8C8C",
              "accent4": "B3141A", "accent5": "F47C80", "accent6": "595959",
              "hlink": "ED1C24", "folHlink": "B3141A"}
    for tag_name, val in colors.items():
        el = scheme.find(f"a:{tag_name}", a)
        for child in list(el):
            el.remove(child)
        etree.SubElement(el, qn("a:srgbClr")).set("val", val)
    fonts = root.find(".//a:fontScheme", a)
    fonts.set("name", "Five Guys")
    for which in ("majorFont", "minorFont"):
        fonts.find(f"a:{which}/a:latin", a).set("typeface", "Arial")
    for effect in root.iterfind(".//a:effectStyle", a):  # flat look: no theme shadows/bevels
        for child in list(effect):
            effect.remove(child)
        etree.SubElement(effect, qn("a:effectLst"))
    theme.blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # The stock template is 4:3; stretch its placeholders to the 16:9 canvas.
    ratio = SLIDE_W_EMU / 9144000
    holders = list(prs.slide_master.placeholders)
    for layout in prs.slide_layouts:
        holders += list(layout.placeholders)
    for ph in holders:
        if ph._element.spPr.find(qn("a:xfrm")) is not None:
            ph.left, ph.width = Emu(int(ph.left * ratio)), Emu(int(ph.width * ratio))


def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(SLIDE_W_EMU), Emu(SLIDE_H_EMU)
    prs.part._element.find(qn("p:sldSz")).attrib.pop("type", None)  # template says 4:3
    apply_brand_theme(prs)
    props = prs.core_properties
    props.title = "FIVE GUYS Review"
    props.author = props.last_modified_by = ""
    # File > Info comments only, never on a slide (python-pptx caps the setter at 255 chars)
    props._element._get_or_add("description").text = PHOTO_CREDITS

    slide_title(prs)
    slide_marketing_1(prs)
    slide_marketing_2(prs)
    slide_burger(prs)
    slide_fries(prs)
    prs.save(OUTPUT)
    print(f"Saved {OUTPUT}")


if __name__ == "__main__":
    build()
