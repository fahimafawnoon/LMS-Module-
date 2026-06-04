"""
QPON Merchant Pitch Deck — PPTX builder
Creates two versions:
  v1: dark navy theme (current style)
  v2: light theme with QPON brand colours
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as nsmap
from lxml import etree
import copy

# ── Widescreen 16:9 ──────────────────────────────────────────────────────────
W = Inches(13.333)
H = Inches(7.5)

# ── Brand palette ────────────────────────────────────────────────────────────
NAVY       = RGBColor(0x0D, 0x1B, 0x3E)
NAVY2      = RGBColor(0x15, 0x22, 0x47)
NAVY3      = RGBColor(0x1A, 0x2B, 0x55)
ORANGE     = RGBColor(0xF4, 0x51, 0x1E)
ORANGE2    = RGBColor(0xFF, 0x6B, 0x35)
CORAL      = RGBColor(0xFF, 0x8C, 0x5A)
TEAL       = RGBColor(0x00, 0xC9, 0xB1)
TEAL2      = RGBColor(0x00, 0xA8, 0x96)
PURPLE     = RGBColor(0x7B, 0x2F, 0xBE)
PURPLE2    = RGBColor(0xA8, 0x55, 0xF7)
GOLD       = RGBColor(0xFF, 0xB8, 0x30)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY = RGBColor(0xF4, 0xF6, 0xFB)
MID_GREY   = RGBColor(0xA8, 0xB4, 0xCF)
DARK_BG    = RGBColor(0x08, 0x0E, 0x22)

# Light-theme equivalents
L_BG       = RGBColor(0xFA, 0xFB, 0xFF)
L_BG2      = RGBColor(0xF0, 0xF3, 0xFF)
L_BG3      = RGBColor(0xE8, 0xEC, 0xF8)
L_TEXT     = RGBColor(0x0D, 0x1B, 0x3E)
L_TEXT2    = RGBColor(0x3A, 0x4A, 0x6B)
L_CARD     = RGBColor(0xFF, 0xFF, 0xFF)
L_BORDER   = RGBColor(0xD8, 0xDE, 0xF0)


# ── Low-level XML helpers ────────────────────────────────────────────────────
def rgb_hex(r):
    return "{:02X}{:02X}{:02X}".format(r[0], r[1], r[2])

def set_cell_bg(cell, colour):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    solidFill = etree.SubElement(tcPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
    srgb = etree.SubElement(solidFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
    srgb.set('val', rgb_hex(colour))


# ── Shape helpers ────────────────────────────────────────────────────────────
def add_rect(slide, x, y, w, h, fill_colour, radius_emu=None, line_colour=None, line_width=None):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if line_colour:
        shape.line.color.rgb = line_colour
        if line_width:
            shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    if radius_emu:
        sp = shape._element
        spPr = sp.find('{http://schemas.openxmlformats.org/drawingml/2006/main}spPr') or \
               sp.find('{http://schemas.openxmlformats.org/presentationml/2006/main}spPr')
        if spPr is not None:
            prstGeom = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom')
            if prstGeom is not None:
                prstGeom.set('prst', 'roundRect')
                avLst = prstGeom.find('{http://schemas.openxmlformats.org/drawingml/2006/main}avLst')
                if avLst is None:
                    avLst = etree.SubElement(prstGeom, '{http://schemas.openxmlformats.org/drawingml/2006/main}avLst')
                gd = etree.SubElement(avLst, '{http://schemas.openxmlformats.org/drawingml/2006/main}gd')
                gd.set('name', 'adj')
                gd.set('fmla', 'val 30000')
    return shape

def add_textbox(slide, text, x, y, w, h,
                font_size=14, bold=False, colour=WHITE,
                align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = colour
    run.font.italic = italic
    run.font.name = 'Segoe UI'
    return txBox

def add_multiline(slide, lines, x, y, w, h,
                  font_size=13, bold=False, colour=WHITE,
                  align=PP_ALIGN.LEFT, line_spacing=None):
    """lines: list of (text, bold, size, colour) tuples or plain strings"""
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for item in lines:
        if isinstance(item, str):
            text, b, s, c = item, bold, font_size, colour
        else:
            text, b, s, c = item
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = Pt(line_spacing)
        run = p.add_run()
        run.text = text
        run.font.size = Pt(s)
        run.font.bold = b
        run.font.color.rgb = c
        run.font.name = 'Segoe UI'
    return txBox

def add_hyperlink_textbox(slide, text, url, x, y, w, h,
                          font_size=16, bold=True, colour=WHITE, align=PP_ALIGN.CENTER):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = colour
    run.font.name = 'Segoe UI'
    # Add hyperlink
    rPr = run._r.get_or_add_rPr()
    hlinkClick = etree.SubElement(
        rPr,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkClick'
    )
    # Add relationship
    slide_part = slide.part
    rId = slide_part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    hlinkClick.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', rId)
    return txBox

def slide_bg(slide, colour):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = colour


# ── Accent bar helper ─────────────────────────────────────────────────────────
def accent_bar(slide, x, y, h, colour):
    add_rect(slide, x, y, 0.05, h, colour)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE BUILDERS  (shared logic, dark=True for V1, dark=False for V2)
# ══════════════════════════════════════════════════════════════════════════════

def bg_col(dark, a, b):
    return a if dark else b

def txt_col(dark):
    return WHITE if dark else L_TEXT

def sub_col(dark):
    return MID_GREY if dark else L_TEXT2

def card_col(dark):
    return NAVY3 if dark else L_CARD

def card_border(dark):
    return RGBColor(0x2A, 0x3B, 0x65) if dark else L_BORDER


# ── Logo text ─────────────────────────────────────────────────────────────────
def add_logo(slide, x=0.3, y=0.12, dark=True):
    # Teal square stand-in for tag icon
    add_rect(slide, x, y+0.02, 0.22, 0.22, TEAL)
    add_textbox(slide, "Q", x+0.04, y+0.01, 0.15, 0.25, font_size=11, bold=True, colour=WHITE)
    add_textbox(slide, "QPON", x+0.27, y, 0.8, 0.28, font_size=18, bold=True,
                colour=txt_col(dark))

def add_slide_num(slide, num, total=13, dark=True):
    add_textbox(slide, f"{num:02d} / {total}", 12.7, 0.13, 0.5, 0.25,
                font_size=10, colour=sub_col(dark), align=PP_ALIGN.RIGHT)


# ── Divider line ──────────────────────────────────────────────────────────────
def add_divider(slide, x, y, w, colour=ORANGE):
    add_rect(slide, x, y, w, 0.04, colour)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 01 — HERO
# ══════════════════════════════════════════════════════════════════════════════
def build_s01(slide, dark=True):
    slide_bg(slide, bg_col(dark, NAVY, L_BG2))
    # Top accent strip
    add_rect(slide, 0, 0, 13.333, 0.06, ORANGE)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 1, dark=dark)

    # Hero heading
    add_textbox(slide, "More Customers.", 0.6, 0.85, 8, 0.65, font_size=46, bold=True, colour=txt_col(dark))
    add_textbox(slide, "Zero Commission.", 0.6, 1.45, 8, 0.65, font_size=46, bold=True, colour=ORANGE)
    add_textbox(slide, "Total Control.", 0.6, 2.05, 8, 0.65, font_size=42, bold=True, colour=TEAL)

    add_divider(slide, 0.6, 2.78, 1.2, ORANGE)

    add_textbox(slide,
        "QPON is a commission-free coupon & deals platform built for local Sri Lankan businesses.\n"
        "You set the deals, you own the relationship — we put you in front of ready-to-spend customers.",
        0.6, 2.9, 7.8, 1.0, font_size=14, colour=sub_col(dark))

    # Stat pills
    for i, (stat, label, col) in enumerate([
        ("0%",       "Commission — Ever",      ORANGE),
        ("3 Months", "Free to Start",          TEAL),
        ("50 Spots", "Founding Partners Only", GOLD),
    ]):
        bx = 0.6 + i * 4.1
        add_rect(slide, bx, 4.1, 3.8, 0.95, card_col(dark), line_colour=col, line_width=1.5)
        add_textbox(slide, stat, bx+0.18, 4.18, 3.4, 0.5, font_size=26, bold=True, colour=col)
        add_textbox(slide, label, bx+0.18, 4.66, 3.4, 0.3, font_size=11, colour=sub_col(dark))

    # Right panel — mock dashboard
    add_rect(slide, 9.5, 0.8, 3.5, 5.9, card_col(dark), radius_emu=200000, line_colour=card_border(dark), line_width=1)
    add_rect(slide, 9.65, 1.0, 3.2, 0.7, ORANGE)
    add_textbox(slide, "📊 Live Dashboard — 14 redemptions today", 9.7, 1.05, 3.1, 0.55, font_size=10, bold=True, colour=WHITE)
    for j, (em, name, pct, col) in enumerate([
        ("🍽️","Lunch Special — 25% OFF · 38 redeemed",  "▮▮▮▮▮▮▮▮", TEAL),
        ("🌅","Happy Hour — 30% OFF · 22 redeemed",      "▮▮▮▮▮▮",   ORANGE2),
        ("☕","Morning Coffee — 20% OFF · 17 redeemed",  "▮▮▮▮▮",    GOLD),
    ]):
        y0 = 1.85 + j * 1.35
        add_rect(slide, 9.65, y0, 3.2, 1.2, bg_col(dark, NAVY2, L_BG3), line_colour=card_border(dark), line_width=0.5)
        add_textbox(slide, em, 9.75, y0+0.08, 0.4, 0.4, font_size=18)
        add_textbox(slide, name, 10.2, y0+0.1, 2.55, 0.35, font_size=9, colour=txt_col(dark))
        add_rect(slide, 9.75, y0+0.55, 3.0, 0.08, bg_col(dark, NAVY, L_BG2))
        add_rect(slide, 9.75, y0+0.55, 1.5 + j*0.4, 0.08, col)
        add_textbox(slide, pct, 9.75, y0+0.72, 3.0, 0.3, font_size=8, colour=col)
    add_rect(slide, 9.65, 5.9, 3.2, 0.5, TEAL2)
    add_textbox(slide, "↑ 43% more visits vs last week", 9.72, 5.97, 3.05, 0.34, font_size=10, bold=True, colour=WHITE, align=PP_ALIGN.CENTER)

    # Bottom bar
    add_rect(slide, 0, 6.9, 13.333, 0.6, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4,
                font_size=10, colour=sub_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 02 — THE MERCHANT PROBLEM
# ══════════════════════════════════════════════════════════════════════════════
def build_s02(slide, dark=True):
    slide_bg(slide, bg_col(dark, DARK_BG, L_BG))
    add_rect(slide, 0, 0, 13.333, 0.06, ORANGE)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 2, dark=dark)

    add_textbox(slide, "THE REALITY FOR LOCAL BUSINESSES", 0.6, 0.45, 10, 0.28, font_size=10, bold=True, colour=ORANGE)
    add_textbox(slide, "You're Working Hard.", 0.6, 0.78, 12, 0.6, font_size=38, bold=True, colour=txt_col(dark))
    add_textbox(slide, "The Middlemen Are Getting Paid.", 0.6, 1.35, 12, 0.6, font_size=38, bold=True, colour=ORANGE)

    problems = [
        ("💸", "Commissions Eat Your Margin",
         "OTAs and delivery apps charge 20–30% on every transaction. For a café doing\nLKR 500K/month, that's LKR 100–150K gone before you pay staff.",
         ORANGE, "Up to 30% lost per sale"),
        ("👥", "No Direct Customer Relationship",
         "Third-party platforms own the customer data. You fulfil the order — they keep\nthe contact, the reviews, and the repeat business opportunity.",
         ORANGE2, "Zero ownership of your audience"),
        ("📉", "Empty Tables in Off-Peak Hours",
         "Weekday lunches, mid-afternoon, low-season months — capacity sits idle.\nNo smart tool to attract customers precisely when you need them.",
         CORAL, "~70% of off-peak capacity wasted"),
    ]
    for i, (em, title, body, col, sub) in enumerate(problems):
        bx = 0.6 + i * 4.25
        add_rect(slide, bx, 2.2, 4.0, 3.4, card_col(dark), line_colour=col, line_width=1.5)
        add_textbox(slide, em, bx+0.2, 2.35, 0.55, 0.55, font_size=26)
        add_textbox(slide, title, bx+0.2, 2.95, 3.6, 0.55, font_size=14, bold=True, colour=txt_col(dark))
        add_textbox(slide, body, bx+0.2, 3.55, 3.6, 1.05, font_size=11, colour=sub_col(dark))
        add_rect(slide, bx+0.2, 4.9, 3.6, 0.04, col)
        add_textbox(slide, sub, bx+0.2, 4.98, 3.6, 0.3, font_size=10, bold=True, colour=col)

    # Bottom row
    for i, (em, title, body) in enumerate([
        ("🚫", "No visibility without paid ads", "Meta/Google ads cost LKR 50–200K/mo with unpredictable ROI for most SMEs."),
        ("📊", "No real analytics", "Merchants have no clear view of which promotions work, when, and for whom."),
        ("🔄", "No repeat visit engine", "One-time walk-ins are never converted into loyal regulars — the highest-value segment."),
    ]):
        bx = 0.6 + i * 4.25
        add_rect(slide, bx, 5.55, 4.0, 1.2, card_col(dark), line_colour=card_border(dark), line_width=1)
        add_textbox(slide, em + "  " + title, bx+0.18, 5.68, 3.6, 0.38, font_size=12, bold=True, colour=txt_col(dark))
        add_textbox(slide, body, bx+0.18, 6.1, 3.6, 0.5, font_size=10, colour=sub_col(dark))

    add_rect(slide, 0, 6.9, 13.333, 0.6, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4, font_size=10, colour=sub_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 03 — QPON SOLUTION
# ══════════════════════════════════════════════════════════════════════════════
def build_s03(slide, dark=True):
    slide_bg(slide, bg_col(dark, NAVY, L_BG2))
    add_rect(slide, 0, 0, 13.333, 0.06, TEAL)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 3, dark=dark)

    add_textbox(slide, "THE QPON SOLUTION", 0.6, 0.45, 8, 0.28, font_size=10, bold=True, colour=TEAL)
    add_textbox(slide, "Your Business.", 0.6, 0.78, 7, 0.58, font_size=38, bold=True, colour=txt_col(dark))
    add_textbox(slide, "Your Rules.", 0.6, 1.33, 7, 0.58, font_size=38, bold=True, colour=TEAL)
    add_divider(slide, 0.6, 1.98, 1.0, TEAL)
    add_textbox(slide,
        "QPON is a commission-free coupon platform built for local Sri Lankan businesses.\n"
        "Run promotions on your terms — we put you in front of the customers you want.",
        0.6, 2.12, 6.4, 0.85, font_size=13, colour=sub_col(dark))

    pillars = [
        (ORANGE,  "No commission. Ever.",       "Pay only when a coupon is redeemed. You keep what you earn."),
        (TEAL,    "Full campaign control.",      "Set your own discount %, caps, duration. Pause or edit instantly."),
        (PURPLE2, "Real customer data.",         "See exactly who redeems, when, and how often — your insights."),
        (GOLD,    "Corporate channel access.",   "Get discovered by employees through company benefit programs."),
    ]
    for i, (col, title, body) in enumerate(pillars):
        y0 = 3.2 + i * 0.88
        add_rect(slide, 0.6, y0, 0.06, 0.6, col)
        add_textbox(slide, title, 0.82, y0, 5.8, 0.32, font_size=13, bold=True, colour=txt_col(dark))
        add_textbox(slide, body, 0.82, y0+0.32, 5.8, 0.45, font_size=11, colour=sub_col(dark))

    # Right: 4-step how it works
    add_textbox(slide, "GETTING STARTED — 4 STEPS", 7.3, 0.82, 5.6, 0.28, font_size=10, bold=True, colour=sub_col(dark))
    steps = [
        ("1", "Create your merchant account",    "Sign up at qpon.lk. Our team onboards you — no tech skills needed."),
        ("2", "Build your first offer",          "Set discount %, redemption cap, dates. Publish in minutes."),
        ("3", "Customers discover & visit",      "Geo-targeted deal surfaces to the right customer near you."),
        ("4", "Track real-time analytics",       "See live data. Pause or edit campaigns anytime."),
    ]
    for i, (num, title, body) in enumerate(steps):
        y0 = 1.25 + i * 1.45
        bg = TEAL2 if i == 3 else card_col(dark)
        bdr = TEAL if i == 3 else card_border(dark)
        add_rect(slide, 7.3, y0, 5.8, 1.28, bg, line_colour=bdr, line_width=1.2)
        add_rect(slide, 7.42, y0+0.35, 0.42, 0.42, ORANGE)
        add_textbox(slide, num, 7.42, y0+0.35, 0.42, 0.42, font_size=16, bold=True, colour=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, title, 7.95, y0+0.12, 5.0, 0.36, font_size=13, bold=True, colour=WHITE if i==3 else txt_col(dark))
        add_textbox(slide, body,  7.95, y0+0.52, 5.0, 0.55, font_size=11, colour=WHITE if i==3 else sub_col(dark))

    add_textbox(slide, "⏱  Average setup time: under 15 minutes", 7.3, 7.02, 5.8, 0.3, font_size=10, colour=TEAL)
    add_rect(slide, 0, 6.9, 13.333, 0.6, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4, font_size=10, colour=sub_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 04 — SMARTER DISCOUNTS
# ══════════════════════════════════════════════════════════════════════════════
def build_s04(slide, dark=True):
    slide_bg(slide, bg_col(dark, DARK_BG, L_BG))
    add_rect(slide, 0, 0, 13.333, 0.06, TEAL)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 4, dark=dark)

    add_textbox(slide, "FEATURE DEEP-DIVE", 0.6, 0.45, 8, 0.28, font_size=10, bold=True, colour=TEAL)
    add_textbox(slide, "Smarter Discounts.", 0.6, 0.78, 8, 0.58, font_size=36, bold=True, colour=txt_col(dark))
    add_textbox(slide, "Bigger Returns.", 0.6, 1.33, 8, 0.55, font_size=36, bold=True, colour=ORANGE)

    features = [
        ("🎯","Target by time slot",   "Run a 20% lunch special only 12–2 pm on weekdays. Fill tables that would sit empty."),
        ("🔢","Set redemption caps",   "Limit to 30 coupons/day so you never over-commit capacity or erode margin."),
        ("⏸️","Pause or edit instantly","Fully booked? Special event? Pause the deal in one tap — no admin, no delays."),
        ("📅","Seasonal campaigns",    "Pre-schedule promotions for low season, holidays, or local events."),
    ]
    for i, (em, title, body) in enumerate(features):
        y0 = 2.05 + i * 1.12
        add_rect(slide, 0.6, y0, 6.0, 0.98, card_col(dark), line_colour=TEAL2, line_width=1)
        add_textbox(slide, em, 0.78, y0+0.18, 0.5, 0.5, font_size=20)
        add_textbox(slide, title, 1.38, y0+0.08, 5.0, 0.35, font_size=13, bold=True, colour=txt_col(dark))
        add_textbox(slide, body, 1.38, y0+0.46, 5.0, 0.42, font_size=11, colour=sub_col(dark))

    # ROI comparison
    add_textbox(slide, "EXAMPLE: A MIRISSA RESTAURANT", 7.2, 0.82, 5.8, 0.28, font_size=10, bold=True, colour=sub_col(dark))

    # Before
    add_rect(slide, 7.2, 1.18, 5.8, 2.52, card_col(dark), line_colour=card_border(dark), line_width=1)
    add_textbox(slide, "BEFORE QPON", 7.38, 1.28, 5.4, 0.28, font_size=10, bold=True, colour=ORANGE2)
    rows_b = [
        ("Monthly revenue",         "LKR 500,000",  txt_col(dark)),
        ("OTA commission (25%)",    "− LKR 125,000", ORANGE),
        ("Ad spend (Meta/Google)",  "− LKR 60,000",  ORANGE),
    ]
    for j, (lbl, val, col) in enumerate(rows_b):
        y0 = 1.65 + j*0.55
        add_textbox(slide, lbl, 7.38, y0, 3.5, 0.38, font_size=12, colour=sub_col(dark))
        add_textbox(slide, val, 9.9, y0, 2.9, 0.38, font_size=12, bold=True, colour=col, align=PP_ALIGN.RIGHT)
        add_rect(slide, 7.38, y0+0.4, 5.4, 0.01, bg_col(dark, NAVY2, L_BG3))
    add_rect(slide, 7.38, 3.33, 5.4, 0.25, bg_col(dark, NAVY2, L_BG3))
    add_textbox(slide, "Net retained", 7.45, 3.34, 3.0, 0.25, font_size=12, bold=True, colour=txt_col(dark))
    add_textbox(slide, "LKR 315,000", 9.9, 3.34, 2.6, 0.25, font_size=14, bold=True, colour=ORANGE, align=PP_ALIGN.RIGHT)

    # After
    add_rect(slide, 7.2, 3.82, 5.8, 2.75, card_col(dark), line_colour=TEAL, line_width=1.5)
    add_textbox(slide, "WITH QPON", 7.38, 3.92, 5.4, 0.28, font_size=10, bold=True, colour=TEAL)
    rows_a = [
        ("Monthly revenue (+15% footfall)", "LKR 575,000",  txt_col(dark)),
        ("QPON pay-per-redemption fee",      "− LKR 18,000",  TEAL),
        ("Ad spend (reduced)",               "− LKR 20,000",  TEAL),
    ]
    for j, (lbl, val, col) in enumerate(rows_a):
        y0 = 4.38 + j*0.55
        add_textbox(slide, lbl, 7.38, y0, 3.8, 0.38, font_size=12, colour=sub_col(dark))
        add_textbox(slide, val, 9.9, y0, 2.9, 0.38, font_size=12, bold=True, colour=col, align=PP_ALIGN.RIGHT)
        add_rect(slide, 7.38, y0+0.4, 5.4, 0.01, bg_col(dark, NAVY2, L_BG3))
    add_rect(slide, 7.38, 6.08, 5.4, 0.28, TEAL2)
    add_textbox(slide, "Net retained", 7.45, 6.1, 3.0, 0.28, font_size=12, bold=True, colour=WHITE)
    add_textbox(slide, "LKR 537,000", 9.9, 6.1, 2.6, 0.28, font_size=14, bold=True, colour=WHITE, align=PP_ALIGN.RIGHT)

    add_rect(slide, 7.2, 6.55, 5.8, 0.35, TEAL)
    add_textbox(slide, "📈  +LKR 222,000 more retained per month", 7.38, 6.6, 5.4, 0.28, font_size=11, bold=True, colour=WHITE)

    add_rect(slide, 0, 6.9, 13.333, 0.6, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4, font_size=10, colour=sub_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 05 — APP FEATURES
# ══════════════════════════════════════════════════════════════════════════════
def build_s05(slide, dark=True):
    slide_bg(slide, bg_col(dark, NAVY, L_BG2))
    add_rect(slide, 0, 0, 13.333, 0.06, ORANGE)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 5, dark=dark)

    add_textbox(slide, "PLATFORM FEATURES", 0.6, 0.45, 8, 0.28, font_size=10, bold=True, colour=ORANGE)
    add_textbox(slide, "Everything You Need.", 0.6, 0.78, 9, 0.58, font_size=36, bold=True, colour=txt_col(dark))
    add_textbox(slide, "Nothing You Don't.", 0.6, 1.33, 9, 0.55, font_size=36, bold=True, colour=ORANGE)

    feats = [
        ("📊","Real-Time Analytics",     TEAL,   "Live view of issued, collected, and redeemed coupons. See peak hours, best-performing deals, and return rates."),
        ("📱","Web + App Dashboard",     WHITE,  "Manage from laptop or phone. Full feature parity across all devices — no separate app needed."),
        ("🎛️","Full Campaign Control",   WHITE,  "Create, pause, edit, or delete campaigns anytime. Set spending caps, time windows, per-user limits."),
        ("📍","Geo-Targeted Discovery",  ORANGE2,"Your deals appear to customers physically near your location — the highest-intent audience possible."),
        ("🏢","Corporate Channel",       PURPLE2,"Your business is visible inside corporate employee benefit packages — driving guaranteed regular visitors."),
        ("🤝","Dedicated Onboarding",    WHITE,  "Personal setup walkthrough, demo videos, FAQ portal. First 50 merchants get a physical Welcome Kit."),
    ]
    cols = [0.6, 4.7, 8.8]
    for i, (em, title, tcol, body) in enumerate(feats):
        col_i = i % 3
        row_i = i // 3
        bx = cols[col_i]
        by = 2.25 + row_i * 2.32
        bg = bg_col(dark, NAVY3 if tcol==WHITE else (RGBColor(0x00,0x30,0x2A) if tcol==TEAL else
             (RGBColor(0x28,0x08,0x3A) if tcol==PURPLE2 else RGBColor(0x38,0x18,0x08))),
             RGBColor(0xE8,0xF8,0xF5) if tcol==TEAL else
             (RGBColor(0xF0,0xE8,0xFA) if tcol==PURPLE2 else
              (RGBColor(0xFF,0xF0,0xE8) if tcol==ORANGE2 else L_CARD)))
        add_rect(slide, bx, by, 4.0, 2.05, bg, line_colour=tcol if tcol!=WHITE else card_border(dark), line_width=1.2)
        add_textbox(slide, em, bx+0.18, by+0.18, 0.5, 0.5, font_size=24)
        add_textbox(slide, title, bx+0.18, by+0.72, 3.6, 0.38, font_size=13, bold=True, colour=txt_col(dark))
        add_textbox(slide, body,  bx+0.18, by+1.12, 3.6, 0.75, font_size=10.5, colour=sub_col(dark))

    add_rect(slide, 0, 6.9, 13.333, 0.6, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4, font_size=10, colour=sub_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 06 — CORPORATE BENEFITS CHANNEL
# ══════════════════════════════════════════════════════════════════════════════
def build_s06(slide, dark=True):
    slide_bg(slide, bg_col(dark, DARK_BG, L_BG))
    add_rect(slide, 0, 0, 13.333, 0.06, PURPLE)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 6, dark=dark)

    add_textbox(slide, "BONUS REVENUE STREAM", 0.6, 0.45, 8, 0.28, font_size=10, bold=True, colour=PURPLE2)
    add_textbox(slide, "A New Customer Channel:", 0.6, 0.78, 10, 0.58, font_size=34, bold=True, colour=txt_col(dark))
    add_textbox(slide, "Corporate Employees.", 0.6, 1.33, 10, 0.55, font_size=34, bold=True, colour=PURPLE2)

    add_textbox(slide,
        "QPON sells branded coupon books to companies as employee perks. Your business gets listed —\n"
        "giving you a stream of repeat visitors from nearby offices every single week.",
        0.6, 2.0, 6.2, 0.85, font_size=13, colour=sub_col(dark))

    steps = [
        ("🏢","Company Signs Up",           "HR purchases annual QPON employee benefit package for all staff."),
        ("👷","Employees Get Access",        "Staff open QPON and see deals near their office — your business is front and centre."),
        ("🔁","They Visit You Regularly",    "Unlike one-off tourists, employees are nearby every weekday. One company = dozens of weekly regulars."),
        ("💰","You Earn Predictably",        "Regular weekday visits. Zero effort — QPON handles the entire corporate relationship."),
    ]
    for i, (em, title, body) in enumerate(steps):
        y0 = 3.05 + i * 0.95
        add_rect(slide, 0.6, y0, 6.2, 0.82, card_col(dark), line_colour=PURPLE, line_width=1)
        add_textbox(slide, em, 0.75, y0+0.18, 0.45, 0.42, font_size=18)
        add_textbox(slide, title, 1.28, y0+0.08, 3.0, 0.32, font_size=12, bold=True, colour=txt_col(dark))
        add_textbox(slide, body,  1.28, y0+0.42, 5.35, 0.3, font_size=10.5, colour=sub_col(dark))
        if i < 3:
            add_textbox(slide, "↓", 0.85, y0+0.87, 0.25, 0.3, font_size=12, colour=PURPLE2)

    # Right: Why it matters box
    add_rect(slide, 7.3, 1.95, 5.7, 4.85, bg_col(dark, NAVY3, L_BG2), line_colour=PURPLE, line_width=1.5)
    add_textbox(slide, "WHY THIS MATTERS FOR YOUR BUSINESS", 7.5, 2.1, 5.3, 0.3, font_size=10, bold=True, colour=PURPLE2)
    add_textbox(slide,
        "Corporate employees are the most predictable customer segment — same area, same routines, "
        "high weekly spend.\n\nQPON's corporate channel turns your existing location into a reliable "
        "revenue stream beyond peak hours.\n\nOne partner company can bring you 20–100 regular "
        "weekly visitors — permanently.",
        7.5, 2.5, 5.3, 2.0, font_size=12, colour=sub_col(dark))

    bullets = ["Bulk coupon books — branded to employer", "Group subscription tiers — 10 to 10,000 employees",
               "HR usage dashboard — visibility on engagement", "Category control — curate for company culture"]
    for j, b in enumerate(bullets):
        y0 = 4.62 + j * 0.48
        add_rect(slide, 7.5, y0+0.06, 0.08, 0.08, PURPLE2)
        add_textbox(slide, b, 7.68, y0, 5.1, 0.38, font_size=11.5, colour=txt_col(dark))

    add_rect(slide, 0, 6.9, 13.333, 0.6, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4, font_size=10, colour=sub_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 07 — PRICING  (3 tiers: Free / Pro / Corporate)   ← MODIFIED
# ══════════════════════════════════════════════════════════════════════════════
def build_s07(slide, dark=True):
    slide_bg(slide, bg_col(dark, NAVY, L_BG2))
    add_rect(slide, 0, 0, 13.333, 0.06, GOLD)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 7, dark=dark)

    add_textbox(slide, "TRANSPARENT PRICING", 0.6, 0.45, 8, 0.28, font_size=10, bold=True, colour=GOLD)
    add_textbox(slide, "Simple, Fair,", 0.6, 0.78, 9, 0.58, font_size=38, bold=True, colour=txt_col(dark))
    add_textbox(slide, "Performance-Based.", 0.6, 1.33, 9, 0.55, font_size=38, bold=True, colour=ORANGE)
    add_textbox(slide, "You only pay QPON when it actually works for you. No monthly retainer, no upfront risk, no hidden fees.",
                0.6, 2.0, 12.0, 0.5, font_size=13, colour=sub_col(dark))

    tiers = [
        # (badge_text, badge_col, title, subtitle, bullet_col, bullets, featured)
        ("FREE",
         TEAL2, "Founding Partner", "First 50 merchants only", TEAL,
         ["Full platform access — no fees",
          "3 Months at zero cost",
          "Priority placement in app",
          "Welcome Kit (QR standee, cards)",
          "Dedicated onboarding support",
          "500 LKR cashback per redemption"],
         True),
        ("PRO",
         GOLD, "Monthly Subscription", "High-volume merchants", GOLD,
         ["Flat monthly fee — no per-use charge",
          "Advanced analytics & ROI reports",
          "Featured placement in app",
          "Multi-branch management",
          "Dedicated account manager",
          "Priority customer support"],
         False),
        ("CORP",
         PURPLE, "Corporate Plan", "Annual B2B packages", PURPLE2,
         ["Listed in corporate benefit books",
          "Featured to employee audiences",
          "B2B footfall & ROI reporting",
          "Company partnership branding",
          "Priority in corporate searches",
          "Tailored onboarding for HR teams"],
         False),
    ]

    # 3 columns, wider
    col_w = 3.8
    col_gap = 0.42
    start_x = (13.333 - 3*col_w - 2*col_gap) / 2

    for i, (badge, bcol, title, sub, bullet_col, bullets, featured) in enumerate(tiers):
        bx = start_x + i*(col_w + col_gap)
        by = 2.75
        bh = 4.0

        # Highlight border for featured
        if featured:
            add_rect(slide, bx-0.04, by-0.04, col_w+0.08, bh+0.08, bcol)
        bg = bg_col(dark,
                    RGBColor(0x00,0x2E,0x28) if bcol==TEAL2 else
                    (RGBColor(0x38,0x20,0x00) if bcol==GOLD else RGBColor(0x22,0x08,0x3A)),
                    RGBColor(0xE8,0xF8,0xF5) if bcol==TEAL2 else
                    (RGBColor(0xFF,0xF5,0xE0) if bcol==GOLD else RGBColor(0xF5,0xEA,0xFF)))
        add_rect(slide, bx, by, col_w, bh, bg, line_colour=bcol, line_width=1.8)

        # Badge
        add_rect(slide, bx + col_w/2 - 0.55, by - 0.22, 1.1, 0.42, bcol)
        add_textbox(slide, badge, bx + col_w/2 - 0.55, by - 0.21, 1.1, 0.4,
                    font_size=12, bold=True, colour=WHITE, align=PP_ALIGN.CENTER)

        add_textbox(slide, title, bx+0.22, by+0.28, col_w-0.44, 0.38, font_size=15, bold=True, colour=txt_col(dark))
        add_textbox(slide, sub,   bx+0.22, by+0.68, col_w-0.44, 0.28, font_size=10, colour=bcol)
        add_rect(slide, bx+0.22, by+1.02, col_w-0.44, 0.03, bcol)

        for j, b in enumerate(bullets):
            y0 = by + 1.15 + j * 0.45
            add_rect(slide, bx+0.22, y0+0.12, 0.1, 0.1, bullet_col)
            add_textbox(slide, b, bx+0.4, y0, col_w-0.6, 0.38, font_size=11, colour=sub_col(dark))

    add_rect(slide, 0.6, 6.88, 12.1, 0.4, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "🔒  Our permanent promise: QPON will never charge a percentage commission on your sales. Ever.",
                0.8, 6.91, 11.8, 0.35, font_size=11, bold=True, colour=txt_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 08 — CUSTOMER SEGMENTS  (reordered)  ← MODIFIED
# ══════════════════════════════════════════════════════════════════════════════
def build_s08(slide, dark=True):
    slide_bg(slide, bg_col(dark, DARK_BG, L_BG))
    add_rect(slide, 0, 0, 13.333, 0.06, TEAL)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 8, dark=dark)

    add_textbox(slide, "YOUR NEW CUSTOMER BASE", 0.6, 0.45, 8, 0.28, font_size=10, bold=True, colour=TEAL)
    add_textbox(slide, "Four Segments.", 0.6, 0.78, 7, 0.58, font_size=36, bold=True, colour=txt_col(dark))
    add_textbox(slide, "All Ready to Spend.", 0.6, 1.33, 7, 0.55, font_size=36, bold=True, colour=TEAL)

    # Reordered: Corporate / Digital Nomads / Local Foodies (compelling) / Tourist
    segments = [
        ("🏢", "Corporate Employees",
         "Weekly Regulars · Predictable Revenue",
         PURPLE, PURPLE2,
         "Office workers with QPON via their employer's benefits plan. "
         "Same area every weekday — your most reliable, repeat customer segment.",
         "Consistent weekly revenue"),
        ("💻", "Digital Nomads",
         "Daily Repeat · Long Stay",
         TEAL2, TEAL,
         "Remote professionals staying for months. They eat out daily, use cafés for work, "
         "and build routines fast. One nomad = multiple weekly visits for months.",
         "7+ visits/week potential"),
        ("🍽️", "Urban Taste-Makers",
         "Trend Setters · Social Amplifiers",
         GOLD, GOLD,
         "Savvy locals who discover new spots first and broadcast them to thousands of "
         "followers. One QPON visit = organic marketing you can't buy.",
         "Free social amplification"),
        ("🧳", "Explorers & Tourists",
         "High Spend · Short Stay",
         ORANGE, ORANGE2,
         "Destination travellers seeking authentic local recommendations. "
         "High spend per visit — and QPON is their first stop when they arrive.",
         "Avg. LKR 45K / week spend"),
    ]

    for i, (em, name, tagline, border_col, stat_col, body, kpi) in enumerate(segments):
        bx = 0.52 + i * 3.22
        add_rect(slide, bx, 2.22, 3.08, 4.6, card_col(dark), line_colour=border_col, line_width=1.5)
        add_textbox(slide, em, bx+0.2, 2.38, 0.55, 0.55, font_size=28)
        add_textbox(slide, name,    bx+0.2, 2.98, 2.65, 0.38, font_size=13, bold=True, colour=txt_col(dark))
        add_textbox(slide, tagline, bx+0.2, 3.38, 2.65, 0.32, font_size=9.5, colour=stat_col)
        add_rect(slide, bx+0.2, 3.76, 2.65, 0.04, border_col)
        add_textbox(slide, body, bx+0.2, 3.88, 2.65, 1.2, font_size=10.5, colour=sub_col(dark))
        add_rect(slide, bx+0.2, 5.55, 2.65, 0.35, bg_col(dark,
            RGBColor(0x1A,0x08,0x2A) if border_col==PURPLE else
            (RGBColor(0x00,0x25,0x22) if border_col==TEAL2 else
             (RGBColor(0x2A,0x1E,0x00) if border_col==GOLD else RGBColor(0x28,0x12,0x00))),
            RGBColor(0xF2,0xE8,0xFF) if border_col==PURPLE else
            (RGBColor(0xE0,0xF8,0xF4) if border_col==TEAL2 else
             (RGBColor(0xFF,0xF5,0xD8) if border_col==GOLD else RGBColor(0xFF,0xED,0xE5)))))
        add_textbox(slide, kpi, bx+0.22, 5.58, 2.62, 0.28, font_size=10, bold=True, colour=stat_col, align=PP_ALIGN.CENTER)

    # Bottom bar
    add_rect(slide, 0.52, 6.05, 12.26, 0.75, card_col(dark), line_colour=card_border(dark), line_width=1)
    add_textbox(slide, "📍 All four segments are shown your deals only when physically near you — zero wasted impressions.\n"
                       "🔄 QPON's repeat-visit engine actively drives each segment back to your business.",
                0.72, 6.12, 11.86, 0.58, font_size=10.5, colour=sub_col(dark))

    add_rect(slide, 0, 6.9, 13.333, 0.6, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4, font_size=10, colour=sub_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 09 — COMPETITIVE COMPARISON
# ══════════════════════════════════════════════════════════════════════════════
def build_s09(slide, dark=True):
    slide_bg(slide, bg_col(dark, NAVY, L_BG2))
    add_rect(slide, 0, 0, 13.333, 0.06, ORANGE)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 9, dark=dark)

    add_textbox(slide, "COMPETITIVE COMPARISON", 0.6, 0.45, 8, 0.28, font_size=10, bold=True, colour=ORANGE)
    add_textbox(slide, "How QPON Compares to", 0.6, 0.78, 9, 0.55, font_size=34, bold=True, colour=txt_col(dark))
    add_textbox(slide, "What You're Using Now.", 0.6, 1.3, 9, 0.55, font_size=34, bold=True, colour=ORANGE)

    headers = ["", "QPON ✦", "OTAs", "Delivery Apps", "Social Ads"]
    col_w   = [3.6, 2.2, 2.2, 2.2, 2.2]
    col_x   = [0.55]
    for w in col_w[:-1]:
        col_x.append(col_x[-1] + w + 0.08)

    # Header row
    add_rect(slide, 0.55, 2.08, 12.6, 0.45, bg_col(dark, NAVY3, L_BG3))
    for i, h in enumerate(headers):
        col_bg = TEAL2 if i == 1 else bg_col(dark, NAVY3, L_BG3)
        if i == 1:
            add_rect(slide, col_x[i], 2.08, col_w[i], 0.45, col_bg)
        align = PP_ALIGN.CENTER if i > 0 else PP_ALIGN.LEFT
        c = WHITE if (i==1 or not dark) else sub_col(dark)
        add_textbox(slide, h, col_x[i]+0.1, 2.12, col_w[i]-0.2, 0.35, font_size=11, bold=True, colour=c, align=align)

    rows = [
        ("Commission on sales",      "0% Always",    "20–30%",  "25–30%",  "Sunk cost"),
        ("Real-time analytics",      "✓ Full",       "—",       "Basic",   "✕"),
        ("Full deal control",        "✓ Complete",   "Limited", "Minimal", "✕"),
        ("Geo-targeted reach",       "✓ Precision",  "—",       "—",       "Approx."),
        ("Corporate B2B channel",    "✓ Built-in",   "—",       "—",       "✕"),
        ("Repeat visit engine",      "✓ Active",     "—",       "Partial", "✕"),
    ]
    row_h = 0.58
    for r, (feat, *vals) in enumerate(rows):
        ry = 2.62 + r * row_h
        row_bg = bg_col(dark, NAVY2 if r%2==0 else NAVY, L_BG2 if r%2==0 else L_BG)
        add_rect(slide, 0.55, ry, 12.6, row_h-0.04, row_bg)
        add_textbox(slide, feat, col_x[0]+0.12, ry+0.12, col_w[0]-0.2, row_h-0.16, font_size=12, bold=True, colour=txt_col(dark))
        for j, val in enumerate(vals):
            ci = j + 1
            c = TEAL if ci==1 else (ORANGE2 if "%" in val and ci>1 else sub_col(dark))
            add_textbox(slide, val, col_x[ci]+0.05, ry+0.12, col_w[ci]-0.1, row_h-0.16, font_size=11, colour=c, align=PP_ALIGN.CENTER)

    add_rect(slide, 0, 6.9, 13.333, 0.6, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4, font_size=10, colour=sub_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 10 — REDESIGNED: JUNE 2026 LAUNCH / CORPORATE-FIRST  ← MODIFIED
# ══════════════════════════════════════════════════════════════════════════════
def build_s10(slide, dark=True):
    slide_bg(slide, bg_col(dark, DARK_BG, L_BG))
    add_rect(slide, 0, 0, 13.333, 0.06, PURPLE)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 10, dark=dark)

    add_textbox(slide, "LAUNCH & GROWTH STRATEGY", 0.6, 0.45, 9, 0.28, font_size=10, bold=True, colour=PURPLE2)
    add_textbox(slide, "June 2026 Launch.", 0.6, 0.78, 9, 0.58, font_size=38, bold=True, colour=txt_col(dark))
    add_textbox(slide, "Corporate-First. Ecosystem-Wide.", 0.6, 1.33, 11, 0.55, font_size=34, bold=True, colour=PURPLE2)

    # ── PRIMARY TARGET BOX ───────────────────────────────────────────────────
    add_rect(slide, 0.55, 2.08, 6.0, 4.75, bg_col(dark, RGBColor(0x20,0x06,0x3A), RGBColor(0xF5,0xEB,0xFF)),
             line_colour=PURPLE, line_width=2.0)
    add_rect(slide, 0.55, 2.08, 6.0, 0.42, PURPLE)
    add_textbox(slide, "★  PRIMARY TARGET  —  CORPORATE", 0.7, 2.12, 5.7, 0.34,
                font_size=11, bold=True, colour=WHITE)

    add_textbox(slide, "🏢", 0.85, 2.68, 0.7, 0.7, font_size=36)
    add_textbox(slide, "Corporate Employees & HR Teams",
                1.65, 2.72, 4.7, 0.48, font_size=16, bold=True, colour=PURPLE2)
    add_textbox(slide,
        "QPON launches June 2026 with Corporate as the anchor market. By partnering with companies "
        "before the public launch, we secure a committed, predictable base of repeat visitors for "
        "every onboarded merchant from day one.",
        0.75, 3.28, 5.65, 1.2, font_size=12, colour=sub_col(dark))

    corp_kpis = [
        ("Target", "50+ companies signed pre-launch"),
        ("Employees", "5,000+ active users at launch"),
        ("Visit cadence", "Weekly recurring footfall"),
    ]
    for j, (k, v) in enumerate(corp_kpis):
        y0 = 4.62 + j * 0.52
        add_rect(slide, 0.75, y0, 5.65, 0.44, bg_col(dark, NAVY3, L_BG3))
        add_textbox(slide, k, 0.9, y0+0.07, 1.3, 0.28, font_size=10.5, bold=True, colour=PURPLE2)
        add_textbox(slide, v, 2.3, y0+0.07, 4.0, 0.28, font_size=10.5, colour=txt_col(dark))

    # ── SECONDARY SEGMENTS ───────────────────────────────────────────────────
    add_rect(slide, 7.0, 2.08, 6.0, 0.42, bg_col(dark, NAVY3, L_BG3))
    add_textbox(slide, "SECONDARY SEGMENTS  —  GROWTH PHASE", 7.12, 2.12, 5.8, 0.34,
                font_size=11, bold=True, colour=sub_col(dark))

    secondaries = [
        ("💻", "Digital Nomads",       TEAL,    "Long-stay remote workers build daily routines. High-frequency return visits."),
        ("🍽️","Urban Taste-Makers",   GOLD,    "Local trend-setters drive word-of-mouth and social discovery after launch."),
        ("🧳","Explorers & Tourists",  ORANGE2, "Seasonal high-spenders activate deal discovery for the broader consumer funnel."),
    ]
    for i, (em, name, col, body) in enumerate(secondaries):
        y0 = 2.65 + i * 1.42
        add_rect(slide, 7.0, y0, 6.0, 1.28, card_col(dark), line_colour=col, line_width=1.2)
        add_textbox(slide, em, 7.15, y0+0.3, 0.5, 0.55, font_size=22)
        add_textbox(slide, name, 7.72, y0+0.1, 3.8, 0.38, font_size=13, bold=True, colour=txt_col(dark))
        add_textbox(slide, body, 7.72, y0+0.52, 5.1, 0.62, font_size=11, colour=sub_col(dark))
        add_rect(slide, 7.0, y0, 0.18, 1.28, col)

    # ── Timeline strip ────────────────────────────────────────────────────────
    add_rect(slide, 0.55, 6.0, 12.45, 0.78, bg_col(dark, NAVY3, L_BG3))
    milestones = [
        ("Jan–May 2026", "Corporate pre-sales & merchant onboarding", PURPLE2),
        ("June 2026",    "Public launch — corporate users live",      TEAL),
        ("Q3 2026",      "Digital nomad & local foodie expansion",    GOLD),
        ("Q4 2026",      "Tourist corridor rollout nationwide",       ORANGE2),
    ]
    for i, (date, desc, col) in enumerate(milestones):
        bx = 0.7 + i * 3.12
        add_rect(slide, bx, 6.05, 0.12, 0.62, col)
        add_textbox(slide, date, bx+0.2, 6.05, 2.8, 0.28, font_size=10, bold=True, colour=col)
        add_textbox(slide, desc, bx+0.2, 6.34, 2.8, 0.38, font_size=9.5, colour=sub_col(dark))

    add_rect(slide, 0, 6.9, 13.333, 0.6, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4, font_size=10, colour=sub_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 11 — ONBOARDING SUPPORT
# ══════════════════════════════════════════════════════════════════════════════
def build_s11(slide, dark=True):
    slide_bg(slide, bg_col(dark, NAVY, L_BG2))
    add_rect(slide, 0, 0, 13.333, 0.06, TEAL)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 11, dark=dark)

    add_textbox(slide, "WE'VE GOT YOU", 0.6, 0.45, 8, 0.28, font_size=10, bold=True, colour=TEAL)
    add_textbox(slide, "You're Never on", 0.6, 0.78, 9, 0.58, font_size=36, bold=True, colour=txt_col(dark))
    add_textbox(slide, "Your Own.", 0.6, 1.33, 9, 0.55, font_size=36, bold=True, colour=TEAL)
    add_textbox(slide, "QPON provides every tool and human support you need to get set up fast and keep growing.",
                0.6, 2.0, 8.5, 0.5, font_size=13, colour=sub_col(dark))

    supports = [
        ("🧑‍💼","Personal Sales Walkthrough",    "A team member guides you through setup — in person or video call. No questions too small."),
        ("🎬","Demo Videos & FAQ Portal",        "Step-by-step guides and a searchable help centre available 24/7."),
        ("📞","Priority Support Line",           "Direct WhatsApp and email support. Founding merchants get priority — response within hours."),
        ("🤖","AI Onboarding (2026 Roadmap)",    "Intelligent chatbot that guides setup and suggests deal optimisations automatically."),
    ]
    for i, (em, title, body) in enumerate(supports):
        y0 = 2.72 + i * 1.05
        add_rect(slide, 0.6, y0, 7.0, 0.92, card_col(dark), line_colour=TEAL2, line_width=1)
        add_textbox(slide, em, 0.78, y0+0.2, 0.5, 0.5, font_size=22)
        add_textbox(slide, title, 1.4, y0+0.1, 5.9, 0.36, font_size=13, bold=True, colour=txt_col(dark))
        add_textbox(slide, body,  1.4, y0+0.5, 5.9, 0.34, font_size=11, colour=sub_col(dark))

    # Welcome Kit
    add_rect(slide, 8.2, 2.5, 4.75, 4.3, bg_col(dark, RGBColor(0x00,0x25,0x22), RGBColor(0xE0,0xF8,0xF5)),
             line_colour=TEAL, line_width=2.0)
    add_textbox(slide, "🎁", 9.9, 2.65, 1.4, 1.1, font_size=52, align=PP_ALIGN.CENTER)
    add_textbox(slide, "Founding Merchant\nWelcome Kit", 8.35, 3.75, 4.45, 0.68,
                font_size=16, bold=True, colour=txt_col(dark), align=PP_ALIGN.CENTER)
    add_textbox(slide, "First 50 merchants only", 8.35, 4.45, 4.45, 0.3,
                font_size=10, colour=TEAL, align=PP_ALIGN.CENTER)
    kit_items = ["📱  Custom QR standee for your counter",
                 "🪧  Branded table cards & stickers",
                 "🏅  'QPON Partner' window badge",
                 "📘  Printed quick-start guide"]
    for j, item in enumerate(kit_items):
        y0 = 4.85 + j * 0.42
        add_rect(slide, 8.35, y0, 4.45, 0.36, bg_col(dark, NAVY3, L_BG3))
        add_textbox(slide, item, 8.5, y0+0.04, 4.2, 0.28, font_size=11, colour=txt_col(dark))

    add_rect(slide, 0, 6.9, 13.333, 0.6, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4, font_size=10, colour=sub_col(dark), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 12 — FOUNDING OFFER  (no cashback)  ← MODIFIED
# ══════════════════════════════════════════════════════════════════════════════
def build_s12(slide, dark=True):
    bg = bg_col(dark, RGBColor(0xB8,0x34,0x10), RGBColor(0xFF,0x6B,0x35))
    slide_bg(slide, bg)
    add_rect(slide, 0, 0, 13.333, 0.06, RGBColor(0xFF,0xFF,0xFF) if not dark else RGBColor(0xC0,0x48,0x18))
    add_logo(slide, dark=False)   # always white on orange
    add_slide_num(slide, 12, dark=False)

    add_textbox(slide, "EXCLUSIVE LAUNCH OFFER", 0.6, 0.45, 10, 0.28, font_size=10, bold=True, colour=RGBColor(0xFF,0xE0,0xD0))
    add_textbox(slide, "Join as a Founding Merchant.", 0.6, 0.8, 12, 0.62, font_size=42, bold=True, colour=WHITE)
    add_textbox(slide, "Win Before Launch.", 0.6, 1.4, 12, 0.58, font_size=40, bold=True, colour=RGBColor(0xFF,0xE8,0xD0))
    add_textbox(slide,
        "Limited to 50 businesses. Once these spots are filled, the founding tier closes permanently\n"
        "and pricing reverts to standard rates.",
        0.6, 2.08, 12.0, 0.65, font_size=14, colour=RGBColor(0xFF,0xDE,0xC8))

    # 3 benefit cards (cashback removed)
    benefits = [
        ("🆓","3 Months Free",       "Full platform access with zero fees.\nNo credit card required to start."),
        ("🏆","Priority Placement",  "Your business appears at the top\nwhen customers search near you."),
        ("🎁","Welcome Kit",         "QR standee, table cards, and stickers\n— everything to signal you're on QPON."),
    ]
    card_start = (13.333 - 3*3.8 - 2*0.35) / 2
    for i, (em, title, body) in enumerate(benefits):
        bx = card_start + i * 4.15
        add_rect(slide, bx, 2.95, 3.8, 2.55, RGBColor(0xFF,0xFF,0xFF), line_colour=RGBColor(0xFF,0xCC,0xAA), line_width=1)
        add_textbox(slide, em, bx+1.5, 3.1, 0.8, 0.7, font_size=36, align=PP_ALIGN.CENTER)
        add_textbox(slide, title, bx+0.2, 3.88, 3.4, 0.45, font_size=15, bold=True, colour=NAVY, align=PP_ALIGN.CENTER)
        add_textbox(slide, body,  bx+0.2, 4.36, 3.4, 0.9, font_size=12, colour=L_TEXT2, align=PP_ALIGN.CENTER)

    # Commitment strip
    strips = [
        ("🤝","Dedicated Onboarding",      "Personal setup walkthrough"),
        ("🔒","0% Commission — Forever",    "Locked in for founding partners"),
        ("📍","Priority geo visibility",    "Top of search in your area"),
    ]
    sx = (13.333 - 3*3.8 - 2*0.35) / 2
    for i, (em, t, s) in enumerate(strips):
        bx = sx + i * 4.15
        add_rect(slide, bx, 5.72, 3.8, 0.88, RGBColor(0x8B,0x22,0x00))
        add_textbox(slide, em + "  " + t, bx+0.15, 5.78, 3.5, 0.35, font_size=12, bold=True, colour=WHITE)
        add_textbox(slide, s, bx+0.15, 6.12, 3.5, 0.3, font_size=10, colour=RGBColor(0xFF,0xCC,0xAA))

    add_rect(slide, 0, 6.9, 13.333, 0.6, RGBColor(0x80,0x1C,0x00))
    add_textbox(slide, "qpon.lk  ·  inquiries@qpon.lk  ·  +94 75 977 8858", 0.5, 6.97, 12.5, 0.4, font_size=10, colour=RGBColor(0xFF,0xCC,0xAA), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE 13 — CALL TO ACTION  (with working hyperlink)  ← MODIFIED
# ══════════════════════════════════════════════════════════════════════════════
def build_s13(slide, dark=True):
    slide_bg(slide, bg_col(dark, RGBColor(0x0D,0x10,0x30), RGBColor(0xF0,0xF3,0xFF)))
    add_rect(slide, 0, 0, 13.333, 0.06, TEAL)
    add_logo(slide, dark=dark)
    add_slide_num(slide, 13, dark=dark)

    add_textbox(slide, "PARTNER WITH QPON TODAY", 0.6, 0.45, 10, 0.28, font_size=10, bold=True, colour=TEAL)
    add_textbox(slide, "Your Business Deserves", 0.6, 0.82, 12, 0.62, font_size=44, bold=True, colour=txt_col(dark))
    add_textbox(slide, "Better Than 30% Commission.", 0.6, 1.42, 12, 0.6, font_size=42, bold=True, colour=ORANGE)
    add_textbox(slide,
        "Join QPON now. Keep your margin, own your customers, and fill your tables —\n"
        "on your terms, with zero commission, starting free.",
        0.6, 2.12, 11.5, 0.75, font_size=15, colour=sub_col(dark))

    # 3 contact blocks
    contacts = [
        ("🌐", "Sign Up Online",    "www.qpon.lk",           "https://www.qpon.lk", ORANGE),
        ("📞", "Call Us",           "+94 75 977 8858",        None,                  TEAL2),
        ("✉️", "Email Us",          "inquiries@qpon.lk",     "mailto:inquiries@qpon.lk", PURPLE),
    ]
    bw = 3.6
    start_bx = (13.333 - 3*bw - 2*0.32) / 2
    for i, (em, label, detail, url, col) in enumerate(contacts):
        bx = start_bx + i*(bw+0.32)
        add_rect(slide, bx, 3.1, bw, 1.55,
                 col if i==0 else card_col(dark),
                 line_colour=col, line_width=1.8)
        add_textbox(slide, em, bx+1.5, 3.2, 0.7, 0.55, font_size=28, align=PP_ALIGN.CENTER)
        add_textbox(slide, label, bx+0.15, 3.8, bw-0.3, 0.38,
                    font_size=14, bold=True,
                    colour=WHITE if i==0 else txt_col(dark),
                    align=PP_ALIGN.CENTER)
        if url:
            add_hyperlink_textbox(slide, detail, url,
                                  bx+0.15, 4.22, bw-0.3, 0.32,
                                  font_size=12, bold=False,
                                  colour=WHITE if i==0 else col,
                                  align=PP_ALIGN.CENTER)
        else:
            add_textbox(slide, detail, bx+0.15, 4.22, bw-0.3, 0.32,
                        font_size=12, colour=TEAL if i==1 else sub_col(dark),
                        align=PP_ALIGN.CENTER)

    # 3-beat summary
    beats = [
        ("0%",      "Commission forever",   ORANGE),
        ("3 Months","Free to start",        TEAL),
        ("50 Spots","Founding tier only",   GOLD),
    ]
    for i, (stat, lbl, col) in enumerate(beats):
        bx = 2.0 + i * 4.0
        add_textbox(slide, stat, bx, 5.0, 3.0, 0.62, font_size=40, bold=True, colour=col, align=PP_ALIGN.CENTER)
        add_textbox(slide, lbl,  bx, 5.62, 3.0, 0.3,  font_size=11, colour=sub_col(dark), align=PP_ALIGN.CENTER)
        if i < 2:
            add_textbox(slide, "→", bx+3.1, 5.15, 0.5, 0.4, font_size=22, colour=sub_col(dark))

    # Footer
    add_rect(slide, 0, 6.62, 13.333, 0.04, bg_col(dark, NAVY3, L_BG3))
    add_rect(slide, 0, 6.66, 13.333, 0.84, bg_col(dark, DARK_BG, L_BG3))
    add_textbox(slide, "© 2026 QPON", 0.5, 6.75, 2.5, 0.3, font_size=10, colour=sub_col(dark))
    add_hyperlink_textbox(slide, "qpon.lk", "https://www.qpon.lk",
                          5.5, 6.75, 2.2, 0.3, font_size=10, bold=False,
                          colour=TEAL, align=PP_ALIGN.CENTER)
    add_textbox(slide, "Issac Heshan Suppiah", 10.0, 6.75, 3.1, 0.3,
                font_size=10, colour=sub_col(dark), align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
#  MASTER BUILD FUNCTION
# ══════════════════════════════════════════════════════════════════════════════
BUILDERS = [
    build_s01, build_s02, build_s03, build_s04, build_s05,
    build_s06, build_s07, build_s08, build_s09, build_s10,
    build_s11, build_s12, build_s13,
]

def build_deck(dark=True):
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    blank_layout = prs.slide_layouts[6]  # blank

    for builder in BUILDERS:
        slide = prs.slides.add_slide(blank_layout)
        builder(slide, dark=dark)

    return prs


if __name__ == "__main__":
    print("Building Version 1 (Dark theme)...")
    v1 = build_deck(dark=True)
    v1.save("/home/user/LMS-Module-/QPON_Merchant_Pitch_V1_Dark.pptx")
    print("  → QPON_Merchant_Pitch_V1_Dark.pptx saved")

    print("Building Version 2 (Light theme)...")
    v2 = build_deck(dark=False)
    v2.save("/home/user/LMS-Module-/QPON_Merchant_Pitch_V2_Light.pptx")
    print("  → QPON_Merchant_Pitch_V2_Light.pptx saved")

    print("Done!")
