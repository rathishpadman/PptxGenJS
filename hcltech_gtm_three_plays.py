"""HCLTech slide: Re-energizing our GTM — three plays, one motion."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

TEMPLATE = '/root/.claude/skills/synced/hcltech-deck/assets/hcltech_base.pptx'
OUT = '/home/user/PptxGenJS/hcltech_gtm_three_plays.pptx'

BLACK       = RGBColor(0x00, 0x00, 0x00)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
PURPLE      = RGBColor(0x5F, 0x1E, 0xBE)
PURPLE_DEEP = RGBColor(0x41, 0x14, 0x82)
BLUE        = RGBColor(0x0F, 0x5F, 0xDC)
ICE_BLUE    = RGBColor(0xDC, 0xE6, 0xF0)
GREY_1      = RGBColor(0x82, 0x91, 0xA0)
GREY_3      = RGBColor(0xC8, 0xD2, 0xDD)
GREY_4      = RGBColor(0xE6, 0xEB, 0xF5)
FONT = 'Aptos'

prs = Presentation(TEMPLATE)
slide = prs.slides.add_slide(prs.slide_layouts[16])
S = slide.shapes
I = Inches


def style(run, size, *, bold=False, color=BLACK):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def fill_tf(tf, paras, *, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
            line_spacing=None, space_after=0):
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, prun in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        if line_spacing:
            p.line_spacing = line_spacing
        for text, size, bold, color in prun:
            style(p.add_run(), size, bold=bold, color=color)
            p.runs[-1].text = text


def textbox(x, y, w, h, runs, **kw):
    tb = S.add_textbox(I(x), I(y), I(w), I(h))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    fill_tf(tf, runs if runs and isinstance(runs[0], list) else [runs], **kw)
    return tb


def shape(kind, x, y, w, h, *, fill=None, line=None, line_w=0.75, adj=None,
          rot=None, dash=False):
    s = S.add_shape(kind, I(x), I(y), I(w), I(h))
    if adj is not None:
        vals = adj if isinstance(adj, (list, tuple)) else [adj]
        for i, v in enumerate(vals):
            if i < len(s.adjustments):
                s.adjustments[i] = v
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
        if dash:
            s.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    if rot is not None:
        s.rotation = rot
    s.shadow.inherit = False
    s.text_frame.word_wrap = True
    return s


def poly(pts, color, width, close=False, dash=False):
    b = S.build_freeform(I(pts[0][0]), I(pts[0][1]))
    b.add_line_segments([(I(x), I(y)) for x, y in pts[1:]], close=close)
    s = b.convert_to_shape()
    s.fill.background()
    s.line.color.rgb = color
    s.line.width = Pt(width)
    if dash:
        s.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    s.shadow.inherit = False
    return s


def rule(x, y, w, color, width=0.75):
    return poly([(x, y), (x + w, y)], color, width)


def chip(kind, x, y, w, h, text, size, txt_color, *, fill=None, line=None,
         adj=None, bold=True, lines=None):
    c = shape(kind, x, y, w, h, fill=fill, line=line, adj=adj)
    tf = c.text_frame
    tf.margin_left = tf.margin_right = I(0.05)
    tf.margin_top = tf.margin_bottom = 0
    fill_tf(tf, lines or [[(text, size, bold, txt_color)]],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=0.95)
    return c


# ------------------------------------------------------------------ glyphs ---
def glyph(name, cx, cy, s, color):
    u = s / 2.0
    lw = max(1.0, s * 2.2)
    if name == 'target':
        shape(MSO_SHAPE.OVAL, cx - u, cy - u, s, s, line=color, line_w=lw)
        shape(MSO_SHAPE.OVAL, cx - u * .55, cy - u * .55, s * .55, s * .55,
              line=color, line_w=lw)
        shape(MSO_SHAPE.OVAL, cx - u * .18, cy - u * .18, s * .18, s * .18, fill=color)
    elif name == 'people':
        for dx in (-u * .44, u * .10):
            shape(MSO_SHAPE.OVAL, dx + cx, cy - u * .74, s * .30, s * .30,
                  line=color, line_w=lw)
            shape(MSO_SHAPE.ROUND_2_SAME_RECTANGLE, dx + cx - s * .05,
                  cy - u * .12, s * .40, s * .44, line=color, line_w=lw,
                  adj=(0.45, 0.0))
    elif name == 'doc':
        shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx - u * .62, cy - u * .86,
              s * .62, s * .86, line=color, line_w=lw, adj=0.12)
        for i in range(3):
            rule(cx - u * .34, cy - u * .48 + i * s * .21, s * .34, color, lw * .85)
    elif name == 'nodes':
        shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx - s * .10, cy - u * .74,
              s * .20, s * .20, fill=color, adj=0.25)
        for dx in (-.32, -.08, .16):
            shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + s * dx, cy + u * .28,
                  s * .16, s * .16, fill=color, adj=0.25)
        rule(cx - s * .24, cy + u * .08, s * .48, color, lw * .85)
    elif name == 'chart_up':
        for i, hh in enumerate((.34, .58, .84)):
            shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx - u * .70 + i * s * .27,
                  cy + u * .62 - s * hh, s * .18, s * hh, fill=color, adj=0.3)
    elif name == 'database':
        for i in range(3):
            shape(MSO_SHAPE.OVAL, cx - u * .70, cy - u * .82 + i * s * .32,
                  s * .70, s * .30, line=color, line_w=lw)
    elif name == 'gear':
        shape(MSO_SHAPE.GEAR_6, cx - u * .86, cy - u * .86, s * .86, s * .86,
              line=color, line_w=lw)
        shape(MSO_SHAPE.OVAL, cx - u * .24, cy - u * .24, s * .24, s * .24,
              line=color, line_w=lw)


# ------------------------------------------------------------------ header ---
textbox(0.65, 0.50, 12.03, 0.58,
        [[('Re-energizing our GTM — ', 28, True, BLACK),
          ('three plays, one motion', 28, True, PURPLE)]])
textbox(0.65, 1.14, 12.03, 0.32,
        [('Vertical sales embedded, DBS and DPO on one story, partners '
          'co-creating — early proof already in.', 16, False, GREY_1)])

# ------------------------------------------------------------------- cards ---
TOP, CH, CW = 1.56, 5.10, 3.85
PX = [0.65, 4.74, 8.83]
PAD = 0.26

PILLARS = [
    {'color': PURPLE_DEEP, 'num': '01', 'head': 'Embedded with\nVertical Sales',
     'desc': 'Deepen vertical alignment and activate BPO platform '
             'conversations at scale.',
     'diagram': 'orbit',
     'highlights': [('1st', ' Advisor Day completed'),
                    ('26+', ' advisors engaged'),
                    ('89%', ' willing to take AFP to market'),
                    ('14', ' successful interactions')],
     'note': 'KPMG advisors, Growth Market (AUS)'},
    {'color': PURPLE, 'num': '02', 'head': 'One story —\nDBS + DPO',
     'desc': 'One integrated narrative that delivers greater value together.',
     'diagram': 'interlock',
     'highlights': [('5+', ' CFO roundtables planned, FY27'),
                    ('15+', ' analyst briefings planned, FY27')],
     'note': ''},
    {'color': BLUE, 'num': '03', 'head': 'Co-created by the\nPartner ecosystem',
     'desc': 'Leverage hyperscalers and partners to build, innovate and '
             'scale together.',
     'diagram': 'network',
     'highlights': [('96', ' media reactions — Autonomous Finance Platform '
                            'with Google'),
                    ('Krisp.ai', ' — real-time voice translation'),
                    ('Stutt.ai', ' — automated collections')],
     'note': ''},
]


def diagram(kind, x, cy, color):
    cx = x + CW / 2
    if kind == 'orbit':
        shape(MSO_SHAPE.OVAL, cx - 1.20, cy - 0.58, 2.40, 1.16,
              line=GREY_3, dash=True)
        shape(MSO_SHAPE.OVAL, cx - 0.44, cy - 0.44, 0.88, 0.88, fill=color)
        glyph('people', cx, cy, 0.40, WHITE)
        for (dx, dy), g in zip([(-1.10, -0.40), (1.10, -0.40),
                                (-1.10, 0.40), (1.10, 0.40)],
                               ['target', 'doc', 'nodes', 'chart_up']):
            shape(MSO_SHAPE.OVAL, cx + dx - 0.21, cy + dy - 0.21, 0.42, 0.42,
                  fill=WHITE, line=GREY_3)
            glyph(g, cx + dx, cy + dy, 0.22, color)
    elif kind == 'interlock':
        for sx, lab, g in ((-1.38, 'DBS', 'database'), (0.38, 'DPO', 'gear')):
            shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + sx, cy - 0.50, 1.00, 1.00,
                  fill=color, adj=0.12)
            glyph(g, cx + sx + 0.50, cy - 0.16, 0.32, WHITE)
            textbox(cx + sx, cy + 0.10, 1.00, 0.26, [(lab, 12, True, WHITE)],
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        chip(MSO_SHAPE.ROUNDED_RECTANGLE, cx - 0.47, cy - 0.47, 0.94, 0.94,
             None, 0, color, fill=WHITE, line=color, adj=0.14,
             lines=[[('Stronger', 9, True, color)],
                    [('Customer', 9, True, color)],
                    [('Value', 9, True, color)]])
    elif kind == 'network':
        shape(MSO_SHAPE.OVAL, cx - 0.40, cy - 0.40, 0.80, 0.80, fill=color)
        glyph('nodes', cx, cy, 0.36, WHITE)
        for (sx, dy), lab in zip([(-1.66, -0.40), (0.54, -0.40),
                                  (-1.66, 0.40), (0.54, 0.40)],
                                 ['Hyperscalers', 'MS & AWS',
                                  'Startup Play', 'GTM Momentum']):
            poly([(cx + sx + (1.12 if sx < 0 else 0.0), cy + dy),
                  (cx + (-0.34 if sx < 0 else 0.34), cy + dy * 0.45)],
                 GREY_3, 0.75, dash=True)
            chip(MSO_SHAPE.ROUNDED_RECTANGLE, cx + sx, cy + dy - 0.16, 1.12, 0.32,
                 lab, 8.5, color, fill=GREY_4, adj=0.5)


for i, p in enumerate(PILLARS):
    x, col = PX[i], p['color']
    shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, TOP, CW, CH,
          fill=WHITE, line=GREY_3, adj=0.025)

    chip(MSO_SHAPE.ROUNDED_RECTANGLE, x + PAD, TOP + 0.24, 0.58, 0.40,
         p['num'], 15, WHITE, fill=col, adj=0.18)
    textbox(x + 0.96, TOP + 0.18, CW - 1.22, 0.52,
            [[(ln, 15, True, col)] for ln in p['head'].split('\n')],
            line_spacing=0.95)
    textbox(x + PAD, TOP + 0.80, CW - 2 * PAD, 0.42,
            [(p['desc'], 10.5, False, GREY_1)], line_spacing=1.1)

    diagram(p['diagram'], x, TOP + 1.95, col)

    rule(x + PAD, TOP + 2.76, CW - 2 * PAD, GREY_4)
    textbox(x + PAD, TOP + 2.88, 2.0, 0.22, [('HIGHLIGHTS', 9, True, col)])
    for j, (metric, rest) in enumerate(p['highlights']):
        ry = TOP + 3.18 + j * 0.40
        shape(MSO_SHAPE.OVAL, x + PAD + 0.02, ry + 0.10, 0.09, 0.09, fill=col)
        textbox(x + PAD + 0.24, ry, CW - 2 * PAD - 0.24, 0.38,
                [[(metric, 11, True, col), (rest, 9.5, False, BLACK)]],
                line_spacing=1.1)
    if p['note']:
        textbox(x + PAD + 0.24, TOP + 3.18 + len(p['highlights']) * 0.40 + 0.06,
                CW - 2 * PAD - 0.24, 0.22, [(p['note'], 8.5, False, GREY_1)])

# ------------------------------------------------------------------ footer ---
FY = 6.86
rule(1.60, FY + 0.13, 3.05, GREY_3)
rule(8.70, FY + 0.13, 3.05, GREY_3)
textbox(0.65, FY, 12.03, 0.26,
        [('Aligned. Integrated. Innovative. Impactful.', 12.5, True, PURPLE_DEEP)],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

prs.save(OUT)
print('saved', OUT, '| shapes:', len(slide.shapes))
