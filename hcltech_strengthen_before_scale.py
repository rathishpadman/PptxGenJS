"""HCLTech slide: Strengthening before we scale out — three-phase layout."""
import sys
sys.path.insert(0, '.')
from pptx import Presentation
from hcltech_board_common import *          # noqa: F401,F403

TEMPLATE = '/root/.claude/skills/synced/hcltech-deck/assets/hcltech_base.pptx'
OUT = '/home/user/PptxGenJS/hcltech_strengthen_before_scale.pptx'

prs = Presentation(TEMPLATE)
slide = prs.slides.add_slide(prs.slide_layouts[16])
use(slide.shapes)

textbox(0.65, 0.50, 12.03, 0.55,
        [[('Strengthening before ', 28, True, BLACK),
          ('we scale out', 28, True, PURPLE)]])
textbox(0.65, 1.12, 12.03, 0.32,
        [('Three phases, eight moves — then one launch decision.',
          16, False, GREY_1)])

c = chip(MSO_SHAPE.ROUNDED_RECTANGLE, 10.62, 1.36, 2.06, 0.46,
         [[('   SCALE OUT', 13, True, BLUE)]], fill=ICE_BLUE, adj=0.22)
shape(MSO_SHAPE.UP_ARROW, 10.90, 1.47, 0.20, 0.24, fill=BLUE, adj=(0.42, 0.42))

PHASES = [
    ('PHASE 1', 'Strengthen the core', PURPLE_DEEP,
     [('cubes', '1', 'Capability units, first time'),
      ('shield', '2', 'Transition and quality')],
     'Delivery quality becomes dependable'),
    ('PHASE 2', 'Simplify the structure', PURPLE,
     [('layers', '3', 'Delayering — fewer layers'),
      ('people', '4', 'Span of control'),
      ('target', '5', 'Decision rights')],
     'Decisions move at speed'),
    ('PHASE 3', 'Align to the customer', BLUE,
     [('union', '6', 'ISD in key accounts'),
      ('nodes', '7', 'Transformation reporting'),
      ('star', '8', 'VDU customer alignment')],
     'Customer outcomes lead delivery'),
]

CT, CH, CW, CGAP = 2.06, 3.94, 3.85, 0.24
for i, (tag, name, col, items, outcome) in enumerate(PHASES):
    x = 0.65 + i * (CW + CGAP)
    shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, CT, CW, CH,
          fill=WHITE, line=GREY_3, adj=0.045)
    chip(MSO_SHAPE.ROUND_2_SAME_RECTANGLE, x, CT, CW, 0.72,
         [[(tag, 9, True, PERIWINKLE)], [(name, 15, True, WHITE)]],
         fill=col, adj=(0.16, 0.0))

    region_top, region_h = CT + 0.90, 2.16
    pitch = 0.72
    y0 = region_top + (region_h - len(items) * pitch) / 2
    for j, (icon, num, label) in enumerate(items):
        iy = y0 + j * pitch
        shape(MSO_SHAPE.OVAL, x + 0.30, iy + 0.06, 0.46, 0.46, fill=GREY_4)
        glyph(icon, x + 0.53, iy + 0.29, 0.24, col)
        textbox(x + 0.92, iy, CW - 1.22, 0.58,
                [[(num + '   ', 10, True, col), (label, 11.5, True, BLACK)]],
                anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)

    rule(x + 0.30, CT + 3.14, CW - 0.60, GREY_4)
    textbox(x + 0.30, CT + 3.28, CW - 0.60, 0.44,
            [(outcome, 10.5, False, col)], anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.1)

for gx in (0.65 + CW + CGAP / 2, 0.65 + 2 * CW + 1.5 * CGAP):
    shape(MSO_SHAPE.ISOSCELES_TRIANGLE, gx - 0.09, CT + 1.82, 0.18, 0.24,
          fill=GREY_3, rot=90)

band = shape(MSO_SHAPE.ROUNDED_RECTANGLE, 0.65, 6.22, 12.03, 0.62, adj=0.14)
gradient(band, PURPLE_DEEP, BLUE)
shape(MSO_SHAPE.OVAL, 1.02, 6.36, 0.34, 0.34, fill=WHITE)
glyph('shield', 1.19, 6.53, 0.19, PURPLE_DEEP)
textbox(1.56, 6.22, 10.80, 0.62,
        [[('Strengthen the core. Define the organization. ', 13, True, WHITE),
          ('Then scale with confidence.', 13, True, PERIWINKLE)]],
        anchor=MSO_ANCHOR.MIDDLE)

prs.save(OUT)
print('saved', OUT, '| shapes:', len(slide.shapes))
