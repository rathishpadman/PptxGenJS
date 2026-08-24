"""Process Automation Due Diligence — training pack (generic, 18 slides)."""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Inches, Pt
from hcltech_board_common import *   # noqa: F401,F403

TEMPLATE = '/root/.claude/skills/synced/hcltech-deck/assets/hcltech_base.pptx'
OUT = '/home/user/PptxGenJS/dd_training_pack.pptx'

ICE_PURPLE = RGBColor(0xEF, 0xEA, 0xFB)

prs = Presentation(TEMPLATE)


def new_slide():
    s = prs.slides.add_slide(prs.slide_layouts[16])
    use(s.shapes)
    return s


def header(kicker, title_runs, subtitle=None, *, dark=False):
    fg = WHITE if dark else BLACK
    sub_col = PERIWINKLE if dark else GREY_1
    if kicker:
        textbox(0.65, 0.36, 11.0, 0.26, [(kicker.upper(), 10.5, True,
                PERIWINKLE if dark else BLUE)])
    textbox(0.65, 0.60, 12.03, 0.55, [[(t, 26, True, c) for t, c in title_runs]])
    if subtitle:
        textbox(0.65, 1.18, 12.03, 0.30, [(subtitle, 13.5, False, sub_col)])


def footer(n, section):
    textbox(0.65, 7.16, 6.0, 0.24, [(section, 8.5, False, GREY_2)])
    textbox(11.9, 7.16, 0.78, 0.24, [(str(n), 8.5, False, GREY_2)],
            align=PP_ALIGN.RIGHT)


def card(x, y, w, h, *, fill=WHITE, line=GREY_3, adj=0.035):
    return shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill, line=line,
                 adj=adj)


def tag(x, y, w, h, text, *, fill, txt_col=WHITE, size=9.5):
    return chip(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
                [[(text, size, True, txt_col)]], fill=fill, adj=0.5)


PY_PATH = '/tmp/claude-0/-home-user-PptxGenJS/8d881808-4bb5-5c08-b7d1-a61f8d8b82e5/scratchpad/dd_pack'


def lerp(a, b, t):
    return RGBColor(*(int(x + (y - x) * t) for x, y in zip(a, b)))


# ═══════════════════════════════════════════════════════ 1 · TITLE ══════════
def slide_title():
    s = new_slide()
    bg = shape(MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5)
    gradient(bg, PURPLE_DEEP, PURPLE, angle=35.0)
    shape(MSO_SHAPE.OVAL, 10.6, -1.7, 4.3, 4.3, fill=LAVENDER)
    shape(MSO_SHAPE.OVAL, -1.4, 5.0, 3.4, 3.4, line=PERIWINKLE, line_w=1.25)

    tag(0.9, 1.55, 3.55, 0.40, 'INTERNAL TRAINING PLAYBOOK', fill=LAVENDER,
        txt_col=WHITE)
    textbox(0.85, 2.15, 11.3, 1.55,
            [[('Transformation', 42, True, WHITE)],
             [('Due Diligence', 42, True, PERIWINKLE)]], line_spacing=1.0)
    textbox(0.9, 3.75, 9.8, 0.5,
            [('A field playbook for running the DD workshop, scoring what '
              "you hear, and building a case that survives review.",
              15, False, RGBColor(0xC7, 0xD6, 0xF5))], line_spacing=1.2)
    rule(0.9, 4.55, 2.2, PERIWINKLE, 1.5)
    textbox(0.9, 4.75, 8, 0.3,
            [('Prepared for internal upskilling  ·  adapt before client use',
              10.5, False, RGBColor(0x9C, 0xB3, 0xE8))])


# ═══════════════════════════════════════════════════════ 2 · AGENDA ═════════
def slide_agenda():
    new_slide()
    header('Training roadmap', [('What we’ll cover ', BLACK),
                                ('today', BLUE)],
           'Eight sections, one worked example, one cheat sheet to keep.')
    ITEMS = [
        ('1', 'The DD lifecycle', 'Five stages, start to sign-off'),
        ('2', 'Who’s in the room, and the timeline', 'The POD, and a typical schedule'),
        ('3', 'Preparing & running the workshop', 'What to gather, what to ask, what to listen for'),
        ('4', 'Scoring what you heard', 'Complexity and impact, simplified for the field'),
        ('5', 'Choosing the technology', 'Five archetypes and when each one fits'),
        ('6', 'Building the business case', 'ROI mechanics in plain terms'),
        ('7', 'Worked example: Order-to-Cash', 'One process, scored, archetyped, priced'),
        ('8', 'Getting to sign-off', 'The gates before an executive ever sees this'),
    ]
    RT, RH = 1.62, 0.62
    for i, (num, title, desc) in enumerate(ITEMS):
        y = RT + i * RH
        if i % 2 == 0:
            shape(MSO_SHAPE.RECTANGLE, 0.65, y, 12.03, RH, fill=GREY_4)
        chip(MSO_SHAPE.OVAL, 0.95, y + 0.11, 0.40, 0.40,
             [[(num, 13, True, WHITE)]], fill=PURPLE if i < 5 else BLUE)
        textbox(1.55, y, 3.85, RH, [(title, 13, True, BLACK)],
                anchor=MSO_ANCHOR.MIDDLE)
        textbox(5.55, y, 6.90, RH, [(desc, 11, False, GREY_1)],
                anchor=MSO_ANCHOR.MIDDLE)
    footer(2, 'Agenda')


# ═══════════════════════════════════════════════════ 1 · THE DD LIFECYCLE ═══
def slide_lifecycle():
    new_slide()
    header('Section 1', [('The DD lifecycle — ', BLACK), ('five stages', BLUE)],
           'Every workshop you run moves a candidate process through this '
           'same pipeline.')
    STAGES = [
        ('doc', 'Pre-DD Prep', 'Collect SOPs, volumes, app inventory'),
        ('search', 'Workshop & Shadowing', 'Walk the process, watch it live'),
        ('target', 'Feasibility & Archetyping', 'Score complexity, pick the technology'),
        ('chart_up', 'Business Case', 'Baseline cost, savings, ROI, payback'),
        ('shield', 'Playback & Sign-off', 'Validate, present, get the green light'),
    ]
    n = len(STAGES)
    W, GAP = 2.10, 0.32
    total = n * W + (n - 1) * GAP
    X0 = (13.333 - total) / 2
    Y, H = 2.60, 2.35
    for i, (icon, title, desc) in enumerate(STAGES):
        x = X0 + i * (W + GAP)
        col = lerp((0x41, 0x14, 0x82), (0x0F, 0x5F, 0xDC), i / (n - 1))
        card(x, Y, W, H, fill=WHITE, line=GREY_3)
        shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Y, W, 0.06, fill=col, adj=0.5)
        shape(MSO_SHAPE.OVAL, x + W / 2 - 0.34, Y + 0.28, 0.68, 0.68, fill=col)
        glyph(icon, x + W / 2, Y + 0.62, 0.34, WHITE)
        chip(MSO_SHAPE.OVAL, x + 0.10, Y + 0.14, 0.32, 0.32,
             [[(str(i + 1), 11, True, WHITE)]], fill=col)
        textbox(x + 0.14, Y + 1.10, W - 0.28, 0.56, [(title, 12.5, True, BLACK)],
                align=PP_ALIGN.CENTER, line_spacing=1.0)
        textbox(x + 0.14, Y + 1.66, W - 0.28, 0.62, [(desc, 9.5, False, GREY_1)],
                align=PP_ALIGN.CENTER, line_spacing=1.1)
        if i < n - 1:
            arrowhead(x + W + GAP / 2, Y + H / 2, GREY_2, 0.08, 'right')
    textbox(X0, Y + H + 0.35, total, 0.30,
            [('You are training on Stages 2–4 today — where most DD leads '
              'either earn or lose credibility.', 10.5, False, GREY_1)],
            align=PP_ALIGN.CENTER)
    footer(3, 'The DD lifecycle')




# ═══════════════════════ 2 · WHO’S IN THE ROOM, AND HOW LONG IT TAKES ═══════
def slide_pod_timeline():
    new_slide()
    header('Section 2', [('Who’s in the room, ', BLACK), ('and how long it takes', BLUE)],
           'The standing POD for a DD workshop, and a typical timeline for '
           'one L4 process.')

    card(0.65, 1.82, 5.75, 4.55)
    tag(0.90, 2.02, 2.4, 0.36, 'THE POD', fill=PURPLE)
    ROLES = [
        ('people', 'Ops Manager', 'Owns the process P&L; makes the prioritization call'),
        ('search', 'Ops SME', 'Runs the process daily — the source of L4 ground truth'),
        ('doc', 'Process Expert', 'Owns the SOP and quality lens; flags where reality has drifted'),
        ('gear', 'Solution Team Member', 'Assesses feasibility; owns the archetype and effort estimate'),
    ]
    for i, (icon, role, desc) in enumerate(ROLES):
        ry = 2.60 + i * 0.92
        col = lerp((0x41, 0x14, 0x82), (0x0F, 0x5F, 0xDC), i / 3)
        shape(MSO_SHAPE.OVAL, 0.90, ry, 0.48, 0.48, fill=col)
        glyph(icon, 1.14, ry + 0.24, 0.26, WHITE)
        textbox(1.54, ry - 0.02, 4.6, 0.28, [(role, 13, True, BLACK)])
        textbox(1.54, ry + 0.28, 4.6, 0.50, [(desc, 10, False, GREY_1)], line_spacing=1.1)

    card(6.83, 1.82, 5.85, 4.55, fill=ICE_BLUE, line=PURPLE)
    tag(7.08, 2.02, 3.5, 0.36, 'HIGH-LEVEL TIMELINE', fill=BLUE)
    STAGES = ['Pre-DD Prep', 'Workshop &\nShadowing', 'Feasibility &\nArchetyping',
              'Business Case', 'Playback &\nSign-off']
    n = len(STAGES)
    TW, TGAP = 0.92, 0.10
    total = n * TW + (n - 1) * TGAP
    TX0 = 7.08 + (5.35 - total) / 2
    TY_ = 3.10
    for i, name in enumerate(STAGES):
        x = TX0 + i * (TW + TGAP)
        col = lerp((0x41, 0x14, 0x82), (0x0F, 0x5F, 0xDC), i / (n - 1))
        shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, TY_, TW, 0.60, fill=col, adj=0.16)
        textbox(x, TY_ + 0.16, TW, 0.28, [('Wk ' + str(i + 1), 11, True, WHITE)],
                align=PP_ALIGN.CENTER)
        for j, ln in enumerate(name.split('\n')):
            textbox(x - 0.05, TY_ + 0.72 + j * 0.24, TW + 0.10, 0.24,
                    [(ln, 8.5, False, BLACK)], align=PP_ALIGN.CENTER)
    rule(TX0, TY_ + 0.30, total, WHITE, 1.0)

    textbox(7.08, 4.70, 5.35, 0.30,
            [('~5 weeks, prep to sign-off — for one L4 process', 11, True, BLUE)])
    textbox(7.08, 5.06, 5.35, 1.10,
            [('Typical for a single process run in isolation. Scale-up '
              'portfolios (multiple processes, multiple workshops) usually '
              'run stages 2–4 in parallel across processes rather than '
              'serially, and compress toward 3–4 weeks per wave.',
              9.5, False, GREY_1)], line_spacing=1.2)
    footer(4, 'Who’s in the room')


# ═══════════════════════════════════════════ 5 · BEFORE YOU WALK IN ═════════
def slide_prep():
    new_slide()
    header('Section 3', [('Prerequisites: ', BLACK), ('what to gather first', BLUE)],
           'Request these eight artifacts before the workshop — arriving '
           'without them wastes the room’s time.')
    ITEMS = [
        ('doc', 'SOPs', 'Latest version, step-by-step, business rules'),
        ('nodes', 'L4 process maps', 'Granular steps, system handoffs, decisions'),
        ('chart_up', 'Volume & seasonality', '12 months, incl. peak intra-day'),
        ('target', 'AHT data', 'Mean, median, 90th percentile per sub-process'),
        ('cubes', 'App & systems inventory', 'Names, versions, hosting type'),
        ('union', 'Process variations', 'Country, entity, currency, language'),
        ('shield', 'Error & exception logs', 'Top 10 reasons, rework loops'),
        ('gear', 'Security mandates', 'PII, PCI-DSS, GDPR, cross-border limits'),
    ]
    CW, CH, GAP = 2.90, 1.86, 0.20
    for i, (icon, title, desc) in enumerate(ITEMS):
        col_i, row_i = i % 4, i // 4
        x = 0.65 + col_i * (CW + GAP)
        y = 1.78 + row_i * (CH + GAP)
        col = PURPLE if row_i == 0 else BLUE
        card(x, y, CW, CH)
        shape(MSO_SHAPE.OVAL, x + 0.22, y + 0.20, 0.52, 0.52, fill=GREY_4)
        glyph(icon, x + 0.48, y + 0.46, 0.28, col)
        textbox(x + 0.20, y + 0.84, CW - 0.40, 0.34, [(title, 12, True, BLACK)])
        textbox(x + 0.20, y + 1.18, CW - 0.40, 0.60, [(desc, 9.5, False, GREY_1)],
                line_spacing=1.1)
    footer(5, 'Preparing for the workshop')


# ═══════════════════════════ 6 · FACILITATION GUIDE — PART 1 ════════════════
def facilitation_slide(n_slide, title_a, title_b, blocks_a, blocks_b, kicker, sub):
    new_slide()
    header(kicker, [(title_a, BLACK), (title_b, BLUE)], sub)
    for col_i, (heading, col_col, qs, tip) in enumerate([blocks_a, blocks_b]):
        x = 0.65 + col_i * 6.19
        card(x, 1.78, 5.95, 4.75)
        tag(x + 0.24, 1.98, 3.6, 0.36, heading, fill=col_col)
        for j, q in enumerate(qs):
            qy = 2.55 + j * 0.72
            shape(MSO_SHAPE.OVAL, x + 0.24, qy, 0.30, 0.30, fill=GREY_4)
            textbox(x + 0.24, qy, 0.30, 0.30, [('Q', 10, True, col_col)],
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            textbox(x + 0.68, qy - 0.06, 5.05, 0.60, [(q, 11, False, BLACK)],
                    line_spacing=1.1)
        ty = 2.55 + len(qs) * 0.72 + 0.10
        shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + 0.24, ty, 5.47,
              4.75 - (ty - 1.78) - 0.24, fill=GREY_4, adj=0.10)
        textbox(x + 0.42, ty + 0.10, 5.15, 0.22, [('LISTEN FOR', 8.5, True, col_col)])
        textbox(x + 0.42, ty + 0.32, 5.15, 1.6, [(tip, 10, False, BLACK)],
                line_spacing=1.15)
    footer(n_slide, 'Running the workshop')


def slide_facilitation_impact():
    facilitation_slide(
        6, 'Ask this: ', 'productivity & business outcome',
        ('PRODUCTIVITY', PURPLE,
         ['What’s the current KPI — cycle time, cost per transaction — '
          'versus internal target or market benchmark?',
          'How many FTEs or hours per week does this absorb today?',
          'If the KPI improves, does that capacity get redeployed — '
          'or just sit idle?'],
         'A benchmark nobody has actually measured against, or “we’d just '
         'have more free time,” signals a productivity case that won’t '
         'survive finance scrutiny.'),
        ('BUSINESS OUTCOME', BLUE,
         ['If this KPI improves by X%, does it move revenue, margin or '
          'compliance risk — or does it stay internal?',
          'What does it cost today when this goes wrong — a late fee, '
          'a lost discount, an audit finding?',
          'Who outside the team actually feels this — customer, '
          'finance, or a regulator?'],
         'If nobody can trace the KPI to a P&L line or a named external '
         'stakeholder, the business-outcome score should stay low, however '
         'painful the process feels internally.'),
        'Section 3', 'Facilitation guide — these questions feed the impact '
        'score directly: Section 4 asks you to weigh exactly what these '
        'answers give you.')



def slide_facilitation_1():
    facilitation_slide(
        7, 'Ask this: ', 'process & data',
        ('PROCESS & OPERATING LOGIC', PURPLE,
         ['What triggers the process — batch, inbox, webhook, manual push?',
          'What % follows the happy path vs. exceptions?',
          'Is it subject to month-end or seasonal spikes (e.g. 5x volume)?'],
         'A trigger nobody can name cleanly, or a happy-path % given with '
         'false confidence, means the SOP and reality have drifted apart.'),
        ('DATA & INPUT CHARACTERISTICS', BLUE,
         ['What format do inputs arrive in — structured, semi-, unstructured?',
          'If scanned: resolution, image quality, handwriting incidence?',
          'Are templates standardized, or do they vary by vendor/region?'],
         'Vague answers on input format almost always mean nobody has '
         'actually opened a sample of 20 real transactions recently.'),
        'Section 3', 'Facilitation guide, part 1 — ask a handful of anchor '
        'questions per category, then dig where the answer is soft.')


def slide_facilitation_2():
    facilitation_slide(
        8, 'Ask this: ', 'systems & exceptions',
        ('SYSTEMS, ARCHITECTURE & ACCESS', PURPLE,
         ['Native access, or Citrix/VDI? Are modern APIs available?',
          'Any migration, major upgrade or decommission in 6–12 months?',
          'What authentication exists — SSO, MFA, RSA tokens, CAPTCHA?'],
         'A “yes” to APIs that turns out to mean “someone mentioned it once” '
         'is the most common false-positive in a DD workshop.'),
        ('EXCEPTIONS, QUALITY & FALLBACK', BLUE,
         ['What are the standard business and technical exceptions?',
          'How are exceptions routed today — dedicated tier-2 queue?',
          'What’s the acceptable human-in-the-loop threshold?'],
         'If nobody can quote an exception rate — only “not that many” — '
         'assume it’s higher than stated until you’ve seen the log.'),
        'Section 3', 'Facilitation guide, part 2 — this is where feasibility '
        'gets decided, not the process walkthrough.')



# ═══════════════════════════ 9 · SCORING — COMPLEXITY (5→3) ═════════════════
def score_bar(x, y, w, col, filled_frac, label_lo, label_hi):
    shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, 0.16, fill=GREY_4, adj=0.5)
    shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w * filled_frac, 0.16, fill=col, adj=0.5)
    textbox(x, y + 0.20, w * 0.6, 0.20, [(label_lo, 8.5, False, GREY_1)])
    textbox(x + w * 0.4, y + 0.20, w * 0.6, 0.20, [(label_hi, 8.5, False, GREY_1)],
            align=PP_ALIGN.RIGHT)


def slide_complexity():
    new_slide()
    header('Section 4', [('Scoring complexity — ', BLACK), ('three dimensions', BLUE)],
           'The source model uses five; in the field, three is enough to '
           'route a decision and easy enough to score live.')
    DIMS = [
        ('35%', 'Data & Input Complexity',
         'Structured digital (API, DB, fixed CSV)',
         'Unstructured — scans, handwriting, free text',
         0.30),
        ('35%', 'Process Logic Complexity',
         'Deterministic rules, <5% exceptions',
         'Subjective judgment, >15% exceptions',
         0.55),
        ('30%', 'Technical & System Complexity',
         '1–2 modern apps, API access, stable',
         '5+ apps, Citrix/VDI/MFA, migration planned',
         0.70),
    ]
    CT, CH, GAP = 1.86, 1.42, 0.22
    for i, (wt, name, lo, hi, frac) in enumerate(DIMS):
        y = CT + i * (CH + GAP)
        col = lerp((0x41, 0x14, 0x82), (0x0F, 0x5F, 0xDC), i / 2)
        card(0.65, y, 12.03, CH)
        tag(0.90, y + 0.22, 0.85, 0.36, wt, fill=col, size=11)
        textbox(1.90, y + 0.20, 5.0, 0.40, [(name, 14, True, BLACK)])
        score_bar(1.90, y + 0.86, 9.4, col, frac, lo, hi)
    textbox(0.65, CT + 3 * CH + 2 * GAP + 0.14, 12.03, 0.3,
            [[('Weighted Complexity Score  ', 11, True, BLACK),
              ('=  Σ (dimension score × weight), on a 1.0–5.0 scale',
               11, False, GREY_1)]])
    footer(9, 'Scoring what you heard')


# ═══════════════════════════════ 10 · SCORING — IMPACT (5→2) ════════════════
def slide_impact():
    new_slide()
    header('Section 4', [('Scoring impact — ', BLACK), ('two dimensions', BLUE)],
           'Everything the source model measures rolls up into one of two '
           'questions: does it free up people, or does it move the P&L?')
    CW, CT, CH = 5.75, 1.86, 4.55
    for i, (title, col, tint, items, wt) in enumerate([
            ('Productivity', PURPLE, ICE_BLUE,
             ['FTE capacity released', 'Cycle time / SLA reduction',
              'Rework and error elimination'], '50%'),
            ('Business Outcome', BLUE, ICE_BLUE,
             ['Revenue protected or accelerated', 'Margin improvement',
              'Cash flow / working capital, compliance risk avoided'], '50%')]):
        x = 0.65 + i * (CW + 0.53)
        card(x, CT, CW, CH, fill=tint, line=col)
        tag(x + 0.30, CT + 0.26, 2.9, 0.42, title.upper(), fill=col, size=12)
        textbox(x + CW - 1.1, CT + 0.26, 0.8, 0.42, [(wt, 16, True, col)],
                align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        for j, it in enumerate(items):
            iy = CT + 0.98 + j * 0.62
            shape(MSO_SHAPE.OVAL, x + 0.30, iy + 0.05, 0.16, 0.16, fill=col)
            textbox(x + 0.60, iy - 0.06, CW - 0.90, 0.56, [(it, 12, False, BLACK)],
                    anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
        rule(x + 0.30, CT + CH - 0.66, CW - 0.60, col)
        low = '< 0.5 FTE, no P&L line' if i == 0 else 'No direct revenue / margin line'
        high = '> 3 FTEs released' if i == 0 else 'Direct revenue, margin or risk line'
        score_bar(x + 0.30, CT + CH - 0.50, CW - 0.60, col, 0.6 if i == 0 else 0.75,
                  low, high)
    footer(10, 'Scoring what you heard')


# ═══════════════════════════════ 11 · PRIORITIZATION QUADRANT ═══════════════
def slide_quadrant():
    new_slide()
    header('Section 4', [('Where it lands: ', BLACK), ('the prioritization quadrant', BLUE)],
           'Plot the two scores — impact on the vertical axis, complexity '
           'on the horizontal.')
    GX, GY, GW, GH = 2.10, 1.90, 8.6, 4.6
    shape(MSO_SHAPE.RECTANGLE, GX, GY, GW / 2, GH / 2, fill=ICE_PURPLE)
    shape(MSO_SHAPE.RECTANGLE, GX + GW / 2, GY, GW / 2, GH / 2, fill=ICE_BLUE)
    shape(MSO_SHAPE.RECTANGLE, GX, GY + GH / 2, GW / 2, GH / 2, fill=GREY_4)
    shape(MSO_SHAPE.RECTANGLE, GX + GW / 2, GY + GH / 2, GW / 2, GH / 2,
          fill=GREY_3)
    card(GX, GY, GW, GH, fill=None, line=GREY_2)
    vrule(GX + GW / 2, GY, GH, GREY_2, 1.0)
    rule(GX, GY + GH / 2, GW, GREY_2, 1.0)

    QUADS = [
        (GX + 0.14, GY + 0.14, 'QUICK WINS', PURPLE_DEEP, 'Priority Wave 1 · high ROI, fast delivery'),
        (GX + GW / 2 + 0.14, GY + 0.14, 'STRATEGIC BETS', PURPLE, 'Priority Wave 2 · high value, needs AI/architecture'),
        (GX + 0.14, GY + GH / 2 + 0.14, 'SCALE FILLERS', GREY_1, 'Priority Wave 3 / backlog · low effort, minor impact'),
        (GX + GW / 2 + 0.14, GY + GH / 2 + 0.14, 'DEPRIORITIZE / RE-ENGINEER',
         BLACK, 'Do NOT automate as-is — fix the process first'),
    ]
    for qx, qy, lab, col, sub in QUADS:
        textbox(qx, qy, 4.0, 0.26, [(lab, 12, True, col)])
        textbox(qx, qy + 0.30, 4.0, 0.42, [(sub, 9, False, GREY_1)], line_spacing=1.1)

    vrule(GX - 0.30, GY + 0.16, GH - 0.30, GREY_2, 1.25)
    arrowhead(GX - 0.30, GY + 0.10, GREY_2, 0.075, 'up')
    textbox(GX - 0.75, GY - 0.06, 1.05, GH, [('IMPACT', 10, True, GREY_1)],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    textbox(GX, GY + GH + 0.10, GW, 0.24, [('COMPLEXITY  →', 10, True, GREY_1)],
            align=PP_ALIGN.CENTER)
    footer(11, 'Scoring what you heard')


# ═══════════════════════════ 12 · TECHNOLOGY DECISION TREE (5 archetypes) ═══
def slide_tech():
    new_slide()
    header('Section 5', [('Choosing the ', BLACK), ('technology', BLUE)],
           'Five archetypes cover almost everything you’ll see. Route the '
           'candidate to the lever that matches how it actually behaves.')
    ROWS = [
        ('gear', 'RPA', PURPLE, 'Structured data, deterministic rules, repetitive UI keystrokes',
         'Order entry, data migration, screen-to-screen re-keying'),
        ('doc', 'IDP', BLUE, 'Semi-structured documents that need extraction',
         'Invoices, remittances, claims forms, IDs'),
        ('bulb', 'AI-Enabled Automation', PURPLE,
         'Classification, matching or prediction — with human oversight',
         'Cash application matching, fraud scoring, demand forecasting'),
        ('union', 'Agentic', BLUE,
         'Multi-step reasoning and dynamic decisioning across systems',
         'Exception triage & resolution, dispute investigation'),
        ('layers', 'BPM', PURPLE,
         'Cross-functional workflow with human handoffs and approvals',
         'Case management, approval routing, onboarding'),
    ]
    TY = 1.78
    hx = 0.65
    W1, W2, W3 = 2.85, 4.30, 4.88
    shape(MSO_SHAPE.RECTANGLE, hx, TY, 12.03, 0.40, fill=PURPLE_DEEP)
    for cx, cw, lab in ((hx, W1, 'ARCHETYPE'), (hx + W1, W2, 'WHEN TO USE IT'),
                        (hx + W1 + W2, W3, 'TYPICAL EXAMPLE')):
        textbox(cx + 0.16, TY, cw - 0.2, 0.40, [(lab, 9.5, True, WHITE)],
                anchor=MSO_ANCHOR.MIDDLE)
    RH = 0.80
    for i, (icon, name, col, when, ex) in enumerate(ROWS):
        ry = TY + 0.40 + i * RH
        if i % 2 == 0:
            shape(MSO_SHAPE.RECTANGLE, hx, ry, 12.03, RH, fill=GREY_4)
        shape(MSO_SHAPE.OVAL, hx + 0.16, ry + RH / 2 - 0.19, 0.38, 0.38, fill=col)
        glyph(icon, hx + 0.35, ry + RH / 2, 0.20, WHITE)
        textbox(hx + 0.64, ry, W1 - 0.70, RH, [(name, 12.5, True, BLACK)],
                anchor=MSO_ANCHOR.MIDDLE)
        textbox(hx + W1 + 0.10, ry, W2 - 0.2, RH, [(when, 10, False, BLACK)],
                anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
        textbox(hx + W1 + W2 + 0.10, ry, W3 - 0.2, RH, [(ex, 10, False, GREY_1)],
                anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)

    fy = TY + 0.40 + len(ROWS) * RH + 0.16
    shape(MSO_SHAPE.ROUNDED_RECTANGLE, hx, fy, 12.03, 0.52, fill=ICE_PURPLE,
          line=PURPLE, line_w=1.0, adj=0.2)
    shape(MSO_SHAPE.OVAL, hx + 0.16, fy + 0.10, 0.32, 0.32, fill=PURPLE_DEEP)
    glyph('target', hx + 0.32, fy + 0.26, 0.18, WHITE)
    textbox(hx + 0.62, fy, 11.2, 0.52,
            [[('Process broken beyond repair? ', 11, True, PURPLE_DEEP),
              ('Fix it first. None of the five archetypes rescue a process '
               'nobody can explain consistently.', 11, False, BLACK)]],
            anchor=MSO_ANCHOR.MIDDLE)
    footer(12, 'Choosing the technology')


# ═══════════════════════════════ 13 · BUSINESS CASE MECHANICS ═══════════════
def slide_mechanics():
    new_slide()
    header('Section 6', [('Business case ', BLACK), ('mechanics', BLUE)],
           'Three benefit streams, weighed against total cost of ownership — '
           'in terms you can explain without a whiteboard.')
    card(0.65, 1.82, 12.03, 2.30)
    tag(0.90, 2.02, 2.6, 0.36, 'BENEFIT FORMULA', fill=PURPLE)
    textbox(0.90, 2.55, 11.5, 0.60,
            [[('Annual FTE Savings  ', 13, True, BLACK),
              ('=  (Volume × AHT ÷ productive hours per FTE)  ×  '
               'Automation Coverage %  ×  Fully Loaded FTE Cost', 12, False, GREY_1)]],
            line_spacing=1.2)
    textbox(0.90, 3.15, 11.5, 0.40,
            [[('Error / Rework Savings  ', 12, True, BLACK),
              ('=  Annual Error Incidents × Average Cost per Correction',
               11.5, False, GREY_1)]])
    textbox(0.90, 3.55, 11.5, 0.44,
            [[('Working Capital / SLA Gains  ', 12, True, BLACK),
              ('=  Early-payment discounts captured + SLA penalties avoided',
               11.5, False, GREY_1)]])

    CT2, CH2, CW2 = 4.34, 1.86, 5.90
    card(0.65, CT2, CW2, CH2, fill=ICE_BLUE, line=PURPLE)
    tag(0.90, CT2 + 0.20, 2.9, 0.36, 'ONE-TIME BUILD COST', fill=PURPLE)
    for i, it in enumerate(['Discovery & solution architecture', 'Build & test',
                            'Change management', 'Deployment']):
        iy = CT2 + 0.70 + i * 0.28
        textbox(0.90, iy, 5.4, 0.26, [('•  ' + it, 10.5, False, BLACK)])

    card(6.78, CT2, CW2, CH2, fill=ICE_BLUE, line=BLUE)
    tag(7.03, CT2 + 0.20, 2.9, 0.36, 'ANNUAL RUN COST', fill=BLUE)
    for i, it in enumerate(['Platform licensing (bot, IDP credits, LLM tokens)',
                            'Cloud infrastructure',
                            'L2/L3 support & maintenance',
                            '~ typically 15–20% of build cost, per year']):
        iy = CT2 + 0.70 + i * 0.28
        textbox(7.03, iy, 5.4, 0.26, [('•  ' + it, 10.5, False, BLACK)])
    footer(13, 'Building the business case')


# ═══════════════ 14 · WORKED EXAMPLE — O2C CASH APPLICATION: THE PROCESS ════
def wx_header(n, title_b, sub):
    header('Section 7 · Worked example', [('Order-to-Cash: ', BLACK), (title_b, BLUE)], sub)
    footer(n, 'Worked example — Order-to-Cash')


def slide_wx_process():
    new_slide()
    wx_header(14, 'cash application',
              'Order-to-Cash spans order, credit, fulfillment, billing, '
              'collections and cash application — we zoom into one L4 '
              'activity: matching incoming payments to open invoices.')
    STEPS = [('doc', 'Remittance received', 'Bank lockbox file or emailed advice'),
             ('search', 'Match to open invoices', 'Invoice #, PO #, amount, customer'),
             ('chart_up', 'Apply cash', 'Post matched payments in SAP'),
             ('target', 'Flag exceptions', 'Short pay, unapplied, disputed'),
             ('union', 'Route exception', 'To collections or deductions team')]
    n = len(STEPS)
    W, GAP = 2.14, 0.28
    total = n * W + (n - 1) * GAP
    X0 = (13.333 - total) / 2
    Y, H = 2.20, 2.05
    for i, (icon, title, desc) in enumerate(STEPS):
        x = X0 + i * (W + GAP)
        col = lerp((0x41, 0x14, 0x82), (0x0F, 0x5F, 0xDC), i / (n - 1))
        card(x, Y, W, H)
        shape(MSO_SHAPE.OVAL, x + W / 2 - 0.30, Y + 0.22, 0.60, 0.60, fill=col)
        glyph(icon, x + W / 2, Y + 0.52, 0.30, WHITE)
        textbox(x + 0.12, Y + 0.94, W - 0.24, 0.50, [(title, 11.5, True, BLACK)],
                align=PP_ALIGN.CENTER, line_spacing=1.0)
        textbox(x + 0.12, Y + 1.44, W - 0.24, 0.56, [(desc, 9, False, GREY_1)],
                align=PP_ALIGN.CENTER, line_spacing=1.1)
        if i < n - 1:
            arrowhead(x + W + GAP / 2, Y + H / 2, GREY_2, 0.075, 'right')

    STATS = [('9,000 / mo', 'Remittance volume'), ('7 min', 'Manual AHT per item'),
             ('22%', 'Require manual matching'), ('~7 FTE', 'Current cash app team')]
    SY = 4.70
    card(X0, SY, total, 1.0, fill=GREY_4, line=None)
    sw = total / 4
    for i, (val, lab) in enumerate(STATS):
        sx = X0 + i * sw
        textbox(sx, SY + 0.14, sw, 0.42, [(val, 22, True, PURPLE)],
                align=PP_ALIGN.CENTER)
        textbox(sx, SY + 0.60, sw, 0.30, [(lab, 9.5, False, GREY_1)],
                align=PP_ALIGN.CENTER)
        if i:
            vrule(sx, SY + 0.16, 0.68, GREY_3)
    textbox(X0, 5.90, total, 0.4,
            [('Illustrative figures for training purposes — replace with '
              'live workshop data.', 9, False, GREY_2)], align=PP_ALIGN.CENTER)


# ═══════════════════ 15 · WORKED EXAMPLE — SCORED ═══════════════════════════
def mini_bar(x, y, w, col, frac, score_txt):
    shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, 0.20, fill=GREY_4, adj=0.5)
    shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w * frac, 0.20, fill=col, adj=0.5)
    textbox(x + w + 0.10, y - 0.03, 0.5, 0.26, [(score_txt, 11, True, col)])


def slide_wx_scored():
    new_slide()
    wx_header(15, 'scored',
              'Same two frameworks from Section 4, applied to this process.')

    card(0.65, 1.80, 5.85, 4.60)
    tag(0.90, 2.00, 2.9, 0.34, 'COMPLEXITY  ·  WEIGHTED 2.70', fill=PURPLE, size=10)
    CDIMS = [('Data & Input (35%)', 3, PURPLE), ('Process Logic (35%)', 3, PURPLE),
             ('Technical & System (30%)', 2, BLUE)]
    for i, (lab, sc, col) in enumerate(CDIMS):
        iy = 2.56 + i * 0.92
        textbox(0.90, iy, 4.4, 0.26, [(lab, 11.5, True, BLACK)])
        mini_bar(0.90, iy + 0.34, 4.35, col, sc / 5, str(sc))
    textbox(0.90, 5.66, 5.3, 0.60,
            [('Remittance emails + scanned lockbox images; matching rules with '
              '22% needing judgment; SAP + bank portal, stable.', 9.5, False, GREY_1)],
            line_spacing=1.15)

    card(6.83, 1.80, 5.85, 4.60)
    tag(7.08, 2.00, 2.9, 0.34, 'IMPACT  ·  WEIGHTED 4.50', fill=BLUE, size=10)
    IDIMS = [('Productivity (50%)', 5, PURPLE), ('Business Outcome (50%)', 4, BLUE)]
    for i, (lab, sc, col) in enumerate(IDIMS):
        iy = 2.56 + i * 0.92
        textbox(7.08, iy, 4.4, 0.26, [(lab, 11.5, True, BLACK)])
        mini_bar(7.08, iy + 0.34, 4.35, col, sc / 5, str(sc))
    textbox(7.08, 4.42, 5.3, 0.60,
            [('~4 FTE released; faster cash recognition improves DSO and '
              'working capital, though no direct revenue capture.', 9.5, False, GREY_1)],
            line_spacing=1.15)

    tag(7.08, 5.30, 4.3, 0.42, 'PLACEMENT → QUICK WIN', fill=PURPLE_DEEP, size=11)
    textbox(7.08, 5.80, 5.3, 0.46,
            [('High impact, low-medium complexity — Wave 1 candidate.',
              10, False, GREY_1)], line_spacing=1.1)


# ═══════════════════ 16 · WORKED EXAMPLE — ARCHETYPE & ROI ══════════════════
def slide_wx_roi():
    new_slide()
    wx_header(16, 'archetype & ROI',
              'Four levers combine — no single archetype covers the whole '
              'process end to end.')

    ARCH = [('doc', 'IDP', BLUE, 'Extract remittance data from PDFs / scans'),
            ('bulb', 'AI-Enabled', PURPLE, 'Fuzzy-match remittance lines to invoices'),
            ('union', 'Agentic', BLUE, 'Triage the 22% exceptions, recommend action'),
            ('gear', 'RPA', PURPLE, 'Post matched cash into SAP')]
    CW, CT = 2.85, 1.80
    for i, (icon, name, col, desc) in enumerate(ARCH):
        x = 0.65 + i * (CW + 0.18)
        card(x, CT, CW, 1.55)
        shape(MSO_SHAPE.OVAL, x + 0.20, CT + 0.18, 0.46, 0.46, fill=col)
        glyph(icon, x + 0.43, CT + 0.41, 0.24, WHITE)
        textbox(x + 0.20, CT + 0.72, CW - 0.40, 0.26, [(name, 12.5, True, BLACK)])
        textbox(x + 0.20, CT + 0.98, CW - 0.40, 0.52, [(desc, 9, False, GREY_1)],
                line_spacing=1.1)
        if i < 3:
            arrowhead(x + CW + 0.09, CT + 0.77, GREY_2, 0.06, 'right')

    card(0.65, 3.66, 12.03, 2.72, fill=ICE_BLUE, line=PURPLE)
    tag(0.90, 3.86, 3.0, 0.36, 'ILLUSTRATIVE ROI SUMMARY', fill=PURPLE)
    METRICS = [('$95K', 'Build cost'), ('$148K', 'Year-1 net benefit'),
               ('~8 mo', 'Payback from go-live'), ('$590K', '3-yr cumulative net value')]
    mw = 11.5 / 4
    for i, (val, lab) in enumerate(METRICS):
        mx = 0.90 + i * mw
        textbox(mx, 4.36, mw - 0.2, 0.62, [(val, 28, True, PURPLE)])
        textbox(mx, 4.98, mw - 0.2, 0.30, [(lab, 10.5, False, GREY_1)])
        if i:
            vrule(mx - 0.14, 4.40, 0.90, GREY_3)
    rule(0.90, 5.50, 11.5, GREY_3)
    textbox(0.90, 5.62, 11.5, 0.60,
            [('Numbers are illustrative, built from the formulas in Section 6 '
              '(full derivation in the companion workbook) — not a substitute '
              'for the live workshop model.', 9.5, False, GREY_1)], line_spacing=1.15)


# ═══════════════════ 17 · SIGN-OFF & EXEC PRESENTATION ══════════════════════
def slide_signoff():
    new_slide()
    header('Section 8', [('Getting to ', BLACK), ('sign-off', BLUE)],
           'Three internal gates before the room ever sees an executive '
           'deck.')
    card(0.65, 1.82, 5.75, 4.55)
    tag(0.90, 2.02, 3.2, 0.36, 'THREE STAGE GATES', fill=PURPLE)
    GATES = [('Tech & architecture review', 'Lead architect verifies VDI, API and credential access'),
             ('Security & compliance audit', 'Data privacy officer clears PII handling, token usage'),
             ('SME playback', 'Ops validates baseline AHT, volume and exception rates')]
    for i, (t, d) in enumerate(GATES):
        gy = 2.58 + i * 1.20
        chip(MSO_SHAPE.OVAL, 0.90, gy, 0.36, 0.36, [[(str(i + 1), 12, True, WHITE)]],
             fill=PURPLE)
        textbox(1.42, gy - 0.04, 4.8, 0.30, [(t, 12.5, True, BLACK)])
        textbox(1.42, gy + 0.30, 4.8, 0.60, [(d, 10, False, GREY_1)], line_spacing=1.15)

    card(6.83, 1.82, 5.85, 4.55, fill=ICE_BLUE, line=BLUE)
    tag(7.08, 2.02, 3.5, 0.36, 'EXECUTIVE DECK STRUCTURE', fill=BLUE)
    PARTS = ['Executive summary — pipeline size, FTE release, 3-yr value',
             'Current-state bottlenecks — heatmap of pain points',
             'Target-state solution blueprint', 'Opportunity matrix & prioritization',
             'Implementation roadmap & pilot phase', 'Commercial business case & ROI',
             'Governance & operating model']
    for i, t in enumerate(PARTS):
        py = 2.58 + i * 0.56
        textbox(7.08, py, 0.32, 0.30, [(str(i + 1) + '.', 10.5, True, BLUE)])
        textbox(7.40, py, 5.2, 0.50, [(t, 10, False, BLACK)], line_spacing=1.05)
    footer(17, 'Getting to sign-off')


# ═══════════════════════════ 18 · CHEAT SHEET ════════════════════════════════
def slide_cheatsheet():
    new_slide()
    header('Keep this at your desk', [('The one-page ', BLACK), ('cheat sheet', BLUE)],
           'Everything above, compressed for the field.')

    card(0.65, 1.78, 3.85, 4.60)
    tag(0.90, 1.98, 2.2, 0.32, 'LIFECYCLE', fill=PURPLE, size=9)
    for i, t in enumerate(['Pre-DD Prep', 'Workshop & Shadowing',
                           'Feasibility & Archetyping', 'Business Case', 'Playback & Sign-off']):
        ty = 2.44 + i * 0.42
        chip(MSO_SHAPE.OVAL, 0.90, ty, 0.24, 0.24, [[(str(i + 1), 8.5, True, WHITE)]],
             fill=PURPLE)
        textbox(1.24, ty - 0.03, 3.15, 0.30, [(t, 10, False, BLACK)],
                anchor=MSO_ANCHOR.MIDDLE)
    rule(0.90, 4.68, 3.35, GREY_3)
    tag(0.90, 4.80, 2.2, 0.32, 'SCORING', fill=BLUE, size=9)
    for i, t in enumerate(['Complexity: Data · Logic · Systems',
                           'Impact: Productivity · Outcome']):
        textbox(0.90, 5.26 + i * 0.32, 3.35, 0.30, [('•  ' + t, 9.5, False, BLACK)])

    card(4.68, 1.78, 3.85, 4.60)
    tag(4.93, 1.98, 2.5, 0.32, 'TECHNOLOGY ARCHETYPES', fill=PURPLE, size=9)
    for i, (icon, t) in enumerate([('gear', 'RPA'), ('doc', 'IDP'),
                                   ('bulb', 'AI-Enabled Automation'),
                                   ('union', 'Agentic'), ('layers', 'BPM')]):
        ty = 2.46 + i * 0.60
        shape(MSO_SHAPE.OVAL, 4.93, ty, 0.36, 0.36, fill=BLUE)
        glyph(icon, 5.11, ty + 0.18, 0.20, WHITE)
        textbox(5.42, ty, 2.95, 0.36, [(t, 10.5, False, BLACK)],
                anchor=MSO_ANCHOR.MIDDLE)
    rule(4.93, 5.55, 3.35, GREY_3)
    textbox(4.93, 5.68, 3.35, 0.60,
            [('Broken process? Fix it first — don’t automate chaos.',
              9.5, True, PURPLE_DEEP)], line_spacing=1.15)

    card(8.71, 1.78, 3.97, 4.60, fill=ICE_BLUE, line=PURPLE)
    tag(8.96, 1.98, 2.6, 0.32, 'THE FORMULA', fill=PURPLE, size=9)
    textbox(8.96, 2.44, 3.55, 1.3,
            [('FTE Savings =', 10.5, True, BLACK),
             ('(Volume × AHT ÷ productive hrs)', 9.5, False, BLACK),
             ('× Automation Coverage %', 9.5, False, BLACK),
             ('× Fully Loaded FTE Cost', 9.5, False, BLACK)],
            line_spacing=1.35)
    rule(8.96, 3.86, 3.5, GREY_3)
    tag(8.96, 3.98, 2.6, 0.32, 'ASK ABOUT', fill=BLUE, size=9)
    textbox(8.96, 4.36, 3.55, 0.42,
            [[('KPI vs. benchmark ', 10, True, BLUE),
              ('— current state vs. market', 9.5, False, BLACK)]], line_spacing=1.1)
    textbox(8.96, 4.72, 3.55, 0.42,
            [[('Top-line linkage ', 10, True, BLUE),
              ('— does the fix move the P&L?', 9.5, False, BLACK)]], line_spacing=1.1)
    footer(18, 'Cheat sheet')


for fn in (slide_title, slide_agenda, slide_lifecycle, slide_pod_timeline,
           slide_prep, slide_facilitation_impact, slide_facilitation_1,
           slide_facilitation_2, slide_complexity, slide_impact, slide_quadrant,
           slide_tech, slide_mechanics, slide_wx_process, slide_wx_scored,
           slide_wx_roi, slide_signoff, slide_cheatsheet):
    fn()

prs.save(OUT)
print('saved', OUT, '| slides:', len(prs.slides._sldIdLst))
