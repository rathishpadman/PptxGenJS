"""Order-to-Cash worked example — companion calculator for the DD training pack."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, Reference

OUT = '/home/user/PptxGenJS/dd_worked_example_o2c.xlsx'

PURPLE = 'FF5F1EBE'
PURPLE_DEEP = 'FF411482'
BLUE = 'FF0F5FDC'
ICE_PURPLE = 'FFEFEAFB'
ICE_BLUE = 'FFDCE6F0'
GREY_4 = 'FFE6EBF5'
GREY_1 = 'FF82919A' if False else 'FF64748B'
WHITE = 'FFFFFFFF'
BLACK = 'FF000000'

wb = openpyxl.Workbook()
wb.remove(wb.active)

THIN = Side(style='thin', color='FFCBD5E1')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def style_header_row(ws, row, col_start, col_end, fill=PURPLE_DEEP, font_color=WHITE):
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill('solid', fgColor=fill)
        cell.font = Font(name='Calibri', bold=True, color=font_color, size=11)
        cell.alignment = Alignment(vertical='center', wrap_text=True)
        cell.border = BORDER

def title_block(ws, title, subtitle):
    ws['A1'] = title
    ws['A1'].font = Font(name='Calibri', bold=True, size=16, color=PURPLE_DEEP)
    ws['A2'] = subtitle
    ws['A2'].font = Font(name='Calibri', italic=True, size=10, color='FF64748B')

def note(ws, cell, text):
    ws[cell] = text
    ws[cell].font = Font(name='Calibri', italic=True, size=9, color='FF64748B')
    ws[cell].alignment = Alignment(wrap_text=True, vertical='top')

# ═══════════════════════════════════ 1. PROCESS OVERVIEW ════════════════════
ws = wb.create_sheet('1. Process Overview')
ws.sheet_view.showGridLines = False
title_block(ws, 'Order-to-Cash — Cash Application', 'Worked example companion to the Transformation Due Diligence training pack.')
ws['A4'] = ('Order-to-Cash spans order management, credit, fulfillment, billing, collections '
            'and cash application. This workbook zooms into one L4 activity: matching incoming '
            'payments to open invoices.')
ws['A4'].alignment = Alignment(wrap_text=True, vertical='top')
ws.merge_cells('A4:H4')
ws.row_dimensions[4].height = 34

ws['A6'] = 'PROCESS STEPS'
ws['A6'].font = Font(bold=True, color=PURPLE_DEEP, size=11)
headers = ['#', 'Step', 'Description', 'System(s) involved']
for i, h in enumerate(headers):
    ws.cell(row=7, column=1 + i, value=h)
style_header_row(ws, 7, 1, 4)
STEPS = [
    (1, 'Remittance received', 'Bank lockbox file or emailed remittance advice', 'Bank portal, Outlook'),
    (2, 'Match to open invoices', 'Match by invoice #, PO #, amount and customer', 'SAP'),
    (3, 'Apply cash', 'Post matched payments against open items', 'SAP'),
    (4, 'Flag exceptions', 'Short pay, unapplied cash, disputed amount', 'SAP'),
    (5, 'Route exception', 'Send to collections or deductions team for resolution', 'SAP, ServiceNow'),
]
for r, (n, step, desc, sysm) in enumerate(STEPS, start=8):
    ws.cell(row=r, column=1, value=n)
    ws.cell(row=r, column=2, value=step).font = Font(bold=True)
    ws.cell(row=r, column=3, value=desc)
    ws.cell(row=r, column=4, value=sysm)
    for c in range(1, 5):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).alignment = Alignment(vertical='center', wrap_text=True)
    if r % 2 == 0:
        for c in range(1, 5):
            ws.cell(row=r, column=c).fill = PatternFill('solid', fgColor=GREY_4)

ws['A15'] = 'CURRENT-STATE INPUTS'
ws['A15'].font = Font(bold=True, color=PURPLE_DEEP, size=11)
INPUTS = [
    ('Remittances per month', 9000, ''),
    ('Remittances per year', '=B16*12', 'Volume/month x 12'),
    ('Manual AHT per item (minutes)', 7, ''),
    ('% requiring manual matching (exceptions)', 0.22, ''),
    ('Productive hours per FTE per year', 1750, 'Standard assumption — adjust to your org'),
    ('Current team FTE-equivalent', '=B17*B18/60/B20', '(Annual volume x AHT in hrs) / productive hrs'),
]
for i, (label, val, cmt) in enumerate(INPUTS):
    r = 16 + i
    ws.cell(row=r, column=1, value=label)
    cell = ws.cell(row=r, column=2, value=val)
    if label.startswith('% '):
        cell.number_format = '0%'
    if label == 'Current team FTE-equivalent':
        cell.font = Font(bold=True, color=PURPLE_DEEP)
        cell.number_format = '0.0'
    ws.cell(row=r, column=3, value=cmt).font = Font(italic=True, size=9, color='FF64748B')

note(ws, 'A23', 'All figures on this sheet are illustrative, built for training — replace with '
               'live workshop data before using this as a real business case.')
ws.merge_cells('A23:H23')

widths = {'A': 34, 'B': 16, 'C': 40, 'D': 22}
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'A8'

# ═══════════════════════════════════ 2. COMPLEXITY SCORING ══════════════════
ws = wb.create_sheet('2. Complexity Scoring')
ws.sheet_view.showGridLines = False
title_block(ws, 'Complexity Scoring', 'Three dimensions (simplified from the source model’s five) — 1 (low) to 5 (high).')

headers = ['Dimension', 'Weight', 'Low (1)', 'High (5)', 'Score (1-5)', 'Weighted', 'Evidence from the workshop']
for i, h in enumerate(headers):
    ws.cell(row=4, column=1 + i, value=h)
style_header_row(ws, 4, 1, 7)

DIMS = [
    ('Data & Input Complexity', 0.35, 'Structured digital (API, DB, fixed CSV)',
     'Unstructured — scans, handwriting, free text', 3,
     'Remittances arrive as emails with PDF/scanned lockbox attachments; format varies by bank.'),
    ('Process Logic Complexity', 0.35, 'Deterministic rules, <5% exceptions',
     'Subjective judgment, >15% exceptions', 3,
     'Matching rules are mostly deterministic, but 22% of volume needs manual judgment.'),
    ('Technical & System Complexity', 0.30, '1-2 modern apps, API access, stable',
     '5+ apps, Citrix/VDI/MFA, migration planned', 2,
     'SAP + bank portal only; both stable, no migration planned in the next 12 months.'),
]
for i, (name, wt, lo, hi, score, evid) in enumerate(DIMS):
    r = 5 + i
    ws.cell(row=r, column=1, value=name).font = Font(bold=True)
    wcell = ws.cell(row=r, column=2, value=wt); wcell.number_format = '0%'
    ws.cell(row=r, column=3, value=lo)
    ws.cell(row=r, column=4, value=hi)
    scell = ws.cell(row=r, column=5, value=score)
    scell.font = Font(bold=True, color=PURPLE_DEEP)
    scell.alignment = Alignment(horizontal='center')
    wtd = ws.cell(row=r, column=6, value='=B%d*E%d' % (r, r))
    wtd.number_format = '0.00'
    ws.cell(row=r, column=7, value=evid).alignment = Alignment(wrap_text=True, vertical='center')
    for c in range(1, 8):
        ws.cell(row=r, column=c).border = BORDER
        if c != 7:
            ws.cell(row=r, column=c).alignment = Alignment(vertical='center', wrap_text=True)

    # data validation-style manual score entry note
ws.row_dimensions[5].height = 30
ws.row_dimensions[6].height = 30
ws.row_dimensions[7].height = 30

r_total = 5 + len(DIMS)
ws.cell(row=r_total, column=1, value='Weighted Complexity Score').font = Font(bold=True, size=12)
ws.merge_cells(start_row=r_total, start_column=1, end_row=r_total, end_column=5)
tot = ws.cell(row=r_total, column=6, value='=SUM(F5:F7)')
tot.font = Font(bold=True, size=14, color=PURPLE_DEEP)
tot.number_format = '0.00'
tot.fill = PatternFill('solid', fgColor=ICE_PURPLE)
for c in range(1, 7):
    ws.cell(row=r_total, column=c).border = BORDER
    ws.cell(row=r_total, column=c).fill = PatternFill('solid', fgColor=ICE_PURPLE)

note(ws, 'A%d' % (r_total + 2),
     'Change any Score (1-5) in column E to re-run the model — the weighted total, the '
     'prioritization sheet, and the quadrant placement all recalculate automatically.')
ws.merge_cells(start_row=r_total + 2, start_column=1, end_row=r_total + 2, end_column=7)

widths = {'A': 26, 'B': 9, 'C': 30, 'D': 30, 'E': 11, 'F': 11, 'G': 42}
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'A5'

COMPLEXITY_TOTAL_REF = "'2. Complexity Scoring'!F%d" % r_total

# ═══════════════════════════════════ 3. IMPACT SCORING ══════════════════════
ws = wb.create_sheet('3. Impact Scoring')
ws.sheet_view.showGridLines = False
title_block(ws, 'Impact Scoring', 'Two dimensions (simplified from the source model’s five) — 1 (low) to 5 (high).')

headers = ['Dimension', 'Weight', 'Low (1)', 'High (5)', 'Score (1-5)', 'Weighted', 'Evidence from the workshop']
for i, h in enumerate(headers):
    ws.cell(row=4, column=1 + i, value=h)
style_header_row(ws, 4, 1, 7, fill=BLUE)

IDIMS = [
    ('Productivity', 0.50, '< 0.5 FTE released, no P&L line', '> 3.0 FTEs released', 5,
     '~4.0 FTEs released (see Sheet 6) — well above the high-impact threshold.'),
    ('Business Outcome', 0.50, 'No direct revenue / margin line', 'Direct revenue, margin or risk line', 4,
     'Faster cash recognition improves DSO and working capital; no direct revenue capture.'),
]
for i, (name, wt, lo, hi, score, evid) in enumerate(IDIMS):
    r = 5 + i
    ws.cell(row=r, column=1, value=name).font = Font(bold=True)
    wcell = ws.cell(row=r, column=2, value=wt); wcell.number_format = '0%'
    ws.cell(row=r, column=3, value=lo)
    ws.cell(row=r, column=4, value=hi)
    scell = ws.cell(row=r, column=5, value=score)
    scell.font = Font(bold=True, color=BLUE)
    scell.alignment = Alignment(horizontal='center')
    wtd = ws.cell(row=r, column=6, value='=B%d*E%d' % (r, r))
    wtd.number_format = '0.00'
    ws.cell(row=r, column=7, value=evid).alignment = Alignment(wrap_text=True, vertical='center')
    for c in range(1, 8):
        ws.cell(row=r, column=c).border = BORDER
        if c != 7:
            ws.cell(row=r, column=c).alignment = Alignment(vertical='center', wrap_text=True)
ws.row_dimensions[5].height = 30
ws.row_dimensions[6].height = 30

r_total = 5 + len(IDIMS)
ws.cell(row=r_total, column=1, value='Weighted Impact Score').font = Font(bold=True, size=12)
ws.merge_cells(start_row=r_total, start_column=1, end_row=r_total, end_column=5)
tot = ws.cell(row=r_total, column=6, value='=SUM(F5:F6)')
tot.font = Font(bold=True, size=14, color=BLUE)
tot.number_format = '0.00'
tot.fill = PatternFill('solid', fgColor=ICE_BLUE)
for c in range(1, 7):
    ws.cell(row=r_total, column=c).border = BORDER
    ws.cell(row=r_total, column=c).fill = PatternFill('solid', fgColor=ICE_BLUE)

note(ws, 'A%d' % (r_total + 2),
     'Productivity = FTE capacity released, cycle time / SLA reduction, rework and error '
     'elimination. Business Outcome = revenue, margin, cash flow / working capital, '
     'compliance risk avoided. Change any Score to re-run the model.')
ws.merge_cells(start_row=r_total + 2, start_column=1, end_row=r_total + 2, end_column=7)
ws.row_dimensions[r_total + 2].height = 30

widths = {'A': 22, 'B': 9, 'C': 30, 'D': 30, 'E': 11, 'F': 11, 'G': 42}
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'A5'

IMPACT_TOTAL_REF = "'3. Impact Scoring'!F%d" % r_total

# ═══════════════════════════════════ 4. PRIORITIZATION ══════════════════════
ws = wb.create_sheet('4. Prioritization')
ws.sheet_view.showGridLines = False
title_block(ws, 'How the Priority Was Arrived At', 'Fully transparent — change nothing here; it recalculates from Sheets 2 and 3.')

ws['A4'] = 'STEP 1 — PULL THE TWO SCORES'
ws['A4'].font = Font(bold=True, color=PURPLE_DEEP, size=11)
headers = ['Score', 'Value', 'Source']
for i, h in enumerate(headers):
    ws.cell(row=5, column=1 + i, value=h)
style_header_row(ws, 5, 1, 3)

ws.cell(row=6, column=1, value='Complexity Score').font = Font(bold=True)
c_val = ws.cell(row=6, column=2, value='=%s' % COMPLEXITY_TOTAL_REF)
c_val.number_format = '0.00'; c_val.font = Font(bold=True, color=PURPLE_DEEP, size=12)
ws.cell(row=6, column=3, value='Sheet 2, Weighted Complexity Score')

ws.cell(row=7, column=1, value='Impact Score').font = Font(bold=True)
i_val = ws.cell(row=7, column=2, value='=%s' % IMPACT_TOTAL_REF)
i_val.number_format = '0.00'; i_val.font = Font(bold=True, color=BLUE, size=12)
ws.cell(row=7, column=3, value='Sheet 3, Weighted Impact Score')
for r in (6, 7):
    for c in range(1, 4):
        ws.cell(row=r, column=c).border = BORDER

ws['A10'] = 'STEP 2 — APPLY THE THRESHOLD (midpoint of the 1-5 scale)'
ws['A10'].font = Font(bold=True, color=PURPLE_DEEP, size=11)
headers2 = ['Axis', 'Rule', 'This process', 'Result']
for i, h in enumerate(headers2):
    ws.cell(row=11, column=1 + i, value=h)
style_header_row(ws, 11, 1, 4)

ws.cell(row=12, column=1, value='Complexity').font = Font(bold=True)
ws.cell(row=12, column=2, value='<= 3.0 -> Low       |       > 3.0 -> High')
ws.cell(row=12, column=3, value='=B6').number_format = '0.00'
r12 = ws.cell(row=12, column=4, value='=IF(B6<=3,"Low Complexity","High Complexity")')
r12.font = Font(bold=True)

ws.cell(row=13, column=1, value='Impact').font = Font(bold=True)
ws.cell(row=13, column=2, value='>= 3.0 -> High       |       < 3.0 -> Low')
ws.cell(row=13, column=3, value='=B7').number_format = '0.00'
r13 = ws.cell(row=13, column=4, value='=IF(B7>=3,"High Impact","Low Impact")')
r13.font = Font(bold=True)
for r in (12, 13):
    for c in range(1, 5):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).alignment = Alignment(vertical='center', wrap_text=True)

ws['A16'] = 'STEP 3 — COMBINE INTO A PRIORITY WAVE'
ws['A16'].font = Font(bold=True, color=PURPLE_DEEP, size=11)
formula = ('=IF(AND(B7>=3,B6<=3),"Quick Win",'
           'IF(AND(B7>=3,B6>3),"Strategic Bet",'
           'IF(AND(B7<3,B6<=3),"Scale Filler","Deprioritize / Re-engineer")))')
ws['A17'] = 'This process is a:'
ws['A17'].font = Font(size=12)
result_cell = ws['B17']
result_cell.value = formula
result_cell.font = Font(bold=True, size=18, color=WHITE)
result_cell.fill = PatternFill('solid', fgColor=PURPLE_DEEP)
result_cell.alignment = Alignment(horizontal='center', vertical='center')
ws.merge_cells('B17:D17')
ws.row_dimensions[17].height = 32

ws['A19'] = ('Rule of thumb: High Impact + Low Complexity = Quick Win (do it first). '
             'High Impact + High Complexity = Strategic Bet (worth it, needs architecture). '
             'Low Impact + Low Complexity = Scale Filler (backlog). Low Impact + High '
             'Complexity = Deprioritize — fix the process before automating it.')
ws['A19'].alignment = Alignment(wrap_text=True, vertical='top')
ws.merge_cells('A19:F19')
ws.row_dimensions[19].height = 48

# ---- 2x2 visual grid with the live result highlighted ----------------------
ws['A22'] = 'VISUAL PLACEMENT'
ws['A22'].font = Font(bold=True, color=PURPLE_DEEP, size=11)
grid_labels = {
    (23, 'B'): ('Quick Win', ICE_PURPLE),
    (23, 'D'): ('Strategic Bet', ICE_BLUE),
    (26, 'B'): ('Scale Filler', GREY_4),
    (26, 'D'): ('Deprioritize / Re-engineer', 'FFC8D2DD'),
}
ws.merge_cells('B23:C25')
ws.merge_cells('D23:E25')
ws.merge_cells('B26:C28')
ws.merge_cells('D26:E28')
for (row, col), (label, fill) in grid_labels.items():
    cell = ws['%s%d' % (col, row)]
    cell.value = label
    cell.fill = PatternFill('solid', fgColor=fill)
    cell.font = Font(bold=True, size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws['A24'] = 'High'
ws['A27'] = 'Low'
ws['A24'].font = ws['A27'].font = Font(bold=True, size=9, color='FF64748B')
ws['A21'] = 'Impact'
ws['A21'].font = Font(bold=True, size=9, color='FF64748B')
ws['C29'] = 'Low'
ws['D29'] = '<— Complexity —>'
ws['E29'] = 'High'
for cc in ('C29', 'E29'):
    ws[cc].font = Font(size=9, color='FF64748B')
    ws[cc].alignment = Alignment(horizontal='center')
ws['D29'].font = Font(bold=True, size=9, color='FF64748B')
ws['D29'].alignment = Alignment(horizontal='center')

thick_gold = Side(style='thick', color='FF411482')
highlight_border = Border(left=thick_gold, right=thick_gold, top=thick_gold, bottom=thick_gold)
for (row, col), (label, _) in grid_labels.items():
    rng = '%s%d:%s%d' % (col, row, chr(ord(col) + 1), row + 2)
    ws.conditional_formatting.add(
        rng,
        FormulaRule(formula=['$B$17="%s"' % label], border=highlight_border))

note(ws, 'A31', 'The bold border marks where this process actually lands, driven live by '
                'the B17 formula above.')
ws.merge_cells('A31:F31')

widths = {'A': 20, 'B': 24, 'C': 24, 'D': 24, 'E': 24}
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'A4'

# ═══════════════════════════════════ 5. TECHNOLOGY ARCHETYPE ════════════════
ws = wb.create_sheet('5. Technology Archetype')
ws.sheet_view.showGridLines = False
title_block(ws, 'Technology Archetype', 'Four levers combine — no single archetype covers this process end to end.')

headers = ['Archetype', 'Role in this process', 'Why this archetype']
for i, h in enumerate(headers):
    ws.cell(row=4, column=1 + i, value=h)
style_header_row(ws, 4, 1, 3)

ARCH = [
    ('IDP', 'Extract remittance data from PDFs / scans',
     'Inputs are semi-structured documents (remittance advices, lockbox scans) that need field extraction.'),
    ('AI-Enabled Automation', 'Fuzzy-match remittance lines to open invoices',
     'Matching is mostly rule-based but needs tolerance for imperfect references — a classic ML matching task.'),
    ('Agentic', 'Triage the 22% exceptions and recommend an action',
     'Exceptions need multi-step reasoning (which invoice, which reason code, who to route to) — not a fixed rule.'),
    ('RPA', 'Post matched cash into SAP',
     'Once matched, posting is a deterministic, repetitive UI/API step.'),
]
for i, (name, role, why) in enumerate(ARCH):
    r = 5 + i
    ws.cell(row=r, column=1, value=name).font = Font(bold=True, color=PURPLE_DEEP)
    ws.cell(row=r, column=2, value=role)
    ws.cell(row=r, column=3, value=why)
    for c in range(1, 4):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).alignment = Alignment(vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = 34
    if i % 2 == 0:
        for c in range(1, 4):
            ws.cell(row=r, column=c).fill = PatternFill('solid', fgColor=GREY_4)

note(ws, 'A10', 'Broken process? Fix it first — none of the five archetypes (RPA, IDP, '
                'AI-Enabled Automation, Agentic, BPM) rescue a process nobody can explain '
                'consistently.')
ws.merge_cells('A10:C10')
ws.row_dimensions[10].height = 30

widths = {'A': 24, 'B': 40, 'C': 55}
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'A5'

# ═══════════════════════════════════ 6. ROI CALCULATION ═════════════════════
ws = wb.create_sheet('6. ROI Calculation')
ws.sheet_view.showGridLines = False
title_block(ws, 'ROI Calculation', 'Fully formula-driven — change any input and every downstream figure recalculates, '
                                    'including the headline card on the PPT slide.')

# ---- assumptions -----------------------------------------------------------
ws['A4'] = 'ASSUMPTIONS'
ws['A4'].font = Font(bold=True, color=PURPLE_DEEP, size=11)
headers = ['Input', 'Value', 'Note']
for i, h in enumerate(headers):
    ws.cell(row=5, column=1 + i, value=h)
style_header_row(ws, 5, 1, 3)

ASSUMP = [
    ('Current team FTE-equivalent', "='1. Process Overview'!B21", 'From Sheet 1', '0.0'),
    ('Automation coverage (%)', 0.55, 'Share of current effort the solution automates', '0%'),
    ('FTE released', '=B6*B7', 'Current FTE x coverage', '0.0'),
    ('Fully loaded FTE cost ($/yr)', 58000, '', '$#,##0'),
    ('Annual FTE savings, run-rate ($)', '=B8*B9', 'FTE released x FTE cost', '$#,##0'),
    ('Annual error incidents avoided', 4500, 'Distinct mis-applied-cash cases per year', '#,##0'),
    ('Avg cost per correction ($)', 15, '', '$#,##0.00'),
    ('Error / rework savings, run-rate ($)', '=B11*B12', '', '$#,##0'),
    ('Working capital / SLA gains, run-rate ($)', 12000, 'Early-payment capture + SLA penalty avoidance', '$#,##0'),
    ('Total gross benefit, run-rate ($/yr)', '=B10+B13+B14', 'FTE + error/rework + working capital', '$#,##0'),
]
for i, (label, val, cmt, fmt) in enumerate(ASSUMP):
    r = 6 + i
    ws.cell(row=r, column=1, value=label)
    vcell = ws.cell(row=r, column=2, value=val)
    vcell.number_format = fmt
    ws.cell(row=r, column=3, value=cmt).font = Font(italic=True, size=9, color='FF64748B')
    for c in (1, 2, 3):
        ws.cell(row=r, column=c).border = BORDER
ws.cell(row=15, column=1).font = Font(bold=True)
ws.cell(row=15, column=2).font = Font(bold=True, color=PURPLE_DEEP)
ws.cell(row=15, column=2).fill = PatternFill('solid', fgColor=ICE_PURPLE)

# ---- ramp + year-by-year build-up ------------------------------------------
ws['A18'] = 'YEAR-BY-YEAR BUILD-UP'
ws['A18'].font = Font(bold=True, color=PURPLE_DEEP, size=11)
cols = ['Year 0 (Build)', 'Year 1', 'Year 2', 'Year 3', 'Total']
for i, h in enumerate(cols):
    ws.cell(row=19, column=2 + i, value=h)
style_header_row(ws, 19, 2, 6, fill=BLUE)
ws.cell(row=19, column=1, value='')

ws.cell(row=20, column=1, value='Ramp — % of run-rate realized').font = Font(italic=True)
ramp_vals = [None, 0.55, 0.90, 1.00, None]
for i, v in enumerate(ramp_vals):
    cell = ws.cell(row=20, column=2 + i)
    if v is not None:
        cell.value = v
        cell.number_format = '0%'

ws.cell(row=21, column=1, value='Gross Financial Benefit ($)').font = Font(bold=True)
ws.cell(row=21, column=2, value=0)
ws.cell(row=21, column=3, value='=$B$15*C20')
ws.cell(row=21, column=4, value='=$B$15*D20')
ws.cell(row=21, column=5, value='=$B$15*E20')
ws.cell(row=21, column=6, value='=SUM(B21:E21)')
for c in range(2, 7):
    ws.cell(row=21, column=c).number_format = '$#,##0'

ws.cell(row=22, column=1, value='Implementation / Build Costs ($)')
ws.cell(row=22, column=2, value=-90000)
ws.cell(row=22, column=3, value=0)
ws.cell(row=22, column=4, value=0)
ws.cell(row=22, column=5, value=0)
ws.cell(row=22, column=6, value='=SUM(B22:E22)')
for c in range(2, 7):
    ws.cell(row=22, column=c).number_format = '$#,##0;($#,##0)'

ws.cell(row=23, column=1, value='Software Licensing & Run Costs ($)')
ws.cell(row=23, column=2, value=-5000)
ws.cell(row=23, column=3, value=-22000)
ws.cell(row=23, column=4, value=-24000)
ws.cell(row=23, column=5, value=-26000)
ws.cell(row=23, column=6, value='=SUM(B23:E23)')
for c in range(2, 7):
    ws.cell(row=23, column=c).number_format = '$#,##0;($#,##0)'

ws.cell(row=24, column=1, value='Net Cash Flow ($)').font = Font(bold=True)
for i, col in enumerate(['B', 'C', 'D', 'E', 'F']):
    src = ['B', 'C', 'D', 'E', 'F'][i]
    formula = '=%s21+%s22+%s23' % (src, src, src)
    cell = ws.cell(row=24, column=2 + i, value=formula)
    cell.font = Font(bold=True)
    cell.number_format = '$#,##0;($#,##0)'
    cell.fill = PatternFill('solid', fgColor=GREY_4)

ws.cell(row=25, column=1, value='Cumulative Net Benefit ($)').font = Font(bold=True, color=PURPLE_DEEP)
ws.cell(row=25, column=2, value='=B24')
ws.cell(row=25, column=3, value='=B25+C24')
ws.cell(row=25, column=4, value='=C25+D24')
ws.cell(row=25, column=5, value='=D25+E24')
for c in range(2, 6):
    ws.cell(row=25, column=c).number_format = '$#,##0;($#,##0)'
    ws.cell(row=25, column=c).font = Font(bold=True, color=PURPLE_DEEP)

for r in range(19, 26):
    for c in range(1, 7):
        ws.cell(row=r, column=c).border = BORDER

# ---- headline metrics -------------------------------------------------------
ws['A28'] = 'HEADLINE METRICS  (these four numbers are on the PPT slide)'
ws['A28'].font = Font(bold=True, color=PURPLE_DEEP, size=11)
METRICS_ROWS = [
    ('Build cost', '=-B22-B23', '$#,##0'),
    ('Year-1 net benefit', '=C24', '$#,##0'),
    ('Payback period (months from go-live)', '=ABS(B24)/(C24/12)', '0.0" months"'),
    ('3-yr cumulative net value', '=E25', '$#,##0'),
    ('3-yr NPV @ 10% discount rate', '=NPV(0.10,C24:E24)+B24', '$#,##0'),
]
for i, (label, formula, fmt) in enumerate(METRICS_ROWS):
    r = 29 + i
    ws.cell(row=r, column=1, value=label).font = Font(bold=True)
    cell = ws.cell(row=r, column=2, value=formula)
    cell.number_format = fmt
    cell.font = Font(bold=True, size=13, color=BLUE)
    cell.fill = PatternFill('solid', fgColor=ICE_BLUE)
    for c in (1, 2):
        ws.cell(row=r, column=c).border = BORDER

note(ws, 'A35', 'All figures on this sheet are illustrative training inputs. Update the '
                'assumptions above with live workshop data before using this as a real '
                'business case — every total, and the priority classification on Sheet 4, '
                'will recalculate.')
ws.merge_cells('A35:F35')
ws.row_dimensions[35].height = 30

widths = {'A': 38, 'B': 16, 'C': 16, 'D': 16, 'E': 16, 'F': 16}
for col, w in widths.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = 'A6'

for ws in wb.worksheets:
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.print_options.horizontalCentered = False

wb.save(OUT)
print('saved', OUT)
