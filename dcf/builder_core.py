#!/usr/bin/env python3
"""
Institutional-grade, fully formula-driven DCF template (8 tabs) built with openpyxl.
Every company-specific number lives in a BLUE input cell (yellow-shaded = fill these first).
Pre-loaded values are ILLUSTRATIVE ("Example Co.") - overwrite them for any listed company.
Usage:  python build_dcf.py [output.xlsx]
"""
import datetime, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment

OUT = globals().get('OUT_PATH') or (sys.argv[1] if len(sys.argv) > 1 else 'DCF_Valuation_Template.xlsx')

# ----------------------------------------------------------------- CONFIG
CUR = globals().get('CUR_OVERRIDE', '$')                                   # change to '₹', '€', '£' ... for share-price format
F_NUM = '#,##0.0;(#,##0.0);"-"'
F_PX = CUR + '#,##0.00'
F_PCT = '0.0%'
F_PCT2 = '0.00%'
F_X = '0.0"x"'
F_DAYS = '0.0'
F_FAC = '0.0000'
F_YR = '0'

BLUE, BLACK, GREEN = '0000FF', '000000', '008000'
NAVY, GREY, YEL, LGREY = '1F3864', 'D9E1F2', 'FFF2CC', 'F2F2F2'

# (scenario lists SC / SK / N are defined in the SCENARIO LIBRARY block below)
HC = ['C', 'D', 'E']
FC = ['F', 'G', 'H', 'I', 'J']
ALL = HC + FC + ['K']

DCF_UFCF_ROW = 37      # row of UFCF on DCF_Valuation (cross-referenced by CFS)
DIL_ROW = 81           # row of diluted shares on DCF_Valuation (cross-referenced by WACC)
WACC_LIVE = 'WACC!$N$46'   # live WACC cell (3 scenario rows 43-45, live row 46)
RATING_FIRST, RATING_LAST = 51, 65  # synthetic-rating table rows on WACC (fixed: LIVE_R=46, RT0=49, rows 51-65)


def pv(c):
    return get_column_letter(column_index_from_string(c) - 1)


wb = Workbook()
NAMES = ['Cover', 'Assumptions', 'Bridge', 'Segments', 'Macro_Scenarios', 'IS', 'BS', 'CFS',
         'WACC', 'DCF_Valuation', 'Comps', 'Scenario_Valuation', 'Sensitivity_Analysis',
         'Precedents', 'Risks', 'Controls', 'Version_Log']
wb.active.title = NAMES[0]
for n in NAMES[1:]:
    wb.create_sheet(n)
WS = {n: wb[n] for n in NAMES}
SECT = {n: set() for n in NAMES}


def put(ws, ref, val, fmt=None, bold=False, kind=None, fill=None, italic=False, align=None, color=None):
    c = ws[ref]
    c.value = val
    if kind is None:
        if isinstance(val, str) and val.startswith('='):
            kind = 'link' if '!' in val else 'calc'
        elif isinstance(val, (int, float, datetime.date)):
            kind = 'input'
        else:
            kind = 'label'
    col = {'input': BLUE, 'calc': BLACK, 'link': GREEN, 'label': BLACK}[kind]
    c.font = Font(name='Arial', size=10, bold=bold, italic=italic, color=color or col)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = PatternFill('solid', fgColor=fill)
    if align:
        c.alignment = Alignment(horizontal=align)
    return c


def band(ws, row, text, fillc=NAVY, fontc='FFFFFF', c1='A', c2='K', sect=True):
    for ci in range(column_index_from_string(c1), column_index_from_string(c2) + 1):
        cell = ws.cell(row=row, column=ci)
        cell.fill = PatternFill('solid', fgColor=fillc)
        cell.font = Font(name='Arial', size=10, bold=True, color=fontc)
    ws[f'{c1}{row}'].value = text
    if sect:
        SECT[ws.title].add(row)


def subband(ws, row, text, unit=None, c2='K'):
    band(ws, row, text, fillc=GREY, fontc=BLACK, c2=c2)
    if unit:
        ws[f'B{row}'].value = unit
        ws[f'B{row}'].font = Font(name='Arial', size=10, italic=True, color=BLACK)


def lab(ws, row, text, unit=None, bold=False, italic=False):
    put(ws, f'A{row}', text, kind='label', bold=bold, italic=italic)
    if unit:
        put(ws, f'B{row}', unit, kind='label', italic=True)


def title(ws, text, sub=None):
    put(ws, 'A1', text, kind='label', bold=True)
    ws['A1'].font = Font(name='Arial', size=14, bold=True, color=NAVY)
    ws.row_dimensions[1].height = 24
    if sub:
        put(ws, 'A2', sub, kind='label', italic=True)


def yr_labels(ws, row, ref, cols=ALL, band_row=True):
    """Year header labels driven by the 'last historical FY' input (ref e.g. Assumptions!$C$12)."""
    for i, c in enumerate(['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']):
        if c not in cols:
            continue
        if i < 3:
            f = f'="FY"&RIGHT({ref}-{2 - i},2)&"A"'
        elif i < 8:
            f = f'="FY"&RIGHT({ref}+{i - 2},2)&"E"'
        else:
            f = f'="FY"&RIGHT({ref}+{i - 2},2)&"E (TY)"'
        cell = ws[f'{c}{row}']
        cell.value = f
        cell.font = Font(name='Arial', size=10, bold=True, color='FFFFFF' if band_row else BLACK)
        cell.alignment = Alignment(horizontal='right')


def dv_list(ws, ref, options):
    dv = DataValidation(type='list', formula1='"' + ','.join(options) + '"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ref)


def P(xs):
    return [x / 100.0 for x in xs]


# =====================================================================================
# SEGMENTS - top-down segment allocation (names + mix filled by hook from CSV lookup)
# =====================================================================================
wsG = WS['Segments']
title(wsG, 'Segments - Revenue Mix & Margin by Business Line (ties to IS)')
put(wsG, 'A2', '="Segment names + revenue mix are pre-filled from the ticker dataset. Overwrite any blue cell; total revenue / EBITDA must tie to the IS."', kind='calc', italic=True)
band(wsG, 4, 'SEGMENT REVENUE MIX (% of total company revenue) - 5Y forecast', c2='K')
yr_labels(wsG, 4, 'Assumptions!$C$12', cols=FC)
SEG_ROWS = list(range(9, 15))
for r in SEG_ROWS:
    lab(wsG, r, f'  Segment {r - 8} (name from dataset)', 'name')
    for c in FC:
        put(wsG, f'{c}{r}', 0.0, F_PCT, kind='input', fill=YEL)
band(wsG, 17, 'CHECK: revenue mix sums to 100% in every forecast year', c2='K')
for c in FC:
    put(wsG, f'{c}18', f'=SUM({c}9:{c}14)', F_PCT, bold=True)
    put(wsG, f'{c}19', f'=IF(ABS({c}18-1)<0.005,"OK","REVIEW")', None, bold=True, align='right')
band(wsG, 21, 'SEGMENT EBITDA MARGIN (% of segment revenue) - 5Y forecast', c2='K')
for r in SEG_ROWS:
    lab(wsG, r + 14, f'  Segment {r - 8} EBITDA margin', '%')
    for c in FC:
        put(wsG, f'{c}{r + 14}', 0.0, F_PCT, kind='input', fill=YEL)
band(wsG, 30, 'SEGMENT REVENUE (units; mix x IS total revenue)', c2='K')
for i, r in enumerate(SEG_ROWS):
    lab(wsG, 31 + i, f'  Segment {i + 1} revenue')
    for c in FC:
        put(wsG, f'{c}{31 + i}', f'={c}{r}*IS!{c}11', F_NUM)
lab(wsG, 37, 'Total segment revenue (ties to IS)', 'units', bold=True)
for c in FC:
    put(wsG, f'{c}37', f'=SUM({c}31:{c}36)', F_NUM, bold=True)
lab(wsG, 38, 'Check vs IS revenue (must be 0)', 'units', bold=True)
for c in FC:
    put(wsG, f'{c}38', f'={c}37-IS!{c}11', F_NUM, bold=True)
band(wsG, 40, 'SEGMENT EBITDA (units; revenue x margin)', c2='K')
for i, r in enumerate(SEG_ROWS):
    lab(wsG, 41 + i, f'  Segment {i + 1} EBITDA')
    for c in FC:
        put(wsG, f'{c}{41 + i}', f'={c}{31 + i}*{c}{r + 14}', F_NUM)
lab(wsG, 47, 'Total segment EBITDA (ties to IS)', 'units', bold=True)
for c in FC:
    put(wsG, f'{c}47', f'=SUM({c}41:{c}46)', F_NUM, bold=True)
lab(wsG, 48, 'Check vs IS EBITDA (must be 0)', 'units', bold=True)
for c in FC:
    put(wsG, f'{c}48', f'={c}47-IS!{c}18', F_NUM, bold=True)
wsG.freeze_panes = 'C5'

# =====================================================================================
# 3. MACRO_SCENARIOS  (blocks first - Assumptions links to these rows)
# =====================================================================================
# ------------------------------------------------------------------ SCENARIO LIBRARY (Bear / Base / Bull)
SHAPES = {'flat': [1, 1, 1, 1, 1], 'fade': [1, .8, .5, .3, .2], 'fadeslow': [1, 1, .7, .5, .5],
          'shock': [1, .6, .2, 0, 0], 'build': [.2, .5, .8, 1, 1]}
BASEV = dict(
    gdp=[.022, .022, .020, .020, .020], cpi=[.030, .028, .025, .025, .025], rf=[.043, .042, .042, .042, .042],
    tam=[.040, .040, .038, .035, .035], comm=[.030, .030, .028, .025, .025], fx=[100.0] * 5,
    share=[.003] * 5, spread=[0.0] * 5, gmb=[.420, .425, .430, .430, .430], sga=[.138, .135, .132, .130, .128],
    rd=[.060] * 5, dso=[45.0] * 5, dio=[60.0] * 5, dpo=[50.0] * 5, gcx=[.026, .025, .024, .024, .024])
SCEN = []


def _S(key, name, desc, shape, prob, **d):
    SCEN.append((key, name, desc, shape, prob, d))


_S('bear', 'Bear', 'Weak growth, sticky inflation, higher rates, margin pressure', 'fadeslow', 20,
   gdp=-.017, cpi=.015, rf=.007, tam=-.020, comm=.030, fx=3, share=-.013, spread=-.010, gmb=-.015, sga=.004, rd=.002, dso=3, dio=5, dpo=-3, gcx=-.006, prem=.010)
_S('base', 'Base', 'Consensus-style path: moderate growth, inflation normalising', 'flat', 60)
_S('bull', 'Bull', 'Strong demand, disinflation, falling rates, operating leverage', 'fadeslow', 20,
   gdp=.010, cpi=-.008, rf=-.005, tam=.020, comm=-.015, fx=-1.5, share=.007, spread=.005, gmb=.010, sga=-.004, rd=-.001, dso=-2, dio=-3, dpo=2, gcx=.004, prem=-.005)
SK = [s[0] for s in SCEN]
SC = [s[1] for s in SCEN]
N = len(SCEN)
assert sum(s[4] for s in SCEN) == 100


def apply_delta(var, base):
    """Scenario path for a driver = Base path + delta x shape (risk-free shift is always flat)."""
    out = {}
    for key, name, desc, shape, prob, d in SCEN:
        sh = SHAPES['flat'] if var == 'rf' else SHAPES[shape]
        out[key] = [base[i] + d.get(var, 0.0) * sh[i] for i in range(5)]
    return out


def VAL(var):
    ad = apply_delta(var, BASEV[var])
    return [ad[k] for k in SK]

wsM = WS['Macro_Scenarios']
title(wsM, f'Macro_Scenarios - Macroeconomic & Industry Assumptions ({N} scenarios: Bear / Base / Bull)',
      'Blue = input | Green = cross-sheet link | Black = formula.  Rows marked "Live" switch with the scenario selector on Assumptions!C5.')
band(wsM, 4, 'MACRO SCENARIO MATRIX')
yr_labels(wsM, 4, 'Assumptions!$C$12', cols=FC)
wsM['E4'].value = 'FY25A / Spot'
wsM['E4'].font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
wsM['E4'].alignment = Alignment(horizontal='right')
M = {}
mr = 5


def macro_block(key, ttl, unit, fmt, vals, e_val=None):
    global mr
    subband(wsM, mr, ttl, unit)
    for i, s in enumerate(SK):
        rr = mr + 1 + i
        lab(wsM, rr, f'  {SC[i]}')
        if e_val is not None:
            put(wsM, f'E{rr}', e_val if i == 0 else f'=E{mr + 1}', fmt)
        for j, col in enumerate(FC):
            put(wsM, f'{col}{rr}', round(vals[i][j], 6), fmt)
    rr = mr + N + 1
    lab(wsM, rr, '  \u25ba Live (active scenario)', bold=True)
    for col in (['E'] if e_val is not None else []) + FC:
        put(wsM, f'{col}{rr}', f'=INDEX({col}{mr + 1}:{col}{mr + N},Assumptions!$C$6)', fmt, bold=True)
    M[key] = {s: mr + 1 + i for i, s in enumerate(SK)}
    M[key]['live'] = rr
    mr += N + 3


macro_block('gdp', 'Real GDP growth', '% YoY', F_PCT, VAL('gdp'))
macro_block('cpi', 'CPI inflation', '% YoY', F_PCT, VAL('cpi'))
macro_block('rf', 'Benchmark risk-free rate / policy rate (long-dated govt yield)', '%', F_PCT2, VAL('rf'))
macro_block('tam', 'Industry TAM growth (volume)', '% YoY', F_PCT, VAL('tam'))
macro_block('comm', 'Commodity / input-cost inflation', '% YoY', F_PCT, VAL('comm'))
macro_block('fx', 'Reporting-currency strength index (rising = stronger home currency = translation headwind)', 'index', '0.0',
            VAL('fx'), e_val=100.0)

subband(wsM, mr, 'Scenario cost-of-equity adjustment (added to Ke on WACC tab)', '% pts')
M['prem'] = {}
for i, s in enumerate(SK):
    lab(wsM, mr + 1 + i, f'  {SC[i]}')
    put(wsM, f'C{mr + 1 + i}', SCEN[i][5].get('prem', 0.0), F_PCT2)
    M['prem'][s] = mr + 1 + i
mr += N + 3
M['transmission_start'] = mr

# =====================================================================================
# 2. ASSUMPTIONS
# =====================================================================================
wsA = WS['Assumptions']
title(wsA, 'Assumptions - Control Center & Drivers',
      'LEGEND: Blue font = hard-coded input | Black = same-sheet formula | Green = cross-sheet link | Yellow fill = fill these cells first for a new company.')
# (top 3 rows are blank/hidden on every sheet - see the formatting pass)
A = {}

band(wsA, 4, 'A. CONTROL PANEL & COMPANY INFORMATION')
lab(wsA, 5, 'Active scenario (Bear / Base / Bull)', 'selector', bold=True)
put(wsA, 'C5', 'Base', kind='input', fill=YEL, bold=True)
_dv = DataValidation(type='list', formula1=f'$N$5:$N${4 + N}', allow_blank=False)
wsA.add_data_validation(_dv)
_dv.add('C5')
lab(wsA, 6, f'Scenario index (1..{N}: 1 = Bear, 2 = Base, 3 = Bull)', 'calc')
put(wsA, 'C6', f'=IFERROR(MATCH(C5,$N$5:$N${4 + N},0),2)', '0')
put(wsA, 'M4', '#', kind='label', bold=True, align='right')
put(wsA, 'N4', f'SCENARIO LIBRARY ({N})', kind='label', bold=True)
put(wsA, 'O4', 'Description', kind='label', bold=True)
for _i, (_k, _nm, _ds, _sh, _pb, _d) in enumerate(SCEN):
    put(wsA, f'M{5 + _i}', _i + 1, '0', kind='label', align='right')
    put(wsA, f'N{5 + _i}', _nm, kind='input')
    put(wsA, f'O{5 + _i}', _ds, kind='label', italic=True)
for row, label, unit, val, fmt in [
    (7, 'Company name', 'text', 'Example Co. (ILLUSTRATIVE)', None),
    (8, 'Ticker', 'text', 'EXCO', None),
    (9, 'Reporting currency', 'text', 'USD', None),
    (10, 'Units for all financials', 'text', 'USD Millions', None),
    (11, 'Valuation date', 'date', datetime.date(2026, 11, 16), 'mmmm d, yyyy'),
    (12, 'Last historical fiscal year (FY of "A" column)', 'year', 2025, F_YR),
]:
    lab(wsA, row, label, unit)
    put(wsA, f'C{row}', val, fmt, kind='input', fill=YEL)
lab(wsA, 13, 'Treat stock-based comp as a real cash cost in UFCF?', 'Yes / No')
put(wsA, 'C13', 'Yes', kind='input', fill=YEL)
dv_list(wsA, 'C13', ['Yes', 'No'])
lab(wsA, 14, 'Days in year (working-capital day counts)', 'days')
put(wsA, 'C14', 365, '0', fill=YEL)

band(wsA, 16, 'B. CAPITAL STRUCTURE & MARKET INPUTS')
mk = [
    (17, 'Current share price', CUR, 48.00, F_PX),
    (18, 'Basic shares outstanding', 'mm', 600.0, F_NUM),
]
for row, label, unit, val, fmt in mk:
    lab(wsA, row, label, unit)
    put(wsA, f'C{row}', val, fmt, fill=YEL)
lab(wsA, 19, 'Dilutive securities (Treasury Stock Method)', None, bold=True)
put(wsA, 'C19', 'Number (mm)', kind='label', bold=True, align='right')
put(wsA, 'D19', 'Strike price', kind='label', bold=True, align='right')
for row, label, n, k in [(20, 'RSUs / PSUs (no strike)', 8.0, 0.0), (21, 'Options - tranche 1', 15.0, 30.0),
                         (22, 'Options - tranche 2', 10.0, 45.0), (23, 'Options - tranche 3', 6.0, 60.0)]:
    lab(wsA, row, label, 'mm / strike')
    put(wsA, f'C{row}', n, F_NUM, fill=YEL)
    put(wsA, f'D{row}', k, F_PX, fill=YEL)
for row, label, unit, val, fmt in [
    (24, 'Total debt (latest balance sheet)', 'units', 1700.0, F_NUM),
    (25, 'Cash & marketable securities (latest balance sheet)', 'units', 1250.0, F_NUM),
    (26, 'Minority interest (non-controlling)', 'units', 0.0, F_NUM),
    (27, 'Preferred stock', 'units', 0.0, F_NUM),
    (28, 'Marginal tax rate (cost of debt shield)', '%', 0.25, F_PCT),
    (29, 'Effective cash tax rate on EBIT', '%', 0.21, F_PCT),
    (30, 'Interest rate on existing debt (supplemental schedule)', '%', 0.055, F_PCT2),
]:
    lab(wsA, row, label, unit)
    put(wsA, f'C{row}', val, fmt, fill=YEL)

band(wsA, 32, 'C. TERMINAL VALUE INPUTS')
for row, label, unit, val, fmt in [
    (33, 'Terminal perpetuity growth rate (g)', '%', 0.025, F_PCT2),
    (34, 'Terminal exit EV / EBITDA multiple', 'x', 11.0, F_X),
    (35, 'Weight on Gordon Growth in blended value (rest = Exit Multiple)', '%', 0.5, F_PCT),
    (36, 'Tolerance: max terminal RONIC above WACC before audit flags REVIEW', '%', 0.05, F_PCT),
    (37, 'Terminal ROIC (g / reinvestment rate; sector default 15%)', '%', 0.15, F_PCT),
]:
    lab(wsA, row, label, unit)
    put(wsA, f'C{row}', val, fmt, fill=YEL)

band(wsA, 38, 'D. HISTORICAL FINANCIALS (INPUTS)')
yr_labels(wsA, 38, '$C$12', cols=HC)
HIST = [
    ('rev', 'Revenue', [10000, 10800, 11700]),
    ('cogs', 'COGS (excl. D&A)', [5850, 6286, 6786]),
    ('sga', 'SG&A (excl. D&A and SBC)', [1450, 1534, 1638]),
    ('rd', 'R&D (excl. D&A and SBC)', [600, 648, 702]),
    ('oth', 'Other operating expenses', [200, 216, 234]),
    ('da', 'Depreciation & amortization', [340, 367, 410]),
    ('sbc', 'Stock-based compensation', [120, 130, 140]),
    ('capex', 'Capital expenditures', [520, 572, 632]),
    ('ar', 'Accounts receivable', [1260, 1332, 1442]),
    ('inv', 'Inventory', [994, 1050, 1115]),
    ('oca', 'Prepaid & other current assets', [300, 324, 351]),
    ('ap', 'Accounts payable', [801, 861, 930]),
    ('ocl', 'Accrued & other current liabilities', [900, 972, 1053]),
    ('ppe', 'Net PP&E (ending)', [3200, 3405, 3627]),
    ('int', 'Interest expense', [105, 100, 95]),
    ('debt', 'Total debt (ending)', [1900, 1800, 1700]),
    ('cash', 'Cash & marketable securities (ending)', [900, 1050, 1250]),
    ('equity', 'Stockholders equity (book, excl. minority)', [5200, 5800, 6400]),
]
HR = {}
for i, (k, label, vals) in enumerate(HIST):
    row = 39 + i
    HR[k] = row
    lab(wsA, row, label, 'units')
    for j, c in enumerate(HC):
        put(wsA, f'{c}{row}', float(vals[j]), F_NUM, fill=YEL)
put(wsA, 'A57', 'ILLUSTRATIVE DATA: overwrite every blue cell (10-K / consensus). EBITDA / EBIT / NI are post-SBC (GAAP-style); the BS plug keeps history balancing.',
    kind='label', italic=True, color='C00000')

band(wsA, 58, 'E. OPERATING DRIVERS (HISTORICAL + 5-YEAR FORECAST ACROSS BEAR / BASE / BULL)')
yr_labels(wsA, 58, '$C$12', cols=ALL[:-1])
lab(wsA, 59, 'Gross-margin sensitivity to input-cost inflation above CPI (x)', 'x')
put(wsA, 'C59', 0.30, '0.00', fill=YEL)
lab(wsA, 60, 'Share of revenue exposed to FX translation', '%')
put(wsA, 'C60', 0.10, F_PCT, fill=YEL)
SENS, FXEXP = '$C$59', '$C$60'
ar_ = 62


def drv_block(key, ttl, unit, fmt, vals=None, fml=None, hist=None):
    """One row per scenario (inputs or formulas) + Live row = INDEX on the scenario selector."""
    global ar_
    subband(wsA, ar_, ttl, unit, c2='K')
    for i, s in enumerate(SK):
        rr = ar_ + 1 + i
        lab(wsA, rr, f'  {SC[i]}')
        for j, col in enumerate(FC):
            if vals is not None:
                put(wsA, f'{col}{rr}', round(vals[i][j], 6), fmt)
            else:
                put(wsA, f'{col}{rr}', fml(s, col, pv(col)), fmt)
    rr = ar_ + N + 1
    lab(wsA, rr, '  \u25ba Live (active scenario)', bold=True)
    if hist:
        for c in HC:
            f = hist(c, pv(c))
            if f:
                put(wsA, f'{c}{rr}', f, fmt, bold=True)
    for col in FC:
        put(wsA, f'{col}{rr}', f'=INDEX({col}{ar_ + 1}:{col}{ar_ + N},$C$6)', fmt, bold=True)
    A[key] = {s: ar_ + 1 + i for i, s in enumerate(SK)}
    A[key]['live'] = rr
    ar_ += N + 3


H = HR
CONST = lambda vec: [list(vec) for _ in range(N)]
drv_block('share', 'Market-share gain / (loss) vs. industry (added to Industry TAM growth)', '% pts', F_PCT, vals=VAL('share'))
drv_block('spread', 'Pricing power: price growth spread over CPI inflation', '% pts', F_PCT, vals=VAL('spread'))
drv_block('rg', 'Revenue growth = (1 + TAM + share) x (1 + CPI + spread) - 1 - FX exposure x \u0394FX index', '% YoY', F_PCT,
          fml=lambda s, c, p: (f"=(1+Macro_Scenarios!{c}{M['tam'][s]}+{c}{A['share'][s]})*(1+Macro_Scenarios!{c}{M['cpi'][s]}+{c}{A['spread'][s]})-1"
                               f"-{FXEXP}*(Macro_Scenarios!{c}{M['fx'][s]}/Macro_Scenarios!{p}{M['fx'][s]}-1)"),
          hist=lambda c, p: None if c == 'C' else f'={c}{H["rev"]}/{p}{H["rev"]}-1')
drv_block('gmb', 'Gross margin before input-cost pressure', '% of revenue', F_PCT, vals=VAL('gmb'))
drv_block('gm', 'Gross margin = pre-pressure margin - sensitivity x (commodity inflation - CPI)', '% of revenue', F_PCT,
          fml=lambda s, c, p: f"={c}{A['gmb'][s]}-{SENS}*(Macro_Scenarios!{c}{M['comm'][s]}-Macro_Scenarios!{c}{M['cpi'][s]})",
          hist=lambda c, p: f'=1-{c}{H["cogs"]}/{c}{H["rev"]}')
drv_block('sga', 'SG&A % of revenue', '% of revenue', F_PCT, vals=VAL('sga'), hist=lambda c, p: f'={c}{H["sga"]}/{c}{H["rev"]}')
drv_block('rd', 'R&D % of revenue', '% of revenue', F_PCT, vals=VAL('rd'), hist=lambda c, p: f'={c}{H["rd"]}/{c}{H["rev"]}')
drv_block('oth', 'Other OpEx % of revenue', '% of revenue', F_PCT, vals=CONST([.02] * 5), hist=lambda c, p: f'={c}{H["oth"]}/{c}{H["rev"]}')
drv_block('sbc', 'Stock-based compensation % of revenue', '% of revenue', F_PCT, vals=CONST([.012] * 5), hist=lambda c, p: f'={c}{H["sbc"]}/{c}{H["rev"]}')
drv_block('dso', 'DSO - accounts receivable days (on revenue)', 'days', F_DAYS, vals=VAL('dso'), hist=lambda c, p: f'={c}{H["ar"]}/{c}{H["rev"]}*$C$14')
drv_block('dio', 'DIO - inventory days (on COGS)', 'days', F_DAYS, vals=VAL('dio'), hist=lambda c, p: f'={c}{H["inv"]}/{c}{H["cogs"]}*$C$14')
drv_block('dpo', 'DPO - accounts payable days (on COGS)', 'days', F_DAYS, vals=VAL('dpo'), hist=lambda c, p: f'={c}{H["ap"]}/{c}{H["cogs"]}*$C$14')
drv_block('oca', 'Prepaid & other current assets % of revenue', '% of revenue', F_PCT, vals=CONST([.03] * 5), hist=lambda c, p: f'={c}{H["oca"]}/{c}{H["rev"]}')
drv_block('ocl', 'Accrued & other current liabilities % of revenue', '% of revenue', F_PCT, vals=CONST([.09] * 5), hist=lambda c, p: f'={c}{H["ocl"]}/{c}{H["rev"]}')
drv_block('mcx', 'Maintenance CapEx % of revenue', '% of revenue', F_PCT, vals=CONST([.038] * 5))
drv_block('gcx', 'Growth CapEx % of revenue', '% of revenue', F_PCT, vals=VAL('gcx'))
drv_block('da', 'Depreciation & amortization % of revenue', '% of revenue', F_PCT, vals=CONST([.035, .035, .036, .036, .037]),
          hist=lambda c, p: f'={c}{H["da"]}/{c}{H["rev"]}')
subband(wsA, ar_, 'Memo: total CapEx % of revenue (Live case)', '% of revenue')
lab(wsA, ar_ + 1, '  Total CapEx % revenue (maintenance + growth)')
for c in HC:
    put(wsA, f'{c}{ar_ + 1}', f'={c}{H["capex"]}/{c}{H["rev"]}', F_PCT)
for c in FC:
    put(wsA, f'{c}{ar_ + 1}', f'={c}{A["mcx"]["live"]}+{c}{A["gcx"]["live"]}', F_PCT)
A['capex_tot'] = ar_ + 1
ar_ += 3
subband(wsA, ar_, 'Debt drawdown / (repayment) - supplemental credit schedule (all scenarios)', 'units')
lab(wsA, ar_ + 1, '  Net debt drawdown / (repayment)')
for c in FC:
    put(wsA, f'{c}{ar_ + 1}', -100.0, F_NUM)
A['draw'] = ar_ + 1
ar_ += 3
band(wsA, ar_, 'F. NOTES & METHODOLOGY')
notes = [
    f'Scenarios: Bear / Base / Bull. Bear and Bull = Base path + a shaped delta (fade-slow) on macro drivers, margins, working capital and CapEx; Base is the consensus path. Every driver cell is an editable blue input.',
    'Macro link: volume growth = Industry TAM growth + share gain; price growth = CPI + spread; revenue growth compounds both and adds FX translation. Change Macro_Scenarios and the DCF updates.',
    'Gross margin = pre-pressure margin less sensitivity x (commodity inflation - CPI): cost shocks beyond inflation compress margin.',
    'Terminal year (FY31E): every flow = FY30E x (1 + g); NWC change = FY30E NWC x g. Terminal value on Gordon uses this normalised UFCF.',
    'Diluted shares use the Treasury Stock Method at the CURRENT share price (avoids circularity); shares are held constant across scenarios.',
    'Valuation is struck at the end of the last historical fiscal year (no stub-period adjustment); roll the base year forward as new actuals are reported.',
    'Synthetic credit rating table (WACC tab) is illustrative - refresh annually from Damodaran / rating-agency data. All pre-loaded numbers are placeholders.',
]
for i, t in enumerate(notes):
    put(wsA, f'A{ar_ + 1 + i}', f'{i + 1}. {t}', kind='label', italic=True)
A['_end'] = ar_ + 1 + len(notes)

# ---- Macro transmission summary (needs Assumptions rows) -------------------------------
mr = M['transmission_start']
subband(wsM, mr, 'MACRO-TO-OPERATING TRANSMISSION (Live case - flows straight into Assumptions / IS)', None)
trans = [
    ('Nominal GDP growth = (1 + real GDP) x (1 + CPI) - 1', lambda c: f'=(1+{c}{M["gdp"]["live"]})*(1+{c}{M["cpi"]["live"]})-1', F_PCT),
    ('Real policy spread: risk-free rate less CPI', lambda c: f'={c}{M["rf"]["live"]}-{c}{M["cpi"]["live"]}', F_PCT),
    ('Volume growth = Industry TAM growth + market-share gain / (loss)', lambda c: f'=IS!{c}7', F_PCT),
    ('Price growth = CPI + pricing spread', lambda c: f'=IS!{c}8', F_PCT),
    ('FX translation impact on revenue', lambda c: f'=IS!{c}9', F_PCT),
    ('Resulting revenue growth', lambda c: f'=IS!{c}10', F_PCT),
    ('Gross-margin drag from input costs (sensitivity x (commodity - CPI))',
     lambda c: f'=-Assumptions!{SENS}*({c}{M["comm"]["live"]}-{c}{M["cpi"]["live"]})', F_PCT2),
    ('Resulting gross margin', lambda c: f'=IS!{c}14', F_PCT),
]
for i, (label, fn, fmt) in enumerate(trans):
    lab(wsM, mr + 1 + i, label)
    for c in FC:
        put(wsM, f'{c}{mr + 1 + i}', fn(c), fmt)

# =====================================================================================
# 4. IS - INCOME STATEMENT (full P&L to net income; Live scenario)
# =====================================================================================
wsO = WS['IS']
title(wsO, 'IS - Income Statement (historical + 5Y forecast; post-SBC, GAAP-style)')
put(wsO, 'A2', '="All figures in "&Assumptions!C10&" | Active scenario: "&Assumptions!C5&" | Last column = terminal year (FY30E x (1 + g))"',
    kind='calc', italic=True)
band(wsO, 4, 'FISCAL YEAR')
for c in ALL:
    wsO[f'{c}4'].value = (f'=IF({c}5<=Assumptions!$C$12,"FY"&RIGHT({c}5,2)&"A","FY"&RIGHT({c}5,2)&"E")' if c != 'K'
                          else '="FY"&RIGHT(K5,2)&"E (TY)"')
    wsO[f'{c}4'].font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    wsO[f'{c}4'].alignment = Alignment(horizontal='right')
put(wsO, 'C5', '=Assumptions!$C$12-2', F_YR)
for c in ALL[1:]:
    put(wsO, f'{c}5', f'={pv(c)}5+1', F_YR)
lab(wsO, 5, 'Fiscal year (numeric)')


def O_row(row, label, unit=None, bold=False, italic=False):
    lab(wsO, row, label, unit, bold=bold, italic=italic)


def O_fill(row, cols, fn, fmt, bold=False):
    for c in cols:
        f = fn(c, pv(c))
        if f is not None:
            put(wsO, f'{c}{row}', f, fmt, bold=bold)


ASM = 'Assumptions!'
HIST_LINK = lambda key: (lambda c, p: f'={ASM}{c}{HR[key]}')
GK = '$K$10'

band(wsO, 6, 'INCOME STATEMENT')
O_row(7, '  Volume growth (TAM + share)', '%', italic=True)
O_fill(7, FC, lambda c, p: f'=Macro_Scenarios!{c}{M["tam"]["live"]}+Assumptions!{c}{A["share"]["live"]}', F_PCT)
O_row(8, '  Price / ASP growth (CPI + spread)', '%', italic=True)
O_fill(8, FC, lambda c, p: f'=Macro_Scenarios!{c}{M["cpi"]["live"]}+Assumptions!{c}{A["spread"]["live"]}', F_PCT)
O_row(9, '  FX translation impact', '%', italic=True)
O_fill(9, FC, lambda c, p: f'=-Assumptions!{FXEXP}*(Macro_Scenarios!{c}{M["fx"]["live"]}/Macro_Scenarios!{p}{M["fx"]["live"]}-1)', F_PCT)
O_row(10, 'Revenue growth', '%', bold=True)
O_fill(10, ['D', 'E'], lambda c, p: f'={c}11/{p}11-1', F_PCT)
O_fill(10, FC, lambda c, p: f'={ASM}{c}{A["rg"]["live"]}', F_PCT)
put(wsO, 'K10', '=Assumptions!$C$33', F_PCT, bold=True)
O_row(11, 'Revenue', 'units', bold=True)
O_fill(11, HC, HIST_LINK('rev'), F_NUM, bold=True)
O_fill(11, FC + ['K'], lambda c, p: f'={p}11*(1+{c}10)', F_NUM, bold=True)
O_row(12, '(-) COGS (excl. D&A)', 'units')
O_fill(12, HC, HIST_LINK('cogs'), F_NUM)
O_fill(12, FC, lambda c, p: f'={c}11*(1-{ASM}{c}{A["gm"]["live"]})', F_NUM)
O_row(13, 'Gross profit', 'units', bold=True)
O_fill(13, ALL, lambda c, p: f'={c}11-{c}12', F_NUM, bold=True)
O_row(14, '  Gross margin', '%', italic=True)
O_fill(14, ALL, lambda c, p: f'={c}13/{c}11', F_PCT)
for row, label, key, akey in [(15, '(-) SG&A', 'sga', 'sga'), (16, '(-) R&D', 'rd', 'rd'), (17, '(-) Other operating expenses (incl. SBC)', 'oth', 'oth')]:
    O_row(row, label, 'units')
    O_fill(row, HC, (lambda c, p: f'={HIST_LINK(key)(c, p)[1:]}+{HIST_LINK("sbc")(c, p)[1:]}') if key == 'oth' else HIST_LINK(key), F_NUM)
    O_fill(row, FC, (lambda c, p, ak=akey: f'={c}11*({ASM}{c}{A[ak]["live"]}+{ASM}{c}{A["sbc"]["live"]})') if key == 'oth' else (lambda c, p, ak=akey: f'={c}11*{ASM}{c}{A[ak]["live"]}'), F_NUM)
O_row(18, 'EBITDA (post-SBC, GAAP-style)', 'units', bold=True)
O_fill(18, ALL, lambda c, p: f'={c}13-{c}15-{c}16-{c}17', F_NUM, bold=True)
O_row(19, '  EBITDA margin', '%', italic=True)
O_fill(19, ALL, lambda c, p: f'={c}18/{c}11', F_PCT)
O_row(20, '(-) Depreciation & amortization', 'units')
O_fill(20, HC, HIST_LINK('da'), F_NUM)
O_fill(20, FC, lambda c, p: f'={c}11*{ASM}{c}{A["da"]["live"]}', F_NUM)
O_row(21, 'EBIT', 'units', bold=True)
O_fill(21, ALL, lambda c, p: f'={c}18-{c}20', F_NUM, bold=True)
O_row(22, '  EBIT margin', '%', italic=True)
O_fill(22, ALL, lambda c, p: f'={c}21/{c}11', F_PCT)
O_row(23, '(-) Cash taxes on EBIT (effective rate)', 'units')
O_fill(23, ALL, lambda c, p: f'={c}21*{ASM}$C$29', F_NUM)
O_row(24, 'NOPAT (EBIAT)', 'units', bold=True)
O_fill(24, ALL, lambda c, p: f'={c}21-{c}23', F_NUM, bold=True)
O_row(25, '  NOPAT margin', '%', italic=True)
O_fill(25, ALL, lambda c, p: f'={c}24/{c}11', F_PCT)
O_row(26, 'Memo: stock-based compensation (in opex above)', 'units', italic=True)
O_fill(26, HC, HIST_LINK('sbc'), F_NUM)
O_fill(26, FC, lambda c, p: f'={c}11*{ASM}{c}{A["sbc"]["live"]}', F_NUM)
O_row(27, '(-) Interest expense (average debt; BS schedule)', 'units')
O_fill(27, HC, HIST_LINK('int'), F_NUM)
O_fill(27, FC, lambda c, p: f'=BS!{c}42', F_NUM)
O_row(28, 'Earnings before taxes', 'units', bold=True)
O_fill(28, ALL, lambda c, p: f'={c}21-{c}27', F_NUM, bold=True)
O_row(29, '(-) Cash taxes on EBT (effective rate)', 'units')
O_fill(29, ALL, lambda c, p: f'={c}28*{ASM}$C$29', F_NUM)
O_row(30, 'Net income (post-SBC)', 'units', bold=True)
O_fill(30, ALL, lambda c, p: f'={c}28-{c}29', F_NUM, bold=True)
O_row(31, '  Net margin', '%', italic=True)
O_fill(31, ALL, lambda c, p: f'={c}30/{c}11', F_PCT)
O_row(32, 'Diluted EPS (post-SBC)', CUR)
O_fill(32, HC, lambda c, p: f'={c}30/DCF_Valuation!$F$104', F_PX)
O_fill(32, FC + ['K'], lambda c, p: f'={c}30/DCF_Valuation!{c}104', F_PX)
for row in (12, 15, 16, 17, 20, 26, 27):
    put(wsO, f'K{row}', f'=J{row}*(1+{GK})', F_NUM)
wsO.freeze_panes = 'C6'

# =====================================================================================
# 5. BS - BALANCE SHEET (articulated; balance check must be zero)
# =====================================================================================
wsBS = WS['BS']
title(wsBS, 'BS - Balance Sheet (historical + 5Y forecast)')
put(wsBS, 'A2', '="All figures in "&Assumptions!C10&" | Active scenario: "&Assumptions!C5&" | Check row 24 = 0 in every year (historical plug = other assets)"',
    kind='calc', italic=True)
band(wsBS, 4, 'FISCAL YEAR')
for c in HC + FC:
    wsBS[f'{c}4'].value = f'=IF({c}5<=Assumptions!$C$12,"FY"&RIGHT({c}5,2)&"A","FY"&RIGHT({c}5,2)&"E")'
    wsBS[f'{c}4'].font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    wsBS[f'{c}4'].alignment = Alignment(horizontal='right')
put(wsBS, 'C5', '=Assumptions!$C$12-2', F_YR)
for c in (HC + FC)[1:]:
    put(wsBS, f'{c}5', f'={pv(c)}5+1', F_YR)
lab(wsBS, 5, 'Fiscal year (numeric)')
ALLBS = HC + FC


def BS_row(row, label, unit=None, bold=False, italic=False):
    lab(wsBS, row, label, unit, bold=bold, italic=italic)


def BS_fill(row, fn, fmt, bold=False):
    for c in ALLBS:
        f = fn(c, pv(c))
        if f is not None:
            put(wsBS, f'{c}{row}', f, fmt, bold=bold)


band(wsBS, 6, 'BALANCE SHEET')
BS_row(7, 'Cash & marketable securities', 'units')
BS_fill(7, lambda c, p: (HIST_LINK('cash')(c, p) if c in HC else f'=CFS!{c}19'), F_NUM)
BS_row(8, 'Accounts receivable', 'units')
BS_fill(8, lambda c, p: (HIST_LINK('ar')(c, p) if c in HC else f'=IS!{c}11*{ASM}{c}{A["dso"]["live"]}/{ASM}$C$14'), F_NUM)
BS_row(9, 'Inventory', 'units')
BS_fill(9, lambda c, p: (HIST_LINK('inv')(c, p) if c in HC else f'=IS!{c}12*{ASM}{c}{A["dio"]["live"]}/{ASM}$C$14'), F_NUM)
BS_row(10, 'Prepaid & other current assets', 'units')
BS_fill(10, lambda c, p: (HIST_LINK('oca')(c, p) if c in HC else f'=IS!{c}11*{ASM}{c}{A["oca"]["live"]}'), F_NUM)
BS_row(11, 'PP&E, net (schedule below)', 'units')
BS_fill(11, lambda c, p: f'={c}30', F_NUM)
BS_row(12, 'Other assets (hist plug; forecast held flat)', 'units')
BS_fill(12, lambda c, p: (f'={c}15+{c}16+{c}17+{c}22-({c}7+{c}8+{c}9+{c}10+{c}11)' if c in HC else '=$E$12'), F_NUM)
BS_row(13, 'TOTAL ASSETS', 'units', bold=True)
BS_fill(13, lambda c, p: f'=SUM({c}7:{c}12)', F_NUM, bold=True)
BS_row(15, 'Accounts payable', 'units')
BS_fill(15, lambda c, p: (HIST_LINK('ap')(c, p) if c in HC else f'=IS!{c}12*{ASM}{c}{A["dpo"]["live"]}/{ASM}$C$14'), F_NUM)
BS_row(16, 'Accrued & other current liabilities', 'units')
BS_fill(16, lambda c, p: (HIST_LINK('ocl')(c, p) if c in HC else f'=IS!{c}11*{ASM}{c}{A["ocl"]["live"]}'), F_NUM)
BS_row(17, 'Total debt (schedule below)', 'units')
BS_fill(17, lambda c, p: f'={c}40', F_NUM)
BS_row(18, 'TOTAL LIABILITIES', 'units', bold=True)
BS_fill(18, lambda c, p: f'={c}15+{c}16+{c}17', F_NUM, bold=True)
BS_row(20, 'Common & APIC (opening + NI + SBC - dividends)', 'units')
BS_fill(20, lambda c, p: (HIST_LINK('equity')(c, p) if c in HC else f'={p}20+IS!{c}30+IS!{c}26+CFS!{c}15+CFS!{c}21'), F_NUM)
BS_row(21, 'Minority interest', 'units')
for c in ALLBS:
    put(wsBS, f'{c}21', '=Assumptions!$C$26', F_NUM)
BS_row(22, 'TOTAL EQUITY', 'units', bold=True)
BS_fill(22, lambda c, p: f'={c}20+{c}21', F_NUM, bold=True)
BS_row(23, 'TOTAL LIABILITIES & EQUITY', 'units', bold=True)
BS_fill(23, lambda c, p: f'={c}18+{c}22', F_NUM, bold=True)
BS_row(24, 'CHECK: Assets - Liabilities - Equity (must be 0)', 'units', bold=True)
BS_fill(24, lambda c, p: f'={c}13-{c}23', F_NUM, bold=True)
BS_row(25, '  Balance check status', None)
BS_fill(25, lambda c, p: f'=IF(ABS({c}24)<0.5,"OK","ERROR")', None, bold=True)
BS_row(26, '  Memo: net working capital (NWC)', 'units', italic=True)
BS_fill(26, lambda c, p: f'={c}8+{c}9+{c}10-{c}15-{c}16', F_NUM)
BS_row(27, '  Memo: NWC % of revenue', '%', italic=True)
BS_fill(27, lambda c, p: f'={c}26/IS!{c}11', F_PCT)

band(wsBS, 29, 'NET PP&E SCHEDULE')
BS_row(30, 'Ending net PP&E', 'units', bold=True)
BS_fill(30, lambda c, p: (HIST_LINK('ppe')(c, p) if c in HC else f'={c}31+{c}32-{c}33'), F_NUM, bold=True)
BS_row(31, 'Beginning net PP&E', 'units')
for c in FC:
    put(wsBS, f'{c}31', f'={pv(c)}30', F_NUM)
BS_row(32, '(+) Total CapEx (maintenance + growth)', 'units')
BS_fill(32, lambda c, p: (HIST_LINK('capex')(c, p) if c in HC else f'=IS!{c}11*({ASM}{c}{A["mcx"]["live"]}+{ASM}{c}{A["gcx"]["live"]})'), F_NUM)
BS_row(33, '(-) Depreciation & amortization', 'units')
BS_fill(33, lambda c, p: f'=IS!{c}20', F_NUM)
lab(wsBS, 34, '  Ending = beginning + CapEx - D&A', None, italic=True)
BS_row(35, '  D&A % of beginning net PP&E (implied)', '%', italic=True)
for c in FC:
    put(wsBS, f'{c}35', f'={c}33/{c}31', F_PCT)

band(wsBS, 37, 'DEBT & INTEREST SCHEDULE (interest on average debt; draws are inputs - no revolver)')
BS_row(38, 'Beginning debt', 'units')
for c in FC:
    put(wsBS, f'{c}38', f'={pv(c)}40', F_NUM)
BS_row(39, '(+) Drawdown / (repayment)', 'units')
for c in FC:
    put(wsBS, f'{c}39', f'={ASM}{c}{A["draw"]}', F_NUM)
BS_row(40, 'Ending debt (row 17)', 'units', bold=True)
BS_fill(40, lambda c, p: (HIST_LINK('debt')(c, p) if c in HC else f'={c}38+{c}39'), F_NUM, bold=True)
BS_row(41, 'Average debt (memo)', 'units', italic=True)
for c in FC:
    put(wsBS, f'{c}41', f'=({c}38+{c}40)/2', F_NUM)
BS_row(42, 'Interest expense (avg debt x rate)', 'units')
BS_fill(42, lambda c, p: (HIST_LINK('int')(c, p) if c in HC else f'={c}41*{ASM}$C$30'), F_NUM)
BS_row(43, '  After-tax interest (memo)', 'units', italic=True)
BS_fill(43, lambda c, p: f'={c}42*(1-{ASM}$C$29)', F_NUM)
BS_row(44, '  Net debt', 'units', bold=True)
BS_fill(44, lambda c, p: f'={c}40-{c}7', F_NUM, bold=True)
BS_row(45, '  Net debt / EBITDA', 'x')
BS_fill(45, lambda c, p: f'={c}44/IS!{c}18', F_X)
BS_row(46, '  EBIT / interest coverage', 'x')
BS_fill(46, lambda c, p: f'=IF({c}42>0,IS!{c}21/{c}42,0)', F_X)
lab(wsBS, 47, '  Note: interest is charged on average book debt; drawdowns are inputs (Assumptions); cash earns no return in this version.', None, italic=True)
wsBS.freeze_panes = 'C6'

# =====================================================================================
# 6. CFS - CASH FLOW STATEMENT (indirect; closes to BS cash)
# =====================================================================================
wsC = WS['CFS']
title(wsC, 'CFS - Cash Flow Statement (indirect method)')
put(wsC, 'A2', '="All figures in "&Assumptions!C10&" | Active scenario: "&Assumptions!C5&" | Net change in cash rolls the BS cash balance"',
    kind='calc', italic=True)
band(wsC, 4, 'FISCAL YEAR')
for c in HC + FC:
    wsC[f'{c}4'].value = f'=IF({c}5<=Assumptions!$C$12,"FY"&RIGHT({c}5,2)&"A","FY"&RIGHT({c}5,2)&"E")'
    wsC[f'{c}4'].font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    wsC[f'{c}4'].alignment = Alignment(horizontal='right')
put(wsC, 'C5', '=Assumptions!$C$12-2', F_YR)
for c in (HC + FC)[1:]:
    put(wsC, f'{c}5', f'={pv(c)}5+1', F_YR)
lab(wsC, 5, 'Fiscal year (numeric)')
ALLCF = HC + FC


def C_row(row, label, unit=None, bold=False, italic=False):
    lab(wsC, row, label, unit, bold=bold, italic=italic)


def C_fill(row, cols, fn, fmt, bold=False):
    for c in cols:
        f = fn(c, pv(c))
        if f is not None:
            put(wsC, f'{c}{row}', f, fmt, bold=bold)


band(wsC, 6, 'CASH FLOW FROM OPERATIONS')
C_row(7, 'Net income (Adjusted)', 'units')
C_fill(7, ALLCF, lambda c, p: f'=IS!{c}30', F_NUM)
C_row(8, '(+) Depreciation & amortization', 'units')
C_fill(8, ALLCF, lambda c, p: f'=IS!{c}20', F_NUM)
C_row(9, '(+) Stock-based compensation (non-cash)', 'units')
C_fill(9, ALLCF, lambda c, p: f'=IS!{c}26', F_NUM)
C_row(10, '(-) Increase in net working capital', 'units')
C_fill(10, ALLCF[1:], lambda c, p: f'=-((BS!{c}8-BS!{p}8)+(BS!{c}9-BS!{p}9)+(BS!{c}10-BS!{p}10)-(BS!{c}15-BS!{p}15)-(BS!{c}16-BS!{p}16))', F_NUM)
C_row(11, 'Cash flow from operations', 'units', bold=True)
C_fill(11, ALLCF[1:], lambda c, p: f'=SUM({c}7:{c}10)', F_NUM, bold=True)
C_row(12, '(-) Capital expenditures', 'units')
C_fill(12, ALLCF, lambda c, p: f'=-BS!{c}32', F_NUM)
C_row(13, 'Free cash flow (pre-financing)', 'units', bold=True)
C_fill(13, ALLCF[1:], lambda c, p: f'={c}11+{c}12', F_NUM, bold=True)
band(wsC, 14, 'FINANCING & NET CHANGE IN CASH')
C_row(15, '(-) Dividends paid (payout input row 23)', 'units')
C_fill(15, ALLCF, lambda c, p: f'=-$C$23*{c}7', F_NUM)
C_row(16, '(+/-) Net borrowing / (repayment)', 'units')
C_fill(16, ALLCF[1:], lambda c, p: (f'=BS!{c}39' if c in FC else f'=BS!{c}40-BS!{p}40'), F_NUM)
C_row(21, '(-) Share repurchases (from DCF)', 'units')
C_fill(21, FC, lambda c, p: f'=-DCF_Valuation!{c}102', F_NUM)
C_row(17, 'Net change in cash', 'units', bold=True)
C_fill(17, ALLCF[1:], lambda c, p: f'={c}13+{c}15+{c}16' + (f'+{c}21' if c in FC else ''), F_NUM, bold=True)
C_row(18, 'Opening cash', 'units')
C_fill(18, ALLCF[1:], lambda c, p: (f'=BS!{p}7' if c == 'F' else f'={p}19'), F_NUM)
C_row(19, 'Closing cash (rolls the BS cash balance)', 'units', bold=True)
C_fill(19, ALLCF[1:], lambda c, p: f'={c}17+{c}18', F_NUM, bold=True)
C_row(20, 'Check: closing cash ties to BS cash (forecast years; buybacks omitted so history never ties)', None)
C_fill(20, FC, lambda c, p: f'=IF(ABS({c}19-BS!{c}7)<0.5,"OK","ERROR")', None, bold=True)
band(wsC, 22, 'INPUTS (this tab)')
lab(wsC, 23, 'Dividend payout ratio (of adjusted net income)', '% of NI')
put(wsC, 'C23', 0.10, F_PCT, kind='input', fill=YEL)
lab(wsC, 24, 'Note: historical dividends are estimated at the payout input (Yahoo does not supply a clean dividend history line here).', None, italic=True)
wsC.freeze_panes = 'C6'





# =====================================================================================
# 5. WACC
# =====================================================================================
wsW = WS['WACC']
title(wsW, 'WACC - Cost of Capital Build')
put(wsW, 'A2', '="Scenario columns differ via the risk-free rate and the scenario risk-premium adjustment on Macro_Scenarios | Live column follows "&Assumptions!C5',
    kind='calc', italic=True)
band(wsW, 4, '1. UNLEVERED PEER BETA TABLE (replace with 4-5 comparable listed peers - Hamada unlevering)')
for c, t in zip('ABCDEF', ['Peer company', 'Ticker', 'Raw (levered) beta', 'Debt / Equity', 'Marginal tax rate', 'Unlevered beta']):
    put(wsW, f'{c}5', t, kind='label', bold=True, align='left' if c in 'AB' else 'right')
peers = [('Peer A (illustrative)', 'PEERA', 1.05, 0.20, 0.25), ('Peer B (illustrative)', 'PEERB', 0.95, 0.15, 0.25),
         ('Peer C (illustrative)', 'PEERC', 1.15, 0.30, 0.25), ('Peer D (illustrative)', 'PEERD', 0.90, 0.10, 0.25),
         ('Peer E (illustrative)', 'PEERE', 1.10, 0.25, 0.25)]
for i, (nm, tk, b, de, tx) in enumerate(peers):
    r = 6 + i
    put(wsW, f'A{r}', nm, kind='input', fill=YEL)
    put(wsW, f'B{r}', tk, kind='input', fill=YEL)
    put(wsW, f'C{r}', b, '0.00', fill=YEL)
    put(wsW, f'D{r}', de, '0.00', fill=YEL)
    put(wsW, f'E{r}', tx, F_PCT, fill=YEL)
    put(wsW, f'F{r}', f'=C{r}/(1+(1-E{r})*D{r})', '0.00')
lab(wsW, 11, 'Median', bold=True)
for c, fm in zip('CDEF', ['0.00', '0.00', F_PCT, '0.00']):
    put(wsW, f'{c}11', f'=MEDIAN({c}6:{c}10)', fm, bold=True)
lab(wsW, 12, 'Mean', italic=True)
for c, fm in zip('CDEF', ['0.00', '0.00', F_PCT, '0.00']):
    put(wsW, f'{c}12', f'=AVERAGE({c}6:{c}10)', fm)

band(wsW, 14, '2. CAPITAL STRUCTURE WEIGHTS (MARKET VALUES)')
rows = [
    (15, 'Current share price', '=Assumptions!C17', F_PX), (16, 'Diluted shares (TSM, DCF_Valuation)', f'=DCF_Valuation!C{DIL_ROW}', F_NUM),
    (17, 'Market value of equity', '=C15*C16', F_NUM), (18, 'Market value of debt (book value proxy)', '=Assumptions!C24', F_NUM),
    (19, 'Market debt / equity', '=C18/C17', F_PCT), (20, 'Market debt / total capital (D/V)', '=C18/(C17+C18)', F_PCT),
]
for r, l, f, fm in rows:
    lab(wsW, r, l)
    put(wsW, f'C{r}', f, fm)
lab(wsW, 21, 'Target D/V override (leave blank to use market weights)')
put(wsW, 'C21', None, F_PCT, kind='input', fill=YEL)
for r, l, f, fm in [(22, 'Applied D/V', '=IF(ISNUMBER(C21),C21,C20)', F_PCT), (23, 'Applied E/V', '=1-C22', F_PCT),
                    (24, 'Applied D/E (used to re-lever beta)', '=C22/C23', F_PCT)]:
    lab(wsW, r, l, bold=(r == 22))
    put(wsW, f'C{r}', f, fm)

band(wsW, 26, '3. COST OF EQUITY INPUTS (CAPM) - common to all scenarios')
for r, l, v, fm in [(27, 'Equity risk premium (ERP)', 0.05, F_PCT2), (28, 'Country risk premium (if applicable)', 0.0, F_PCT2),
                    (29, 'Size / company-specific risk premium (optional)', 0.005, F_PCT2)]:
    lab(wsW, r, l)
    put(wsW, f'C{r}', v, fm, fill=YEL)
for r, l, f, fm in [(30, 'Marginal tax rate', '=Assumptions!C28', F_PCT), (31, 'Median unlevered peer beta', '=F11', '0.00'),
                    (32, 'Re-levered beta = Bu x [1 + (1 - T) x D/E]  (Hamada)', '=C31*(1+(1-C30)*C24)', '0.00')]:
    lab(wsW, r, l, bold=(r == 32))
    put(wsW, f'C{r}', f, fm)

band(wsW, 34, '4. COST OF DEBT (SYNTHETIC RATING FROM INTEREST COVERAGE)')
for r, l, f, fm in [(35, 'EBIT - last historical year', '=IS!E21', F_NUM),
                    (36, 'Interest expense - last historical year', '=BS!E42', F_NUM),
                    (37, 'Interest coverage (EBIT / interest; 999 if no debt)', '=IF(C36>0,C35/C36,999)', F_X),
                    (38, 'Synthetic credit rating', f'=INDEX($A${RATING_FIRST}:$A${RATING_LAST},MATCH(C37,$B${RATING_FIRST}:$B${RATING_LAST},1))', None),
                    (39, 'Credit spread over risk-free', f'=INDEX($C${RATING_FIRST}:$C${RATING_LAST},MATCH(C37,$B${RATING_FIRST}:$B${RATING_LAST},1))', F_PCT2)]:
    lab(wsW, r, l)
    put(wsW, f'C{r}', f, fm)

band(wsW, 41, '5. WACC BY SCENARIO (risk-free rate and scenario premium differ; everything else is common)', c2='N')
lab(wsW, 42, 'Scenario', bold=True)
for c, t in zip('CDEFGHIJKLMN', ['Risk-free', 'Beta x ERP', 'Country risk', 'Size / CSRP', 'Scenario adj.', 'Cost of equity',
                                  'Pre-tax Kd', 'Tax rate', 'After-tax Kd', 'E/V', 'D/V', 'WACC']):
    put(wsW, f'{c}42', t, kind='label', bold=True, align='right')
for i, s in enumerate(SK):
    r = 43 + i
    put(wsW, f'A{r}', f'=Assumptions!$N${5 + i}', None, color=BLACK)
    fs = {'C': (f'=Macro_Scenarios!F{M["rf"][s]}', F_PCT2), 'D': ('=$C$32*$C$27', F_PCT2), 'E': ('=$C$28', F_PCT2), 'F': ('=$C$29', F_PCT2),
          'G': (f'=Macro_Scenarios!$C${M["prem"][s]}', F_PCT2), 'H': (f'=SUM(C{r}:G{r})', F_PCT2), 'I': (f'=C{r}+$C$39', F_PCT2),
          'J': ('=$C$30', F_PCT), 'K': (f'=I{r}*(1-J{r})', F_PCT2), 'L': ('=$C$23', F_PCT), 'M': ('=$C$22', F_PCT),
          'N': (f'=H{r}*L{r}+K{r}*M{r}', F_PCT2)}
    for c, (f, fm) in fs.items():
        put(wsW, f'{c}{r}', f, fm, bold=(c in 'HN'))
LIVE_R = 43 + N
RT0 = LIVE_R + 3
assert RT0 + 2 == RATING_FIRST
assert f'WACC!$N${LIVE_R}' == WACC_LIVE
lab(wsW, LIVE_R, '\u25ba Live (active scenario)', bold=True)
for c in 'CDEFGHIJKLMN':
    put(wsW, f'{c}{LIVE_R}', f'=INDEX({c}43:{c}{42 + N},Assumptions!$C$6)', F_PCT if c in 'JLM' else F_PCT2, bold=True)


band(wsW, RT0, '6. SYNTHETIC RATING TABLE (illustrative large-cap bands - refresh from Damodaran / agency data)')
for c, t in zip('ABC', ['Rating', 'Coverage \u2265', 'Spread']):
    put(wsW, f'{c}{RT0 + 1}', t, kind='label', bold=True, align='left' if c == 'A' else 'right')
rt = [('D', -100000, 0.1400), ('C', 0.20, 0.1100), ('CC', 0.65, 0.0885), ('CCC', 0.80, 0.0820), ('B-', 1.25, 0.0515),
      ('B', 1.50, 0.0421), ('B+', 1.75, 0.0351), ('BB', 2.00, 0.0240), ('BB+', 2.25, 0.0200), ('BBB', 2.50, 0.0156),
      ('A-', 3.00, 0.0122), ('A', 4.25, 0.0108), ('A+', 5.50, 0.0098), ('AA', 6.50, 0.0078), ('AAA', 8.50, 0.0063)]
assert RT0 + 2 == RATING_FIRST
for i, (rg, lo, sp) in enumerate(rt):
    r = RT0 + 2 + i
    put(wsW, f'A{r}', rg, kind='input')
    put(wsW, f'B{r}', float(lo), '0.00', fill=YEL)
    put(wsW, f'C{r}', sp, F_PCT2, fill=YEL)
wsW.freeze_panes = 'B4'

# =====================================================================================
# 7. COMPS - PEER TRADING MULTIPLES (cross-check on the DCF)
# =====================================================================================
wsC2 = WS['Comps']
title(wsC2, 'Comps - Peer Trading Multiples (implied value from EV/EBITDA and P/E)')
put(wsC2, 'A2', '="Yahoo does not supply peer financials - peer columns are inputs (yellow). Subject row is fully linked. Replace Peer A-E."',
    kind='calc', italic=True)
band(wsC2, 4, 'PEER SET - trading metrics', c2='P')
hdr = ['Peer', 'Ticker', 'Price', 'Shares (mm)', 'Mkt cap', 'Net debt', 'EV', 'EBITDA FY25', 'EV/EBITDA',
       'EBITDA FY26E', 'EV/EBITDA FY26E', 'Net income FY25', 'P/E FY25', 'EV/Revenue FY25', 'Revenue FY25']
for c, t in zip(['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O'], hdr):
    put(wsC2, f'{c}5', t, kind='label', bold=True, align='left' if c in 'AB' else 'right')
sub_r = 6
lab(wsC2, sub_r, 'Subject company', bold=True)
put(wsC2, f'B{sub_r}', '=Assumptions!C8', None, color=BLACK)
put(wsC2, f'C{sub_r}', '=Assumptions!C17', F_PX)
put(wsC2, f'D{sub_r}', '=DCF_Valuation!$F$104', F_NUM)
put(wsC2, f'E{sub_r}', f'=C{sub_r}*D{sub_r}', F_NUM)
put(wsC2, f'F{sub_r}', '=Assumptions!C24-Assumptions!C25', F_NUM)
put(wsC2, f'G{sub_r}', f'=E{sub_r}+F{sub_r}', F_NUM)
put(wsC2, f'H{sub_r}', '=IS!E18', F_NUM)
put(wsC2, f'I{sub_r}', f'=G{sub_r}/H{sub_r}', F_X)
put(wsC2, f'J{sub_r}', '=IS!F18', F_NUM)
put(wsC2, f'K{sub_r}', f'=G{sub_r}/J{sub_r}', F_X)
put(wsC2, f'L{sub_r}', '=IS!E30', F_NUM)
put(wsC2, f'M{sub_r}', f'=E{sub_r}/L{sub_r}', F_X)
put(wsC2, f'N{sub_r}', f'=G{sub_r}/IS!E11', F_X)
put(wsC2, f'O{sub_r}', '=IS!E11', F_NUM)
peers = [('Peer A', 'PEERA'), ('Peer B', 'PEERB'), ('Peer C', 'PEERC'), ('Peer D', 'PEERD'), ('Peer E', 'PEERE')]
for i, (nm, tk) in enumerate(peers):
    r = sub_r + 1 + i
    lab(wsC2, r, nm + ' (REPLACE)')
    put(wsC2, f'B{r}', tk, kind='input', fill=YEL)
    for col, fmt, default in [('C', F_PX, 48.00), ('D', F_NUM, 600.0), ('F', F_NUM, 450.0), ('H', F_NUM, 900.0),
                              ('J', F_NUM, 990.0), ('L', F_NUM, 650.0), ('O', F_NUM, 10000.0)]:
        put(wsC2, f'{col}{r}', default, fmt, fill=YEL)
    put(wsC2, f'E{r}', f'=C{r}*D{r}', F_NUM)
    put(wsC2, f'G{r}', f'=E{r}+F{r}', F_NUM)
    put(wsC2, f'I{r}', f'=G{r}/H{r}', F_X)
    put(wsC2, f'K{r}', f'=G{r}/J{r}', F_X)
    put(wsC2, f'M{r}', f'=E{r}/L{r}', F_X)
    put(wsC2, f'N{r}', f'=G{r}/O{r}', F_X)
MED, MEAN = 12, 13
lab(wsC2, MED, 'Peer median (excl. subject)', bold=True)
lab(wsC2, MEAN, 'Peer mean (excl. subject)', italic=True)
for c in ['C','D','F','G','H','I','J','K','L','M','N','O']:
    put(wsC2, f'{c}{MED}', f'=MEDIAN({c}7:{c}11)', F_X if c in 'IKMN' else (F_PX if c=='C' else F_NUM), bold=True)
    put(wsC2, f'{c}{MEAN}', f'=AVERAGE({c}7:{c}11)', F_X if c in 'IKMN' else (F_PX if c=='C' else F_NUM))

band(wsC2, 15, 'IMPLIED VALUE PER SHARE FROM PEER MULTIPLES', c2='P')
for c, t in zip(['D','E'], ['Peer median', 'Peer mean']):
    put(wsC2, f'{c}16', t, kind='label', bold=True, align='right')
NC = '(Assumptions!$C$24-Assumptions!$C$25-Assumptions!$C$26-Assumptions!$C$27)'
lab(wsC2, 17, 'EV/EBITDA FY26E x peer multiple -> price', 'x')
put(wsC2, 'D17', f'=($K${MED}*IS!F18-{NC})/DCF_Valuation!$F$104', F_PX)
put(wsC2, 'E17', f'=($K${MEAN}*IS!F18-{NC})/DCF_Valuation!$F$104', F_PX)
lab(wsC2, 18, 'P/E FY25 x peer multiple -> price', 'x')
put(wsC2, 'D18', f'=$M${MED}*IS!E30/DCF_Valuation!$F$104', F_PX)
put(wsC2, 'E18', f'=$M${MEAN}*IS!E30/DCF_Valuation!$F$104', F_PX)
lab(wsC2, 19, 'Memo: peer FY26E EV/EBITDA multiple (median | mean)', 'x', italic=True)
put(wsC2, 'D19', f'=$K${MED}', F_X)
put(wsC2, 'E19', f'=$K${MEAN}', F_X)
lab(wsC2, 21, 'Implied price - average of the two methods', None, bold=True)
put(wsC2, 'D21', '=(D17+D18)/2', F_PX, bold=True)
put(wsC2, 'E21', '=(E17+E18)/2', F_PX, bold=True)
lab(wsC2, 22, 'Football-field range across comps methods', None, italic=True)
put(wsC2, 'D22', '=MIN(D17:E18)', F_PX)
put(wsC2, 'E22', '=MAX(D17:E18)', F_PX)
band(wsC2, 24, 'METHODOLOGY', c2='P')
notes_c = [
    'EV/EBITDA case: (peer FY26E multiple x subject FY26E EBITDA) = implied EV; bridge to equity through latest debt, cash, minority and preferred; divide by TSM diluted shares.',
    'P/E case: peer FY25 P/E x subject FY25 adjusted net income; divide by TSM diluted shares.',
    'Peer set is editable by design: replace Peer A-E placeholders with true listed peers (ticker, price, shares, net debt, FY25/FY26 EBITDA, FY25 NI, FY25 revenue).',
    'Subject row is fully linked to the rest of this model (Assumptions / IS / DCF_Valuation) so the tab always reflects the loaded company.',
]
for i, t in enumerate(notes_c):
    put(wsC2, f'A{25 + i}', f'{i + 1}. {t}', kind='label', italic=True)
wsC2.freeze_panes = 'B6'

# =====================================================================================
# 6. DCF_VALUATION
# =====================================================================================
wsD = WS['DCF_Valuation']
title(wsD, 'DCF_Valuation - Core Discounted Cash Flow & Equity Bridge')
put(wsD, 'A2', '="All figures in "&Assumptions!C10&" except per-share data | Active scenario: "&Assumptions!C5', kind='calc', italic=True)
band(wsD, 13, 'VALUATION SETTINGS')
for r, l, f, fm in [(14, 'Company', '=Assumptions!C7', None), (15, 'Valuation date', '=Assumptions!C11', 'mmmm d, yyyy'),
                    (16, 'Active scenario', '=Assumptions!C5', None)]:
    lab(wsD, r, l)
    put(wsD, f'C{r}', f, fm)
lab(wsD, 17, 'Mid-year convention (Yes / No)', 'toggle')
put(wsD, 'C17', 'Yes', kind='input', fill=YEL)
dv_list(wsD, 'C17', ['Yes', 'No'])
lab(wsD, 18, 'Discount-period shift (0.5 if mid-year)', 'years')
put(wsD, 'C18', '=IF(C17="Yes",0.5,0)', '0.0')
for r, l, f, fm in [(19, 'WACC (live scenario)', f'={WACC_LIVE}', F_PCT2), (20, 'Terminal growth rate (g)', '=Assumptions!C33', F_PCT2),
                    (21, 'Exit EV / EBITDA multiple', '=Assumptions!C34', F_X), (22, 'SBC included in opex (GAAP-style)', 'Yes', None),
                    (23, 'Weight on Gordon Growth in blended value', '=Assumptions!C35', F_PCT)]:
    lab(wsD, r, l)
    put(wsD, f'C{r}', f, fm)

band(wsD, 26, 'UNLEVERED FREE CASH FLOW BUILD')
yr_labels(wsD, 26, 'Assumptions!$C$12')
OMr = 'IS!'
ufcf = [
    (27, 'Revenue', lambda c, p: f'={OMr}{c}11', F_NUM, ALL, False),
    (28, 'EBITDA (post-SBC)', lambda c, p: f'={OMr}{c}18', F_NUM, ALL, False),
    (29, '  EBITDA margin', lambda c, p: f'={c}28/{c}27', F_PCT, ALL, False),
    (30, 'EBIT', lambda c, p: f'={OMr}{c}21', F_NUM, ALL, False),
    (31, '(-) Cash taxes on EBIT', lambda c, p: f'=-{OMr}{c}23', F_NUM, ALL, False),
    (32, 'NOPAT', lambda c, p: f'={c}30+{c}31', F_NUM, ALL, True),
    (33, '(+) Depreciation & amortization', lambda c, p: f'={OMr}{c}20', F_NUM, ALL, False),
    (34, '(-) Capital expenditures', lambda c, p: f'=-BS!{c}32', F_NUM, ALL[:-1], False),
    (35, '(-) Increase in net working capital', lambda c, p: f'=-(BS!{c}26-BS!{p}26)', F_NUM, ALL[1:-1], False),
    (36, 'Memo: SBC already in EBITDA (no double-count)', lambda c, p: '=0', F_NUM, ALL, False),
    (37, 'Unlevered free cash flow (UFCF)', lambda c, p: f'=SUM({c}32:{c}36)', F_NUM, ALL[1:], True),
    (38, '  FCF conversion (UFCF / EBITDA)', lambda c, p: f'={c}37/{c}28', F_PCT, ALL[1:], False),
    (39, '  UFCF growth', lambda c, p: f'={c}37/{p}37-1', F_PCT, ALL[2:], False),
]
for r, l, fn, fm, cols, bd in ufcf:
    lab(wsD, r, l, 'units' if fm == F_NUM else '%', bold=bd)
    for c in cols:
        put(wsD, f'{c}{r}', fn(c, pv(c)), fm, bold=bd)
put(wsD, 'K34', '=-(K32*Assumptions!$C$33/Assumptions!$C$37)', F_NUM)  # terminal reinvestment = g / ROIC x NOPAT
put(wsD, 'K35', '=0', F_NUM)  # NWC build folded into ROIC-driven reinvestment above
assert DCF_UFCF_ROW == 37

band(wsD, 41, 'DISCOUNTING MECHANICS')
yr_labels(wsD, 41, 'Assumptions!$C$12', cols=FC)
lab(wsD, 42, 'Year index (t)')
put(wsD, 'F42', 1, '0')
for c in FC[1:]:
    put(wsD, f'{c}42', f'={pv(c)}42+1', '0')
lab(wsD, 43, 'Discount period (t - 0.5 if mid-year)', 'years')
lab(wsD, 44, 'Discount factor = 1 / (1 + WACC)^t')
lab(wsD, 45, 'Present value of UFCF', 'units', bold=True)
for c in FC:
    put(wsD, f'{c}43', f'={c}42-$C$18', '0.0')
    put(wsD, f'{c}44', f'=1/(1+$C$19)^{c}43', F_FAC)
    put(wsD, f'{c}45', f'={c}37*{c}44', F_NUM, bold=True)

band(wsD, 47, 'TERMINAL VALUE - METHOD A (GORDON GROWTH) vs. METHOD B (EXIT MULTIPLE)')
for c, t in zip('CD', ['Gordon', 'Exit multiple']):
    wsD[f'{c}47'].value = t
    wsD[f'{c}47'].alignment = Alignment(horizontal='right')
lab(wsD, 48, 'Terminal metric: UFCF FY31E (A) / EBITDA FY30E (B)', 'units')
put(wsD, 'C48', '=K37', F_NUM)
put(wsD, 'D48', '=J28', F_NUM)
lab(wsD, 49, 'Assumption: g (A) / exit multiple (B)')
put(wsD, 'C49', '=C20', F_PCT2)
put(wsD, 'D49', '=C21', F_X)
lab(wsD, 50, 'Terminal value at FY30E (undiscounted)', 'units', bold=True)
put(wsD, 'C50', '=C48/(C19-C20)', F_NUM, bold=True)
put(wsD, 'D50', '=D48*D49', F_NUM, bold=True)
lab(wsD, 51, 'Discount period for TV (Gordon: N - shift; Exit: N)', 'years')
put(wsD, 'C51', '=J43', '0.0')
put(wsD, 'D51', '=J42', '0.0')
lab(wsD, 52, 'Discount factor')
put(wsD, 'C52', '=1/(1+$C$19)^C51', F_FAC)
put(wsD, 'D52', '=1/(1+$C$19)^D51', F_FAC)
lab(wsD, 53, 'Present value of terminal value', 'units', bold=True)
put(wsD, 'C53', '=C50*C52', F_NUM, bold=True)
put(wsD, 'D53', '=D50*D52', F_NUM, bold=True)
lab(wsD, 54, 'Sanity: implied exit EV/EBITDA from Gordon TV', 'x')
put(wsD, 'C54', '=C50*(1+C19)^C18/D48', F_X)
lab(wsD, 55, 'Sanity: implied perpetuity growth from exit multiple (approx.)', '%')
put(wsD, 'D55', '=(D50/(1+C19)^C18*C19-J37)/(D50/(1+C19)^C18+J37)', F_PCT2)

band(wsD, 57, 'ENTERPRISE VALUE TO EQUITY VALUE BRIDGE')
for c, t in zip('CDE', ['Gordon', 'Exit multiple', 'Blended']):
    wsD[f'{c}57'].value = t
    wsD[f'{c}57'].alignment = Alignment(horizontal='right')
br = [
    (58, 'Sum of PV of explicit UFCF (FY26E-FY30E)', 'units', lambda c: '=SUM($F$45:$J$45)', F_NUM, False),
    (59, '(+) PV of terminal value', 'units', None, F_NUM, False),
    (60, 'Implied enterprise value (EV)', 'units', lambda c: f'={c}58+{c}59', F_NUM, True),
    (61, '(+) Cash & short-term investments', 'units', lambda c: '=Assumptions!$C$25', F_NUM, False),
    (62, '(-) Total debt', 'units', lambda c: '=-Assumptions!$C$24', F_NUM, False),
    (63, '(-) Preferred stock & minority interest', 'units', lambda c: '=-(Assumptions!$C$26+Assumptions!$C$27)', F_NUM, False),
    (64, 'Implied equity value', 'units', lambda c: f'=SUM({c}60:{c}63)', F_NUM, True),
    (65, '(\u00f7) Diluted shares outstanding (TSM)', 'mm', lambda c: '=$F$104', F_NUM, False),
    (66, 'Implied intrinsic value per share', CUR, lambda c: f'={c}64/{c}65', F_PX, True),
    (67, 'Current share price', CUR, lambda c: '=Assumptions!$C$17', F_PX, False),
    (68, 'Upside / (downside)', '%', lambda c: f'={c}66/{c}67-1', F_PCT, True),
    (69, 'Terminal value as % of EV', '%', lambda c: f'={c}59/{c}60', F_PCT, False),
    (70, 'Implied EV / EBITDA (last historical year)', 'x', lambda c: f'={c}60/{OMr}$E$18', F_X, False),
    (71, 'Implied EV / EBITDA (FY26E)', 'x', lambda c: f'={c}60/{OMr}$F$18', F_X, False),
    (72, 'Implied P/E (FY26E EPS)', 'x', lambda c: f'={c}66/$F$104', F_X, False),
]
for r, l, u, fn, fm, bd in br:
    lab(wsD, r, l, u, bold=bd)
    for c in 'CDE':
        if r == 59:
            f = {'C': '=C53', 'D': '=D53', 'E': '=$C$23*C53+(1-$C$23)*D53'}[c]
        else:
            f = fn(c)
        put(wsD, f'{c}{r}', f, fm, bold=bd)

band(wsD, 74, 'DILUTED SHARES - TREASURY STOCK METHOD (at current share price)')
for c, t in zip('CDEF', ['Number (mm)', 'Strike', 'In the money?', 'Net new shares']):
    wsD[f'{c}74'].value = t
    wsD[f'{c}74'].alignment = Alignment(horizontal='right')
lab(wsD, 75, 'Current share price used for TSM')
put(wsD, 'C75', '=Assumptions!C17', F_PX)
lab(wsD, 76, 'Basic shares outstanding', 'mm')
put(wsD, 'C76', '=Assumptions!C18', F_NUM)
lab(wsD, 77, 'RSUs / PSUs (fully dilutive)', 'mm')
put(wsD, 'C77', '=Assumptions!C20', F_NUM)
put(wsD, 'F77', '=C77', F_NUM)
for r, ar in [(78, 21), (79, 22), (80, 23)]:
    lab(wsD, r, f'Options tranche {r - 77}: n x (1 - strike / price) if in the money', 'mm')
    put(wsD, f'C{r}', f'=Assumptions!C{ar}', F_NUM)
    put(wsD, f'D{r}', f'=Assumptions!D{ar}', F_PX)
    put(wsD, f'E{r}', f'=IF($C$75>D{r},"Yes","No")', None, align='right')
    put(wsD, f'F{r}', f'=IF($C$75>D{r},C{r}*(1-D{r}/$C$75),0)', F_NUM)
lab(wsD, 81, 'Diluted shares outstanding (basic + TSM)', 'mm', bold=True)
put(wsD, 'C81', '=C76+SUM(F77:F80)', F_NUM, bold=True)
assert DIL_ROW == 81
# share repurchases & rolling diluted shares placed at rows 100-105 (rows 83-85 are the EPS memo)
band(wsD, 87, 'TERMINAL-YEAR CONVERGENCE CHECK (reinvestment vs. growth)')
cv = [
    (88, 'Terminal NOPAT (FY31E)', '=K32', F_NUM),
    (89, 'Net reinvestment = CapEx - D&A + increase in NWC', '=-K34-K33-K35', F_NUM),
    (90, 'Reinvestment rate (% of NOPAT)', '=C89/C88', F_PCT),
    (91, 'Implied RONIC = g / reinvestment rate', '=IF(C90>0,C20/C90,0)', F_PCT),
    (92, 'ROIC on existing capital (FY31E NOPAT / FY30E NWC + net PP&E)', f'=K32/(BS!J26+BS!J30)', F_PCT),
    (93, 'WACC', '=C19', F_PCT2),
    (94, 'RONIC - WACC spread', '=C91-C93', F_PCT2),
    (95, 'Tolerance (Assumptions)', '=Assumptions!C36', F_PCT2),
    (96, 'Convergence status', '=IF(AND(C90>0,C94<=C95),"OK","REVIEW")', None),
]
for r, l, f, fm in cv:
    lab(wsD, r, l, bold=(r == 96))
    put(wsD, f'C{r}', f, fm, bold=(r == 96), align='right' if r == 96 else None)

# =====================================================================================
# 7. SCENARIO_VALUATION
# =====================================================================================
wsS = WS['Scenario_Valuation']
title(wsS, f'Scenario_Valuation - {N}-Scenario Valuation Matrix (all cases live simultaneously)')
put(wsS, 'A2', '="All figures in "&Assumptions!C10&" except per-share data | Every scenario is fully re-computed in the blocks below the matrix - no data tables, no toggle flipping."',
    kind='calc', italic=True)
band(wsS, 4, 'VALUATION MATRIX - one row per scenario (probabilities are inputs; must total 100%)', c2='Q')
lab(wsS, 5, 'Scenario', bold=True)
for c, t in zip('CDEFGHIJKLMNOPQ', ['Probability', 'Revenue CAGR', 'FY30E EBITDA mgn', 'FY30E UFCF', 'WACC', 'EV - Gordon', 'EV - Exit',
                                    'EV - Blended', 'Equity value', 'Price - Gordon', 'Price - Exit', 'Price - Blended', 'Upside/(down)',
                                    'TV % of EV', 'Active']):
    put(wsS, f'{c}5', t, kind='label', bold=True, align='right')
SV_FIRST, SV_LAST = 6, 5 + N
SV_PW, SV_TOT = SV_LAST + 1, SV_LAST + 2
SVB = {s: SV_TOT + 3 + 29 * i for i, s in enumerate(SK)}
SV_END = SV_TOT + 3 + 29 * N


def sblock(i, s, b):
    subband(wsS, b, f'{SC[i].upper()} - CALCULATION BLOCK  (links to the {SC[i]} rows on Assumptions)', None)
    yr_labels(wsS, b, 'Assumptions!$C$12', cols=['E'] + FC + ['K'], band_row=False)
    names = ['Revenue growth', 'Revenue', 'Gross margin', 'EBITDA margin', 'EBITDA', 'D&A', 'EBIT', 'Cash taxes on EBIT', 'NOPAT',
             'Stock-based compensation', 'CapEx (maintenance + growth)', 'Net working capital', 'Increase in NWC', 'UFCF',
             'Discount factor', 'PV of UFCF']
    for k, nme in enumerate(names):
        lab(wsS, b + 1 + k, nme)
    for c in ['E'] + FC + ['K']:
        p = pv(c)
        isF, isK = c in FC, c == 'K'
        if isF:
            put(wsS, f'{c}{b + 1}', f'=Assumptions!{c}{A["rg"][s]}', F_PCT)
            put(wsS, f'{c}{b + 3}', f'=Assumptions!{c}{A["gm"][s]}', F_PCT)
            put(wsS, f'{c}{b + 4}', f'={c}{b + 3}-Assumptions!{c}{A["sga"][s]}-Assumptions!{c}{A["rd"][s]}-Assumptions!{c}{A["oth"][s]}-Assumptions!{c}{A["sbc"][s]}', F_PCT)
            put(wsS, f'{c}{b + 6}', f'={c}{b + 2}*Assumptions!{c}{A["da"][s]}', F_NUM)
            put(wsS, f'{c}{b + 10}', f'={c}{b + 2}*Assumptions!{c}{A["sbc"][s]}', F_NUM)
            put(wsS, f'{c}{b + 11}', f'={c}{b + 2}*(Assumptions!{c}{A["mcx"][s]}+Assumptions!{c}{A["gcx"][s]})', F_NUM)
            put(wsS, f'{c}{b + 12}', (f'={c}{b + 2}*(Assumptions!{c}{A["dso"][s]}/Assumptions!$C$14+(1-{c}{b + 3})*(Assumptions!{c}{A["dio"][s]}-Assumptions!{c}{A["dpo"][s]})/Assumptions!$C$14'
                                      f'+Assumptions!{c}{A["oca"][s]}-Assumptions!{c}{A["ocl"][s]})'), F_NUM)
            put(wsS, f'{c}{b + 15}', f'=1/(1+$C${b + 17})^DCF_Valuation!{c}$43', F_FAC)
            put(wsS, f'{c}{b + 16}', f'={c}{b + 14}*{c}{b + 15}', F_NUM)
        if isK:
            put(wsS, f'K{b + 1}', '=Assumptions!$C$33', F_PCT)
            put(wsS, f'K{b + 3}', f'=J{b + 3}', F_PCT)
            put(wsS, f'K{b + 4}', f'=J{b + 4}', F_PCT)
            for off in (6, 10, 11, 12):
                put(wsS, f'K{b + off}', f'=J{b + off}*(1+K{b + 1})', F_NUM)
        if c == 'E':
            put(wsS, f'E{b + 2}', f'={OMr}E11', F_NUM)
            put(wsS, f'E{b + 5}', f'={OMr}E18', F_NUM)
            put(wsS, f'E{b + 12}', f'=BS!E26', F_NUM)
            continue
        put(wsS, f'{c}{b + 2}', f'={p}{b + 2}*(1+{c}{b + 1})', F_NUM)
        put(wsS, f'{c}{b + 5}', f'={c}{b + 2}*{c}{b + 4}', F_NUM)
        put(wsS, f'{c}{b + 7}', f'={c}{b + 5}-{c}{b + 6}', F_NUM)
        put(wsS, f'{c}{b + 8}', f'={c}{b + 7}*Assumptions!$C$29', F_NUM)
        put(wsS, f'{c}{b + 9}', f'={c}{b + 7}-{c}{b + 8}', F_NUM)
        put(wsS, f'{c}{b + 13}', f'={c}{b + 12}-{p}{b + 12}', F_NUM)
        put(wsS, f'{c}{b + 14}', f'={c}{b + 9}+{c}{b + 6}-{c}{b + 11}-{c}{b + 13}', F_NUM, bold=True)
    sc = [
        (b + 17, 'WACC (WACC tab)', f'=WACC!$N${43 + i}', F_PCT2),
        (b + 18, 'Sum of PV of explicit UFCF', f'=SUM(F{b + 16}:J{b + 16})', F_NUM),
        (b + 19, 'Terminal value - Gordon growth', f'=K{b + 14}/(C{b + 17}-Assumptions!$C$33)', F_NUM),
        (b + 20, 'PV of terminal value - Gordon', f'=C{b + 19}/(1+C{b + 17})^DCF_Valuation!$J$43', F_NUM),
        (b + 21, 'Terminal value - Exit multiple', f'=J{b + 5}*Assumptions!$C$34', F_NUM),
        (b + 22, 'PV of terminal value - Exit multiple', f'=C{b + 21}/(1+C{b + 17})^DCF_Valuation!$J$42', F_NUM),
        (b + 23, 'Enterprise value - Gordon', f'=C{b + 18}+C{b + 20}', F_NUM),
        (b + 24, 'Enterprise value - Exit multiple', f'=C{b + 18}+C{b + 22}', F_NUM),
        (b + 25, 'Enterprise value - Blended', f'=Assumptions!$C$35*C{b + 23}+(1-Assumptions!$C$35)*C{b + 24}', F_NUM),
    ]
    for r, l, f, fm in sc:
        lab(wsS, r, l, bold=(r == b + 25))
        put(wsS, f'C{r}', f, fm, bold=(r == b + 25))


NC = '(DCF_Valuation!$C$61+DCF_Valuation!$C$62+DCF_Valuation!$C$63)'
for i, s in enumerate(SK):
    b = SVB[s]
    sblock(i, s, b)
    r = SV_FIRST + i
    put(wsS, f'A{r}', f'=Assumptions!$N${5 + i}', None, color=BLACK)
    put(wsS, f'C{r}', SCEN[i][4] / 100.0, F_PCT, fill=YEL)
    put(wsS, f'D{r}', f'=(J{b + 2}/E{b + 2})^(1/COUNT(F{b + 2}:J{b + 2}))-1', F_PCT)
    put(wsS, f'E{r}', f'=J{b + 4}', F_PCT)
    put(wsS, f'F{r}', f'=J{b + 14}', F_NUM)
    put(wsS, f'G{r}', f'=C{b + 17}', F_PCT2)
    put(wsS, f'H{r}', f'=C{b + 23}', F_NUM)
    put(wsS, f'I{r}', f'=C{b + 24}', F_NUM)
    put(wsS, f'J{r}', f'=C{b + 25}', F_NUM, bold=True)
    put(wsS, f'K{r}', f'=J{r}+{NC}', F_NUM)
    put(wsS, f'L{r}', f'=(H{r}+{NC})/DCF_Valuation!$F$104', F_PX)
    put(wsS, f'M{r}', f'=(I{r}+{NC})/DCF_Valuation!$F$104', F_PX)
    put(wsS, f'N{r}', f'=K{r}/DCF_Valuation!$F$104', F_PX, bold=True)
    put(wsS, f'O{r}', f'=N{r}/Assumptions!$C$17-1', F_PCT)
    put(wsS, f'P{r}', f'=1-C{b + 18}/C{b + 25}', F_PCT)
    put(wsS, f'Q{r}', f'=IF(Assumptions!$C$5=$A{r},"ACTIVE","")', None, bold=True, align='right')
lab(wsS, SV_PW, 'Probability-weighted expected value', bold=True)
for c, fm in zip('DEFGHIJKLMNP', [F_PCT, F_PCT, F_NUM, F_PCT2, F_NUM, F_NUM, F_NUM, F_NUM, F_PX, F_PX, F_PX, F_PCT]):
    put(wsS, f'{c}{SV_PW}', f'=SUMPRODUCT($C${SV_FIRST}:$C${SV_LAST},{c}{SV_FIRST}:{c}{SV_LAST})', fm, bold=True)
put(wsS, f'O{SV_PW}', f'=N{SV_PW}/Assumptions!$C$17-1', F_PCT, bold=True)
put(wsS, f'C{SV_PW}', f'=SUM(C{SV_FIRST}:C{SV_LAST})', F_PCT, bold=True)
lab(wsS, SV_TOT, 'Total probability (must equal 100%)')
put(wsS, f'C{SV_TOT}', f'=SUM(C{SV_FIRST}:C{SV_LAST})', F_PCT, bold=True)
lab(wsS, SV_TOT + 1, 'Current share price (colour-scale midpoint)')
put(wsS, f'C{SV_TOT + 1}', '=Assumptions!$C$17', F_PX)
wsS.conditional_formatting.add(f'N{SV_FIRST}:N{SV_LAST}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='formula',
                               mid_value=f'$C${SV_TOT + 1}', mid_color='FFEB84', end_type='max', end_color='63BE7B'))
wsS.freeze_panes = 'C6'

# =====================================================================================
# 8. SENSITIVITY_ANALYSIS
# =====================================================================================
wsT = WS['Sensitivity_Analysis']
title(wsT, 'Sensitivity_Analysis - Two-Way Valuation Sensitivities (closed-form, live in every cell)')
put(wsT, 'A2', '="Implied share price in "&Assumptions!C9&" | Colour scale: green = above current price, red = below, yellow = at current price ("&TEXT(Assumptions!C17,"0.00")&")"',
    kind='calc', italic=True)
band(wsT, 4, 'BASE INPUTS (LIVE LINKS - do not overwrite)', c2='P')
si = [
    (5, 'WACC (live scenario)', f'={WACC_LIVE}', F_PCT2), (6, 'Terminal growth rate (g)', '=Assumptions!C33', F_PCT2),
    (7, 'Exit EV / EBITDA multiple', '=Assumptions!C34', F_X), (8, 'Current share price', '=Assumptions!C17', F_PX),
    (9, 'Diluted shares (mm)', '=DCF_Valuation!$F$104', F_NUM),
    (10, 'Net claims (cash - debt - pref - minority)', '=DCF_Valuation!C61+DCF_Valuation!C62+DCF_Valuation!C63', F_NUM),
    (11, 'Mid-year shift (years)', '=DCF_Valuation!C18', '0.0'), (12, 'Number of explicit years (N)', '=DCF_Valuation!J42', '0'),
    (13, 'Weight on Gordon method (Table 3 blend)', '=Assumptions!C35', F_PCT), (14, 'FY30E EBITDA', '=IS!J18', F_NUM),
    (15, 'Terminal base flow A = terminal UFCF x (1+g) adjusted (ROIC-driven)', '=DCF_Valuation!K37/(1+Assumptions!$C$33)', F_NUM),
    (16, 'NWC build (folded into ROIC reinvestment)', '=0', F_NUM), (17, 'Effective tax rate', '=Assumptions!C29', F_PCT),
    (18, 'FY30E EBITDA margin', '=IS!J19', F_PCT),
]
for r, l, f, fm in si:
    lab(wsT, r, l)
    put(wsT, f'C{r}', f, fm)
for r, l, v, fm in [(19, 'Step: WACC', 0.005, F_PCT2), (20, 'Step: terminal growth', 0.005, F_PCT2), (21, 'Step: exit multiple', 1.0, F_X),
                    (22, 'Step: revenue growth spread (every forecast year)', 0.01, F_PCT), (23, 'Step: terminal EBITDA margin spread', 0.01, F_PCT)]:
    lab(wsT, r, l)
    put(wsT, f'C{r}', v, fm, fill=YEL)
lab(wsT, 25, 'UFCF FY26E-FY30E (live)')
lab(wsT, 26, 'Discount period (live)')
for c in FC:
    put(wsT, f'{c}25', f'=DCF_Valuation!{c}37', F_NUM)
    put(wsT, f'{c}26', f'=DCF_Valuation!{c}43', '0.0')
PVX = '(SUMPRODUCT($F$25:$J$25/(1+$B{r})^$F$26:$J$26))'
offs = [-2, -1, 0, 1, 2]
cols5 = ['C', 'D', 'E', 'F', 'G']


def axis(prefix_cell, step_cell, k, fmt):
    if k == 0:
        return f'={prefix_cell}'
    return f'={prefix_cell}{"+" if k > 0 else "-"}{abs(k)}*{step_cell}' if abs(k) != 1 else f'={prefix_cell}{"+" if k > 0 else "-"}{step_cell}'


# Table 1
band(wsT, 28, 'TABLE 1: WACC (rows) vs. TERMINAL GROWTH g (columns)  ->  Implied share price, Gordon Growth method', c2='P')
put(wsT, 'A29', 'WACC \u2193   |   g \u2192', kind='label', bold=True)
for c, k in zip(cols5, offs):
    put(wsT, f'{c}29', axis('$C$6', '$C$20', k, F_PCT2), F_PCT2, bold=True)
for i, k in enumerate(offs):
    r = 30 + i
    put(wsT, f'B{r}', axis('$C$5', '$C$19', k, F_PCT2), F_PCT2, bold=True)
    for c in cols5:
        put(wsT, f'{c}{r}', (f'=IF($B{r}>{c}$29,(SUMPRODUCT($F$25:$J$25/(1+$B{r})^$F$26:$J$26)'
                             f'+(($C$15*(1+{c}$29)-$C$16*{c}$29)/($B{r}-{c}$29))/(1+$B{r})^($C$12-$C$11)+$C$10)/$C$9,"n.m.")'), F_PX)
# Table 2
band(wsT, 36, 'TABLE 2: WACC (rows) vs. EXIT EV/EBITDA MULTIPLE (columns)  ->  Implied share price, Exit Multiple method', c2='P')
put(wsT, 'A37', 'WACC \u2193   |   Exit multiple \u2192', kind='label', bold=True)
for c, k in zip(cols5, offs):
    put(wsT, f'{c}37', axis('$C$7', '$C$21', k, F_X), F_X, bold=True)
for i, k in enumerate(offs):
    r = 38 + i
    put(wsT, f'B{r}', axis('$C$5', '$C$19', k, F_PCT2), F_PCT2, bold=True)
    for c in cols5:
        put(wsT, f'{c}{r}', (f'=(SUMPRODUCT($F$25:$J$25/(1+$B{r})^$F$26:$J$26)'
                             f'+$C$14*{c}$37/(1+$B{r})^$C$12+$C$10)/$C$9'), F_PX)
# Table 3
band(wsT, 44, 'TABLE 3: REVENUE GROWTH SPREAD (rows) vs. TERMINAL EBITDA MARGIN SPREAD (columns)  ->  Implied share price, Blended method', c2='P')
put(wsT, 'A45', 'Rev. growth spread \u2193   |   Margin spread \u2192', kind='label', bold=True)
for c, k in zip(cols5, offs):
    put(wsT, f'{c}45', f'={k}*$C$23', F_PCT, bold=True)
# helper section
band(wsT, 52, 'HELPER ROWS FOR TABLE 3 (revenue / NWC paths per growth spread; closed-form PV components)', c2='P')
yr_labels(wsT, 52, 'Assumptions!$C$12', cols=FC, band_row=True)
lab(wsT, 53, 'Discount factor (live WACC)')
lab(wsT, 54, 'After-tax cash margin before NWC: c_t = (m - da)(1-t) + da - capex% (post-SBC)')
lab(wsT, 55, 'Margin-spread ramp (t / N)')
lab(wsT, 56, 'NWC % of revenue (k_t)')
lab(wsT, 57, 'Live revenue growth')
for c in FC:
    put(wsT, f'{c}53', f'=1/(1+$C$5)^{c}26', F_FAC)
    put(wsT, f'{c}54', (f'=({OMr}{c}19-{OMr}{c}20/{OMr}{c}11)*(1-$C$17)+{OMr}{c}20/{OMr}{c}11-BS!{c}32/{OMr}{c}11'
                        f'-{OMr}{c}26/{OMr}{c}11*0'), '0.0000')
    put(wsT, f'{c}55', f'=DCF_Valuation!{c}42/$C$12', '0.00')
    put(wsT, f'{c}56', f'=BS!{c}26/{OMr}{c}11', F_PCT)
    put(wsT, f'{c}57', f'={OMr}{c}10', F_PCT)
for c, t in zip('LMNOP', ['PV c_t x Rev', 'PV ramp x Rev x (1-t)', 'PV of \u0394NWC', 'Revenue FY30E', 'NWC FY30E']):
    put(wsT, f'{c}59', t, kind='label', bold=True, align='right')
lab(wsT, 59, 'Per spread: (a) revenue path | (b) NWC path | (c) prior-year NWC', bold=True)
HLP = {}
for i, k in enumerate(offs):
    a = 60 + 4 * i
    b_, c_ = a + 1, a + 2
    HLP[i] = a
    tr = 46 + i
    put(wsT, f'B{tr}', f'={k}*$C$22', F_PCT, bold=True)
    lab(wsT, a, f'  Revenue path | spread {k:+d} x step')
    put(wsT, f'B{a}', f'=B{tr}', F_PCT)
    put(wsT, f'E{a}', f'={OMr}$E$11', F_NUM)
    lab(wsT, b_, '  NWC path')
    lab(wsT, c_, '  Prior-year NWC')
    for c in FC:
        p = pv(c)
        put(wsT, f'{c}{a}', f'={p}{a}*(1+{c}$57+$B{a})', F_NUM)
        put(wsT, f'{c}{b_}', f'={c}{a}*{c}$56', F_NUM)
        put(wsT, f'{c}{c_}', f'=BS!$E$26' if c == 'F' else f'={p}{b_}', F_NUM)
    put(wsT, f'L{a}', f'=SUMPRODUCT($F$53:$J$53,F{a}:J{a},$F$54:$J$54)', F_NUM)
    put(wsT, f'M{a}', f'=SUMPRODUCT($F$53:$J$53,F{a}:J{a},$F$55:$J$55)*(1-$C$17)', F_NUM)
    put(wsT, f'N{a}', f'=SUMPRODUCT($F$53:$J$53,F{b_}:J{b_}-F{c_}:J{c_})', F_NUM)
    put(wsT, f'O{a}', f'=J{a}', F_NUM)
    put(wsT, f'P{a}', f'=J{b_}', F_NUM)
    for c in cols5:
        d = f'{c}$45'
        put(wsT, f'{c}{tr}', (f'=($L{a}+{d}*$M{a}-$N{a}'
                              f'+$C$13*(($O{a}*($J$54+{d}*(1-$C$17))*(1+$C$6)-$P{a}*$C$6)/($C$5-$C$6))/(1+$C$5)^($C$12-$C$11)'
                              f'+(1-$C$13)*($O{a}*($C$18+{d})*$C$7)/(1+$C$5)^$C$12+$C$10)/$C$9'), F_PX)
for rng in ('C30:G34', 'C38:G42', 'C46:G50'):
    wsT.conditional_formatting.add(rng, ColorScaleRule(start_type='min', start_color='F8696B', mid_type='formula', mid_value='$C$8',
                                                       mid_color='FFEB84', end_type='max', end_color='63BE7B'))
wsT.freeze_panes = 'C4'
SENS_CENTER = {'t1': 'E32', 't2': 'E40', 't3': 'E48'}

# ---- DCF audit block (needs all sheets) -------------------------------------------------
band(wsD, 4, 'AUDIT / ERROR CHECK BLOCK')
err_formula = ('=SUMPRODUCT(--ISERROR(Assumptions!A1:K%d))+SUMPRODUCT(--ISERROR(Macro_Scenarios!A1:K%d))'
               '+SUMPRODUCT(--ISERROR(IS!A1:K40))+SUMPRODUCT(--ISERROR(BS!A1:K48))+SUMPRODUCT(--ISERROR(CFS!A1:K25))+SUMPRODUCT(--ISERROR(WACC!A1:N80))'
               '+SUMPRODUCT(--ISERROR(DCF_Valuation!A13:K96))+SUMPRODUCT(--ISERROR(Comps!A1:P45))+SUMPRODUCT(--ISERROR(Scenario_Valuation!A1:Q%d))'
               '+SUMPRODUCT(--ISERROR(Sensitivity_Analysis!A1:P82))') % (A['_end'] + 2, mr + 12, SV_END)
checks = [
    (5, 'No formula errors (REF / DIV0 / VALUE / NAME) anywhere in the model (error count)', '=IF(D5=0,"OK","ERROR")', err_formula, '0'),
    (6, 'Scenario probabilities sum to 100% (total)', '=IF(ABS(D6-1)<0.000001,"OK","ERROR")', f'=Scenario_Valuation!C{SV_TOT}', F_PCT),
    (7, 'WACC > terminal growth rate g (spread)', '=IF(C19>C20,"OK","ERROR")', '=C19-C20', F_PCT2),
    (8, 'Terminal-year convergence: implied RONIC vs. WACC (spread)', '=C96', '=C94', F_PCT2),
    (9, 'BS balance check = 0 in every historical / forecast year (A and FC)',
     '=IF(ABS(D9)<0.5,"OK","ERROR")', '=SUMPRODUCT(ABS(BS!C24:J24))', '0.0000'),
    (10, 'CFS closes tie to BS cash, every forecast year', '=IF(COUNTIF(CFS!F20:J20,"ERROR")=0,"OK","ERROR")', '=COUNTIF(CFS!F20:J20,"ERROR")', '0'),
    (11, 'Sensitivity Table 1 & 2 centres tie to DCF (max abs price difference)',
     '=IF(ABS(D11)<0.005,"OK","ERROR")',
     '=MAX(ABS(Sensitivity_Analysis!E32-C66),ABS(Sensitivity_Analysis!E40-D66))', '0.0000'),
]
for r, l, st, det, fm in checks:
    lab(wsD, r, l)
    put(wsD, f'C{r}', st, None, bold=True, align='right')
    put(wsD, f'D{r}', det, fm)
lab(wsD, 12, 'OVERALL MODEL STATUS', bold=True)
put(wsD, 'C12', '=IF(AND(C5="OK",C6="OK",C7="OK",C9="OK",C10="OK",C11="OK"),IF(C8="OK","ALL CHECKS PASSED","PASSED - REVIEW TERMINAL ROIC"),"CHECK ERRORS")',
    None, bold=True, align='right')
wsD.freeze_panes = 'C4'

# =====================================================================================
# SHARE REPURCHASES & ROLLING DILUTED SHARES (rows 100-105)
# =====================================================================================
band(wsD, 100, 'SHARE REPURCHASES & ROLLING DILUTED SHARE COUNT', c2='K')
lab(wsD, 101, 'Repurchase budget (% of UFCF; editable)', '%')
put(wsD, 'C101', 0.25, F_PCT, kind='input', fill=YEL)
lab(wsD, 102, 'Share repurchases ($)', 'units')
for c in FC:
    put(wsD, f'{c}102', f'=MAX(0,{c}37*$C$101)', F_NUM)
lab(wsD, 103, 'Shares repurchased (mm)', 'mm')
for c in FC:
    put(wsD, f'{c}103', f'={c}102/Assumptions!$C$17', F_NUM)
lab(wsD, 104, 'Rolling diluted shares (mm)', 'mm', bold=True)
for i, c in enumerate(FC):
    prev = f'$C${DIL_ROW}' if i == 0 else f'{pv(c)}104'
    put(wsD, f'{c}104', f'={prev}-{c}103', F_NUM, bold=True)
put(wsD, 'K104', '=J104', F_NUM, bold=True)
for c in FC:
    wsC[f'{c}21'] = f'=-DCF_Valuation!{c}102'

# =====================================================================================
# SUPPORTING TABS: PRECEDENTS / RISKS / CONTROLS / VERSION LOG
# =====================================================================================
wsP = WS['Precedents']
title(wsP, 'Precedents - M&A Multiples (editable; illustrative placeholders)')
put(wsP, 'A2', 'Replace placeholders with real precedent transactions (announced deals in the sector). Not fetched from Yahoo.', kind='calc', italic=True)
band(wsP, 4, 'PRECEDENT TRANSACTIONS', c2='K')
for c, t in zip('ABCDEF', ['Target', 'Acquirer', 'Announced', 'EV ($m)', 'EBITDA LTM ($m)', 'EV/EBITDA']):
    put(wsP, f'{c}5', t, kind='label', bold=True)
for i in range(8):
    r = 6 + i
    lab(wsP, r, f'  Transaction {i + 1} (replace)')
    for c in 'BC':
        put(wsP, f'{c}{r}', 'REPLACE', kind='input', fill=YEL)
    put(wsP, f'D{r}', 1000.0, F_NUM, kind='input', fill=YEL)
    put(wsP, f'E{r}', 100.0, F_NUM, kind='input', fill=YEL)
    put(wsP, f'F{r}', f'=D{r}/E{r}', F_X)
lab(wsP, 15, 'Median precedent EV/EBITDA', bold=True)
put(wsP, 'F15', '=MEDIAN(F6:F13)', F_X, bold=True)
lab(wsP, 16, 'Implied price from median precedent (EV/EBITDA x FY26E EBITDA)', None)
put(wsP, 'F16', '=(F15*IS!F18-(Assumptions!C24-Assumptions!C25-Assumptions!C26-Assumptions!C27))/DCF_Valuation!$F$104', F_PX, bold=True)
wsP.freeze_panes = 'C6'

wsR = WS['Risks']
title(wsR, 'Risks - Company-Specific Risk Register (seeded from dataset)')
put(wsR, 'A2', 'Seeded from ticker_full_dataset_v2.csv. Analyst adds probability, impact and mitigation; links to valuation via WACC/ROIC overrides.', kind='calc', italic=True)
band(wsR, 4, 'RISK REGISTER', c2='H')
for c, t in zip('ABCDEF', ['#', 'Risk', 'Likelihood (1-5)', 'Impact (1-5)', 'Mitigation / watch item', 'Valuation link']):
    put(wsR, f'{c}5', t, kind='label', bold=True)
for i in range(6):
    r = 6 + i
    put(wsR, f'A{r}', i + 1, '0', kind='label')
    put(wsR, f'B{r}', f'Risk {i + 1} (from dataset)', kind='input', fill=YEL)
    put(wsR, f'C{r}', 3, '0', kind='input', fill=YEL)
    put(wsR, f'D{r}', 3, '0', kind='input', fill=YEL)
    put(wsR, f'E{r}', 'REPLACE', kind='input', fill=YEL)
    put(wsR, f'F{r}', 'WACC / margin / growth', kind='label')
wsR.freeze_panes = 'C6'

wsK = WS['Controls']
title(wsK, 'Controls - Model Integrity Dashboard')
put(wsK, 'A2', '="Live checks from DCF_Valuation + BS / CFS / Segments. All must read OK before the model is circulated."', kind='calc', italic=True)
band(wsK, 4, 'CONTROL CHECKS', c2='H')
checks_ctrl = [
    (6, 'Formula errors (all tabs)', "=DCF_Valuation!C5"),
    (7, 'Scenario probabilities = 100%', "=DCF_Valuation!C6"),
    (8, 'WACC > g', "=DCF_Valuation!C7"),
    (9, 'Terminal ROIC vs WACC', "=DCF_Valuation!C8"),
    (10, 'BS balance = 0 (all years)', "=DCF_Valuation!C9"),
    (11, 'CFS closes to BS (all years)', "=DCF_Valuation!C10"),
    (12, 'Sensitivity ties to DCF', "=DCF_Valuation!C11"),
    (13, 'Segment mix = 100% (all years)', '=IF(SUMPRODUCT(--(Segments!F19:J19="OK"))=5,"OK","ERROR")'),
    (14, 'Segment revenue ties to IS', '=IF(SUMPRODUCT(ABS(Segments!F38:J38))<0.5,"OK","ERROR")'),
    (15, 'Segment EBITDA ties to IS', '=IF(SUMPRODUCT(ABS(Segments!F48:J48))<0.5,"OK","ERROR")'),
    (16, 'Listing coverage semi-annual', '=IF(OR(CFS!F20="OK",IS!E32>0),"OK","ERROR")'),
]
for r, l, f in checks_ctrl:
    lab(wsK, r, l)
    put(wsK, f'C{r}', f, None, bold=True, align='right')
lab(wsK, 18, 'OVERALL CONTROL STATUS', bold=True)
put(wsK, 'C18', '=IF(COUNTIF(C6:C16,"ERROR")=0,"ALL CONTROLS PASSED","REVIEW CONTROLS")', None, bold=True, align='right')
wsK.freeze_panes = 'C6'

wsV = WS['Version_Log']
title(wsV, 'Version Log - Analyst Change History')
put(wsV, 'A2', 'Record every material change: date, analyst, what changed, reason/approval.', kind='calc', italic=True)
band(wsV, 4, 'CHANGE HISTORY', c2='H')
for c, t in zip('ABCD', ['Date', 'Analyst', 'Change made', 'Approval / source']):
    put(wsV, f'{c}5', t, kind='label', bold=True)
for i in range(12):
    r = 6 + i
    put(wsV, f'A{r}', '2026-11-17' if i == 0 else '', None, kind='input', fill=YEL)
    for c in 'BCD':
        put(wsV, f'{c}{r}', 'Initial model build' if (i == 0 and c == 'C') else '', kind='input', fill=YEL)
wsV.freeze_panes = 'C6'

# =====================================================================================
# BRIDGE - valuation bridge (tornado) with bar chart
# =====================================================================================
wsBr = WS['Bridge']
title(wsBr, 'Bridge - Valuation Sensitivity to Key Drivers')
put(wsBr, 'A2', '="Blended DCF value per share at +/- one step on each driver. Amber = base case."', kind='calc', italic=True)
band(wsBr, 4, 'VALUATION BRIDGE - IMPACT OF KEY DRIVERS (± step, blended value per share)', c2='H')
for c, t in zip('CDE', ['Downside', 'Base', 'Upside']):
    put(wsBr, f'{c}5', t, kind='label', bold=True, align='right')
D = 'DCF_Valuation!'
W = f'{D}$C$19'; G = f'{D}$C$20'; N = f'{D}$J$42'
pv_exp = f'(SUMPRODUCT({D}$F$37:$J$37/(1+{{w}})^{D}$F$43:$J$43))'
tv_g = f'({D}$K$37*(1+{{g}})/({{w}}-{{g}}))/(1+{{w}})^{N}'
tv_x = f'({D}$J$28*{{x}})/(1+{{w}})^{N}'
shares = f'{D}$F$104'
bridge = [
    (8, 'WACC ±100bps', f'=({pv_exp.format(w=f"({W}-0.01)")}+{tv_g.format(w=f"({W}-0.01)", g=G)}+{D}$C$61+{D}$C$62+{D}$C$63)/{shares}',
     f'=DCF_Valuation!$E$66', f'=({pv_exp.format(w=f"({W}+0.01)")}+{tv_g.format(w=f"({W}+0.01)", g=G)}+{D}$C$61+{D}$C$62+{D}$C$63)/{shares}'),
    (9, 'Terminal growth g ±50bps', f'=({pv_exp.format(w=W)}+{tv_g.format(w=W, g=f"({G}-0.005)")}+{D}$C$61+{D}$C$62+{D}$C$63)/{shares}',
     f'=DCF_Valuation!$E$66', f'=({pv_exp.format(w=W)}+{tv_g.format(w=W, g=f"({G}+0.005)")}+{D}$C$61+{D}$C$62+{D}$C$63)/{shares}'),
    (10, 'Exit multiple ±1.0x', f'=({pv_exp.format(w=W)}+{tv_x.format(w=W, x=f"({D}$C$21-1)")}+{D}$C$61+{D}$C$62+{D}$C$63)/{shares}',
     f'=DCF_Valuation!$E$66', f'=({pv_exp.format(w=W)}+{tv_x.format(w=W, x=f"({D}$C$21+1)")}+{D}$C$61+{D}$C$62+{D}$C$63)/{shares}'),
]
for r, l, lo, mid, hi in bridge:
    lab(wsBr, r, l)
    for col, v in zip('CDE', (lo, mid, hi)):
        put(wsBr, f'{col}{r}', v, F_PX)
for r in range(8, 11):
    put(wsBr, f'F{r}', f'=D{r}-C{r}', F_PX)
    put(wsBr, f'G{r}', f'=E{r}-D{r}', F_PX)

from openpyxl.chart import BarChart
barb = BarChart()
barb.type = 'bar'
barb.grouping = 'clustered'
for col, nm in [('C','Downside'),('E','Upside')]:
    barb.series.append(Series(Reference(wsBr, min_col=3 if col=='C' else 5, min_row=8, max_row=10), title=nm))
barb.set_categories(Reference(wsBr, min_col=1, min_row=8, max_row=10))
barb.y_axis.numFmt = '#,##0.00'
barb.title = 'Valuation Bridge: Downside vs Upside per Driver'
barb.title.overlay = False
barb.x_axis.delete = False
barb.y_axis.delete = False
barb.height = 9
barb.legend = None
wsBr.add_chart(barb, 'A13')
wsBr.freeze_panes = 'C6'
# =====================================================================================
# 1. COVER - EQUITY RESEARCH SUMMARY (sell-side format)
# =====================================================================================
wsB = WS['Cover']
put(wsB, 'A1', '=Assumptions!C7&" ("&Assumptions!C8&") - Equity Research"', kind='calc')
wsB['A1'].font = Font(name='Arial', size=14, bold=True, color=NAVY)
wsB.row_dimensions[1].height = 24
put(wsB, 'A2', '="All figures in "&Assumptions!C10&" except per-share data | Valuation date: "&TEXT(Assumptions!C11,"mmm d, yyyy")&" | Active scenario: "&Assumptions!C5&" | DCF (blended) + peer multiples cross-check"',
    kind='calc', italic=True)
SVn = 'Scenario_Valuation!'
r_Bear, r_Base, r_Bull, r_PW = SV_FIRST, SV_FIRST + 1, SV_FIRST + 2, SV_PW

band(wsB, 4, 'ANALYST FRAME & COMPANY SNAPSHOT', c2='H')
lab(wsB, 5, 'Analyst / desk (placeholder - replace)', 'name')
put(wsB, 'C5', 'Equity Research', kind='input', fill=YEL)
lab(wsB, 6, 'Buy threshold - implied upside >=', '%')
put(wsB, 'C6', 0.15, F_PCT, kind='input', fill=YEL)
lab(wsB, 7, 'Sell threshold - implied upside <=', '%')
put(wsB, 'C7', -0.15, F_PCT, kind='input', fill=YEL)
snap = [
    (9, 'Company', '=Assumptions!C7', None), (10, 'Ticker', '=Assumptions!C8', None),
    (11, 'Currency / units', '=Assumptions!C9&" / "&Assumptions!C10', None),
    (12, 'Valuation date', '=Assumptions!C11', 'mmmm d, yyyy'), (13, 'Current share price', '=Assumptions!C17', F_PX),
    (14, 'Market value of equity (diluted, TSM)', f'=Assumptions!C17*DCF_Valuation!C{DIL_ROW}', F_NUM),
    (15, 'Implied enterprise value (blended)', '=DCF_Valuation!E60', F_NUM),
    (16, 'Implied net debt / EBITDA (FY25)', '=(Assumptions!C24-Assumptions!C25)/IS!E18', F_X),
    (17, 'Implied EV / Revenue (FY26E)', '=DCF_Valuation!E60/IS!F11', F_X),
    (18, 'Net debt (latest actual)', '=Assumptions!C24-Assumptions!C25', F_NUM),
    (19, 'Diluted shares outstanding (TSM)', '=DCF_Valuation!$F$104', '#,##0.0 "mm"'),
]
for r, l, f, fm in snap:
    lab(wsB, r, l)
    put(wsB, f'C{r}', f, fm)

band(wsB, 21, 'DCF VALUATION SUMMARY (ACTIVE SCENARIO)', c2='H')
for c, t in zip('CDE', ['Gordon', 'Exit multiple', 'Blended']):
    put(wsB, f'{c}22', t, kind='label', bold=True, align='right')
vsum = [
    (23, 'Implied intrinsic value per share', lambda c: f'=DCF_Valuation!{c}66', F_PX, True),
    (24, 'Upside / (downside) vs. current price', lambda c: f'=DCF_Valuation!{c}68', F_PCT, True),
    (25, 'Terminal value as % of enterprise value', lambda c: f'=DCF_Valuation!{c}69', F_PCT, False),
    (26, 'Implied EV / EBITDA (FY26E)', lambda c: f'=DCF_Valuation!{c}71', F_X, False),
    (27, 'Implied P/E (FY26E)', lambda c: f'=DCF_Valuation!{c}72', F_X, False),
]
for r, l, fn, fm, bd in vsum:
    lab(wsB, r, l, CUR if fm == F_PX else ('%' if fm == F_PCT else 'x'), bold=bd)
    for c in 'CDE':
        put(wsB, f'{c}{r}', fn(c), fm, bold=bd)
lab(wsB, 28, 'Implied rating (blended upside vs. thresholds above)', None, bold=True)
put(wsB, 'C28', '=IF($E$24>=$C$6,"BUY",IF($E$24<=$C$7,"SELL","HOLD"))', None, bold=True, align='right')
lab(wsB, 29, 'WACC (live scenario)', '%')
put(wsB, 'C29', f'={WACC_LIVE}', F_PCT2)
lab(wsB, 30, 'Probability-weighted value per share (20 / 60 / 20)', None)
put(wsB, 'C30', f'={SVn}N{r_PW}', F_PX, bold=True)
lab(wsB, 31, '  Bear | Base | Bull blended prices', None, italic=True)
put(wsB, 'C31', f'=TEXT({SVn}N{r_Bear},"0.00")&"  |  "&TEXT({SVn}N{r_Base},"0.00")&"  |  "&TEXT({SVn}N{r_Bull},"0.00")', None)

band(wsB, 33, 'FOOTBALL FIELD - IMPLIED VALUE PER SHARE (low | mid | high)', c2='H')
for c, t in zip('CDEFGHI', ['Low', 'Mid', 'High', 'EV low', 'EV mid', 'EV high', 'Range']):
    put(wsB, f'{c}34', t, kind='label', bold=True, align='right')
lab(wsB, 34, 'Method', bold=True)
nc = 'DCF_Valuation!$C$61+DCF_Valuation!$C$62+DCF_Valuation!$C$63'
def ffc(col, lo, mid, hi, elo, emid, ehi):
    return [lo, mid, hi, elo, emid, ehi]
ffrows = [
    (35, 'DCF - Gordon Growth (Bear | Base | Bull)', 'L', 'H'),
    (36, 'DCF - Exit Multiple (Bear | Base | Bull)', 'M', 'I'),
    (37, 'DCF - Blended (Bear | Base | Bull)', 'N', 'J'),
]
for r, l, pc, ec in ffrows:
    lab(wsB, r, l)
    for col, v in zip('CDEFGH', ffc(None, f'=MIN({SVn}{pc}{r_Bear}:{SVn}{pc}{r_Bull})', f'={SVn}{pc}{r_Base}',
                                              f'=MAX({SVn}{pc}{r_Bear}:{SVn}{pc}{r_Bull})',
                                              f'=MIN({SVn}{ec}{r_Bear}:{SVn}{ec}{r_Bull})', f'={SVn}{ec}{r_Base}',
                                              f'=MAX({SVn}{ec}{r_Bear}:{SVn}{ec}{r_Bull})')[0:6]):
        put(wsB, f'{col}{r}', v, F_PX if col in 'CDE' else F_NUM)
lab(wsB, 38, 'Comps - peer multiples (min | median-avg | max)', None)
put(wsB, 'C38', '=Comps!D22', F_PX); put(wsB, 'D38', '=Comps!D21', F_PX); put(wsB, 'E38', '=Comps!E22', F_PX)
for col, pc in zip('FGH', 'CDE'):
    put(wsB, f'{col}38', f'={pc}38*DCF_Valuation!$C${DIL_ROW}-({nc})', F_NUM)
lab(wsB, 39, 'Sensitivity - WACC x terminal growth (grid min | centre | max)', None)
put(wsB, 'C39', '=MIN(Sensitivity_Analysis!C30:G34)', F_PX); put(wsB, 'D39', '=Sensitivity_Analysis!E32', F_PX); put(wsB, 'E39', '=MAX(Sensitivity_Analysis!C30:G34)', F_PX)
for col, pc in zip('FGH', 'CDE'):
    put(wsB, f'{col}39', f'={pc}39*DCF_Valuation!$C${DIL_ROW}-({nc})', F_NUM)
for r in range(35, 40):
    put(wsB, f'I{r}', f'=E{r}-C{r}', F_PX)
lab(wsB, 40, 'Current share price (reference)', None)
put(wsB, 'D40', '=Assumptions!C17', F_PX, bold=True)

band(wsB, 42, 'KEY KPIs (ACTIVE SCENARIO)', c2='H')
kp = [
    (43, '5-year revenue CAGR (last actual to FY30E)', f'=(IS!J11/IS!E11)^(1/5)-1', F_PCT),
    (44, 'EBITDA margin - last actual year', '=IS!E19', F_PCT), (45, 'EBITDA margin - FY30E', '=IS!J19', F_PCT),
    (46, 'EBITDA margin expansion (bps)', '=(C45-C44)*10000', '#,##0;(#,##0);"-"'),
    (47, 'FCF conversion (avg UFCF / EBITDA, FY26E-FY30E)', '=SUM(DCF_Valuation!F37:J37)/SUM(DCF_Valuation!F28:J28)', F_PCT),
    (48, 'FY30E UFCF', '=DCF_Valuation!J37', F_NUM), (49, 'WACC', f'={WACC_LIVE}', F_PCT2),
    (50, 'Terminal value as % of EV (blended)', '=DCF_Valuation!E69', F_PCT),
    (51, 'Diluted shares outstanding (TSM)', '=DCF_Valuation!$F$104', '#,##0.0 "mm"'),
]
for r, l, f, fm in kp:
    lab(wsB, r, l)
    put(wsB, f'C{r}', f, fm)

band(wsB, 53, 'MODEL INTEGRITY - AUDIT BLOCK', c2='H')
for r in range(54, 66):
    src_r = r - 49          # DCF_Valuation rows 5..16 for 54..65; only through row 12 mirror
    if src_r > 12:
        break
    put(wsB, f'A{r}', f'=DCF_Valuation!A{src_r}', None, bold=(src_r == 12), color=BLACK)
    put(wsB, f'C{r}', f'=DCF_Valuation!C{src_r}', None, bold=True, align='right')

# ---- chart data + charts
CD = 68
band(wsB, CD, 'CHART DATA (live links)', c2='H')
rY, rR, rM, rU, rP, rE, rT = CD + 1, CD + 2, CD + 3, CD + 4, CD + 5, CD + 6, CD + 7
for r, l, u in [(rY, 'Fiscal year', None), (rR, 'Revenue', 'units'), (rM, 'EBITDA margin', '%'), (rU, 'UFCF', 'units'),
                (rP, 'PV of UFCF', 'units'), (rE, 'PV of explicit-period UFCF (blended bridge)', None), (rT, 'PV of terminal value (blended bridge)', None)]:
    lab(wsB, r, l, u)
for c in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
    put(wsB, f'{c}{rY}', f'=IS!{c}4', None, bold=True, align='right')
    put(wsB, f'{c}{rR}', f'=IS!{c}11', F_NUM)
    put(wsB, f'{c}{rM}', f'=IS!{c}19', F_PCT)
for c in FC:
    put(wsB, f'{c}{rU}', f'=DCF_Valuation!{c}37', F_NUM)
    put(wsB, f'{c}{rP}', f'=DCF_Valuation!{c}45', F_NUM)
put(wsB, f'C{rE}', '=DCF_Valuation!E58', F_NUM)
put(wsB, f'C{rT}', '=DCF_Valuation!E59', F_NUM)

def chart_polish(ch, ttl, xt=None, yt=None, axes=True):
    ch.title = ttl
    ch.title.overlay = False
    ch.roundedCorners = False
    ch.width, ch.height = 22, 9
    if ch.legend is not None:
        ch.legend.position = 'b'
        ch.legend.overlay = False
    if axes:
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        if xt:
            ch.x_axis.title = xt
            ch.x_axis.title.overlay = False
        if yt:
            ch.y_axis.title = yt
            ch.y_axis.title.overlay = False

cats_all = Reference(wsB, min_col=3, max_col=10, min_row=rY)
bar = BarChart()
bar.type = 'col'
bar.series.append(Series(Reference(wsB, min_col=3, max_col=10, min_row=rR), title='Revenue'))
bar.set_categories(cats_all)
bar.y_axis.numFmt = '#,##0'
bar.y_axis.majorGridlines = None
line = LineChart()
line.series.append(Series(Reference(wsB, min_col=3, max_col=10, min_row=rM), title='EBITDA margin (RHS)'))
line.set_categories(cats_all)
line.y_axis.axId = 200
line.y_axis.numFmt = '0%'
line.y_axis.crosses = 'max'
line.y_axis.delete = False
from openpyxl.chart.marker import Marker
line.series[0].marker = Marker(symbol='circle')
chart_polish(bar, 'Revenue & EBITDA Margin Trajectory', None, 'Revenue')
bar += line
CH = rT + 3
wsB.add_chart(bar, f'A{CH}')

# football field (floating horizontal bars: low offset + range)
fb = BarChart()
fb.type = 'bar'
fb.grouping = 'stacked'
fb.overlap = 100
cats_ff = Reference(wsB, min_col=1, min_row=35, max_row=39)
s_low = Series(Reference(wsB, min_col=3, min_row=35, max_row=39), title='Low')
s_low.graphicalProperties.noFill = True
fb.series.append(s_low)
fb.series.append(Series(Reference(wsB, min_col=9, min_row=35, max_row=39), title='Range (low-high)'))
fb.set_categories(cats_ff)
fb.x_axis.title = 'Implied value per share'
fb.y_axis.delete = False
chart_polish(fb, 'Football Field - Implied Value per Share (Bear | Base | Bull and ranges)', None, 'Method')
fb.height = 11
fb.legend = None
wsB.add_chart(fb, f'A{CH + 20}')

pie = PieChart()
pie.add_data(Reference(wsB, min_col=3, min_row=rE, max_row=rT), titles_from_data=False)
pie.set_categories(Reference(wsB, min_col=1, min_row=rE, max_row=rT))
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showVal = False
pie.dataLabels.showCatName = False
pie.dataLabels.showSerName = False
pie.dataLabels.showLeaderLines = False
chart_polish(pie, 'Enterprise Value Composition: PV of Explicit UFCF vs. PV of Terminal Value', axes=False)
wsB.add_chart(pie, f'A{CH + 40}')

# scenario bar chart (3 cases)
T0, T1 = SV_FIRST, SV_FIRST + 3 - 1
b4 = BarChart()
b4.type = 'bar'
b4.series.append(Series(Reference(wsB, min_col=3, min_row=CD + 9, max_row=CD + 11), title='Blended value per share'))
srows = []
for i, nm in enumerate(['Bear', 'Base', 'Bull']):
    put(wsB, f'A{CD + 9 + i}', nm, kind='label')
    put(wsB, f'C{CD + 9 + i}', f'={SVn}N{SV_FIRST + i}', F_PX)
b4.set_categories(Reference(wsB, min_col=1, min_row=CD + 9, max_row=CD + 11))
b4.y_axis.numFmt = '#,##0.00'
b4.x_axis.scaling.orientation = 'maxMin'
chart_polish(b4, 'Blended DCF Value per Share: Bear | Base | Bull', None, 'Value per share')
b4.height = 9
b4.legend = None
wsB.add_chart(b4, f'A{CH + 60}')
wsB.freeze_panes = 'A4'

# =====================================================================================
# FORMATTING PASS: status colouring, auto-fit, grid, save
# =====================================================================================
ok = dict(fill=PatternFill(bgColor='C6EFCE', fill_type='solid'), font=Font(color='006100', bold=True))
bad = dict(fill=PatternFill(bgColor='FFC7CE', fill_type='solid'), font=Font(color='9C0006', bold=True))
warn = dict(fill=PatternFill(bgColor='FFEB9C', fill_type='solid'), font=Font(color='9C5700', bold=True))
status_ranges = [(WS['DCF_Valuation'], 'C5:C12'), (WS['DCF_Valuation'], 'C96'), (WS['Cover'], 'C54:C61'),
                 (WS['BS'], 'C25:J25'), (WS['CFS'], 'F20:J20')]
for ws, rng in status_ranges:
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"OK"'], **ok))
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"ALL CHECKS PASSED"'], **ok))
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"ERROR"'], **bad))
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"CHECK ERRORS"'], **bad))
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"REVIEW"'], **warn))
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"PASSED - REVIEW TERMINAL ROIC"'], **warn))


def autofit(ws, skip_rows=(1, 2, 3), pad=5):
    widths = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None or c.row in skip_rows or c.row in SECT[ws.title]:
                continue
            v = c.value
            if isinstance(v, str) and v.startswith('='):
                L = 11
            elif isinstance(v, (int, float)):
                L = 11
            elif isinstance(v, datetime.date):
                L = 18
            else:
                L = len(str(v))
                if c.column_letter != 'A':
                    L = min(L, 22)
            widths[c.column_letter] = max(widths.get(c.column_letter, 0), L)
    for col, w in widths.items():
        cap = 78 if col == 'A' else 26
        ws.column_dimensions[col].width = min(max(w + pad, 14 if col != 'B' else 12), cap)


from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties
for n in NAMES:
    ws = WS[n]
    autofit(ws)
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.4, bottom=0.4, header=0.2, footer=0.2)
    ws.oddFooter.left.text = '&8&K7F8C9BIllustrative model - not investment advice'
    ws.oddFooter.left.size = 8
    ws.oddFooter.left.color = '7F8C9B'
    ws.oddFooter.center.text = '&8&K1F3864' + n
    ws.oddFooter.center.size = 8
    ws.oddFooter.center.color = '1F3864'
    ws.oddFooter.right.text = '&8&K7F8C9BPage &P of &N'
    ws.oddFooter.right.size = 8
    ws.oddFooter.right.color = '7F8C9B'
    ws.oddHeader.right.text = '&8&K7F8C9B&D'
    ws.oddHeader.right.size = 8
    ws.oddHeader.right.color = '7F8C9B'
WS['Cover'].column_dimensions['A'].width = 64
WS['Cover'].column_dimensions['B'].width = 12
for n in ('IS', 'BS', 'CFS', 'DCF_Valuation', 'Assumptions', 'Macro_Scenarios', 'Scenario_Valuation'):
    WS[n].column_dimensions['A'].width = min(max(WS[n].column_dimensions['A'].width or 0, 50), 78)
for n, rows_ in [('Assumptions', range(1, 4))]:
    pass
wsA.freeze_panes = 'C4'
wsM.freeze_panes = 'C5'
# hide the top 3 rows (blank) on every sheet - view and print start at the first content row
from openpyxl.utils import get_column_letter
for _ws in WS.values():
    _ws['A1'].value = None
    _ws['A2'].value = None
    _ws['A3'].value = None
    for _r in (1, 2, 3):
        _ws.row_dimensions[_r].hidden = True
    if _ws.max_row >= 4:
        _ws.print_area = f'A4:{get_column_letter(_ws.max_column)}{_ws.max_row}'
# ---- input validation + named ranges (analyst control conventions) ----
from openpyxl.workbook.defined_name import DefinedName
for rng, lo, hi, msg in [
    ("Assumptions!$C$29", 0.05, 0.45, "Effective tax rate 5%-45%"),
    ("Assumptions!$C$33", 0.0, 0.05, "Terminal growth 0%-5%, must stay below WACC"),
    ("Assumptions!$C$34", 3.0, 30.0, "Exit EV/EBITDA 3x-30x"),
    ("Assumptions!$C$35", 0.0, 1.0, "Gordon weight 0%-100%"),
    ("Assumptions!$C$37", 0.0, 0.40, "Terminal ROIC 0%-40%"),
    ("DCF_Valuation!$C$101", 0.0, 1.0, "Repurchase budget 0%-100% of UFCF"),
]:
    sheet, addr = rng.split('!')
    dv = DataValidation(type='decimal', operator='between', formula1=str(lo), formula2=str(hi),
                        allow_blank=False, showErrorMessage=True, errorTitle='Input out of range', error=msg)
    WS[sheet.strip('$')].add_data_validation(dv)
    dv.add(addr.replace('$', ''))
for nm, ref in [('WACC_Live', WACC_LIVE), ('TerminalGrowth', 'Assumptions!$C$33'),
                ('ExitMultiple', 'Assumptions!$C$34'), ('SharePrice', 'Assumptions!$C$17'),
                ('DilutedShares', 'DCF_Valuation!$F$104'), ('BasePrice', 'DCF_Valuation!$E$66')]:
    wb.defined_names.add(DefinedName(nm, attr_text=ref))
wb.properties.title = 'DCF Valuation Template'
if globals().get('HOOK'):
    HOOK(wb, globals())   # lets the app pre-fill inputs for a ticker
wb.save(OUT)
print('saved', OUT)
