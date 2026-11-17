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

DCF_UFCF_ROW = 37      # row of UFCF on DCF_Valuation (cross-referenced by Operating_Model)
DIL_ROW = 81           # row of diluted shares on DCF_Valuation (cross-referenced by WACC)
WACC_LIVE = 'WACC!$N$66'   # live WACC cell (23 scenario rows 43-65, live row 66)


def pv(c):
    return get_column_letter(column_index_from_string(c) - 1)


wb = Workbook()
NAMES = ['Dashboard', 'Assumptions', 'Macro_Scenarios', 'Operating_Model', 'WACC',
         'DCF_Valuation', 'Scenario_Valuation', 'Sensitivity_Analysis']
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
# 3. MACRO_SCENARIOS  (blocks first - Assumptions links to these rows)
# =====================================================================================
# ------------------------------------------------------------------ SCENARIO LIBRARY (23 scenarios)
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


_S('bear', 'Bear', 'Weak growth, sticky inflation, higher rates, margin pressure', 'fadeslow', 8,
   gdp=-.017, cpi=.015, rf=.007, tam=-.020, comm=.030, fx=3, share=-.013, spread=-.010, gmb=-.015, sga=.004, rd=.002, dso=3, dio=5, dpo=-3, gcx=-.006, prem=.010)
_S('base', 'Base', 'Consensus-style path: moderate growth, inflation normalising', 'flat', 30)
_S('bull', 'Bull', 'Strong demand, disinflation, falling rates, operating leverage', 'fadeslow', 8,
   gdp=.010, cpi=-.008, rf=-.005, tam=.020, comm=-.015, fx=-1.5, share=.007, spread=.005, gmb=.010, sga=-.004, rd=-.001, dso=-2, dio=-3, dpo=2, gcx=.004, prem=-.005)
_S('s04', 'Soft Landing', 'Disinflation without recession; gradual rate cuts', 'flat', 7,
   gdp=.002, cpi=-.004, rf=-.005, tam=.005, comm=-.005, spread=.001, gmb=.003, sga=-.001, dso=-1, dio=-1, gcx=.002, prem=-.0025)
_S('s05', 'Mild Recession', 'Two-quarter contraction, demand dip then recovery', 'shock', 6,
   gdp=-.030, cpi=-.010, rf=-.010, tam=-.030, comm=-.020, fx=1, share=-.003, spread=-.010, gmb=-.010, sga=.004, rd=.001, dso=4, dio=8, dpo=-2, gcx=-.008, prem=.0075)
_S('s06', 'Deep Recession', 'GFC-style demand collapse, credit stress, slow recovery', 'shock', 2,
   gdp=-.060, cpi=-.015, rf=-.020, tam=-.070, comm=-.040, fx=3, share=-.008, spread=-.020, gmb=-.030, sga=.010, rd=.002, dso=10, dio=15, dpo=-5, gcx=-.015, prem=.020)
_S('s07', 'Stagflation', 'Low growth with high inflation and rising rates', 'fade', 3,
   gdp=-.020, cpi=.030, rf=.015, tam=-.020, comm=.040, fx=2, share=-.005, spread=-.015, gmb=-.015, sga=.004, dso=4, dio=8, dpo=-3, gcx=-.004, prem=.015)
_S('s08', 'Rate Shock (Hawkish Central Banks)', 'Aggressive tightening: higher discount rates, stronger home currency', 'flat', 3,
   gdp=-.010, cpi=.010, rf=.015, tam=-.010, comm=.005, fx=4, spread=-.005, gmb=-.005, sga=.002, dso=3, dio=3, dpo=-2, gcx=-.005, prem=.010)
_S('s09', 'Dovish Pivot (Rate Cuts)', 'Faster-than-expected easing supports demand and multiples', 'flat', 4,
   gdp=.005, rf=-.010, tam=.010, fx=-2, spread=.002, gmb=.002, dso=-1, gcx=.003, prem=-.005)
_S('s10', 'Oil / Commodity Price Shock', 'Energy and input-cost spike squeezes margins', 'shock', 2,
   gdp=-.010, cpi=.020, rf=.005, tam=-.005, comm=.060, spread=-.005, gmb=-.020, sga=.001, dso=2, dio=5, dpo=-2, gcx=-.002, prem=.0075)
_S('s11', 'Input-Cost Relief (Commodity Deflation)', 'Falling raw-material prices lift gross margin', 'flat', 3,
   gdp=.003, cpi=-.010, rf=-.003, tam=.003, comm=-.040, spread=-.003, gmb=.008, prem=-.0025)
_S('s12', 'Strong Home Currency (Translation Headwind)', 'Reporting currency appreciates; foreign revenue translates lower', 'flat', 2,
   cpi=-.003, comm=-.010, fx=8, share=-.002, spread=-.002, gmb=-.003)
_S('s13', 'Weak Home Currency (FX Tailwind)', 'Reporting currency depreciates; translation tailwind, imported inflation', 'flat', 2,
   cpi=.005, comm=.010, fx=-8, share=.002, spread=.002, gmb=.002, prem=.0025)
_S('s14', 'AI / Technology Productivity Boom', 'Productivity-led growth, SG&A leverage, higher R&D and CapEx', 'build', 3,
   gdp=.005, cpi=-.003, rf=.002, tam=.025, share=.010, spread=.003, gmb=.010, sga=-.008, rd=.003, dso=-2, dio=-4, gcx=.006, prem=-.005)
_S('s15', 'Supply-Chain Disruption', 'Shortages: scarcity pricing but higher costs and inventory build', 'shock', 2,
   gdp=-.010, cpi=.015, rf=.005, tam=-.010, comm=.030, share=-.004, spread=.005, gmb=-.015, sga=.002, dso=2, dio=15, dpo=-4, gcx=.002, prem=.0075)
_S('s16', 'Trade War / Tariffs', 'Deglobalisation: tariffs, weaker volumes, input-cost inflation', 'build', 3,
   gdp=-.010, cpi=.010, rf=.003, tam=-.015, comm=.015, fx=3, share=-.005, spread=-.003, gmb=-.012, sga=.002, dio=6, prem=.010)
_S('s17', 'Geopolitical Shock', 'Conflict-driven energy spike, risk-off, weaker demand', 'shock', 2,
   gdp=-.015, cpi=.020, rf=.005, tam=-.020, comm=.040, fx=3, share=-.003, spread=-.005, gmb=-.012, sga=.002, dso=3, dio=8, dpo=-3, gcx=-.004, prem=.015)
_S('s18', 'Market-Share Outperformer', 'Company-specific: gains share from competitors', 'build', 3,
   share=.015, spread=.002, gmb=.004, sga=-.003, rd=.001, gcx=.004, prem=-.0025)
_S('s19', 'Competitive Disruption (Share Loss)', 'New entrant / disruptor erodes share and margin', 'build', 2,
   share=-.020, spread=-.008, gmb=-.012, sga=.005, rd=.003, dso=2, dio=4, gcx=.001, prem=.010)
_S('s20', 'Pricing-Power Boom', 'Company-specific: sustained price increases stick', 'flat', 2,
   cpi=.005, spread=.015, gmb=.010, prem=-.0025)
_S('s21', 'Price War / Margin Squeeze', 'Industry discounting compresses price and margin', 'fade', 1,
   share=.004, spread=-.020, gmb=-.020, sga=.003, prem=.0075)
_S('s22', 'Cost-Efficiency Programme', 'Company-specific: restructuring lowers SG&A and working capital', 'build', 1,
   gmb=.006, sga=-.010, dso=-2, dio=-5, gcx=-.002, prem=-.0025)
_S('s23', 'Demand Boom / Reflation', 'Hot economy: strong volumes and pricing, higher rates', 'fade', 1,
   gdp=.020, cpi=.015, rf=.010, tam=.030, comm=.020, fx=-2, share=.003, spread=.010, gmb=.003, sga=-.003, dso=-1, dio=-3, gcx=.005, prem=.005)
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
title(wsM, f'Macro_Scenarios - Macroeconomic & Industry Assumptions ({N} scenarios: Bear / Base / Bull + 20 regimes)',
      'Blue = input | Green = cross-sheet link | Black = formula.  Rows marked "Live" switch with the scenario selector on Assumptions!C5 (library in Assumptions!N5:O27).')
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
put(wsA, 'A3', 'ILLUSTRATIVE DATA: all pre-loaded numbers are for a fictional "Example Co." - overwrite every blue cell (sources: 10-K/annual report, market data, consensus) before using.',
    kind='label', italic=True, color='C00000')
A = {}

band(wsA, 4, 'A. CONTROL PANEL & COMPANY INFORMATION')
lab(wsA, 5, 'Active scenario (pick from the scenario library in column N)', 'selector', bold=True)
put(wsA, 'C5', 'Base', kind='input', fill=YEL, bold=True)
_dv = DataValidation(type='list', formula1=f'$N$5:$N${4 + N}', allow_blank=False)
wsA.add_data_validation(_dv)
_dv.add('C5')
lab(wsA, 6, f'Scenario index (1..{N}: 1 = Bear, 2 = Base, 3 = Bull, 4+ = library)', 'calc')
put(wsA, 'C6', f'=IFERROR(MATCH(C5,$N$5:$N${4 + N},0),2)', '0')
put(wsA, 'M4', '#', kind='label', bold=True, align='right')
put(wsA, 'N4', f'SCENARIO LIBRARY ({N}) - selectable in C5', kind='label', bold=True)
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
]
HR = {}
for i, (k, label, vals) in enumerate(HIST):
    row = 39 + i
    HR[k] = row
    lab(wsA, row, label, 'units')
    for j, c in enumerate(HC):
        put(wsA, f'{c}{row}', float(vals[j]), F_NUM, fill=YEL)
put(wsA, 'A56', 'Note: cost lines exclude D&A and SBC (EBITDA shown is Adjusted / pre-SBC). SBC is deducted in UFCF when the switch in C13 = Yes.',
    kind='label', italic=True)

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
    f'Scenario library: {N} pre-built regimes (Assumptions!N5:O27). Each = Base path + shaped delta (flat / fade / shock / build) on macro, margins, working capital and CapEx; every number is an editable blue input.',
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
subband(wsM, mr, 'MACRO-TO-OPERATING TRANSMISSION (Live case - flows straight into Assumptions / Operating_Model)', None)
trans = [
    ('Nominal GDP growth = (1 + real GDP) x (1 + CPI) - 1', lambda c: f'=(1+{c}{M["gdp"]["live"]})*(1+{c}{M["cpi"]["live"]})-1', F_PCT),
    ('Real policy spread: risk-free rate less CPI', lambda c: f'={c}{M["rf"]["live"]}-{c}{M["cpi"]["live"]}', F_PCT),
    ('Volume growth = Industry TAM growth + market-share gain / (loss)', lambda c: f'=Operating_Model!{c}7', F_PCT),
    ('Price growth = CPI + pricing spread', lambda c: f'=Operating_Model!{c}8', F_PCT),
    ('FX translation impact on revenue', lambda c: f'=Operating_Model!{c}9', F_PCT),
    ('Resulting revenue growth', lambda c: f'=Operating_Model!{c}10', F_PCT),
    ('Gross-margin drag from input costs (sensitivity x (commodity - CPI))',
     lambda c: f'=-Assumptions!{SENS}*({c}{M["comm"]["live"]}-{c}{M["cpi"]["live"]})', F_PCT2),
    ('Resulting gross margin', lambda c: f'=Operating_Model!{c}14', F_PCT),
]
for i, (label, fn, fmt) in enumerate(trans):
    lab(wsM, mr + 1 + i, label)
    for c in FC:
        put(wsM, f'{c}{mr + 1 + i}', fn(c), fmt)

# =====================================================================================
# 4. OPERATING_MODEL
# =====================================================================================
wsO = WS['Operating_Model']
title(wsO, 'Operating_Model - Integrated Unlevered Operating Schedule (Live scenario)',
      None)
put(wsO, 'A2', '="All figures in "&Assumptions!C10&" | Active scenario: "&Assumptions!C5&" | Last column = terminal year (FY30E x (1 + g))"',
    kind='calc', italic=True)
band(wsO, 4, 'FISCAL YEAR')
for c in ALL:
    pass
# numeric year row 5 + label row 4
put(wsO, 'C5', '=Assumptions!$C$12-2', F_YR)
for c in ALL[1:]:
    put(wsO, f'{c}5', f'={pv(c)}5+1', F_YR)
lab(wsO, 5, 'Fiscal year (numeric)')
for c in ALL:
    cell = wsO[f'{c}4']
    cell.value = (f'=IF({c}5<=Assumptions!$C$12,"FY"&RIGHT({c}5,2)&"A","FY"&RIGHT({c}5,2)&"E")' if c != 'K'
                  else '="FY"&RIGHT(K5,2)&"E (TY)"')
    cell.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    cell.alignment = Alignment(horizontal='right')


def O_row(row, label, unit=None, bold=False, italic=False):
    lab(wsO, row, label, unit, bold=bold, italic=italic)


def O_fill(row, cols, fn, fmt, bold=False):
    for c in cols:
        f = fn(c, pv(c))
        if f is not None:
            put(wsO, f'{c}{row}', f, fmt, bold=bold)


ASM = 'Assumptions!'
HIST_LINK = lambda key: (lambda c, p: f'={ASM}{c}{HR[key]}')
LIVE = lambda key: (lambda c, p: f'={ASM}{c}{A[key]["live"]}')
GK = '$K$10'

band(wsO, 6, 'INCOME STATEMENT (TO NOPAT)')
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
for row, label, key, akey in [(15, '(-) SG&A', 'sga', 'sga'), (16, '(-) R&D', 'rd', 'rd'), (17, '(-) Other operating expenses', 'oth', 'oth')]:
    O_row(row, label, 'units')
    O_fill(row, HC, HIST_LINK(key), F_NUM)
    O_fill(row, FC, lambda c, p, ak=akey: f'={c}11*{ASM}{c}{A[ak]["live"]}', F_NUM)
O_row(18, 'EBITDA (Adjusted, pre-SBC)', 'units', bold=True)
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
O_row(26, 'Memo: stock-based compensation', 'units', italic=True)
O_fill(26, HC, HIST_LINK('sbc'), F_NUM)
O_fill(26, FC, lambda c, p: f'={c}11*{ASM}{c}{A["sbc"]["live"]}', F_NUM)
# terminal-year flows = FY30E x (1+g)
for row in (12, 15, 16, 17, 20, 26):
    put(wsO, f'K{row}', f'=J{row}*(1+{GK})', F_NUM)

band(wsO, 28, 'WORKING CAPITAL SCHEDULE')
wc = [(29, 'Accounts receivable', 'ar', 11, 'dso', True), (30, 'Inventory', 'inv', 12, 'dio', True),
      (31, 'Prepaid & other current assets', 'oca', 11, 'oca', False)]
for row, label, hk, base, ak, isdays in wc:
    O_row(row, label, 'units')
    O_fill(row, HC, HIST_LINK(hk), F_NUM)
    if isdays:
        O_fill(row, FC, lambda c, p, b=base, a_=ak: f'={c}{b}*{ASM}{c}{A[a_]["live"]}/{ASM}$C$14', F_NUM)
    else:
        O_fill(row, FC, lambda c, p, b=base, a_=ak: f'={c}{b}*{ASM}{c}{A[a_]["live"]}', F_NUM)
O_row(32, 'Accounts payable', 'units')
O_fill(32, HC, HIST_LINK('ap'), F_NUM)
O_fill(32, FC, lambda c, p: f'={c}12*{ASM}{c}{A["dpo"]["live"]}/{ASM}$C$14', F_NUM)
O_row(33, 'Accrued & other current liabilities', 'units')
O_fill(33, HC, HIST_LINK('ocl'), F_NUM)
O_fill(33, FC, lambda c, p: f'={c}11*{ASM}{c}{A["ocl"]["live"]}', F_NUM)
for row in (29, 30, 31, 32, 33):
    put(wsO, f'K{row}', f'=J{row}*(1+{GK})', F_NUM)
O_row(34, 'Net working capital (NWC)', 'units', bold=True)
O_fill(34, ALL, lambda c, p: f'={c}29+{c}30+{c}31-{c}32-{c}33', F_NUM, bold=True)
O_row(35, '  NWC % of revenue', '%', italic=True)
O_fill(35, ALL, lambda c, p: f'={c}34/{c}11', F_PCT)
O_row(36, 'Change in NWC (increase = cash outflow)', 'units', bold=True)
O_fill(36, ALL[1:], lambda c, p: f'={c}34-{p}34', F_NUM, bold=True)
O_row(37, '  Cash OUTFLOW: NWC build', 'units')
O_fill(37, ALL[1:], lambda c, p: f'=MAX({c}36,0)', F_NUM)
O_row(38, '  Cash INFLOW: NWC release', 'units')
O_fill(38, ALL[1:], lambda c, p: f'=MAX(-{c}36,0)', F_NUM)
O_row(39, '  Net cash flow impact of NWC', 'units')
O_fill(39, ALL[1:], lambda c, p: f'=-{c}36', F_NUM)

band(wsO, 41, 'PP&E AND CAPEX SCHEDULE')
O_row(42, 'Beginning net PP&E', 'units')
O_fill(42, ALL[1:], lambda c, p: f'={p}47', F_NUM)
O_row(43, '(+) Maintenance CapEx', 'units')
O_fill(43, FC, lambda c, p: f'={c}11*{ASM}{c}{A["mcx"]["live"]}', F_NUM)
put(wsO, 'K43', f'=J43*(1+{GK})', F_NUM)
O_row(44, '(+) Growth CapEx', 'units')
O_fill(44, FC, lambda c, p: f'={c}11*{ASM}{c}{A["gcx"]["live"]}', F_NUM)
put(wsO, 'K44', f'=J44*(1+{GK})', F_NUM)
O_row(45, 'Total CapEx', 'units', bold=True)
O_fill(45, HC, HIST_LINK('capex'), F_NUM, bold=True)
O_fill(45, FC + ['K'], lambda c, p: f'={c}43+{c}44', F_NUM, bold=True)
O_row(46, '(-) Depreciation & amortization', 'units')
O_fill(46, ALL, lambda c, p: f'={c}20', F_NUM)
O_row(47, 'Ending net PP&E', 'units', bold=True)
O_fill(47, HC, HIST_LINK('ppe'), F_NUM, bold=True)
O_fill(47, FC + ['K'], lambda c, p: f'={c}42+{c}45-{c}46', F_NUM, bold=True)
O_row(48, '  CapEx % of revenue', '%', italic=True)
O_fill(48, ALL, lambda c, p: f'={c}45/{c}11', F_PCT)
O_row(49, '  Maintenance CapEx / D&A', 'x', italic=True)
O_fill(49, FC + ['K'], lambda c, p: f'={c}43/{c}46', F_X)
O_row(50, '  Check: total CapEx covers D&A', 'check', italic=True)
O_fill(50, FC + ['K'], lambda c, p: f'=IF({c}45>={c}46,"OK","REVIEW")', None)
O_row(51, '  D&A % of beginning net PP&E (implied)', '%', italic=True)
O_fill(51, ALL[1:], lambda c, p: f'={c}46/{c}42', F_PCT)

band(wsO, 53, 'SUPPLEMENTAL DEBT & INTEREST SCHEDULE (credit context only - not used in UFCF)')
O_row(54, 'Beginning debt', 'units')
O_fill(54, ALL[1:-1], lambda c, p: f'={p}56', F_NUM)
O_row(55, 'Debt drawdown / (repayment)', 'units')
O_fill(55, ['D', 'E'], lambda c, p: f'={c}56-{c}54', F_NUM)
O_fill(55, FC, lambda c, p: f'={ASM}{c}{A["draw"]}', F_NUM)
O_row(56, 'Ending debt', 'units', bold=True)
O_fill(56, HC, HIST_LINK('debt'), F_NUM, bold=True)
O_fill(56, FC, lambda c, p: f'={c}54+{c}55', F_NUM, bold=True)
O_row(57, 'Interest expense', 'units')
O_fill(57, HC, HIST_LINK('int'), F_NUM)
O_fill(57, FC, lambda c, p: f'={c}54*{ASM}$C$30', F_NUM)
O_row(58, 'Cash & marketable securities (prior + UFCF - after-tax interest + net borrowing)', 'units')
O_fill(58, HC, HIST_LINK('cash'), F_NUM)
O_fill(58, FC, lambda c, p: f'={p}58+DCF_Valuation!{c}{DCF_UFCF_ROW}-{c}57*(1-{ASM}$C$29)+{c}55', F_NUM)
O_row(59, 'Net debt', 'units', bold=True)
O_fill(59, ALL[:-1], lambda c, p: f'={c}56-{c}58', F_NUM, bold=True)
O_row(60, 'Net debt / EBITDA', 'x')
O_fill(60, ALL[:-1], lambda c, p: f'={c}59/{c}18', F_X)
O_row(61, 'EBIT / interest coverage', 'x')
O_fill(61, ALL[:-1], lambda c, p: f'=IF({c}57>0,{c}21/{c}57,0)', F_X)
O_row(62, 'EBITDA / interest coverage', 'x')
O_fill(62, ALL[:-1], lambda c, p: f'=IF({c}57>0,{c}18/{c}57,0)', F_X)
wsO.freeze_panes = 'C6'

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
for r, l, f, fm in [(35, 'EBIT - last historical year', '=Operating_Model!E21', F_NUM),
                    (36, 'Interest expense - last historical year', '=Operating_Model!E57', F_NUM),
                    (37, 'Interest coverage (EBIT / interest; 999 if no debt)', '=IF(C36>0,C35/C36,999)', F_X),
                    (38, 'Synthetic credit rating', '=INDEX($A$71:$A$85,MATCH(C37,$B$71:$B$85,1))', None),
                    (39, 'Credit spread over risk-free', '=INDEX($C$71:$C$85,MATCH(C37,$B$71:$B$85,1))', F_PCT2)]:
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
assert f'WACC!$N${LIVE_R}' == WACC_LIVE
lab(wsW, LIVE_R, '\u25ba Live (active scenario)', bold=True)
for c in 'CDEFGHIJKLMN':
    put(wsW, f'{c}{LIVE_R}', f'=INDEX({c}43:{c}{42 + N},Assumptions!$C$6)', F_PCT if c in 'JLM' else F_PCT2, bold=True)

RT0 = LIVE_R + 3
band(wsW, RT0, '6. SYNTHETIC RATING TABLE (illustrative large-cap bands - refresh from Damodaran / agency data)')
for c, t in zip('ABC', ['Rating', 'Coverage \u2265', 'Spread']):
    put(wsW, f'{c}{RT0 + 1}', t, kind='label', bold=True, align='left' if c == 'A' else 'right')
rt = [('D', -100000, 0.1400), ('C', 0.20, 0.1100), ('CC', 0.65, 0.0885), ('CCC', 0.80, 0.0820), ('B-', 1.25, 0.0515),
      ('B', 1.50, 0.0421), ('B+', 1.75, 0.0351), ('BB', 2.00, 0.0240), ('BB+', 2.25, 0.0200), ('BBB', 2.50, 0.0156),
      ('A-', 3.00, 0.0122), ('A', 4.25, 0.0108), ('A+', 5.50, 0.0098), ('AA', 6.50, 0.0078), ('AAA', 8.50, 0.0063)]
assert RT0 + 2 == 71
for i, (rg, lo, sp) in enumerate(rt):
    r = RT0 + 2 + i
    put(wsW, f'A{r}', rg, kind='input')
    put(wsW, f'B{r}', float(lo), '0.00', fill=YEL)
    put(wsW, f'C{r}', sp, F_PCT2, fill=YEL)
wsW.freeze_panes = 'B4'

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
                    (21, 'Exit EV / EBITDA multiple', '=Assumptions!C34', F_X), (22, 'SBC treated as cash cost?', '=Assumptions!C13', None),
                    (23, 'Weight on Gordon Growth in blended value', '=Assumptions!C35', F_PCT)]:
    lab(wsD, r, l)
    put(wsD, f'C{r}', f, fm)

band(wsD, 26, 'UNLEVERED FREE CASH FLOW BUILD')
yr_labels(wsD, 26, 'Assumptions!$C$12')
OMr = 'Operating_Model!'
ufcf = [
    (27, 'Revenue', lambda c, p: f'={OMr}{c}11', F_NUM, ALL, False),
    (28, 'EBITDA (Adjusted)', lambda c, p: f'={OMr}{c}18', F_NUM, ALL, False),
    (29, '  EBITDA margin', lambda c, p: f'={c}28/{c}27', F_PCT, ALL, False),
    (30, 'EBIT', lambda c, p: f'={OMr}{c}21', F_NUM, ALL, False),
    (31, '(-) Cash taxes on EBIT', lambda c, p: f'=-{OMr}{c}23', F_NUM, ALL, False),
    (32, 'NOPAT', lambda c, p: f'={c}30+{c}31', F_NUM, ALL, True),
    (33, '(+) Depreciation & amortization', lambda c, p: f'={OMr}{c}20', F_NUM, ALL, False),
    (34, '(-) Capital expenditures', lambda c, p: f'=-{OMr}{c}45', F_NUM, ALL, False),
    (35, '(-) Increase in net working capital', lambda c, p: f'=-{OMr}{c}36', F_NUM, ALL[1:], False),
    (36, '(-) Stock-based compensation (if cash cost)', lambda c, p: f'=-{OMr}{c}26*IF($C$22="Yes",1,0)', F_NUM, ALL, False),
    (37, 'Unlevered free cash flow (UFCF)', lambda c, p: f'=SUM({c}32:{c}36)', F_NUM, ALL[1:], True),
    (38, '  FCF conversion (UFCF / EBITDA)', lambda c, p: f'={c}37/{c}28', F_PCT, ALL[1:], False),
    (39, '  UFCF growth', lambda c, p: f'={c}37/{p}37-1', F_PCT, ALL[2:], False),
]
for r, l, fn, fm, cols, bd in ufcf:
    lab(wsD, r, l, 'units' if fm == F_NUM else '%', bold=bd)
    for c in cols:
        put(wsD, f'{c}{r}', fn(c, pv(c)), fm, bold=bd)
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
    (65, '(\u00f7) Diluted shares outstanding (TSM)', 'mm', lambda c: f'=$C${DIL_ROW}', F_NUM, False),
    (66, 'Implied intrinsic value per share', CUR, lambda c: f'={c}64/{c}65', F_PX, True),
    (67, 'Current share price', CUR, lambda c: '=Assumptions!$C$17', F_PX, False),
    (68, 'Upside / (downside)', '%', lambda c: f'={c}66/{c}67-1', F_PCT, True),
    (69, 'Terminal value as % of EV', '%', lambda c: f'={c}59/{c}60', F_PCT, False),
    (70, 'Implied EV / EBITDA (last historical year)', 'x', lambda c: f'={c}60/{OMr}$E$18', F_X, False),
    (71, 'Implied EV / EBITDA (FY26E)', 'x', lambda c: f'={c}60/{OMr}$F$18', F_X, False),
    (72, 'Implied P/E (FY26E EPS)', 'x', lambda c: f'={c}66/$F$85', F_X, False),
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
lab(wsD, 81, 'Diluted shares outstanding', 'mm', bold=True)
put(wsD, 'C81', '=C76+SUM(F77:F80)', F_NUM, bold=True)
assert DIL_ROW == 81

band(wsD, 83, 'MEMO: EARNINGS FOR IMPLIED P/E  [ (EBIT - SBC - interest) x (1 - tax) ]')
yr_labels(wsD, 83, 'Assumptions!$C$12', cols=['E', 'F'])
lab(wsD, 84, 'Net income (approximate)', 'units')
lab(wsD, 85, 'EPS on diluted shares', CUR)
for c in ('E', 'F'):
    put(wsD, f'{c}84', f'=({OMr}{c}21-{OMr}{c}26-{OMr}{c}57)*(1-Assumptions!$C$29)', F_NUM)
    put(wsD, f'{c}85', f'={c}84/$C$81', F_PX)

band(wsD, 87, 'TERMINAL-YEAR CONVERGENCE CHECK (reinvestment vs. growth)')
cv = [
    (88, 'Terminal NOPAT (FY31E)', '=K32', F_NUM),
    (89, 'Net reinvestment = CapEx - D&A + increase in NWC', '=-K34-K33-K35', F_NUM),
    (90, 'Reinvestment rate (% of NOPAT)', '=C89/C88', F_PCT),
    (91, 'Implied RONIC = g / reinvestment rate', '=IF(C90>0,C20/C90,0)', F_PCT),
    (92, 'ROIC on existing capital (FY31E NOPAT / FY30E NWC + net PP&E)', f'=K32/({OMr}J34+{OMr}J47)', F_PCT),
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
            put(wsS, f'{c}{b + 4}', f'={c}{b + 3}-Assumptions!{c}{A["sga"][s]}-Assumptions!{c}{A["rd"][s]}-Assumptions!{c}{A["oth"][s]}', F_PCT)
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
            put(wsS, f'E{b + 12}', f'={OMr}E34', F_NUM)
            continue
        put(wsS, f'{c}{b + 2}', f'={p}{b + 2}*(1+{c}{b + 1})', F_NUM)
        put(wsS, f'{c}{b + 5}', f'={c}{b + 2}*{c}{b + 4}', F_NUM)
        put(wsS, f'{c}{b + 7}', f'={c}{b + 5}-{c}{b + 6}', F_NUM)
        put(wsS, f'{c}{b + 8}', f'={c}{b + 7}*Assumptions!$C$29', F_NUM)
        put(wsS, f'{c}{b + 9}', f'={c}{b + 7}-{c}{b + 8}', F_NUM)
        put(wsS, f'{c}{b + 13}', f'={c}{b + 12}-{p}{b + 12}', F_NUM)
        put(wsS, f'{c}{b + 14}', f'={c}{b + 9}+{c}{b + 6}-{c}{b + 11}-{c}{b + 13}-{c}{b + 10}*IF(Assumptions!$C$13="Yes",1,0)', F_NUM, bold=True)
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
    put(wsS, f'L{r}', f'=(H{r}+{NC})/DCF_Valuation!$C${DIL_ROW}', F_PX)
    put(wsS, f'M{r}', f'=(I{r}+{NC})/DCF_Valuation!$C${DIL_ROW}', F_PX)
    put(wsS, f'N{r}', f'=K{r}/DCF_Valuation!$C${DIL_ROW}', F_PX, bold=True)
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
    (9, 'Diluted shares (mm)', f'=DCF_Valuation!C{DIL_ROW}', F_NUM),
    (10, 'Net claims (cash - debt - pref - minority)', '=DCF_Valuation!C61+DCF_Valuation!C62+DCF_Valuation!C63', F_NUM),
    (11, 'Mid-year shift (years)', '=DCF_Valuation!C18', '0.0'), (12, 'Number of explicit years (N)', '=DCF_Valuation!J42', '0'),
    (13, 'Weight on Gordon method (Table 3 blend)', '=Assumptions!C35', F_PCT), (14, 'FY30E EBITDA', '=Operating_Model!J18', F_NUM),
    (15, 'Terminal base flow A = NOPAT + D&A - CapEx - SBC (FY30E)', '=DCF_Valuation!J32+DCF_Valuation!J33+DCF_Valuation!J34+DCF_Valuation!J36', F_NUM),
    (16, 'FY30E net working capital', '=Operating_Model!J34', F_NUM), (17, 'Effective tax rate', '=Assumptions!C29', F_PCT),
    (18, 'FY30E EBITDA margin', '=Operating_Model!J19', F_PCT),
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
lab(wsT, 54, 'After-tax cash margin before NWC: c_t = (m - da)(1-t) + da - capex% - SBC%')
lab(wsT, 55, 'Margin-spread ramp (t / N)')
lab(wsT, 56, 'NWC % of revenue (k_t)')
lab(wsT, 57, 'Live revenue growth')
for c in FC:
    put(wsT, f'{c}53', f'=1/(1+$C$5)^{c}26', F_FAC)
    put(wsT, f'{c}54', (f'=({OMr}{c}19-{OMr}{c}20/{OMr}{c}11)*(1-$C$17)+{OMr}{c}20/{OMr}{c}11-{OMr}{c}45/{OMr}{c}11'
                        f'-{OMr}{c}26/{OMr}{c}11*IF(Assumptions!$C$13="Yes",1,0)'), '0.0000')
    put(wsT, f'{c}55', f'=DCF_Valuation!{c}42/$C$12', '0.00')
    put(wsT, f'{c}56', f'={OMr}{c}34/{OMr}{c}11', F_PCT)
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
        put(wsT, f'{c}{c_}', f'={OMr}$E$34' if c == 'F' else f'={p}{b_}', F_NUM)
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
               '+SUMPRODUCT(--ISERROR(Operating_Model!A1:K64))+SUMPRODUCT(--ISERROR(WACC!A1:N90))'
               '+SUMPRODUCT(--ISERROR(DCF_Valuation!A13:K96))+SUMPRODUCT(--ISERROR(Scenario_Valuation!A1:Q%d))'
               '+SUMPRODUCT(--ISERROR(Sensitivity_Analysis!A1:P82))') % (A['_end'] + 2, mr + 12, SV_END)
checks = [
    (5, 'No formula errors (REF / DIV0 / VALUE / NAME) anywhere in the model (error count)', '=IF(D5=0,"OK","ERROR")', err_formula, '0'),
    (6, 'Scenario probabilities sum to 100% (total)', '=IF(ABS(D6-1)<0.000001,"OK","ERROR")', f'=Scenario_Valuation!C{SV_TOT}', F_PCT),
    (7, 'WACC > terminal growth rate g (spread)', '=IF(C19>C20,"OK","ERROR")', '=C19-C20', F_PCT2),
    (8, 'Terminal-year convergence: implied RONIC vs. WACC (spread)', '=C96', '=C94', F_PCT2),
    (9, 'Scenario block ties to live DCF (price difference, active case)',
     '=IF(ABS(D9)<0.005,"OK","ERROR")', f'=E66-INDEX(Scenario_Valuation!N{SV_FIRST}:N{SV_LAST},Assumptions!C6)', '0.0000'),
    (10, 'Sensitivity centre cells tie to DCF (max abs price difference)',
     '=IF(ABS(D10)<0.005,"OK","ERROR")',
     '=MAX(ABS(Sensitivity_Analysis!E32-C66),ABS(Sensitivity_Analysis!E40-D66),ABS(Sensitivity_Analysis!E48-E66))', '0.0000'),
]
for r, l, st, det, fm in checks:
    lab(wsD, r, l)
    put(wsD, f'C{r}', st, None, bold=True, align='right')
    put(wsD, f'D{r}', det, fm)
lab(wsD, 11, 'OVERALL MODEL STATUS', bold=True)
put(wsD, 'C11', '=IF(AND(C5="OK",C6="OK",C7="OK",C9="OK",C10="OK"),IF(C8="OK","ALL CHECKS PASSED","PASSED - REVIEW TERMINAL ROIC"),"CHECK ERRORS")',
    None, bold=True, align='right')
wsD.freeze_panes = 'C4'

# =====================================================================================
# 1. DASHBOARD
# =====================================================================================
wsB = WS['Dashboard']
put(wsB, 'A1', '=Assumptions!C7&" ("&Assumptions!C8&") - DCF Valuation Dashboard"', kind='calc')
wsB['A1'].font = Font(name='Arial', size=14, bold=True, color=NAVY)
wsB.row_dimensions[1].height = 24
put(wsB, 'A2', '="All figures in "&Assumptions!C10&" except per-share data | Valuation date: "&TEXT(Assumptions!C11,"mmm d, yyyy")&" | Active scenario: "&Assumptions!C5',
    kind='calc', italic=True)
band(wsB, 4, 'MODEL INTEGRITY - AUDIT BLOCK', c2='H')
for r in range(5, 12):
    put(wsB, f'A{r}', f'=DCF_Valuation!A{r}', None, bold=(r == 11), color=BLACK)
    put(wsB, f'C{r}', f'=DCF_Valuation!C{r}', None, bold=True, align='right')
band(wsB, 13, 'VALUATION SUMMARY (ACTIVE SCENARIO)', c2='H')
hdr = [
    (14, 'Company', '=Assumptions!C7', None), (15, 'Ticker', '=Assumptions!C8', None),
    (16, 'Currency / units', '=Assumptions!C9&" / "&Assumptions!C10', None),
    (17, 'Valuation date', '=Assumptions!C11', 'mmmm d, yyyy'), (18, 'Active scenario', '=Assumptions!C5', None),
    (19, 'Current share price', '=Assumptions!C17', F_PX),
    (21, 'Implied upside / (downside) vs. current', '=DCF_Valuation!E68', F_PCT),
    (22, 'Implied enterprise value', '=DCF_Valuation!E60', F_NUM), (23, 'Implied equity value', '=DCF_Valuation!E64', F_NUM),
    (24, 'Implied EV / EBITDA (FY26E)', '=DCF_Valuation!E71', F_X), (25, 'Implied P/E (FY26E)', '=DCF_Valuation!E72', F_X),
    (26, 'Implied share price - Gordon Growth only', '=DCF_Valuation!C66', F_PX),
    (27, 'Implied share price - Exit Multiple only', '=DCF_Valuation!D66', F_PX),
]
for r, l, f, fm in hdr:
    lab(wsB, r, l)
    put(wsB, f'C{r}', f, fm)
put(wsB, 'A20', '="Implied intrinsic share price - "&Assumptions!C5&" case (blended)"', kind='calc', bold=True)
put(wsB, 'C20', '=DCF_Valuation!E66', F_PX, bold=True)
band(wsB, 29, 'KEY FINANCIAL & VALUATION KPIs', c2='H')
kp = [
    (30, '5-year revenue CAGR (last actual to FY30E)', f'=({OMr}J11/{OMr}E11)^(1/5)-1', F_PCT),
    (31, 'EBITDA margin - last actual year', f'={OMr}E19', F_PCT), (32, 'EBITDA margin - FY30E', f'={OMr}J19', F_PCT),
    (33, 'EBITDA margin expansion (bps)', '=(C32-C31)*10000', '#,##0;(#,##0);"-"'),
    (34, 'FCF conversion (avg UFCF / EBITDA, FY26E-FY30E)', '=SUM(DCF_Valuation!F37:J37)/SUM(DCF_Valuation!F28:J28)', F_PCT),
    (35, 'WACC', '=DCF_Valuation!C19', F_PCT2), (36, 'Terminal value as % of enterprise value (blended)', '=DCF_Valuation!E69', F_PCT),
    (37, 'FY30E UFCF', '=DCF_Valuation!J37', F_NUM),
]
for r, l, f, fm in kp:
    lab(wsB, r, l)
    put(wsB, f'C{r}', f, fm)

band(wsB, 39, 'VALUATION FOOTBALL FIELD SUMMARY (Bear / Base / Bull, all scenarios and sensitivity ranges)', c2='H')
for c, t in zip('CDEFGH', ['Price low', 'Price mid', 'Price high', 'EV low', 'EV mid', 'EV high']):
    put(wsB, f'{c}40', t, kind='label', bold=True, align='right')
SVn = 'Scenario_Valuation!'
nc = 'DCF_Valuation!$C$61+DCF_Valuation!$C$62+DCF_Valuation!$C$63'
r_bear, r_base, r_bull = SV_FIRST, SV_FIRST + 1, SV_FIRST + 2


def srng(col, a, b):
    return f'{SVn}{col}{a}:{col}{b}'


ffrows = [(41, 'DCF - Gordon Growth (Bear | Base | Bull)', 'L', 'H', r_bull), (42, 'DCF - Exit Multiple (Bear | Base | Bull)', 'M', 'I', r_bull),
          (43, 'DCF - Blended (Bear | Base | Bull)', 'N', 'J', r_bull),
          (44, f'DCF - Blended, all {N} scenarios (min | Base | max)', 'N', 'J', SV_LAST)]
for r, l, pc, ec, last in ffrows:
    lab(wsB, r, l)
    for col, fm_, v in zip('CDEFGH', [F_PX] * 3 + [F_NUM] * 3,
                           [f'=MIN({srng(pc, SV_FIRST, last)})', f'={SVn}{pc}{r_base}', f'=MAX({srng(pc, SV_FIRST, last)})',
                            f'=MIN({srng(ec, SV_FIRST, last)})', f'={SVn}{ec}{r_base}', f'=MAX({srng(ec, SV_FIRST, last)})']):
        put(wsB, f'{col}{r}', v, fm_)
sens_rows = [(45, 'Sensitivity: WACC x terminal growth', 'C30:G34', 'E32'), (46, 'Sensitivity: WACC x exit multiple', 'C38:G42', 'E40'),
             (47, 'Sensitivity: revenue growth x terminal margin', 'C46:G50', 'E48')]
for r, l, rng, ctr in sens_rows:
    lab(wsB, r, l)
    put(wsB, f'C{r}', f'=MIN(Sensitivity_Analysis!{rng})', F_PX)
    put(wsB, f'D{r}', f'=Sensitivity_Analysis!{ctr}', F_PX)
    put(wsB, f'E{r}', f'=MAX(Sensitivity_Analysis!{rng})', F_PX)
    for pc, ec in zip('CDE', 'FGH'):
        put(wsB, f'{ec}{r}', f'={pc}{r}*DCF_Valuation!$C${DIL_ROW}-({nc})', F_NUM)
lab(wsB, 48, 'Probability-weighted expected value (all scenarios)', bold=True)
put(wsB, 'D48', f'={SVn}N{SV_PW}', F_PX, bold=True)
put(wsB, 'G48', f'={SVn}J{SV_PW}', F_NUM, bold=True)
lab(wsB, 49, 'Current share price (reference line)')
put(wsB, 'D49', '=Assumptions!C17', F_PX)

# ---- all-scenario summary table
ST = 51
band(wsB, ST, f'ALL-SCENARIO SUMMARY ({N} scenarios, blended DCF)', c2='H')
for c, t in zip('CDEFG', ['Probability', 'Price', 'Upside/(down)', 'WACC', 'Enterprise value']):
    put(wsB, f'{c}{ST + 1}', t, kind='label', bold=True, align='right')
lab(wsB, ST + 1, 'Scenario', bold=True)
T0 = ST + 2
for i in range(N):
    r, sr = T0 + i, SV_FIRST + i
    put(wsB, f'A{r}', f'={SVn}A{sr}', None, color=BLACK)
    for c, sc_, fm in zip('CDEFG', ['C', 'N', 'O', 'G', 'J'], [F_PCT, F_PX, F_PCT, F_PCT2, F_NUM]):
        put(wsB, f'{c}{r}', f'={SVn}{sc_}{sr}', fm)
T1 = T0 + N - 1
wsB.conditional_formatting.add(f'D{T0}:D{T1}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='formula', mid_value='$D$49',
                                                             mid_color='FFEB84', end_type='max', end_color='63BE7B'))

# ---- chart data
CD = T1 + 2
band(wsB, CD, 'CHART DATA (live links)', c2='H')
rY, rR, rM, rU, rP, rE, rT = CD + 1, CD + 2, CD + 3, CD + 4, CD + 5, CD + 6, CD + 7
for r, l, u in [(rY, 'Fiscal year', None), (rR, 'Revenue', 'units'), (rM, 'EBITDA margin', '%'), (rU, 'UFCF', 'units'), (rP, 'PV of UFCF', 'units'),
                (rE, 'PV of explicit-period UFCF (blended bridge)', None), (rT, 'PV of terminal value (blended bridge)', None)]:
    lab(wsB, r, l, u)
for c in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
    put(wsB, f'{c}{rY}', f'={OMr}{c}4', None, bold=True, align='right')
    put(wsB, f'{c}{rR}', f'={OMr}{c}11', F_NUM)
    put(wsB, f'{c}{rM}', f'={OMr}{c}19', F_PCT)
    if c in FC:
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
chart_polish(bar, 'Revenue & EBITDA Margin Trajectory (FY23A-FY30E)', None, 'Revenue')
bar += line
CH = rT + 3
wsB.add_chart(bar, f'A{CH}')

cats_f = Reference(wsB, min_col=6, max_col=10, min_row=rY)
b2 = BarChart()
b2.type = 'col'
b2.series.append(Series(Reference(wsB, min_col=6, max_col=10, min_row=rU), title='UFCF'))
b2.series.append(Series(Reference(wsB, min_col=6, max_col=10, min_row=rP), title='PV of UFCF'))
b2.set_categories(cats_f)
b2.y_axis.numFmt = '#,##0'
chart_polish(b2, 'Unlevered Free Cash Flow Build & Present Value (Explicit Period)', None, 'Units')
wsB.add_chart(b2, f'A{CH + 20}')

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

b4 = BarChart()
b4.type = 'bar'
b4.series.append(Series(Reference(wsB, min_col=4, min_row=T0, max_row=T1), title='Implied value per share (blended)'))
b4.set_categories(Reference(wsB, min_col=1, min_row=T0, max_row=T1))
b4.y_axis.numFmt = '#,##0.00'
b4.x_axis.scaling.orientation = 'maxMin'
chart_polish(b4, f'Implied Value per Share by Scenario ({N} scenarios)', None, 'Value per share')
b4.height = 15
b4.legend = None
wsB.add_chart(b4, f'A{CH + 60}')
wsB.freeze_panes = 'A4'


# =====================================================================================
# FORMATTING PASS: status colouring, auto-fit, grid, save
# =====================================================================================
ok = dict(fill=PatternFill(bgColor='C6EFCE', fill_type='solid'), font=Font(color='006100', bold=True))
bad = dict(fill=PatternFill(bgColor='FFC7CE', fill_type='solid'), font=Font(color='9C0006', bold=True))
warn = dict(fill=PatternFill(bgColor='FFEB9C', fill_type='solid'), font=Font(color='9C5700', bold=True))
status_ranges = [(WS['DCF_Valuation'], 'C5:C11'), (WS['DCF_Valuation'], 'C96'), (WS['Dashboard'], 'C5:C11'),
                 (WS['Operating_Model'], 'F50:K50')]
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


for n in NAMES:
    ws = WS[n]
    autofit(ws)
    ws.sheet_view.showGridLines = False
WS['Dashboard'].column_dimensions['A'].width = 62
WS['Dashboard'].column_dimensions['B'].width = 12
for n in ('Operating_Model', 'DCF_Valuation', 'Assumptions', 'Macro_Scenarios', 'Scenario_Valuation'):
    WS[n].column_dimensions['A'].width = min(max(WS[n].column_dimensions['A'].width or 0, 50), 78)
for n, rows_ in [('Assumptions', range(1, 4))]:
    pass
wsA.freeze_panes = 'C4'
wsM.freeze_panes = 'C5'
# remove title + legend lines (rows 1-2) from every sheet
for _ws in WS.values():
    _ws['A1'].value = None
    _ws['A2'].value = None
    _ws.row_dimensions[1].height = None
wb.properties.title = 'DCF Valuation Template'
if globals().get('HOOK'):
    HOOK(wb, globals())   # lets the app pre-fill inputs for a ticker
wb.save(OUT)
print('saved', OUT)
