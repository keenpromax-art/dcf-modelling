"""Builds the workbook from the builder + ticker data, and (optionally) reads results via LibreOffice for an on-page preview."""
import os, shutil, subprocess, sys, tempfile
import pandas as pd
from openpyxl import load_workbook
from . import config as C
from .fill import make_hook

CORE = os.path.join(os.path.dirname(__file__), 'builder_core.py')


def build_workbook(d, rf, out_path):
    src = open(CORE, encoding='utf-8').read()
    ns = {'__name__': 'dcf_builder_core', 'OUT_PATH': out_path, 'HOOK': make_hook(d, rf),
          'CUR_OVERRIDE': C.CURRENCY_SYMBOLS.get(d['currency'], '')}
    old_argv, sys.argv = sys.argv, ['builder']
    try:
        exec(compile(src, CORE, 'exec'), ns)
    finally:
        sys.argv = old_argv
    return out_path


def have_libreoffice():
    return shutil.which('soffice') is not None


def _f(v):
    return float(v) if isinstance(v, (int, float)) else None


def preview(path):
    """Recalculate a copy with LibreOffice; return numeric results for charts/cards, or None if unavailable."""
    if not have_libreoffice():
        return None
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, 'in.xlsx'); shutil.copy(path, src)
    out = os.path.join(tmp, 'out'); os.makedirs(out)
    subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', out, src], capture_output=True, timeout=180)
    f = os.path.join(out, 'in.xlsx')
    if not os.path.exists(f):
        return None
    wb = load_workbook(f, data_only=True)
    sv, dcf, om, st_, asm = wb['Scenario_Valuation'], wb['DCF_Valuation'], wb['Operating_Model'], wb['Sensitivity_Analysis'], wb['Assumptions']
    spec = {'Revenue CAGR (5Y)': 9, 'FY30E EBITDA margin': 10, 'FY30E UFCF': 11, 'WACC': 12, 'Enterprise value': 16,
            'Equity value': 18, 'Implied share price': 22, 'Upside / (downside)': 24}
    scen = pd.DataFrame({k: [_f(sv.cell(r, c).value) for c in (3, 4, 5, 6)] for k, r in spec.items()},
                        index=['Bear', 'Base', 'Bull', 'Prob-weighted']).T
    g_axis = [_f(st_.cell(29, c).value) for c in range(3, 8)]
    w_axis = [_f(st_.cell(r, 2).value) for r in range(30, 35)]
    sens = [(w, g, _f(st_.cell(30 + i, 3 + j).value)) for i, w in enumerate(w_axis) for j, g in enumerate(g_axis)]
    years = [om.cell(4, c).value for c in range(3, 11)]
    trend = pd.DataFrame({'Year': years, 'Revenue': [_f(om.cell(11, c).value) for c in range(3, 11)],
                          'EBITDA margin': [_f(om.cell(19, c).value) for c in range(3, 11)],
                          'Forecast': [str(y).endswith('E') for y in years]})
    ufcf = pd.DataFrame({'Year': years[3:], 'UFCF': [_f(dcf.cell(37, c).value) for c in range(6, 11)],
                         'PV of UFCF': [_f(dcf.cell(45, c).value) for c in range(6, 11)]})
    audit = [(dcf.cell(r, 1).value, dcf.cell(r, 3).value) for r in range(5, 12)]
    return dict(scen=scen, sens=pd.DataFrame(sens, columns=['WACC', 'g', 'Price']), trend=trend, ufcf=ufcf,
                pv_explicit=_f(dcf['E58'].value), pv_tv=_f(dcf['E59'].value), tv_pct=_f(dcf['E69'].value),
                gordon=_f(dcf['C66'].value), exitm=_f(dcf['D66'].value), audit=audit, price=_f(asm['C17'].value),
                sens_center=_f(st_['E32'].value))
