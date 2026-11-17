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
    sv, dcf, is_, bs, st_, asm, cover = (wb['Scenario_Valuation'], wb['DCF_Valuation'], wb['IS'], wb['BS'],
                                         wb['Sensitivity_Analysis'], wb['Assumptions'], wb['Cover'])
    names = []
    for r in range(5, 12):                                    # static scenario list in Assumptions!N5:N7 (3 cases)
        v = asm.cell(r, 14).value
        if v is None or str(v).strip() == '':
            break
        names.append(str(v).strip())
    if not names:                                             # fallback: cached formula results, then generic labels
        names = [str(sv.cell(r, 1).value).strip() for r in range(6, 9) if sv.cell(r, 1).value] or ['Bear', 'Base', 'Bull']
    names = names + ['Probability-weighted']
    cols = {'Probability': 'C', 'Revenue CAGR': 'D', 'FY30E EBITDA margin': 'E', 'FY30E UFCF': 'F', 'WACC': 'G', 'Enterprise value': 'J',
            'Equity value': 'K', 'Price (Gordon)': 'L', 'Price (Exit)': 'M', 'Price (blended)': 'N', 'Upside / (downside)': 'O', 'TV % of EV': 'P'}
    scen = pd.DataFrame({k: [_f(sv[f'{c}{r}'].value) for r in range(6, 6 + len(names))] for k, c in cols.items()}, index=names)
    g_axis = [_f(st_.cell(29, c).value) for c in range(3, 8)]
    w_axis = [_f(st_.cell(r, 2).value) for r in range(30, 35)]
    sens = [(w, g, _f(st_.cell(30 + i, 3 + j).value)) for i, w in enumerate(w_axis) for j, g in enumerate(g_axis)]
    years = [is_.cell(4, c).value for c in range(3, 11)]
    trend = pd.DataFrame({'Year': years, 'Revenue': [_f(is_.cell(11, c).value) for c in range(3, 11)],
                          'EBITDA margin': [_f(is_.cell(19, c).value) for c in range(3, 11)],
                          'Forecast': [str(y).endswith('E') for y in years]})
    bs_check = [_f(bs.cell(24, c).value) for c in range(3, 11)]
    ufcf = pd.DataFrame({'Year': years[3:], 'UFCF': [_f(dcf.cell(37, c).value) for c in range(6, 11)],
                         'PV of UFCF': [_f(dcf.cell(45, c).value) for c in range(6, 11)]})
    audit = [(dcf.cell(r, 1).value, dcf.cell(r, 3).value) for r in range(5, 13)]
    rating = cover.cell(28, 3).value
    price_target = _f(cover.cell(23, 5).value)
    return dict(scen=scen, sens=pd.DataFrame(sens, columns=['WACC', 'g', 'Price']), trend=trend, ufcf=ufcf,
                pv_explicit=_f(dcf['E58'].value), pv_tv=_f(dcf['E59'].value), tv_pct=_f(dcf['E69'].value),
                gordon=_f(dcf['C66'].value), exitm=_f(dcf['D66'].value), audit=audit, price=_f(asm['C17'].value),
                sens_center=_f(st_['E32'].value), rating=rating, price_target=price_target, bs_check=bs_check,
                recalc_path=f)   # use the recalculated copy for download: it carries cached formula values
