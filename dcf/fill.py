"""Writes the ticker's data + basic default assumptions into the workbook inputs (everything stays editable in Excel)."""
import datetime
from . import config as C

FC = ['F', 'G', 'H', 'I', 'J']


def make_hook(d, rf):
    def hook(wb, g):
        A, Mx, HR = g['A'], g['M'], g['HR']          # row registries from the builder
        ws, wm, ww, wd = wb['Assumptions'], wb['Macro_Scenarios'], wb['WACC'], wb['DCF_Valuation']
        h, rev = d['hist'], d['hist']['rev']
        used = {}

        def setc(ref, val, label=None):
            ws[ref] = val
            if label:
                used[label] = val

        setc('C7', d['name'] + ' (auto-filled)'); setc('C8', d['ticker']); setc('C9', d['currency'])
        setc('C10', f"{d['currency']} Millions"); setc('C11', datetime.date.today()); setc('C12', d['last_fy'])
        setc('C17', round(d['price'], 2), 'Share price'); setc('C18', round(d['shares'], 3), 'Basic shares (mm)')
        for r in (20, 21, 22, 23):
            ws[f'C{r}'] = 0.0; ws[f'D{r}'] = 0.0           # options / RSUs: not reliably available -> enter in Excel
        setc('C24', round(d['debt'], 1), 'Total debt'); setc('C25', round(d['cash'], 1), 'Cash & securities')
        setc('C26', round(d['minority'], 1), 'Minority interest'); setc('C27', round(d['pref'], 1), 'Preferred stock')
        ws['C28'] = C.MARGINAL_TAX; setc('C29', round(d['eff_tax'], 3), 'Effective tax rate')
        ws['C30'] = round(min(max(h['int'][2] / h['debt'][2], 0.02), 0.10), 4) if h['debt'][2] > 0 else 0.05
        setc('C33', C.TERMINAL_GROWTH, 'Terminal growth')
        lo, hi = C.EXIT_MULTIPLE_CLIP
        mult = min(max(d['ev_ebitda'], lo), hi) if d['ev_ebitda'] > 0 else C.EXIT_MULTIPLE_FALLBACK
        setc('C34', round(mult, 1), 'Exit EV/EBITDA')
        setc('C35', C.GORDON_WEIGHT, 'Weight on Gordon Growth')
        for k, row in HR.items():
            for j, c in enumerate('CDE'):
                ws[f'{c}{row}'] = round(h[k][j], 2)
        wd['C17'] = C.MID_YEAR

        def ratio(k, den='rev'):
            return sum(h[k][j] / h[den][j] for j in (1, 2) if h[den][j]) / 2.0

        def setrow(key, vals):                              # (bear, base, bull), scalar or list of 5
            for s, v in zip(('bear', 'base', 'bull'), vals):
                for i, c in enumerate(FC):
                    ws[f'{c}{A[key][s]}'] = round(v[i] if isinstance(v, list) else v, 5)

        gm = 1 - ratio('cogs'); da = h['da'][2] / rev[2]; capex = h['capex'][2] / rev[2]
        setrow('gmb', (gm - 0.01, gm, gm + 0.01))
        setrow('sga', (ratio('sga') + 0.005, ratio('sga'), ratio('sga') - 0.005))
        setrow('rd', (ratio('rd') + 0.002, ratio('rd'), ratio('rd') - 0.001))
        setrow('oth', (ratio('oth'),) * 3); setrow('sbc', (ratio('sbc'),) * 3)
        has_cogs = h['cogs'][2] > 0
        dso = ratio('ar') * 365
        dio = ratio('inv', 'cogs') * 365 if has_cogs else 0.0
        dpo = ratio('ap', 'cogs') * 365 if has_cogs else 0.0
        setrow('dso', (dso + 3, dso, dso - 2)); setrow('dio', (dio + 5, dio, dio - 3)); setrow('dpo', (dpo - 3, dpo, dpo + 3))
        setrow('oca', (ratio('oca'),) * 3); setrow('ocl', (ratio('ocl'),) * 3)
        setrow('da', (da,) * 3)
        mcx, gcx = max(da, 0.005), max(capex - da, 0.005)
        setrow('mcx', (mcx,) * 3); setrow('gcx', (gcx * 0.7, gcx, gcx * 1.2))
        for c in FC:
            ws[f'{c}{A["draw"]}'] = 0.0
        # revenue: fade historical CAGR toward FADE_TO_GROWTH; solve the market-share term so Base growth follows that path
        cagr = min(max((rev[2] / rev[0]) ** 0.5 - 1, -0.05), 0.20) if rev[0] > 0 else 0.05
        tam = [wm[f'{c}{Mx["tam"]["base"]}'].value for c in FC]
        cpi = [wm[f'{c}{Mx["cpi"]["base"]}'].value for c in FC]
        spr = ws[f'F{A["spread"]["base"]}'].value or 0.0
        base = [(1 + cagr + (C.FADE_TO_GROWTH - cagr) * i / 4.0) / (1 + cpi[i] + spr) - 1 - tam[i] for i in range(5)]
        setrow('share', ([x - 0.015 for x in base], base, [x + 0.01 for x in base]))
        used['Hist. revenue CAGR'] = round(cagr, 4)
        for s, adj in zip(('bear', 'base', 'bull'), (0.007, 0.0, -0.005)):
            for c in FC:
                wm[f'{c}{Mx["rf"][s]}'] = round(rf + adj, 4)
        used['Risk-free (10Y)'] = round(rf, 4)
        own = (d['name'] + ' (subject co. - replace with real peers)', d['ticker'], d['beta'], d['de'])
        for i in range(5):                                   # placeholder peer set = the company itself
            r = 6 + i
            ww[f'A{r}'] = own[0]; ww[f'B{r}'] = own[1]; ww[f'C{r}'] = round(own[2], 2); ww[f'D{r}'] = round(own[3], 3); ww[f'E{r}'] = C.MARGINAL_TAX
        wb.calculation.fullCalcOnLoad = True
        d['_used'] = used
    return hook
