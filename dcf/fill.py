"""Writes the ticker's data + basic default assumptions into the workbook inputs (everything stays editable in Excel)."""
import datetime
from . import config as C

FC = ['F', 'G', 'H', 'I', 'J']


def make_hook(d, rf):
    def hook(wb, g):
        A, Mx, HR, SK = g['A'], g['M'], g['HR'], g['SK']      # row registries + scenario keys from the builder
        ad = g['apply_delta']                                  # Base path + scenario delta x shape
        ws, wm, ww, wd = wb['Assumptions'], wb['Macro_Scenarios'], wb['WACC'], wb['DCF_Valuation']
        h, rev = d['hist'], d['hist']['rev']
        # segment names + revenue mix from the curated dataset (CSV); equal-weight fallback
        wsg = wb['Segments']
        ref = d.get('ref') or {}
        segs = [s[:44] for s in (ref.get('b') or [])][:6]
        if not segs:
            segs = ['Total company']
        for i, nm in enumerate(segs):
            for row in (9 + i, 23 + i, 31 + i, 41 + i):
                wsg[f'A{row}'] = nm
            mix = round(1.0 / len(segs), 5)
            for j, c in enumerate('FGHIJ'):
                wsg[f'{c}{9 + i}'] = mix if i < len(segs) - 1 else round(1.0 - mix * (len(segs) - 1), 5)
                wsg[f'{c}{23 + i}'] = f'=IS!{c}19'   # default: company margin; analyst overrides per segment; total ties exactly
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

        def put_all(key, vec_by_s, lo=None, hi=None):         # write one vector per scenario into Assumptions rows
            for s in SK:
                for i, c in enumerate(FC):
                    v = vec_by_s[s][i]
                    v = max(v, lo) if lo is not None else v
                    v = min(v, hi) if hi is not None else v
                    ws[f'{c}{A[key][s]}'] = round(v, 5)

        def flat(key, v):
            put_all(key, {s: [v] * 5 for s in SK})

        def shaped(key, base_vec, lo=None, hi=None):          # company Base path + the 23-scenario deltas
            put_all(key, ad(key, base_vec), lo, hi)

        gm = 1 - ratio('cogs'); da = h['da'][2] / rev[2]; capex = h['capex'][2] / rev[2]
        shaped('gmb', [gm] * 5, 0.05, 0.95)
        shaped('sga', [ratio('sga')] * 5, 0.0); shaped('rd', [ratio('rd')] * 5, 0.0)
        flat('oth', ratio('oth')); flat('sbc', ratio('sbc'))
        has_cogs = h['cogs'][2] > 0
        dso = ratio('ar') * 365
        dio = ratio('inv', 'cogs') * 365 if has_cogs else 0.0
        dpo = ratio('ap', 'cogs') * 365 if has_cogs else 0.0
        shaped('dso', [dso] * 5, 0.0); shaped('dio', [dio] * 5, 0.0); shaped('dpo', [dpo] * 5, 0.0)
        flat('oca', ratio('oca')); flat('ocl', ratio('ocl')); flat('da', da)
        flat('mcx', max(da, 0.005))
        shaped('gcx', [max(capex - da, 0.005)] * 5, 0.002)
        for c in FC:
            ws[f'{c}{A["draw"]}'] = 0.0
        # revenue: fade historical CAGR toward FADE_TO_GROWTH; solve the market-share term so Base growth follows that path
        cagr = min(max((rev[2] / rev[0]) ** 0.5 - 1, -0.05), 0.20) if rev[0] > 0 else 0.05
        tam = [wm[f'{c}{Mx["tam"]["base"]}'].value for c in FC]
        cpi = [wm[f'{c}{Mx["cpi"]["base"]}'].value for c in FC]
        spr = ws[f'F{A["spread"]["base"]}'].value or 0.0
        base = [(1 + cagr + (C.FADE_TO_GROWTH - cagr) * i / 4.0) / (1 + cpi[i] + spr) - 1 - tam[i] for i in range(5)]
        shaped('share', base)
        used['Hist. revenue CAGR'] = round(cagr, 4)
        rfv = ad('rf', [rf] * 5)                              # risk-free anchored to the live 10Y yield (+ scenario shifts)
        for s in SK:
            for i, c in enumerate(FC):
                wm[f'{c}{Mx["rf"][s]}'] = round(rfv[s][i], 4)
        used['Risk-free (10Y)'] = round(rf, 4)
        own = (d['name'] + ' (subject co. - replace with real peers)', d['ticker'], d['beta'], d['de'])
        ref = d.get('ref') or {}
        peers = [p[0] for p in (ref.get('p') or [])][:5]
        pdata = d.get('peers_data') or []
        for i in range(5):                                   # curated peer tickers + live beta/D-E; fallback = company itself
            r = 6 + i
            nm = (ref.get('p') or [])[i][1] if i < len(ref.get('p') or []) else own[0]
            tk = peers[i] if i < len(peers) else own[1]
            pd_ = pdata[i] if i < len(pdata) else {}
            beta = pd_.get('beta') if pd_.get('ok') else own[2]
            de = pd_.get('de') if pd_.get('ok') else own[3]
            ww[f'A{r}'] = nm; ww[f'B{r}'] = tk
            ww[f'C{r}'] = round(beta or own[2], 2); ww[f'D{r}'] = round(de or own[3], 3)
            ww[f'E{r}'] = C.MARGINAL_TAX
        # Risks tab seeded from curated dataset
        wr = wb['Risks']
        for i, rtext in enumerate((ref.get('r') or [])[:6]):
            wr[f'B{6 + i}'] = rtext[:90]
        # Comps peer data seeded from curated tickers + live Yahoo snapshot multiples
        wc = wb['Comps']
        for i in range(5):
            r = 7 + i
            pd_ = pdata[i] if i < len(pdata) else {}
            if pd_.get('ok'):
                wc[f'B{r}'] = pd_['ticker']
                if pd_.get('price'): wc[f'C{r}'] = round(pd_['price'], 2)
                if pd_.get('shares'): wc[f'D{r}'] = round(pd_['shares'], 3)
                if pd_.get('net_debt'): wc[f'F{r}'] = round(pd_['net_debt'], 1)
                if pd_.get('ebitda'): wc[f'H{r}'] = round(pd_['ebitda'], 1)
                if pd_.get('ebitda'):
                    growth = pd_.get('growth') or 0.06
                    wc[f'J{r}'] = round(pd_['ebitda'] * (1 + growth), 1)
                if pd_.get('ni'): wc[f'L{r}'] = round(pd_['ni'], 1)
                if pd_.get('rev'): wc[f'O{r}'] = round(pd_['rev'], 1)
            elif i < len(peers):
                wc[f'B{7 + i}'] = peers[i]
        wb.calculation.fullCalcOnLoad = True
        d['_used'] = used
    return hook
