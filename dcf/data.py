"""Yahoo Finance (yfinance) extraction -> plain dict, all money in millions of reporting currency."""
import math
import pandas as pd
import yfinance as yf
from .config import DEFAULT_RF

M = 1e6
KEYS = ['rev', 'cogs', 'sga', 'rd', 'oth', 'da', 'sbc', 'capex', 'ar', 'inv', 'oca', 'ap', 'ocl', 'ppe', 'int', 'debt', 'cash', 'equity']


def _num(x):
    try:
        x = float(x)
        return 0.0 if math.isnan(x) else x
    except Exception:
        return 0.0


def _pick(df, names, col, default=0.0, agg=False):
    if df is None or df.empty or col is None or col not in df.columns:
        return default
    tot, hit = 0.0, False
    for n in names:
        if n in df.index and not pd.isna(df.loc[n, col]):
            if not agg:
                return float(df.loc[n, col])
            tot += float(df.loc[n, col])
            hit = True
    return tot if hit else default


def _by_year(df):
    return {c.year: c for c in df.columns} if df is not None and not df.empty else {}


def _safe_info(t):
    try:
        return t.info or {}
    except Exception:
        return {}


def _safe_df(fn):
    try:
        df = fn()
        return df if df is not None and not df.empty else None
    except Exception:
        return None


def fetch_company(ticker):
    t = yf.Ticker(ticker)
    info = _safe_info(t)
    fin, bs, cf = _safe_df(lambda: t.financials), _safe_df(lambda: t.balance_sheet), _safe_df(lambda: t.cashflow)
    fy, by, cy = _by_year(fin), _by_year(bs), _by_year(cf)
    years = sorted(fy)[-3:]
    if len(years) < 3:
        raise ValueError(f'Need 3 years of annual statements for {ticker}; Yahoo returned {len(years)}. Check the ticker / exchange suffix.')
    h = {k: [] for k in KEYS}
    cash_names = ['Cash Cash Equivalents And Short Term Investments', 'Cash And Cash Equivalents']
    for y in years:
        f, b, c = fy[y], by.get(y), cy.get(y)
        rev = _pick(fin, ['Total Revenue', 'Operating Revenue'], f)
        cogs = _pick(fin, ['Cost Of Revenue', 'Reconciled Cost Of Revenue'], f)
        sga = _pick(fin, ['Selling General And Administration'], f)
        rd = _pick(fin, ['Research And Development'], f)
        ebit = _pick(fin, ['Operating Income', 'EBIT'], f)
        da = (_pick(cf, ['Depreciation And Amortization', 'Depreciation Amortization Depletion', 'Depreciation'], c)
              or _pick(fin, ['Reconciled Depreciation'], f))
        sbc = _pick(cf, ['Stock Based Compensation'], c)
        plug = rev - cogs - sga - rd - (ebit + da + sbc)     # other opex so that Adj. EBITDA = EBIT + D&A + SBC
        if plug < 0:
            sga, plug = sga + plug, 0.0
        vals = dict(rev=rev, cogs=cogs, sga=sga, rd=rd, oth=plug, da=da, sbc=sbc,
                    capex=abs(_pick(cf, ['Capital Expenditure'], c)),
                    ar=_pick(bs, ['Accounts Receivable', 'Receivables'], b), inv=_pick(bs, ['Inventory'], b),
                    oca=_pick(bs, ['Other Current Assets', 'Prepaid Assets'], b),
                    ap=_pick(bs, ['Accounts Payable', 'Payables'], b),
                    ocl=_pick(bs, ['Current Accrued Expenses', 'Other Current Liabilities'], b, agg=True),
                    ppe=_pick(bs, ['Net PPE'], b),
                    int=abs(_pick(fin, ['Interest Expense', 'Interest Expense Non Operating'], f)),
                    debt=_pick(bs, ['Total Debt'], b), cash=_pick(bs, cash_names, b),
                    equity=_pick(bs, ['Stockholders Equity', 'Total Equity Gross Minority Interest', 'Common Stock Equity'], b))
        for k in KEYS:
            h[k].append(vals[k] / M)
    lb, lf = by.get(years[-1]), fy[years[-1]]
    pretax, taxexp = _pick(fin, ['Pretax Income'], lf), _pick(fin, ['Tax Provision'], lf)
    eff = min(max(taxexp / pretax, 0.10), 0.35) if pretax > 0 and taxexp > 0 else 0.21
    price = _num(info.get('currentPrice') or info.get('regularMarketPrice') or info.get('previousClose'))
    shares = _num(info.get('sharesOutstanding') or info.get('impliedSharesOutstanding')) / M
    if not price or not shares:
        raise ValueError('Yahoo Finance returned no share price / share count for this ticker.')
    return dict(
        ticker=ticker.upper(), name=info.get('longName') or info.get('shortName') or ticker.upper(),
        currency=(info.get('financialCurrency') or info.get('currency') or 'USD').upper(),
        price_ccy=(info.get('currency') or '').upper(), last_fy=years[-1], hist=h, price=price, shares=shares,
        debt=_pick(bs, ['Total Debt'], lb) / M, cash=_pick(bs, cash_names, lb) / M,
        minority=_pick(bs, ['Minority Interest'], lb) / M, pref=_pick(bs, ['Preferred Stock'], lb) / M,
        eff_tax=eff, beta=_num(info.get('beta')) or 1.0, de=_num(info.get('debtToEquity')) / 100.0,
        ev_ebitda=_num(info.get('enterpriseToEbitda')), sector=info.get('sector', ''), industry=info.get('industry', ''),
        market_cap=_num(info.get('marketCap')) / M)


def risk_free():
    """Live US 10Y Treasury yield; override on the Macro_Scenarios tab for non-USD companies."""
    try:
        v = _num(yf.Ticker('^TNX').history(period='5d')['Close'].dropna().iloc[-1]) / 100.0
        return v if 0.005 < v < 0.15 else DEFAULT_RF
    except Exception:
        return DEFAULT_RF
