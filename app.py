"""Streamlit front-end: enter a ticker -> download a fully formula-driven DCF workbook. All DCF inputs are edited in Excel."""
import os, tempfile
import altair as alt
import pandas as pd
import streamlit as st
from dcf.data import fetch_company, risk_free
from dcf.build import build_workbook, preview, have_libreoffice

st.set_page_config(page_title='Ticker to DCF', page_icon='\U0001F4C8', layout='wide', initial_sidebar_state='expanded')

st.markdown("""
<style>
.block-container {padding-top: 1.4rem; max-width: 1200px;}
.hero {position:relative; overflow:hidden; color:#fff; padding:0; border-radius:20px; margin-bottom:18px;
       border:1px solid rgba(255,255,255,.14);
       background:
         radial-gradient(1100px 460px at 88% -30%, rgba(56,189,248,.55), transparent 60%),
         radial-gradient(900px 520px at -12% 130%, rgba(29,78,216,.75), transparent 55%),
         linear-gradient(115deg,#081C3F 0%,#122C5E 35%,#1D4ED8 78%,#0E9BD8 100%);
       box-shadow: inset 0 1px 0 rgba(255,255,255,.22), 0 22px 44px -20px rgba(8,28,63,.55);}
.hero::after {content:""; position:absolute; top:0; right:0; bottom:0; width:44%;
              background-image: radial-gradient(rgba(255,255,255,.32) 1px, transparent 1.35px);
              background-size:17px 17px; opacity:.45; pointer-events:none;
              -webkit-mask-image: linear-gradient(to left, #000 30%, transparent); mask-image: linear-gradient(to left, #000 30%, transparent);}
.hero-inner {position:relative; padding:30px 36px 22px;}
.hero h1 {margin:0 0 8px 0; font-size:2.35rem; font-weight:800; color:#fff; letter-spacing:-0.03em; line-height:1.15; text-wrap:balance; max-width:26ch;}
.hero p {margin:0; font-size:1.02rem; line-height:1.5; color:#CBDDFB; max-width:62ch;}
.hero ::selection {background:#FBBF24; color:#081C3F;}
.hero-tape {position:relative; border-top:1px solid rgba(255,255,255,.16); padding:8px 0; overflow:hidden;
            -webkit-mask-image: linear-gradient(to right, transparent, #000 4%, #000 96%, transparent);
            mask-image: linear-gradient(to right, transparent, #000 4%, #000 96%, transparent);}
.hero-tape-inner {display:inline-flex; white-space:nowrap; animation:hero-tape 110s linear infinite; will-change:transform;}
.hero-tape span {font-size:.74rem; font-weight:600; letter-spacing:.06em; color:rgba(219,234,254,.62); padding:0 .9rem;}
.hero-tape b {color:rgba(255,255,255,.85); font-weight:600;}
@keyframes hero-tape { from { transform: translateX(0); } to { transform: translateX(-50%); } }
@media (prefers-reduced-motion: reduce) { .hero-tape-inner { animation: none; } }
@media (max-width: 900px) { .hero::after { display:none; } .hero-inner { padding:22px 22px 18px; } .hero h1 { font-size:1.7rem; } }
.chip {display:inline-block; padding:3px 11px; border-radius:999px; background:#E0E7FF; color:#1E3A8A; font-size:.8rem; margin:0 6px 6px 0; font-weight:600;}
.chip.g {background:#DCFCE7; color:#166534;} .chip.a {background:#FEF3C7; color:#92400E;} .chip.r {background:#FEE2E2; color:#991B1B;}
.card {border:1px solid #E2E8F0; border-radius:14px; padding:14px 18px; background:#fff; box-shadow:0 1px 3px rgba(15,23,42,.06);}
.card .lbl {font-size:.78rem; color:#64748B; text-transform:uppercase; letter-spacing:.6px; font-weight:600;}
.card .val {font-size:1.65rem; font-weight:700; color:#0F172A; line-height:1.25;}
.card .sub {font-size:.82rem; color:#64748B;}
.up {color:#15803D !important;} .down {color:#B91C1C !important;}
div[data-testid="stDownloadButton"] button {width:100%; font-weight:700; padding:.65rem 1rem; border-radius:12px;}
div[data-testid="stForm"] {border:none; padding:0;}
.stTabs [data-baseweb="tab"] {font-weight:600;}
</style>
""", unsafe_allow_html=True)

EXAMPLES = ['AAPL', 'MSFT', 'NVDA', 'KO', 'RELIANCE.NS', 'TCS.NS', 'SAP.DE', 'NESN.SW']
NAVY, BLUE, GREEN, RED, AMBER = '#1E3A8A', '#2563EB', '#16A34A', '#DC2626', '#F59E0B'


def card(label, value, sub='', cls=''):
    return f'<div class="card"><div class="lbl">{label}</div><div class="val {cls}">{value}</div><div class="sub">{sub}</div></div>'


def sget(sc, row, col):
    try:
        v = sc.loc[row, col]
        return float(v) if v == v else None
    except (KeyError, TypeError, ValueError):
        return None


def money(v, ccy=''):
    if v is None:
        return 'n/a'
    a = abs(v)
    s = f'{a/1e6:,.2f}T' if a >= 1e6 else f'{a/1e3:,.1f}B' if a >= 1e3 else f'{a:,.0f}M'
    return f'{"-" if v < 0 else ""}{s} {ccy}'.strip()


# ------------------------------------------------------------------ sidebar
with st.sidebar:
    st.markdown('### \U0001F4C8 Ticker \u2192 DCF')
    st.markdown('**How it works**')
    st.markdown('1. Enter a Yahoo Finance ticker\n2. We pull 3 years of financials & build the model\n'
                '3. Download the Excel \u2014 every input is live there\n4. Change drivers in **Excel**; everything recalculates')
    st.divider()
    st.markdown('**Workbook tabs**')
    st.caption('Cover \u00b7 Assumptions \u00b7 Segments \u00b7 Macro_Scenarios \u00b7 IS \u00b7 BS \u00b7 CFS \u00b7 WACC \u00b7 DCF_Valuation \u00b7 Comps \u00b7 Scenario_Valuation \u00b7 Sensitivity_Analysis \u00b7 Precedents \u00b7 Risks \u00b7 Controls \u00b7 Version_Log')
    st.divider()
    st.markdown('**Ticker format**')
    st.caption('US: `AAPL` \u00b7 India: `RELIANCE.NS` \u00b7 Germany: `SAP.DE` \u00b7 UK: `BP.L` \u00b7 Japan: `7203.T` \u00b7 Switzerland: `NESN.SW`')
    st.divider()
    st.caption('Data: Yahoo Finance (unofficial, may be incomplete). Not for banks / insurers / REITs. Illustrative output, not investment advice.')

# ------------------------------------------------------------------ hero + input
REGIMES = ['16-tab institutional model', 'Post-SBC GAAP IS', 'Segments + buybacks', '3-statement integrated model', 'Bear / Base / Bull', 'WACC via CAPM + peer betas', 'Gordon + exit multiple blend', 'TSM dilution', 'Peer multiples cross-check', 'Football field & implied rating', 'Driver-based forecast', 'Full BS articulation', 'Balance check = 0']
_tape = ' \u2002&middot;\u2002 '.join(REGIMES) + ' \u2002&middot;\u2002 '
st.markdown('<div class="hero"><div class="hero-inner"><h1>Ticker \u2192 DCF Valuation Workbook</h1>'
            '<p>Enter a ticker. Get a bank-grade, fully formula-driven Excel DCF: 3-statement integrated model, WACC via CAPM with peer betas, Bear / Base / Bull scenarios, comps cross-check, football field and sensitivity tables.</p></div>'
            f'<div class="hero-tape"><div class="hero-tape-inner"><span>{_tape}</span><span>{_tape}</span></div></div></div>',
            unsafe_allow_html=True)


def pick(t):
    st.session_state['ticker_in'] = t
    st.session_state['autorun'] = True


with st.form('f', border=False):
    c1, c2 = st.columns([4, 1], vertical_alignment='bottom')
    ticker = c1.text_input('Ticker symbol', key='ticker_in', value='AAPL', placeholder='e.g. AAPL, RELIANCE.NS, SAP.DE').strip().upper()
    go = c2.form_submit_button('Generate Excel  \u2192', type='primary', width='stretch')
st.caption('Try one:')
cols = st.columns(len(EXAMPLES))
for c, t in zip(cols, EXAMPLES):
    c.button(t, key=f'ex_{t}', on_click=pick, args=(t,), width='stretch')

run = go or st.session_state.pop('autorun', False)
if run:
    if not ticker:
        st.error('Enter a ticker symbol.')
    else:
        try:
            with st.status(f'Building DCF for {ticker} ...', expanded=True) as stat:
                st.write('\U0001F4E5 Fetching financial statements & market data from Yahoo Finance')
                d = fetch_company(ticker)
                rf = risk_free()
                st.write('\U0001F9EE Building the 16-tab formula-driven workbook and pre-filling assumptions')
                out = os.path.join(tempfile.mkdtemp(), f'DCF_{ticker.replace("^", "").replace("/", "_")}.xlsx')
                build_workbook(d, rf, out)
                prev = None
                if have_libreoffice():
                    st.write('\U0001F501 Recalculating to preview results')
                    prev = preview(out)
                with open(out, 'rb') as fh:
                    st.session_state['res'] = dict(d=d, rf=rf, bytes=fh.read(), name=os.path.basename(out), prev=prev)
                stat.update(label=f'DCF ready for {d["name"]}', state='complete', expanded=False)
        except Exception as e:
            st.session_state.pop('res', None)
            st.error(f'**{type(e).__name__}:** {e}')
            st.caption('Check the ticker (add the exchange suffix, e.g. `.NS`, `.DE`, `.L`) or try again in a minute (Yahoo may rate-limit).')
            with st.expander('Developer details (full traceback)'):
                st.exception(e)

res = st.session_state.get('res')
if not res:
    st.info('Enter a ticker above or pick an example to generate your workbook.')
    st.stop()

d, prev = res['d'], res['prev']
ccy = d['currency']
hist = d['hist']
ebitda_hist = [hist['rev'][i] - hist['cogs'][i] - hist['sga'][i] - hist['rd'][i] - hist['oth'][i] for i in range(3)]

# ------------------------------------------------------------------ header
chips = f'<span class="chip">{d["ticker"]}</span><span class="chip">{d["sector"] or "n/a"}</span><span class="chip">{d["industry"] or "n/a"}</span>' \
        f'<span class="chip">{ccy} millions</span><span class="chip">Base year FY{d["last_fy"]}</span>'
if d['price_ccy'] and d['price_ccy'] != ccy:
    chips += f'<span class="chip a">Price in {d["price_ccy"]} \u2260 reporting {ccy}</span>'
if hist['cogs'][2] == 0:
    chips += '<span class="chip r">No COGS: not suited to financials</span>'
st.markdown(f'## {d["name"]}')
st.markdown(chips, unsafe_allow_html=True)

k = st.columns(5)
k[0].markdown(card('Share price', f'{d["price"]:,.2f}', d['price_ccy'] or ccy), unsafe_allow_html=True)
k[1].markdown(card('Market cap', money(d['market_cap']), f'{d["shares"]:,.0f}m shares'), unsafe_allow_html=True)
if prev and 'Base' in prev['scen'].index:
    sc = prev['scen']
    base_px, up, base_wacc = sget(sc, 'Base', 'Price (blended)'), sget(sc, 'Base', 'Upside / (downside)'), sget(sc, 'Base', 'WACC')
    k[2].markdown(card('Base-case value / share', f'{base_px:,.2f}' if base_px is not None else 'n/a', 'blended DCF'), unsafe_allow_html=True)
    k[3].markdown(card('Upside / (downside)', f'{up:+.1%}' if up is not None else 'n/a', 'vs. current price', 'up' if (up or 0) >= 0 else 'down'), unsafe_allow_html=True)
    k[4].markdown(card('Base WACC', f'{base_wacc:.2%}' if base_wacc is not None else 'n/a', f'Rf {res["rf"]:.2%}'), unsafe_allow_html=True)
else:
    k[2].markdown(card('Revenue FY%d' % d['last_fy'], money(hist['rev'][2]), ''), unsafe_allow_html=True)
    k[3].markdown(card('Adj. EBITDA margin', f'{ebitda_hist[2]/hist["rev"][2]:.1%}', ''), unsafe_allow_html=True)
    k[4].markdown(card('Risk-free (10Y)', f'{res["rf"]:.2%}', ''), unsafe_allow_html=True)
st.write('')

dl, note = st.columns([1, 2], vertical_alignment='center')
dl.download_button('\u2B07  Download Excel workbook', res['bytes'], file_name=res['name'],
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
note.caption('Open the file \u2192 **Assumptions** tab: set the scenario (C5), edit drivers, terminal inputs & probabilities. '
             'Mid-year toggle is on **DCF_Valuation**. Excel recalculates on open.')
if prev and prev.get('rating'):
    pt = prev.get('price_target')
    pt_txt = f' · blended DCF value **{pt:,.2f}**' if pt else ''
    note.caption(f'DCF-implied rating: **{prev["rating"]}**{pt_txt}')

# ------------------------------------------------------------------ tabs
t1, t2, t3, t4, t5 = st.tabs(['\U0001F4CA Valuation', '\U0001F4C8 Financials', '\U0001F3AF Sensitivity', '\U0001F9FE Pre-filled inputs', '\u2705 Model checks'])

with t1:
    if not prev:
        st.info('Results preview needs LibreOffice on the server (it is optional). Download the workbook to see the full valuation in Excel.')
    else:
        sc = prev['scen']
        n_sc = len(sc) - 1
        mc = st.columns(4)
        for col, nm in zip(mc, ['Bear', 'Base', 'Bull', 'Probability-weighted']):
            p, u_ = sget(sc, nm, 'Price (blended)'), sget(sc, nm, 'Upside / (downside)')
            col.metric(nm if nm != 'Probability-weighted' else 'Prob-weighted',
                       f'{p:,.2f}' if p is not None else 'n/a', f'{u_:+.1%}' if u_ is not None else None)
        a, b = st.columns([3, 2])
        with a:
            st.markdown(f'**Implied value per share \u2014 all {n_sc} scenarios** (blended DCF)')
            df = sc.drop('Probability-weighted', errors='ignore')[['Price (blended)']].reset_index()
            df.columns = ['Scenario', 'Price']
            df['Group'] = df['Scenario'].map(lambda x: x if x in ('Bear', 'Base', 'Bull') else 'Other')
            cmap = alt.Scale(domain=['Bear', 'Base', 'Bull', 'Other'], range=[RED, BLUE, GREEN, '#94A3B8'])
            bars = alt.Chart(df).mark_bar(cornerRadiusTopRight=5, cornerRadiusBottomRight=5).encode(
                y=alt.Y('Scenario:N', sort='-x', title=None, axis=alt.Axis(labelLimit=260)), x=alt.X('Price:Q', title=f'{ccy} per share'),
                color=alt.Color('Group:N', scale=cmap, legend=None), tooltip=['Scenario', alt.Tooltip('Price:Q', format=',.2f')])
            rule = alt.Chart(pd.DataFrame({'p': [prev['price']]})).mark_rule(color=AMBER, strokeDash=[6, 4], size=2).encode(x='p:Q')
            st.altair_chart((bars + rule).properties(height=max(320, 22 * n_sc + 40)), width='stretch')
            st.caption(f'Amber line = current price ({prev["price"]:,.2f}). Red / blue / green = Bear / Base / Bull.')
        with b:
            st.markdown('**EV composition (active scenario, blended)**')
            comp = pd.DataFrame({'Part': ['PV of explicit UFCF', 'PV of terminal value'], 'Value': [prev['pv_explicit'], prev['pv_tv']]})
            donut = alt.Chart(comp).mark_arc(innerRadius=62, outerRadius=105).encode(
                theta='Value:Q', color=alt.Color('Part:N', scale=alt.Scale(range=[BLUE, NAVY]), legend=alt.Legend(orient='bottom', title=None)),
                tooltip=['Part', alt.Tooltip('Value:Q', format=',.0f')])
            st.altair_chart(donut.properties(height=300), width='stretch')
            st.caption(f'Terminal value = **{prev["tv_pct"]:.0%}** of enterprise value.')
            m1, m2 = st.columns(2)
            m1.metric('Gordon only', f'{prev["gordon"]:,.2f}')
            m2.metric('Exit multiple only', f'{prev["exitm"]:,.2f}')
        st.markdown('**Scenario matrix**')
        fm = {'Probability': '{:.0%}', 'Revenue CAGR': '{:.1%}', 'FY30E EBITDA margin': '{:.1%}', 'FY30E UFCF': '{:,.0f}', 'WACC': '{:.2%}',
              'Enterprise value': '{:,.0f}', 'Equity value': '{:,.0f}', 'Price (Gordon)': '{:,.2f}', 'Price (Exit)': '{:,.2f}',
              'Price (blended)': '{:,.2f}', 'Upside / (downside)': '{:+.1%}', 'TV % of EV': '{:.0%}'}
        st.dataframe(sc.style.format(fm, na_rep='n/a'), width='stretch', height=min(900, 36 * len(sc) + 40))
        st.caption('Probabilities are inputs in Excel (`Scenario_Valuation`, default 20 / 60 / 20 = 100%). '
                   'Pick the active scenario in `Assumptions!C5`.')

with t2:
    if prev:
        tr = prev['trend']
        base = alt.Chart(tr).encode(x=alt.X('Year:N', sort=None, title=None, axis=alt.Axis(labelAngle=0)))
        rev = base.mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
            y=alt.Y('Revenue:Q', title=f'Revenue ({ccy} m)'), color=alt.condition(alt.datum.Forecast, alt.value('#93C5FD'), alt.value(BLUE)),
            tooltip=['Year', alt.Tooltip('Revenue:Q', format=',.0f'), alt.Tooltip('EBITDA margin:Q', format='.1%')])
        mar = base.mark_line(color=AMBER, point=alt.OverlayMarkDef(color=AMBER, size=60), strokeWidth=3).encode(
            y=alt.Y('EBITDA margin:Q', title='Adj. EBITDA margin', axis=alt.Axis(format='%', orient='right')))
        st.markdown('**Revenue & Adj. EBITDA margin** (dark = actual, light = Base-case forecast)')
        st.altair_chart(alt.layer(rev, mar).resolve_scale(y='independent').properties(height=330), width='stretch')
        u = prev['ufcf'].melt('Year', var_name='Series', value_name='Value')
        st.markdown('**Unlevered free cash flow vs. present value (Base)**')
        st.altair_chart(alt.Chart(u).mark_bar().encode(
            x=alt.X('Year:N', sort=None, axis=alt.Axis(labelAngle=0), title=None), xOffset='Series:N',
            y=alt.Y('Value:Q', title=f'{ccy} m'), color=alt.Color('Series:N', scale=alt.Scale(range=[BLUE, NAVY]), legend=alt.Legend(orient='bottom', title=None)),
            tooltip=['Year', 'Series', alt.Tooltip('Value:Q', format=',.0f')]).properties(height=280), width='stretch')
    st.markdown('**Historical financials pulled from Yahoo Finance** (millions)')
    yrs = [f'FY{d["last_fy"]-2}A', f'FY{d["last_fy"]-1}A', f'FY{d["last_fy"]}A']
    rows = {'Revenue': hist['rev'], 'COGS': hist['cogs'], 'SG&A': hist['sga'], 'R&D': hist['rd'], 'Other opex': hist['oth'],
            'Adj. EBITDA': ebitda_hist, 'D&A': hist['da'], 'SBC': hist['sbc'], 'CapEx': hist['capex'], 'Net PP&E': hist['ppe'],
            'Total debt': hist['debt'], 'Cash': hist['cash']}
    hd = pd.DataFrame(rows, index=yrs).T
    st.dataframe(hd.style.format('{:,.0f}'), width='stretch')

with t3:
    if not prev:
        st.info('Sensitivity preview needs LibreOffice on the server. The full tables are in the workbook (`Sensitivity_Analysis`).')
    else:
        s = prev['sens'].dropna()
        s['W'] = s['WACC'].map(lambda x: f'{x:.2%}'); s['G'] = s['g'].map(lambda x: f'{x:.2%}')
        px0 = prev['price']
        hm = alt.Chart(s).mark_rect().encode(
            x=alt.X('G:N', title='Terminal growth g', sort=None, axis=alt.Axis(labelAngle=0)), y=alt.Y('W:N', title='WACC', sort=None),
            color=alt.Color('Price:Q', scale=alt.Scale(scheme='redyellowgreen', domainMid=px0), legend=None),
            tooltip=['W', 'G', alt.Tooltip('Price:Q', format=',.2f')])
        tx = hm.mark_text(fontWeight='bold').encode(text=alt.Text('Price:Q', format=',.2f'), color=alt.value('#0F172A'))
        st.markdown(f'**Implied share price: WACC \u00d7 terminal growth** (Gordon method, active case) \u2014 green above current price ({px0:,.2f}), red below')
        st.altair_chart((hm + tx).properties(height=320), width='stretch')
        st.caption('Two more tables (WACC \u00d7 exit multiple, revenue growth \u00d7 terminal margin) are in the workbook.')

with t4:
    st.markdown('Everything below is written into the Excel as **editable inputs**. Review and overwrite there.')
    u = d['_used']
    st.dataframe(pd.DataFrame({'Input': list(u), 'Value': [f'{v:,.4g}' if isinstance(v, float) else str(v) for v in u.values()]}),
                 hide_index=True, width='stretch')
    st.markdown('- **Also pre-filled:** 3-year historicals (IS / BS / CFS), Bear / Base / Bull drivers, macro paths; revenue growth fading from the historical CAGR to 3.5%. Balance check ties to zero by construction.\n'
                '- **You should replace:** peer betas on `WACC` (placeholder = company itself), options/RSUs on `Assumptions` (default 0), the macro paths for your region.\n'
                '- **Defaults:** terminal growth 2.5%, 50/50 Gordon vs exit multiple, mid-year on (see `dcf/config.py`).')

with t5:
    if not prev:
        st.info('Model checks run inside the workbook (`DCF_Valuation` audit block). Preview needs LibreOffice on the server.')
    else:
        for name, status in prev['audit']:
            good = str(status).startswith(('OK', 'ALL'))
            icon = '\u2705' if good else ('\u26A0\uFE0F' if 'REVIEW' in str(status) else '\u274C')
            st.markdown(f'{icon} **{status}** \u2014 {name}')
        st.caption('"REVIEW terminal ROIC" means the terminal reinvestment rate implies returns above WACC by more than the tolerance \u2014 sensible for high-ROIC businesses; adjust CapEx/growth if you disagree.')
