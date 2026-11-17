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
.hero {background: linear-gradient(120deg,#1E3A8A 0%,#2563EB 55%,#0EA5E9 100%); color:#fff; padding: 30px 34px; border-radius: 18px; margin-bottom: 18px;}
.hero h1 {margin:0 0 6px 0; font-size: 2.1rem; color:#fff; letter-spacing:-0.5px;}
.hero p {margin:0; font-size: 1.02rem; opacity:.93;}
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
    st.caption('Dashboard \u00b7 Assumptions \u00b7 Macro_Scenarios \u00b7 Operating_Model \u00b7 WACC \u00b7 DCF_Valuation \u00b7 Scenario_Valuation \u00b7 Sensitivity_Analysis')
    st.divider()
    st.markdown('**Ticker format**')
    st.caption('US: `AAPL` \u00b7 India: `RELIANCE.NS` \u00b7 Germany: `SAP.DE` \u00b7 UK: `BP.L` \u00b7 Japan: `7203.T` \u00b7 Switzerland: `NESN.SW`')
    st.divider()
    st.caption('Data: Yahoo Finance (unofficial, may be incomplete). Not for banks / insurers / REITs. Illustrative output, not investment advice.')

# ------------------------------------------------------------------ hero + input
st.markdown('<div class="hero"><h1>Ticker \u2192 DCF Valuation Workbook</h1>'
            '<p>Enter a ticker. Get an institutional-style, fully formula-driven Excel DCF with Bear / Base / Bull scenarios, WACC build and sensitivity tables.</p></div>',
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
                st.write('\U0001F9EE Building the 8-tab formula-driven workbook and pre-filling assumptions')
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
if prev:
    sc = prev['scen']
    base_px, up = sc.loc['Implied share price', 'Base'], sc.loc['Upside / (downside)', 'Base']
    k[2].markdown(card('Base-case value / share', f'{base_px:,.2f}', 'blended DCF'), unsafe_allow_html=True)
    k[3].markdown(card('Upside / (downside)', f'{up:+.1%}', 'vs. current price', 'up' if up >= 0 else 'down'), unsafe_allow_html=True)
    k[4].markdown(card('Base WACC', f'{sc.loc["WACC", "Base"]:.2%}', f'Rf {res["rf"]:.2%}'), unsafe_allow_html=True)
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

# ------------------------------------------------------------------ tabs
t1, t2, t3, t4, t5 = st.tabs(['\U0001F4CA Valuation', '\U0001F4C8 Financials', '\U0001F3AF Sensitivity', '\U0001F9FE Pre-filled inputs', '\u2705 Model checks'])

with t1:
    if not prev:
        st.info('Results preview needs LibreOffice on the server (it is optional). Download the workbook to see the full valuation in Excel.')
    else:
        sc = prev['scen']
        a, b = st.columns([3, 2])
        with a:
            st.markdown('**Implied value per share by scenario**')
            df = pd.DataFrame({'Scenario': sc.columns, 'Price': sc.loc['Implied share price'].values})
            cmap = alt.Scale(domain=['Bear', 'Base', 'Bull', 'Prob-weighted'], range=[RED, BLUE, GREEN, NAVY])
            bars = alt.Chart(df).mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, size=55).encode(
                x=alt.X('Scenario:N', sort=None, axis=alt.Axis(labelAngle=0, title=None)), y=alt.Y('Price:Q', title=f'{ccy} per share'),
                color=alt.Color('Scenario:N', scale=cmap, legend=None), tooltip=['Scenario', alt.Tooltip('Price:Q', format=',.2f')])
            txt = bars.mark_text(dy=-8, fontWeight='bold').encode(text=alt.Text('Price:Q', format=',.2f'))
            rule = alt.Chart(pd.DataFrame({'p': [prev['price']]})).mark_rule(color=AMBER, strokeDash=[6, 4], size=2).encode(y='p:Q')
            lbl = alt.Chart(pd.DataFrame({'p': [prev['price']], 't': [f'Current {prev["price"]:,.2f}']})).mark_text(
                align='left', dx=4, dy=-8, color=AMBER, fontWeight='bold').encode(y='p:Q', text='t:N', x=alt.value(4))
            st.altair_chart((bars + txt + rule + lbl).properties(height=320), width='stretch')
        with b:
            st.markdown('**EV composition (Base blended)**')
            comp = pd.DataFrame({'Part': ['PV of explicit UFCF', 'PV of terminal value'], 'Value': [prev['pv_explicit'], prev['pv_tv']]})
            donut = alt.Chart(comp).mark_arc(innerRadius=62, outerRadius=105).encode(
                theta='Value:Q', color=alt.Color('Part:N', scale=alt.Scale(range=[BLUE, NAVY]), legend=alt.Legend(orient='bottom', title=None)),
                tooltip=['Part', alt.Tooltip('Value:Q', format=',.0f')])
            st.altair_chart(donut.properties(height=320), width='stretch')
            st.caption(f'Terminal value = **{prev["tv_pct"]:.0%}** of enterprise value (active scenario).')
        m1, m2 = st.columns(2)
        m1.metric('Gordon Growth only (active case)', f'{prev["gordon"]:,.2f}')
        m2.metric('Exit Multiple only (active case)', f'{prev["exitm"]:,.2f}')
        st.markdown('**Scenario matrix**')
        fm = {'Revenue CAGR (5Y)': '{:.1%}', 'FY30E EBITDA margin': '{:.1%}', 'FY30E UFCF': '{:,.0f}', 'WACC': '{:.2%}',
              'Enterprise value': '{:,.0f}', 'Equity value': '{:,.0f}', 'Implied share price': '{:,.2f}', 'Upside / (downside)': '{:+.1%}'}
        show = sc.copy().astype(object)
        for r in show.index:
            show.loc[r] = [fm[r].format(v) if v is not None else 'n/a' for v in sc.loc[r]]
        st.dataframe(show, width='stretch')
        st.caption('Probability weights default to 25 / 50 / 25 \u2014 change in Excel (`Scenario_Valuation`). Values use the default assumptions written by this app.')

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
    st.markdown('- **Also pre-filled:** 3-year historicals; Bear/Base/Bull drivers from last-2-year ratios; macro paths; revenue growth fading from the historical CAGR to 3.5%.\n'
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
