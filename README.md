# Ticker → DCF (Streamlit)

Enter a ticker, download a fully formula-driven, sell-side-style Excel DCF.

**Fully integrated 3-statement model** — IS / BS / CFS all articulate (historical + 5-year forecast):
balance-sheet cash is rolled from the cash-flow statement, equity rolls as opening + NI + SBC − dividends,
and the BS **balance check row = 0** in every year by construction.

**Bank-standard conventions throughout** — blue inputs (yellow = fill-first), black formulas,
green cross-sheet links, one driver cell per assumption, every scenario switchable on the `Assumptions` tab,
and a full audit block on `DCF_Valuation` (formula errors, BS = 0, CFS↔BS tie, WACC > g, sensitivity ties).

## Tabs

Cover · Assumptions · Macro_Scenarios · IS · BS · CFS · WACC · DCF_Valuation · Comps · Scenario_Valuation · Sensitivity_Analysis

- **Cover** — analyst frame (buy/hold/sell thresholds you set), company snapshot, blended Gordon/exit/EV bridge,
  **football field** (floating bars, Bear | Base | Bull and ranges by method), implied rating, key KPIs, audit strip.
- **IS / BS / CFS** — driver-based forecast (TAM × volume share × price spread, gross-margin pressure),
  DSO/DIO/DPO working capital from the BS, PP&E and debt schedules, average-debt interest.
- **WACC** — 5-peer Hamada unlever→median→relever at applied capital structure, CAPM + ERP + CRP + size premium,
  synthetic credit rating from interest coverage (large-cap bands).
- **DCF_Valuation** — UFCF build, mid-year convention, Gordon + exit blend, **TSM dilution**,
  implied terminal RONIC vs WACC check, sensitivity-tied audit block.
- **Comps** — own company fully linked + 5 editable peers → implied price from EV/EBITDA FY26E and P/E FY25.
- **Scenario_Valuation / Macro_Scenarios** — Bear / Base / Bull (20 / 60 / 20 probabilities), live simultaneously,
  probability-weighted value; active case drives the whole model via `Assumptions!C5`.
- **Sensitivity_Analysis** — WACC × g, WACC × exit multiple, growth × terminal-margin grids, all live.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
Optional: install LibreOffice (`soffice` on PATH) to show a quick results preview on the page.
On Streamlit Community Cloud, `packages.txt` installs it automatically.

## Layout
| File | Role |
|---|---|
| `app.py` | Streamlit UI (ticker → download) |
| `dcf/data.py` | Yahoo Finance extraction (`yfinance`) |
| `dcf/fill.py` | Writes ticker data into the workbook inputs (hacks every override you may want) |
| `dcf/config.py` | Default assumptions (terminal growth, exit-multiple clip, tax, fade rate…) |
| `dcf/build.py` | Runs the builder, optional LibreOffice preview |
| `dcf/builder_core.py` | The workbook generator (also runs standalone: `python dcf/builder_core.py out.xlsx`) |

## Notes
- Yahoo data is unofficial and can be incomplete; review every input in the workbook.
- Not suitable for banks / insurers / REITs (no COGS-based operating model).
- Placeholder peers (Comps + WACC betas) and option/RSU tranches are obvious placeholders — replace them.
- Buybacks are not modelled; dividends are a payout-ratio input on `CFS`. Equity rolls by NI + SBC − dividends.
- Historical cash flow never ties (buybacks/actuals omitted); the BS/CFS tie checks apply to forecast years.
- Risk-free = live US 10Y yield; override on `Macro_Scenarios` for non-USD companies.
