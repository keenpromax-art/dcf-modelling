# Ticker → DCF (Streamlit)

Enter a ticker, download a fully formula-driven 8-tab Excel DCF
with **23 scenarios** (Bear / Base / Bull + 20 macro & company regimes, library in `Assumptions!N5:O27`)
(Dashboard · Assumptions · Macro_Scenarios · Operating_Model · WACC · DCF_Valuation · Scenario_Valuation · Sensitivity_Analysis).
The web page asks **only for the ticker**; all DCF inputs are edited in Excel afterwards.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
Optional: install LibreOffice (`soffice` on PATH) to show a quick results preview on the page.
On Streamlit Community Cloud, `packages.txt` installs it automatically (remove that file to skip the preview and deploy faster).

## Layout
| File | Role |
|---|---|
| `app.py` | Streamlit UI (ticker → download) |
| `dcf/data.py` | Yahoo Finance extraction (`yfinance`) |
| `dcf/fill.py` | Writes ticker data + basic default assumptions into the workbook inputs |
| `dcf/config.py` | Default assumptions (terminal growth, exit-multiple clip, tax, fade rate...) |
| `dcf/build.py` | Runs the builder, optional LibreOffice preview |
| `dcf/builder_core.py` | The workbook generator (also runs standalone: `python dcf/builder_core.py out.xlsx` → blank "Example Co." template) |

## Notes
- Yahoo data is unofficial and can be incomplete; review every input in the workbook.
- Not suitable for banks / insurers / REITs (no COGS-based operating model).
- Placeholder peer betas = the company's own beta; replace on the `WACC` tab. Options/RSUs default to 0.
- Risk-free = live US 10Y yield; override on `Macro_Scenarios` for non-USD companies.
