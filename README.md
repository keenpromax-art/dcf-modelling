# Ticker → DCF (Streamlit) — Institutional Model

Enter a ticker, download a 16-tab, fully formula-driven, sell-side-style Excel DCF.

**Model architecture**
- **Post-SBC GAAP-style IS**: EBITDA / EBIT / NI after stock-based compensation (no double count in UFCF).
- **Segments**: revenue mix + margin by business line seeded from `ticker_full_dataset_v2.csv` (e.g. AAPL → iPhone/Mac/iPad/Wearables/Services); total revenue and EBITDA tie to the IS.
- **Integrated BS / CFS**: cash rolls from CFS, equity = opening + NI + SBC − dividends − buybacks, balance check = 0 every year, CFS↔BS tie.
- **Buybacks**: repurchase budget as % of UFCF; rolling diluted share count feeds EPS and per-share valuation.
- **ROIC-driven terminal value**: reinvestment = g / terminal ROIC; Gordon + exit multiple blend; implied RONIC vs WACC check.
- **WACC**: Hamada unlever→median→relever of 5 curated peers (from CSV), CAPM + ERP + CRP + size premium, synthetic rating from interest coverage.
- **Comps / Precedents**: peer tickers and M&A placeholders from the dataset; editable and overridable.
- **Scenarios**: Bear / Base / Bull (20 / 60 / 20), all live simultaneously; two-way sensitivity grids.
- **Cover**: implied rating vs thresholds, football field, key KPIs, audit strip.
- **Controls / Version_Log / Risks**: 11 live model checks, change log, seeded company-specific risk register.

**Peer / segment data** comes from a compact 1.3 MB lookup (`dcf/peers_lookup.json`, 1,509 tickers — top companies per country from the 210 MB source CSV). `dcf/data.py` fetches live multiples/betas for peer tickers from Yahoo with resilient fallbacks.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
LibreOffice (`soffice` on PATH) enables the on-page preview; Streamlit Cloud installs it via `packages.txt`.

## Notes
- Yahoo data is unofficial and incomplete — review every blue input before circulating.
- Not suitable for banks / insurers / REITs.
- Terminal ROIC REVIEW flag is expected for high-ROIC businesses (g/ROIC reinvestment is very low); adjust the ROIC input or rely on the exit-multiple method.
- Buybacks are modelled on forecast years only; M&A precedents are illustrative placeholders.
