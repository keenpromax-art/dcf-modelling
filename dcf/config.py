"""Basic default assumptions written into the Excel file. Edit them in Excel afterwards (blue / yellow cells)."""
TERMINAL_GROWTH = 0.025        # Assumptions!C33
GORDON_WEIGHT = 0.50           # Assumptions!C35 (rest = exit multiple)
MID_YEAR = 'Yes'               # DCF_Valuation!C17
EXIT_MULTIPLE_FALLBACK = 11.0  # used if Yahoo has no EV/EBITDA; otherwise current multiple clipped to 6x-20x
EXIT_MULTIPLE_CLIP = (6.0, 20.0)
MARGINAL_TAX = 0.25
FADE_TO_GROWTH = 0.035         # revenue growth fades from historical CAGR to this nominal rate by FY30E
DEFAULT_RF = 0.043             # used if the live 10Y yield cannot be fetched
CURRENCY_SYMBOLS = {'USD': '$', 'EUR': '\u20ac', 'GBP': '\u00a3', 'INR': '\u20b9', 'JPY': '\u00a5', 'CNY': '\u00a5', 'KRW': '\u20a9'}
