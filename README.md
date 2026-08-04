# DCF Valuation Tool

A single-file Python DCF (Discounted Cash Flow) valuation model. Give it a
ticker, and it fetches live financial data, calculates WACC from first
principles, projects free cash flow under Bull/Base/Bear scenarios, and
exports a formatted 3-sheet Excel valuation model — complete with
built-in checks that flag when the model's own assumptions don't hold
for a given company, instead of silently returning a misleading number.

Built as one file on purpose — no folder structure to navigate, no
`src` imports to break. Just download, install requirements, edit the
ticker, and run.

## What it produces

A 3-sheet Excel workbook:

1. **Cover** — scenario summary, current market price, implied growth,
   base FCF (with transparency into how it was calculated), WACC
   breakdown, and any reliability warnings
2. **DCF Model** — full 5-year FCF projections, terminal value, and
   enterprise-to-equity bridge for each scenario
3. **Sensitivity** — a WACC x terminal-growth-rate grid of intrinsic
   value per share

## How to run

### In a terminal
```
pip install -r requirements.txt
python dcf_valuation.py
```

### In Jupyter
Open this folder in Jupyter (Files → navigate to where these three
files live), then in a notebook cell:
```python
from dcf_valuation import run_valuation
filepath = run_valuation("AAPL", manual_growth=0.10)
print(filepath)
```

### Output location
Every run saves to `Documents/Project/Python Projects/{TICKER}_valuation.xlsx`
in your home folder — created automatically if it doesn't exist yet, so
you always know where to look regardless of your working directory.

## Changing the ticker

Open `dcf_valuation.py`, scroll to the very bottom, and edit this line:
```python
run_valuation("AAPL", manual_growth=0.10)
```
Swap `"AAPL"` for any public ticker. `manual_growth` is your assumed
Base Case FCF growth rate — leave it out entirely (or set to `None`) to
let the model estimate it from historical FCF growth instead.

### Non-US tickers
Works on any exchange `yfinance` covers — just use the right suffix:
- US (NASDAQ/NYSE): no suffix, e.g. `AAPL`, `JNJ`
- India (NSE): `.NS`, e.g. `TCS.NS`
- India (BSE): `.BO`, e.g. `RELIANCE.BO`
- UK (LSE): `.L`, e.g. `HSBA.L`

For non-US tickers, consider overriding the risk-free rate and equity
risk premium, since the defaults are US Treasury-based:
```python
run_valuation("TCS.NS", manual_growth=0.12, risk_free_rate=0.069, equity_risk_premium=0.07)
```

## Methodology

- **Base FCF**: the average of the last 3 years of free cash flow
  (fewer if less history is available), rather than only the most
  recent year. A single unusual year — a working capital swing, a
  one-off charge, a heavy CapEx cycle — can otherwise dominate the
  entire valuation even when it doesn't reflect normal cash generation.
  Both the smoothed average and the raw most-recent-year figure are
  shown side by side in the output for transparency.
- **WACC**: calculated from first principles via CAPM (risk-free rate +
  beta × equity risk premium for cost of equity; interest expense ÷
  total debt for cost of debt), weighted by market cap vs. total debt.
- **Scenarios**: Bull/Base/Bear apply +5% / flat / -6% to the base
  growth assumption, with growth tapering slightly each year toward
  the 5-year projection horizon.
- **Terminal value**: standard Gordon Growth formula at a 2.5%
  long-run terminal growth rate.
- **Implied growth**: binary-searches for the FCF growth rate that
  would justify the current market price, so you can see what the
  market is "pricing in" versus your own assumptions.

## Known limitations (and how the tool handles them)

No DCF model — this one included — produces a reliable answer for
every company automatically. Real analysts adjust their approach by
hand for cases like these; this tool tries to at least flag them
instead of guessing silently:

- **Negative free cash flow.** Companies in a heavy investment or
  cash-burn phase (e.g. Intel's fab expansion) don't fit a model that
  assumes a positive, growing FCF base. The tool detects this and
  prints/writes a clear warning rather than returning a nonsensical
  negative "intrinsic value."
- **Captive finance arms.** Companies with large lending/financing
  divisions (e.g. GM Financial) carry debt on their balance sheet
  that's matched by loan receivables — not real operating leverage.
  Subtracting all of it as net debt overstates the drag on equity
  value. The tool flags this when net debt exceeds market cap, and a
  proper sum-of-the-parts valuation is recommended instead.
- **Growth beyond what the model can express.** For richly-valued
  growth names (e.g. mega-cap tech trading on future AI narratives),
  the market price may imply growth the model's search range (-10% to
  +40%) can't reach. Rather than silently reporting a misleading
  boundary value, the tool returns "n/a" and explains why.

## Notes

- Requires an internet connection (pulls live data from Yahoo Finance).
- If you see a `429 Too Many Requests` error, Yahoo is temporarily
  rate-limiting — wait a few minutes and try again, or run
  `pip install --upgrade yfinance` to get the latest version.
- For educational purposes only. Not financial advice.
