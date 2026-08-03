# DCF Valuation Tool

A single-file Python DCF (Discounted Cash Flow) valuation model. Give it a
ticker, and it fetches live financial data, calculates WACC from first
principles, projects free cash flow under Bull/Base/Bear scenarios, and
exports a formatted 3-sheet Excel valuation model.

Built as one file on purpose — no folder structure to navigate, no
`src` imports to break. Just download, install requirements, edit the
ticker, and run.

## What it produces

A 3-sheet Excel workbook:

1. **Cover** — scenario summary, current market price, implied growth,
   and WACC breakdown
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
Open `dcf_valuation.py`'s folder in Jupyter, then in a notebook cell:
```python
from dcf_valuation import run_valuation
filepath = run_valuation("AAPL", manual_growth=0.10)
print(filepath)
```

## Changing the ticker

Open `dcf_valuation.py`, scroll to the very bottom, and edit this line:
```python
run_valuation("AAPL", manual_growth=0.10)
```
Swap `"AAPL"` for any public ticker. `manual_growth` is your assumed
Base Case FCF growth rate — leave it out entirely (or set to `None`) to
let the model estimate it from historical FCF growth instead.

## Notes

- Requires an internet connection (pulls live data from Yahoo Finance).
- If you see a `429 Too Many Requests` error, Yahoo is temporarily
  rate-limiting — wait a few minutes and try again, or run
  `pip install --upgrade yfinance` to get the latest version.
- For educational purposes only. Not financial advice.
