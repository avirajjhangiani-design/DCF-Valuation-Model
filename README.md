# DCF-Valuation-Model
# DCF Valuation Model

An automated discounted cash flow (DCF) valuation tool built in Python. Give it a stock ticker and it pulls live financial data, estimates WACC, projects free cash flow, and outputs an intrinsic value per share alongside a formatted Excel report.

## What it does

1. **Retrieves live data** from Yahoo Finance (via `yfinance`): cash flow statement, balance sheet, income statement, share price, market cap, and beta.
2. **Calculates historical free cash flow** (operating cash flow minus capital expenditure).
3. **Estimates WACC** using CAPM for the cost of equity and the company's own interest expense and debt for the cost of debt, weighted by market capitalisation and total debt.
4. **Projects free cash flow** over a set number of years (default 5), using a growth rate estimated from historical FCF (or one you supply).
5. **Discounts the cash flows** and calculates a terminal value using the Gordon growth method.
6. **Bridges enterprise value to equity value** (adds cash, subtracts debt) and divides by shares outstanding to get an intrinsic value per share.
7. **Compares this with the current share price** to show implied upside or downside.
8. **Exports an Excel workbook** with the full breakdown.

## Installation

```bash
git clone https://github.com/avirajjhangiani-design/DCF-Valuation-Model.git
cd DCF-Valuation-Model
pip install -r requirements.txt
```

Requires Python 3.8+. Main dependencies: `yfinance`, `pandas`, `numpy`, `openpyxl`.

## Usage

Basic run:

```bash
python dcf_valuation.py CAT
```

With custom assumptions:

```bash
python dcf_valuation.py CAT --years 7 --growth 0.06 --terminal-growth 0.02 --output cat_valuation.xlsx
```

### Options

| Flag | Description | Default |
|---|---|---|
| `ticker` | Stock ticker symbol (e.g. `CAT`, `MSFT`, `RR.L`) | required |
| `--years` | Number of years to project FCF | `5` |
| `--growth` | Manual FCF growth rate (e.g. `0.08` for 8%) | auto-estimated from history |
| `--terminal-growth` | Terminal (perpetuity) growth rate | `0.025` |
| `--risk-free-rate` | Risk-free rate used in CAPM | `0.043` |
| `--market-risk-premium` | Equity market risk premium | `0.055` |
| `--tax-rate` | Manual effective tax rate | auto-estimated from statements |
| `--output` | Output Excel filename | `TICKER_DCF_Valuation.xlsx` |

## Output

### Terminal summary

```
--- [TICKER] DCF Summary ---
Current price: [xxx.xx]
WACC: [x.xx]%
FCF growth rate used: [x.xx]%
Terminal growth rate: 2.50%
Enterprise value: [x,xxx,xxx,xxx]
Equity value: [x,xxx,xxx,xxx]
Intrinsic value / share: [xxx.xx]
Implied upside/downside: [x.x]% ([UNDERVALUED / OVERVALUED])
```

The tool also lists any fallback assumptions it had to use (see below), so you can see where the data was incomplete.

### Excel workbook

| Sheet | Contents |
|---|---|
| **Summary** | Price, WACC, growth assumptions, enterprise value, equity value, intrinsic value per share, upside/downside, verdict |
| **FCF Projection** | Projected FCF by year, discount factors, present values, and a line chart |
| **Historical Data** | Historical free cash flow by fiscal year |
| **Assumptions** | Projection years, risk-free rate, market risk premium, beta, terminal growth, tax rate |


## Methodology

- **Cost of equity:** risk-free rate + beta × market risk premium (CAPM).
- **WACC:** (E / (E + D)) × cost of equity + (D / (E + D)) × cost of debt × (1 − tax rate).
- **Projected FCF:** base year FCF grown at a constant rate for the projection period.
- **Terminal value:** FCF in final year × (1 + g) / (WACC − g), discounted back to present value.
- **Equity value:** enterprise value + cash − total debt.

### Automatic estimates and safeguards

When inputs aren't supplied or the data is missing, the model applies the following and reports it in the output:

- **Growth rate:** CAGR of historical positive FCF, capped between −10% and +25%. Falls back to 5% if unreliable.
- **WACC:** constrained to a 3%–20% range to avoid unrealistic results.
- **Beta:** defaults to 1.0 if unavailable.
- **Tax rate:** derived from the latest income statement; defaults to 21% if not derivable.
- **Cost of debt:** interest expense divided by total debt; defaults to 4.5% if not derivable.

## Limitations

- Uses a single-stage growth assumption for the projection period rather than a fade to terminal growth.
- Relies on Yahoo Finance data, which can be incomplete or inconsistent, particularly for non-US companies.
- Default risk-free rate, market risk premium, and fallback tax rate are US-based and should be adjusted for other markets.
- Terminal growth must be lower than WACC or the terminal value is not meaningful.
- The Excel output contains calculated values rather than live formulas.
- No sensitivity analysis yet.
- DCF valuations are highly sensitive to assumptions. Output should be read as an estimate, not investment advice.

## Possible improvements

- Sensitivity table (WACC vs terminal growth)
- Multi-stage growth projection
- Live Excel formulas in the output
- Comparable company cross-check

## Disclaimer

This project was built for educational purposes and does not constitute financial advice.
