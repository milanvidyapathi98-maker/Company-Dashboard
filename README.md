# 📊 Financial Dashboard

A browser dashboard that turns a company's **income statement, balance sheet and cash flow
statement** into KPIs, charts and a ratio scorecard, alongside its share price, market cap and valuation.

**Live app:** *https://milanvidyapathi98-maker.github.io/Company-Dashboard/*

> AI-based learning and portfolio project. Not investment advice.

\---

## Features

* **10 companies built in**: NVIDIA, Apple, Microsoft, Alphabet, Amazon, Meta, Tesla, Broadcom, AMD and Netflix. Five years of annual data, ready without any upload
* **Open your own Excel file**: any workbook with `Income Statement`, `Balance Sheet` and `Cash Flow` sheets. The file is read in the browser and never uploaded
* **KPI cards with sparklines**: revenue, net income, diluted EPS, free cash flow, operating margin and net cash, each with the change vs. the prior year
* **Share price and valuation**: price chart (1M / 6M / 1Y / 5Y), a performance table from 1 day to 5 years, 52-week range, market cap, P/E, P/S, P/B, FCF yield and dividend yield
* **12 charts**: revenue and net income, margins, revenue growth, cash flow, shareholder returns, operating expenses, asset mix, cash vs. debt, working-capital days, EPS and three gauges
* **Ratio scorecard**: 17 ratios in four groups (profitability, liquidity and leverage, efficiency, cash quality), with a trend badge for each
* **Works offline and installs as an app**: a PWA with a service worker, so it can be added to the home screen or desktop

## How it works

|Step|File|What it does|
|-|-|-|
|1|`config.js`|Holds the company list for the drop-down, the default ticker and an optional Twelve Data API key|
|2|`data.js`|Stores the financial statements of the built-in companies (USD millions), taken from SEC 10-K filings and Excel workbooks|
|3|`script.js`|Reads an Excel file with SheetJS and matches row names to a fixed list (e.g. *Revenue*, *Total assets*, *Net cash from operating activities*). It works out derived figures and ratios, then draws everything with Chart.js|
|4|`prices-data.js`|A saved snapshot of about five years of daily closing prices for every built-in company|
|5|`script.js`|Fetches live prices from Twelve Data if a key is set and keeps them in the browser cache. Otherwise it shows the newer of the snapshot and the cache|
|6|`sw.js`, `manifest.json`|Service worker and app manifest for offline use and installation (network first, cache when offline)|

### Definitions

* **Free cash flow** = operating cash flow − capital expenditures (acquisitions are not deducted)
* **Net cash** = cash + marketable securities − total debt (debt includes lease liabilities where the data has them)
* **Market cap** ≈ latest price × diluted shares of the latest fiscal year; P/E uses that year's diluted EPS
* Company-specific choices (e.g. how Amazon's or Netflix's lines are mapped) are listed in the footnotes under the dashboard

## Run it on your own computer

```bash
git clone https://github.com/<your-username>/financial-dashboard.git
cd financial-dashboard
python -m http.server 8000
```

Then open `http://localhost:8000` in your browser. The **Live Server** extension in VS Code works too.
Opening `index.html` directly also works, but offline mode, installation and the automatic loading of `data.xlsx` need a local server.

**Live prices (optional):** get a free key at [twelvedata.com](https://twelvedata.com) and paste it into `TWELVE\_DATA\_KEY` in `config.js`.
Without a key the dashboard uses the saved prices in `prices-data.js`.

> ⚠️ The key in `config.js` can be seen by anyone who opens the site. Don't commit a real key to a public repository.

## Use your own Excel file

Click **Open Excel file**, or put a file named `data.xlsx` in the project folder. Each sheet needs a row that starts with
`Fiscal year` followed by the years, then one row per line item. Required rows:

* **Income Statement**: Revenue, Operating income, Net income, Weighted avg. shares - diluted (millions), EPS - diluted ($)
* **Balance Sheet**: Cash \& cash equivalents, Total current assets, Total assets, Total current liabilities, Total shareholders' equity
* **Cash Flow**: Net cash from operating activities, Capital expenditures

Other rows (R\&D, SG\&A, inventories, debt, buybacks, dividends and more) are optional. Any optional row that is missing counts as 0 and is listed in the footnotes.
An optional `Settings` sheet can set `Company`, `Ticker`, `Currency` (e.g. EUR, GBP, LKR), `Subtitle` and `Note`.

## Add a company

1. Add `{ symbol: "INTC", name: "Intel" }` to `TICKERS` in `config.js`
2. Add an entry with the same structure as the others to `data.js` (or load an Excel file for it)
3. Add its daily prices to `prices-data.js`, or rely on the Twelve Data key
4. Bump `CACHE` in `sw.js` (e.g. `findash-v3`) so installed copies pick up the new files

## Project structure

```
index.html        page layout
style.css         dark theme and responsive grid
script.js         Excel reader, calculations, charts, prices
config.js         company list and API key
data.js           built-in financial statements
prices-data.js    saved price history
sw.js             service worker (offline)
manifest.json     app manifest
icon-192.png      app icons
icon-512.png
```

## Tech stack

HTML · CSS · JavaScript · Chart.js · SheetJS (xlsx) · Twelve Data API · Service Worker / PWA

## Limitations

* Annual data only, no quarterly figures or trailing twelve months.
* Companies don't all report the same lines (e.g. Microsoft and Meta show no separate interest expense), so some ratios read *n/a* or are not fully comparable between companies.
* The saved price snapshot only changes when `prices-data.js` is regenerated, and the free Twelve Data plan has request limits.
* Market cap is approximate because it uses diluted shares from the last fiscal year rather than today's share count.

## Author

Milan Vidyapathige · M.A. International Management and Information Systems · B.Sc. Finance

