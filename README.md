# EFP Stock Ticker

Weekly PCBT (Pacific Coast Business Times) stock report. Calculates 1-week,
4-week, and 52-week changes for a list of stocks, builds the PCBT / Tri-County /
California custom indexes, and publishes the results as:

- an **Excel workbook** (four sheets), archived in [`weekly_reports/`](weekly_reports), and
- a **web page** with four data tabs (per-tab / full-report download buttons) plus
  an **Archive** tab listing recent reports.

The whole thing regenerates automatically **every Saturday** via GitHub Actions
and publishes to GitHub Pages.

Only the **5 most recent** reports are kept; older ones are deleted automatically
each run (change `ARCHIVE_KEEP` at the top of `stock_ticker.py` to keep more).

## Files

| File | Purpose |
|------|---------|
| `stock_ticker.py` | Main program: downloads data, computes changes, writes the workbook + website payload. |
| `stock_ticker.R` | The original R version (kept for reference). |
| `stock_list.csv` | Master list of tickers and their index memberships. |
| `site/index.html` | The four-tab web page (reads `site/data.json`). |
| `weekly_reports/` | Archive of every weekly `.xlsx` report. |
| `.github/workflows/weekly-report.yml` | Saturday automation + Pages deploy. |

## Run it locally

```bash
pip install -r requirements.txt
python3 stock_ticker.py
```

This writes `weekly_reports/PCBT_YYYY_MM_DD.xlsx` and rebuilds the `site/`
payload (`data.json` + `downloads/`). To preview the page:

```bash
cd site && python3 -m http.server 8000
# then open http://localhost:8000
```

## One-time GitHub setup (needed to go live)

1. Push this repo to GitHub.
2. In the repo, go to **Settings → Pages** and set **Source** to
   **GitHub Actions**.
3. In **Settings → Actions → General → Workflow permissions**, select
   **Read and write permissions** (lets the weekly job commit the archived report).
4. Trigger the first run: **Actions → Weekly Stock Report → Run workflow**
   (or wait for Saturday). The site URL appears in the workflow's `deploy` step.

## Notes

- Prices are dividend/split-adjusted (`yfinance` with `auto_adjust=True`), so
  52-week figures can differ slightly from the original R script, which used raw
  Yahoo closes.
- The email step from the R version is intentionally not ported.
- The excluded-tickers list (delisted/acquired names) lives at the top of
  `stock_ticker.py` — edit it there as the roster changes.
