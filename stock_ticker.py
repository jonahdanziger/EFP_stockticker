#!/usr/bin/env python3
"""
stock_ticker.py - Python port of stock_ticker.R

Original R author: Rish Singhania
Additional R editors: Zach Bethune, Kellie Forrester, Brian Thomas,
                      Travis Cyronek, Ryan Sherrard, Sarah Papich

Purpose
-------
Calculates 1-week, 4-week, and 52-week changes for a list of stocks specified
in a master CSV, builds the PCBT / Tri-County / California custom indexes, and
writes the results to an Excel workbook (one file per week) in ./weekly_reports.

Changes are computed Friday-to-Friday (or Thursday-to-Thursday on a holiday
week). Because of that, the program should be run early in the week to capture
last week's completed data.

Usage
-----
    python3 stock_ticker.py

Optional arguments:
    --stock-list PATH   Path to the master CSV (default: ./stock_list.csv)
    --outdir PATH       Output folder for the report (default: ./weekly_reports)
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd
import yfinance as yf

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

# Tickers to drop before downloading (delisted / acquired / error-producing).
# Mirrors the manual filter block in stock_ticker.R.
#   ATVI  - produces a download error
#   CAMP  - dropped in R
#   GPS   - dropped in R
#   JNPR  - Juniper Networks sold 2025-07-02, ticker retired
#   PPBI  - Pacific Premier Bancorp acquired by Columbia Banking 2025-08-31
#   CVGW  - dropped in R
EXCLUDED_TICKERS = {"ATVI", "CAMP", "GPS", "JNPR", "PPBI", "CVGW"}

# Number of most-recent weekly reports to keep in the archive; older ones are
# deleted after each run.
ARCHIVE_KEEP = 5

# Custom index definitions: (ticker, display name, membership-flag column).
INDEX_DEFS = [
    ("PCBTEFP", "Pacific Coast Business Times Stock Index", "pcbt.index"),
    ("TCEFP", "Tri County Stock Index", "tri.county.index"),
    ("CAEFP", "California Stock Index", "ca.index"),
]

# Column labels used in the output workbook.
COL_RENAME = {
    "ticker": "Symbol",
    "company.name": "Company Name",
    "n1.week.change": "Change from Previous Week",
    "n4.week.change": "Change from Last Four Weeks",
    "n52.week.change": "Change from Last 52 Weeks",
    "latest.quote": "Latest Quote",
}
OUTPUT_COLS = list(COL_RENAME.keys())


# --------------------------------------------------------------------------
# Data download and weekly conversion
# --------------------------------------------------------------------------

def download_daily_closes(tickers: list[str]) -> dict[str, pd.Series]:
    """Download daily adjusted-close series for each ticker from Yahoo Finance.

    Returns a dict mapping ticker -> daily Close series (indexed by date).
    Tickers that fail to download or return no data are skipped with a warning.
    """
    closes: dict[str, pd.Series] = {}
    print(f"Downloading {len(tickers)} tickers from Yahoo Finance...")
    # auto_adjust=True gives split/dividend-adjusted prices (Close column).
    data = yf.download(
        tickers,
        period="5y",
        interval="1d",
        auto_adjust=True,
        progress=False,
        group_by="ticker",
        threads=True,
    )

    for t in tickers:
        try:
            if len(tickers) == 1:
                series = data["Close"]
            else:
                series = data[t]["Close"]
            series = series.dropna()
            if series.empty:
                print(f"  WARNING: no data for {t}, skipping.")
                continue
            closes[t] = series
        except (KeyError, TypeError):
            print(f"  WARNING: could not read data for {t}, skipping.")
    print(f"Successfully downloaded {len(closes)} tickers.")
    return closes


def to_weekly(daily: pd.Series) -> pd.Series:
    """Convert a daily close series to a Friday-to-Friday weekly close series.

    Mirrors R's to.weekly + the incomplete-week drop: each weekly value is the
    last trading close of that week, indexed by that last trading date. If the
    most recent week's last trading day is not a Friday or Thursday, the week is
    treated as incomplete and dropped.
    """
    daily = daily.sort_index()
    # Group by ISO week ending Friday; take the last close in each week,
    # labelled with the actual last trading date of the week.
    grouped = daily.groupby(pd.Grouper(freq="W-FRI"))
    weekly_close = grouped.last().dropna()
    last_dates = grouped.apply(lambda s: s.index[-1] if len(s) else pd.NaT).dropna()
    weekly_close.index = last_dates.values

    if len(weekly_close):
        last_day = pd.Timestamp(weekly_close.index[-1]).day_name()
        if last_day not in ("Friday", "Thursday"):
            weekly_close = weekly_close.iloc[:-1]
    return weekly_close


def pct_change_k(weekly_close: pd.Series, k: int) -> float | None:
    """Most recent k-period arithmetic percentage change, in percent.

    Equivalent to R's last(Delt(close, k=k)) * 100. Requires more than k
    observations; otherwise returns None (NA).
    """
    if len(weekly_close) > k:
        return (weekly_close.iloc[-1] / weekly_close.iloc[-1 - k] - 1.0) * 100.0
    return None


# --------------------------------------------------------------------------
# Custom index construction
# --------------------------------------------------------------------------

def build_index_weekly(
    flag_col: str,
    stocklist: pd.DataFrame,
    daily_closes: dict[str, pd.Series],
    most_recent_eow: pd.Timestamp,
    horizon_weeks: int | None,
) -> pd.Series:
    """Build a weekly close series for a custom index.

    The index level on each date is the sum of member closes on that date. When
    `horizon_weeks` is given, membership is restricted to stocks whose history
    begins before (most_recent_eow - horizon_weeks), so the composition is
    consistent across the comparison window (mirrors the _1week/_4week/_52week
    filtered sums in stock_ticker.R). `horizon_weeks=None` uses all members.
    """
    members = stocklist.loc[stocklist[flag_col] == 1, "ticker"].tolist()

    if horizon_weeks is not None:
        cutoff = most_recent_eow - pd.Timedelta(weeks=horizon_weeks)
        members = [
            t for t in members
            if t in daily_closes and daily_closes[t].index[0] < cutoff
        ]
    else:
        members = [t for t in members if t in daily_closes]

    if not members:
        return pd.Series(dtype=float)

    frame = pd.concat({t: daily_closes[t] for t in members}, axis=1)
    # Sum members present on each date (min_count=1 avoids all-NaN -> 0).
    index_daily = frame.sum(axis=1, min_count=1).dropna()
    return to_weekly(index_daily)


# --------------------------------------------------------------------------
# Formatting helpers
# --------------------------------------------------------------------------

def fmt_pct(value) -> str:
    if value is None or pd.isna(value):
        return "NA"
    return f"{value:.2f}%"


def fmt_dollar(value) -> str:
    if value is None or pd.isna(value):
        return "NA"
    return f"${value:,.2f}"


def format_sheet(df: pd.DataFrame, dollar_quote: bool) -> pd.DataFrame:
    """Rename columns to their display labels and format numbers as text."""
    out = df[OUTPUT_COLS].copy()
    for col in ("n1.week.change", "n4.week.change", "n52.week.change"):
        out[col] = out[col].map(fmt_pct)
    if dollar_quote:
        out["latest.quote"] = out["latest.quote"].map(fmt_dollar)
    else:
        out["latest.quote"] = pd.to_numeric(out["latest.quote"], errors="coerce")
    return out.rename(columns=COL_RENAME)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def write_skipped_report(skipped: list[str], company_by_ticker: dict,
                         path: Path) -> None:
    """Write the list of tickers that returned no data (likely delisted).

    Always (re)writes the file: a non-empty file is the signal the CI workflow
    uses to notify you; an empty file means everything downloaded cleanly.
    """
    lines = [f"{t}\t{company_by_ticker.get(t, '')}" for t in skipped]
    path.write_text("\n".join(lines))
    if skipped:
        print("\n" + "=" * 60)
        print(f"ATTENTION: {len(skipped)} ticker(s) returned NO data "
              f"(possibly delisted or renamed):")
        for t in skipped:
            print(f"  - {t}  ({company_by_ticker.get(t, '')})")
        print("Update stock_list.csv / EXCLUDED_TICKERS if these are permanent.")
        print("=" * 60 + "\n")
    else:
        print("All tickers downloaded successfully (none skipped).")


def main() -> int:
    parser = argparse.ArgumentParser(description="Weekly PCBT stock report generator.")
    here = Path(__file__).resolve().parent
    parser.add_argument("--stock-list", default=str(here / "stock_list.csv"))
    parser.add_argument("--outdir", default=str(here / "weekly_reports"))
    parser.add_argument("--site-dir", default=str(here / "site"),
                        help="Directory for the website payload (data.json + downloads).")
    args = parser.parse_args()

    stock_list_path = Path(args.stock_list)
    outdir = Path(args.outdir)
    site_dir = Path(args.site_dir)
    outdir.mkdir(parents=True, exist_ok=True)

    if not stock_list_path.exists():
        print(f"ERROR: stock list not found at {stock_list_path}", file=sys.stderr)
        return 1

    # ---- Read and clean the master list ----
    stocklist = pd.read_csv(stock_list_path, dtype={"ticker": str})
    stocklist["ticker"] = stocklist["ticker"].str.strip()
    stocklist = stocklist[~stocklist["ticker"].isin(EXCLUDED_TICKERS)].reset_index(drop=True)

    for col in ("pcbt.index", "ca.index", "tri.county.index", "us.index"):
        stocklist[col] = pd.to_numeric(stocklist[col], errors="coerce").fillna(0).astype(int)

    # ---- Download daily data (keep the ^ prefix; strip it for display) ----
    tickers = stocklist["ticker"].tolist()
    company_by_ticker = dict(zip(stocklist["ticker"], stocklist["company.name"]))
    daily_closes_raw = download_daily_closes(tickers)

    # Any expected ticker that returned no data (likely delisted/renamed).
    skipped = [t for t in tickers if t not in daily_closes_raw]
    write_skipped_report(skipped, company_by_ticker, here / "skipped_tickers.txt")

    # Strip the ^ that Yahoo uses for index symbols, in both data and list.
    daily_closes = {t.lstrip("^"): s for t, s in daily_closes_raw.items()}
    stocklist["ticker"] = stocklist["ticker"].str.lstrip("^")

    # ---- Most recent completed end-of-week, anchored to NASDAQ (IXIC) ----
    if "IXIC" not in daily_closes:
        print("ERROR: NASDAQ (^IXIC) data unavailable; cannot anchor the week.",
              file=sys.stderr)
        return 1
    nasdaq_weekly = to_weekly(daily_closes["IXIC"])
    most_recent_eow = pd.Timestamp(nasdaq_weekly.index[-1])
    print(f"Most recent completed week ends: {most_recent_eow.date()}")

    # ---- Append custom index rows to the stock list ----
    index_rows = []
    for tkr, name, _flag in INDEX_DEFS:
        index_rows.append({
            "ticker": tkr, "company.name": name,
            "pcbt.index": 0, "ca.index": 0, "tri.county.index": 0, "us.index": 1,
        })
    stocklist = pd.concat([stocklist, pd.DataFrame(index_rows)], ignore_index=True)

    # ---- Pre-compute weekly close series for the custom indexes ----
    index_tickers = {tkr for tkr, _, _ in INDEX_DEFS}
    index_weekly_full: dict[str, pd.Series] = {}
    index_change: dict[str, dict[int, float | None]] = {}
    for tkr, _name, flag_col in INDEX_DEFS:
        index_weekly_full[tkr] = build_index_weekly(
            flag_col, stocklist, daily_closes, most_recent_eow, horizon_weeks=None
        )
        index_change[tkr] = {}
        for k in (1, 4, 52):
            wk = build_index_weekly(
                flag_col, stocklist, daily_closes, most_recent_eow, horizon_weeks=k
            )
            index_change[tkr][k] = pct_change_k(wk, k)

    # ---- Compute changes and latest quote for every row ----
    stocklist["n1.week.change"] = pd.NA
    stocklist["n4.week.change"] = pd.NA
    stocklist["n52.week.change"] = pd.NA
    stocklist["latest.quote"] = pd.NA

    for i, row in stocklist.iterrows():
        tkr = row["ticker"]

        if tkr in index_tickers:
            weekly_close = index_weekly_full.get(tkr, pd.Series(dtype=float))
            n1 = index_change[tkr][1]
            n4 = index_change[tkr][4]
            n52 = index_change[tkr][52]
        elif tkr in daily_closes:
            weekly_close = to_weekly(daily_closes[tkr])
            n1 = pct_change_k(weekly_close, 1)
            n4 = pct_change_k(weekly_close, 4)
            n52 = pct_change_k(weekly_close, 52)
        else:
            continue  # no data; leave as NA

        stocklist.at[i, "n1.week.change"] = n1
        stocklist.at[i, "n4.week.change"] = n4
        stocklist.at[i, "n52.week.change"] = n52
        if len(weekly_close):
            stocklist.at[i, "latest.quote"] = float(weekly_close.iloc[-1])

    # ---- Slice groups ----
    nat = stocklist[stocklist["us.index"] == 1]
    pcbt = stocklist[stocklist["pcbt.index"] == 1]

    # Winners / losers within the PCBT group, by 1-week change (drop NA).
    pcbt_ranked = pcbt.dropna(subset=["n1.week.change"]).copy()
    pcbt_ranked["n1.week.change"] = pd.to_numeric(pcbt_ranked["n1.week.change"])
    top5_winners = pcbt_ranked.sort_values("n1.week.change", ascending=False).head(5)
    top5_losers = pcbt_ranked.sort_values("n1.week.change", ascending=True).head(5)

    # ---- Filename from the last S&P 500 (GSPC) weekly date ----
    if "GSPC" in daily_closes:
        report_date = pd.Timestamp(to_weekly(daily_closes["GSPC"]).index[-1])
    else:
        report_date = most_recent_eow
    fname = f"PCBT_{report_date.strftime('%Y_%m_%d')}.xlsx"
    out_path = outdir / fname

    # ---- Build the four display sheets once, reuse for workbook + website ----
    # (id, sheet label, formatted DataFrame)
    sheets = [
        ("best_performers", "Best Performers", format_sheet(top5_winners, dollar_quote=True)),
        ("worst_performers", "Worst Performers", format_sheet(top5_losers, dollar_quote=True)),
        ("pcbt", "Pacific Coast Business Times", format_sheet(pcbt, dollar_quote=True)),
        ("index", "Index", format_sheet(nat, dollar_quote=False)),
    ]

    # ---- Write the archive workbook to weekly_reports/ ----
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for _sid, label, df in sheets:
            df.to_excel(writer, sheet_name=label, index=False)
    print(f"Report written to: {out_path}")

    # ---- Prune the archive to the newest ARCHIVE_KEEP reports ----
    archive = prune_archive(outdir, keep=ARCHIVE_KEEP)
    print(f"Archive holds {len(archive)} report(s): "
          f"{', '.join(p.name for p, _ in archive)}")

    # ---- Build the website payload (data.json + per-tab / full workbooks) ----
    build_site(sheets, out_path, report_date, site_dir, archive)
    print(f"Website payload written to: {site_dir}")
    return 0


def _report_date_from_name(path: Path) -> pd.Timestamp | None:
    """Parse the date out of a 'PCBT_YYYY_MM_DD.xlsx' filename."""
    stem = path.stem  # PCBT_YYYY_MM_DD
    parts = stem.split("_")
    if len(parts) != 4 or parts[0] != "PCBT":
        return None
    try:
        return pd.Timestamp(f"{parts[1]}-{parts[2]}-{parts[3]}")
    except ValueError:
        return None


def prune_archive(outdir: Path, keep: int) -> list[tuple[Path, pd.Timestamp]]:
    """Keep only the `keep` most recent PCBT_*.xlsx reports; delete the rest.

    Returns the kept reports as (path, date) tuples, newest first.
    """
    reports = []
    for p in outdir.glob("PCBT_*.xlsx"):
        d = _report_date_from_name(p)
        if d is not None:
            reports.append((p, d))
    reports.sort(key=lambda pd_: pd_[1], reverse=True)  # newest first

    for path, _d in reports[keep:]:
        path.unlink()
        print(f"  Deleted old report: {path.name}")

    return reports[:keep]


def build_site(sheets, full_workbook_path: Path, report_date: pd.Timestamp,
               site_dir: Path,
               archive: list[tuple[Path, pd.Timestamp]] | None = None) -> None:
    """Write the static-site payload consumed by index.html.

    Produces, under `site_dir`:
      - downloads/<full_workbook_name>      full 4-sheet workbook (copy)
      - downloads/<tab_id>.xlsx             one single-sheet workbook per tab
      - downloads/<archived reports>.xlsx   the kept archive workbooks (copies)
      - data.json                           metadata + table rows + archive list
    index.html (committed alongside) fetches data.json at load time.
    """
    import json
    import shutil

    downloads = site_dir / "downloads"
    # Rebuild the downloads folder from scratch so deleted archives don't linger.
    if downloads.exists():
        shutil.rmtree(downloads)
    downloads.mkdir(parents=True, exist_ok=True)

    # Full workbook: copy the archive into the site's downloads folder.
    full_name = full_workbook_path.name
    shutil.copyfile(full_workbook_path, downloads / full_name)

    tabs = []
    for sid, label, df in sheets:
        # One single-sheet workbook per tab ("Download this tab").
        tab_file = f"{sid}.xlsx"
        with pd.ExcelWriter(downloads / tab_file, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=label[:31], index=False)

        # Rows for the web table (NaN -> None so it serialises as JSON null).
        safe = df.astype(object).where(pd.notna(df), None)
        tabs.append({
            "id": sid,
            "label": label,
            "tab_file": tab_file,
            "columns": list(df.columns),
            "rows": safe.to_dict(orient="records"),
        })

    # Archive: copy each kept report into downloads and list it (newest first).
    archive_entries = []
    for path, adate in (archive or []):
        if path.name != full_name:  # current report already copied above
            shutil.copyfile(path, downloads / path.name)
        archive_entries.append({
            "date": adate.strftime("%Y-%m-%d"),
            "date_pretty": adate.strftime("%B %-d, %Y"),
            "file": path.name,
        })

    payload = {
        "report_date": report_date.strftime("%Y-%m-%d"),
        "report_date_pretty": report_date.strftime("%B %-d, %Y"),
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "full_report_file": full_name,
        "tabs": tabs,
        "archive": archive_entries,
    }
    with open(site_dir / "data.json", "w") as f:
        json.dump(payload, f, indent=2)


if __name__ == "__main__":
    raise SystemExit(main())
