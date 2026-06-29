"""
chart_generator_hourly.py  —  NSE Live Edition (HOURLY)
=========================================================
Combines:
  • STEP 1  — Live NSE "Stocks Traded" download  (from nse_pipeline.py)
  • STEP 2  — Filter EQ series, Value > ₹10 Cr   (from nse_pipeline.py)
  • STEP 3  — TradingView-style dark HOURLY charts

Charts contain:
  - Candlestick price chart  (wider, more visible bars)
  - 9-bar EMA overlay
  - Standard MACD subplot (12, 26, 9)
  - 20-bar recent low annotation (dashed line + % below current close)

Requirements:
    pip install selenium webdriver-manager yfinance pandas openpyxl matplotlib

Usage:
    python chart_generator_hourly.py                   # headless browser
    python chart_generator_hourly.py --visible         # visible browser (local debug)
    python chart_generator_hourly.py --from-csv FILE   # skip browser, use saved CSV

Notes on yfinance hourly data:
  - yfinance returns 1h data up to ~730 calendar days back, but quality
    degrades beyond ~60 days. PERIOD_DAYS=60 is a safe default.
  - Timestamps are UTC; converted to IST (UTC+5:30) for display.
"""

import sys
import os
import glob
import json
import time
import shutil
import datetime
import tempfile
import argparse
import traceback
import warnings
from datetime import timedelta

import pandas as pd
import numpy as np

# Selenium (optional - graceful fallback if missing)
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, WebDriverException
    from webdriver_manager.chrome import ChromeDriverManager
    SELENIUM_OK = True
except ImportError:
    SELENIUM_OK = False

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from matplotlib.lines import Line2D
import yfinance as yf

warnings.filterwarnings("ignore")


# ===============================================================
#  CONFIG  - NSE scrape settings  (unchanged)
# ===============================================================

TRADED_VALUE_MIN_CR = 10

NSE_HOME  = "https://www.nseindia.com"
NSE_PAGE  = "https://www.nseindia.com/market-data/stocks-traded"
NSE_APIS  = [
    "https://www.nseindia.com/api/live-analysis-stocksTraded",
    "https://www.nseindia.com/json/liveAnalysis/stocks-traded.json",
]

PAGE_WAIT     = 30
API_SETTLE    = 10
DOWNLOAD_WAIT = 40

SYNC_XHR = """
var xhr = new XMLHttpRequest();
xhr.open('GET', arguments[0], false);
xhr.setRequestHeader('Accept', 'application/json, text/plain, */*');
xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
xhr.setRequestHeader('Referer',
    'https://www.nseindia.com/market-data/stocks-traded');
try {
    xhr.send(null);
    return {status: xhr.status, body: xhr.responseText};
} catch(e) {
    return {status: -1, body: e.toString()};
}
"""

# ===============================================================
#  CONFIG  - Chart settings  (HOURLY adaptations marked with *)
# ===============================================================

EXCHANGE_SFX    = ".NS"
OUTPUT_DIR      = "YF_H_TO_10Cr"        # * new output folder
INTERVAL        = "1h"                   # * hourly candles
PERIOD_DAYS     = 60                     # * ~60 days of hourly data
MAX_BARS        = 500                    # * show up to 500 hourly bars

EMA_PERIOD      = 9
MACD_FAST       = 12
MACD_SLOW       = 26
MACD_SIGNAL     = 9
RECENT_LOW_BARS = 20

CANDLE_BODY_WIDTH = 0.6                 # * slightly narrower for hourly density
CANDLE_WICK_WIDTH = 0.10

IST_OFFSET = timedelta(hours=5, minutes=30)   # * UTC to IST

STYLE = {
    "bg":          "#131722",
    "panel_bg":    "#1E222D",
    "grid":        "#2A2E39",
    "up_candle":   "#26A69A",
    "down_candle": "#EF5350",
    "wick_up":     "#26A69A",
    "wick_down":   "#EF5350",
    "ema_color":   "#FF9800",
    "macd_line":   "#2962FF",
    "signal_line": "#FF6D00",
    "hist_up":     "#26A69A",
    "hist_down":   "#EF5350",
    "text":        "#D1D4DC",
    "subtext":     "#787B86",
    "border":      "#2A2E39",
    "recent_low":  "#FFD700",
}


# ===============================================================
#  STEP 1A - BROWSER SETUP  (unchanged)
# ===============================================================

def build_driver(headless, download_dir):
    if not SELENIUM_OK:
        print("[ERROR] selenium / webdriver-manager not installed.")
        print("  Fix: pip install selenium webdriver-manager")
        sys.exit(1)

    opts = Options()
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-setuid-sandbox")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
    if headless:
        opts.add_argument("--headless=new")

    opts.add_experimental_option("prefs", {
        "download.default_directory":   download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade":   True,
        "safebrowsing.enabled":         True,
    })

    try:
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opts,
        )
        print("  ChromeDriver ready (webdriver-manager)")
        return driver
    except Exception as e:
        print(f"  [WARN] webdriver-manager failed: {e}")

    try:
        driver = webdriver.Chrome(options=opts)
        print("  ChromeDriver ready (system)")
        return driver
    except Exception as e:
        print(f"\n[ERROR] Chrome unavailable: {e}")
        sys.exit(1)


def patch_driver(driver):
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": "Object.defineProperty(navigator,'webdriver',"
                       "{get:()=>undefined});"}
        )
    except Exception:
        pass


def warm_session(driver):
    print("  [1/3]  Loading NSE homepage ...")
    driver.get(NSE_HOME)
    time.sleep(4)
    print(f"         Cookies: {[c['name'] for c in driver.get_cookies()]}")

    print("  [2/3]  Loading Stocks Traded page ...")
    driver.get(NSE_PAGE)
    try:
        WebDriverWait(driver, PAGE_WAIT).until(
            EC.presence_of_element_located(
                (By.XPATH,
                 "//*[@id='cm_9'] | "
                 "//h2[contains(text(),'Stocks Traded')] | "
                 "//*[contains(text(),'Stocks Traded') and "
                 "    not(contains(@class,'nav'))]")
            )
        )
        print("         Page loaded")
    except TimeoutException:
        print("         [WARN] timeout - continuing")

    print(f"         Settling {API_SETTLE}s for XHR to complete ...")
    time.sleep(API_SETTLE)
    print(f"         Cookies: {[c['name'] for c in driver.get_cookies()]}")


# ===============================================================
#  STEP 1B - XHR DATA FETCH  (unchanged)
# ===============================================================

def _find_records_in_json(payload):
    if isinstance(payload, list) and payload and isinstance(payload[0], dict):
        if any(k in payload[0] for k in ("symbol", "Symbol", "SYMBOL")):
            return payload
    if isinstance(payload, dict):
        for key in ("data", "stocksTradedData", "result", "rows",
                    "stockData", "DATA", "records", "stocks", "dataList"):
            val = payload.get(key)
            if isinstance(val, list) and val and isinstance(val[0], dict):
                return val
    return []


def fetch_via_xhr(driver):
    print("  [3/3]  Calling NSE API via sync XHR ...")
    for url in NSE_APIS:
        print(f"         -> {url}")
        try:
            res    = driver.execute_script(SYNC_XHR, url)
            status = res.get("status", -1)
            body   = res.get("body", "")
            print(f"           HTTP {status}  |  {len(body):,} chars")

            if status != 200 or not body:
                continue

            payload = json.loads(body)
            if isinstance(payload, dict):
                print(f"           JSON keys: {list(payload.keys())}")

            records = _find_records_in_json(payload)
            if records:
                print(f"  {len(records)} records via XHR")
                print(f"     Fields: {list(records[0].keys())[:10]}")
                return records
            else:
                print("           No stock list found in response")

        except json.JSONDecodeError as e:
            print(f"           JSON error: {e}")
        except Exception as e:
            print(f"           Error: {e}")
    return []


# ===============================================================
#  STEP 1C - CSV BUTTON FALLBACK  (unchanged)
# ===============================================================

def _wait_for_csv(dl_dir):
    print(f"         Waiting {DOWNLOAD_WAIT}s for file", end="", flush=True)
    deadline = time.time() + DOWNLOAD_WAIT
    while time.time() < deadline:
        time.sleep(1)
        print(".", end="", flush=True)
        files = [
            f for f in
            glob.glob(os.path.join(dl_dir, "*.csv")) +
            glob.glob(os.path.join(dl_dir, "*.CSV"))
            if not f.endswith(".crdownload")
        ]
        if files:
            latest = max(files, key=os.path.getmtime)
            print(f"\n  {os.path.basename(latest)}")
            return latest
    print("\n  Timed out.")
    return ""


def fetch_via_csv_button(driver, dl_dir):
    print("  CSV button fallback ...")
    xpaths = [
        "//a[contains(@onclick,'StocksTraded-download')]",
        "//a[contains(@onclick,'StocksTraded')]",
        ".//a[.//img[contains(@src,'xls') or contains(@src,'csv')]]",
        "//a[contains(@onclick,'download') and "
        "    not(contains(@onclick,'First')) and "
        "    not(contains(@onclick,'Prev')) and "
        "    not(contains(@onclick,'Next'))]",
    ]
    for xpath in xpaths:
        try:
            for el in driver.find_elements(By.XPATH, xpath):
                if el.is_displayed():
                    print(f"  Clicking: {el.get_attribute('outerHTML')[:100]}")
                    driver.execute_script(
                        "arguments[0].scrollIntoView(true);", el)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", el)
                    path = _wait_for_csv(dl_dir)
                    if path:
                        return path
        except Exception:
            continue
    try:
        driver.execute_script("downloadCSV('StocksTraded-download');")
        return _wait_for_csv(dl_dir)
    except Exception as e:
        print(f"  JS call failed: {e}")
    return ""


# ===============================================================
#  STEP 1D - NORMALISE TO STANDARD DATAFRAME  (unchanged)
# ===============================================================

def safe_num(v):
    try:
        return float(str(v).replace(",", "").replace("\u2013", "0")
                     .replace("\u2212", "0").strip())
    except Exception:
        return 0.0


def normalise_json(records):
    rows = []
    for d in records:
        sym = str(d.get("symbol", d.get("Symbol", ""))).strip()
        if not sym:
            continue
        tv_raw = safe_num(d.get("totalTradedValue",
                          d.get("tradedValue", 0)))
        vol    = safe_num(d.get("totalTradedVolume",
                          d.get("tradedQuantity", 0)))
        rows.append({
            "Symbol":           sym,
            "Company":          str(d.get("companyName", "")).strip(),
            "Series":           str(d.get("series", "EQ")).strip(),
            "LTP (Rs)":         round(safe_num(d.get("lastPrice",
                                     d.get("closePrice", 0))), 2),
            "% Change":         round(safe_num(d.get("pChange", 0)), 2),
            "Mkt Cap (Rs Cr)":  round(safe_num(d.get("marketCap",
                                     d.get("market_cap", 0))), 2),
            "Volume (Lakhs)":   round(vol / 1e5, 2),
            "Value (Rs Crores)": round(tv_raw / 1e7, 2),
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df.sort_values("Value (Rs Crores)", ascending=False, inplace=True)
        df.reset_index(drop=True, inplace=True)
        df.index += 1
    return df


def normalise_csv(path):
    try:
        df = pd.read_csv(path, thousands=",")
    except Exception as e:
        print(f"  CSV read error: {e}")
        return pd.DataFrame()

    df.columns = df.columns.str.strip()
    print(f"  CSV columns: {list(df.columns)}")
    if not df.empty:
        print(f"  First row:   {df.iloc[0].to_dict()}")

    col_map = {}
    for col in df.columns:
        cl = col.lower().strip()
        if cl == "symbol":
            col_map[col] = "Symbol"
        elif cl == "series":
            col_map[col] = "Series"
        elif cl in ("ltp", "last price", "close", "lastprice"):
            col_map[col] = "LTP (Rs)"
        elif cl in ("%chng", "%change", "% change", "pchange",
                    "% chng", "per change", "%chg"):
            col_map[col] = "% Change"
        elif "mkt cap" in cl or "market cap" in cl:
            col_map[col] = "Mkt Cap (Rs Cr)"
        elif "volume" in cl:
            col_map[col] = "Volume (Lakhs)"
        elif "value" in cl:
            col_map[col] = "Value (Rs Crores)"
        elif "company" in cl or cl == "name":
            col_map[col] = "Company"

    df.rename(columns=col_map, inplace=True)
    print(f"  Mapped to:   {list(df.columns)}")

    for col in ["LTP (Rs)", "% Change", "Mkt Cap (Rs Cr)",
                "Volume (Lakhs)", "Value (Rs Crores)"]:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str)
                .str.replace(",", "", regex=False)
                .str.strip(),
                errors="coerce"
            ).fillna(0)

    for col, default in [("Company", ""), ("Series", "EQ"),
                         ("Mkt Cap (Rs Cr)", 0.0)]:
        if col not in df.columns:
            df[col] = default

    if "Value (Rs Crores)" in df.columns:
        df.sort_values("Value (Rs Crores)", ascending=False, inplace=True)
        top = df["Value (Rs Crores)"].iloc[0]
        print(f"  Top Value (Rs Crores): {top:,.2f}  "
              f"({'looks correct' if top > 10 else 'suspiciously low'})")

    df.reset_index(drop=True, inplace=True)
    df.index += 1
    print(f"  {len(df)} rows loaded from CSV")
    return df


# ===============================================================
#  STEP 1 - MASTER DOWNLOAD FUNCTION  (unchanged)
# ===============================================================

def download_nse_data(headless, from_csv):
    if from_csv:
        print(f"\n  Loading manual CSV: {from_csv}")
        df = normalise_csv(from_csv)
        if df.empty:
            print("  [ERROR] CSV empty or unreadable.")
            sys.exit(1)
        return df

    if not SELENIUM_OK:
        print("[ERROR] selenium not installed.")
        sys.exit(1)

    dl_dir = tempfile.mkdtemp()
    driver = build_driver(headless, dl_dir)
    driver.set_page_load_timeout(60)
    patch_driver(driver)
    df = pd.DataFrame()

    try:
        warm_session(driver)
        records = fetch_via_xhr(driver)
        if records:
            df = normalise_json(records)
        if df.empty:
            print("\n  XHR returned no data - trying CSV button ...")
            csv_path = fetch_via_csv_button(driver, dl_dir)
            if csv_path:
                df = normalise_csv(csv_path)
    except WebDriverException as e:
        print(f"\n[ERROR] WebDriver: {e}")
    finally:
        driver.quit()
        shutil.rmtree(dl_dir, ignore_errors=True)
        print("  Browser closed.")

    if df.empty:
        print("\n  Could not retrieve data from NSE.")
        print(f"  MANUAL FALLBACK:")
        print(f"  1. Open {NSE_PAGE} in Chrome")
        print("  2. Click the down arrow CSV button")
        print("  3. Run: python chart_generator_hourly.py --from-csv StocksTraded.csv")
        sys.exit(1)

    return df


# ===============================================================
#  STEP 2 - FILTER  (unchanged, uses renamed column keys)
# ===============================================================

def filter_stocks(df):
    # Support both column naming conventions (Rs vs rupee symbol)
    val_col = "Value (Rs Crores)" if "Value (Rs Crores)" in df.columns else "Value (\u20b9 Crores)"
    ser_col = "Series"
    ltp_col = "LTP (Rs)" if "LTP (Rs)" in df.columns else "LTP (\u20b9)"
    pct_col = "% Change"
    cmp_col = "Company"

    print(f"\n  Top 5 by Value before filter:")
    top5 = df[df[val_col] > 0].nlargest(5, val_col)
    for _, r in top5.iterrows():
        print(f"    {r['Symbol']:<12}  Rs{r[val_col]:>10,.2f} Cr  Series={r[ser_col]}")

    mask = (
        (df[ser_col].str.strip().str.upper() == "EQ") &
        (df[val_col] > TRADED_VALUE_MIN_CR)
    )
    out = df[mask].copy()
    out.sort_values(val_col, ascending=False, inplace=True)
    out.reset_index(drop=True, inplace=True)

    # Normalise column names for downstream use
    out.rename(columns={
        val_col: "Value (Rs Crores)",
        ltp_col: "LTP (Rs)",
        cmp_col: "Company",
    }, inplace=True, errors="ignore")

    return out


# ===============================================================
#  STEP 3 - CHART HELPERS
# ===============================================================

def ema(series, period):
    return series.ewm(span=period, adjust=False).mean()


def macd_calc(close, fast=12, slow=26, signal=9):
    ml = ema(close, fast) - ema(close, slow)
    sl = ema(ml, signal)
    return ml, sl, ml - sl


def to_ist(ts):
    """Convert a pandas Timestamp (possibly tz-aware) to IST datetime."""
    if hasattr(ts, "tzinfo") and ts.tzinfo is not None:
        return ts.to_pydatetime().astimezone(
            datetime.timezone(IST_OFFSET)
        )
    return ts.to_pydatetime() + IST_OFFSET


# ===============================================================
#  STEP 3 - CHART  (* hourly-specific changes)
# ===============================================================

def plot_chart(symbol, df, output_path):
    s  = STYLE
    xs = np.arange(len(df))

    # * Convert index to IST for display
    ist_times = [to_ist(ts) for ts in df.index]

    ema9                  = ema(df["Close"], EMA_PERIOD)
    macd_l, sig, hist     = macd_calc(df["Close"], MACD_FAST, MACD_SLOW, MACD_SIGNAL)

    # Recent-low calculation
    lookback           = min(RECENT_LOW_BARS, len(df))
    recent_window      = df["Low"].iloc[-lookback:]
    recent_low_price   = recent_window.min()
    recent_low_bar_idx = len(df) - lookback + int(recent_window.values.argmin())
    current_close      = df["Close"].iloc[-1]
    pct_below          = (current_close - recent_low_price) / current_close * 100

    fig = plt.figure(figsize=(24, 10), facecolor=s["bg"])
    gs  = gridspec.GridSpec(2, 1, height_ratios=[7, 3],
                            hspace=0.04, top=0.93, bottom=0.08,
                            left=0.05, right=0.96)
    ax1 = fig.add_subplot(gs[0])
    ax2 = fig.add_subplot(gs[1], sharex=ax1)

    for ax in (ax1, ax2):
        ax.set_facecolor(s["panel_bg"])
        ax.tick_params(colors=s["subtext"], labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor(s["border"])
        ax.grid(True, color=s["grid"], linewidth=0.4, alpha=0.6)

    # Candlesticks
    for i, (_, row) in enumerate(df.iterrows()):
        o, h, l, c = row["Open"], row["High"], row["Low"], row["Close"]
        is_bull    = c >= o
        body_col   = s["up_candle"]  if is_bull else s["down_candle"]
        wick_col   = s["wick_up"]    if is_bull else s["wick_down"]

        ax1.bar(i, h - l,      bottom=l,        width=CANDLE_WICK_WIDTH,
                color=wick_col, zorder=2)
        ax1.bar(i, abs(c - o), bottom=min(o, c), width=CANDLE_BODY_WIDTH,
                color=body_col, zorder=3)

    # EMA
    ax1.plot(xs, ema9.values, color=s["ema_color"], linewidth=1.6,
             label=f"EMA {EMA_PERIOD}", zorder=4)

    # Recent-low horizontal dashed line
    low_start_x = len(df) - lookback
    ax1.hlines(
        y=recent_low_price,
        xmin=low_start_x,
        xmax=len(df) - 0.5,
        colors=s["recent_low"],
        linewidths=1.2,
        linestyles="--",
        zorder=5,
    )
    ax1.plot(
        recent_low_bar_idx, recent_low_price,
        marker="D", markersize=5,
        color=s["recent_low"], markeredgecolor=s["bg"],
        markeredgewidth=0.8, zorder=6,
    )
    ax1.annotate(
        f"  {pct_below:.2f}%",
        xy=(len(df) - 1, recent_low_price),
        xytext=(len(df) - 1 + 0.8, recent_low_price),
        color=s["recent_low"],
        fontsize=7.5,
        fontweight="bold",
        va="center",
        ha="left",
        zorder=7,
        annotation_clip=False,
        bbox=dict(
            boxstyle="round,pad=0.3",
            facecolor=s["panel_bg"],
            edgecolor=s["recent_low"],
            alpha=0.85,
            linewidth=0.8,
        ),
    )

    ax1.set_xlim(-1, len(xs) + 1)
    pad = (df["High"].max() - df["Low"].min()) * 0.04
    ax1.set_ylim(df["Low"].min() - pad, df["High"].max() + pad)
    ax1.set_ylabel("Price (Rs)", color=s["text"], fontsize=9)
    ax1.yaxis.set_label_position("right")
    ax1.yaxis.tick_right()

    leg = [mpatches.Patch(facecolor=s["up_candle"],   label="Bullish"),
           mpatches.Patch(facecolor=s["down_candle"], label="Bearish"),
           Line2D([0], [0], color=s["ema_color"], linewidth=1.8,
                  label=f"EMA {EMA_PERIOD}"),
           Line2D([0], [0], color=s["recent_low"], linewidth=1.2,
                  linestyle="--", label=f"{RECENT_LOW_BARS}-bar Low")]
    ax1.legend(handles=leg, loc="upper left", fontsize=8,
               framealpha=0.6, facecolor=s["bg"],
               edgecolor=s["border"], labelcolor=s["text"])

    # Price labels on right y-axis
    last_close = df["Close"].iloc[-1]
    last_ema   = ema9.iloc[-1]
    close_col  = s["up_candle"] if df["Close"].iloc[-1] >= df["Open"].iloc[-1] \
                 else s["down_candle"]

    for val, col, label in [
        (last_close, close_col,      f"Rs{last_close:,.2f}"),
        (last_ema,   s["ema_color"], f"Rs{last_ema:,.2f}"),
    ]:
        ax1.annotate(label,
                     xy=(1, val), xycoords=("axes fraction", "data"),
                     xytext=(4, 0), textcoords="offset points",
                     fontsize=8, fontweight="bold", color=s["bg"],
                     ha="left", va="center",
                     bbox=dict(boxstyle="round,pad=0.3", facecolor=col,
                               edgecolor="none", alpha=0.95),
                     annotation_clip=False)

    # MACD
    hcols = [s["hist_up"] if v >= 0 else s["hist_down"] for v in hist.values]
    ax2.bar(xs, hist.values,    color=hcols, alpha=0.8,  width=0.7, zorder=2,
            label="Histogram")
    ax2.plot(xs, macd_l.values, color=s["macd_line"],   linewidth=1.3,
             label="MACD",   zorder=3)
    ax2.plot(xs, sig.values,    color=s["signal_line"], linewidth=1.1,
             label="Signal", zorder=3)
    ax2.axhline(0, color=s["subtext"], linewidth=0.6, linestyle="--")
    ax2.set_ylabel("MACD", color=s["text"], fontsize=9)
    ax2.yaxis.set_label_position("right")
    ax2.yaxis.tick_right()
    ax2.legend(loc="upper left", fontsize=7.5, framealpha=0.6,
               facecolor=s["bg"], edgecolor=s["border"], labelcolor=s["text"])

    # * X-axis: date+time labels in IST, ~16 ticks for hourly density
    n    = len(xs)
    step = max(n // 16, 1)
    ax2.set_xticks(xs[::step])
    ax2.set_xticklabels(
        [ist_times[i].strftime("%d %b  %H:%M") for i in range(0, n, step)],
        rotation=35, ha="right", fontsize=7, color=s["subtext"])
    plt.setp(ax1.get_xticklabels(), visible=False)

    # * Title block
    lc   = df["Close"].iloc[-1]
    fc   = df["Close"].iloc[0]
    pct  = (lc - fc) / fc * 100
    sign = "+" if pct >= 0 else ""
    ccol = s["up_candle"] if pct >= 0 else s["down_candle"]

    fig.text(0.05, 0.955, f"{symbol}  |  NSE  |  Hourly",
             color=s["text"], fontsize=13, fontweight="bold")
    fig.text(0.05, 0.935, f"Rs{lc:,.2f}   {sign}{pct:.2f}%  ({PERIOD_DAYS}D)",
             color=ccol, fontsize=10)
    latest_ist = ist_times[-1].strftime("%d %b %Y  %H:%M IST")
    fig.text(0.96, 0.955, f"Latest: {latest_ist}",
             color=s["text"], fontsize=9, ha="right", fontweight="bold")
    fig.text(0.96, 0.935,
             f"MACD ({MACD_FAST},{MACD_SLOW},{MACD_SIGNAL})  |  EMA {EMA_PERIOD}"
             f"  |  Bars: {len(df)}  |  Data: yfinance",
             color=s["subtext"], fontsize=8, ha="right")

    plt.savefig(output_path, dpi=150, bbox_inches="tight",
                facecolor=s["bg"], edgecolor="none")
    plt.close(fig)


# ===============================================================
#  STEP 3 - BATCH CHART GENERATOR  (* hourly download)
# ===============================================================

def generate_charts(filtered_df):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    end_date   = datetime.datetime.today() + timedelta(days=1)
    start_date = end_date - timedelta(days=PERIOD_DAYS + 1)   # * 60-day window
    success, failed = [], []
    total = len(filtered_df)

    for idx, row in enumerate(filtered_df.itertuples(), 1):
        sym    = row.Symbol
        ticker = sym if sym.endswith(EXCHANGE_SFX) else sym + EXCHANGE_SFX
        print(f"\n[{idx:>4}/{total}]  {ticker:<22}", end="  ", flush=True)

        try:
            df = yf.download(
                ticker,
                start=start_date.strftime("%Y-%m-%d"),
                end=end_date.strftime("%Y-%m-%d"),
                interval=INTERVAL,          # * "1h"
                auto_adjust=True,
                progress=False,
            )

            if df.empty or len(df) < MACD_SLOW + MACD_SIGNAL + 5:
                print(f"Insufficient data ({len(df)} rows)")
                failed.append(sym)
                continue

            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)

            df = df[["Open", "High", "Low", "Close", "Volume"]].dropna()
            df.index = pd.to_datetime(df.index)

            # * Keep only NSE trading hours (09:15-15:30 IST)
            if df.index.tzinfo is not None:
                df_ist = df.copy()
                df_ist.index = df_ist.index.tz_convert("Asia/Kolkata")
                market_mask = (
                    (df_ist.index.time >= datetime.time(9, 15)) &
                    (df_ist.index.time <= datetime.time(15, 30))
                )
                df = df[market_mask.values]

            df = df.tail(MAX_BARS)         # * cap at MAX_BARS hourly candles

            if len(df) < MACD_SLOW + MACD_SIGNAL + 5:
                print(f"Too few bars after market-hours filter ({len(df)})")
                failed.append(sym)
                continue

            out = os.path.join(OUTPUT_DIR, f"{sym}.png")
            plot_chart(sym, df, out)
            print(f"OK  {len(df)} bars  ->  {out}")
            success.append(sym)

        except Exception:
            print("Exception")
            traceback.print_exc()
            failed.append(sym)

    return success, failed


# ===============================================================
#  MAIN
# ===============================================================

def main():
    parser = argparse.ArgumentParser(
        description="NSE Live -> Filter -> Hourly Charts (EMA9 + MACD + Recent Low)"
    )
    parser.add_argument(
        "--visible", action="store_true",
        help="Run Chrome with a visible window (local debug only)",
    )
    parser.add_argument(
        "--from-csv", metavar="FILE",
        help="Skip browser - parse a manually downloaded NSE CSV",
    )
    args     = parser.parse_args()
    headless = not args.visible

    run_time = datetime.datetime.now().strftime("%d %b %Y  %H:%M:%S")
    print(f"\n{'='*60}")
    print(f"  Chart Generator  -  NSE Live  |  HOURLY  |  {run_time}")
    print(f"  Step 1 : Download NSE Stocks Traded")
    print(f"  Step 2 : Filter  Value > Rs{TRADED_VALUE_MIN_CR} Cr  (EQ series)")
    print(f"  Step 3 : Hourly charts ({PERIOD_DAYS}D window)  ->  {OUTPUT_DIR}/")
    print(f"{'='*60}")

    # STEP 1
    print(f"\n{'-'*60}")
    print("  STEP 1  -  NSE data download")
    print(f"{'-'*60}")
    all_df = download_nse_data(
        headless=headless,
        from_csv=getattr(args, "from_csv", None),
    )
    print(f"\n  Total records : {len(all_df)}")
    all_df.to_csv("nse_all_stocks.csv", index=False)
    print(f"  Saved         : nse_all_stocks.csv")

    # STEP 2
    print(f"\n{'-'*60}")
    print(f"  STEP 2  -  Filter: Value > Rs{TRADED_VALUE_MIN_CR} Cr  (EQ only)")
    print(f"{'-'*60}")
    filtered = filter_stocks(all_df)
    print(f"\n  Stocks passing filter : {len(filtered)}")

    if filtered.empty:
        print("\n  No stocks passed the filter.")
        print("     Inspect nse_all_stocks.csv to check 'Value (Rs Crores)'.")
        sys.exit(0)

    val_col = "Value (Rs Crores)"
    ltp_col = "LTP (Rs)" if "LTP (Rs)" in filtered.columns else "LTP (Rs)"
    print(f"\n  {'#':>4}  {'Symbol':<12} {'Company':<28} "
          f"{'LTP':>8}  {'Value (Cr)':>11}  {'%Chg':>7}")
    print(f"  {'-'*76}")
    for i, row in filtered.head(25).iterrows():
        ltp = row.get(ltp_col, 0)
        val = row.get(val_col, 0)
        pct = row.get("% Change", 0)
        print(f"  {i+1:>4}  {row['Symbol']:<12} "
              f"{str(row.get('Company',''))[:26]:<28} "
              f"  Rs{ltp:>7,.2f}"
              f"  Rs{val:>9,.1f} Cr"
              f"  {pct:>+7.2f}%")
    if len(filtered) > 25:
        print(f"  ... and {len(filtered) - 25} more stocks")

    filtered.to_csv("nse_filtered_stocks.csv", index=False)
    print(f"\n  Saved : nse_filtered_stocks.csv  ({len(filtered)} stocks)")

    # STEP 3
    print(f"\n{'-'*60}")
    print(f"  STEP 3  -  Generating HOURLY charts  ->  {OUTPUT_DIR}/")
    print(f"{'-'*60}")
    success, failed = generate_charts(filtered)

    print(f"\n{'='*60}")
    print(f"  DONE  -  {run_time}")
    print(f"  NSE records  : {len(all_df)}")
    print(f"  Filtered     : {len(filtered)}  (Value > Rs{TRADED_VALUE_MIN_CR} Cr)")
    print(f"  Charts saved : {len(success)}  ->  {OUTPUT_DIR}/")
    if failed:
        print(f"  Failed       : {len(failed)}: "
              + ", ".join(failed[:20])
              + (" ..." if len(failed) > 20 else ""))
    print(f"  nse_all_stocks.csv      - full NSE list")
    print(f"  nse_filtered_stocks.csv - filtered list")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
