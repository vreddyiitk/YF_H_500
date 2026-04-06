"""
chart_generator_hourly.py
=========================
Reads scrip symbols from Excel, downloads 60-day hourly OHLC data
from NSE via yfinance, and saves a TradingView-style (Dark theme)
PNG chart for each stock containing:
  - Candlestick price chart
  - 9-period EMA overlay
  - Standard MACD subplot (12, 26, 9)

NOTE: yfinance limits hourly data to last 60 days only.

Requirements:
    pip install yfinance pandas openpyxl matplotlib

Usage:
    Place your Excel file in the same directory and run:
        python chart_generator_hourly.py
"""

import os
import warnings
import traceback
from datetime import datetime, timedelta

import matplotlib
matplotlib.use("Agg")
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from matplotlib.lines import Line2D
import yfinance as yf

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
EXCEL_FILE   = "Scrips_500.xlsx"
SYMBOL_COL   = None
EXCHANGE_SFX = ".NS"
OUTPUT_DIR   = "YF_H_500"     # separate folder from daily charts
PERIOD_DAYS  = 58             # yfinance allows max 60 days for hourly

EMA_PERIOD  = 9
MACD_FAST   = 12
MACD_SLOW   = 26
MACD_SIGNAL = 9

CANDLE_BODY_WIDTH = 0.75
CANDLE_WICK_WIDTH = 0.12

# ── Dark TradingView palette ─────────────────
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
}


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def read_symbols(path, col):
    df = pd.read_excel(path, sheet_name=0)
    if col and col in df.columns:
        return df[col].dropna().astype(str).str.strip().tolist()
    for c in df.columns:
        sample = df[c].dropna().astype(str).str.strip()
        mask = sample.str.match(r'^[A-Z0-9&\-]{2,20}$')
        if mask.sum() > len(sample) * 0.5:
            print(f"  Auto-detected symbol column: '{c}'")
            return sample[mask].tolist()
    return df.iloc[:, 0].dropna().astype(str).str.strip().tolist()


def ema(series, period):
    return series.ewm(span=period, adjust=False).mean()


def macd(close, fast=12, slow=26, signal=9):
    ml = ema(close, fast) - ema(close, slow)
    sl = ema(ml, signal)
    return ml, sl, ml - sl


def format_xtick(ts, prev_ts):
    """Show date only when day changes, else show time only."""
    if prev_ts is None or ts.date() != prev_ts.date():
        return ts.strftime("%d %b\n%H:%M")
    return ts.strftime("%H:%M")


# ─────────────────────────────────────────────
# CHART
# ─────────────────────────────────────────────

def plot_chart(symbol, df, output_path):
    s       = STYLE
    xs      = np.arange(len(df))
    dlabels = df.index

    ema9              = ema(df["Close"], EMA_PERIOD)
    macd_l, sig, hist = macd(df["Close"], MACD_FAST, MACD_SLOW, MACD_SIGNAL)

    fig = plt.figure(figsize=(24, 10), facecolor=s["bg"])
    gs  = gridspec.GridSpec(2, 1, height_ratios=[7, 3],
                            hspace=0.04, top=0.93, bottom=0.09,
                            left=0.05, right=0.96)
    ax1 = fig.add_subplot(gs[0])
    ax2 = fig.add_subplot(gs[1], sharex=ax1)

    for ax in (ax1, ax2):
        ax.set_facecolor(s["panel_bg"])
        ax.tick_params(colors=s["subtext"], labelsize=7.5)
        for spine in ax.spines.values():
            spine.set_edgecolor(s["border"])
        ax.grid(True, color=s["grid"], linewidth=0.4, alpha=0.6)

    # ── Candlesticks ──────────────────────────
    for i, (_, row) in enumerate(df.iterrows()):
        o, h, l, c = row["Open"], row["High"], row["Low"], row["Close"]
        is_bull  = c >= o
        body_col = s["up_candle"]  if is_bull else s["down_candle"]
        wick_col = s["wick_up"]    if is_bull else s["wick_down"]
        ax1.bar(i, h - l,      bottom=l,        width=CANDLE_WICK_WIDTH,
                color=wick_col, zorder=2)
        ax1.bar(i, abs(c - o), bottom=min(o, c), width=CANDLE_BODY_WIDTH,
                color=body_col, zorder=3)

    # ── EMA ───────────────────────────────────
    ax1.plot(xs, ema9.values, color=s["ema_color"], linewidth=1.6,
             label=f"EMA {EMA_PERIOD}", zorder=4)

    ax1.set_xlim(-1, len(xs) + 2)
    pad = (df["High"].max() - df["Low"].min()) * 0.04
    ax1.set_ylim(df["Low"].min() - pad, df["High"].max() + pad)
    ax1.set_ylabel("Price (₹)", color=s["text"], fontsize=9)
    ax1.yaxis.set_label_position("right")
    ax1.yaxis.tick_right()

    leg = [mpatches.Patch(facecolor=s["up_candle"],   label="Bullish"),
           mpatches.Patch(facecolor=s["down_candle"], label="Bearish"),
           Line2D([0], [0], color=s["ema_color"], linewidth=1.8,
                  label=f"EMA {EMA_PERIOD}")]
    ax1.legend(handles=leg, loc="upper left", fontsize=8,
               framealpha=0.6, facecolor=s["bg"],
               edgecolor=s["border"], labelcolor=s["text"])

    # ── Price + EMA pill labels on right y-axis ──
    last_close = df["Close"].iloc[-1]
    last_ema   = ema9.iloc[-1]
    close_col  = s["up_candle"] if df["Close"].iloc[-1] >= df["Open"].iloc[-1] \
                 else s["down_candle"]

    for val, col, label in [
        (last_close, close_col,      f"₹{last_close:,.2f}"),
        (last_ema,   s["ema_color"], f"₹{last_ema:,.2f}"),
    ]:
        ax1.annotate(label,
                     xy=(1, val), xycoords=("axes fraction", "data"),
                     xytext=(4, 0), textcoords="offset points",
                     fontsize=8, fontweight="bold", color=s["bg"],
                     ha="left", va="center",
                     bbox=dict(boxstyle="round,pad=0.3", facecolor=col,
                               edgecolor="none", alpha=0.95),
                     annotation_clip=False)

    # ── MACD ──────────────────────────────────
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

    # ── X-axis: smart date+time labels ────────
    n    = len(xs)
    step = max(n // 16, 1)   # more ticks for hourly (denser data)
    tick_positions = xs[::step]
    tick_labels = []
    prev = None
    for i in range(0, n, step):
        ts = dlabels[i]
        # Timezone-naive comparison
        ts_naive = ts.replace(tzinfo=None) if hasattr(ts, 'tzinfo') else ts
        prev_naive = prev.replace(tzinfo=None) if (prev is not None and hasattr(prev, 'tzinfo')) else prev
        if prev_naive is None or ts_naive.date() != prev_naive.date():
            tick_labels.append(ts_naive.strftime("%d %b\n%H:%M"))
        else:
            tick_labels.append(ts_naive.strftime("%H:%M"))
        prev = ts

    ax2.set_xticks(tick_positions)
    ax2.set_xticklabels(tick_labels, fontsize=7, color=s["subtext"],
                        ha="center", linespacing=1.3)
    plt.setp(ax1.get_xticklabels(), visible=False)

    # ── Title block ───────────────────────────
    lc   = df["Close"].iloc[-1]
    fc   = df["Close"].iloc[0]
    pct  = (lc - fc) / fc * 100
    sign = "+" if pct >= 0 else ""
    ccol = s["up_candle"] if pct >= 0 else s["down_candle"]

    # Last candle timestamp (IST = UTC+5:30)
    last_ts = df.index[-1]
    try:
        last_ts_ist = last_ts.tz_convert("Asia/Kolkata")
    except Exception:
        last_ts_ist = last_ts
    latest_str = last_ts_ist.strftime("%d %b %Y  %H:%M IST") \
                 if hasattr(last_ts_ist, 'strftime') else str(last_ts_ist)

    fig.text(0.05, 0.955, f"{symbol}  |  NSE  |  1H",
             color=s["text"], fontsize=13, fontweight="bold")
    fig.text(0.05, 0.935, f"₹{lc:,.2f}   {sign}{pct:.2f}%  (60 days)",
             color=ccol, fontsize=10)
    fig.text(0.96, 0.955, f"Latest: {latest_str}",
             color=s["text"], fontsize=9, ha="right", fontweight="bold")
    fig.text(0.96, 0.935,
             f"MACD ({MACD_FAST},{MACD_SLOW},{MACD_SIGNAL})  |  EMA {EMA_PERIOD}"
             f"  |  Bars: {len(df)}  |  Data: yfinance",
             color=s["subtext"], fontsize=8, ha="right")

    plt.savefig(output_path, dpi=150, bbox_inches="tight",
                facecolor=s["bg"], edgecolor="none")
    plt.close(fig)


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"\n{'='*55}")
    print(f"  Chart Generator (Hourly)  -  NSE  |  TradingView Dark")
    print(f"{'='*55}")

    if not os.path.exists(EXCEL_FILE):
        print(f"\n[ERROR] '{EXCEL_FILE}' not found.")
        return

    symbols = read_symbols(EXCEL_FILE, SYMBOL_COL)
    print(f"\n  Loaded {len(symbols)} symbols from '{EXCEL_FILE}'")

    end_date   = datetime.today() + timedelta(days=1)
    start_date = end_date - timedelta(days=PERIOD_DAYS)
    success, failed = [], []

    for i, sym in enumerate(symbols, 1):
        ticker = sym if sym.endswith(EXCHANGE_SFX) else sym + EXCHANGE_SFX
        print(f"\n[{i:>3}/{len(symbols)}]  {ticker:<20}", end="", flush=True)
        try:
            df = yf.download(ticker,
                             start=start_date.strftime("%Y-%m-%d"),
                             end=end_date.strftime("%Y-%m-%d"),
                             interval="60m",        # <── hourly
                             auto_adjust=True,
                             progress=False)

            if df.empty or len(df) < MACD_SLOW + MACD_SIGNAL + 5:
                print(f"  X  Insufficient data ({len(df)} rows)")
                failed.append(sym)
                continue

            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)

            df = df[["Open", "High", "Low", "Close", "Volume"]].dropna()

            # Filter to NSE market hours only: 09:15 – 15:30 IST
            try:
                df.index = df.index.tz_convert("Asia/Kolkata")
                df = df.between_time("09:15", "15:30")
            except Exception:
                pass   # if tz conversion fails, keep all bars

            if len(df) < MACD_SLOW + MACD_SIGNAL + 5:
                print(f"  X  Insufficient bars after market filter ({len(df)})")
                failed.append(sym)
                continue

            out = os.path.join(OUTPUT_DIR, f"{sym}.png")
            plot_chart(sym, df, out)
            print(f"  OK  {len(df)} bars  ->  {out}")
            success.append(sym)

        except Exception:
            print(f"  X  Error")
            traceback.print_exc()
            failed.append(sym)

    print(f"\n{'='*55}")
    print(f"  Done!  {len(success)} charts saved to '{OUTPUT_DIR}/'")
    if failed:
        print(f"  Failed ({len(failed)}): {', '.join(failed[:20])}"
              + (" ..." if len(failed) > 20 else ""))
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()