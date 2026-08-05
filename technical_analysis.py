#!/usr/bin/env python3
"""
technical_analysis.py
=====================

A learning tool for reading the key technical indicators across several companies
at once. Give it a list of companies; it pulls daily price history, computes a
standard indicator suite, and rolls each company up into four plain-English
parameters, a reversal flag, and one composite verdict.

Output, one self-contained HTML file per mode:

    Report_day.html     hourly bars, momentum-weighted
    Report_swing.html   daily bars, balanced
    Report_long.html    weekly bars, trend-weighted

    Double-click any of them to open in a browser; they work offline. Hover a
    summary cell to see the exact conditions that produced it.

Reads tickers from watchlist.xlsx, and optionally holdings from portfolio.xlsx
showing cost basis, market value, gain/loss, weight and days held.

The output parameters:

    Trend                  Uptrend / Downtrend / Sideways
    Trend Strength         How much conviction is behind the move (ADX-based),
                           qualified as strengthening / weakening / stable
    Momentum               Whether the move is speeding up or running out of steam
    Overbought / Oversold  Whether price is stretched, with how long it has been
                           stretched and how unusual the RSI reading is for THIS stock
    Reversal Signals       Corroboration across MA cross, MACD cross, RSI divergence
    Verdict                STRONG BUY / BUY / HOLD / SELL / STRONG SELL - a weighted
                           sum of the four parameters plus a reversal modifier

LOOKBACKS ARE FIXED, not user-supplied. Each one is tied to the period of the
indicator it reads rather than to a calendar number:

    Trend                  latest values + 20-bar slope of SMA-50
    Trend Strength         latest values + 20-bar slope of ADX
    Momentum               histogram sign + 5-bar and 10-bar slopes
    Overbought / Oversold  latest values + 60-bar persistence + 250-bar RSI percentile
    Reversal Signals       10-bar cross window + swing pivots within 120 bars

Roughly 3 years of history is pulled per company so every one of those windows
has enough bars behind it. Only the final readings are reported.

--------------------------------------------------------------------------------
INSTALL (one time):
    pip install yfinance pandas numpy openpyxl

RUN:
    python technical_analysis.py                 # all three modes
    python technical_analysis.py --mode long     # just one

    # or edit DEFAULTS below and just run:  python technical_analysis.py
--------------------------------------------------------------------------------
NOTE: descriptive technical analysis only, not investment advice. The verdict is a
mechanical aggregation of indicators that have NOT been tested against future
outcomes, and it is deliberately blind to what you paid for a position.
"""

import argparse
import os
import sys
import html as _html
from datetime import datetime

import numpy as np
import pandas as pd

try:
    import yfinance as yf
except ImportError:
    sys.exit("Missing dependency 'yfinance'. Run:  pip install yfinance pandas numpy openpyxl")

try:
    import openpyxl  # noqa: F401  - pandas needs it to read .xlsx inputs
except ImportError:
    sys.exit("Missing dependency 'openpyxl' (needed to read the .xlsx inputs). "
             "Run:  pip install openpyxl")



# ============================================================================
# DEFAULTS / CONFIG
# ============================================================================
DEFAULTS = {
    # Fallback tickers, used only when the watchlist workbook is absent or empty.
    "companies": ["NVDA", "AMD", "AVGO", "MRVL", "ASML",
                  "MU", "GOOGL", "INTC", "TSM", "QCOM"],
    "watchlist": "watchlist.xlsx",   # ticker, note
    "portfolio": "portfolio.xlsx",   # ticker, shares, price, date, note
    "output": "Report",              # -> Report_day.html, Report_swing.html, ...
}

# ----------------------------------------------------------------------------
# TRADING MODES
#
# Every period in CFG counts BARS, not calendar time, so changing the interval
# rescales the whole indicator suite at once: RSI-14 becomes 14 hours, 14 days
# or 14 weeks without touching a single period.
#
# The weights are what make the verdict fit the horizon. A day trader and a
# long-term investor should not weight these components equally: long moving
# averages are nearly meaningless intraday and decisive over years, while
# mean reversion pays intraday and is entry-timing at most over years.
#
# Scores are normalised to a fixed -10..+10 scale afterwards, so the verdict
# bands mean the same thing in every mode despite the different weights.
# ----------------------------------------------------------------------------
MODES = {
    "day": {
        "label": "Day trading",
        "interval": "60m",
        "bar": "hour",
        "confirm_bars": 1,          # 1 = report every bar, no confirmation
        "weights": {"Trend": 0.5, "Trend Strength": 1.0, "Momentum": 2.0,
                    "Overbought / Oversold": 2.0, "Reversal": 1.5},
    },
    "swing": {
        "label": "Swing trading",
        "interval": "1d",
        "bar": "day",
        "confirm_bars": 1,
        "weights": {"Trend": 1.0, "Trend Strength": 1.0, "Momentum": 1.0,
                    "Overbought / Oversold": 1.0, "Reversal": 1.0},
    },
    "long": {
        "label": "Long-term investing",
        "interval": "1wk",
        "bar": "week",
        "confirm_bars": 2,          # verdict must hold 2 bars before it changes
        "weights": {"Trend": 2.0, "Trend Strength": 1.5, "Momentum": 0.5,
                    "Overbought / Oversold": 0.25, "Reversal": 0.5},
    },
}

# Largest history yfinance will serve for each interval. Asking for more comes
# back empty, so the request is clamped to these.
INTERVAL_MAX_DAYS = {
    "1m": 7, "2m": 60, "5m": 60, "15m": 60, "30m": 60,
    "60m": 730, "90m": 60, "1h": 730,
}
INTRADAY = set(INTERVAL_MAX_DAYS)

# How long one bar covers, and roughly how many bars a calendar day yields.
# The second figure is what lets history be derived from the interval instead of
# hard-coded per mode, so an --interval override cannot silently under-fetch.
INTERVAL_SPEC = {
    "1m":  (pd.Timedelta(minutes=1),  390 / 1.45),
    "2m":  (pd.Timedelta(minutes=2),  195 / 1.45),
    "5m":  (pd.Timedelta(minutes=5),   78 / 1.45),
    "15m": (pd.Timedelta(minutes=15),  26 / 1.45),
    "30m": (pd.Timedelta(minutes=30),  13 / 1.45),
    "60m": (pd.Timedelta(hours=1),      7 / 1.45),
    "1h":  (pd.Timedelta(hours=1),      7 / 1.45),
    "90m": (pd.Timedelta(minutes=90),   5 / 1.45),
    "1d":  (pd.Timedelta(days=1),       1 / 1.45),
    "1wk": (pd.Timedelta(days=7),       1 / 7),
    "1mo": (pd.Timedelta(days=30),      1 / 30.5),
}

# Bars the slowest component needs: a 250-bar RSI percentile sitting on top of
# RSI's own ~42-bar settling period, with margin.
TARGET_BARS = 450


def history_days_for(interval):
    """Calendar days to request so TARGET_BARS is reachable at this interval."""
    spec = INTERVAL_SPEC.get(interval)
    days = int(TARGET_BARS / spec[1] * 1.15) if spec else 1100
    cap = INTERVAL_MAX_DAYS.get(interval)
    return min(days, cap) if cap else days


def min_bars_required():
    """Below this, the slow parameters cannot be computed as documented."""
    return CFG["sma_slow"] + 20

# Widest value each component can contribute before weighting.
COMPONENT_RANGE = {"Trend": 2.0, "Trend Strength": 2.5, "Momentum": 2.0,
                   "Overbought / Oversold": 2.0, "Reversal": 1.0}

CFG = {
    # ---- indicator periods (textbook standards) ----
    "sma_fast": 20, "sma_mid": 50, "sma_slow": 200,
    "ema_fast": 12, "ema_slow": 26,
    "rsi_period": 14, "rsi_ob": 70, "rsi_os": 30,
    "macd_fast": 12, "macd_slow": 26, "macd_signal": 9,
    "adx_period": 14, "adx_trend": 25, "adx_weak": 20,
    "bb_period": 20, "bb_std": 2.0,
    "stoch_k": 14, "stoch_d": 3, "stoch_ob": 80, "stoch_os": 20,

    # ---- classifier lookbacks (bars) ----
    "trend_slope_bars": 20,      # slope of SMA-50: is the trend tilting up or down
    "adx_slope_bars": 20,        # slope of ADX: is conviction building or draining
    "adx_slope_flat": 0.05,      # |slope| below this counts as "stable"
    "mom_slope_short": 5,        # MACD histogram slope, short horizon
    "mom_slope_long": 10,        # MACD histogram slope, long horizon
    "persist_bars": 60,          # window for counting stretched sessions
    "pct_bars": 250,             # window for ranking RSI against its own history
    "pct_min_bars": 100,         # refuse a percentile on less history than this
    "cross_lookback": 10,        # a MA/MACD cross must be this recent to count
    "pivot_k": 5,                # swing-pivot half-window for RSI divergence
    "div_bars": 120,             # only hunt divergence pivots within this window

    # ---- verdict scoring ----
    # Overbought/oversold is contrarian, but damped where it is known to
    # mislead: inside a strong trend, and on names that stay stretched.
    "persist_damp": 20,          # stretched this many of persist_bars -> halve
    "verdict_bands": {           # cut points on the normalised -10..+10 scale
        "strong_buy": 6.0,
        "buy": 2.5,
        "sell": -2.5,
        "strong_sell": -6.0,
    },

}


# ----------------------------------------------------------------------------
# Category -> colour. Shared so Excel and HTML always agree.
# ----------------------------------------------------------------------------
CATEGORY = {  # (html hex, excel ARGB)
    "bull":    ("#127a3d", "FF127A3D"),
    "bear":    ("#b3261e", "FFB3261E"),
    "neutral": ("#5f6368", "FF5F6368"),
    "over":    ("#b26a00", "FFB26A00"),   # overbought - amber
    "under":   ("#1256a0", "FF1256A0"),   # oversold   - blue
}
FILL = {
    "bull":    ("#e6f4ea", "FFE6F4EA"),
    "bear":    ("#fce8e6", "FFFCE8E6"),
    "neutral": ("#f1f3f4", "FFF1F3F4"),
    "over":    ("#fef7e0", "FFFEF7E0"),
    "under":   ("#e8f0fe", "FFE8F0FE"),
}

# Output parameter names, in report order.
PARAMS = ["Trend", "Trend Strength", "Momentum", "Overbought / Oversold"]
REVERSAL_NAME = "Reversal Signals"
VERDICT_NAME = "Verdict"
VERDICT_CAT = {"STRONG BUY": "bull", "BUY": "bull", "HOLD": "neutral",
               "SELL": "bear", "STRONG SELL": "bear"}


# ============================================================================
# 1. TICKER RESOLUTION
# ============================================================================
def resolve_ticker(name):
    """Return (ticker, display_name). Accepts a raw ticker or a company name.
    Tickers are tested first: faster, and unambiguous."""
    raw = name.strip()
    if not raw:
        return None, None

    try:
        t = yf.Ticker(raw)
        if not t.history(period="5d").empty:
            long_name = raw
            try:
                info = getattr(t, "info", {}) or {}
                long_name = info.get("longName") or info.get("shortName") or raw
            except Exception:
                pass
            return raw.upper(), long_name
    except Exception:
        pass

    try:
        quotes = getattr(yf.Search(raw), "quotes", None) or []
        for q in quotes:
            sym = q.get("symbol")
            if sym and q.get("quoteType", "EQUITY") in ("EQUITY", "ETF", ""):
                return sym.upper(), (q.get("shortname") or q.get("longname") or sym)
        if quotes and quotes[0].get("symbol"):
            return quotes[0]["symbol"].upper(), quotes[0].get("shortname", raw)
    except Exception:
        pass

    return None, raw


# ============================================================================
# 2. DATA FETCH
# ============================================================================
def _now_exchange():
    """Current time in US exchange time, or None if tz data is unavailable."""
    try:
        return pd.Timestamp.now(tz="America/New_York")
    except Exception:
        return None


def drop_incomplete_bar(df, interval):
    """Remove the final bar if its period has not closed yet.

    yfinance happily returns the in-progress bar. Every period and threshold here
    assumes COMPLETED bars, so a partial one is not a less-accurate input, it is
    a wrong one - and it bites hardest in weekly mode, where on four weekdays out
    of five the current bar is a stub. Where completeness cannot be established
    the bar is dropped anyway: losing one bar of recency is far cheaper than
    scoring a fragment.

    Returns (df, dropped_timestamp_or_None).
    """
    if df is None or len(df) < 2:
        return df, None
    spec = INTERVAL_SPEC.get(interval)
    if spec is None:
        return df, None

    last, dur = df.index[-1], spec[0]
    if interval in INTRADAY:
        try:
            now = (pd.Timestamp.now(tz=last.tz) if getattr(last, "tz", None)
                   else pd.Timestamp.now())
            incomplete = (last + dur) > now
        except Exception:
            incomplete = True
    else:
        now = _now_exchange()
        if now is None:
            incomplete = True                       # cannot tell -> assume partial
        else:
            end = (last.tz_localize(now.tz) if getattr(last, "tz", None) is None
                   else last.tz_convert(now.tz))
            if interval == "1wk":
                end = end + pd.Timedelta(days=4)    # Monday-dated bar ends Friday
            elif interval == "1mo":
                end = end + pd.offsets.MonthEnd(0)
            end = end.normalize() + pd.Timedelta(hours=16, minutes=15)
            incomplete = now < end

    return (df.iloc[:-1].copy(), last) if incomplete else (df, None)


def fetch_history(ticker, interval="1d", history_days=None):
    """Completed bars at the requested interval. History is derived from the
    interval, so an --interval override cannot silently under-fetch."""
    if history_days is None:
        history_days = history_days_for(interval)
    cap = INTERVAL_MAX_DAYS.get(interval)
    if cap:
        history_days = min(history_days, cap)
    end = pd.Timestamp.today().normalize() + pd.Timedelta(days=1)
    start = end - pd.Timedelta(days=history_days)
    try:
        df = yf.download(ticker, start=start.strftime("%Y-%m-%d"),
                         interval=interval, auto_adjust=True,
                         progress=False, threads=False)
    except Exception as e:
        print(f"    ! download failed for {ticker}: {e}")
        return None, None
    if df is None or df.empty:
        return None, None
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    df = df.rename(columns=str.title)
    if not {"Open", "High", "Low", "Close"}.issubset(df.columns):
        return None, None
    return drop_incomplete_bar(df.dropna(subset=["Close"]), interval)


# ============================================================================
# 3. INDICATORS  (pure pandas/numpy, no TA-Lib)
# ============================================================================
def _wilder(series, period):
    """Wilder's smoothing (RMA) - what RSI and ADX are actually defined with."""
    return series.ewm(alpha=1 / period, adjust=False, min_periods=period).mean()


def add_indicators(df):
    c, h, l = df["Close"], df["High"], df["Low"]

    df["SMA_fast"] = c.rolling(CFG["sma_fast"]).mean()
    df["SMA_mid"] = c.rolling(CFG["sma_mid"]).mean()
    df["SMA_slow"] = c.rolling(CFG["sma_slow"]).mean()
    df["EMA_fast"] = c.ewm(span=CFG["ema_fast"], adjust=False).mean()
    df["EMA_slow"] = c.ewm(span=CFG["ema_slow"], adjust=False).mean()

    # RSI
    delta = c.diff()
    gain = _wilder(delta.clip(lower=0), CFG["rsi_period"])
    loss = _wilder(-delta.clip(upper=0), CFG["rsi_period"])
    rs = gain / loss.replace(0, np.nan)
    df["RSI"] = (100 - 100 / (1 + rs)).fillna(100)

    # MACD
    ema_f = c.ewm(span=CFG["macd_fast"], adjust=False).mean()
    ema_s = c.ewm(span=CFG["macd_slow"], adjust=False).mean()
    df["MACD"] = ema_f - ema_s
    df["MACD_signal"] = df["MACD"].ewm(span=CFG["macd_signal"], adjust=False).mean()
    df["MACD_hist"] = df["MACD"] - df["MACD_signal"]

    # Bollinger
    mid = c.rolling(CFG["bb_period"]).mean()
    sd = c.rolling(CFG["bb_period"]).std()
    upper, lower = mid + CFG["bb_std"] * sd, mid - CFG["bb_std"] * sd
    df["BB_upper"], df["BB_mid"], df["BB_lower"] = upper, mid, lower
    df["BB_pctB"] = (c - lower) / (upper - lower)

    # ADX / DI
    up, dn = h.diff(), -l.diff()
    # Bind to df.index explicitly: np.where returns a bare array, and if the
    # frame arrived as a slice its cached shape can differ from the array's.
    plus_dm = pd.Series(np.where((up > dn) & (up > 0), up, 0.0), index=df.index)
    minus_dm = pd.Series(np.where((dn > up) & (dn > 0), dn, 0.0), index=df.index)
    prev_c = c.shift(1)
    tr = pd.concat([h - l, (h - prev_c).abs(), (l - prev_c).abs()], axis=1).max(axis=1)
    atr = _wilder(tr, CFG["adx_period"])
    plus_di = 100 * _wilder(plus_dm, CFG["adx_period"]) / atr
    minus_di = 100 * _wilder(minus_dm, CFG["adx_period"]) / atr
    dx = 100 * (plus_di - minus_di).abs() / (plus_di + minus_di).replace(0, np.nan)
    df["ADX"] = _wilder(dx, CFG["adx_period"])
    df["plus_DI"], df["minus_DI"] = plus_di, minus_di

    # Stochastic
    ll, hh = l.rolling(CFG["stoch_k"]).min(), h.rolling(CFG["stoch_k"]).max()
    df["Stoch_K"] = 100 * (c - ll) / (hh - ll)
    df["Stoch_D"] = df["Stoch_K"].rolling(CFG["stoch_d"]).mean()
    return df


# ============================================================================
# 4. SMALL HELPERS
# ============================================================================
def _v(df, col):
    """Latest non-NaN value of a column, or None."""
    s = df[col].dropna()
    return float(s.iloc[-1]) if len(s) else None


def _slope(series, bars):
    """Least-squares slope per bar over the last `bars` points. Uses every
    point, so a single odd bar cannot flip the result the way an endpoint-to-
    endpoint comparison can."""
    s = series.dropna()
    if len(s) < 3:
        return None
    y = s.iloc[-bars:].values
    if len(y) < 3:
        return None
    return float(np.polyfit(np.arange(len(y)), y, 1)[0])


def _percentile(series, bars, min_bars):
    """Where the latest value sits within its own recent history, 0-100.
    Answers 'is this reading unusual FOR THIS STOCK', which a fixed 70/30
    threshold cannot."""
    s = series.dropna()
    if len(s) < min_bars:
        return None
    w = s.iloc[-bars:]
    return float((w <= w.iloc[-1]).sum() / len(w) * 100)


def _count_recent(bool_series, bars):
    """(hits, window_length) over the last `bars` bars."""
    s = bool_series.iloc[-bars:]
    if len(s) == 0:
        return None, 0
    return int(s.sum()), int(len(s))


# ============================================================================
# 5. THE FOUR OUTPUT PARAMETERS
# ============================================================================
def classify_trend(df):
    """Where price sits relative to its own moving averages."""
    price = _v(df, "Close")
    f, m, s = _v(df, "SMA_fast"), _v(df, "SMA_mid"), _v(df, "SMA_slow")
    slope = _slope(df["SMA_mid"], CFG["trend_slope_bars"])
    fastn, midn, slown = CFG["sma_fast"], CFG["sma_mid"], CFG["sma_slow"]

    score = 0
    trace = []

    def check(cond, text_true, text_false):
        nonlocal score
        score += 1 if cond else -1
        trace.append((text_true if cond else text_false, "+1" if cond else "-1"))

    if price is not None and m is not None:
        check(price > m,
              f"Price {_fmt_num(price)} > SMA-{midn} {_fmt_num(m)}",
              f"Price {_fmt_num(price)} < SMA-{midn} {_fmt_num(m)}")
    if f is not None and m is not None:
        check(f > m,
              f"SMA-{fastn} {_fmt_num(f)} > SMA-{midn} {_fmt_num(m)}",
              f"SMA-{fastn} {_fmt_num(f)} < SMA-{midn} {_fmt_num(m)}")
    if slope is not None:
        check(slope > 0,
              f"SMA-{midn} slope {_fmt_num(slope)} > 0 (rising)",
              f"SMA-{midn} slope {_fmt_num(slope)} < 0 (falling)")
    if price is not None and s is not None:
        check(price > s,
              f"Price {_fmt_num(price)} > SMA-{slown} {_fmt_num(s)}",
              f"Price {_fmt_num(price)} < SMA-{slown} {_fmt_num(s)}")

    # The threshold must scale with how many checks actually contributed.
    # With a fixed +/-2 on short history, three checks made a verdict EASIER to
    # reach than four did - worse data producing more confident labels.
    checks = len(trace)
    need = 2 if checks >= 4 else max(2, (checks + 1) // 2 + 1) if checks else 2
    need = min(need, checks) if checks else 2
    if checks and score >= need:
        label, cat = "Uptrend", "bull"
        why = f"Score {score:+d} >= +{need} (of {checks} checks)"
    elif checks and score <= -need:
        label, cat = "Downtrend", "bear"
        why = f"Score {score:+d} <= -{need} (of {checks} checks)"
    else:
        label, cat = "Sideways", "neutral"
        why = (f"Score {score:+d} within +/-{need} (of {checks} checks)"
               if checks else "no usable moving averages")
    trace.append((why, label))

    nums = {
        "Price": price,
        f"SMA-{fastn}": f,
        f"SMA-{midn}": m,
        f"SMA-{slown}": s,
        f"SMA-{midn} slope ({CFG['trend_slope_bars']}-bar)": slope,
        "Trend score (of +/-4)": score,
    }
    return label, cat, nums, trace


def classify_trend_strength(df):
    """How much conviction is behind the move, and whether that conviction is
    building or draining. ADX measures strength without direction; the DI lines
    say which way."""
    adx, p_di, m_di = _v(df, "ADX"), _v(df, "plus_DI"), _v(df, "minus_DI")
    slope = _slope(df["ADX"], CFG["adx_slope_bars"])

    if slope is None:
        qual = "-"
    elif slope > CFG["adx_slope_flat"]:
        qual = "strengthening"
    elif slope < -CFG["adx_slope_flat"]:
        qual = "weakening"
    else:
        qual = "stable"

    if adx is None:
        base, cat = "Insufficient data", "neutral"
        trace = [("ADX unavailable - too few bars", "insufficient data")]
    else:
        trace = []
        if adx < CFG["adx_weak"]:
            base, cat = "Ranging / no clear direction", "neutral"
            trace.append((f"ADX {_fmt_num(adx)} < {CFG['adx_weak']}",
                          "no conviction -> ranging"))
        else:
            strong = adx >= CFG["adx_trend"]
            if strong:
                trace.append((f"ADX {_fmt_num(adx)} >= {CFG['adx_trend']}", "strong"))
            else:
                trace.append((f"ADX {_fmt_num(adx)} between {CFG['adx_weak']} "
                              f"and {CFG['adx_trend']}", "emerging"))
            bullish = (p_di or 0) > (m_di or 0)
            trace.append((f"+DI {_fmt_num(p_di)} {'>' if bullish else '<'} "
                          f"-DI {_fmt_num(m_di)}", "bullish" if bullish else "bearish"))
            if bullish:
                base, cat = ("Strong bullish" if strong else "Bullish (emerging)"), "bull"
            else:
                base, cat = ("Strong bearish" if strong else "Bearish (emerging)"), "bear"

        if slope is None:
            trace.append((f"ADX slope ({CFG['adx_slope_bars']}-bar) unavailable", "-"))
        else:
            flat = CFG["adx_slope_flat"]
            if qual == "stable":
                cmp_txt = f"within +/-{flat}"
            else:
                cmp_txt = f"{'>' if slope > 0 else '<'} {'+' if slope > 0 else '-'}{flat}"
            trace.append((f"ADX slope {_fmt_num(slope)} {cmp_txt}", qual))

    label = base if base == "Insufficient data" or qual == "-" else f"{base}, {qual}"
    nums = {
        f"ADX-{CFG['adx_period']}": adx,
        "+DI": p_di,
        "-DI": m_di,
        f"ADX slope ({CFG['adx_slope_bars']}-bar)": slope,
        "Strength direction": qual,
    }
    return label, cat, nums, trace


def classify_momentum(df):
    """Territory from the MACD histogram's sign; acceleration from least-squares
    slopes over two horizons. Agreement between them = confirmed; disagreement
    flags an early, unconfirmed turn."""
    hist = df["MACD_hist"].dropna()
    macd, sig = _v(df, "MACD"), _v(df, "MACD_signal")
    h_now = float(hist.iloc[-1]) if len(hist) else None
    s_short = _slope(hist, CFG["mom_slope_short"])
    s_long = _slope(hist, CFG["mom_slope_long"])

    if h_now is None or s_short is None or s_long is None:
        label, cat, agree_txt = "Insufficient data", "neutral", "-"
        trace = [("Fewer than 10 usable histogram bars", "insufficient data")]
    elif h_now == 0:
        label, cat, agree_txt = "Flat", "neutral", "-"
        trace = [("Histogram exactly 0", "flat")]
    else:
        agree = (s_short > 0) == (s_long > 0)
        agree_txt = "confirmed" if agree else "diverging"
        rising = s_short > 0
        trace = [
            (f"Histogram {_fmt_num(h_now)} {'>' if h_now > 0 else '<'} 0",
             "bullish territory" if h_now > 0 else "bearish territory"),
            (f"Slope ({CFG['mom_slope_short']}-bar) {_fmt_num(s_short)} "
             f"{'>' if s_short > 0 else '<'} 0", "rising" if s_short > 0 else "falling"),
            (f"Slope ({CFG['mom_slope_long']}-bar) {_fmt_num(s_long)} "
             f"{'>' if s_long > 0 else '<'} 0", "rising" if s_long > 0 else "falling"),
            ("Both slopes same sign" if agree else "Slopes disagree",
             "confirmed" if agree else "early - unconfirmed"),
        ]
        if h_now > 0:
            cat = "bull"
            if agree:
                label = "Bullish & accelerating" if rising else "Bullish but fading"
            else:
                label = ("Bullish, turning up (early)" if rising
                         else "Bullish, turning down (early)")
        else:
            cat = "bear"
            if agree:
                label = "Bearish but recovering" if rising else "Bearish & accelerating"
            else:
                label = ("Bearish, turning up (early)" if rising
                         else "Bearish, turning down (early)")
        trace.append(("Territory + slope verdict", label))

    nums = {
        "MACD": macd, "Signal": sig, "Histogram": h_now,
        f"Slope ({CFG['mom_slope_short']}-bar)": s_short,
        f"Slope ({CFG['mom_slope_long']}-bar)": s_long,
        "Slope agreement": agree_txt,
    }
    return label, cat, nums, trace


def classify_overbought_oversold(df):
    """Whether price is stretched, how long it has been stretched, and how
    unusual the RSI reading is for this particular stock."""
    rsi, pctb, k = _v(df, "RSI"), _v(df, "BB_pctB"), _v(df, "Stoch_K")

    ob_now = sum([rsi is not None and rsi > CFG["rsi_ob"],
                  pctb is not None and pctb > 1,
                  k is not None and k > CFG["stoch_ob"]])
    os_now = sum([rsi is not None and rsi < CFG["rsi_os"],
                  pctb is not None and pctb < 0,
                  k is not None and k < CFG["stoch_os"]])

    # per-bar stretch counts, for persistence over the last N sessions
    ob_bars = ((df["RSI"] > CFG["rsi_ob"]).astype(int)
               + (df["BB_pctB"] > 1).astype(int)
               + (df["Stoch_K"] > CFG["stoch_ob"]).astype(int))
    os_bars = ((df["RSI"] < CFG["rsi_os"]).astype(int)
               + (df["BB_pctB"] < 0).astype(int)
               + (df["Stoch_K"] < CFG["stoch_os"]).astype(int))
    ob_days, win = _count_recent(ob_bars >= 2, CFG["persist_bars"])
    os_days, _ = _count_recent(os_bars >= 2, CFG["persist_bars"])

    rsi_pct = _percentile(df["RSI"], CFG["pct_bars"], CFG["pct_min_bars"])

    def vote_line(nm, val, hi, lo):
        if val is None:
            return (f"{nm} unavailable", "-")
        if val > hi:
            return (f"{nm} {_fmt_num(val)} > {hi}", "overbought")
        if val < lo:
            return (f"{nm} {_fmt_num(val)} < {lo}", "oversold")
        return (f"{nm} {_fmt_num(val)} between {lo} and {hi}", "neutral")

    trace = [
        vote_line(f"RSI-{CFG['rsi_period']}", rsi, CFG["rsi_ob"], CFG["rsi_os"]),
        vote_line("Bollinger %B", pctb, 1, 0),
        vote_line("Stochastic %K", k, CFG["stoch_ob"], CFG["stoch_os"]),
    ]

    if ob_now >= 2:
        base, cat, persist = f"Overbought ({ob_now}/3)", "over", ob_days
        trace.append((f"{ob_now} of 3 overbought (needs 2)", base))
    elif os_now >= 2:
        base, cat, persist = f"Oversold ({os_now}/3)", "under", os_days
        trace.append((f"{os_now} of 3 oversold (needs 2)", base))
    elif ob_now == 1:
        base, cat, persist = "Approaching overbought", "over", None
        trace.append(("1 of 3 overbought (needs 2)", base))
    elif os_now == 1:
        base, cat, persist = "Approaching oversold", "under", None
        trace.append(("1 of 3 oversold (needs 2)", base))
    else:
        base, cat, persist = "Neutral", "neutral", None
        trace.append(("No indicator stretched either way", base))

    if persist:
        trace.append((f"2-of-3 also met on {persist} of last {win} sessions",
                      "persistence"))
    if rsi_pct is not None:
        trace.append((f"RSI ranks at {_fmt_num(rsi_pct)} percentile of its own "
                      f"{CFG['pct_bars']} bars", "context only"))

    label = base if not persist else f"{base}, {persist}/{win} sessions"
    nums = {
        f"RSI-{CFG['rsi_period']}": rsi,
        "Bollinger %B": pctb,
        "Stochastic %K": k,
        f"RSI percentile ({CFG['pct_bars']}-bar)": rsi_pct,
        f"Overbought sessions (last {win})": ob_days,
        f"Oversold sessions (last {win})": os_days,
    }
    return label, cat, nums, trace


# ============================================================================
# 6. REVERSAL SIGNALS
# ============================================================================
def _recent_cross(fast, slow, lookback):
    """+1 fast crossed above slow recently, -1 crossed below, 0 neither."""
    f, s = fast.dropna(), slow.dropna()
    idx = f.index.intersection(s.index)
    if len(idx) < lookback + 2:
        return 0
    sign = np.sign(f.loc[idx] - s.loc[idx])
    recent = sign.iloc[-(lookback + 1):]
    for i in range(1, len(recent)):
        if recent.iloc[i - 1] < 0 and recent.iloc[i] > 0:
            return 1
        if recent.iloc[i - 1] > 0 and recent.iloc[i] < 0:
            return -1
    return 0


def _find_pivots(values, k):
    """Indices that are the max (or min) of their +/-k neighbourhood."""
    highs, lows = [], []
    for i in range(k, len(values) - k):
        win = values[i - k:i + k + 1]
        if values[i] == win.max():
            highs.append(i)
        if values[i] == win.min():
            lows.append(i)
    return highs, lows


def _rsi_divergence(price, rsi, k, max_bars):
    """+1 bullish divergence, -1 bearish, 0 none. Pivot search is bounded to
    the last `max_bars` bars so a stale pivot from long ago cannot drive it."""
    p, r = price.dropna(), rsi.dropna()
    idx = p.index.intersection(r.index)
    if len(idx) < 2 * k + 3:
        return 0
    p = p.loc[idx].iloc[-max_bars:].reset_index(drop=True)
    r = r.loc[idx].iloc[-max_bars:].reset_index(drop=True)
    highs, lows = _find_pivots(p.values, k)
    if len(highs) >= 2:
        a, b = highs[-2], highs[-1]
        if p[b] > p[a] and r[b] < r[a]:       # higher price high, lower RSI high
            return -1
    if len(lows) >= 2:
        a, b = lows[-2], lows[-1]
        if p[b] < p[a] and r[b] > r[a]:       # lower price low, higher RSI low
            return 1
    return 0


def detect_reversal(df):
    lb = CFG["cross_lookback"]
    ma = _recent_cross(df["EMA_fast"], df["EMA_slow"], lb)
    macd = _recent_cross(df["MACD"], df["MACD_signal"], lb)
    div = _rsi_divergence(df["Close"], df["RSI"], CFG["pivot_k"], CFG["div_bars"])

    votes = [ma, macd, div]
    bull = sum(1 for v in votes if v > 0)
    bear = sum(1 for v in votes if v < 0)
    word = {1: "bullish", -1: "bearish", 0: "none"}

    detail = {
        f"MA cross (EMA-{CFG['ema_fast']}/{CFG['ema_slow']})": word[int(ma)],
        "MACD cross": word[int(macd)],
        "RSI divergence": word[int(div)],
    }
    trace = [
        (f"MA cross (EMA-{CFG['ema_fast']}/{CFG['ema_slow']}) "
         f"within {CFG['cross_lookback']} bars", word[int(ma)]),
        (f"MACD cross within {CFG['cross_lookback']} bars", word[int(macd)]),
        (f"RSI divergence within {CFG['div_bars']} bars", word[int(div)]),
    ]
    if bull >= 2 and bull > bear:
        label, cat = f"Possible bullish reversal ({bull}/3)", "bull"
    elif bear >= 2 and bear > bull:
        label, cat = f"Possible bearish reversal ({bear}/3)", "bear"
    elif bull == 1 and bear == 0:
        label, cat = "Watch - weak bullish signal (1/3)", "neutral"
    elif bear == 1 and bull == 0:
        label, cat = "Watch - weak bearish signal (1/3)", "neutral"
    else:
        label, cat = "None", "neutral"
    trace.append((f"{bull} bullish vs {bear} bearish (2 needed to call)", label))
    return label, cat, detail, trace


# ============================================================================
# 7. ASSEMBLE ONE COMPANY
# ============================================================================
def _signals_at(df, back=0):
    """All parameters as they stood `back` bars ago. Slicing an already-computed
    indicator frame is what makes the confirmation filter honest: it re-runs the
    same classifiers on truncated history rather than guessing at prior state."""
    sub = df.iloc[:len(df) - back] if back else df
    if len(sub) < 60:
        return None, None
    signals = {
        "Trend": classify_trend(sub),
        "Trend Strength": classify_trend_strength(sub),
        "Momentum": classify_momentum(sub),
        "Overbought / Oversold": classify_overbought_oversold(sub),
    }
    return signals, detect_reversal(sub)


def analyse(company_input, mode="swing"):
    m = MODES[mode]
    ticker, name = resolve_ticker(company_input)
    if not ticker:
        return {"ok": False, "input": company_input, "error": "Could not resolve to a ticker."}
    df, dropped = fetch_history(ticker, m["interval"])
    need = min_bars_required()
    if df is None or len(df) < need:
        got = 0 if df is None else len(df)
        return {"ok": False, "input": company_input, "ticker": ticker, "name": name,
                "error": f"Needs {need} complete {m['bar']} bars, has {got}. "
                         f"Scoring this would use inputs the parameters do not support."}
    df = add_indicators(df)

    # Note anything computed over a narrower window than designed, rather than
    # letting it pass as a normal reading.
    degraded = []
    rsi_obs = int(df["RSI"].notna().sum())
    if rsi_obs < CFG["pct_bars"]:
        degraded.append(f"RSI percentile ranked over {rsi_obs} readings, "
                        f"not the full {CFG['pct_bars']}")
    if len(df) < CFG["div_bars"]:
        degraded.append(f"divergence searched over {len(df)} bars, "
                        f"not {CFG['div_bars']}")

    signals, reversal = _signals_at(df)
    vlabel, vcat, vnums, vtrace = compute_verdict(signals, reversal, m["weights"])

    # ---- confirmation: a verdict must repeat for N bars before it takes effect
    confirm = m["confirm_bars"]
    pending = None
    if confirm > 1:
        seq = []
        for back in range(confirm + 5, -1, -1):        # oldest first
            sg, rv = _signals_at(df, back)
            if sg:
                seq.append(compute_verdict(sg, rv, m["weights"])[0])
        if seq:
            held, run = seq[0], 1
            for i in range(1, len(seq)):
                run = run + 1 if seq[i] == seq[i - 1] else 1
                if run >= confirm:
                    held = seq[i]
            if held != vlabel:
                pending, vlabel = vlabel, held
                vcat = VERDICT_CAT.get(held, "neutral")
        vtrace.append((f"Confirmation: {confirm} consecutive {m['bar']}s required",
                       f"{pending} pending" if pending else "held"))

    display = f"{vlabel} ({pending} pending)" if pending else vlabel
    return {
        "ok": True, "input": company_input, "ticker": ticker, "name": name,
        "last_close": _v(df, "Close"),
        "last_date": df.index.max(),
        "bars": len(df),
        "mode": mode,
        "dropped_bar": dropped,
        "degraded": degraded,
        "signals": signals,
        "reversal": reversal,
        "verdict": (display, vcat, vnums, vtrace),
    }


# ============================================================================
# 7b. PORTFOLIO
# ============================================================================
def _read_sheet(path, preferred):
    """Open a workbook and return a DataFrame with normalised column names.
    Prefers a sheet named like `preferred`, else falls back to the first sheet,
    so reordering or renaming sheets does not break the read."""
    try:
        book = pd.read_excel(path, sheet_name=None, dtype=object)
    except ImportError:
        return None, [f"reading .xlsx needs openpyxl - run: pip install openpyxl"]
    except FileNotFoundError:
        return None, [f"file not found: {path}"]
    except Exception as e:
        return None, [f"could not read {path}: {e}"]

    if not book:
        return None, [f"{path} has no sheets"]
    name = next((s for s in book if s.strip().lower() == preferred.lower()), None)
    if name is None:
        name = next((s for s in book if s.strip().lower() != "instructions"), None)
    if name is None:
        name = list(book)[0]

    df = book[name]
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df, []


def _cell_str(v):
    """Excel blanks arrive as NaN/None; normalise to a stripped string."""
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    return str(v).strip()


def load_watchlist(path):
    """Read tickers from watchlist.xlsx. Returns (tickers, problems)."""
    df, problems = _read_sheet(path, "Watchlist")
    if df is None:
        return [], problems
    if "ticker" not in df.columns:
        return [], [f"{path}: no 'ticker' column found "
                    f"(saw: {', '.join(df.columns) or 'nothing'})"]

    tickers = []
    for n, val in enumerate(df["ticker"], 2):   # row 2 = first data row
        t = _cell_str(val).upper()
        if not t:
            continue
        if " " in t:
            problems.append(f"row {n}: '{t}' contains a space, skipped")
            continue
        if t not in tickers:
            tickers.append(t)
    return tickers, problems


def load_portfolio(path):
    """Read holdings from portfolio.xlsx. One row per PURCHASE, not per company,
    so multiple lots of the same ticker stay separate and are averaged later.
    Columns: ticker, shares, price, date, note (last two optional).
    Returns (lots, problems)."""
    df, problems = _read_sheet(path, "Portfolio")
    if df is None:
        return [], problems

    missing = [c for c in ("ticker", "shares", "price") if c not in df.columns]
    if missing:
        return [], [f"{path}: missing column(s) {', '.join(missing)} "
                    f"(saw: {', '.join(df.columns) or 'nothing'})"]

    lots = []
    for n, row in enumerate(df.to_dict("records"), 2):
        ticker = _cell_str(row.get("ticker")).upper()
        if not ticker:
            continue
        try:
            shares = float(row.get("shares"))
            price = float(row.get("price"))
        except (TypeError, ValueError):
            problems.append(f"row {n} ({ticker}): shares and price must be numbers")
            continue
        if not np.isfinite(shares) or not np.isfinite(price) or shares <= 0 or price <= 0:
            problems.append(f"row {n} ({ticker}): shares and price must be positive")
            continue

        when = None
        raw_date = row.get("date")
        if _cell_str(raw_date):
            try:
                when = pd.Timestamp(raw_date)
                if pd.isna(when):
                    when = None
            except (ValueError, TypeError):
                problems.append(f"row {n} ({ticker}): unreadable date "
                                f"'{_cell_str(raw_date)}', ignored")
        lots.append({"ticker": ticker, "shares": shares, "price": price,
                     "date": when, "note": _cell_str(row.get("note"))})
    return lots, problems


def build_portfolio(lots, results):
    """Aggregate lots by ticker and value them against the analysed prices."""
    by_ticker = {}
    for lot in lots:
        t = by_ticker.setdefault(lot["ticker"], {"shares": 0.0, "cost": 0.0,
                                                 "first": None, "lots": 0})
        t["shares"] += lot["shares"]
        t["cost"] += lot["shares"] * lot["price"]
        t["lots"] += 1
        if lot["date"] is not None and (t["first"] is None or lot["date"] < t["first"]):
            t["first"] = lot["date"]

    lookup = {r["ticker"]: r for r in results if r.get("ok")}
    rows, total_cost, total_value = [], 0.0, 0.0
    today = pd.Timestamp.today().normalize()

    for ticker, t in by_ticker.items():
        res = lookup.get(ticker)
        last = res["last_close"] if res else None
        avg = t["cost"] / t["shares"] if t["shares"] else None
        value = t["shares"] * last if last is not None else None
        pl = (value - t["cost"]) if value is not None else None
        plpct = (pl / t["cost"] * 100) if pl is not None and t["cost"] else None
        held = int((today - t["first"]).days) if t["first"] is not None else None
        verdict = res["verdict"][:2] if res and res.get("verdict") else None
        rows.append({
            "ticker": ticker, "shares": t["shares"], "lots": t["lots"],
            "avg": avg, "cost": t["cost"], "last": last, "value": value,
            "pl": pl, "plpct": plpct, "held": held, "verdict": verdict,
            "missing": res is None,
        })
        total_cost += t["cost"]
        if value is not None:
            total_value += value

    for r in rows:
        r["weight"] = (r["value"] / total_value * 100) if r["value"] and total_value else None
    rows.sort(key=lambda r: (r["value"] is None, -(r["value"] or 0)))

    totals = {
        "cost": total_cost, "value": total_value,
        "pl": total_value - total_cost,
        "plpct": ((total_value - total_cost) / total_cost * 100) if total_cost else None,
    }
    return rows, totals


# ============================================================================
# 8. FORMATTING HELPERS
# ============================================================================
def _fmt_num(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "-"
    if isinstance(x, (int, np.integer)):
        return f"{x:,}"
    if abs(x) >= 100:
        return f"{x:,.2f}"
    if abs(x) >= 1:
        return f"{x:.2f}"
    return f"{x:.4f}"


def _cell_value(v):
    return _fmt_num(v) if isinstance(v, (int, float, np.integer, np.floating)) or v is None else str(v)


# ============================================================================
# 9. HTML OUTPUT  (self-contained)
# ============================================================================
def _pick(nums, prefix):
    """Fetch a value whose key starts with `prefix` (keys embed CFG values)."""
    for k, v in nums.items():
        if k.startswith(prefix):
            return v
    return None


def compute_verdict(signals, reversal, weights=None):
    """Mechanical aggregation of the four parameters plus a reversal modifier
    into one rating. Reads the parameters' own outputs; it does not recompute
    or alter them. Deliberately blind to any purchase price - the indicators
    are, so anchoring the rating on cost basis would encode the disposition
    effect (selling winners early, holding losers long) rather than analysis.

    Component scores are multiplied by the mode's weights and then normalised to
    a fixed -10..+10 scale, so the bands mean the same thing in every mode."""
    weights = weights or MODES["swing"]["weights"]
    trace = []
    total = 0.0

    def sc(x):
        """Format a contribution, avoiding a displayed '-0.00'."""
        return f"{x + 0.0 if x else 0.0:+.2f}"

    # ---- Trend: +/-2 ----
    t_label = signals["Trend"][0]
    t_score = {"Uptrend": 2.0, "Downtrend": -2.0}.get(t_label, 0.0)
    total += t_score
    trace.append((f"Trend: {t_label}", sc(t_score)))

    # ---- Trend Strength: magnitude from ADX, sign from DI, scaled by qualifier ----
    ts_nums = signals["Trend Strength"][2]
    adx = _pick(ts_nums, "ADX-")
    p_di, m_di = ts_nums.get("+DI"), ts_nums.get("-DI")
    qual = ts_nums.get("Strength direction", "-")
    if adx is None:
        s_score, s_txt = 0.0, "no ADX"
    else:
        mag = 2.0 if adx >= CFG["adx_trend"] else (1.0 if adx >= CFG["adx_weak"] else 0.0)
        sign = 1.0 if (p_di or 0) > (m_di or 0) else -1.0
        factor = {"strengthening": 1.25, "weakening": 0.75}.get(qual, 1.0)
        s_score = sign * mag * factor
        s_txt = f"ADX {_fmt_num(adx)} x{factor} ({qual})"
    total += s_score
    trace.append((f"Trend Strength: {s_txt}", sc(s_score)))

    # ---- Momentum: territory +/-1, short slope +/-0.75, long slope +/-0.25 ----
    m_nums = signals["Momentum"][2]
    hist = m_nums.get("Histogram")
    s5 = m_nums.get(f"Slope ({CFG['mom_slope_short']}-bar)")
    s10 = m_nums.get(f"Slope ({CFG['mom_slope_long']}-bar)")
    if hist is None or s5 is None or s10 is None:
        m_score, m_txt = 0.0, "insufficient data"
    else:
        m_score = ((1.0 if hist > 0 else -1.0)
                   + (0.75 if s5 > 0 else -0.75)
                   + (0.25 if s10 > 0 else -0.25))
        m_txt = (f"hist {'+' if hist > 0 else '-'}, "
                 f"slopes {'+' if s5 > 0 else '-'}/{'+' if s10 > 0 else '-'}")
    total += m_score
    trace.append((f"Momentum: {m_txt}", sc(m_score)))

    # ---- Overbought / Oversold: contrarian, damped inside strong trends ----
    o_nums = signals["Overbought / Oversold"][2]
    rsi = _pick(o_nums, "RSI-")
    pctb, k = o_nums.get("Bollinger %B"), o_nums.get("Stochastic %K")
    rsi_pct = _pick(o_nums, "RSI percentile")
    ob_days = _pick(o_nums, "Overbought sessions")
    os_days = _pick(o_nums, "Oversold sessions")

    ob = sum([rsi is not None and rsi > CFG["rsi_ob"],
              pctb is not None and pctb > 1,
              k is not None and k > CFG["stoch_ob"]])
    osd = sum([rsi is not None and rsi < CFG["rsi_os"],
               pctb is not None and pctb < 0,
               k is not None and k < CFG["stoch_os"]])

    if ob >= 2:
        o_score, o_txt = (-1.5 if ob == 3 else -1.0), f"overbought {ob}/3"
    elif osd >= 2:
        o_score, o_txt = (1.5 if osd == 3 else 1.0), f"oversold {osd}/3"
    elif ob == 1:
        o_score, o_txt = -0.5, "approaching overbought"
    elif osd == 1:
        o_score, o_txt = 0.5, "approaching oversold"
    else:
        o_score, o_txt = 0.0, "neutral"

    dampers = []
    if o_score != 0 and adx is not None and adx >= CFG["adx_trend"]:
        o_score *= 0.5
        dampers.append("strong trend")
    persist = ob_days if o_score < 0 else os_days
    if o_score != 0 and persist and persist >= CFG["persist_damp"]:
        o_score *= 0.5
        dampers.append(f"chronic {persist}d")
    if rsi_pct is not None:
        if rsi_pct >= 95:
            o_score -= 0.5
            dampers.append("pctile>=95")
        elif rsi_pct <= 5:
            o_score += 0.5
            dampers.append("pctile<=5")
    if dampers:
        o_txt += " [" + ", ".join(dampers) + "]"
    total += o_score
    trace.append((f"Overbought/Oversold: {o_txt}", sc(o_score)))

    # ---- Reversal Signals: +/-1 modifier ----
    r_detail = reversal[2]
    words = list(r_detail.values())
    r_bull = words.count("bullish")
    r_bear = words.count("bearish")
    if r_bull >= 2 and r_bull > r_bear:
        r_score, r_txt = 1.0, f"bullish {r_bull}/3"
    elif r_bear >= 2 and r_bear > r_bull:
        r_score, r_txt = -1.0, f"bearish {r_bear}/3"
    elif r_bull == 1 and r_bear == 0:
        r_score, r_txt = 0.5, "weak bullish 1/3"
    elif r_bear == 1 and r_bull == 0:
        r_score, r_txt = -0.5, "weak bearish 1/3"
    else:
        r_score, r_txt = 0.0, "none"
    total += r_score
    trace.append((f"Reversal: {r_txt}", sc(r_score)))

    # ---- weight, normalise to -10..+10, then band ----
    raw = (t_score * weights["Trend"]
           + s_score * weights["Trend Strength"]
           + m_score * weights["Momentum"]
           + o_score * weights["Overbought / Oversold"]
           + r_score * weights["Reversal"])
    ceiling = sum(COMPONENT_RANGE[k] * weights[k] for k in COMPONENT_RANGE)
    total = (raw / ceiling * 10.0) if ceiling else 0.0

    b = CFG["verdict_bands"]
    if total >= b["strong_buy"]:
        label, cat = "STRONG BUY", "bull"
    elif total >= b["buy"]:
        label, cat = "BUY", "bull"
    elif total > b["sell"]:
        label, cat = "HOLD", "neutral"
    elif total > b["strong_sell"]:
        label, cat = "SELL", "bear"
    else:
        label, cat = "STRONG SELL", "bear"

    trace.append((f"Weighted {raw:+.2f} of {ceiling:.2f} max "
                  f"-> normalised {total:+.2f} on -10..+10", label))
    nums = {
        "Score (-10 to +10)": round(total, 2),
        "Trend": round(t_score, 2),
        "Trend Strength": round(s_score, 2),
        "Momentum": round(m_score, 2),
        "Overbought / Oversold": round(o_score, 2),
        "Reversal": round(r_score, 2),
    }
    return label, cat, nums, trace


def _pill(label, cat):
    return (f'<span class="pill" style="color:{CATEGORY[cat][0]};'
            f'background:{FILL[cat][0]}">{_html.escape(str(label))}</span>')


def _pill_tip(label, cat, trace, heading):
    """Summary-table pill that reveals, on hover, the conditions that were
    evaluated to produce this label - each with the values substituted in and
    the outcome it contributed."""
    rows = "".join(
        f'<div><span>{_html.escape(str(cond))}</span>'
        f'<span>{_html.escape(str(outcome))}</span></div>'
        for cond, outcome in trace)
    return (f'<span class="tw">{_pill(label, cat)}'
            f'<span class="tip"><span class="tiph">{_html.escape(heading)}'
            f' &mdash; how this was decided</span>{rows}</span></span>')


def write_html(results, path, prows=None, ptotals=None, pproblems=None,
               mode_label="", interval=""):
    css = """
    :root{--ink:#1a1c1e;--muted:#6b7075;--line:#e6e8eb;--bg:#fbfbfc;}
    *{box-sizing:border-box}
    body{margin:0;background:var(--bg);color:var(--ink);
      font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
      -webkit-font-smoothing:antialiased;line-height:1.45}
    .wrap{max-width:1120px;margin:0 auto;padding:40px 28px 80px}
    header h1{font-size:26px;margin:0 0 4px;letter-spacing:-.4px}
    header .sub{color:var(--muted);font-size:13px}
    h2{font-size:13px;text-transform:uppercase;letter-spacing:1.2px;color:var(--muted);
      font-weight:600;margin:48px 0 14px;padding-bottom:8px;border-bottom:1px solid var(--line)}
    table{width:100%;border-collapse:collapse;font-size:14px}
    th,td{text-align:left;padding:11px 12px;border-bottom:1px solid var(--line);vertical-align:middle}
    th{font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);font-weight:600}
    tbody tr:hover{background:#f4f6f8}
    .co{font-weight:600}
    .tk{color:var(--muted);font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:12px}
    .pill{display:inline-block;padding:3px 10px;border-radius:999px;font-size:12.5px;
      font-weight:600;line-height:1.35}
    .card{background:#fff;border:1px solid var(--line);border-radius:14px;padding:22px 24px;
      margin:16px 0;box-shadow:0 1px 2px rgba(0,0,0,.03)}
    .card h3{margin:0 0 2px;font-size:18px}
    .card .meta{color:var(--muted);font-size:12px;font-family:ui-monospace,Menlo,monospace;
      margin-bottom:16px}
    .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:14px}
    .sig{border:1px solid var(--line);border-radius:10px;padding:13px 15px}
    .sig .name{font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);
      margin-bottom:7px}
    .sig .nums{margin-top:10px;font-size:12.5px;color:#3c4043}
    .sig .nums div{display:flex;justify-content:space-between;gap:10px;padding:3px 0;
      border-top:1px dashed #eef0f2}
    .sig .nums span:last-child{font-family:ui-monospace,Menlo,monospace;white-space:nowrap}
    .rev{grid-column:1/-1;background:#fafbfc}
    .err{color:#b3261e;font-style:italic}
    .tw{position:relative;display:inline-block}
    .tw .tip{position:absolute;left:0;top:calc(100% + 6px);z-index:30;width:340px;
      background:#fff;border:1px solid #d4d7db;border-radius:8px;padding:9px 11px;
      box-shadow:0 6px 20px rgba(0,0,0,.13);opacity:0;visibility:hidden;
      transition:opacity .09s;pointer-events:none;text-align:left;white-space:normal}
    .tw:hover .tip,.tw.active .tip{opacity:1;visibility:visible}
    .tw .tiph{display:block;font-size:10.5px;text-transform:uppercase;letter-spacing:.6px;
      color:var(--muted);font-weight:600;margin-bottom:5px}
    .tw .tip div{display:flex;justify-content:space-between;gap:12px;padding:3px 0;
      font-size:12px;line-height:1.4;color:#3c4043;border-top:1px dashed #eef0f2}
    .tw .tip div span:first-child{flex:1;font-family:ui-monospace,Menlo,monospace}
    .tw .tip div span:last-child{flex-shrink:0;font-weight:600;color:#202124}
    .tw .tip div:last-child{border-top:1px solid #d4d7db;margin-top:3px;padding-top:5px}
    td:last-child .tw .tip, td:nth-last-child(2) .tw .tip{left:auto;right:0}
    tbody tr:last-child .tw .tip{top:auto;bottom:calc(100% + 6px)}
    footer{margin-top:60px;color:var(--muted);font-size:11.5px;text-align:center;line-height:1.8}
    table.pf{font-size:13px}
    table.pf td.n,table.pf th.n{text-align:right;font-family:ui-monospace,Menlo,monospace}
    table.pf th.n{font-family:inherit}
    .pos{color:#127a3d;font-weight:600}
    .neg{color:#b3261e;font-weight:600}
    .mut{color:var(--muted)}
    tr.tot td{border-top:2px solid #c8ccd0;font-weight:700;background:#f7f8f9}
    .pwarn{background:#fef7e0;border:1px solid #f0d79a;border-radius:8px;padding:8px 11px;
      font-size:12.5px;color:#7a5300;margin-bottom:12px}
    .pnote{margin-top:9px;font-size:11.5px;color:var(--muted);line-height:1.6}
    @media(max-width:620px){.wrap{padding:24px 14px 60px}table{font-size:12.5px}
      th,td{padding:8px 6px}}
    """

    def nums_html(nums):
        rows = "".join(
            f'<div><span>{_html.escape(k)}</span><span>{_cell_value(v)}</span></div>'
            for k, v in nums.items())
        return f'<div class="nums">{rows}</div>'

    srows = ""
    for res in results:
        if not res.get("ok"):
            srows += (f'<tr><td class="co">{_html.escape(res.get("name") or res["input"])}</td>'
                      f'<td class="tk">{_html.escape(res.get("ticker","-"))}</td>'
                      f'<td class="err" colspan="6">{_html.escape(res.get("error","error"))}</td></tr>')
            continue
        cells = ""
        for k in PARAMS:
            label, cat, _, trace = res["signals"][k]
            cells += f'<td>{_pill_tip(label, cat, trace, k)}</td>'
        rlabel, rcat, _, rtrace = res["reversal"]
        vlabel, vcat, _, vtrace = res["verdict"]
        srows += (f'<tr><td class="co">{_html.escape(res["name"])}</td>'
                  f'<td class="tk">{_html.escape(res["ticker"])}</td>{cells}'
                  f'<td>{_pill_tip(rlabel, rcat, rtrace, REVERSAL_NAME)}</td>'
                  f'<td>{_pill_tip(vlabel, vcat, vtrace, VERDICT_NAME)}</td></tr>')

    cards = ""
    for res in results:
        if not res.get("ok"):
            continue
        blocks = ""
        for key in PARAMS:
            label, cat, nums, _ = res["signals"][key]
            blocks += (f'<div class="sig"><div class="name">{_html.escape(key)}</div>'
                       f'{_pill(label, cat)}{nums_html(nums)}</div>')
        rlabel, rcat, rdetail, _ = res["reversal"]
        blocks += (f'<div class="sig"><div class="name">{REVERSAL_NAME}</div>'
                   f'{_pill(rlabel, rcat)}{nums_html(rdetail)}</div>')
        vlabel, vcat, vnums, _ = res["verdict"]
        blocks += (f'<div class="sig rev"><div class="name">{VERDICT_NAME}</div>'
                   f'{_pill(vlabel, vcat)}{nums_html(vnums)}</div>')
        degnote = (f'<div class="pwarn">{_html.escape("; ".join(res["degraded"]))}</div>'
                   if res.get("degraded") else "")
        cards += (f'<div class="card"><h3>{_html.escape(res["name"])} '
                  f'<span class="tk">{_html.escape(res["ticker"])}</span></h3>'
                  f'<div class="meta">last close {_fmt_num(res["last_close"])} &nbsp;·&nbsp; '
                  f'last complete bar {res["last_date"]:%Y-%m-%d} &nbsp;·&nbsp; '
                  f'{res["bars"]} bars analysed</div>{degnote}'
                  f'<div class="grid">{blocks}</div></div>')

    # ---- portfolio block (only when holdings were supplied) ----
    pblock = ""
    if prows:
        def money(x):
            return "-" if x is None else f"{x:,.2f}"

        def signed(x, pct=False):
            if x is None:
                return '<span class="mut">-</span>'
            cls = "pos" if x >= 0 else "neg"
            txt = f"{x:+,.2f}%" if pct else f"{x:+,.2f}"
            return f'<span class="{cls}">{txt}</span>'

        body = ""
        for r in prows:
            v = (f'{_pill(*r["verdict"])}' if r["verdict"]
                 else '<span class="mut">-</span>')
            shares = f'{r["shares"]:,.0f}' if float(r["shares"]).is_integer() else f'{r["shares"]:,.2f}'
            lots = f' <span class="mut">({r["lots"]} lots)</span>' if r["lots"] > 1 else ""
            weight = "-" if r["weight"] is None else f'{r["weight"]:.1f}%'
            held = "-" if r["held"] is None else str(r["held"])
            body += (f'<tr><td class="tk">{_html.escape(r["ticker"])}</td>'
                     f'<td class="n">{shares}{lots}</td>'
                     f'<td class="n">{money(r["avg"])}</td>'
                     f'<td class="n">{money(r["cost"])}</td>'
                     f'<td class="n">{money(r["last"])}</td>'
                     f'<td class="n">{money(r["value"])}</td>'
                     f'<td class="n">{signed(r["pl"])}</td>'
                     f'<td class="n">{signed(r["plpct"], True)}</td>'
                     f'<td class="n">{weight}</td>'
                     f'<td class="n">{held}</td>'
                     f'<td>{v}</td></tr>')
        body += (f'<tr class="tot"><td>TOTAL</td><td></td><td></td>'
                 f'<td class="n">{money(ptotals["cost"])}</td><td></td>'
                 f'<td class="n">{money(ptotals["value"])}</td>'
                 f'<td class="n">{signed(ptotals["pl"])}</td>'
                 f'<td class="n">{signed(ptotals["plpct"], True)}</td>'
                 f'<td></td><td></td><td></td></tr>')
        warn = ""
        if pproblems:
            warn = ('<div class="pwarn">Skipped lines in the holdings file: '
                    + _html.escape("; ".join(pproblems)) + "</div>")
        pblock = (f'<h2>Portfolio</h2>{warn}'
                  f'<table class="pf"><thead><tr><th>Ticker</th><th class="n">Shares</th>'
                  f'<th class="n">Avg cost</th><th class="n">Cost basis</th>'
                  f'<th class="n">Last price</th><th class="n">Market value</th>'
                  f'<th class="n">Gain/loss</th><th class="n">Gain/loss %</th>'
                  f'<th class="n">Weight</th><th class="n">Days held</th>'
                  f'<th>{_html.escape(VERDICT_NAME)}</th></tr></thead>'
                  f'<tbody>{body}</tbody></table>'
                  f'<div class="pnote">Values use the last close shown above, not live prices. '
                  f'The verdict is the same mechanical rating as the summary and is blind to '
                  f'what you paid &mdash; it reflects the indicators only, and has not been '
                  f'tested against outcomes.</div>')

    th = "".join(f"<th>{_html.escape(p)}</th>" for p in PARAMS)
    ok_n = len([r for r in results if r.get("ok")])
    doc = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Technical Analysis Report</title><style>{css}</style></head><body><div class="wrap">
<header><h1>Technical Analysis Report</h1>
<div class="sub">{ok_n} companies · {_html.escape(mode_label)} · {_html.escape(interval)} bars · generated {datetime.now():%Y-%m-%d %H:%M}</div></header>
<h2>Cross-Company Summary</h2>
<table><thead><tr><th>Company</th><th>Ticker</th>{th}
<th>{_html.escape(REVERSAL_NAME)}</th><th>{_html.escape(VERDICT_NAME)}</th></tr></thead>
<tbody>{srows}</tbody></table>
{pblock}
<h2>Company Detail</h2>
{cards}
<footer>
Indicators: RSI {CFG['rsi_period']} · MACD {CFG['macd_fast']}/{CFG['macd_slow']}/{CFG['macd_signal']} ·
ADX {CFG['adx_period']} · Bollinger {CFG['bb_period']}/{CFG['bb_std']} · Stochastic {CFG['stoch_k']}/{CFG['stoch_d']} ·
SMA {CFG['sma_fast']}/{CFG['sma_mid']}/{CFG['sma_slow']}<br>
Lookbacks: trend &amp; strength slopes {CFG['trend_slope_bars']} bars ·
momentum slopes {CFG['mom_slope_short']} and {CFG['mom_slope_long']} bars ·
persistence {CFG['persist_bars']} bars · RSI percentile {CFG['pct_bars']} bars ·
crosses {CFG['cross_lookback']} bars · divergence within {CFG['div_bars']} bars<br>
Descriptive technical analysis for learning purposes — not investment advice.
These labels report what the indicators currently say and have not been tested against
future outcomes. Data via yfinance.
</footer></div>
<script>
document.addEventListener('click', function(e){{
  var tw = e.target.closest('.tw');
  document.querySelectorAll('.tw.active').forEach(function(el){{
    if (el !== tw) el.classList.remove('active');
  }});
  if (tw) {{
    e.stopPropagation();
    tw.classList.toggle('active');
  }}
}});
</script>
</body></html>"""

    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)


# ============================================================================
# 10. MAIN
# ============================================================================
def run_mode(mode, companies, lots, pproblems, base, interval_override=None):
    """Analyse every company in one mode and write that mode's report.
    Each mode fetches its own bars, because the interval differs."""
    m = MODES[mode]
    if interval_override:
        m = dict(m)
        m["interval"] = interval_override
        m["bar"] = {"1d": "day", "1wk": "week",
                    "1mo": "month"}.get(interval_override, "bar")
        MODES[mode] = m
    interval = m["interval"]

    print(f"\n=== {m['label']}  ({interval} bars, confirmation "
          f"{m['confirm_bars']} bar{'s' if m['confirm_bars'] != 1 else ''}) ===")
    if interval in INTRADAY:
        print("    ! intraday data is free and delayed ~15 min, and this tool "
              "ignores spreads,\n      liquidity and fees, which dominate at "
              "that timescale. Day mode is for\n      watching indicator "
              "behaviour, not for trading on.")

    results = []
    for comp in companies:
        print(f"  - {comp} ...", end=" ", flush=True)
        try:
            res = analyse(comp, mode)
        except Exception as e:
            res = {"ok": False, "input": comp, "error": f"unexpected: {e}"}
        if res.get("ok"):
            print(f"{res['ticker']}: {res['verdict'][0]}  "
                  f"({res['signals']['Trend'][0]}, {res['bars']} bars)")
            if res.get("dropped_bar") is not None:
                print(f"        dropped in-progress bar "
                      f"{res['dropped_bar']:%Y-%m-%d %H:%M} - not yet closed")
            for d in res.get("degraded", []):
                print(f"        degraded: {d}")
        else:
            print(f"SKIPPED ({res.get('error')})")
        results.append(res)

    prows, ptotals = (build_portfolio(lots, results) if lots else (None, None))
    path = f"{base}_{mode}.html"
    write_html(results, path, prows, ptotals, pproblems, m["label"], interval)

    ok = sum(1 for r in results if r.get("ok"))
    print(f"  -> {path}   ({ok}/{len(results)} analysed)")
    if prows:
        print(f"     portfolio: {len(prows)} holdings, "
              f"value {ptotals['value']:,.2f}, P/L {ptotals['pl']:+,.2f}")
    return path, ok, len(results)


def main():
    ap = argparse.ArgumentParser(
        description="Technical analysis reports driven by watchlist.xlsx and "
                    "portfolio.xlsx. With no --mode, all three modes are run.")
    ap.add_argument("--watchlist", type=str,
                    help="Watchlist workbook (default: watchlist.xlsx).")
    ap.add_argument("--portfolio", type=str,
                    help="Holdings workbook (default: portfolio.xlsx).")
    ap.add_argument("--mode", type=str, choices=list(MODES), default=None,
                    help="Run a single mode: day | swing | long. "
                         "Omit to run all three.")
    ap.add_argument("--interval", type=str, default=None,
                    help="Override the bar interval, e.g. 15m, 1d, 1wk. "
                         "Applies to the single mode given by --mode.")
    ap.add_argument("--output", type=str,
                    help="Output prefix; the mode is appended (default: Report).")
    args = ap.parse_args()

    if args.interval and not args.mode:
        ap.error("--interval needs --mode, since it applies to one mode's bars.")

    modes = [args.mode] if args.mode else list(MODES)
    base = args.output or DEFAULTS["output"]
    if base.lower().endswith((".xlsx", ".html")):
        base = base.rsplit(".", 1)[0]

    # ---- watchlist: named file, else the default name, else built-in list ----
    wl_path = args.watchlist or DEFAULTS["watchlist"]
    companies = []
    if os.path.exists(wl_path):
        companies, wproblems = load_watchlist(wl_path)
        for msg in wproblems:
            print(f"    ! {msg}")
        if companies:
            print(f"Watchlist: {len(companies)} tickers from {wl_path}")
        else:
            print(f"Watchlist: {wl_path} held no usable tickers - "
                  f"using the built-in list instead")
    else:
        print(f"Watchlist: no {wl_path} found - using the built-in list "
              f"(copy watchlist_example.xlsx to {wl_path} to change it)")
    if not companies:
        companies = list(DEFAULTS["companies"])

    # ---- holdings: optional ----
    pf_path = args.portfolio or DEFAULTS["portfolio"]
    lots, pproblems = ([], [])
    if os.path.exists(pf_path):
        lots, pproblems = load_portfolio(pf_path)
        print(f"Holdings: {len(lots)} lots from {pf_path}"
              + (f" ({len(pproblems)} rows skipped)" if pproblems else ""))
        for msg in pproblems:
            print(f"    ! {msg}")
        # anything held is analysed too, whether or not it is on the watchlist
        have = {c.strip().upper() for c in companies}
        for t in dict.fromkeys(lot["ticker"] for lot in lots):
            if t not in have:
                companies.append(t)
    elif args.portfolio:
        print(f"Holdings: {pf_path} not found - no portfolio table")

    print(f"Analysing {len(companies)} companies across "
          f"{len(modes)} mode{'s' if len(modes) != 1 else ''}: {', '.join(modes)}")
    if len(modes) > 1:
        print("Each mode fetches its own bars, so this makes "
              f"{len(modes)} passes over the list.")

    written = []
    for mode in modes:
        try:
            path, ok, total = run_mode(mode, companies, lots, pproblems,
                                       base, args.interval)
            written.append(path)
        except Exception as e:
            # one mode failing must not lose the others
            print(f"  ! {mode} mode failed: {e}")

    print(f"\nDone. {len(written)} report"
          f"{'s' if len(written) != 1 else ''} written:")
    for path in written:
        print(f"  {path}")


if __name__ == "__main__":
    main()
