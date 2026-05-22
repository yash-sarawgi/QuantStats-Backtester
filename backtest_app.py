#!/usr/bin/env python3
"""
QuantStats Strategy Backtesting Report Generator
Professional GUI for strategy backtesting with CSV or Alpaca data
"""

# ─── AUTO-INSTALL DEPENDENCIES ──────────────────────────────────────────────
import subprocess, sys

REQUIRED = {
    "quantstats": "quantstats",
    "pandas": "pandas",
    "numpy": "numpy",
    "alpaca": "alpaca-py",
    "pytz": "pytz",
}

def install_if_missing():
    for mod, pkg in REQUIRED.items():
        try:
            __import__(mod)
        except ImportError:
            print(f"Installing {pkg}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

install_if_missing()

# ─── IMPORTS ────────────────────────────────────────────────────────────────
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading, traceback, os, tempfile, webbrowser, re
import pandas as pd
import numpy as np
import pytz
import quantstats as qs
from datetime import datetime, timedelta

# ─── THEME CONSTANTS ────────────────────────────────────────────────────────
BG        = "#0d1117"
BG2       = "#161b22"
BG3       = "#21262d"
BORDER    = "#30363d"
ACCENT    = "#58a6ff"
ACCENT2   = "#3fb950"
WARNING   = "#d29922"
DANGER    = "#f85149"
FG        = "#e6edf3"
FG2       = "#8b949e"
FONT_MONO = ("Cascadia Code", 10) if sys.platform == "win32" else ("Menlo", 10)
FONT_UI   = ("Segoe UI", 10)      if sys.platform == "win32" else ("SF Pro Display", 10)
FONT_HEAD = ("Segoe UI Semibold", 11) if sys.platform == "win32" else ("SF Pro Display", 11)

# ─── DEFAULT STRATEGY ───────────────────────────────────────────────────────
DEFAULT_STRATEGY = '''\
# ══════════════════════════════════════════════════════════════════════════
# STRATEGY CODE — Required Function Signature:
#
#   def strategy(data: pd.DataFrame) -> pd.Series
#
# INPUT  `data` columns (all lowercase, guaranteed):
#   timestamp  – index (DatetimeTZAware, UTC)
#   open, high, low, close  – OHLC prices (float)
#   volume                  – volume (float)
#   vwap                    – volume-weighted avg price (float, if available)
#
# OUTPUT  pd.Series of SIGNALS aligned to `data.index`:
#   +1  = Long
#    0  = Flat / Out
#   -1  = Short
#
# The engine converts signals → next-bar-open fill → daily returns → tearsheet.
#
# You can import anything installed: numpy, pandas, talib (if installed), etc.
# ══════════════════════════════════════════════════════════════════════════

import pandas as pd
import numpy as np

def strategy(data: pd.DataFrame) -> pd.Series:
    """
    Example: Simple Golden Cross (SMA 20 / SMA 50)
    Go long when fast MA crosses above slow MA, flat otherwise.
    """
    close = data["close"]

    fast = close.rolling(20).mean()
    slow = close.rolling(50).mean()

    signal = pd.Series(0, index=data.index, dtype=float)
    signal[fast > slow] = 1      # long
    signal[fast <= slow] = 0     # flat

    return signal
'''

# ─── ALPACA TIMEFRAME MAP ────────────────────────────────────────────────────
TF_MAP = {
    "1 Min":  ("1Min",  1),
    "5 Min":  ("5Min",  5),
    "15 Min": ("15Min", 15),
    "30 Min": ("30Min", 30),
    "1 Hour": ("1Hour", 60),
    "4 Hour": ("4Hour", 240),
    "1 Day":  ("1Day",  1440),
}

# ════════════════════════════════════════════════════════════════════════════
# DATA LAYER
# ════════════════════════════════════════════════════════════════════════════

def detect_and_load_csv(filepath: str) -> pd.DataFrame:
    """
    Robust CSV loader. Auto-detects timestamp, OHLCV columns regardless of
    casing, ordering, or extra columns.  Handles Alpaca-style exports.
    """
    df = pd.read_csv(filepath)
    df.columns = [c.strip().lower() for c in df.columns]

    # ── Detect timestamp column ──────────────────────────────────────────
    ts_candidates = ["timestamp", "datetime", "date", "time", "t",
                     "date_time", "bar_time", "open_time"]
    ts_col = None
    for c in ts_candidates:
        if c in df.columns:
            ts_col = c
            break
    if ts_col is None:
        # Fallback: first column that looks like a date string
        for c in df.columns:
            sample = str(df[c].iloc[0])
            if re.search(r"\d{4}-\d{2}-\d{2}", sample):
                ts_col = c
                break
    if ts_col is None:
        raise ValueError("Cannot detect timestamp column in CSV. "
                         "Expected a column named 'timestamp', 'datetime', or 'date'.")

    # ── Parse timestamp ──────────────────────────────────────────────────
    df["timestamp"] = pd.to_datetime(df[ts_col], utc=True, errors="coerce")
    df = df.dropna(subset=["timestamp"])
    df = df.set_index("timestamp").sort_index()

    # ── Detect OHLCV columns ─────────────────────────────────────────────
    col_aliases = {
        "open":   ["open", "o", "open_price", "open_bid"],
        "high":   ["high", "h", "high_price", "ask_high"],
        "low":    ["low",  "l", "low_price",  "bid_low"],
        "close":  ["close","c", "close_price","last","last_price","adj_close"],
        "volume": ["volume","vol","v","qty","quantity"],
        "vwap":   ["vwap","vw","weighted_avg"],
    }
    renamed = {}
    for canonical, aliases in col_aliases.items():
        if canonical in df.columns:
            renamed[canonical] = canonical
            continue
        for alias in aliases:
            if alias in df.columns:
                df = df.rename(columns={alias: canonical})
                renamed[canonical] = canonical
                break

    required = ["open", "high", "low", "close"]
    missing = [r for r in required if r not in df.columns]
    if missing:
        raise ValueError(f"CSV missing required columns: {missing}. "
                         f"Detected columns: {list(df.columns)}")

    if "volume" not in df.columns:
        df["volume"] = 0.0
    if "vwap" not in df.columns:
        df["vwap"] = df["close"]

    for col in ["open", "high", "low", "close", "volume", "vwap"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["open", "high", "low", "close"])
    return df[["open", "high", "low", "close", "volume", "vwap"]]


def fetch_alpaca_data(ticker: str, start: str, end: str,
                      timeframe_label: str,
                      api_key: str, secret_key: str) -> pd.DataFrame:
    """Fetch OHLCV bars from Alpaca Markets free data API."""
    try:
        from alpaca.data.historical import StockHistoricalDataClient
        from alpaca.data.requests import StockBarsRequest
        from alpaca.data.timeframe import TimeFrame, TimeFrameUnit
    except ImportError:
        raise RuntimeError("alpaca-py not installed. Run: pip install alpaca-py")

    if not api_key or not secret_key:
        raise ValueError("Alpaca API Key and Secret Key are required.")

    client = StockHistoricalDataClient(api_key=api_key, secret_key=secret_key)

    tf_str, _ = TF_MAP.get(timeframe_label, ("1Day", 1440))
    tf_map_alpaca = {
        "1Min":  TimeFrame(1,  TimeFrameUnit.Minute),
        "5Min":  TimeFrame(5,  TimeFrameUnit.Minute),
        "15Min": TimeFrame(15, TimeFrameUnit.Minute),
        "30Min": TimeFrame(30, TimeFrameUnit.Minute),
        "1Hour": TimeFrame(1,  TimeFrameUnit.Hour),
        "4Hour": TimeFrame(4,  TimeFrameUnit.Hour),
        "1Day":  TimeFrame(1,  TimeFrameUnit.Day),
    }
    alpaca_tf = tf_map_alpaca[tf_str]

    start_dt = datetime.strptime(start, "%Y-%m-%d").replace(
        tzinfo=pytz.timezone("America/New_York"))
    end_dt   = datetime.strptime(end,   "%Y-%m-%d").replace(
        hour=23, minute=59, second=59,
        tzinfo=pytz.timezone("America/New_York"))

    req = StockBarsRequest(
        symbol_or_symbols=ticker.upper(),
        timeframe=alpaca_tf,
        start=start_dt,
        end=end_dt,
        adjustment="all",
    )
    bars = client.get_stock_bars(req)
    df = bars.df

    if df.empty:
        raise ValueError(f"No data returned for {ticker} in the given date range.")

    # Alpaca returns MultiIndex (symbol, timestamp) for single symbol too
    if isinstance(df.index, pd.MultiIndex):
        df = df.reset_index(level=0, drop=True)

    df.index = pd.to_datetime(df.index, utc=True)
    df.index.name = "timestamp"
    df.columns = [c.lower() for c in df.columns]

    rename = {"vwap": "vwap", "trade_count": "trade_count"}
    for src, dst in rename.items():
        if src in df.columns:
            df = df.rename(columns={src: dst})

    required = ["open", "high", "low", "close"]
    missing  = [r for r in required if r not in df.columns]
    if missing:
        raise ValueError(f"Alpaca data missing columns: {missing}")

    if "volume" not in df.columns:
        df["volume"] = 0.0
    if "vwap" not in df.columns:
        df["vwap"] = df["close"]

    return df[["open", "high", "low", "close", "volume", "vwap"]].sort_index()


# ════════════════════════════════════════════════════════════════════════════
# BACKTEST ENGINE
# ════════════════════════════════════════════════════════════════════════════

def run_backtest(data: pd.DataFrame, strategy_code: str,
                 strategy_name: str) -> dict:
    """
    Execute strategy code, compute portfolio returns, return QuantStats stats.
    Returns dict with keys: returns, benchmark_returns, stats_df, html_path
    """
    # ── Execute strategy code ────────────────────────────────────────────
    ns = {"pd": pd, "np": np, "__builtins__": __builtins__}
    try:
        exec(compile(strategy_code, "<strategy>", "exec"), ns)
    except Exception as e:
        raise RuntimeError(f"Strategy code compilation error:\n{e}")

    if "strategy" not in ns:
        raise ValueError("Strategy code must define a function named `strategy(data)`.")

    strategy_fn = ns["strategy"]
    try:
        signals = strategy_fn(data.copy())
    except Exception as e:
        raise RuntimeError(f"Strategy execution error:\n{traceback.format_exc()}")

    if not isinstance(signals, pd.Series):
        raise ValueError("`strategy()` must return a pandas Series of signals.")

    signals = signals.reindex(data.index).fillna(0)
    signals = signals.shift(1).fillna(0)   # next-bar execution (avoid look-ahead)

    # ── Compute returns ──────────────────────────────────────────────────
    price_returns  = data["close"].pct_change().fillna(0)
    portfolio_rets = (signals * price_returns).rename(strategy_name)

    # ── Force consistent timezone handling ──────────────────────────────
    portfolio_rets.index = pd.to_datetime(portfolio_rets.index, utc=True)

    # Daily aggregation
    daily_rets = portfolio_rets.resample("D").apply(
        lambda x: (1 + x).prod() - 1).dropna()

    daily_rets = daily_rets[daily_rets.abs() < 1]

    # Make timezone NAIVE for QuantStats compatibility
    if daily_rets.index.tz is not None:
        daily_rets.index = daily_rets.index.tz_convert("UTC").tz_localize(None)

    if daily_rets.empty:
        raise ValueError("Backtest produced no returns. Check date range and strategy logic.")

    # ── Benchmark: buy-and-hold ──────────────────────────────────────────
    bh_daily = price_returns.resample("D").apply(
        lambda x: (1 + x).prod() - 1).dropna()

    bh_daily = bh_daily.reindex(daily_rets.index).fillna(0)

    # Ensure same dtype
    if bh_daily.index.tz is not None:
        bh_daily.index = bh_daily.index.tz_convert("UTC").tz_localize(None)

    bh_daily.name = "Buy & Hold"

    # ── Generate HTML Tearsheet ──────────────────────────────────────────
    html_path = os.path.join(tempfile.gettempdir(),
                             f"backtest_{strategy_name.replace(' ','_')}_{datetime.now():%Y%m%d_%H%M%S}.html")

    qs.extend_pandas()
    try:
        qs.reports.html(
            returns=daily_rets,
            benchmark=bh_daily,
            title=strategy_name,
            output=html_path,
            download_filename=html_path,
        )
    except Exception as e:
        # Fallback: basic stats only
        raise RuntimeError(f"QuantStats report generation failed:\n{e}")

    # ── Quick summary stats ──────────────────────────────────────────────
    summary = {
        "Total Return":     f"{qs.stats.comp(daily_rets)*100:.2f}%",
        "CAGR":             f"{qs.stats.cagr(daily_rets)*100:.2f}%",
        "Sharpe Ratio":     f"{qs.stats.sharpe(daily_rets):.3f}",
        "Sortino Ratio":    f"{qs.stats.sortino(daily_rets):.3f}",
        "Max Drawdown":     f"{qs.stats.max_drawdown(daily_rets)*100:.2f}%",
        "Volatility (Ann)": f"{qs.stats.volatility(daily_rets)*100:.2f}%",
        "Win Rate":         f"{qs.stats.win_rate(daily_rets)*100:.2f}%",
        "Calmar Ratio":     f"{qs.stats.calmar(daily_rets):.3f}",
        "Skew":             f"{qs.stats.skew(daily_rets):.3f}",
        "Kurtosis":         f"{qs.stats.kurtosis(daily_rets):.3f}",
        "VaR (95%)":        f"{qs.stats.value_at_risk(daily_rets)*100:.2f}%",
        "CVaR (95%)":       f"{qs.stats.cvar(daily_rets)*100:.2f}%",
    }

    bars_traded = int((signals.abs() > 0).sum())
    summary["Bars in Market"] = f"{bars_traded:,} bars"
    summary["Total Bars"]     = f"{len(data):,} bars"

    return {
        "summary":   summary,
        "html_path": html_path,
        "daily_rets": daily_rets,
        "data_rows": len(data),
    }


# ════════════════════════════════════════════════════════════════════════════
# TRADINGVIEW LIST OF TRADES ANALYZER
# ════════════════════════════════════════════════════════════════════════════

def analyze_tradingview_trades(filepath: str) -> dict:
    """
    Analyze TradingView exported backtest XLSX file
    Expects sheet: 'List of trades'
    """

    import pandas as pd
    import numpy as np

    xls = pd.ExcelFile(filepath)
    sheet_name = None

    # Detect correct sheet dynamically
    for s in xls.sheet_names:
        if "list" in s.lower() and "trade" in s.lower():
            sheet_name = s
            break

    if sheet_name is None:
        raise ValueError("Could not find 'List of trades' sheet.")

    df = pd.read_excel(filepath, sheet_name=sheet_name)
    df.columns = [c.strip().lower() for c in df.columns]

    # Detect profit column automatically
    print("Detected columns:", df.columns.tolist())
    
    profit_col = None

    for c in df.columns:
        lc = c.lower()

        if any(key in lc for key in [
            "net p&l",
            "p&l",
            "net profit",
            "profit",
            "result"
        ]) and "%" not in lc:
            profit_col = c
            break

    if profit_col is None:
        raise ValueError(
            f"Could not detect Profit column.\nDetected columns:\n{list(df.columns)}"
        )


    if profit_col is None:
        raise ValueError("Could not detect Profit column.")

    df["profit"] = pd.to_numeric(df[profit_col], errors="coerce")
    df = df.dropna(subset=["profit"])

    total_trades = len(df)
    wins = df[df["profit"] > 0]
    losses = df[df["profit"] < 0]

    gross_profit = wins["profit"].sum()
    gross_loss = losses["profit"].sum()

    profit_factor = (
        abs(gross_profit / gross_loss)
        if gross_loss != 0 else float("inf")
    )

    avg_win = wins["profit"].mean() if not wins.empty else 0
    avg_loss = losses["profit"].mean() if not losses.empty else 0

    expectancy = (
        (wins.shape[0] / total_trades) * avg_win +
        (losses.shape[0] / total_trades) * avg_loss
    ) if total_trades > 0 else 0

    win_rate = wins.shape[0] / total_trades if total_trades > 0 else 0

    equity_curve = df["profit"].cumsum()
    max_dd = (equity_curve.cummax() - equity_curve).max()

    return {
        "Total Trades": total_trades,
        "Win Rate %": round(win_rate * 100, 2),
        "Profit Factor": round(profit_factor, 3) if np.isfinite(profit_factor) else "∞",
        "Expectancy": round(expectancy, 4),
        "Gross Profit": round(gross_profit, 2),
        "Gross Loss": round(gross_loss, 2),
        "Average Win": round(avg_win, 2),
        "Average Loss": round(avg_loss, 2),
        "Max Drawdown": round(max_dd, 2),
        "Max Win": round(df["profit"].max(), 2),
        "Max Loss": round(df["profit"].min(), 2),
        "Median Trade": round(df["profit"].median(), 2),
        "Std Dev": round(df["profit"].std(), 4),
    }

# ════════════════════════════════════════════════════════════════════════════
# GUI APPLICATION
# ════════════════════════════════════════════════════════════════════════════

class BacktestApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QuantStats Backtest Generator")
        self.configure(bg=BG)
        self.geometry("1300x860")
        self.minsize(1100, 700)

        self._csv_path   = tk.StringVar()
        self._ticker     = tk.StringVar(value="SOFI")
        self._start_date = tk.StringVar(value="2020-01-01")
        self._end_date   = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        self._timeframe  = tk.StringVar(value="1 Day")
        self._data_src   = tk.StringVar(value="alpaca")
        self._api_key    = tk.StringVar()
        self._secret_key = tk.StringVar()
        self._strat_name = tk.StringVar(value="My Strategy")
        self._status     = tk.StringVar(value="Ready.")
        self._html_path  = None

        self._apply_styles()
        self._build_ui()

    # ── Style ─────────────────────────────────────────────────────────────
    def _apply_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        base = dict(background=BG2, foreground=FG, fieldbackground=BG3,
                    insertbackground=FG, selectbackground=ACCENT,
                    selectforeground=BG, bordercolor=BORDER,
                    lightcolor=BORDER, darkcolor=BORDER,
                    relief="flat", padding=4)
        s.configure("TFrame",      background=BG)
        s.configure("Card.TFrame", background=BG2, relief="flat")
        s.configure("TLabel",      background=BG,  foreground=FG,  font=FONT_UI)
        s.configure("Card.TLabel", background=BG2, foreground=FG,  font=FONT_UI)
        s.configure("Dim.TLabel",  background=BG2, foreground=FG2, font=("", 9))
        s.configure("Head.TLabel", background=BG,  foreground=ACCENT,
                    font=(FONT_HEAD[0], 12, "bold"))
        s.configure("TEntry",      **base)
        s.configure("TCombobox",   **base)
        s.map("TCombobox", fieldbackground=[("readonly", BG3)],
              background=[("readonly", BG3)])
        s.configure("TButton", background=BG3, foreground=FG,
                    bordercolor=BORDER, font=FONT_UI, padding=(10, 5),
                    relief="flat")
        s.map("TButton",
              background=[("active", BORDER)],
              foreground=[("active", FG)])
        s.configure("Run.TButton", background=ACCENT, foreground=BG,
                    font=(FONT_UI[0], 11, "bold"), padding=(16, 8))
        s.map("Run.TButton", background=[("active", "#79c0ff")])
        s.configure("Open.TButton", background=ACCENT2, foreground=BG,
                    font=(FONT_UI[0], 10, "bold"), padding=(12, 6))
        s.map("Open.TButton", background=[("active", "#56d364")])
        s.configure("TRadiobutton", background=BG2, foreground=FG,
                    indicatorbackground=BG3, selectcolor=ACCENT)
        s.configure("TNotebook",     background=BG,  bordercolor=BORDER, tabmargins=0)
        s.configure("TNotebook.Tab", background=BG3, foreground=FG2,
                    padding=(14, 6), bordercolor=BORDER)
        s.map("TNotebook.Tab",
              background=[("selected", BG2)],
              foreground=[("selected", FG)])
        s.configure("TProgressbar", troughcolor=BG3, background=ACCENT,
                    bordercolor=BORDER, thickness=4)
        s.configure("Horizontal.TSeparator", background=BORDER)

    # ── Build UI ───────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=BG2, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="⚡ QuantStats Backtest Generator",
                 bg=BG2, fg=FG, font=(FONT_HEAD[0], 15, "bold")).pack(
                     side="left", padx=20, pady=12)
        tk.Label(hdr, text="Powered by QuantStats + Alpaca / CSV",
                 bg=BG2, fg=FG2, font=("", 9)).pack(side="right", padx=20)

        sep = ttk.Separator(self, orient="horizontal")
        sep.pack(fill="x")

        # Main content
        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=0, pady=0)

        # Left panel
        left = tk.Frame(main, bg=BG, width=380)
        left.pack(side="left", fill="y", padx=0)
        left.pack_propagate(False)

        self._build_config_panel(left)


        # Divider
        tk.Frame(main, bg=BORDER, width=1).pack(side="left", fill="y")

        # Right panel
        right = tk.Frame(main, bg=BG)
        right.pack(side="left", fill="both", expand=True)

        self._build_right_panel(right)

        # Status bar
        status_bar = tk.Frame(self, bg=BG3, height=26)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)
        self._prog = ttk.Progressbar(status_bar, mode="indeterminate",
                                     style="TProgressbar", length=120)
        self._prog.pack(side="right", padx=10, pady=5)
        tk.Label(status_bar, textvariable=self._status,
                 bg=BG3, fg=FG2, font=("", 9),
                 anchor="w").pack(side="left", padx=10, pady=4)

    def _build_config_panel(self, parent):
        canvas = tk.Canvas(parent, bg=BG, bd=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=BG)

        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        p = 16

        def section(title):
            f = tk.Frame(inner, bg=BG2, bd=0)
            f.pack(fill="x", padx=p, pady=(10, 0))
            tk.Label(f, text=title, bg=BG2, fg=ACCENT,
                     font=(FONT_HEAD[0], 10, "bold")).pack(
                         anchor="w", padx=10, pady=(8, 4))
            return f

        def row(parent, label, widget_factory, **kw):
            r = tk.Frame(parent, bg=BG2)
            r.pack(fill="x", padx=10, pady=3)
            tk.Label(r, text=label, bg=BG2, fg=FG2,
                     font=("", 9), width=14, anchor="w").pack(side="left")
            w = widget_factory(r, **kw)
            w.pack(side="left", fill="x", expand=True)
            return w

        def entry(parent, textvariable=None, show=None, **kw):
            e = ttk.Entry(parent, textvariable=textvariable,
                          font=FONT_MONO, show=show)
            return e

        # ─ Strategy Name ─────────────────────────────────────────────────
        f0 = section("◈  STRATEGY")
        row(f0, "Name", lambda p: ttk.Entry(
            p, textvariable=self._strat_name, font=FONT_UI))

        # ─ Data Source ───────────────────────────────────────────────────
        f1 = section("◈  DATA SOURCE")
        rb_frame = tk.Frame(f1, bg=BG2)
        rb_frame.pack(fill="x", padx=10, pady=4)
        for text, val in [("Alpaca API (Free)", "alpaca"), ("CSV File", "csv")]:
            ttk.Radiobutton(
                rb_frame, text=text, variable=self._data_src, value=val,
                command=self._toggle_source,
                style="TRadiobutton"
            ).pack(side="left", padx=(0, 16))

        # Alpaca fields
        self._alpaca_frame = tk.Frame(f1, bg=BG2)
        self._alpaca_frame.pack(fill="x")
        row(self._alpaca_frame, "API Key",
            lambda p: ttk.Entry(p, textvariable=self._api_key, font=FONT_MONO))
        row(self._alpaca_frame, "Secret Key",
            lambda p: ttk.Entry(p, textvariable=self._secret_key,
                                font=FONT_MONO, show="•"))
        tk.Label(self._alpaca_frame,
                 text="  Get free keys at alpaca.markets",
                 bg=BG2, fg=FG2, font=("", 8)).pack(anchor="w", padx=10, pady=(0,6))

        # CSV fields
        self._csv_frame = tk.Frame(f1, bg=BG2)
        csv_r = tk.Frame(self._csv_frame, bg=BG2)
        csv_r.pack(fill="x", padx=10, pady=4)
        tk.Label(csv_r, text="CSV File", bg=BG2, fg=FG2,
                 font=("", 9), width=14, anchor="w").pack(side="left")
        ttk.Entry(csv_r, textvariable=self._csv_path,
                  font=("", 9)).pack(side="left", fill="x", expand=True, padx=(0,4))
        ttk.Button(csv_r, text="Browse",
                   command=self._browse_csv).pack(side="left")
        tk.Label(self._csv_frame,
                 text="  Columns: timestamp, open, high, low, close, volume",
                 bg=BG2, fg=FG2, font=("", 8)).pack(anchor="w", padx=10, pady=(0,6))

        # ─ Ticker & Time ─────────────────────────────────────────────────
        f2 = section("◈  TICKER & TIMEFRAME")
        row(f2, "Ticker",
            lambda p: ttk.Entry(p, textvariable=self._ticker, font=FONT_MONO))
        row(f2, "Timeframe",
            lambda p: ttk.Combobox(
                p, textvariable=self._timeframe, state="readonly",
                values=list(TF_MAP.keys()), font=FONT_UI))
        row(f2, "Start Date",
            lambda p: ttk.Entry(p, textvariable=self._start_date, font=FONT_MONO))
        row(f2, "End Date",
            lambda p: ttk.Entry(p, textvariable=self._end_date, font=FONT_MONO))
        tk.Label(f2, text="  Format: YYYY-MM-DD",
                 bg=BG2, fg=FG2, font=("", 8)).pack(anchor="w", padx=10, pady=(0,8))

        # ─ Run button ────────────────────────────────────────────────────
        btn_f = tk.Frame(inner, bg=BG)
        btn_f.pack(fill="x", padx=p, pady=14)
        ttk.Button(btn_f, text="▶  Run Backtest",
                   style="Run.TButton",
                   command=self._run_threaded).pack(fill="x")

        ttk.Button(
            btn_f,
            text="📊 Analyze TradingView Backtest",
            command=self._analyze_tradingview_file
        ).pack(fill="x", pady=(8, 0))


        self._open_btn = ttk.Button(
            btn_f, text="🌐  Open Tearsheet in Browser",
            style="Open.TButton",
            command=self._open_report,
            state="disabled")
        self._open_btn.pack(fill="x", pady=(8, 0))

        tk.Frame(inner, bg=BG, height=10).pack()
        self._toggle_source()

    def _build_right_panel(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True, padx=0, pady=0)

        # ── Tab 1: Strategy Editor ─────────────────────────────────────
        editor_tab = tk.Frame(nb, bg=BG2)
        nb.add(editor_tab, text="  Strategy Editor  ")

        toolbar = tk.Frame(editor_tab, bg=BG3, height=34)
        toolbar.pack(fill="x")
        toolbar.pack_propagate(False)
        tk.Label(toolbar, text="strategy.py", bg=BG3,
                 fg=FG2, font=FONT_MONO).pack(side="left", padx=12, pady=6)
        ttk.Button(toolbar, text="Reset to Default",
                   command=self._reset_strategy).pack(side="right", padx=8, pady=4)
        ttk.Button(toolbar, text="Clear",
                   command=lambda: (self._editor.delete("1.0", "end"))).pack(
                       side="right", padx=0, pady=4)

        self._editor = scrolledtext.ScrolledText(
            editor_tab,
            bg="#0d1117", fg="#e6edf3",
            font=FONT_MONO,
            insertbackground=FG,
            selectbackground=ACCENT,
            selectforeground=BG,
            relief="flat", bd=0,
            wrap="none",
            undo=True,
            tabs="1c",
        )
        self._editor.pack(fill="both", expand=True, padx=0, pady=0)
        self._editor.insert("1.0", DEFAULT_STRATEGY)
        self._apply_syntax_highlight()

        # ── Tab 2: Results Summary ─────────────────────────────────────
        results_tab = tk.Frame(nb, bg=BG2)
        nb.add(results_tab, text="  Results  ")
        self._results_frame = results_tab
        self._build_results_area(results_tab)

        # ── Tab 3: Log ─────────────────────────────────────────────────
        log_tab = tk.Frame(nb, bg=BG2)
        nb.add(log_tab, text="  Log  ")
        self._log = scrolledtext.ScrolledText(
            log_tab, bg="#0d1117", fg="#8b949e",
            font=("Cascadia Code", 9) if sys.platform=="win32" else ("Menlo", 9),
            relief="flat", bd=0, state="disabled",
            wrap="word",
        )
        self._log.pack(fill="both", expand=True)
        self._log.tag_config("ok",    foreground=ACCENT2)
        self._log.tag_config("err",   foreground=DANGER)
        self._log.tag_config("warn",  foreground=WARNING)
        self._log.tag_config("info",  foreground=ACCENT)
        self._log.tag_config("plain", foreground=FG2)

        self._nb = nb

    def _build_results_area(self, parent):
        placeholder = tk.Frame(parent, bg=BG2)
        placeholder.pack(fill="both", expand=True)
        tk.Label(placeholder,
                 text="Run a backtest to see results here.",
                 bg=BG2, fg=FG2,
                 font=(FONT_UI[0], 12)).pack(expand=True)
        self._results_placeholder = placeholder

    # ── Helpers ────────────────────────────────────────────────────────────
    def _toggle_source(self):
        src = self._data_src.get()
        if src == "alpaca":
            self._alpaca_frame.pack(fill="x")
            self._csv_frame.pack_forget()
        else:
            self._csv_frame.pack(fill="x")
            self._alpaca_frame.pack_forget()

    def _browse_csv(self):
        path = filedialog.askopenfilename(
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
        if path:
            self._csv_path.set(path)

    def _analyze_tradingview_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if not path:
            return

        try:
            stats = analyze_tradingview_trades(path)
            msg = "\n".join([f"{k}: {v}" for k, v in stats.items()])
            messagebox.showinfo("TradingView Trade Analysis", msg)
        except Exception as e:
            messagebox.showerror("Error", str(e))


    def _reset_strategy(self):
        self._editor.delete("1.0", "end")
        self._editor.insert("1.0", DEFAULT_STRATEGY)

    def _apply_syntax_highlight(self):
        """Basic Python keyword highlighting."""
        kw_color = "#ff7b72"
        str_color = "#a5d6ff"
        comment_color = "#8b949e"

        content = self._editor.get("1.0", "end")
        # Clear existing tags
        for tag in ["kw", "str", "comment", "num"]:
            self._editor.tag_remove(tag, "1.0", "end")

        self._editor.tag_config("kw",      foreground=kw_color)
        self._editor.tag_config("str",     foreground=str_color)
        self._editor.tag_config("comment", foreground=comment_color)
        self._editor.tag_config("num",     foreground="#79c0ff")
        self._editor.tag_config("func",    foreground="#d2a8ff")

        keywords = r"\b(def|class|import|from|return|if|else|elif|for|while|in|not|and|or|True|False|None|try|except|raise|with|as|lambda|yield|pass|break|continue|global|nonlocal|del|is)\b"
        for m in re.finditer(keywords, content):
            s = f"1.0 + {m.start()} chars"
            e = f"1.0 + {m.end()} chars"
            self._editor.tag_add("kw", s, e)

        for m in re.finditer(r"(\"\"\".*?\"\"\"|'''.*?'''|\"[^\"]*\"|'[^']*')",
                              content, re.DOTALL):
            self._editor.tag_add("str",
                                  f"1.0 + {m.start()} chars",
                                  f"1.0 + {m.end()} chars")

        for m in re.finditer(r"#[^\n]*", content):
            self._editor.tag_add("comment",
                                  f"1.0 + {m.start()} chars",
                                  f"1.0 + {m.end()} chars")

    def _log_write(self, msg, tag="plain"):
        self._log.configure(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.insert("end", f"[{ts}] {msg}\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _set_status(self, msg, color=FG2):
        self._status.set(msg)

    # ── Run Backtest ───────────────────────────────────────────────────────
    def _run_threaded(self):
        self._prog.start(12)
        self._set_status("Running backtest...")
        self._log_write("═" * 60, "info")
        self._log_write(f"Starting backtest: {self._strat_name.get()}", "info")
        t = threading.Thread(target=self._run_backtest_task, daemon=True)
        t.start()

    def _run_backtest_task(self):
        try:
            # 1. Load data
            self._log_write("Loading data...", "plain")
            data = self._load_data()
            self._log_write(
                f"Data loaded: {len(data):,} bars | "
                f"{data.index[0].date()} → {data.index[-1].date()}", "ok")

            # 2. Get strategy code
            code = self._editor.get("1.0", "end")

            # 3. Run
            self._log_write("Executing strategy...", "plain")
            result = run_backtest(data, code, self._strat_name.get())

            # 4. Display
            self._html_path = result["html_path"]
            self.after(0, lambda: self._display_results(result))
            self._log_write("Backtest complete!", "ok")
            for k, v in result["summary"].items():
                self._log_write(f"  {k:<22} {v}", "plain")
            self._log_write(f"Tearsheet saved: {self._html_path}", "ok")

        except Exception as e:
            err = traceback.format_exc()
            self._log_write(f"ERROR: {e}", "err")
            self._log_write(err, "err")
            self.after(0, lambda: messagebox.showerror(
                "Backtest Error", str(e)))
            self.after(0, lambda: self._set_status(f"Error: {e}"))
        finally:
            self.after(0, self._prog.stop)

    def _load_data(self) -> pd.DataFrame:
        src = self._data_src.get()
        ticker = self._ticker.get().strip().upper()
        start  = self._start_date.get().strip()
        end    = self._end_date.get().strip()
        tf     = self._timeframe.get()

        if src == "alpaca":
            return fetch_alpaca_data(
                ticker, start, end, tf,
                self._api_key.get().strip(),
                self._secret_key.get().strip(),
            )
        else:
            csv_path = self._csv_path.get().strip()
            if not csv_path or not os.path.exists(csv_path):
                raise ValueError("Please select a valid CSV file.")
            df = detect_and_load_csv(csv_path)

            # Apply date filter
            try:
                s = pd.Timestamp(start, tz="UTC")
                e = pd.Timestamp(end,   tz="UTC") + pd.Timedelta(days=1)
                mask = (df.index >= s) & (df.index <= e)
                df = df[mask]
            except Exception:
                pass

            if df.empty:
                raise ValueError("CSV data is empty after date filtering.")
            return df

    def _display_results(self, result: dict):
        # Clear placeholder
        for w in self._results_frame.winfo_children():
            w.destroy()

        # ── Header ───────────────────────────────────────────────────────
        hdr = tk.Frame(self._results_frame, bg=BG2)
        hdr.pack(fill="x", padx=20, pady=(16, 0))
        tk.Label(hdr, text=f"✓  {self._strat_name.get()}",
                 bg=BG2, fg=ACCENT2,
                 font=(FONT_HEAD[0], 14, "bold")).pack(side="left")
        ttk.Button(hdr, text="🌐 Open Full Tearsheet",
                   style="Open.TButton",
                   command=self._open_report).pack(side="right")

        ttk.Separator(self._results_frame, orient="horizontal").pack(
            fill="x", padx=20, pady=12)

        # ── Stats grid ───────────────────────────────────────────────────
        grid = tk.Frame(self._results_frame, bg=BG2)
        grid.pack(fill="x", padx=20)

        summary = result["summary"]
        items = list(summary.items())
        cols = 3
        for i, (k, v) in enumerate(items):
            row_, col_ = divmod(i, cols)
            cell = tk.Frame(grid, bg=BG3, bd=0)
            cell.grid(row=row_, column=col_,
                      padx=6, pady=6, sticky="ew")
            grid.columnconfigure(col_, weight=1)

            color = ACCENT2 if "%" in v and not v.startswith("-") else \
                    DANGER  if v.startswith("-") else FG
            if "Drawdown" in k and v.startswith("-"):
                color = DANGER
            elif "Drawdown" in k:
                color = WARNING

            tk.Label(cell, text=k, bg=BG3, fg=FG2,
                     font=("", 9)).pack(anchor="w", padx=10, pady=(8, 2))
            tk.Label(cell, text=v, bg=BG3, fg=color,
                     font=(FONT_HEAD[0], 14, "bold")).pack(
                         anchor="w", padx=10, pady=(0, 8))

        # ── Sparkline (ASCII-ish equity) ──────────────────────────────────
        try:
            self._draw_equity_curve(result["daily_rets"])
        except Exception:
            pass

        # Enable open button
        self._open_btn.configure(state="normal")
        self._nb.select(1)  # Switch to Results tab
        self._set_status("Backtest complete. Tearsheet ready.")

    def _draw_equity_curve(self, returns: pd.Series):
        equity = (1 + returns).cumprod()

        chart_f = tk.Frame(self._results_frame, bg=BG2)
        chart_f.pack(fill="x", padx=20, pady=(14, 20))
        tk.Label(chart_f, text="Equity Curve (preview)",
                 bg=BG2, fg=FG2, font=("", 9)).pack(anchor="w", padx=8)

        canvas = tk.Canvas(chart_f, bg=BG3, height=120,
                           highlightthickness=0, bd=0)
        canvas.pack(fill="x", padx=8, pady=4)
        canvas.update()

        W = canvas.winfo_width() or 800
        H = 120
        pad = 10

        vals = equity.values
        mn, mx = vals.min(), vals.max()
        rng = (mx - mn) or 1e-9

        def yx(i, v):
            x = pad + (i / (len(vals) - 1)) * (W - 2 * pad)
            y = H - pad - ((v - mn) / rng) * (H - 2 * pad)
            return x, y

        pts = [yx(i, v) for i, v in enumerate(vals)]
        flat_pts = [c for pt in pts for c in pt]

        # Gradient fill (approximation with lines)
        for i in range(len(pts) - 1):
            x1, y1 = pts[i]
            x2, y2 = pts[i + 1]
            canvas.create_polygon(
                x1, y1, x2, y2, x2, H - pad, x1, H - pad,
                fill="#1f4068", outline="")

        # Line
        if len(flat_pts) >= 4:
            canvas.create_line(*flat_pts, fill=ACCENT, width=2, smooth=True)

        # Baseline
        base_y = yx(0, 1.0)[1]
        canvas.create_line(pad, base_y, W - pad, base_y,
                           fill=BORDER, dash=(4, 4))

        # Labels
        canvas.create_text(W - pad, pad, text=f"{mx:.2f}x",
                           fill=FG2, anchor="ne", font=("", 8))
        canvas.create_text(W - pad, H - pad, text=f"{mn:.2f}x",
                           fill=FG2, anchor="se", font=("", 8))

    def _open_report(self):
        if self._html_path and os.path.exists(self._html_path):
            webbrowser.open(f"file://{self._html_path}")
        else:
            messagebox.showwarning("No Report",
                                   "Run a backtest first to generate the tearsheet.")


# ════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = BacktestApp()
    app.mainloop()
