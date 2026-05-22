"""
X Strategy with CPR Validation v4 - Long Only (Python Conversion)
=================================================================
Converted from Pine Script® by Pivot-in-action.

This strategy calculates pivot levels (X, TC, BC) from Day-2 OHLC data,
validates them against Day-1 price action, and enters long positions
when intraday price interacts with validated X levels.

Two entry types:
  - Type A (Standard): Day-1 traded entirely ABOVE the CPR zone (bullish confirmation)
  - Type B (Delayed):  Day-1 traded entirely BELOW the CPR zone (reversal setup)

Entry triggers (same for both types):
  - Standard Breach: Day opens >= X, candle low touches X, candle open >= X
  - Recovery Entry:   Day opens < X, candle opens < X but closes > X

Exit rules:
  - ATR trailing stop (configurable)
  - End-of-day exit at 15:30

Requirements:
  pip install pandas numpy
"""

import pandas as pd
import numpy as np
from dataclasses import dataclass, field
from typing import Optional
import warnings

warnings.filterwarnings("ignore")


# ═══════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════

@dataclass
class StrategyConfig:
    """All tunable parameters — mirrors the Pine Script input() calls."""

    max_alerts: int = 10               # Max alert history to keep
    use_fixed_qty: bool = False        # True = fixed shares, False = % of equity
    fixed_qty: float = 50.0            # Shares when use_fixed_qty is True
    default_qty_pct: float = 100.0     # % of equity when use_fixed_qty is False
    use_atr_trail: bool = True         # Enable ATR trailing stop
    atr_length: int = 14              # ATR lookback period
    atr_multiplier: float = 3.0       # ATR multiplier for trail distance
    enable_delayed_x: bool = True     # Enable Type B (delayed) setups
    initial_capital: float = 100_000.0 # Starting equity


# ═══════════════════════════════════════════════════════════════
# DATA PREPARATION
# ═══════════════════════════════════════════════════════════════

def compute_atr(df: pd.DataFrame, length: int = 14) -> pd.Series:
    """
    Average True Range — same as ta.atr(length) in Pine Script.

    True Range = max(H-L, |H - prev_close|, |L - prev_close|)
    ATR = RMA (Wilder's smoothing) of True Range over `length` bars.
    """
    high = df["high"]
    low = df["low"]
    prev_close = df["close"].shift(1)

    tr = pd.concat([
        high - low,
        (high - prev_close).abs(),
        (low - prev_close).abs()
    ], axis=1).max(axis=1)

    # Wilder's smoothing (RMA) — equivalent to Pine's ta.rma / ta.atr
    atr = tr.ewm(alpha=1.0 / length, min_periods=length, adjust=False).mean()
    return atr


def prepare_daily_ohlc(df: pd.DataFrame) -> pd.DataFrame:
    """
    From intraday bars, build a daily OHLC table.
    Expects `df` to have a DatetimeIndex and columns: open, high, low, close.
    """
    daily = df.resample("1D").agg({
        "open": "first",
        "high": "max",
        "low": "min",
        "close": "last"
    }).dropna()
    return daily


def add_daily_context(df: pd.DataFrame, daily: pd.DataFrame) -> pd.DataFrame:
    """
    Attach daily-level columns to every intraday bar:
      - day2_high/low/close  → previous-previous day (Day-2 in Pine Script)
      - day1_high/low/close  → previous day (Day-1)
      - day_open_price       → today's opening price
      - date                 → calendar date of the bar
    """
    df = df.copy()
    df["date"] = df.index.date

    # Shift daily rows: day1 = yesterday, day2 = day before yesterday
    daily = daily.copy()
    daily["day1_high"] = daily["high"]       # "current" daily bar (used as Day-1 for next day)
    daily["day1_low"] = daily["low"]
    daily["day1_close"] = daily["close"]
    daily["day2_high"] = daily["high"].shift(1)
    daily["day2_low"] = daily["low"].shift(1)
    daily["day2_close"] = daily["close"].shift(1)
    daily["day_open_price"] = daily["open"]  # Today's open

    # The Pine Script uses request.security with [1] offsets.
    # On any given trading day, Day-2 data = two calendar days back,
    # Day-1 data = one calendar day back.
    # We shift by 1 so that on "today" we see yesterday as Day-1 and
    # day-before-yesterday as Day-2.
    daily_shifted = daily[
        ["day2_high", "day2_low", "day2_close",
         "day1_high", "day1_low", "day1_close",
         "day_open_price"]
    ].shift(1)  # shift so today's row carries *yesterday's* values as Day-1

    # Actually let me re-derive this more carefully to match Pine Script exactly.
    #
    # In Pine Script:
    #   day2_high = request.security(syminfo.tickerid, "1D", high[1])  → high of 2 days ago
    #   day2_low  = request.security(syminfo.tickerid, "1D", low[1])   → low of 2 days ago
    #   day2_close= request.security(syminfo.tickerid, "1D", close[1]) → close of 2 days ago
    #   day1_high = request.security(syminfo.tickerid, "1D", high)     → high of previous completed day
    #   day1_low  = request.security(syminfo.tickerid, "1D", low)      → low of previous completed day
    #   day1_close= request.security(syminfo.tickerid, "1D", close)    → close of previous completed day
    #
    # request.security on "1D" with high[1] gives the high from 2 trading days ago.
    # request.security on "1D" with high (no offset) gives the high from the most recently
    # *completed* daily bar (i.e., yesterday).
    #
    # So on any given day:
    #   Day-2 = 2 trading days back
    #   Day-1 = 1 trading day back (yesterday)

    daily_lookup = pd.DataFrame(index=daily.index)
    daily_lookup["day2_high"] = daily["high"].shift(2)
    daily_lookup["day2_low"] = daily["low"].shift(2)
    daily_lookup["day2_close"] = daily["close"].shift(2)
    daily_lookup["day1_high"] = daily["high"].shift(1)
    daily_lookup["day1_low"] = daily["low"].shift(1)
    daily_lookup["day1_close"] = daily["close"].shift(1)
    daily_lookup["day_open_price"] = daily["open"]

    daily_lookup.index = daily_lookup.index.date  # convert to date for merge
    daily_lookup.index.name = "date"

    df = df.merge(daily_lookup, left_on="date", right_index=True, how="left")
    return df


# ═══════════════════════════════════════════════════════════════
# CORE STRATEGY LOGIC
# ═══════════════════════════════════════════════════════════════

@dataclass
class TradeRecord:
    """One completed round-trip trade."""
    entry_time: pd.Timestamp
    entry_price: float
    exit_time: pd.Timestamp
    exit_price: float
    entry_type: str          # "Type A Standard Breach", "Type B Recovery Entry", etc.
    exit_reason: str         # "ATR Trail Stop", "EOD Exit", "Flip", etc.
    qty: float
    pnl: float = 0.0

    def __post_init__(self):
        self.pnl = (self.exit_price - self.entry_price) * self.qty


@dataclass
class Position:
    """Represents an open long position."""
    entry_time: pd.Timestamp
    entry_price: float
    qty: float
    entry_type: str
    trail_stop: float = 0.0


@dataclass
class StrategyState:
    """Mutable state carried across bars — mirrors Pine Script `var` variables."""

    position: Optional[Position] = None
    x_long_levels: list = field(default_factory=list)           # Type A X levels
    x_delayed_long_levels: list = field(default_factory=list)   # Type B X levels
    current_x: Optional[float] = None
    current_tc: Optional[float] = None
    current_bc: Optional[float] = None
    current_day_pivot: Optional[float] = None
    day_open_price: Optional[float] = None
    long_entry_triggered_today: bool = False
    trades: list = field(default_factory=list)
    equity: float = 100_000.0
    current_date: Optional[object] = None   # track date changes
    alert_levels: list = field(default_factory=list)
    alert_times: list = field(default_factory=list)


def calculate_cpr(day2_high: float, day2_low: float, day2_close: float):
    """
    Central Pivot Range from Day-2 data.

    X  (Pivot) = (H + L + C) / 3
    BC (Bottom Central) = (H + L) / 2
    TC (Top Central) = 2*X - BC

    Returns: (X, TC, BC)
    """
    x = (day2_high + day2_low + day2_close) / 3.0
    bc = (day2_high + day2_low) / 2.0
    tc = 2.0 * x - bc
    return x, tc, bc


def check_entry_conditions(
    candle_open: float,
    candle_low: float,
    candle_close: float,
    x_level: float,
    day_open_price: float,
) -> tuple:
    """
    Mirrors the Pine Script `check_entry_conditions()` function.

    Two cases:
      1. Standard Breach: Day opened at/above X → candle low touches X, candle open >= X
      2. Recovery Entry:  Day opened below X → candle opens below X but closes above X

    Returns (can_enter: bool, entry_type: str)
    """
    if pd.isna(day_open_price) or pd.isna(x_level):
        return False, ""

    # CASE 1: Day opens at or above X
    if day_open_price >= x_level:
        if candle_low <= x_level and candle_open >= x_level:
            return True, "Standard Breach"

    # CASE 2: Day opens below X
    else:
        if candle_open < x_level and candle_close > x_level:
            return True, "Recovery Entry"

    return False, ""


def run_backtest(df: pd.DataFrame, config: StrategyConfig = None) -> StrategyState:
    """
    Main backtest loop — iterates bar-by-bar through intraday data.

    Parameters
    ----------
    df : pd.DataFrame
        Intraday OHLCV data with DatetimeIndex and columns:
        open, high, low, close, volume (volume optional).
        Must also have daily context columns added by `add_daily_context()`.
    config : StrategyConfig
        Strategy parameters.

    Returns
    -------
    StrategyState with completed trades list and final equity.
    """
    if config is None:
        config = StrategyConfig()

    state = StrategyState(equity=config.initial_capital)
    atr_series = compute_atr(df, config.atr_length)

    for idx in range(len(df)):
        row = df.iloc[idx]
        bar_time = df.index[idx]
        bar_date = bar_time.date() if hasattr(bar_time, 'date') else bar_time

        o, h, l, c = row["open"], row["high"], row["low"], row["close"]
        atr_val = atr_series.iloc[idx]

        # --- Day-level columns (pre-computed) ---
        d2h = row.get("day2_high", np.nan)
        d2l = row.get("day2_low", np.nan)
        d2c = row.get("day2_close", np.nan)
        d1h = row.get("day1_high", np.nan)
        d1l = row.get("day1_low", np.nan)
        day_open = row.get("day_open_price", np.nan)

        # --- Detect new day ---
        is_new_day = (state.current_date is None) or (bar_date != state.current_date)

        if is_new_day:
            state.current_date = bar_date
            state.day_open_price = o
            state.long_entry_triggered_today = False

            # ─── STEP 1: Calculate CPR from Day-2 ───
            if not (pd.isna(d2h) or pd.isna(d2l) or pd.isna(d2c)):
                x, tc, bc = calculate_cpr(d2h, d2l, d2c)
                state.current_x = x
                state.current_tc = tc
                state.current_bc = bc
                state.current_day_pivot = x

            # ─── STEP 2A: Validate Type A (Standard Long X) ───
            # Day-1 Low >= TC AND Day-1 Low >= X AND Day-1 Low >= BC
            if (state.current_x is not None and
                state.current_tc is not None and
                state.current_bc is not None and
                not pd.isna(d1l)):

                if (d1l >= state.current_tc and
                    d1l >= state.current_x and
                    d1l >= state.current_bc):
                    state.x_long_levels.append(state.current_x)

            # ─── STEP 2B: Validate Type B (Delayed Long X) ───
            # Day-1 High <= TC AND Day-1 High <= X AND Day-1 High <= BC
            if (config.enable_delayed_x and
                state.current_x is not None and
                state.current_tc is not None and
                state.current_bc is not None and
                not pd.isna(d1h)):

                if (d1h <= state.current_tc and
                    d1h <= state.current_x and
                    d1h <= state.current_bc):
                    state.x_delayed_long_levels.append(state.current_x)

        # --- Time filters ---
        bar_hour = bar_time.hour if hasattr(bar_time, 'hour') else 0
        bar_minute = bar_time.minute if hasattr(bar_time, 'minute') else 0
        is_entry_allowed = bar_hour < 15
        is_day_close = (bar_hour == 15 and bar_minute >= 30)

        effective_day_open = state.day_open_price if state.day_open_price is not None else day_open

        # ─── Helper: Execute Entry ───
        def execute_entry(x_level: float, entry_type_label: str):
            """Open a new long position at the current bar's close."""
            if config.use_fixed_qty:
                qty = config.fixed_qty
            else:
                qty = state.equity / c if c > 0 else 0

            pos = Position(
                entry_time=bar_time,
                entry_price=c,       # Pine Script strategy.entry fills at close of signal bar
                qty=qty,
                entry_type=entry_type_label,
            )
            if config.use_atr_trail and not pd.isna(atr_val):
                pos.trail_stop = c - (atr_val * config.atr_multiplier)
            state.position = pos

        # ─── Helper: Close Position ───
        def close_position(exit_reason: str):
            """Record a completed trade and clear position."""
            if state.position is None:
                return
            trade = TradeRecord(
                entry_time=state.position.entry_time,
                entry_price=state.position.entry_price,
                exit_time=bar_time,
                exit_price=c,
                entry_type=state.position.entry_type,
                exit_reason=exit_reason,
                qty=state.position.qty,
            )
            state.trades.append(trade)
            state.equity += trade.pnl
            state.position = None

        # ═══════════════════════════════════════════════════
        # STEP 3A: TYPE A LONG ENTRY (no existing position)
        # ═══════════════════════════════════════════════════
        if (len(state.x_long_levels) > 0 and
            state.position is None and
            not state.long_entry_triggered_today and
            is_entry_allowed):

            for i in range(len(state.x_long_levels) - 1, -1, -1):
                x_level = state.x_long_levels[i]
                can_enter, etype = check_entry_conditions(o, l, c, x_level, effective_day_open)

                if can_enter:
                    state.x_long_levels.pop(i)

                    # Alert bookkeeping
                    state.alert_levels.insert(0, x_level)
                    state.alert_times.insert(0, str(bar_time))
                    if len(state.alert_levels) > config.max_alerts:
                        state.alert_levels.pop()
                        state.alert_times.pop()

                    execute_entry(x_level, f"Type A {etype}")
                    state.long_entry_triggered_today = True
                    break

        # ═══════════════════════════════════════════════════
        # STEP 3B: TYPE B DELAYED LONG ENTRY
        # ═══════════════════════════════════════════════════
        if (len(state.x_delayed_long_levels) > 0 and
            config.enable_delayed_x and
            state.position is None and
            not state.long_entry_triggered_today and
            is_entry_allowed):

            for i in range(len(state.x_delayed_long_levels) - 1, -1, -1):
                x_level = state.x_delayed_long_levels[i]
                can_enter, etype = check_entry_conditions(o, l, c, x_level, effective_day_open)

                if can_enter:
                    state.x_delayed_long_levels.pop(i)

                    state.alert_levels.insert(0, x_level)
                    state.alert_times.insert(0, str(bar_time))
                    if len(state.alert_levels) > config.max_alerts:
                        state.alert_levels.pop()
                        state.alert_times.pop()

                    execute_entry(x_level, f"Type B {etype}")
                    state.long_entry_triggered_today = True
                    break

        # ═══════════════════════════════════════════════════
        # STEP 4A: FLIP LONG — TYPE A X (already in position)
        # ═══════════════════════════════════════════════════
        if (len(state.x_long_levels) > 0 and
            state.position is not None and
            is_entry_allowed):

            for i in range(len(state.x_long_levels) - 1, -1, -1):
                x_level = state.x_long_levels[i]

                # Block flip if day_open < current_day_pivot
                if (effective_day_open is not None and
                    state.current_day_pivot is not None and
                    effective_day_open < state.current_day_pivot):
                    if l <= x_level:
                        continue  # flip blocked
                    continue

                can_enter, etype = check_entry_conditions(o, l, c, x_level, effective_day_open)

                if can_enter:
                    close_position(f"Flip to Type A X ({etype})")
                    execute_entry(x_level, f"Type A Flip {etype}")
                    state.x_long_levels.pop(i)
                    break

        # ═══════════════════════════════════════════════════
        # STEP 4B: FLIP LONG — TYPE B X (already in position)
        # ═══════════════════════════════════════════════════
        if (len(state.x_delayed_long_levels) > 0 and
            config.enable_delayed_x and
            state.position is not None and
            is_entry_allowed):

            for i in range(len(state.x_delayed_long_levels) - 1, -1, -1):
                x_level = state.x_delayed_long_levels[i]

                # Block flip if day_open < current_day_pivot
                if (effective_day_open is not None and
                    state.current_day_pivot is not None and
                    effective_day_open < state.current_day_pivot):
                    if l <= x_level:
                        pass  # flip blocked (Pine Script shows label here)
                    continue

                can_enter, etype = check_entry_conditions(o, l, c, x_level, effective_day_open)

                if can_enter:
                    close_position(f"Flip to Type B X ({etype})")
                    execute_entry(x_level, f"Type B Flip {etype}")
                    state.x_delayed_long_levels.pop(i)
                    break

        # ═══════════════════════════════════════════════════
        # ATR TRAILING STOP
        # ═══════════════════════════════════════════════════
        if state.position is not None and config.use_atr_trail and not pd.isna(atr_val):
            new_trail = c - (atr_val * config.atr_multiplier)
            state.position.trail_stop = max(state.position.trail_stop, new_trail)

            if c <= state.position.trail_stop:
                close_position("ATR Trail Stop")

        # ═══════════════════════════════════════════════════
        # END OF DAY EXIT
        # ═══════════════════════════════════════════════════
        if state.position is not None and is_day_close:
            close_position("EOD Exit")

    # Close any dangling position at end of data
    if state.position is not None:
        state.trades.append(TradeRecord(
            entry_time=state.position.entry_time,
            entry_price=state.position.entry_price,
            exit_time=df.index[-1],
            exit_price=df.iloc[-1]["close"],
            entry_type=state.position.entry_type,
            exit_reason="End of Data",
            qty=state.position.qty,
        ))
        state.equity += state.trades[-1].pnl
        state.position = None

    return state


# ═══════════════════════════════════════════════════════════════
# REPORTING
# ═══════════════════════════════════════════════════════════════

def generate_report(state: StrategyState, config: StrategyConfig) -> pd.DataFrame:
    """
    Build a trade log DataFrame and print summary statistics.
    """
    if not state.trades:
        print("No trades executed.")
        return pd.DataFrame()

    records = []
    for t in state.trades:
        records.append({
            "entry_time": t.entry_time,
            "entry_price": round(t.entry_price, 2),
            "exit_time": t.exit_time,
            "exit_price": round(t.exit_price, 2),
            "qty": round(t.qty, 4),
            "pnl": round(t.pnl, 2),
            "entry_type": t.entry_type,
            "exit_reason": t.exit_reason,
        })

    trade_df = pd.DataFrame(records)

    total_pnl = trade_df["pnl"].sum()
    num_trades = len(trade_df)
    winners = trade_df[trade_df["pnl"] > 0]
    losers = trade_df[trade_df["pnl"] <= 0]
    win_rate = len(winners) / num_trades * 100 if num_trades > 0 else 0
    avg_win = winners["pnl"].mean() if len(winners) > 0 else 0
    avg_loss = losers["pnl"].mean() if len(losers) > 0 else 0
    profit_factor = (
        abs(winners["pnl"].sum() / losers["pnl"].sum())
        if len(losers) > 0 and losers["pnl"].sum() != 0
        else float("inf")
    )

    # Max drawdown
    equity_curve = config.initial_capital + trade_df["pnl"].cumsum()
    running_max = equity_curve.cummax()
    drawdown = equity_curve - running_max
    max_dd = drawdown.min()

    print("=" * 60)
    print("  X STRATEGY CPR VALIDATION v4 — BACKTEST RESULTS")
    print("=" * 60)
    print(f"  Initial Capital:   ${config.initial_capital:,.2f}")
    print(f"  Final Equity:      ${state.equity:,.2f}")
    print(f"  Total PnL:         ${total_pnl:,.2f}")
    print(f"  Total Trades:      {num_trades}")
    print(f"  Winners:           {len(winners)}  ({win_rate:.1f}%)")
    print(f"  Losers:            {len(losers)}")
    print(f"  Avg Win:           ${avg_win:,.2f}")
    print(f"  Avg Loss:          ${avg_loss:,.2f}")
    print(f"  Profit Factor:     {profit_factor:.2f}")
    print(f"  Max Drawdown:      ${max_dd:,.2f}")
    print("=" * 60)

    # Breakdown by entry type
    print("\n  Breakdown by Entry Type:")
    for etype, group in trade_df.groupby("entry_type"):
        print(f"    {etype}: {len(group)} trades, "
              f"PnL=${group['pnl'].sum():,.2f}, "
              f"Win%={len(group[group['pnl'] > 0]) / len(group) * 100:.1f}%")

    print("\n  Breakdown by Exit Reason:")
    for reason, group in trade_df.groupby("exit_reason"):
        print(f"    {reason}: {len(group)} trades, "
              f"PnL=${group['pnl'].sum():,.2f}")

    return trade_df


# ═══════════════════════════════════════════════════════════════
# EXAMPLE USAGE / DEMO
# ═══════════════════════════════════════════════════════════════

def create_sample_data() -> pd.DataFrame:
    """
    Generate synthetic intraday data for demonstration.
    Replace this with your own data loading (CSV, API, etc.).

    Expected format:
      - DatetimeIndex (timezone-aware or naive)
      - Columns: open, high, low, close  (volume optional)
      - Intraday frequency (e.g., 5min, 15min)
    """
    np.random.seed(42)
    dates = pd.bdate_range("2024-01-02", "2024-03-29", freq="B")
    bars = []

    price = 100.0
    for day in dates:
        # Generate ~78 five-minute bars per trading day (9:30 to 16:00)
        times = pd.date_range(
            start=day.replace(hour=9, minute=30),
            end=day.replace(hour=15, minute=55),
            freq="5min"
        )
        for t in times:
            change = np.random.normal(0, 0.3)
            o = price
            h = o + abs(np.random.normal(0, 0.4))
            l = o - abs(np.random.normal(0, 0.4))
            c = o + change
            h = max(h, o, c)
            l = min(l, o, c)
            bars.append({"datetime": t, "open": o, "high": h, "low": l, "close": c})
            price = c

    df = pd.DataFrame(bars).set_index("datetime")
    return df


def load_csv_data(filepath: str, datetime_col: str = "datetime") -> pd.DataFrame:
    """
    Load intraday OHLC data from a CSV file.

    Expected columns: datetime (or date+time), open, high, low, close
    """
    df = pd.read_csv(filepath, parse_dates=[datetime_col], index_col=datetime_col)
    df.columns = df.columns.str.lower().str.strip()
    required = {"open", "high", "low", "close"}
    if not required.issubset(set(df.columns)):
        raise ValueError(f"CSV must contain columns: {required}. Found: {set(df.columns)}")
    df.sort_index(inplace=True)
    return df


# ═══════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # --- Configuration ---
    config = StrategyConfig(
        use_atr_trail=True,
        atr_length=14,
        atr_multiplier=3.0,
        enable_delayed_x=True,
        initial_capital=100_000.0,
        use_fixed_qty=False,
    )

    # --- Load Data ---
    # Option 1: Use synthetic demo data
    print("Loading sample data...")
    df = create_sample_data()

    # Option 2: Load from CSV (uncomment and adjust path)
    # df = load_csv_data("your_intraday_data.csv")

    print(f"Data: {len(df)} bars from {df.index[0]} to {df.index[-1]}")

    # --- Prepare daily context ---
    daily = prepare_daily_ohlc(df)
    df = add_daily_context(df, daily)

    # --- Run Backtest ---
    print("Running backtest...")
    state = run_backtest(df, config)

    # --- Report ---
    trade_df = generate_report(state, config)

    if len(trade_df) > 0:
        print("\n  Last 10 Trades:")
        print(trade_df.tail(10).to_string(index=False))
