"""
strategy.py
One-&-Done Opening Range Breakout (ORB) strategy — per-symbol state machine.

States:
  IDLE          → Waiting for market open / first candle
  WATCHING      → First candle confirmed, monitoring for breakout
  ENTRY_PENDING → Breakout detected, limit entry order placed
  IN_TRADE      → Entry filled, managing TP/SL
  DONE          → Trade completed or conditions invalidated; no more trades today
"""

import logging
import threading
from dataclasses import dataclass
from enum import Enum, auto
from typing import Optional
from datetime import datetime, time as dtime
import pytz

from candle_manager import SymbolCandleManager
from order_manager import OrderManager, OrderSide, OrderStatus, Position

logger = logging.getLogger("Strategy")

MARKET_TZ = pytz.timezone("America/New_York")
MARKET_OPEN = dtime(9, 30)


def _now_et() -> datetime:
    return datetime.now(tz=MARKET_TZ)


class StrategyState(Enum):
    IDLE          = auto()
    WATCHING      = auto()
    ENTRY_PENDING = auto()
    IN_TRADE      = auto()
    DONE          = auto()


@dataclass
class StrategyConfig:
    quantity: int           = 10
    tp1_ratio: float        = 1.0     # Risk multiplier for TP1
    tp2_ratio: float        = 2.0     # Risk multiplier for TP2
    tp1_pct: float          = 0.50    # % of position closed at TP1
    tp2_pct: float          = 0.50    # % of position closed at TP2
    entry_buffer: float     = 0.05    # Added to breakout level for limit price
    sl_buffer: float        = 0.00    # Buffer inside SL for stop-limit trigger
    min_gap_pct: float      = 0.10    # Minimum gap % to qualify
    max_entry_time: dtime   = dtime(10, 30)
    eod_exit_time: dtime    = dtime(15, 58)
    stop_limit_buffer: float = 0.05   # Limit below stop for SL stop-limit order


class SymbolStrategy:
    """
    Full ORB strategy state machine for a single symbol.
    Call `on_tick(price, volume, timestamp)` on every market data update.
    """

    def __init__(
        self,
        symbol: str,
        candle_mgr: SymbolCandleManager,
        order_mgr: OrderManager,
        config: StrategyConfig,
        paper_mode: bool = True,
    ):
        self.symbol = symbol
        self.candle_mgr = candle_mgr
        self.order_mgr = order_mgr
        self.cfg = config
        self.paper_mode = paper_mode

        self._state = StrategyState.IDLE
        self._lock = threading.Lock()

        # Set during WATCHING
        self._direction: Optional[str] = None   # "LONG" or "SHORT"
        self._breakout_price: Optional[float] = None
        self._sl_price: Optional[float] = None
        self._tp1_price: Optional[float] = None
        self._tp2_price: Optional[float] = None

        # Orders
        self._entry_order = None
        self._tp1_order = None
        self._tp2_order = None
        self._sl_order = None

        # Position reference
        self._position: Optional[Position] = None

    # ─────────────────────────────────────────────
    # Main tick handler (called on every price update)
    # ─────────────────────────────────────────────

    def on_tick(self, price: float, cum_volume: int, timestamp: datetime):
        """Route tick to candle manager and run state machine."""
        with self._lock:
            # Feed the candle manager
            closed_candle = self.candle_mgr.on_tick(price, cum_volume, timestamp)

            now_time = timestamp.astimezone(MARKET_TZ).time()

            # EOD exit check — always runs regardless of state
            if now_time >= self.cfg.eod_exit_time and self._state == StrategyState.IN_TRADE:
                self._eod_exit(price)
                return

            state = self._state

            if state == StrategyState.IDLE:
                self._handle_idle(closed_candle, now_time)

            elif state == StrategyState.WATCHING:
                # Check entry time gate
                if now_time >= self.cfg.max_entry_time:
                    logger.info(f"{self.symbol}: Max entry time reached — no trade today.")
                    self._state = StrategyState.DONE
                    return
                self._handle_watching(price, now_time)

            elif state == StrategyState.ENTRY_PENDING:
                self._handle_entry_pending(price)

            elif state == StrategyState.IN_TRADE:
                self._handle_in_trade(price)

    # ─────────────────────────────────────────────
    # State handlers
    # ─────────────────────────────────────────────

    def _handle_idle(self, closed_candle, now_time: dtime):
        """Wait for 9:30 first candle to close (at ~9:31)."""
        if self.candle_mgr.first_candle is None:
            return

        fc = self.candle_mgr.first_candle
        prev_close = self.candle_mgr.prev_close

        if prev_close is None:
            logger.warning(f"{self.symbol}: No prev_close — cannot determine gap. Skipping.")
            self._state = StrategyState.DONE
            return

        gap_pct = self.candle_mgr.gap_pct

        # ── Validate gap ──────────────────────────────────────────
        gap_up = gap_pct is not None and gap_pct >= self.cfg.min_gap_pct
        gap_down = gap_pct is not None and gap_pct <= -self.cfg.min_gap_pct

        if not (gap_up or gap_down):
            logger.info(
                f"{self.symbol}: Insufficient gap ({gap_pct:.2f}% vs ±{self.cfg.min_gap_pct}%). "
                f"Skipping."
            )
            self._state = StrategyState.DONE
            return

        # ── Validate candle color ─────────────────────────────────
        if gap_up and not fc.is_green:
            logger.info(f"{self.symbol}: Gap-up but first candle is NOT green. Skipping.")
            self._state = StrategyState.DONE
            return

        if gap_down and not fc.is_red:
            logger.info(f"{self.symbol}: Gap-down but first candle is NOT red. Skipping.")
            self._state = StrategyState.DONE
            return

        # ── Set up breakout levels ────────────────────────────────
        if gap_up:
            self._direction = "LONG"
            self._breakout_price = round(fc.high + self.cfg.entry_buffer, 2)
            self._sl_price = round(fc.low, 2)
        else:
            self._direction = "SHORT"
            self._breakout_price = round(fc.low - self.cfg.entry_buffer, 2)
            self._sl_price = round(fc.high, 2)

        logger.info(
            f"{self.symbol}: Setup {self._direction} | "
            f"Gap={gap_pct:.2f}% | "
            f"FC H={fc.high:.2f} L={fc.low:.2f} | "
            f"Breakout={self._breakout_price:.2f} | SL={self._sl_price:.2f}"
        )
        self._state = StrategyState.WATCHING

    def _handle_watching(self, price: float, now_time: dtime):
        """Monitor live price for breakout of the opening range."""
        ema = self.candle_mgr.ema_value
        vwap = self.candle_mgr.vwap_value
        fc = self.candle_mgr.first_candle
        if fc is None:
            return

        if ema is None or vwap is None:
            logger.debug(f"{self.symbol}: EMA/VWAP not ready yet.")
            return

        current_low = self.candle_mgr.current_candle_low
        current_high = self.candle_mgr.current_candle_high

        if self._direction == "LONG":
            # Price must break above first candle high
            if price >= fc.high:
                # Entry candle low must be above EMA and VWAP
                if current_low is not None and current_low > ema and current_low > vwap:
                    self._fire_long_entry()
                else:
                    logger.info(
                        f"{self.symbol}: Long breakout but EMA/VWAP filter failed. "
                        f"CandleLow={current_low:.2f} EMA={ema:.2f} VWAP={vwap:.2f}"
                    )
                    self._state = StrategyState.DONE

        elif self._direction == "SHORT":
            # Price must break below first candle low
            if price <= fc.low:
                # Entry candle high must be below EMA and VWAP
                if current_high is not None and current_high < ema and current_high < vwap:
                    self._fire_short_entry()
                else:
                    logger.info(
                        f"{self.symbol}: Short breakout but EMA/VWAP filter failed. "
                        f"CandleHigh={current_high:.2f} EMA={ema:.2f} VWAP={vwap:.2f}"
                    )
                    self._state = StrategyState.DONE

    def _handle_entry_pending(self, price: float):
        """Check if the entry order has been filled."""
        if self._entry_order is None:
            return

        order = self._entry_order

        if order.status == OrderStatus.FILLED:
            entry_price = order.avg_fill_price if order.avg_fill_price > 0 else order.price
            self._on_entry_filled(entry_price)

        elif order.status in (OrderStatus.CANCELLED, OrderStatus.REJECTED):
            logger.warning(f"{self.symbol}: Entry order {order.status.name}. No trade.")
            self._state = StrategyState.DONE

        elif order.status == OrderStatus.PARTIAL:
            # For this strategy we wait for full fill (10 shares)
            logger.debug(f"{self.symbol}: Entry partial fill {order.filled_qty}/{order.qty}")

    def _handle_in_trade(self, price: float):
        """Monitor TP/SL fill status. DAS handles the actual order execution."""
        pos = self._position
        if pos is None:
            return

        # Check TP1 fill
        if not pos.tp1_hit and self._tp1_order and self._tp1_order.status == OrderStatus.FILLED:
            pos.tp1_hit = True
            logger.info(
                f"{self.symbol}: TP1 HIT @ {self._tp1_order.avg_fill_price:.2f} "
                f"({pos.tp1_qty} shares)"
            )
            # Reduce SL order qty to remaining shares
            self._update_sl_for_remaining(pos)

        # Check TP2 fill
        if pos.tp1_hit and not pos.tp2_hit and self._tp2_order and self._tp2_order.status == OrderStatus.FILLED:
            pos.tp2_hit = True
            logger.info(
                f"{self.symbol}: TP2 HIT @ {self._tp2_order.avg_fill_price:.2f} "
                f"({pos.tp2_qty} shares)"
            )
            self._close_out()

        # Check SL fill
        if self._sl_order and self._sl_order.status == OrderStatus.FILLED:
            logger.info(
                f"{self.symbol}: STOP LOSS HIT @ {self._sl_order.avg_fill_price:.2f}"
            )
            self._close_out()

    # ─────────────────────────────────────────────
    # Order firing
    # ─────────────────────────────────────────────

    def _fire_long_entry(self):
        qty = self.cfg.quantity
        entry_price = self._breakout_price
        risk = entry_price - self._sl_price

        if risk <= 0:
            logger.error(f"{self.symbol}: Invalid risk for LONG (entry={entry_price} sl={self._sl_price}). Skipping.")
            self._state = StrategyState.DONE
            return

        self._tp1_price = round(entry_price + risk * self.cfg.tp1_ratio, 2)
        self._tp2_price = round(entry_price + risk * self.cfg.tp2_ratio, 2)

        logger.info(
            f"{self.symbol}: LONG ENTRY → price={entry_price:.2f} "
            f"SL={self._sl_price:.2f} TP1={self._tp1_price:.2f} TP2={self._tp2_price:.2f} "
            f"Risk/share={risk:.2f}"
        )

        self._entry_order = self.order_mgr.place_limit_order(
            self.symbol, OrderSide.BUY, qty, entry_price, tag="ENTRY"
        )
        self._state = StrategyState.ENTRY_PENDING

    def _fire_short_entry(self):
        qty = self.cfg.quantity
        entry_price = self._breakout_price
        risk = self._sl_price - entry_price

        if risk <= 0:
            logger.error(f"{self.symbol}: Invalid risk for SHORT. Skipping.")
            self._state = StrategyState.DONE
            return

        self._tp1_price = round(entry_price - risk * self.cfg.tp1_ratio, 2)
        self._tp2_price = round(entry_price - risk * self.cfg.tp2_ratio, 2)

        logger.info(
            f"{self.symbol}: SHORT ENTRY → price={entry_price:.2f} "
            f"SL={self._sl_price:.2f} TP1={self._tp1_price:.2f} TP2={self._tp2_price:.2f} "
            f"Risk/share={risk:.2f}"
        )

        self._entry_order = self.order_mgr.place_limit_order(
            self.symbol, OrderSide.SHORT, qty, entry_price, tag="ENTRY"
        )
        self._state = StrategyState.ENTRY_PENDING

    def _on_entry_filled(self, fill_price: float):
        """Called when entry limit order is confirmed filled."""
        qty = self.cfg.quantity
        tp1_qty = round(qty * self.cfg.tp1_pct)    # 5 shares
        tp2_qty = qty - tp1_qty                     # 5 shares

        logger.info(
            f"{self.symbol}: Entry FILLED @ {fill_price:.2f} | "
            f"TP1_qty={tp1_qty} TP2_qty={tp2_qty}"
        )

        pos = Position(
            symbol=self.symbol,
            side=self._direction,
            entry_price=fill_price,
            qty=qty,
            sl_price=self._sl_price,
            tp1_price=self._tp1_price,
            tp2_price=self._tp2_price,
            tp1_qty=tp1_qty,
            tp2_qty=tp2_qty,
        )
        self.order_mgr.register_position(pos)
        self._position = pos

        if self._direction == "LONG":
            # TP orders (sell limit)
            self._tp1_order = self.order_mgr.place_limit_order(
                self.symbol, OrderSide.SELL, tp1_qty, self._tp1_price, tag="TP1"
            )
            self._tp2_order = self.order_mgr.place_limit_order(
                self.symbol, OrderSide.SELL, tp2_qty, self._tp2_price, tag="TP2"
            )
            # SL stop-limit (sell)
            sl_limit = round(self._sl_price - self.cfg.stop_limit_buffer, 2)
            self._sl_order = self.order_mgr.place_stop_limit_order(
                self.symbol, OrderSide.SELL, qty, self._sl_price, sl_limit, tag="SL"
            )

        else:  # SHORT
            # TP orders (buy-to-cover limit)
            self._tp1_order = self.order_mgr.place_limit_order(
                self.symbol, OrderSide.COVER, tp1_qty, self._tp1_price, tag="TP1"
            )
            self._tp2_order = self.order_mgr.place_limit_order(
                self.symbol, OrderSide.COVER, tp2_qty, self._tp2_price, tag="TP2"
            )
            # SL stop-limit (cover)
            sl_limit = round(self._sl_price + self.cfg.stop_limit_buffer, 2)
            self._sl_order = self.order_mgr.place_stop_limit_order(
                self.symbol, OrderSide.COVER, qty, self._sl_price, sl_limit, tag="SL"
            )

        self._state = StrategyState.IN_TRADE

    def _update_sl_for_remaining(self, pos: Position):
        """After TP1 hit, cancel original SL and re-place for remaining qty only."""
        if self._sl_order and self._sl_order.status not in (
            OrderStatus.FILLED, OrderStatus.CANCELLED
        ):
            self.order_mgr.cancel_order(self._sl_order)

        remaining = pos.tp2_qty
        if self._direction == "LONG":
            sl_limit = round(self._sl_price - self.cfg.stop_limit_buffer, 2)
            self._sl_order = self.order_mgr.place_stop_limit_order(
                self.symbol, OrderSide.SELL, remaining,
                self._sl_price, sl_limit, tag="SL_PARTIAL"
            )
        else:
            sl_limit = round(self._sl_price + self.cfg.stop_limit_buffer, 2)
            self._sl_order = self.order_mgr.place_stop_limit_order(
                self.symbol, OrderSide.COVER, remaining,
                self._sl_price, sl_limit, tag="SL_PARTIAL"
            )

    def _eod_exit(self, price: float):
        """Market-close forced exit for all remaining position."""
        pos = self._position
        if pos is None:
            return

        logger.info(f"{self.symbol}: EOD exit triggered @ {price:.2f}")

        # Cancel outstanding TP/SL orders
        for order in [self._tp1_order, self._tp2_order, self._sl_order]:
            if order and order.status not in (
                OrderStatus.FILLED, OrderStatus.CANCELLED
            ):
                self.order_mgr.cancel_order(order)

        # Calculate remaining shares
        tp1_filled = pos.tp1_qty if pos.tp1_hit else 0
        remaining = pos.qty - tp1_filled

        if remaining > 0:
            eod_price = round(price, 2)
            if self._direction == "LONG":
                self.order_mgr.place_limit_order(
                    self.symbol, OrderSide.SELL, remaining, eod_price, tag="EOD_EXIT"
                )
            else:
                self.order_mgr.place_limit_order(
                    self.symbol, OrderSide.COVER, remaining, eod_price, tag="EOD_EXIT"
                )

        self._close_out()

    def _close_out(self):
        self.order_mgr.close_position(self.symbol)
        self._state = StrategyState.DONE
        logger.info(f"{self.symbol}: Strategy DONE for today.")

    # ─────────────────────────────────────────────
    # Public status
    # ─────────────────────────────────────────────

    @property
    def state(self) -> StrategyState:
        return self._state

    @property
    def direction(self) -> Optional[str]:
        return self._direction

    def status_summary(self) -> str:
        ema = self.candle_mgr.ema_value
        vwap = self.candle_mgr.vwap_value
        fc = self.candle_mgr.first_candle
        return (
            f"{self.symbol} | State={self._state.name} | Dir={self._direction} | "
            f"EMA={ema:.4f if ema else 'N/A'} | VWAP={vwap:.4f if vwap else 'N/A'} | "
            f"FC={'set' if fc else 'pending'}"
        )
