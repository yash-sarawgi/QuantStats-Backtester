# ⚡ QuantStats Backtest Generator
### Professional Python GUI for Strategy Backtesting

---

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install quantstats pandas numpy alpaca-py pytz
```
*(The app also auto-installs missing packages on first launch.)*

### 2. Run the App
```bash
python backtest_app.py
```

---

## 📋 Strategy Code Format

Your strategy must define a **`strategy(data)`** function.

### Input: `data` — `pd.DataFrame`
| Column      | Type    | Description                            |
|-------------|---------|----------------------------------------|
| `open`      | float   | Bar open price                         |
| `high`      | float   | Bar high price                         |
| `low`       | float   | Bar low price                          |
| `close`     | float   | Bar close price                        |
| `volume`    | float   | Bar volume                             |
| `vwap`      | float   | Volume-weighted average price          |
| *(index)*   | DatetimeTZAware (UTC) | Timestamp of each bar    |

### Output: `pd.Series` of Signals
| Value | Meaning          |
|-------|------------------|
| `+1`  | Long position    |
|  `0`  | Flat / No trade  |
| `-1`  | Short position   |

> **Note:** The engine automatically shifts signals by 1 bar to avoid look-ahead bias (signal generated at bar N fills at bar N+1 open).

---

## 📝 Strategy Examples

### Example 1 — Golden Cross (SMA)
```python
import pandas as pd
import numpy as np

def strategy(data):
    close = data["close"]
    fast = close.rolling(20).mean()
    slow = close.rolling(50).mean()
    signal = pd.Series(0, index=data.index, dtype=float)
    signal[fast > slow] = 1
    return signal
```

### Example 2 — RSI Mean Reversion
```python
import pandas as pd
import numpy as np

def strategy(data):
    close = data["close"]
    delta = close.diff()
    gain = delta.clip(lower=0).rolling(14).mean()
    loss = (-delta.clip(upper=0)).rolling(14).mean()
    rs = gain / loss.replace(0, 1e-9)
    rsi = 100 - (100 / (1 + rs))

    signal = pd.Series(0, index=data.index, dtype=float)
    signal[rsi < 30] = 1    # oversold → long
    signal[rsi > 70] = 0    # overbought → exit
    return signal
```

### Example 3 — Bollinger Band Breakout
```python
import pandas as pd
import numpy as np

def strategy(data):
    close = data["close"]
    mid   = close.rolling(20).mean()
    std   = close.rolling(20).std()
    upper = mid + 2 * std
    lower = mid - 2 * std

    signal = pd.Series(0, index=data.index, dtype=float)
    signal[close > upper] =  1   # breakout long
    signal[close < lower] = -1   # breakout short
    return signal
```

### Example 4 — VWAP Reversion
```python
import pandas as pd

def strategy(data):
    signal = pd.Series(0, index=data.index, dtype=float)
    signal[data["close"] < data["vwap"]] = 1   # below VWAP → long
    signal[data["close"] > data["vwap"]] = 0   # above VWAP → exit
    return signal
```

---

## 📁 CSV File Format

The loader is **robust** and auto-detects columns. Supported column name aliases:

| Required | Accepted Names                                          |
|----------|---------------------------------------------------------|
| Timestamp| `timestamp`, `datetime`, `date`, `time`, `bar_time`     |
| Open     | `open`, `o`, `open_price`                               |
| High     | `high`, `h`, `high_price`                               |
| Low      | `low`, `l`, `low_price`                                 |
| Close    | `close`, `c`, `last`, `adj_close`                       |
| Volume   | `volume`, `vol`, `v`, `qty`    *(optional)*             |
| VWAP     | `vwap`, `vw`                   *(optional)*             |

### Sample CSV Layout (Alpaca Export Format)
```
timestamp,symbol,open,high,low,close,volume,trade_count,vwap
2020-11-30 09:41:00-05:00,SOFI,11,11,11,11,100,1,11
2020-11-30 09:47:00-05:00,SOFI,11.5,11.5,11.5,11.5,200,1,11.5
2020-11-30 09:48:00-05:00,SOFI,11.836,11.836,11.6,11.6,1000,2,11.718
```

### Timezone Handling
- Timestamps with UTC offset (e.g. `-05:00`) are parsed automatically
- Naive timestamps are assumed UTC
- All data is normalized to UTC internally

---

## 🔑 Alpaca API Setup

1. Go to **[alpaca.markets](https://alpaca.markets)** and create a free account
2. Navigate to **Paper Trading** → **API Keys**
3. Generate a new key pair (API Key + Secret Key)
4. Paste both into the app's API Key fields

> The free tier provides historical data access for all US equities.

---

## 📊 Output: QuantStats Tearsheet

The full HTML tearsheet includes:
- Equity curve & drawdown chart
- Monthly returns heatmap
- Rolling Sharpe/Sortino/Volatility
- Distribution of returns
- vs Benchmark (Buy & Hold) comparison
- Full statistics table (CAGR, Sharpe, Calmar, Win Rate, VaR, CVaR, etc.)

The tearsheet opens automatically in your browser after each backtest.

---

## ⚙️ Requirements

```
Python >= 3.8
quantstats >= 0.0.62
pandas >= 1.3
numpy >= 1.20
alpaca-py >= 0.8
pytz
tkinter (standard library)
```
