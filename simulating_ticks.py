import os
import time
import pandas as pd
import xlwings as xw
import numpy as np

# === CONFIG ===
FOLDER_PATH = "NIFTY50_1min_20JUN"
EXCEL_FILE = "nifty50_live.xlsx"
STATUS_FILE = "nifty50_status.xlsx"
UPDATE_INTERVAL = 0.1  # seconds

RSI_WINDOW = 14
VOLUME_SMA_WINDOW = 20
PERC_CHANGE_THRESHOLD = 0.5

SUPERTREND_PERIOD = 10
SUPERTREND_MULTIPLIER = 3

def calculate_supertrend(df, period=10, multiplier=3):
    hl2 = (df['High'] + df['Low']) / 2
    atr = df['High'].rolling(period).max() - df['Low'].rolling(period).min()
    atr = atr.rolling(period).mean()

    upper_band = hl2 + multiplier * atr
    lower_band = hl2 - multiplier * atr

    supertrend = [np.nan] * len(df)
    in_uptrend = True

    for i in range(period, len(df)):
        if df['Close'][i] > upper_band[i - 1]:
            in_uptrend = True
        elif df['Close'][i] < lower_band[i - 1]:
            in_uptrend = False

        if in_uptrend:
            supertrend[i] = lower_band[i]
        else:
            supertrend[i] = upper_band[i]

    df['Supertrend'] = supertrend
    df = df["Supertrend"].fillna(method="bfill")
    return df

# === READ STOCK DATA ===
csv_files = [f for f in os.listdir(FOLDER_PATH) if f.endswith(".csv")]
stock_names = [f.replace(".csv", "") for f in csv_files]
stock_data = {name: pd.read_csv(os.path.join(FOLDER_PATH, name + ".csv")) for name in stock_names}
min_rows = min(len(df) for df in stock_data.values())

# === INIT EXCEL WITH xlwings ===
app = xw.App(visible=True)  # make Excel visible
wb = xw.Book(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else xw.Book()
sht = wb.sheets[0]

# Write header if empty
if sht.range("A1").value != "Symbol":
    sht.range("A1").value = ["Symbol", "Open", "High", "Low", "Close", "VWAP", "RSI", "Supertrend"]
    for i, symbol in enumerate(stock_names):
        sht.range(f"A{i+2}").value = symbol

# === INIT STATUS FILE ===
if os.path.exists(STATUS_FILE):
    status_wb = xw.Book(STATUS_FILE)
    status_sht = status_wb.sheets[0]
else:
    status_wb = xw.Book()
    status_sht = status_wb.sheets[0]
    status_sht.range("A1").value = "BULLISH"
    status_sht.range("A2").value = ["Symbol", "Open", "High", "Low", "Close", "PercChange", "VWAP", "RSI"]
    status_sht.range("A20").value = "BEARISH"
    status_sht.range("A21").value = ["Symbol", "Open", "High", "Low", "Close", "PercChange", "VWAP", "RSI"]
    status_wb.save(STATUS_FILE)

first_open = {symbol: df.iloc[0]["Open"] for symbol, df in stock_data.items()}

volume_sma = VOLUME_SMA_WINDOW
for symbol in stock_names:
    df = stock_data[symbol]
    df[f"Volume_SMA_{volume_sma}"] = df["Volume"].rolling(window=volume_sma).mean()

from ta.momentum import RSIIndicator
from ta.trend import STCIndicator

for symbol in stock_names:
    df = stock_data[symbol]
    df["VWAP"] = (df["Close"] * df["Volume"]).cumsum() / df["Volume"].cumsum()
    df["RSI"] = RSIIndicator(close=df["Close"], window=RSI_WINDOW).rsi().fillna(method="bfill")
    df = calculate_supertrend(df, period=SUPERTREND_PERIOD, multiplier=SUPERTREND_MULTIPLIER)

# === LIVE UPDATE LOOP ===
for i in range(min_rows):
    
    for row_idx, symbol in enumerate(stock_names, start=2):
        df = stock_data[symbol]
        row = df.iloc[i]

        vwap = row["VWAP"]
        rsi = row["RSI"]
        supertrend = row["Supertrend"]

        current_open = row["Open"]
        current_close = row["Close"]
        current_high = row["High"]
        current_low = row["Low"]
        volume= row["Volume"]

        base_open = first_open[symbol]
        change_pct = ((current_close - base_open) / base_open) * 100

        vwap_text = ""
        rsi_value = ""
        if current_close > vwap and change_pct >= PERC_CHANGE_THRESHOLD and supertrend < current_close:
            vwap_text = "BUY"
            rsi_value = round(rsi, 1)
        elif current_close < vwap and change_pct <= -PERC_CHANGE_THRESHOLD and supertrend > current_close:
            vwap_text = "SELL"
            rsi_value = round(rsi, 1)

        row_values = [symbol, current_open, current_high, current_low, current_close, round(change_pct, 2), vwap_text, rsi_value]

        # Clean existing entries
        for row_num in range(3, 20):
            if status_sht.range(f"A{row_num}").value == symbol:
                status_sht.range(f"A{row_num}:H{row_num}").value = None
        for row_num in range(22, 100):
            if status_sht.range(f"A{row_num}").value == symbol:
                status_sht.range(f"A{row_num}:H{row_num}").value = None

        # Add to bullish section
        if change_pct >= PERC_CHANGE_THRESHOLD and current_close > vwap:
            for row_num in range(3, 20):
                if status_sht.range(f"A{row_num}").value is None:
                    status_sht.range(f"A{row_num}:H{row_num}").value = row_values
                    break

        # Add to bearish section
        elif change_pct <= -PERC_CHANGE_THRESHOLD and current_close < vwap:
            for row_num in range(22, 100):
                if status_sht.range(f"A{row_num}").value is None:
                    status_sht.range(f"A{row_num}:H{row_num}").value = row_values
                    break

        sht.range(f"B{row_idx}").value = row["Open"]
        sht.range(f"C{row_idx}").value = row["High"]
        sht.range(f"D{row_idx}").value = row["Low"]
        sht.range(f"E{row_idx}").value = row["Close"]
        sht.range(f"F{row_idx}").value = vwap
        sht.range(f"G{row_idx}").value = rsi
        sht.range(f"H{row_idx}").value = supertrend

    print(f"Updated minute {i+1}")
    time.sleep(UPDATE_INTERVAL)

# Optional: keep workbook open
# wb.save(EXCEL_FILE)
# wb.close()

status_wb.save(STATUS_FILE)