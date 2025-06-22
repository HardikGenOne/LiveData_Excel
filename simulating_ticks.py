import os
import time
import pandas as pd
import xlwings as xw

# === CONFIG ===
FOLDER_PATH = "NIFTY50_1min_20JUN"
EXCEL_FILE = "nifty50_live.xlsx"
STATUS_FILE = "nifty50_status.xlsx"
UPDATE_INTERVAL = 0.1  # seconds

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
    sht.range("A1").value = ["Symbol", "Open", "High", "Low", "Close"]
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
    status_sht.range("A2").value = ["Symbol", "Open", "High", "Low", "Close", "PercChange"]
    status_sht.range("A20").value = "BEARISH"
    status_sht.range("A21").value = ["Symbol", "Open", "High", "Low", "Close", "PercChange"]
    status_wb.save(STATUS_FILE)

first_open = {symbol: df.iloc[0]["Open"] for symbol, df in stock_data.items()}

# === LIVE UPDATE LOOP ===
for i in range(min_rows):
    
    for row_idx, symbol in enumerate(stock_names, start=2):
        df = stock_data[symbol]
        row = df.iloc[i]

        current_open = row["Open"]
        current_close = row["Close"]
        current_high = row["High"]
        current_low = row["Low"]
        volume= row["Volume"]
        row["Volume_SMA"] = 
        base_open = first_open[symbol]
        change_pct = ((current_close - base_open) / base_open) * 100
        volume
        row_values = [symbol, current_open, current_high, current_low, current_close, round(change_pct, 2)]

        # Clean existing entries
        for row_num in range(3, 20):
            if status_sht.range(f"A{row_num}").value == symbol:
                status_sht.range(f"A{row_num}:F{row_num}").value = None
        for row_num in range(22, 100):
            if status_sht.range(f"A{row_num}").value == symbol:
                status_sht.range(f"A{row_num}:F{row_num}").value = None

        # Add to bullish section
        if change_pct >= 0.2:
            for row_num in range(3, 20):
                if status_sht.range(f"A{row_num}").value is None:
                    status_sht.range(f"A{row_num}:F{row_num}").value = row_values
                    break

        # Add to bearish section
        elif change_pct <= -.2:
            for row_num in range(22, 100):
                if status_sht.range(f"A{row_num}").value is None:
                    status_sht.range(f"A{row_num}:F{row_num}").value = row_values
                    break

        sht.range(f"B{row_idx}").value = row["Open"]
        sht.range(f"C{row_idx}").value = row["High"]
        sht.range(f"D{row_idx}").value = row["Low"]
        sht.range(f"E{row_idx}").value = row["Close"]

    print(f"Updated minute {i+1}")
    time.sleep(UPDATE_INTERVAL)

# Optional: keep workbook open
# wb.save(EXCEL_FILE)
# wb.close()

status_wb.save(STATUS_FILE)