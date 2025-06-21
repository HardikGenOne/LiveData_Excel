from SmartApi.smartWebSocketV2 import SmartWebSocketV2
from SmartApi.smartConnect import SmartConnect
from logzero import logger
import pyotp
import time
from datetime import datetime
import threading
import requests
import os
import csv
import pandas as pd
import xlwings as xw

# ------------------ CONFIG ------------------
AUTH_TOKEN = "IV7PPFHDE4RAWYS7OOXQIBLKTI"
API_KEY = "gsS3VOae"
CLIENT_CODE = "AAAJ289396"
PASS = "5689"
totp_secret = "IV7PPFHDE4RAWYS7OOXQIBLKTI"

EXCEL_FILE = "live_data.xlsx"
SHEET_NAME = "LiveData"

# === SYMBOLS ===
SYMBOL_MAP = {
    '10940': 'DIVISLAB-EQ', '11532': 'ULTRACEMCO-EQ', '1348': 'HEROMOTOCO-EQ',
    '1594': 'INFY-EQ', '17963': 'NESTLEIND-EQ', '11630': 'NTPC-EQ',
    '13538': 'TECHM-EQ', '1394': 'HINDUNILVR-EQ', '14977': 'POWERGRID-EQ',
    '11483': 'LT-EQ', '10604': 'BHARTIARTL-EQ', '1363': 'HINDALCO-EQ',
    '15083': 'ADANIPORTS-EQ', '16675': 'BAJAJFINSV-EQ', '1624': 'IOC-EQ',
    '1660': 'ITC-EQ', '11287': 'UPL-EQ', '10999': 'MARUTI-EQ', '11536': 'TCS-EQ',
    '11723': 'JSWSTEEL-EQ', '1232': 'GRASIM-EQ', '3506': 'TITAN-EQ',
    '1333': 'HDFCBANK-EQ', '1922': 'KOTAKBANK-EQ', '2885': 'RELIANCE-EQ',
    '3432': 'TATACONSUM-EQ', '467': 'HDFCLIFE-EQ', '20374': 'COALINDIA-EQ',
    '317': 'BAJFINANCE-EQ', '236': 'ASIANPAINT-EQ', '5258': 'INDUSINDBK-EQ',
    '5900': 'AXISBANK-EQ', '3103': 'SHREECEM-EQ', '3351': 'SUNPHARMA-EQ',
    '3045': 'SBIN-EQ', '526': 'BPCL-EQ', '3499': 'TATASTEEL-EQ',
    '4963': 'ICICIBANK-EQ', '2031': 'M&M-EQ', '21808': 'SBILIFE-EQ',
    '2475': 'ONGC-EQ', '3787': 'WIPRO-EQ', '547': 'BRITANNIA-EQ',
    '910': 'EICHERMOT-EQ', '881': 'DRREDDY-EQ', '694': 'CIPLA-EQ', '7229': 'HCLTECH-EQ'
}
token_list = [{"exchangeType": 1, "tokens": list(SYMBOL_MAP.keys())}]
correlation_id = "momentum-scan"
mode = 3

# ------------------ TELEGRAM ------------------
bot_token = "8021458974:AAH_U6vWbr877Cv669Ig88MFqWVUWZtf5Mk"
chat_id = 5609562789
url = f"https://api.telegram.org/bot{bot_token}/sendMessage"

def send_telegram_message(text):
    try:
        requests.post(url, data={"chat_id": chat_id, "text": text})
    except:
        pass

send_telegram_message("✅ Momentum Bot connected.")

# ------------------ EXCEL SETUP ------------------
# Create Excel file if needed
if not os.path.exists(EXCEL_FILE):
    wb = xw.Book()
    sheet = wb.sheets[0]
    sheet.name = SHEET_NAME
    sheet.range("A1:D1").value = ["Symbol", "LTP", "Bid Qty", "Ask Qty"]
    wb.save(EXCEL_FILE)
    wb.close()

# Open Excel workbook
wb = xw.Book(EXCEL_FILE)
sheet = wb.sheets[SHEET_NAME]
symbol_row_map = {}
current_row = 2

# ------------------ TICK DATA ------------------
tick_data = []

def on_data(wsapp, message):
    try:
        ltp = message['last_traded_price'] / 100
        bid_qty = message.get("total_buy_quantity", 0)
        ask_qty = message.get("total_sell_quantity", 0)
        full_symbol = SYMBOL_MAP.get(message["token"], "UNKNOWN")
        symbol = full_symbol.replace("-EQ", "")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")

        print(f"{symbol} | LTP: ₹{ltp:.2f}, Bid: {bid_qty}, Ask: {ask_qty}")

        # Save to CSV per symbol
        os.makedirs("NIFTY50 JUN20", exist_ok=True)
        file_path = os.path.join("NIFTY50 JUN20", f"{symbol}.csv")
        file_exists = os.path.isfile(file_path)
        with open(file_path, mode='a', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=["timestamp", "ltp", "bid_qty", "ask_qty"])
            if not file_exists:
                writer.writeheader()
            writer.writerow({
                "timestamp": timestamp,
                "ltp": ltp,
                "bid_qty": bid_qty,
                "ask_qty": ask_qty
            })

        # Excel live update
        global current_row
        if symbol not in symbol_row_map:
            symbol_row_map[symbol] = current_row
            sheet.range(f"A{current_row}").value = symbol
            current_row += 1

        row = symbol_row_map[symbol]
        sheet.range(f"B{row}").value = ltp
        sheet.range(f"C{row}").value = bid_qty
        sheet.range(f"D{row}").value = ask_qty

    except Exception as e:
        logger.error(f"Tick error: {e}")

def on_open(wsapp):
    logger.info("✅ WebSocket opened")
    sws.subscribe(correlation_id, mode, token_list)

def on_error(wsapp, error):
    logger.error(f"❌ WebSocket error: {error}")

def on_close(wsapp):
    logger.info("🔌 WebSocket closed")

# ------------------ START ENGINE ------------------
# Authenticate
totp = pyotp.TOTP(totp_secret)
otp = totp.now()
obj = SmartConnect(api_key=API_KEY)
session_data = obj.generateSession(CLIENT_CODE, PASS, otp)
FEED_TOKEN = session_data['data']['feedToken']

# Start WebSocket
sws = SmartWebSocketV2(AUTH_TOKEN, API_KEY, CLIENT_CODE, FEED_TOKEN)
sws.on_open = on_open
sws.on_data = on_data
sws.on_error = on_error
sws.on_close = on_close

# Optional dummy background thread
def dummy_detect_momentum():
    while True:
        time.sleep(60)

threading.Thread(target=dummy_detect_momentum, daemon=True).start()

# Connect
sws.connect()
