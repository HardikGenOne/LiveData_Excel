import csv
from datetime import datetime
from collections import deque
from statistics import mean
from SmartApi.smartWebSocketV2 import SmartWebSocketV2
from SmartApi.smartConnect import SmartConnect
import pyotp
from logzero import logger
import time
import pandas as pd

# ====== Credentials ======
TOTP_SECRET = "J4DWDXYMDAKVV6VFJW6RHMS3RI"
API_KEY = "gsS3VOae"
CLIENT_CODE = "L52128673"
PASS = "7990"
totp = pyotp.TOTP(TOTP_SECRET)

# ====== Tokens ======
FUT_TOKEN = "53213"
CE_TOKEN = "57845"
PE_TOKEN = "57846"

TOKEN_LIST = [{"exchangeType": 2, "tokens": [FUT_TOKEN, CE_TOKEN, PE_TOKEN]}]

# ====== CSV Setup ======
RESULT_PATH = "BANKNIFTY/result_fastbreakout.csv"
result_file = open(RESULT_PATH, mode='a', newline='')
result_writer = csv.writer(result_file)
result_writer.writerow(["OptionType", "Entry Time", "Entry Price", "Exit Time", "Exit Price", "PnL"])

# ====== Strategy Params ======
TARGET_TICKS = 5.0   # higher profit to catch breakout
STOPLOSS_TICKS = 3.0
COOLDOWN_SECONDS = 5

fut_ltp_history = deque(maxlen=3)
ce_ltp_history = deque(maxlen=3)
pe_ltp_history = deque(maxlen=3)

position = None
ltp_map = {}
last_exit_time = 0

def calculate_atr(prices, period=14):
    if len(prices) < period + 1:
        return None
    df = pd.DataFrame(prices, columns=["price"])
    df["previous"] = df["price"].shift(1)
    df["tr"] = (df["price"] - df["previous"]).abs()
    atr = df["tr"].rolling(window=period).mean().iloc[-1]
    return round(atr, 2)

# ====== Helpers ======
def log_result(option_type, entry_time, entry_price, exit_time, exit_price):
    global last_exit_time
    pnl = round(exit_price - entry_price, 2) if option_type == "CE" else round(entry_price - exit_price, 2)
    result_writer.writerow([option_type, entry_time, entry_price, exit_time, exit_price, pnl])
    result_file.flush()
    print(f"[TRADE EXIT] {option_type} | Entry: {entry_price} at {entry_time} | Exit: {exit_price} at {exit_time} | PnL: {pnl}")
    last_exit_time = time.time()

# ====== Tick Handler ======
def handle_tick(message):

    global position, last_exit_time
    token = message.get("token")
    ltp = message.get("last_traded_price", 0) / 100
    exch_ts = message.get("exchange_timestamp", 0)
    timestamp = datetime.fromtimestamp(exch_ts / 1000).strftime("%Y-%m-%d %H:%M:%S.%f") if exch_ts else datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")

    ltp_map[token] = (ltp, timestamp)

    ce_atr_prices = deque(maxlen=20)
    pe_atr_prices = deque(maxlen=20)

    if token == CE_TOKEN:
        ce_ltp_history.append(ltp)
        ce_atr_prices.append(ltp)

    elif token == PE_TOKEN:
        pe_ltp_history.append(ltp)
        pe_atr_prices.append(ltp)

    elif token == FUT_TOKEN:
        fut_ltp_history.append(ltp)

        # Debugging all incoming live tick data
        bid_qty = sum(b["quantity"] for b in message.get("best_5_buy_data", []))
        ask_qty = sum(s["quantity"] for s in message.get("best_5_sell_data", []))
        pressure_ratio = bid_qty / max(ask_qty, 1)

        ce_now = ltp_map.get(CE_TOKEN, (0, ""))[0]
        pe_now = ltp_map.get(PE_TOKEN, (0, ""))[0]

        # print("──── Tick Update ────")
        # print("📈 FUTURES LTP:", ltp)
        # print("🟢 Bid Qty:", bid_qty, "| 🔴 Ask Qty:", ask_qty, "| ⚖️ Ratio:", round(pressure_ratio, 2))
        # print("📊 CE LTPs:", list(ce_ltp_history), "| Now:", ce_now)
        # print("📊 PE LTPs:", list(pe_ltp_history), "| Now:", pe_now)
        # print("⏩ Futures Trend:", fut_ltp_history)

        # === Simple Fast Entry/Exit Strategy with ATR ===
        if position is None and len(fut_ltp_history) >= 2 and len(ce_ltp_history) >= 1 and len(pe_ltp_history) >= 1:
            ce_now = ce_ltp_history[-1]
            pe_now = pe_ltp_history[-1]

            trending_up = fut_ltp_history[-2] < fut_ltp_history[-1]
            trending_down = fut_ltp_history[-2] > fut_ltp_history[-1]

            if trending_up:
                atr = calculate_atr(list(ce_ltp_history), period=5)
                if atr:
                    entry_price = ce_now
                    entry_time = ltp_map.get(CE_TOKEN, ("", ""))[1]
                    sl = round(entry_price - atr, 2)
                    tgt = round(entry_price + 2 * atr, 2)
                    position = {
                        "type": "CE",
                        "entry_price": entry_price,
                        "entry_time": entry_time,
                        "stop_loss": sl,
                        "target": tgt
                    }
                    print(f"[ENTRY] CE at {entry_price} | SL: {sl} | Target: {tgt} | {entry_time}")

            elif trending_down:
                atr = calculate_atr(list(pe_ltp_history), period=5)
                if atr:
                    entry_price = pe_now
                    entry_time = ltp_map.get(PE_TOKEN, ("", ""))[1]
                    sl = round(entry_price + atr, 2)
                    tgt = round(entry_price - 2 * atr, 2)
                    position = {
                        "type": "PE",
                        "entry_price": entry_price,
                        "entry_time": entry_time,
                        "stop_loss": sl,
                        "target": tgt
                    }
                    print(f"[ENTRY] PE at {entry_price} | SL: {sl} | Target: {tgt} | {entry_time}")

    if position:
        token_id = CE_TOKEN if position['type'] == "CE" else PE_TOKEN
        if token_id in ltp_map:
            cur_price, cur_time = ltp_map[token_id]
            sl = position["stop_loss"]
            tgt = position["target"]

            if (position["type"] == "CE" and (cur_price <= sl or cur_price >= tgt)) or \
               (position["type"] == "PE" and (cur_price >= sl or cur_price <= tgt)):
                log_result(position["type"], position["entry_time"], position["entry_price"], cur_time, cur_price)
                position = None

# ====== WebSocket Handlers ======
def on_data(wsapp, message):
    handle_tick(message)

def on_open(wsapp):
    logger.info("WebSocket opened")
    sws.subscribe("banknifty-hft-sub", mode=3, token_list=TOKEN_LIST)

def on_error(wsapp, error):
    logger.error(f"WebSocket error: {error}")

def on_close(wsapp):
    logger.info("WebSocket closed")
    result_file.close()

# ====== Main ======
def main():
    global sws
    otp = totp.now()
    obj = SmartConnect(api_key=API_KEY)
    session_data = obj.generateSession(CLIENT_CODE, PASS, otp)
    feed_token = session_data['data']['feedToken']

    sws = SmartWebSocketV2(TOTP_SECRET, API_KEY, CLIENT_CODE, feed_token)
    sws.on_open = on_open
    sws.on_data = on_data
    sws.on_error = on_error
    sws.on_close = on_close
    sws.connect()

if __name__ == "__main__":
    main()