from SmartApi.smartWebSocketV2 import SmartWebSocketV2
from SmartApi.smartConnect import SmartConnect
from logzero import logger
import pyotp
import time
from datetime import datetime, timedelta
import threading
import requests
import os ,csv
from collections import defaultdict
import pandas as pd

# ------------------ CONFIG ------------------
AUTH_TOKEN = "IV7PPFHDE4RAWYS7OOXQIBLKTI"
API_KEY = "gsS3VOae"
CLIENT_CODE = "AAAJ289396"
PASS = "5689"
totp_secret = "IV7PPFHDE4RAWYS7OOXQIBLKTI"

SYMBOL_MAP = {'10940': 'DIVISLAB-EQ', '11532': 'ULTRACEMCO-EQ', '1348': 'HEROMOTOCO-EQ', '1594': 'INFY-EQ', '17963': 'NESTLEIND-EQ', '11630': 'NTPC-EQ', '13538': 'TECHM-EQ', '1394': 'HINDUNILVR-EQ', '14977': 'POWERGRID-EQ', '11483': 'LT-EQ', '10604': 'BHARTIARTL-EQ', '1363': 'HINDALCO-EQ', '15083': 'ADANIPORTS-EQ', '16675': 'BAJAJFINSV-EQ', '1624': 'IOC-EQ', '1660': 'ITC-EQ', '11287': 'UPL-EQ', '10999': 'MARUTI-EQ', '11536': 'TCS-EQ', '11723': 'JSWSTEEL-EQ', '1232': 'GRASIM-EQ', '3506': 'TITAN-EQ', '1333': 'HDFCBANK-EQ', '1922': 'KOTAKBANK-EQ', '2885': 'RELIANCE-EQ', '3432': 'TATACONSUM-EQ', '467': 'HDFCLIFE-EQ', '20374': 'COALINDIA-EQ', '317': 'BAJFINANCE-EQ', '236': 'ASIANPAINT-EQ', '5258': 'INDUSINDBK-EQ', '5900': 'AXISBANK-EQ', '3103': 'SHREECEM-EQ', '3351': 'SUNPHARMA-EQ', '3045': 'SBIN-EQ', '526': 'BPCL-EQ', '3499': 'TATASTEEL-EQ', '4963': 'ICICIBANK-EQ', '2031': 'M&M-EQ', '21808': 'SBILIFE-EQ', '2475': 'ONGC-EQ', '3787': 'WIPRO-EQ', '547': 'BRITANNIA-EQ', '910': 'EICHERMOT-EQ', '881': 'DRREDDY-EQ', '694': 'CIPLA-EQ', '7229': 'HCLTECH-EQ'}
token_list = [{"exchangeType": 1, "tokens": list(SYMBOL_MAP.keys())}]
correlation_id = "momentum-scan"
mode = 3

bot_token = "8021458974:AAH_U6vWbr877Cv669Ig88MFqWVUWZtf5Mk"
chat_id = 5609562789
url = f"https://api.telegram.org/bot{bot_token}/sendMessage"

tick_data = []

# ------------------ TELEGRAM ------------------
def send_telegram_message(text):
    try:
        data = {"chat_id": chat_id, "text": text}
        requests.post(url, data=data)
    except:
        pass

send_telegram_message("✅ Momentum Bot connected.")

# ------------------ WEBSOCKET CALLBACKS ------------------
def on_data(wsapp, message):
    try:
        ltp = message['last_traded_price'] / 100
        bid_qty = message.get("total_buy_quantity", 0)
        ask_qty = message.get("total_sell_quantity", 0)
        full_symbol = SYMBOL_MAP.get(message["token"], "UNKNOWN")
        symbol = full_symbol.replace("-EQ", "")  # Clean filename

        tick = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f"),
            "ltp": ltp,
            "bid_qty": bid_qty,
            "ask_qty": ask_qty
        }

        tick_data.append({
            "ltp": ltp,
            "bid_qty": bid_qty,
            "ask_qty": ask_qty,
            "timestamp": datetime.now(),
            "symbol": full_symbol
        })

        print(f"{symbol} | LTP: ₹{ltp:.2f}, Bid: {bid_qty}, Ask: {ask_qty}")

        # Save tick to CSV
        os.makedirs("NIFTY50 JUN20", exist_ok=True)
        file_path = os.path.join("NIFTY50 JUN20", f"{symbol}.csv")

        file_exists = os.path.isfile(file_path)
        with open(file_path, mode='a', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=["timestamp", "ltp", "bid_qty", "ask_qty"])
            if not file_exists:
                writer.writeheader()
            writer.writerow(tick)

    except Exception as e:
        logger.error(f"Error parsing tick: {e}")


def on_open(wsapp):
    logger.info("Socket opened")
    sws.subscribe(correlation_id, mode, token_list)

def on_error(wsapp, error):
    logger.error(error)

def on_close(wsapp):
    logger.info("Socket closed")

# ------------------ MOMENTUM LOGIC ------------------
def resample_candles(ticks, timeframe_min):
    df = pd.DataFrame(ticks)
    df['timestamp'] = pd.to_datetime(df['timestamp'])
    df.set_index('timestamp', inplace=True)
    ohlc_dict = {
        'ltp': 'ohlc',
        'bid_qty': 'sum',
        'ask_qty': 'sum'
    }
    resampled = df.groupby('symbol').resample(f'{timeframe_min}min').agg(ohlc_dict)['ltp']
    resampled = resampled.reset_index()
    return resampled

def detect_momentum():
    while True:
        tick_copy = list(tick_data)
        if len(tick_copy) < 200:
            time.sleep(90)
            continue

        candle_3min = resample_candles(tick_copy, 3)
        candle_15min = resample_candles(tick_copy, 15)

        now = datetime.now()
        cutoff = now - timedelta(minutes=1.5)
        recent_ticks = [t for t in tick_data if t["timestamp"] >= cutoff]
        symbols = set(t["symbol"] for t in recent_ticks)

        for symbol in symbols:
            df3 = candle_3min[candle_3min['symbol'] == symbol].copy()
            df15 = candle_15min[candle_15min['symbol'] == symbol].copy()
            if len(df3) < 3 or len(df15) < 3:
                continue

            latest3 = df3.iloc[-1]
            latest15 = df15.iloc[-1]

            is_3min_bullish = latest3['close'] > latest3['open']
            is_15min_bullish = latest15['close'] > latest15['open']

            if not (is_3min_bullish and is_15min_bullish):
                continue

            sym_ticks = [t for t in recent_ticks if t["symbol"] == symbol]
            if len(sym_ticks) < 30:
                continue

            sym_ticks.sort(key=lambda x: x["timestamp"])
            prices = [t["ltp"] for t in sym_ticks]
            bid_qtys = [t["bid_qty"] for t in sym_ticks]
            ask_qtys = [t["ask_qty"] for t in sym_ticks]

            start_price = prices[0]
            end_price = prices[-1]
            price_change = end_price - start_price
            pct_change = (price_change / start_price) * 100

            total_bid = sum(bid_qtys)
            total_ask = sum(ask_qtys)
            ob_imbalance = (total_bid - total_ask) / max((total_bid + total_ask), 1)

            sma_20 = pd.Series(prices[-20:]).mean()
            sma_200 = pd.Series(prices[-200:]).mean()

            support = min(prices[-15:])  # last 15 ticks for support
            resistance = max(prices[-15:])  # last 15 ticks for resistance
            range_pct = (resistance - support) / support * 100

            # Skip if range was too wide (no consolidation)
            if range_pct > 0.3:
                continue

            is_breakout_up = (
                end_price > resistance and
                ob_imbalance > 0.2 and
                end_price > sma_20 and
                end_price > sma_200
            )
            is_breakout_down = (
                end_price < support and
                ob_imbalance < -0.2 and
                end_price < sma_20 and
                end_price < sma_200
            )

            if not (is_breakout_up or is_breakout_down):
                continue

            direction = "📈 BREAKOUT UP" if is_breakout_up else "📉 BREAKOUT DOWN"
            entry = end_price
            if is_breakout_up:
                target = entry + (entry - support) * 1.5
                stop_loss = support
            else:
                target = entry - (resistance - entry) * 1.5
                stop_loss = resistance

            msg = (f"🚨 {symbol} Dual TF Breakout\n"
                   f"{direction}\n"
                   f"Entry: ₹{entry:.2f}, Target: ₹{target:.2f}, SL: ₹{stop_loss:.2f}\n"
                   f"S&R: ₹{support:.2f} - ₹{resistance:.2f}\n"
                   f"SMA20: ₹{sma_20:.2f}, SMA200: ₹{sma_200:.2f}\n"
                   f"3min Bullish: {is_3min_bullish}, 15min Bullish: {is_15min_bullish}\n"
                   f"OB Imbalance: {ob_imbalance:.2f}")
            send_telegram_message(msg)

            # Log trade to CSV
            os.makedirs("logs", exist_ok=True)
            with open("logs/paper_trades.csv", mode="a", newline="") as f:
                writer = csv.writer(f)
                if f.tell() == 0:
                    writer.writerow(["timestamp", "symbol", "direction", "entry", "target", "stop_loss", "support", "resistance", "sma20", "sma200", "ob_imbalance"])
                writer.writerow([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    symbol,
                    direction,
                    f"{entry:.2f}",
                    f"{target:.2f}",
                    f"{stop_loss:.2f}",
                    f"{support:.2f}",
                    f"{resistance:.2f}",
                    f"{sma_20:.2f}",
                    f"{sma_200:.2f}",
                    f"{ob_imbalance:.2f}"
                ])

        time.sleep(90)  # Run every 1.5 minutes

# ------------------ START ENGINE ------------------
# Session auth
totp = pyotp.TOTP(totp_secret)
otp = totp.now()
obj = SmartConnect(api_key=API_KEY)
session_data = obj.generateSession(CLIENT_CODE, PASS, otp)
FEED_TOKEN = session_data['data']['feedToken']

# Websocket
sws = SmartWebSocketV2(AUTH_TOKEN, API_KEY, CLIENT_CODE, FEED_TOKEN)
sws.on_open = on_open
sws.on_data = on_data
sws.on_error = on_error
sws.on_close = on_close

# Start momentum detection in background
threading.Thread(target=detect_momentum, daemon=True).start()

# Start socket
sws.connect()
