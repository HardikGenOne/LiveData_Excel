from SmartApi.smartWebSocketV2 import SmartWebSocketV2
from SmartApi.smartConnect import SmartConnect
from logzero import logger
import pyotp
from datetime import datetime
import xlwings as xw
import threading
import pandas as pd
import time
from datetime import datetime
import pandas_ta as ta
from ta.momentum import RSIIndicator
from ta.volume import VolumeWeightedAveragePrice

# ------------------ CONFIG ------------------
AUTH_TOKEN = "IV7PPFHDE4RAWYS7OOXQIBLKTI"
API_KEY = "gsS3VOae"
CLIENT_CODE = "AAAJ289396"
PASS = "5689"
totp_secret = "IV7PPFHDE4RAWYS7OOXQIBLKTI"

class Excel_Live_Ticks():
    def __init__(self, auth_token, api_key, client_code, password, totp_secret):
        self.AUTH_TOKEN = auth_token
        self.API_KEY = api_key
        self.CLIENT_CODE = client_code
        self.PASS = password
        self.totp_secret = totp_secret
        self.sws = None
        self.symbol_map = self.get_symbol_map()
        self.tick_buffer = []
        self.last_tick_time = time.time()

    def connect_to_AngleOne(self):
        totp = pyotp.TOTP(self.totp_secret)
        otp = totp.now()
        obj = SmartConnect(api_key=self.API_KEY)
        session_data = obj.generateSession(self.CLIENT_CODE, self.PASS, otp)
        FEED_TOKEN = session_data['data']['feedToken']

        self.sws = SmartWebSocketV2(self.AUTH_TOKEN, self.API_KEY, self.CLIENT_CODE, FEED_TOKEN)

        self.sws.on_error = self.on_error
        self.sws.on_close = self.on_close

    def start_NIFTY50_streaming(self):
        if not self.sws:
            logger.error("WebSocket not initialized. Call connect_to_AngleOne() first.")
            return

        self.sws.on_open = self.on_open
        self.sws.on_data = self.on_data

        self.sws.connect()
        self.monitor_connection()

    def on_open(self, wsapp):
        logger.info("✅ WebSocket opened")
        token_list = [{"exchangeType": 1, "tokens": list(self.symbol_map.keys())}]
        self.sws.subscribe("momentum-scan", 3, token_list)

    def on_data(self, wsapp, message):
        try:
            ltp = message['last_traded_price'] / 100
            best_5_buy = message.get("best_5_buy_data", [])
            best_5_sell = message.get("best_5_sell_data", [])

            bid_qty = sum(level["quantity"] for level in best_5_buy)
            ask_qty = sum(level["quantity"] for level in best_5_sell)

            full_symbol = self.symbol_map.get(message["token"], "UNKNOWN")
            symbol = full_symbol.replace("-EQ", "")
            print(f"🕒 Tick received at: {datetime.now().strftime('%H:%M:%S')} | {symbol}")
            print(f"{symbol} | LTP: ₹{ltp:.2f}, 5Bid: {bid_qty}, 5Ask: {ask_qty}")

            tick_ts = message.get("exchange_timestamp", 0)
            tick_time = datetime.fromtimestamp(tick_ts / 1000).strftime("%H:%M:%S")

            used_range = self.sheet.range("A1:A1000").value
            symbol_rows = {row_val: idx + 1 for idx, row_val in enumerate(used_range) if row_val}

            if symbol in symbol_rows:
                row_num = symbol_rows[symbol]
            else:
                row_num = len(symbol_rows) + 2
                self.sheet.range(f"A{row_num}").value = symbol

            ltq = message.get("last_traded_quantity", "")
            atp = message.get("average_traded_price", "") / 100
            volume = message.get("volume_trade_for_the_day", "")
            open_ = message.get("open_price_of_the_day", "") / 100
            high = message.get("high_price_of_the_day", "") / 100
            low = message.get("low_price_of_the_day", "") / 100
            prev_close = message.get("closed_price", "") / 100
            oi = message.get("open_interest", "")
            upper_circuit = message.get("upper_circuit_limit", "") / 100
            lower_circuit = message.get("lower_circuit_limit", "") / 100
            week_52_high = message.get("52_week_high_price", "") / 100
            week_52_low = message.get("52_week_low_price", "") / 100

            self.sheet.range(f"B{row_num}").value = [
                ltp, bid_qty, ask_qty, tick_time,
                ltq, atp, volume, open_, high, low, prev_close,
                oi, upper_circuit, lower_circuit, week_52_high, week_52_low
            ]
            self.tick_buffer.append({
                "symbol": symbol,
                "ltp": ltp,
                "bid_qty": bid_qty,
                "ask_qty": ask_qty,
                "volume": volume,
                "atp": atp,
                "tick_time": tick_time,
                "open": open_,
                "high": high,
                "low": low
            })
            self.last_tick_time = time.time()
        except Exception as e:
            logger.error(f"Tick error: {e}")

    def on_error(self, wsapp, error):
        logger.error(f"❌ WebSocket error: {error}")

    def on_close(self, wsapp):
        logger.info("🔌 WebSocket closed")

    def get_symbol_map(self):
        return {
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
        # return {'11483': 'LT-EQ', '10453': 'SATIN-EQ', '10666': 'PNB-EQ', '10794': 'CANBK-EQ', '11373': 'BIOCON-EQ', '11536': 'TCS-EQ', '15141': 'COLPAL-EQ', '15355': 'RECLTD-EQ', '1023': 'FEDERALBNK-EQ', '10738': 'OFSS-EQ', '11184': 'IDFCFIRSTB-EQ', '15313': 'IRB-EQ', '15337': 'RAIN-EQ', '10440': 'LUPIN-EQ', '11287': 'UPL-EQ', '13285': 'M&MFIN-EQ', '11630': 'NTPC-EQ', '13404': 'SUNTV-EQ', '11654': 'LALPATHLAB-EQ', '13086': 'AIAENG-EQ', '17903': 'ABBOTINDIA-EQ', '1922': 'KOTAKBANK-EQ', '20302': 'PRESTIGE-EQ', '21690': 'DIXON-EQ', '17875': 'GODREJPROP-EQ', '16713': 'UBL-EQ', '11351': 'PETRONET-EQ', '17534': 'MGL-EQ', '11543': 'COFORGE-EQ', '19061': 'MANAPPURAM-EQ', '2043': 'RAMCOCEM-EQ', '212': 'ASHOKLEY-EQ', '2263': 'BANDHANBNK-EQ', '10925': 'GODREJIND-EQ', '1153': 'GLAXO-EQ', '23650': 'MUTHOOTFIN-EQ', '13538': 'TECHM-EQ', '14413': 'PAGEIND-EQ', '14977': 'POWERGRID-EQ', '3273': 'SRF-EQ', '3339': 'SUNDARMFIN-EQ', '15332': 'NMDC-EQ', '157': 'APOLLOHOSP-EQ', '1901': 'CUMMINSIND-EQ', '2029': 'IRFC-EQ', '13611': 'IRCTC-EQ', '18011': 'WHIRLPOOL-EQ', '25780': 'APLAPOLLO-EQ', '1997': 'LICHSGFIN-EQ', '342': 'CAMS-EQ', '275': 'AUROPHARMA-EQ', '31415': 'NBCC-EQ', '335': 'BALKRISIND-EQ', '3417': 'HONAUT-EQ', '17869': 'JSWENERGY-EQ', '17971': 'SBICARD-EQ', '2475': 'ONGC-EQ', '2277': 'MRF-EQ', '3063': 'VEDL-EQ', '163': 'APOLLOTYRE-EQ', '3150': 'SIEMENS-EQ', '3426': 'TATAPOWER-EQ', '3405': 'TATACHEM-EQ', '3787': 'WIPRO-EQ', '21238': 'AUBANK-EQ', '739': 'COROMANDEL-EQ', '2142': 'MFSL-EQ', '22': 'ACC-EQ', '1627': 'LINDEINDIA-EQ', '3499': 'TATASTEEL-EQ', '438': 'BHEL-EQ', '474': '3MINDIA-EQ', '4749': 'CONCOR-EQ', '1808': 'KAJARIACER-EQ', '19020': 'JSWINFRA-EQ', '676': 'EXIDEIND-EQ', '772': 'DABUR-EQ', '21614': 'ABCAPITAL-EQ', '4244': 'HDFCAMC-EQ', '910': 'EICHERMOT-EQ', '422': 'BHARATFORG-EQ', '6364': 'NATIONALUM-EQ', '277': 'GICRE-EQ', '3351': 'SUNPHARMA-EQ', '20374': 'COALINDIA-EQ', '4668': 'BANKBARODA-EQ', '4717': 'GAIL-EQ', '21951': 'HFCL-EQ', '7229': 'HCLTECH-EQ', '881': 'DRREDDY-EQ', '2412': 'PEL-EQ', '24184': 'PIIND-EQ', '3103': 'SHREECEM-EQ', '3518': 'TORNTPHARM-EQ', '3721': 'TATACOMM-EQ', '383': 'BEL-EQ', '3812': 'ZEEL-EQ', '6733': 'JINDALSTEL-EQ', '3718': 'VOLTAS-EQ', '4503': 'MPHASIS-EQ', '4745': 'BANKINDIA-EQ', '958': 'ESCORTS-EQ', '3432': 'TATACONSUM-EQ', '404': 'BERGEPAINT-EQ', '4067': 'MARICO-EQ', '9819': 'HAVELLS-EQ'}


    def create_excel_sheet(self, workbookName="liveTicks.xlsx", sheetName="LiveTicks"):
        import os
        try:
            workbook_path = os.path.abspath(workbookName)
            if os.path.exists(workbook_path):
                os.remove(workbook_path)
            self.wb = xw.Book()  # Create a new workbook
            self.wb.save(workbook_path)
        except Exception as e:
            print(f"Error recreating workbook: {e}")
            raise

        self.sheet = self.wb.sheets.add(sheetName)
        self.sheet.range("A1").value = [
            "Symbol", "LTP", "Bid Qty", "Ask Qty", "Tick Time",
            "LTQ", "ATP", "Volume", "Open", "High", "Low", "Prev Close",
            "OI", "Upper Circuit", "Lower Circuit", "52WH", "52WL"
        ]
        self.wb.save(workbook_path)


    def aggeregrate_data_mins(self, mins, conditions=None, ema_periods=[6, 9, 11]):
        if conditions is None:
            conditions = {
                "price_change": 0.01,
                "vwap_relation": False,
                "supertrend_relation": False,
                "volume_spike_threshold": 0,
                "vwap_distance_min": 0.001,
                "bid_ask_ratio_min": 0.01,
                "price_above_ema_above_of": None,
                "volume_ma_window": 1
            }

        sheet_name = f"Aggregated_{mins}min"
        if sheet_name not in [s.name for s in self.wb.sheets]:
            agg_sheet = self.wb.sheets.add(sheet_name)
            agg_sheet.range("A1").value = [
                "Symbol", "Open", "High", "Low", "Close",
                "Volume", "Avg ATP", "Start Time", "End Time",
                "RSI", "Supertrend", "VWAP", "VwapDist", "BidAskRatio"
            ] + [f"EMA_{p}" for p in ema_periods]
            self.wb.save("liveTicks.xlsx")

        status_sheet_name = f"Status_{mins}min"
        if status_sheet_name not in [s.name for s in self.wb.sheets]:
            status_sheet = self.wb.sheets.add(status_sheet_name)
            status_sheet.range("A1").value = ["BULLISH"]
            status_sheet.range("A2").value = ["Symbol", "LTP", "%Change", "VWAP", "Supertrend", "Volume"]
            status_sheet.range("G1").value = ["BEARISH"]
            status_sheet.range("G2").value = ["Symbol", "LTP", "%Change", "VWAP", "Supertrend", "Volume"]
            self.wb.save("liveTicks.xlsx")

        def aggregator():
            agg_sheet = self.wb.sheets[sheet_name]
            status_sheet = self.wb.sheets[status_sheet_name]

            while True:
                print(f"⏳ Sleeping for {mins} minute(s)...")
                time.sleep(mins * 60)
                buffer_copy = self.tick_buffer
                self.tick_buffer = []
                print(f"🧪 Aggregating {len(buffer_copy)} ticks at {datetime.now()}")

                if not buffer_copy:
                    continue

                df = pd.DataFrame(buffer_copy)
                df["datetime"] = pd.to_datetime("2024-01-01 " + df["tick_time"])
                df.set_index("datetime", inplace=True)
                df.sort_index(inplace=True)

                df["RSI"] = RSIIndicator(close=df["ltp"], window=14).rsi()
                supertrend = ta.supertrend(high=df["high"], low=df["low"], close=df["ltp"], length=10, multiplier=3.0)
                df["Supertrend"] = supertrend.iloc[:, 0]  # usually 'SUPERT_10_3.0'
                df["VWAP"] = VolumeWeightedAveragePrice(
                    high=df["high"], low=df["low"], close=df["ltp"], volume=df["volume"]
                ).volume_weighted_average_price()

                # Add dynamic EMAs
                for period in ema_periods:
                    df[f"EMA_{period}"] = df["ltp"].ewm(span=period, adjust=False).mean()

                df["vwap_distance"] = df["ltp"] - df["VWAP"]
                df["bid_ask_ratio"] = (df["bid_qty"] - df["ask_qty"]) / (df["bid_qty"] + df["ask_qty"] + 1e-9)

                grouped = df.groupby("symbol")

                bullish_list = []
                bearish_list = []

                for symbol, group in grouped:
                    open_ = group.iloc[0]["ltp"]
                    high = group["high"].max()
                    low = group["low"].min()
                    close = group.iloc[-1]["ltp"]
                    volume = group["volume"].max() - group["volume"].min()
                    avg_atp = group["atp"].mean()
                    start_time = group.iloc[0]["tick_time"]
                    end_time = group.iloc[-1]["tick_time"]

                    latest_row = group.iloc[-1]
                    rsi = latest_row["RSI"]
                    supertrend = latest_row["Supertrend"]
                    vwap = latest_row["VWAP"]
                    vwap_dist = latest_row["vwap_distance"]
                    bid_ask_ratio = latest_row["bid_ask_ratio"]
                    # Gather EMA values dynamically
                    ema_values = [latest_row[f"EMA_{p}"] for p in ema_periods]

                    perc_change = ((close - group["ltp"].iloc[0]) / group["ltp"].iloc[0]) * 100

                    # Compute rolling average of volume
                    group = group.copy()
                    group["volume_ma"] = group["volume"].rolling(window=conditions["volume_ma_window"]).mean()

                    used_range = agg_sheet.range("A1:A1000").value
                    symbol_rows = {row_val: idx + 1 for idx, row_val in enumerate(used_range) if row_val}

                    if symbol in symbol_rows:
                        row_num = symbol_rows[symbol]
                    else:
                        row_num = len(symbol_rows) + 2
                        agg_sheet.range(f"A{row_num}").value = symbol

                    agg_sheet.range(f"B{row_num}").value = [
                        open_, high, low, close, volume, avg_atp, start_time, end_time,
                        rsi, supertrend, vwap, vwap_dist, bid_ask_ratio
                    ] + ema_values

                    # Volume rolling average comparison
                    volume_ma_last = group["volume_ma"].iloc[-1]

                    if (
                        perc_change > conditions["price_change"]
                        and (close > vwap if conditions["vwap_relation"] else True)
                        and (close > supertrend if conditions["supertrend_relation"] else True)
                        and (volume > volume_ma_last)
                        and (vwap_dist > conditions["vwap_distance_min"])
                        and (bid_ask_ratio > conditions["bid_ask_ratio_min"])
                        and (close > latest_row[f"EMA_{conditions['price_above_ema_above_of']}"] if conditions["price_above_ema_above_of"] else True)
                    ):
                        bullish_list.append([symbol, close, perc_change, vwap, supertrend, volume])
                    elif (
                        perc_change < -conditions["price_change"]
                        and (close < vwap if conditions["vwap_relation"] else True)
                        and (close < supertrend if conditions["supertrend_relation"] else True)
                        and (volume > volume_ma_last)
                        and (vwap_dist > conditions["vwap_distance_min"])
                        and (bid_ask_ratio > conditions["bid_ask_ratio_min"])
                        and (close < latest_row[f"EMA_{conditions['price_above_ema_above_of']}"] if conditions["price_above_ema_above_of"] else True)
                    ):
                        bearish_list.append([symbol, close, perc_change, vwap, supertrend, volume])

                if bullish_list:
                    status_sheet.range("A3").value = bullish_list
                if bearish_list:
                    status_sheet.range("G3").value = bearish_list
                self.wb.save("liveTicks.xlsx")

        threading.Thread(target=aggregator, daemon=True).start()

    def monitor_connection(self):
        def watchdog():
            while True:
                if time.time() - self.last_tick_time > 30:
                    logger.error("⚠️ WebSocket likely disconnected! No ticks in last 30 seconds.")
                time.sleep(10)
        threading.Thread(target=watchdog, daemon=True).start()
