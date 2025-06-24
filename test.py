from fetching_tickData import Excel_Live_Ticks
from dotenv import load_dotenv
import os

load_dotenv()

AUTH_TOKEN = os.getenv("AUTH_TOKEN")
API_KEY = os.getenv("API_KEY")
CLIENT_CODE = os.getenv("CLIENT_CODE")
PASS = os.getenv("PASS")
TOTP_SECRET = os.getenv("TOTP_SECRET")

# === INIT ===
tickData = Excel_Live_Ticks(AUTH_TOKEN, API_KEY, CLIENT_CODE, PASS, totp_secret)

tickData.create_excel_sheet(workbookName="Temp.xlsx",sheetName="tempLive")

tickData.connect_to_AngleOne()

# Call aggregation functions with optional EMA periods and filtering conditions
tickData.aggeregrate_data_mins(
    mins=1,
    ema_periods=[6, 9],
    conditions={
        "price_change": 0.3,
        "vwap_relation": True,
        "supertrend_relation": True,
        "volume_spike_threshold": 20000,
        "vwap_distance_min": 0.2,
        "bid_ask_ratio_min": 0.3,
        "price_above_ema_above_of": 9
    }
)

tickData.aggeregrate_data_mins(
    mins=5,
    ema_periods=[9, 21],
    conditions={
        "price_change": 0.5,
        "vwap_relation": True,
        "supertrend_relation": False,
        "volume_spike_threshold": 30000,
        "vwap_distance_min": 0.1,
        "bid_ask_ratio_min": 0.2,
        "price_above_ema_above_of": 21
    }
)

tickData.start_NIFTY50_streaming()