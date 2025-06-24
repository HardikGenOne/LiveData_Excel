# 📈 Live Tick Analyzer

A real-time stock tick data monitoring and analysis system built using Angel One SmartAPI.  
It identifies high-momentum intraday stocks based on smart technical indicators and writes filtered insights to Excel for actionable trading decisions.

## 🧠 Why I Built This

As a retail trader, I needed a lightweight but powerful intraday engine that could:
- Process tick-by-tick market data with minimal latency.
- Apply live filtering logic based on VWAP, Supertrend, EMAs, volume spikes, etc.
- Work offline without expensive tools like TradingView Premium or broker terminals.
- Be fully customizable, developer-friendly, and extensible.

This project acts as both a **research engine** and a **practical live signal tool** for scalping or short-term trading.

## ⚙️ Key Features

- ✅ Live NIFTY50 tick streaming via Angel One SmartAPI.
- ✅ Tick aggregation into custom intervals (1-min, 5-min, etc.).
- ✅ Momentum filters using:
  - Price % change
  - VWAP distance
  - Supertrend position
  - Rolling volume spikes
  - Bid/Ask ratio analysis
  - EMA thresholds (e.g., price above EMA 9)
- ✅ Output written directly to Excel in real time.
- ✅ Fully parameterized filtering conditions — no hardcoding.
- ✅ Designed for backtesting-ready extensions.

## 📦 Folder Structure

```
.
├── fetching_tickData.py       # Main class for streaming + aggregation
├── main.py                    # Entry point to run everything
├── .env.template              # Sample environment config
├── requirements.txt           # Dependencies
├── liveTicks.xlsx             # Output (auto-generated)
└── README.md                  # You’re reading it :)
```

## 🔐 API & Secret Handling

Credentials (API keys, passwords, etc.) are managed securely using a `.env` file.

**Steps:**
1. Create a file named `.env`:
   ```
   AUTH_TOKEN=your_auth_token
   API_KEY=your_api_key
   CLIENT_CODE=your_client_code
   PASS=your_password
   TOTP_SECRET=your_totp_secret
   ```

2. Ensure `.env` is ignored in `.gitignore`.

3. Load these in your script using:
   ```python
   from dotenv import load_dotenv
   import os

   load_dotenv()
   AUTH_TOKEN = os.getenv("AUTH_TOKEN")
   ```

## 🚀 How to Run

1. Clone this repo:
   ```bash
   git clone https://github.com/your-username/live-tick-analyzer.git
   cd live-tick-analyzer
   ```

2. Create virtual environment and install dependencies:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

3. Add your credentials to `.env`.

4. Run the app:
   ```bash
   python test.py
   ```

## 🔮 Roadmap

- [ ] Real-time alerts (sound, Telegram)
- [ ] Web-based dashboard for visualization
- [ ] BankNifty & Option chain signal integration
- [ ] Historical data + backtesting
- [ ] Trade execution simulation/paper-trading module

## 📄 License

This project is released under the [MIT License](LICENSE).

## ⚠️ Disclaimer

This project is for **educational and research use only**.  
It does **not** place live trades.  
Use at your own risk — especially when applying it to financial markets.
