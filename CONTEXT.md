# Project Context: Google Receipt Extractor & Budget Tracker

## 1. Architecture Overview
* **Environment:** Google Apps Script (`clasp` deployed).
* **Ingestion (`doGet`):** Webhook triggered by MacroDroid on Android. Listens to Google Wallet, credit card apps (Isracard, MAX), and SMS.
* **Reporting (`checkAndNotify`):** Time-based trigger that calculates daily diffs, generates a 30-day moving average chart, and sends a payload to Telegram.

## 2. Critical Bug Fixes & "Do Not Touch" Logic
The current codebase contains specific logic to handle real-world Android and Israeli banking quirks. **Do not remove or simplify these safeguards:**

* **Concurrency (`LockService`):** Because MacroDroid listens to multiple sources simultaneously to prevent blind spots, duplicate webhooks will fire. The script uses `LockService.getScriptLock()` with a 30-second wait to process requests sequentially.
* **Ghost Replay Filter (Rule B):** Android notification grouping causes old notifications to replay hours later. We drop exact text + exact amount matches that arrive within a 12-hour window (`43200` seconds).
* **Invisible Character Exorcism:** Israeli bank notifications inject invisible Unicode direction marks (e.g., `\u200f`, `\u202A`). We run `.replace(/[\u200B-\u200F\u202A-\u202E\u061C]/g, " ")` on the raw text before running regex to prevent match failures.
* **Strict Amount Extraction:** * **Level 1:** Checks keywords (`חוייב ב`, `בסך`, `Total`).
    * **Level 2:** Checks for explicit currency symbols/codes (`₪`, `NIS`, `ILS`, `$`, `USD`).
    * **No Fallback:** If there is no currency symbol or keyword, the script must ignore the payload entirely to reject promotional spam (e.g., "YUMMI 28.2.26").
* **Live Currency Conversion:** Non-ILS transactions are converted using the free `api.exchangerate-api.com` endpoint before saving. Supported currencies and their trigger tokens:
    * **USD:** `USD`, `$`
    * **EUR:** `EUR`, `EURO`, `€`
    * **CZK:** `CZK`, `Kč`
    * **GBP:** `GBP`, `£`
    * To add a new currency: add its tokens to `CURRENCY_PATTERNS` and its code to `FOREIGN_CURRENCIES` in `parsing.js`, and to `foreignCurrencies` in `fixUnconvertedForeignTransactions`.
    * If the live conversion fails, the row keeps the foreign amount and is tagged `[UNCONVERTED: …]`; the "Fix Unconverted Foreign Transactions" menu item repairs it with historical rates.
* **Shared logic & tests:** Pure parsing (ad filter, merchant, amount, category validation) lives in `parsing.js`. `node tests.js` loads the real pushed files with mocked Apps Script services. `tests.js` is excluded in `.claspignore` — keep it that way.
* **Secrets (`PropertiesService`):** Never hardcode tokens, API keys, or Telegram chat IDs. Read `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_IDS` (comma-separated), and `GEMINI_API_KEY` from `PropertiesService.getScriptProperties()`. `.clasp.json` is gitignored.

## 3. Google Sheets Quirks
* **Tables / Column Types:** The target Google Sheet utilizes the newer "Tables / Typed Columns" feature for the `Category` dropdown. Do not attempt to use classic Data Validation APIs (`DataValidationBuilder`) to manage the dropdowns in the script, as the Typed Column strictly restricts standard data validation.
* **Execution Time Limits:** The `checkAndNotify` chart generation uses an O(N) sliding window array for the 30-day moving average. Do not use `Utilities.formatDate` inside nested loops for historical lookbacks, as the API call volume will trigger Google's hard 6-minute execution timeout.