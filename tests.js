/**
 * Test Suite for Receipt Extractor.
 * Run with: node tests.js
 *
 * Loads the REAL production files (every .js file in this folder except this one —
 * the same set clasp pushes) into Node with mocked Apps Script services, then
 * exercises the actual entry points: doGet, checkAndNotify, getGeminiCategory and
 * the maintenance tools. Nothing here re-implements production logic.
 *
 * NOTE: This file must stay excluded from `clasp push` (see .claspignore) —
 * Apps Script runs top-level code of every file on every execution.
 */
var fs = require("fs");
var path = require("path");
var vm = require("vm");

// =====================================================================
// Mock Apps Script environment
// =====================================================================

var RATES = { USD: 3.7, EUR: 4.0, CZK: 0.16, GBP: 4.7 };
var TX_HEADER = ["Date", "Merchant", "Amount", "Category", "Raw Text"];
var env; // Per-test state, rebuilt by setupEnv()

function FakeSheet(name, rows, id) {
  this.name = name;
  this.rows = rows.map(function (r) { return r.slice(); });
  this.id = id;
  this.appendError = null;
}
FakeSheet.prototype.getLastRow = function () { return this.rows.length; };
FakeSheet.prototype.getSheetId = function () { return this.id; };
FakeSheet.prototype.appendRow = function (values) {
  if (this.appendError) throw this.appendError;
  this.rows.push(values.slice());
};
FakeSheet.prototype.deleteRow = function (n) { this.rows.splice(n - 1, 1); };
FakeSheet.prototype.getRange = function (row, col, numRows, numCols) {
  numRows = numRows === undefined ? 1 : numRows;
  numCols = numCols === undefined ? 1 : numCols;
  // Mirrors the real Apps Script error
  if (numRows < 1) throw new Error("The number of rows in the range must be at least 1.");
  if (numCols < 1) throw new Error("The number of columns in the range must be at least 1.");
  var sheet = this;
  return {
    getValues: function () {
      var out = [];
      for (var r = 0; r < numRows; r++) {
        var src = sheet.rows[row - 1 + r] || [];
        var line = [];
        for (var c = 0; c < numCols; c++) {
          var v = src[col - 1 + c];
          line.push(v === undefined ? "" : v);
        }
        out.push(line);
      }
      return out;
    },
    setValues: function (values) {
      for (var r = 0; r < numRows; r++) {
        if (!sheet.rows[row - 1 + r]) sheet.rows[row - 1 + r] = [];
        for (var c = 0; c < numCols; c++) sheet.rows[row - 1 + r][col - 1 + c] = values[r][c];
      }
    }
  };
};

function makeSpreadsheet(sheetRows) {
  var map = {};
  var nextId = 0;
  Object.keys(sheetRows).forEach(function (n) { map[n] = new FakeSheet(n, sheetRows[n], nextId++); });
  return {
    sheets: map,
    getSheetByName: function (n) { return map[n] || null; },
    insertSheet: function (n) { map[n] = new FakeSheet(n, [], nextId++); return map[n]; },
    getUrl: function () { return "https://docs.google.com/spreadsheets/d/TEST/edit"; },
    getSpreadsheetTimeZone: function () { return "Asia/Jerusalem"; }
  };
}

function jsonResponse(code, obj) {
  return {
    getResponseCode: function () { return code; },
    getContentText: function () { return JSON.stringify(obj); }
  };
}

function fakeFetch(url, options) {
  env.fetches.push(url);
  if (url.indexOf("exchangerate-api.com") > -1) {
    if (env.rateFail) throw new Error("Exception: Address unavailable");
    var cur = url.split("/").pop();
    return jsonResponse(200, { rates: env.rateMissing ? {} : { ILS: RATES[cur] } });
  }
  if (url.indexOf("frankfurter.app") > -1) {
    var from = /from=(\w+)/.exec(url)[1];
    return jsonResponse(200, { rates: { ILS: RATES[from] } });
  }
  if (url.indexOf("generativelanguage.googleapis.com") > -1) {
    env.geminiCalls.push({ lockHeld: env.lockHeld, payload: options && options.payload });
    return jsonResponse(200, { candidates: [{ content: { parts: [{ text: env.geminiReply }] } }] });
  }
  if (url.indexOf("api.telegram.org") > -1) {
    env.telegram.push({ url: url, payload: options && options.payload });
    return jsonResponse(200, { ok: true });
  }
  throw new Error("Unexpected fetch in test: " + url);
}

// Chainable stub for the Charts builder API (any property / call returns the stub)
var chartsStub = new Proxy(function () {}, {
  get: function (t, prop) {
    if (prop === "getAs") return function () { return { fakeBlob: true }; };
    return chartsStub;
  },
  apply: function () { return chartsStub; }
});

function setupEnv(opts) {
  opts = opts || {};
  var ss = makeSpreadsheet({
    Transactions: [TX_HEADER].concat(opts.transactions || []),
    Categories: [["Merchant", "Category"]].concat(opts.categories || []),
    Dashboard: opts.dashboard || [["", "", "", "", "Category", "Budget", "Spent"]]
  });
  env = {
    ss: ss,
    lockHeld: false,
    lockBusy: !!opts.lockBusy,
    rateFail: !!opts.rateFail,
    rateMissing: !!opts.rateMissing,
    geminiReply: opts.geminiReply === undefined ? "Groceries" : opts.geminiReply,
    geminiCalls: [],
    telegram: [],
    fetches: [],
    alerts: []
  };

  globalThis.SpreadsheetApp = {
    getActiveSpreadsheet: function () { return ss; },
    flush: function () {},
    getUi: function () { return { alert: function (m) { env.alerts.push(String(m)); } }; }
  };
  globalThis.ContentService = {
    createTextOutput: function (t) { return { text: t, getContent: function () { return t; } }; }
  };
  globalThis.LockService = {
    getScriptLock: function () {
      return {
        waitLock: function () {
          if (env.lockBusy) throw new Error("Lock timeout: another process was holding the lock for too long.");
          env.lockHeld = true;
        },
        releaseLock: function () { env.lockHeld = false; },
        hasLock: function () { return env.lockHeld; }
      };
    }
  };
  var scriptProps = Object.assign({
    TELEGRAM_BOT_TOKEN: "TEST_BOT_TOKEN",
    TELEGRAM_CHAT_IDS: "111111111,222222222",
    GEMINI_API_KEY: "TEST_GEMINI_KEY"
  }, opts.scriptProperties || {});

  globalThis.PropertiesService = {
    getScriptProperties: function () {
      return {
        getProperty: function (key) { return scriptProps[key] || null; }
      };
    }
  };
  globalThis.UrlFetchApp = { fetch: fakeFetch };
  globalThis.Utilities = {
    sleep: function () {},
    formatDate: function () { return "October"; }
  };
  globalThis.Charts = chartsStub;
  return env;
}

// =====================================================================
// Load the real production code (same file set clasp pushes)
// =====================================================================

var SOURCE_FILES = fs.readdirSync(__dirname)
  .filter(function (f) { return /\.js$/.test(f) && f !== path.basename(__filename); })
  .sort();
SOURCE_FILES.forEach(function (f) {
  vm.runInThisContext(fs.readFileSync(path.join(__dirname, f), "utf8"), { filename: f });
});

// =====================================================================
// Helpers
// =====================================================================

function send(merchant, text) {
  var params = { text: text };
  if (merchant !== undefined) params.merchant = merchant;
  var out = doGet({ parameter: params });
  return out.text;
}
function txRows() { return env.ss.sheets.Transactions.rows.slice(1); }
function lastTx() {
  var r = txRows()[txRows().length - 1];
  if (!r) return null;
  return { date: r[0], merchant: r[1], amount: r[2], category: r[3], raw: r[4] };
}
function cachedMerchants() {
  return env.ss.sheets.Categories.rows.slice(1).map(function (r) { return String(r[0]); });
}
function hoursAgo(h) { return new Date(Date.now() - h * 3600 * 1000); }

function AssertionError(msg) { this.message = msg; }
function eq(actual, expected, label) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new AssertionError((label ? label + ": " : "") + "expected " + JSON.stringify(expected) + ", got " + JSON.stringify(actual));
  }
}
function approx(actual, expected, label) {
  if (typeof actual !== "number" || isNaN(actual) || Math.abs(actual - expected) > 0.011) {
    throw new AssertionError((label ? label + ": " : "") + "expected ≈" + expected + ", got " + JSON.stringify(actual) + " (" + typeof actual + ")");
  }
}
function ok(cond, msg) { if (!cond) throw new AssertionError(msg); }

// =====================================================================
// Runner
// =====================================================================

var passed = 0;
var failures = [];
var realLog = console.log;

function section(title) { realLog("\n" + title); }
function test(name, fn) {
  var logs = [];
  console.log = function () { logs.push(Array.prototype.slice.call(arguments).join(" ")); };
  var error = null;
  try { fn(); } catch (e) { error = e; }
  console.log = realLog;
  if (!error) {
    passed++;
    realLog("  ✅ PASS: " + name);
  } else {
    failures.push(name);
    realLog("  ❌ FAIL: " + name);
    realLog("       " + (error instanceof AssertionError ? error.message : "Uncaught " + (error && error.stack || error)));
    if (logs.length) realLog("       (last log: " + logs[logs.length - 1] + ")");
  }
}

realLog("==========================================");
realLog("🧪 Running Receipt Extractor Tests...");
realLog("   Loaded: " + SOURCE_FILES.join(", "));
realLog("==========================================");

// ---------------------------------------------------------------------
section("1. Spam & Ad Filter (regression)");

test("Ad with promo keyword should be rejected", function () {
  setupEnv();
  eq(send("SMS", "מבצע מיוחד 20% הנחה על כל המוצרים!"), "Ignored: Ad");
  eq(txRows().length, 0, "rows written");
});
test("SMS with no transaction keywords should be rejected", function () {
  setupEnv();
  eq(send("SMS", "הודעה כללית ללא מילת מפתח עסקאות ב-100 שח"), "Ignored: SMS Ad");
});
test("Real transaction with 'חוייב' safety word should pass even if ad word is present", function () {
  setupEnv();
  eq(send("SMS", "חשבונך חוייב בסך 150.00 ₪ עבור מבצע הרהיטים"), "Success");
  approx(lastTx().amount, 150);
});
test("Promotional SMS with 'במקום', 'בלבד', 'מועדון', 'כנסו' should be rejected as Ad", function () {
  setupEnv();
  var promo = "דאגנו לכם למכונת ברד וגלידה 3 ב-1 של SAUTER ב-₪599 בלבד במקום ₪1,490 ועוד מוצרים מטורפים במחירי מועדון שווים. כנסו להציץ >>";
  eq(send("SAUTER", promo), "Ignored: Ad");
  eq(send("SMS", promo), "Ignored: Ad");
  eq(send("050-1234567", promo), "Ignored: Ad");
  eq(txRows().length, 0, "rows written");
});
test("Promotional SMS with 'להסרה' opt-out should be rejected as Ad", function () {
  setupEnv();
  eq(send("Castro", "סופר סייל! כל החנות עד 50% הנחה. להסרה השב הסר"), "Ignored: Ad");
  eq(txRows().length, 0, "rows written");
});
test("Credit card promo SMS containing 'בכרטיסי' and 'בסך 10,000 ₪' must be rejected as Ad", function () {
  setupEnv();
  var flyCardAd = "מצטרפים עכשיו ל-FLY CARD PREMIUM של אמריקן אקספרס, נהנים מפטור מדמי כרטיס ומקבלים כרטיס טיסה מתנה...\nרק בכרטיסי אמריקן אקספרס\nלפרטים המלאים: https://example.com למצטרפים חדשים בהתאם לתקנון המבצע, עד 31.12.26. כפוף לרכישות בסך 10,000 ₪ עד 31.03.27. אי־עמידה בפירעון ההלוואה או האשראי עלולה לגרור ריבית פיגורים.\nלהסרה מרשימת התפוצה ב SMS יש לשלוח הסר למספר 0530000000";
  eq(send("Amex", flyCardAd), "Ignored: Ad");
  eq(send("SMS", flyCardAd), "Ignored: Ad");
  eq(txRows().length, 0, "rows written");
});
test("Promo SMS with unsubscribe footer is rejected even if it mentions a transaction safety word", function () {
  setupEnv();
  eq(send("PromoCard", "בצע עסקה בסך 500 ₪ וקבל מתנה! להסרה מרשימת התפוצה יש לשלוח הסר"), "Ignored: Ad");
  eq(txRows().length, 0, "rows written");
});
test("MacroDroid temporary receipt notification should be rejected", function () {
  setupEnv();
  eq(send("Preparing your receipt...", "Processing"), "Ignored: Temporary");
});

// ---------------------------------------------------------------------
section("2. Merchant Name Parsing (regression)");

test("Clean Isracard format", function () {
  setupEnv();
  send("SMS", "בסך 52.94 ILS בaliexpress - LUXEMBOURG.");
  eq(lastTx().merchant, "aliexpress - LUXEMBOURG");
});
test("Clean Isracard format with transaction date prefix", function () {
  setupEnv();
  send("Isracard", "שלום,\nבכרטיסך המסתיים ב- 1234, אושרה עסקה ב-14/10 בסך 500.00 ש\"ח בד\"ר בוקסנבאום בע\"מ.\nלמידע נוסף באפליקציה ובאתר: https://isracard.onelink.me/bajD/example1.\nלשירותך,  ישראכרט");
  eq(lastTx().merchant, "ד\"ר בוקסנבאום בע\"מ");
  approx(lastTx().amount, 500);
});
test("Clean Isracard format with internal dots in merchant acronym (אי.אם.פי.אם)", function () {
  setupEnv();
  send("Isracard", "שלום,\nבכרטיסך המסתיים ב- 1234, אושרה עסקה ב-24/10 בסך 60.71 ש\"ח באי.אם.פי.אם כצנלסון.\nלמידע נוסף באפליקציה ובאתר: https://isracard.onelink.me/bajD/example1.\nלשירותך,  ישראכרט");
  eq(lastTx().merchant, "אי.אם.פי.אם כצנלסון");
  approx(lastTx().amount, 60.71);
});
test("Clean Isracard format with internal dots in merchant acronym (קיי.אס.פי)", function () {
  setupEnv();
  send("Isracard", "שלום,\nבכרטיסך 1234 אושרה עסקה ב-25/10 בסך 49.00 ש\"ח בקיי.אס.פי מחשבים רמת.\nלמידע נוסף באפליקציה ובאתר: https://isracard.onelink.me/bajD/example2.\nלשירותך,  ישראכרט");
  eq(lastTx().merchant, "קיי.אס.פי מחשבים רמת");
  approx(lastTx().amount, 49.00);
});
test("Clean Cal/Max format", function () {
  setupEnv();
  send("SMS", "עסקה חדשה ב-KSP בסך 499 ש\"ח");
  eq(lastTx().merchant, "KSP");
  approx(lastTx().amount, 499);
});
test("Clean Cal format with 'Cal' sender and 'בכרטיס' (without ך)", function () {
  setupEnv();
  send("Cal", "עסקה ב-שופרסל בסך 120.50 ₪ בכרטיס 1234");
  eq(lastTx().merchant, "שופרסל");
  approx(lastTx().amount, 120.5);
});
test("Clean Cibus format", function () {
  setupEnv();
  send("Cibus", "חויבת ב-45.00 ₪ עבור עסקה במסעדת ג'אפן ג'אפן הסיבוס שלך");
  eq(lastTx().merchant, "ג'אפן ג'אפן");
  approx(lastTx().amount, 45);
});
test("Clean 'Payment to' format", function () {
  setupEnv();
  send("Payment to McDonald's", "Payment to McDonald's ₪45.00");
  eq(lastTx().merchant, "McDonald's");
});

// ---------------------------------------------------------------------
section("3. Amount, Currency, Dedup & Cache (regression)");

test("Extract ILS amount from Israeli credit card message", function () {
  setupEnv();
  send("SMS", "עסקה בסך 150.00 ₪ ב-שופרסל");
  approx(lastTx().amount, 150);
});
test("USD is converted to ILS and tagged [Orig:]", function () {
  setupEnv();
  send("Amazon", "Charged $50.25 at Amazon");
  approx(lastTx().amount, 50.25 * RATES.USD);
  ok(lastTx().raw.indexOf("[Orig:") > -1, "raw text should contain [Orig:]");
});
test("Extract amount containing invisible text-direction Unicode marks", function () {
  setupEnv();
  send("SMS", "\u200Fעסקה בסך\u202A 99.90 ₪");
  approx(lastTx().amount, 99.9);
});
test("Reject payload with no valid amount or currency", function () {
  setupEnv();
  eq(send("Google Wallet", "Just a random text message without money"), "Ignored: No Amount");
});
test("Rule A: same amount within 60s is a Rapid Duplicate", function () {
  setupEnv({ categories: [["wolt", "Restaurants"]] });
  eq(send("Wolt", "Payment ₪60"), "Success");
  eq(send("Wolt", "Payment to Wolt ₪60"), "Ignored: Rapid Duplicate");
  eq(txRows().length, 1, "rows written");
});
test("Rule B: identical text + amount within 12h is a Ghost Replay", function () {
  setupEnv({ transactions: [[hoursAgo(5), "Wolt", 60, "Restaurants", "Payment to Wolt ₪60"]] });
  eq(send("Wolt", "Payment to Wolt ₪60"), "Ignored: Ghost Replay");
});
test("Cache hit categorizes without calling Gemini", function () {
  setupEnv({ categories: [["Wolt", "Restaurants"]] });
  send("Wolt", "Payment ₪60");
  eq(lastTx().category, "Restaurants");
  eq(env.geminiCalls.length, 0, "Gemini calls");
});
test("Partial word does not match in cache, calls Gemini and caches full merchant", function () {
  setupEnv({ categories: [["קפה", "Restaurants"]], geminiReply: "Restaurants" });
  send("SMS", "עסקה ב-רד קפה בייקרי בע''מ בסך 50 ₪");
  eq(lastTx().merchant, "רד קפה בייקרי בע''מ");
  eq(lastTx().category, "Restaurants");
  eq(env.geminiCalls.length, 1, "Gemini should be called for full name");
  ok(cachedMerchants().indexOf("רד קפה בייקרי בע''מ") > -1, "full merchant should be cached");
});
test("Cache hit matches nearly identical merchant names (quotes, hyphens, בעמ/בע''מ)", function () {
  setupEnv({ categories: [["רד קפה בייקרי בע\"מ", "Restaurants"]] });
  send("SMS", "עסקה ב-רד קפה בייקרי בע''מ בסך 50 ₪");
  eq(lastTx().category, "Restaurants");
  eq(env.geminiCalls.length, 0, "Gemini should NOT be called — matched nearly identical cached name");
});
test("Cache hit matches drugstore branch prefix and handles abbreviation dot", function () {
  setupEnv({ categories: [["י. דראגסטורס", "Health"]] });
  send("SMS", "עסקה ב-י דראגסטורס - פרישמ בסך 75.90 ₪");
  eq(lastTx().category, "Health");
  eq(env.geminiCalls.length, 0, "Gemini should NOT be called — matched chain prefix with dot");
});
test("Cache hit handles terminal truncation (פרישמ vs פרישמן) and Unicode RLM", function () {
  setupEnv({ categories: [["י דראגסטורס - פרישמן", "Health"]] });
  send("SMS", "עסקה ב-\u200Fי דראגסטורס - פרישמ בסך 50 ₪");
  eq(lastTx().category, "Health");
  eq(env.geminiCalls.length, 0, "Gemini should NOT be called — matched truncated name despite RLM");
});
test("Unknown merchant is categorized by Gemini and cached", function () {
  setupEnv({ geminiReply: "Groceries" });
  send("Shufersal Deal", "Payment ₪80");
  eq(lastTx().category, "Groceries");
  ok(cachedMerchants().indexOf("Shufersal Deal") > -1, "merchant should be cached");
});

// ---------------------------------------------------------------------
section("4. Bug #4 — merchant regexes overriding each other");

test("#4a Cal SMS with trailing date keeps the merchant", function () {
  setupEnv();
  send("SMS", "עסקה חדשה ב-KSP בסך 499 ש\"ח בתאריך 01/10");
  eq(lastTx().merchant, "KSP");
});
test("#4b Cal SMS with trailing card number keeps the merchant", function () {
  setupEnv();
  send("SMS", "עסקה ב-שופרסל בסך 120.50 ₪ בכרטיס 1234");
  eq(lastTx().merchant, "שופרסל");
});
test("#4c (new) 'בסך X ₪ ב-Merchant' has no leading hyphen", function () {
  setupEnv();
  send("SMS", "עסקה בסך 150.00 ₪ ב-שופרסל");
  eq(lastTx().merchant, "שופרסל");
});

// ---------------------------------------------------------------------
section("5. Bug #5 — decimal / thousands separators, GBP");

test("#5a 'Total 1.234,56 EUR' → 1234.56 EUR converted", function () {
  setupEnv();
  send("Booking.com", "Total 1.234,56 EUR");
  approx(lastTx().amount, 1234.56 * RATES.EUR);
});
test("#5b Comma decimal '250,00 Kč' → 250 CZK (not 25000)", function () {
  setupEnv();
  send("Restaurace U Fleku", "Platba 250,00 Kč");
  approx(lastTx().amount, 250 * RATES.CZK);
});
test("#5c GBP '£20.00' is detected and converted", function () {
  setupEnv();
  eq(send("Tesco", "Charged £20.00 at Tesco"), "Success");
  approx(lastTx().amount, 20 * RATES.GBP);
});
test("#5d (regression) Israeli '1,234.56 ₪' → 1234.56", function () {
  setupEnv();
  send("SMS", "חוייב בסך 1,234.56 ₪");
  approx(lastTx().amount, 1234.56);
});
test("#5e (regression) Thousands-only '1,234 ₪' → 1234", function () {
  setupEnv();
  send("IKEA", "Payment 1,234 ₪");
  approx(lastTx().amount, 1234);
});

// ---------------------------------------------------------------------
section("6/7. Bugs #6 & #7 — NaN amounts and failed conversions");

test("#6 Exchange API without an ILS rate must not save NaN", function () {
  setupEnv({ rateMissing: true });
  send("Amazon", "Charged $10 at Amazon");
  ok(lastTx() !== null, "row should still be saved");
  approx(lastTx().amount, 10, "amount");
});
test("#7 Failed conversion is tagged [UNCONVERTED], not [Orig:]", function () {
  setupEnv({ rateFail: true });
  send("Amazon", "Charged $10 at Amazon");
  ok(lastTx().raw.indexOf("[Orig:") === -1, "must not claim it was converted: " + lastTx().raw);
  ok(lastTx().raw.indexOf("[UNCONVERTED") > -1, "should be tagged [UNCONVERTED]: " + lastTx().raw);
});
test("#7b fixUnconvertedForeignTransactions repairs [UNCONVERTED] rows", function () {
  setupEnv({ transactions: [[hoursAgo(30), "Amazon", 10, "Shopping", "Charged $10 at Amazon [UNCONVERTED: $10]"]] });
  fixUnconvertedForeignTransactions();
  var row = lastTx();
  approx(row.amount, 10 * RATES.USD);
  ok(row.raw.indexOf("[Orig: 10 USD]") > -1, "should be tagged [Orig:]: " + row.raw);
  ok(row.raw.indexOf("[UNCONVERTED") === -1, "UNCONVERTED tag should be removed: " + row.raw);
});

// ---------------------------------------------------------------------
section("8. Bug #8 — ad keywords matched as substrings / case-sensitive");

test("#8a 'Costco Wholesale' is not rejected as an ad", function () {
  setupEnv();
  eq(send("Costco", "Costco Wholesale ₪300"), "Success");
});
test("#8b Upper-case 'SALE' is rejected as an ad", function () {
  setupEnv();
  eq(send("SuperStore", "BIG SALE today only ₪99"), "Ignored: Ad");
});
test("#8c fixAdTransactions keeps the Costco Wholesale row", function () {
  setupEnv({
    transactions: [
      [hoursAgo(48), "Costco", 300, "Groceries", "Costco Wholesale ₪300"],
      [hoursAgo(24), "SMS", 99, "Uncategorized", "BIG SALE today only ₪99"]
    ]
  });
  fixAdTransactions();
  eq(txRows().map(function (r) { return r[1]; }), ["Costco"], "remaining merchants");
});
test("#8d fixAdTransactions removes SAUTER promotional row even if merchant is not 'SMS'", function () {
  setupEnv({
    transactions: [
      [hoursAgo(48), "Shufersal", 250, "Groceries", "בכרטיסך אושרה עסקה בסך 250 ש\"ח בשופרסל"],
      [hoursAgo(2), "SAUTER", 599, "Home Improvement", "דאגנו לכם למכונת ברד וגלידה 3 ב-1 של SAUTER ב-₪599 בלבד במקום ₪1,490 ועוד מוצרים מטורפים במחירי מועדון שווים. כנסו להציץ >>"]
    ]
  });
  fixAdTransactions();
  eq(txRows().map(function (r) { return r[1]; }), ["Shufersal"], "remaining merchants");
});
test("#8e fixAdTransactions removes FlyCard 10,000 ILS promo row containing 'בכרטיסי'", function () {
  setupEnv({
    transactions: [
      [hoursAgo(48), "Shufersal", 250, "Groceries", "בכרטיסך אושרה עסקה בסך 250 ש\"ח בשופרסל"],
      [hoursAgo(2), "Amex", 10000, "Uncategorized", "רק בכרטיסי אמריקן אקספרס. כפוף לרכישות בסך 10,000 ₪. להסרה מרשימת התפוצה ב SMS יש לשלוח הסר"]
    ]
  });
  fixAdTransactions();
  eq(txRows().map(function (r) { return r[1]; }), ["Shufersal"], "remaining merchants");
});

// ---------------------------------------------------------------------
section("9. Bug #9 — merchant cache poisoning");

test("#9a Unparsed SMS ('SMS' merchant) is not sent to Gemini or cached", function () {
  setupEnv({ geminiReply: "Groceries" });
  send("SMS", "חויבת 50 ₪ בכרטיסך");
  eq(env.geminiCalls.length, 0, "Gemini calls");
  ok(cachedMerchants().indexOf("SMS") === -1, "'SMS' must not be cached");
  eq(lastTx().category, "Uncategorized");
});
test("#9b Phone-number sender is not cached as a merchant", function () {
  setupEnv({ geminiReply: "Groceries" });
  send("+972-52-1234567", "חויבת 50 ₪");
  ok(cachedMerchants().indexOf("+972-52-1234567") === -1, "phone number must not be cached");
});

// ---------------------------------------------------------------------
section("10. Bug #10 — unvalidated Gemini output");

test("#10a Off-list answer 'Food' → Uncategorized", function () {
  setupEnv({ geminiReply: "Food" });
  eq(getGeminiCategory("Some Place"), "Uncategorized");
});
test("#10b Trailing punctuation 'Restaurants.' → Restaurants", function () {
  setupEnv({ geminiReply: "Restaurants." });
  eq(getGeminiCategory("Some Place"), "Restaurants");
});
test("#10c JSON + wrong case '{\"category\": \"groceries\"}' → Groceries", function () {
  setupEnv({ geminiReply: "{\"category\": \"groceries\"}" });
  eq(getGeminiCategory("Some Place"), "Groceries");
});
test("#10d doGet does not cache an off-list AI answer", function () {
  setupEnv({ geminiReply: "Food" });
  send("Mystery Place", "Payment ₪40");
  eq(lastTx().category, "Uncategorized");
  eq(cachedMerchants().length, 0, "cache rows");
});
test("#10e getGeminiCategory enables Google Search grounding tool for merchant lookup (e.g. KOKPIT 67)", function () {
  setupEnv({ geminiReply: "Restaurants" });
  eq(getGeminiCategory("KOKPIT 67"), "Restaurants");
  eq(env.geminiCalls.length, 1, "Gemini calls");
  var body = JSON.parse(env.geminiCalls[0].payload);
  ok(body.tools && body.tools[0] && body.tools[0].google_search !== undefined, "google_search tool should be enabled in payload");
});
test("#10f normalizeCategory handles singular 'Restaurant', citations 'Restaurants [1]', and multi-line search answers", function () {
  eq(normalizeCategory("Restaurant"), "Restaurants");
  eq(normalizeCategory("Restaurants [1]"), "Restaurants");
  eq(normalizeCategory("KOKPIT 67 is a cocktail bar in Givatayim [1].\nRestaurants"), "Restaurants");
});

// ---------------------------------------------------------------------
section("11/12. Bugs #11 & #12 — Telegram report");

test("#11 Category names are HTML-escaped for Telegram", function () {
  setupEnv({ dashboard: [["", "", "", "", "Category", "Budget", "Spent"], ["", "", "", "", "Food & Drinks", 1000, 200]] });
  checkAndNotify();
  ok(env.telegram.length > 0, "a message should be sent");
  var text = env.telegram[0].payload.text;
  ok(text.indexOf("Food &amp; Drinks") > -1, "expected escaped '&amp;' in: " + text);
});
test("#12 checkAndNotify does not crash on an empty Dashboard", function () {
  setupEnv();
  checkAndNotify();
  ok(env.telegram.length > 0, "report should still be sent");
});
test("#12b checkAndNotify sends to all configured TELEGRAM_CHAT_IDS", function () {
  setupEnv({ scriptProperties: { TELEGRAM_CHAT_IDS: "111, 222 , 333" } });
  checkAndNotify();
  eq(env.telegram.map(function (m) { return m.payload.chat_id; }), ["111", "222", "333"], "chat IDs");
});
test("#12c Missing Script Properties fail safely without network calls", function () {
  setupEnv({ scriptProperties: { TELEGRAM_BOT_TOKEN: "" } });
  checkAndNotify();
  eq(env.telegram.length, 0, "no telegram calls without token");

  setupEnv({ scriptProperties: { TELEGRAM_CHAT_IDS: "" } });
  checkAndNotify();
  eq(env.telegram.length, 0, "no telegram calls without chat IDs");

  setupEnv({ scriptProperties: { GEMINI_API_KEY: "" } });
  eq(getGeminiCategory("Some Place"), "Uncategorized");
  eq(env.geminiCalls.length, 0, "no Gemini calls without API key");
});

// ---------------------------------------------------------------------
section("13. Bug #13 — lock handling");

test("#13a Lock timeout → nothing written, error returned", function () {
  setupEnv({ lockBusy: true, categories: [["wolt", "Restaurants"]] });
  var res = send("Wolt", "Payment ₪60");
  eq(txRows().length, 0, "rows written without the lock");
  ok(/^Error/.test(res), "response should be an error, got: " + res);
});
test("#13b Gemini is not called while holding the lock", function () {
  setupEnv({ geminiReply: "Groceries" });
  send("Brand New Store", "Payment ₪70");
  eq(env.geminiCalls.length, 1, "Gemini calls");
  eq(env.geminiCalls[0].lockHeld, false, "lock held during Gemini call");
});
test("#13c Lock is released even if the sheet write throws", function () {
  setupEnv({ categories: [["wolt", "Restaurants"]] });
  env.ss.sheets.Transactions.appendError = new Error("Service Spreadsheets failed");
  try { send("Wolt", "Payment ₪60"); } catch (e) { /* expected */ }
  eq(env.lockHeld, false, "lock still held");
});

// ---------------------------------------------------------------------
section("18. Bug #18 — fixZeroAmounts");

test("#18 fixZeroAmounts picks the real amount (not card digits) as a number", function () {
  setupEnv({ transactions: [[hoursAgo(24), "SMS", 0, "Uncategorized", "בכרטיס 1234 חויבת 50 ₪"]] });
  fixZeroAmounts();
  eq(lastTx().amount, 50);
});

// ---------------------------------------------------------------------
section("19. fixMisparsedMerchants");

test("fixMisparsedMerchants repairs misparsed/truncated merchants and updates category", function () {
  setupEnv({
    transactions: [
      [hoursAgo(24), "אי", 60.71, "Uncategorized", "שלום,\nבכרטיסך המסתיים ב- 1234, אושרה עסקה ב-24/10 בסך 60.71 ש\"ח באי.אם.פי.אם כצנלסון.\nלמידע נוסף..."],
      [hoursAgo(12), "קיי", 49.00, "Uncategorized", "שלום,\nבכרטיסך 1234 אושרה עסקה ב-25/10 בסך 49.00 ש\"ח בקיי.אס.פי מחשבים רמת.\nלמידע נוסף..."]
    ],
    categories: [["קיי.אס.פי", "Electronics"], ["אי.אם.פי.אם", "Groceries"]]
  });
  fixMisparsedMerchants();
  eq(txRows()[0][1], "אי.אם.פי.אם כצנלסון");
  eq(txRows()[0][3], "Groceries");
  eq(txRows()[1][1], "קיי.אס.פי מחשבים רמת");
  eq(txRows()[1][3], "Electronics");
});

// ---------------------------------------------------------------------
section("20. syncAndDeduplicateCategories");

test("syncAndDeduplicateCategories removes newer duplicates and syncs missing merchants", function () {
  setupEnv({
    categories: [
      ["שופרסל", "Groceries"],
      ["י. דראגסטורס - פרישמ", "Health"],
      ["וולט", "Restaurants"],
      ["י דראגסטורס - פרישמ", "Restaurants"], // Newer duplicate of row 2 -> should be deleted!
      ["וולט", "Restaurants"]                 // Newer duplicate of row 3 -> should be deleted!
    ],
    transactions: [
      [hoursAgo(24), "אי.אם.פי.אם כצנלסון", 60.71, "Groceries", "..."],
      [hoursAgo(20), "שופרסל", 120.00, "Groceries", "..."],
      [hoursAgo(12), "קיי.אס.פי מחשבים רמת", 49.00, "Electronics", "..."],
      [hoursAgo(5), "אי.אם.פי.אם כצנלסון", 30.00, "Groceries", "..."],
      [hoursAgo(1), "SMS", 0, "Uncategorized", "..."]
    ]
  });

  syncAndDeduplicateCategories();

  var catData = env.ss.sheets.Categories.rows.slice(1);
  eq(catData.length, 5); // 3 original unique + 2 new from transactions (am:pm, ksp)
  eq(catData[0][0], "שופרסל");
  eq(catData[1][0], "י. דראגסטורס - פרישמ");
  eq(catData[1][1], "Health"); // Preserved older Health category!
  eq(catData[2][0], "וולט");
  eq(catData[3][0], "אי.אם.פי.אם כצנלסון");
  eq(catData[3][1], "Groceries");
  eq(catData[4][0], "קיי.אס.פי מחשבים רמת");
  eq(catData[4][1], "Electronics");
});

// =====================================================================
realLog("\n==========================================");
realLog("📊 Summary: " + passed + " Passed, " + failures.length + " Failed.");
realLog("==========================================");
if (failures.length > 0) process.exit(1);
