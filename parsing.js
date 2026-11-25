/**
 * Pure parsing / classification helpers shared by doGet, the maintenance tools
 * (fixAdTransactions, fixZeroAmounts, ...) and the Node test suite (tests.js).
 * Nothing in this file calls Apps Script services.
 *
 * NOTE: Top-level declarations use `var` (not const/let) so they are visible
 * from every script file regardless of Apps Script's file load order.
 */

// Categories the AI is allowed to assign (keep in sync with the Dashboard / typed Category column)
var CATEGORIES = ["Groceries", "Restaurants", "Transport", "Bills", "Clothing", "Health",
  "Entertainment", "Home Improvement", "Beauty", "Vacation", "Cash Withdraw"];

// --- AD FILTER KEYWORDS ---

// Hebrew ad keywords are matched as substrings (Hebrew attaches prefixes like ה/ו/ל to words)
var AD_KEYWORDS_HE = [
  "הנחה", "קבלו", "מבצע", "הצטרפו", "מצטרפים", "קופון", "שובר", "הטבה", "הטבות",
  "אל תפספסו", "לא תפספסו", "להזמנה", "לפרטים", "בתוקף עד", "עד גמר המלאי",
  "הירשמו", "חינם", "מתנה", "לרכישה", "לרוכשים", "למזמינים",
  "במקום", "בלבד", "להסרה", "כנסו", "היכנסו", "להציץ", "הציצו",
  "מחירי מועדון", "חברי מועדון", "הצטרפו למועדון", "דאגנו לכם",
  "מטורף", "שווים", "בלעדי", "סייל", "חיסול", "תקנון"
];

// English ad keywords: whole words, case-insensitive ("SALE" matches, "Wholesale" doesn't)
var AD_KEYWORDS_EN = /\b(?:sales?|promo\w*|discounts?)\b|%\s*off\b/i;

// Hard ad patterns: unambiguous marketing/opt-out footers & credit disclaimers that
// NEVER appear in real purchase receipts — cannot be bypassed by safety words.
var HARD_AD_PATTERNS = [
  /להסרה\s*(?:מרשימת התפוצה|השב\s+הסר|יש\s+לשלוח\s+הסר|שלחו?\s+הסר)/,
  /יש\s+לשלוח\s+הסר/,
  /אי[־\-\s]?עמידה\s+בפירעון/,
  /לתקנון\s+המבצע/,
  /למצטרפים\s+חדשים/
];

// Safety check: an ad keyword is ignored if the text also looks like a real receipt/transaction.
// Must contain an actual charge/approval verb (never bare "בכרטיס", which matches promo "בכרטיסי").
var TX_SAFETY_WORDS = [
  "shulam", "paid", "Payment", "payment",
  "חוייב", "חויב", "חוייבת", "חויבת", "אושרה", "עסקה",
  "שולם", "שילמת", "הסיבוס שלך", "משיכת מזומן"
];

// Real bank/card SMSes ALWAYS contain at least one transaction keyword
var TX_KEYWORDS = [
  "עסקה", "חוייב", "חויב", "חוייבת", "חויבת", "לחיוב",
  "בכרטיסך", "אושרה", "שולם", "שילמת", "בסך", "על סך", "בסכום", "הסיבוס שלך", "משיכת מזומן",
  "Total", "Amount", "charged", "paid", "Payment", "payment"
];

// --- CURRENCIES ---
// To add a currency: add its tokens here and its code to FOREIGN_CURRENCIES
// (and to `foreignCurrencies` in fixUnconvertedForeignTransactions).
var CURRENCY_PATTERNS = {
  USD: "(?:USD|\\$)",
  EUR: "(?:EURO?|€)",
  CZK: "(?:CZK|Kč)",
  GBP: "(?:GBP|£)",
  ILS: "(?:ILS|NIS|ש\"?ח|₪)"
};
var FOREIGN_CURRENCIES = ["USD", "EUR", "CZK", "GBP"];

// A number with optional thousands/decimal separators: "50", "1,234.56", "1.234,56", "250,00"
var NUMBER_PATTERN = "\\d(?:[\\d.,]*\\d)?";

function containsAny_(text, words) {
  for (var i = 0; i < words.length; i++) {
    if (text.indexOf(words[i]) > -1) return true;
  }
  return false;
}

/**
 * Two-layer ad filter. Returns null for a (likely) real transaction, or
 * { response, detail } describing why the text should be rejected.
 *   Layer 0: Unambiguous marketing opt-out / credit-offer disclaimer (hard block)
 *   Layer 1: Ad keyword blocklist (skipped if a transaction safety word is present)
 *   Layer 2: SMS structural check — an SMS with no transaction keyword is promotional
 */
function findAdReason(rawText, isSMS) {
  var text = String(rawText);

  // Layer 0: Unambiguous opt-out / legal credit-marketing disclaimers
  for (var h = 0; h < HARD_AD_PATTERNS.length; h++) {
    if (HARD_AD_PATTERNS[h].test(text)) {
      return { response: "Ignored: Ad", detail: "Hard ad pattern matched: " + HARD_AD_PATTERNS[h].source };
    }
  }

  // Layer 1: Keyword blocklist
  var adHit = null;
  for (var k = 0; k < AD_KEYWORDS_HE.length && !adHit; k++) {
    if (text.indexOf(AD_KEYWORDS_HE[k]) > -1) adHit = AD_KEYWORDS_HE[k];
  }
  if (!adHit) {
    var enMatch = text.match(AD_KEYWORDS_EN);
    if (enMatch) adHit = enMatch[0];
  }
  if (adHit && !containsAny_(text, TX_SAFETY_WORDS)) {
    return { response: "Ignored: Ad", detail: "Ad Keyword found: " + adHit };
  }

  // Layer 2: SMS structural filter
  if (isSMS && !containsAny_(text, TX_KEYWORDS) && !/בכרטיס\s+(?:המסתיים|\d{4})/.test(text)) {
    return { response: "Ignored: SMS Ad", detail: "SMS with no transaction keywords (likely ad/promo)" };
  }
  return null;
}

/** True when the notification/SMS text represents an ATM / cash withdrawal. */
function isCashWithdrawal(rawText) {
  var text = String(rawText || "");
  return /משיכת\s+מזומ(?:ן|נים)|משיכה\s+מכספומט|\b(?:cash|atm)\s+withdrawal\b/i.test(text);
}

/** Identifies the credit card issuer from sender name or SMS body for ATM withdrawals. */
function extractCardIssuer(rawMerchant, rawText) {
  var sender = String(rawMerchant || "").trim();
  var text = String(rawText || "");
  var sources = [sender, text];
  for (var s = 0; s < sources.length; s++) {
    var src = sources[s];
    if (!src) continue;
    if (/\bamex\b|american\s*express|אמריקן\s*אקספרס/i.test(src)) return "Amex";
    if (/\bdiners\b|דיינרס/i.test(src)) return "Diners";
    if (/\bcal\b|cal-online|כאל/i.test(src)) return "Cal";
    if (/\bmax\b|max\.co\.il|לאומי\s*קארד|מקס/i.test(src)) return "MAX";
    if (/\bisracard\b|ישראכרט/i.test(src)) return "Isracard";
  }
  if (sender && !isGenericMerchant(sender)) return sender;
  return "ATM";
}

/**
 * Extracts and cleans the merchant name from the notification sender + text.
 */
function parseMerchant(rawMerchant, rawText) {
  var merchant = String(rawMerchant || "");
  var text = String(rawText || "").replace(/[\u200B-\u200F\u202A-\u202E\u061C]/g, " ");
  if (isCashWithdrawal(text)) {
    return extractCardIssuer(merchant, text);
  }
  var isCibus = text.indexOf("הסיבוס שלך") > -1;

  // A. Cibus / Pluxee Logic
  if (isCibus) {
    var cibusMatch = text.match(/עבור עסקה במסעדת (.*?) הסיבוס שלך/);
    if (cibusMatch && cibusMatch[1]) {
      merchant = cibusMatch[1];
    }
  }

  // B. Google Wallet vs SMS Logic
  // Wallet notifications already carry the real merchant in rawMerchant, while their text
  // ("עסקה בסך 15.00 ₪ בכרטיס Mastercard ... ••1234") can contain "בכרטיס" / "בסך ... ב...".
  var isWalletNotification = !isGenericMerchant(merchant) &&
    (/••\s*[\dX]{4}/i.test(text) || /^עסקה בסך [^\r\n]+ בכרטיס /.test(text));
  // Check if it's an SMS (sender is generic/card issuer/phone number, OR text contains "בכרטיס")
  var isSMS = !isWalletNotification && (isGenericMerchant(merchant) || text.indexOf("בכרטיס") > -1);

  // If it is an SMS (and not Cibus), we aggressively hunt for the merchant name.
  // FIRST MATCH WINS, most specific pattern first. (Previously every later pattern
  // overrode the earlier ones, so "... בסך 499 ש"ח בתאריך 01/10" became "תאריך 01/10".)
  if (isSMS && !isCibus) {
    var patterns = [
      // Attempt 0: "המסתיים ב-? 1234 ב-?[Merchant]\nבסך" (Cal multiline/inline card-ending format)
      /המסתיים\s+ב-?\s*\d{4}\s+ב-?\s*(?!\d{1,2}[./]\d{1,2})([^\r\n]+?)\s+(?:בסך|על סך)/,
      // Attempt 1: "לחיוב ב-[Merchant]... על סך" (Merchant FIRST)
      /לחיוב ב-(.*?)\s+על סך/,
      // Attempt 2: "עסקה/עסקת [...] ב-[Merchant] בסך" (Merchant MIDDLE - Cal/Max, allows newlines)
      // Negative lookahead skips transaction dates ("ב-14/10") and card digits ("ב-1234 ב...")
      /עסק[הת]\s+(?:[\s\S]*?\s+)?ב-(?!\s*(?:\d{1,2}[./]\d{1,2}|\d{4}\s+ב))([^\r\n]+?)\s+(?:בסך|על סך)/,
      // Attempt 3: "בסך/בסכום ... ב[Merchant]" (Merchant LAST - Isracard, incl. 3DS verification SMSes)
      // Matches until the end of the sentence/line, without stopping at internal dots within acronyms
      /(?:בסך|בסכום) .*? ב(?!-?\s*(?:כרטיס(?:ך|\s|$)|תאריך(?:\s|$)))(?:-?\s*)(.*?)(?:\.\s*(?:\r?\n|$|למידע|לשירותך)|,\s*(?:\r?\n|$)|\r?\n|$)/
    ];
    for (var p = 0; p < patterns.length; p++) {
      var m = text.match(patterns[p]);
      if (m) {
        merchant = m[1];
        break;
      }
    }
  }

  // C. Universal Cleaner
  return merchant
    .replace("Payment to ", "").replace("You paid ", "")
    .replace("עבור עסקה במסעדת ", "")
    .replace("עסקה במסעדת ", "")
    .replace("עסקה ב-", "").replace("עסקה ב", "")
    .replace("לחיוב ב-", "").replace("לחיוב ב", "")
    .replace("שילמת ב-", "").replace("שילמת ב", "")
    .replace(/^-\s*/, "") // Remove leading hyphen from "ב-[Merchant]" captured by Attempt 3
    .split(" הסיבוס שלך")[0]
    .split(" בכרטיס")[0]
    .split(" על סך")[0]
    .split(" למידע נוסף")[0]
    .split(" למידע")[0]
    .replace(/[.,]+$/, "") // Remove trailing sentence dot or comma
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Parses a number string using either "1,234.56" or "1.234,56" conventions.
 *   - Both separators present: the LAST one is the decimal separator.
 *   - Only commas: a single comma followed by 1-2 digits is decimal ("250,00"),
 *     otherwise commas are thousands separators ("1,234").
 *   - Only dots: a single dot is decimal ("52.94"), several are thousands ("1.234.567").
 * Returns NaN for anything that isn't a number.
 */
function parseLocaleNumber(str) {
  var s = String(str).replace(/\s/g, "");
  if (!/^\d[\d.,]*$/.test(s)) return NaN;

  var lastComma = s.lastIndexOf(",");
  var lastDot = s.lastIndexOf(".");
  var decimalSep = null;

  if (lastComma > -1 && lastDot > -1) {
    decimalSep = lastComma > lastDot ? "," : ".";
  } else if (lastComma > -1) {
    decimalSep = (s.split(",").length === 2 && /,\d{1,2}$/.test(s)) ? "," : null;
  } else if (lastDot > -1) {
    decimalSep = s.split(".").length === 2 ? "." : null;
  }

  if (!decimalSep) return Number(s.replace(/[.,]/g, ""));
  var idx = s.lastIndexOf(decimalSep);
  return Number(s.slice(0, idx).replace(/[.,]/g, "") + "." + s.slice(idx + 1));
}

/**
 * Finds the transaction amount and its currency.
 *   Level 1: Keywords ("Charged $50", "בסך 3960.00 EUR")
 *   Level 2: Explicit currency ("$50", "50 EUR")
 *   No fallback: without a keyword or currency the amount is 0 (rejects promo spam).
 * Returns { amount, currency, rawNum } — amount is 0 when nothing valid was found.
 */
function extractAmount(rawText) {
  // --- 👻 THE FIX: Remove invisible text-direction marks (RLM, LRM) ---
  // This cleans characters \u200B through \u200F (Zero width spaces, direction marks)
  var cleanText = String(rawText).replace(/[\u200B-\u200F\u202A-\u202E\u061C]/g, " ");

  var allCurr = "(?:" + [CURRENCY_PATTERNS.USD, CURRENCY_PATTERNS.EUR, CURRENCY_PATTERNS.CZK,
    CURRENCY_PATTERNS.GBP, CURRENCY_PATTERNS.ILS].join("|") + ")";

  // Level 1: Keywords
  var keywordRegex = new RegExp("(?:בסך|על סך|בסכום|חוייב ב|Amount|Sum|Total)\\s*(" + allCurr + "?\\s*" + NUMBER_PATTERN + "\\s*" + allCurr + "?)", "i");
  // Level 2: Explicit Currency
  var currencyRegex = new RegExp("(\\s*" + NUMBER_PATTERN + "\\s*" + allCurr + "|" + allCurr + "\\s*" + NUMBER_PATTERN + ")", "i");

  var match = cleanText.match(keywordRegex) || cleanText.match(currencyRegex);
  if (!match) return { amount: 0, currency: "ILS", rawNum: "" };

  var rawNum = match[1];
  var currency = "ILS";
  for (var c = 0; c < FOREIGN_CURRENCIES.length; c++) {
    if (new RegExp(CURRENCY_PATTERNS[FOREIGN_CURRENCIES[c]], "i").test(rawNum)) {
      currency = FOREIGN_CURRENCIES[c];
      break;
    }
  }

  var numberMatch = rawNum.match(new RegExp(NUMBER_PATTERN));
  var amount = numberMatch ? parseLocaleNumber(numberMatch[0]) : NaN;
  return { amount: amount > 0 ? amount : 0, currency: currency, rawNum: rawNum };
}

/**
 * True for merchant names that don't identify a merchant (unparsed "SMS",
 * card-issuer sender names, masked card labels, phone-number senders, empty).
 * These must never be sent to the AI or cached.
 */
function isGenericMerchant(merchant) {
  var m = String(merchant || "").trim();
  return m.length < 2
    || /^(?:sms|isracard|cal|max|amex|diners|atm|ישראכרט|כאל|מקס|אמריקן\s*אקספרס|דיינרס)$/i.test(m)
    || /^[\d\s+\-()./]+$/.test(m)
    || /^כרטיס\s/i.test(m)
    || /••\s*[\dX]{4}/i.test(m);
}

function normalizeRawTextForMatch_(text) {
  return String(text || "")
    .replace(/\s*\[(?:Orig|UNCONVERTED):[^\]]*\]/g, "")
    .replace(/[\u200B-\u200F\u202A-\u202E\u061C]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Recovers a lost merchant name (e.g. from Google Wallet notifications whose merchant
 * was only in e.parameter.merchant, not Raw Text) by matching against parsed Debug rows
 * ({ date, merchant, text, normText }). Prefers the closest timestamp within [-5s, +120s];
 * falls back outside that window ONLY if all matching Debug rows agree on one unique merchant.
 */
function findMerchantInDebug(txDate, txRawText, debugEntries) {
  if (!debugEntries || debugEntries.length === 0) return null;
  var normTx = normalizeRawTextForMatch_(txRawText);
  if (!normTx) return null;

  var bestCandidate = null;
  var bestAbsDiff = Infinity;
  var uniqueFallback = null;
  var fallbackAmbiguous = false;
  var hasTxDate = (txDate instanceof Date && !isNaN(txDate.getTime()));

  for (var i = 0; i < debugEntries.length; i++) {
    var entry = debugEntries[i];
    if (entry.normText !== normTx || !entry.merchant) continue;

    var candidate = parseMerchant(entry.merchant, entry.text);
    if (!candidate || isGenericMerchant(candidate)) continue;

    // Track unique fallback candidate (case/punctuation-normalized)
    if (!fallbackAmbiguous) {
      if (uniqueFallback === null) {
        uniqueFallback = candidate;
      } else if (normalizeMerchant(uniqueFallback) !== normalizeMerchant(candidate)) {
        fallbackAmbiguous = true;
        uniqueFallback = null;
      }
    }

    if (hasTxDate && entry.date instanceof Date && !isNaN(entry.date.getTime())) {
      var diffMs = txDate.getTime() - entry.date.getTime();
      if (diffMs >= -5000 && diffMs <= 120000) {
        var absDiff = Math.abs(diffMs);
        if (absDiff < bestAbsDiff) {
          bestAbsDiff = absDiff;
          bestCandidate = candidate;
        }
      }
    }
  }

  return bestCandidate || (!fallbackAmbiguous ? uniqueFallback : null);
}

/**
 * Normalizes a merchant name for exact full-name matching:
 * Removes invisible Unicode marks, quotes, internal abbreviation dots,
 * converts hyphens to spaces, collapses whitespace, and lowercases.
 */
function normalizeMerchant(name) {
  return String(name || "")
    // Remove invisible Unicode direction marks (RLM, LRM, zero-width spaces)
    .replace(/[\u200B-\u200F\u202A-\u202E\u061C]/g, "")
    .toLowerCase()
    // Remove all quotes, apostrophes, gershayim, smart quotes: ' " ` ´ ״ ׳ “ ” ‘ ’
    .replace(/['"`´״׳“”‘’]+/g, "")
    // Remove all dots (abbreviations like י. -> י, co. -> co, etc.)
    .replace(/\./g, "")
    // Convert hyphens and dashes to spaces: - – —
    .replace(/[\-–—_]+/g, " ")
    // Remove leading/trailing punctuation
    .replace(/^[.,;:\s]+|[.,;:\s]+$/g, "")
    // Collapse multiple spaces to single space
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Checks if an incoming merchant matches a cached merchant name:
 * 1. Exact match after normalization
 * 2. Base chain match: incoming is "י דראגסטורס - פרישמ", cached is "י דראגסטורס"
 * 3. Truncation match: terminal cut 1-2 chars at the end (e.g. פרישמ vs פרישמן)
 */
function isMerchantMatch(incomingRaw, cachedRaw) {
  var cleanIncoming = normalizeMerchant(incomingRaw);
  var cleanCached = normalizeMerchant(cachedRaw);

  if (!cleanCached || !cleanIncoming) return false;

  // 1. Exact full-name match
  if (cleanIncoming === cleanCached) return true;

  // 2. Base chain match:
  // a) incoming has branch with dash: "י דראגסטורס - פרישמ", cached is "י דראגסטורס"
  var parts = String(incomingRaw || "").split(/\s*[-–—]\s*/);
  if (parts.length > 1) {
    var cleanBase = normalizeMerchant(parts[0]);
    if (cleanBase && cleanBase === cleanCached) return true;
  }

  // b) incoming starts with the cached chain name followed by space:
  // e.g. "אי.אם.פי.אם כצנלסון" starts with "אי.אם.פי.אם "
  // Guard: cleanCached must be at least 5 chars or have multiple words (avoids short generic words like 'קפה' / 'בר')
  if ((cleanCached.indexOf(" ") > -1 || cleanCached.length >= 5) &&
      cleanIncoming.indexOf(cleanCached + " ") === 0) {
    return true;
  }

  // 3. Truncation match: terminal cut 1-2 characters at end (e.g. פרישמ vs פרישמן)
  if (cleanIncoming.length >= 8 && cleanCached.startsWith(cleanIncoming) && (cleanCached.length - cleanIncoming.length) <= 2) return true;
  if (cleanCached.length >= 8 && cleanIncoming.startsWith(cleanCached) && (cleanIncoming.length - cleanCached.length) <= 2) return true;

  return false;
}

var CATEGORY_ALIASES = {
  "restaurant": "Restaurants",
  "grocery": "Groceries",
  "bill": "Bills",
  "cash withdrawal": "Cash Withdraw",
  "cash withdrawals": "Cash Withdraw",
  "atm": "Cash Withdraw"
};

function matchSingleCategory_(line) {
  var t = String(line || "")
    .replace(/\[\d+(?:\s*,\s*\d+)*\]/g, "") // Strip Google Search citation markers like [1]
    .replace(/["'`{}*]/g, "")
    .replace(/^\s*category\s*:\s*/i, "")
    .replace(/[.,!;:\s]+$/, "")
    .trim()
    .toLowerCase();
  for (var i = 0; i < CATEGORIES.length; i++) {
    if (CATEGORIES[i].toLowerCase() === t) return CATEGORIES[i];
  }
  if (CATEGORY_ALIASES[t]) return CATEGORY_ALIASES[t];
  return null;
}

/**
 * Maps a free-text AI answer onto one of CATEGORIES (case-insensitive,
 * tolerating quotes / JSON / trailing punctuation / search citations / singular forms).
 * Anything else → "Uncategorized".
 */
function normalizeCategory(text) {
  var raw = String(text || "").trim();
  var direct = matchSingleCategory_(raw);
  if (direct) return direct;

  // If Google Search grounding returned multiple lines, check the last non-empty line
  var lines = raw.split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
  if (lines.length > 1) {
    var lastLineMatch = matchSingleCategory_(lines[lines.length - 1]);
    if (lastLineMatch) return lastLineMatch;
  }
  return "Uncategorized";
}

/** Escapes text for Telegram's HTML parse_mode. */
function escapeHtml(text) {
  return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Date -> "yyyy-MM-dd" without any API calls (cheap enough for big loops). */
function dateToKey(d) {
  var y = d.getFullYear();
  var m = d.getMonth() + 1;
  var day = d.getDate();
  return y + "-" + (m < 10 ? "0" : "") + m + "-" + (day < 10 ? "0" : "") + day;
}
