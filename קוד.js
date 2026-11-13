function doGet(e) {
  // --- DEBUG SPY START ---
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var debugSheet = ss.getSheetByName("Debug");
  if (!debugSheet) {
    debugSheet = ss.insertSheet("Debug");
    debugSheet.appendRow(["Date", "Parameters"]);
  }

  var params = JSON.stringify(e.parameter);
  debugSheet.appendRow([new Date(), params]);
  // --- DEBUG SPY END ---

  console.log("--- NEW REQUEST RECEIVED ---");
  console.log("Params:", JSON.stringify(e.parameter));

  var sheet = ss.getSheetByName("Transactions");

  // 1. Get Data
  var rawMerchant = e.parameter.merchant || "SMS";
  var rawText = e.parameter.text || "";

  // --- SPAM FILTER ---
  if (rawMerchant.indexOf("Preparing your receipt") > -1) {
    console.log("❌ Rejected: 'Preparing receipt' detected");
    return ContentService.createTextOutput("Ignored: Temporary");
  }

  // --- AD FILTER (Two Layers, see findAdReason in parsing.js) ---
  var isSMS = (rawMerchant === "SMS" || isGenericMerchant(rawMerchant) || /sms|messages?/i.test(rawMerchant));
  var adReason = findAdReason(rawText, isSMS);
  if (adReason) {
    console.log("❌ Rejected: " + adReason.detail);
    return ContentService.createTextOutput(adReason.response);
  }

  // --- 3. PARSE SMS & CLEAN MERCHANT ---
  var merchant = parseMerchant(rawMerchant, rawText);
  console.log("Cleaned Merchant: " + merchant);

  // --- 4. EXTRACT AMOUNT & CONVERT CURRENCY ---
  var parsed = extractAmount(rawText);
  var amount = parsed.amount;

  if (!(amount > 0)) { // Also catches NaN
    console.log("❌ Rejected: No amount found (Hidden chars likely)");
    return ContentService.createTextOutput("Ignored: No Amount");
  }

  if (parsed.currency !== "ILS") {
    var converted = convertToILS(amount, parsed.currency);
    if (converted !== null) {
      amount = converted;
      rawText += " [Orig: " + parsed.rawNum.trim() + "]";
    } else {
      // Keep the foreign amount but flag it, so fixUnconvertedForeignTransactions repairs it later
      rawText += " [UNCONVERTED: " + parsed.rawNum.trim() + "]";
    }
  }

  // --- 5. CATEGORIZE (before taking the lock — AI calls are slow) ---
  var category = "Uncategorized";
  var shouldCache = false;
  var catSheet = ss.getSheetByName("Categories");

  // A. Manual Lookup (Cache First)
  // We check if we already know this merchant to save AI tokens and time.
  if (catSheet && catSheet.getLastRow() > 1) {
    var catData = catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 2).getValues();

    for (var c = 0; c < catData.length; c++) {
      if (isMerchantMatch(merchant, catData[c][0])) {
        category = catData[c][1];
        console.log("✅ Found in Cache: " + category);
        break;
      }
    }
  }

  // B. AI Lookup (If still unknown). Generic names ("SMS", phone numbers) are
  // never sent to the AI or cached — a cached "SMS" would match every unparsed SMS.
  if (category === "Uncategorized" && !isGenericMerchant(merchant)) {
    console.log("🤔 Unknown merchant. Asking Gemini...");
    var aiCategory = getGeminiCategory(merchant);
    if (aiCategory !== "Uncategorized") {
      category = aiCategory;
      shouldCache = true;
    }
  }

  // --- 6. DEDUPLICATION + SAVE (Critical Section) ---
  // Lock prevents race conditions between the duplicate webhooks MacroDroid fires
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    // Never write without the lock — the duplicate check would be meaningless
    console.log("❌ Could not acquire lock (Busy): " + err.toString());
    return ContentService.createTextOutput("Error: Busy, retry later");
  }

  var savedRow;
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // Look back at the last 15 rows
      var startRow = Math.max(2, lastRow - 14);
      // Fetch Date (0), Amount (2), and Raw Text (4) — Cols A to E
      var recentData = sheet.getRange(startRow, 1, lastRow - startRow + 1, 5).getValues();
      var now = new Date();

      for (var i = 0; i < recentData.length; i++) {
        var rDate = new Date(recentData[i][0]);
        var rAmount = Number(recentData[i][2]);
        var rRawText = recentData[i][4]; // Column E

        var diffSeconds = Math.abs(now - rDate) / 1000;
        var isAmountMatch = Math.abs(rAmount - amount) < 0.01;

        // RULE A: Machine Gun (Technical Glitch)
        // Same Amount + Arrived within 60 seconds
        if (isAmountMatch && diffSeconds < 60) {
          console.log("❌ Rejected: Rapid Duplicate (Row " + (startRow + i) + ")");
          return ContentService.createTextOutput("Ignored: Rapid Duplicate");
        }

        // RULE B: Ghost Replay (Notification Bug)
        // Identical Text + Arrived within 12 Hours
        // (Legitimate repeat purchases usually have different timestamps/IDs in text)
        if (isAmountMatch && rawText === rRawText && diffSeconds < 43200) {
          console.log("❌ Rejected: Ghost Replay (Row " + (startRow + i) + ")");
          return ContentService.createTextOutput("Ignored: Ghost Replay");
        }
      }
    }

    sheet.appendRow([new Date(), merchant, amount, category, rawText]);
    SpreadsheetApp.flush(); // Force write to sheet immediately
    savedRow = sheet.getLastRow();

    // Add to cache so next time we skip the AI
    if (shouldCache) addToMerchantCache(merchant, category);
  } finally {
    lock.releaseLock(); // Always release, even if a sheet call throws
  }

  console.log("✅ SUCCESS: Saved to row " + savedRow);
  return ContentService.createTextOutput("Success");
}

// --- NEW TOOLS ---

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Wallet Tools')
    .addItem('Update All Categories', 'reprocessCategories')
    .addItem('Sync & Deduplicate Categories', 'syncAndDeduplicateCategories')
    .addItem('Fix Misparsed Merchants', 'fixMisparsedMerchants')
    .addItem('Fix Zero Amounts', 'fixZeroAmounts')
    .addItem('Fix CZK Transactions (June 12-17)', 'fixCZKTransactions')
    .addItem('Remove SMS Ad Transactions', 'fixAdTransactions')
    .addItem('Fix Unconverted Foreign Transactions', 'fixUnconvertedForeignTransactions')
    .addToUi();
}

function reprocessCategories() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transactions");
  var catSheet = ss.getSheetByName("Categories");

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No transactions found.');
    return;
  }

  // --- OPTIMIZATION START ---
  // Only look at the last 15 rows (User assumption: max 10 uncategorized)
  var SEARCH_DEPTH = 15;
  var startRow = Math.max(2, lastRow - SEARCH_DEPTH + 1);
  var numRows = lastRow - startRow + 1;

  // Get range: Col B (Merchant) to Col D (Category) for the bottom rows only
  var range = sheet.getRange(startRow, 2, numRows, 3);
  var data = range.getValues();
  // --- OPTIMIZATION END ---

  // Load Cache (Entire cache is fine, it's usually small)
  var catData = [];
  if (catSheet && catSheet.getLastRow() > 1) {
    catData = catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 2).getValues();
  }

  var updatesCount = 0;
  var aiCallsCount = 0;

  for (var i = 0; i < data.length; i++) {
    var merchant = data[i][0].toString();
    var currentCat = data[i][2].toString();

    // Process only if Uncategorized/Empty
    if (currentCat === "Uncategorized" || currentCat === "") {

      var newCategory = "Uncategorized";

      // 1. FAST: Check Cache First
      for (var j = 0; j < catData.length; j++) {
        if (isMerchantMatch(merchant, catData[j][0])) {
          newCategory = catData[j][1];
          break;
        }
      }

      // 2. SLOW: Ask AI (Only if Cache failed)
      if (newCategory === "Uncategorized") {
        if (aiCallsCount >= 5) { // Safety limit for one run
          console.log("⚠️ Batch limit reached. Run again for more.");
          break;
        }

        try {
          newCategory = getGeminiCategory(merchant);
          aiCallsCount++;

          // Sleep 4s to stay safe (Free Tier = 15 req/min)
          Utilities.sleep(4000);

          if (newCategory !== "Uncategorized") {
            addToMerchantCache(merchant, newCategory);
            // Add to local memory so next row uses it instantly
            catData.push([merchant, newCategory]);
          }
        } catch (e) {
          console.log("AI Failed for " + merchant + ": " + e.toString());
        }
      }

      // 3. Update data array if we found something
      if (newCategory !== "Uncategorized") {
        data[i][2] = newCategory;
        updatesCount++;
      }
    }
  }

  // 4. Write back ONLY the bottom rows
  if (updatesCount > 0) {
    range.setValues(data);
    SpreadsheetApp.getUi().alert('Success! Updated ' + updatesCount + ' recent transactions.');
  } else {
    SpreadsheetApp.getUi().alert('No new matches found in the last ' + SEARCH_DEPTH + ' rows.');
  }
}

/**
 * Retroactive fix: Scans all transactions and re-runs parseMerchant on the
 * raw text (Column E) to repair any merchant names that were truncated or
 * misparsed due to acronym dots (e.g., "אי" -> "אי.אם.פי.אם כצנלסון",
 * "קיי" -> "קיי.אס.פי מחשבים רמת", or "14/10" -> "ד"ר בוקסנבאום בע"מ").
 * Also updates the Category if the repaired merchant matches a category in Categories.
 */
function fixMisparsedMerchants() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transactions");
  var catSheet = ss.getSheetByName("Categories");

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No transactions found.');
    return;
  }

  // Load Categories sheet for instant re-categorization
  var catData = [];
  if (catSheet && catSheet.getLastRow() > 1) {
    catData = catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 2).getValues();
  }

  // Cols A-E: Date(0), Merchant(1), Amount(2), Category(3), Raw Text(4)
  var range = sheet.getRange(2, 1, lastRow - 1, 5);
  var data = range.getValues();

  var updatesCount = 0;

  for (var i = 0; i < data.length; i++) {
    var currentMerchant = String(data[i][1] || "").trim();
    var currentCat = String(data[i][3] || "").trim();
    var rawText = String(data[i][4] || "").trim();

    if (!rawText) continue;

    var newMerchant = parseMerchant(currentMerchant, rawText);

    // If parseMerchant with currentMerchant didn't change it, but currentMerchant is generic ("SMS", "Isracard"):
    if (newMerchant === currentMerchant && (currentMerchant === "SMS" || currentMerchant === "Isracard")) {
      newMerchant = parseMerchant("SMS", rawText);
    }

    if (newMerchant && newMerchant !== currentMerchant) {
      console.log("Fixing row " + (i + 2) + ": '" + currentMerchant + "' -> '" + newMerchant + "'");
      data[i][1] = newMerchant;
      updatesCount++;

      // Check if we can also match/fix the category from cache
      for (var c = 0; c < catData.length; c++) {
        if (isMerchantMatch(newMerchant, catData[c][0])) {
          data[i][3] = catData[c][1];
          console.log("  -> Updated category to: " + catData[c][1]);
          break;
        }
      }
    }
  }

  if (updatesCount > 0) {
    range.setValues(data);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert('Fixed ' + updatesCount + ' merchant names in Transactions!');
  } else {
    SpreadsheetApp.getUi().alert('No misparsed merchant names found to fix.');
  }
}

/**
 * Maintenance Tool:
 * 1. Identifies and removes duplicate entries in the 'Categories' sheet,
 *    keeping the older/earlier entry and deleting the newer duplicates.
 * 2. Scans the 'Transactions' sheet to ensure every non-generic merchant has
 *    a corresponding entry in 'Categories'. If missing, it adds it using the
 *    transaction's category (or querying Gemini AI / setting Uncategorized).
 */
function syncAndDeduplicateCategories() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var catSheet = ss.getSheetByName("Categories");
  var txSheet = ss.getSheetByName("Transactions");

  if (!catSheet) {
    SpreadsheetApp.getUi().alert('Categories sheet not found.');
    return;
  }

  // --- PART 1: Deduplicate Categories sheet (keep older, delete newer) ---
  var catLastRow = catSheet.getLastRow();
  var dupesRemoved = 0;
  var cleanCatList = []; // Array of [merchant, category]

  if (catLastRow > 1) {
    var catRange = catSheet.getRange(2, 1, catLastRow - 1, 2);
    var catData = catRange.getValues();

    var seenMerchants = {};
    var rowsToDelete = []; // 1-indexed sheet rows

    for (var i = 0; i < catData.length; i++) {
      var mName = catData[i][0];
      var clean = normalizeMerchant(mName);
      if (!clean) continue;

      var sheetRowNum = i + 2; // header is row 1

      if (seenMerchants[clean]) {
        rowsToDelete.push(sheetRowNum);
        console.log("Found duplicate category at row " + sheetRowNum + ": '" + mName + "'");
      } else {
        seenMerchants[clean] = true;
        cleanCatList.push([mName, catData[i][1]]);
      }
    }

    // Delete rows bottom-up to prevent index shifting
    if (rowsToDelete.length > 0) {
      for (var r = rowsToDelete.length - 1; r >= 0; r--) {
        catSheet.deleteRow(rowsToDelete[r]);
      }
      dupesRemoved = rowsToDelete.length;
      console.log("Removed " + dupesRemoved + " duplicate category rows.");
    }
  }

  // --- PART 2: Sync missing merchants from Transactions to Categories ---
  var addedCount = 0;
  var aiCallsCount = 0;

  if (txSheet && txSheet.getLastRow() > 1) {
    var txLastRow = txSheet.getLastRow();
    // Fetch Merchant (Col B, index 1) and Category (Col D, index 3)
    var txData = txSheet.getRange(2, 1, txLastRow - 1, 5).getValues();

    for (var t = 0; t < txData.length; t++) {
      var txMerchant = String(txData[t][1] || "").trim();
      var txCat = String(txData[t][3] || "").trim();

      // Skip empty or generic placeholders like "SMS", phone numbers, etc.
      if (!txMerchant || isGenericMerchant(txMerchant)) continue;

      // Check if this merchant already matches any entry in cleanCatList
      var exists = false;
      for (var c = 0; c < cleanCatList.length; c++) {
        if (isMerchantMatch(txMerchant, cleanCatList[c][0])) {
          exists = true;
          break;
        }
      }

      if (!exists) {
        var finalCat = (txCat && txCat !== "Uncategorized") ? txCat : "Uncategorized";

        // If transaction has no category, try Gemini (up to 5 calls safety limit)
        if (finalCat === "Uncategorized" && aiCallsCount < 5) {
          try {
            var aiCat = getGeminiCategory(txMerchant);
            aiCallsCount++;
            if (aiCat !== "Uncategorized") finalCat = aiCat;
          } catch (e) {
            console.log("AI error: " + e.toString());
          }
        }

        catSheet.appendRow([txMerchant, finalCat]);
        cleanCatList.push([txMerchant, finalCat]); // Update in-memory so future rows don't re-add
        addedCount++;
        console.log("Added missing merchant to Categories: '" + txMerchant + "' -> '" + finalCat + "'");
      }
    }
  }

  if (dupesRemoved > 0 || addedCount > 0) {
    SpreadsheetApp.flush();
  }

  var msg = "Categories Sync Complete!\n\n" +
    "• Duplicate categories removed: " + dupesRemoved + "\n" +
    "• Missing merchants added from Transactions: " + addedCount;
  SpreadsheetApp.getUi().alert(msg);
}

function fixZeroAmounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transactions");

  // 1. Get all data (Rows 2 to Last)
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Cols A-E: Amount is index 2, Raw Text is index 4
  var range = sheet.getRange(2, 1, lastRow - 1, 5);
  var data = range.getValues();

  var updatesCount = 0;

  // 2. Loop through rows, using the same extractAmount as doGet
  for (var i = 0; i < data.length; i++) {
    var currentAmount = data[i][2]; // Column C
    var rawText = String(data[i][4]); // Column E (Raw Text)

    // Only fix if Amount is 0 or empty, AND we have raw text to read
    if ((currentAmount == 0 || currentAmount === "") && rawText !== "") {
      var parsed = extractAmount(rawText);
      if (parsed.amount > 0) {
        data[i][2] = parsed.amount; // Number, not a string
        updatesCount++;
      }
    }
  }

  // 3. Write back to sheet ONLY if we found fixes
  if (updatesCount > 0) {
    range.setValues(data);
    SpreadsheetApp.getUi().alert('Fixed ' + updatesCount + ' transactions!');
  } else {
    SpreadsheetApp.getUi().alert('No zero-amount rows found to fix.');
  }
}

/**
 * One-time repair: Retroactively converts CZK transactions from a
 * Czech Republic trip (June 12-17, 2026) that were mislabeled as ILS.
 * Safe to run multiple times — skips rows already marked with [Orig:].
 */
function fixCZKTransactions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transactions");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No transactions found.');
    return;
  }

  // Trip date window: June 12-17, 2026
  var tripStart = new Date(2026, 5, 12, 0, 0, 0);   // Month is 0-indexed
  var tripEnd = new Date(2026, 5, 17, 23, 59, 59);

  var range = sheet.getRange(2, 1, lastRow - 1, 5); // Cols A-E
  var data = range.getValues();

  // Fetch current CZK -> ILS rate
  var rate;
  try {
    var url = "https://api.exchangerate-api.com/v4/latest/CZK";
    var response = UrlFetchApp.fetch(url);
    var json = JSON.parse(response.getContentText());
    rate = json.rates.ILS;
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error fetching CZK rate: ' + e.toString());
    return;
  }

  var updatesCount = 0;

  for (var i = 0; i < data.length; i++) {
    var tDate = data[i][0];
    var amount = Number(data[i][2]);
    var rawText = String(data[i][4]);

    if (!(tDate instanceof Date)) continue;
    if (tDate < tripStart || tDate > tripEnd) continue;

    // Only fix rows that contain CZK/Kč in raw text AND weren't already converted
    var hasCZK = /CZK|Kč/i.test(rawText);
    var alreadyConverted = /\[Orig:/.test(rawText);

    if (hasCZK && !alreadyConverted && amount > 0) {
      var newAmount = Number((amount * rate).toFixed(2));
      data[i][2] = newAmount;
      data[i][4] = rawText + " [Orig: " + amount + " CZK]";
      updatesCount++;
      console.log("Fixed row " + (i + 2) + ": " + amount + " CZK -> " + newAmount + " ILS");
    }
  }

  if (updatesCount > 0) {
    range.setValues(data);
    SpreadsheetApp.getUi().alert('Fixed ' + updatesCount + ' CZK transactions!\nRate used: 1 CZK = ' + rate + ' ILS');
  } else {
    SpreadsheetApp.getUi().alert('No unfixed CZK transactions found between June 12-17.');
  }
}

/**
 * Retroactive fix: Scans all transactions and removes rows that are SMS ads
 * (not real transactions). Uses the same two-layer detection as the live filter:
 *   Layer 1: Ad keyword blocklist
 *   Layer 2: SMS structural check (no transaction keywords = ad)
 * Safe to run multiple times.
 */
function fixAdTransactions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transactions");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No transactions found.');
    return;
  }

  // Get all data: Date(A), Merchant(B), Amount(C), Category(D), RawText(E)
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  // Collect row numbers (1-indexed sheet rows) to delete
  var rowsToDelete = [];

  for (var i = 0; i < data.length; i++) {
    var merchant = String(data[i][1]);
    var rawText = String(data[i][4]);
    var sheetRow = i + 2; // 1-indexed, header is row 1

    // Layer 2 applies to rows where merchant is "SMS", empty, or generic (phone number / SMS)
    var isSMSRow = (merchant === "SMS" || merchant === "" || isGenericMerchant(merchant) || /sms|messages?/i.test(merchant));
    var reason = findAdReason(rawText, isSMSRow); // Same filter as the live doGet

    if (reason) {
      rowsToDelete.push(sheetRow);
      console.log("🗑️ " + reason.detail + " (row " + sheetRow + "): " + rawText.substring(0, 60));
    }
  }

  // Delete rows bottom-up to avoid index shifting
  if (rowsToDelete.length > 0) {
    for (var r = rowsToDelete.length - 1; r >= 0; r--) {
      sheet.deleteRow(rowsToDelete[r]);
    }
    SpreadsheetApp.getUi().alert('Removed ' + rowsToDelete.length + ' SMS ad transaction(s).');
  } else {
    SpreadsheetApp.getUi().alert('No SMS ad transactions found.');
  }
}

/**
 * Retroactive fix: Scans all transactions for foreign currency amounts
 * (USD, EUR, CZK) that were stored as-is without ILS conversion.
 * Uses HISTORICAL exchange rates for each transaction's actual date
 * via frankfurter.app (free, no key needed).
 * Safe to run multiple times — skips already-converted rows.
 */
function fixUnconvertedForeignTransactions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transactions");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No transactions found.');
    return;
  }

  var range = sheet.getRange(2, 1, lastRow - 1, 5); // Cols A-E
  var data = range.getValues();

  // Currency detection patterns (same tokens as doGet)
  var foreignCurrencies = [
    { code: "USD", pattern: /USD|\$/i },
    { code: "EUR", pattern: /EURO?|€/i },
    { code: "CZK", pattern: /CZK|Kč/i },
    { code: "GBP", pattern: /GBP|£/i }
  ];

  // First pass: identify rows to fix and collect unique (date, currency) pairs
  var rowsToFix = []; // { index, currency, dateKey }
  var rateKeysNeeded = {}; // "YYYY-MM-DD|EUR" -> true

  for (var i = 0; i < data.length; i++) {
    var rawText = String(data[i][4]);
    var amount = Number(data[i][2]);
    var tDate = data[i][0];
    if (amount <= 0 || /\[Orig:/.test(rawText)) continue;
    if (!(tDate instanceof Date)) continue;

    // Check which foreign currency this row contains
    var detectedCurrency = null;
    for (var c = 0; c < foreignCurrencies.length; c++) {
      if (foreignCurrencies[c].pattern.test(rawText)) {
        // Extra safety: verify the currency token appears near a number
        var nearNumberCheck = new RegExp("[\\d.,]+\\s*" + foreignCurrencies[c].pattern.source + "|" + foreignCurrencies[c].pattern.source + "\\s*[\\d.,]+", "i");
        if (nearNumberCheck.test(rawText)) {
          detectedCurrency = foreignCurrencies[c].code;
          break;
        }
      }
    }

    if (!detectedCurrency) continue;

    // Skip if the text also contains ILS/₪/NIS/ש"ח near a number
    var ilsPattern = /(?:ILS|NIS|₪|ש"?ח)\s*[\d.,]+|[\d.,]+\s*(?:ILS|NIS|₪|ש"?ח)/i;
    if (ilsPattern.test(rawText)) continue;

    var dateKey = dateToKey(tDate);
    rowsToFix.push({ index: i, currency: detectedCurrency, dateKey: dateKey });
    rateKeysNeeded[dateKey + "|" + detectedCurrency] = true;
  }

  if (rowsToFix.length === 0) {
    SpreadsheetApp.getUi().alert('No unconverted foreign transactions found.');
    return;
  }

  // Fetch historical rates for each unique (date, currency) pair
  // Uses frankfurter.app: free, no API key, supports historical lookups
  var rateCache = {}; // "YYYY-MM-DD|EUR" -> rate
  var rateKeys = Object.keys(rateKeysNeeded);

  for (var r = 0; r < rateKeys.length; r++) {
    var parts = rateKeys[r].split("|");
    var dateStr = parts[0]; // "YYYY-MM-DD"
    var currency = parts[1]; // "EUR"

    try {
      var url = "https://api.frankfurter.app/" + dateStr + "?from=" + currency + "&to=ILS";
      var response = UrlFetchApp.fetch(url);
      var json = JSON.parse(response.getContentText());
      rateCache[rateKeys[r]] = json.rates.ILS;
      console.log("📅 Rate for " + currency + " on " + dateStr + ": " + json.rates.ILS + " ILS");
    } catch (e) {
      SpreadsheetApp.getUi().alert('Error fetching ' + currency + ' rate for ' + dateStr + ': ' + e.toString());
      return;
    }

    // Brief pause between API calls to be polite
    if (r < rateKeys.length - 1) Utilities.sleep(500);
  }

  // Second pass: apply per-date rates
  var updatesCount = 0;
  var details = [];

  for (var f = 0; f < rowsToFix.length; f++) {
    var row = rowsToFix[f];
    var i = row.index;
    var amount = Number(data[i][2]);
    var rawText = String(data[i][4]);
    var rate = rateCache[row.dateKey + "|" + row.currency];

    if (!rate) {
      console.log("⚠️ No rate found for row " + (i + 2) + ", skipping");
      continue;
    }

    var newAmount = Number((amount * rate).toFixed(2));
    data[i][2] = newAmount;
    // Drop the [UNCONVERTED: ...] flag left by doGet when the live conversion failed
    data[i][4] = rawText.replace(/\s*\[UNCONVERTED:[^\]]*\]/, "") + " [Orig: " + amount + " " + row.currency + "]";
    updatesCount++;
    details.push("Row " + (i + 2) + ": " + amount + " " + row.currency + " → " + newAmount + " ILS (rate " + rate + " on " + row.dateKey + ")");
    console.log("Fixed row " + (i + 2) + ": " + amount + " " + row.currency + " -> " + newAmount + " ILS (rate " + rate + ")");
  }

  if (updatesCount > 0) {
    range.setValues(data);
    SpreadsheetApp.getUi().alert('Fixed ' + updatesCount + ' transaction(s):\n\n' + details.join('\n'));
  } else {
    SpreadsheetApp.getUi().alert('No unconverted foreign transactions found.');
  }
}

function checkAndNotify() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = ss.getSheetByName("Dashboard");
  var transSheet = ss.getSheetByName("Transactions");

  // --- 1. CONFIG ---
  var scriptProps = PropertiesService.getScriptProperties();
  var botToken = scriptProps.getProperty("TELEGRAM_BOT_TOKEN");
  if (!botToken) {
    console.log("❌ Missing TELEGRAM_BOT_TOKEN in Script Properties");
    return;
  }
  var rawChatIds = scriptProps.getProperty("TELEGRAM_CHAT_IDS") || "";
  var recipientIds = rawChatIds.split(",")
    .map(function (id) { return id.trim(); })
    .filter(Boolean);
  if (recipientIds.length === 0) {
    console.log("❌ Missing TELEGRAM_CHAT_IDS in Script Properties");
    return;
  }

  // --- 2. CALCULATE "SINCE YESTERDAY" ---
  // We scan transactions to see what happened in the last 24 hours
  var recentSpendMap = {};
  var oneDayAgo = new Date();
  oneDayAgo.setDate(oneDayAgo.getDate() - 1); // Exact time 24 hours ago

  try {
    var lastTransRow = transSheet.getLastRow();
    if (lastTransRow > 1) {
      // Get Date (Col A), Amount (Col C), Category (Col D)
      // A=0, B=1, C=2, D=3
      var transData = transSheet.getRange(2, 1, lastTransRow - 1, 4).getValues();

      for (var t = 0; t < transData.length; t++) {
        var tDate = transData[t][0];
        var tAmount = Number(transData[t][2]);
        var tCat = transData[t][3];

        // If transaction is newer than 24 hours ago
        if (tDate instanceof Date && tDate > oneDayAgo) {
          if (!recentSpendMap[tCat]) recentSpendMap[tCat] = 0;
          recentSpendMap[tCat] += tAmount;
        }
      }
    }
  } catch (e) {
    console.log("Error calculating diffs: " + e.toString());
  }

  // --- 3. BUILD REPORT ---
  var lastRow = dashSheet.getLastRow();
  var dashData = lastRow > 1 ? dashSheet.getRange(2, 5, lastRow - 1, 3).getValues() : []; // Cols E, F, G
  var monthName = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMMM");
  var messageLines = ["📊 <b>Daily Budget Report (" + monthName + ")</b>"];

  var hasUncategorized = false;
  var totalSpent = 0;

  for (var i = 0; i < dashData.length; i++) {
    var category = dashData[i][0];
    var budget = Number(dashData[i][1]);
    var spent = Number(dashData[i][2]);

    // Check for uncategorized items
    if (category === "Uncategorized" && spent > 0) hasUncategorized = true;

    // Skip empty rows, but SHOW categories even if 0 spent if there is a budget
    if (category === "" || (budget === 0 && spent === 0)) continue;

    var pct = (budget > 0) ? (spent / budget) : 0;
    var emoji = "🟢";
    if (pct > 0.8) emoji = "🟠";
    if (pct >= 1.0) emoji = "🔴";

    // THE NEW PART: Add the diff
    var line = emoji + " " + escapeHtml(category) + ": " + spent.toFixed(0) + " / " + budget.toFixed(0);
    totalSpent += spent;

    // If we spent money in this category recently, add "(+Amount)"
    if (recentSpendMap[category] && recentSpendMap[category] > 0) {
      line += " (+" + recentSpendMap[category].toFixed(0) + ")";
    }

    messageLines.push(line);
  }

  var totalSpentLine = "💰" + " " + "Total Spent: " + totalSpent.toFixed(0);
  messageLines.push(totalSpentLine);

  if (hasUncategorized) {
    messageLines.unshift("⚠️ <b>ACTION REQUIRED: Uncategorized Items!</b>");
  }

  // --- 4. CHART LOGIC (Optimized: no Utilities.formatDate in loops) ---
  var chartTotalBlob = null;
  var chartAvgBlob = null;

  // Local helper: Date -> "dd/MM" label
  function dateToLabel(d) {
    var m = d.getMonth() + 1;
    var day = d.getDate();
    return (day < 10 ? "0" : "") + day + "/" + (m < 10 ? "0" : "") + m;
  }

  try {
    var transRows = transSheet.getLastRow();
    if (transRows > 1) {
      var rawData = transSheet.getRange(2, 1, transRows - 1, 3).getValues();
      var currentYear = new Date().getFullYear();
      var dailyMap = {};

      for (var j = 0; j < rawData.length; j++) {
        var rDate = rawData[j][0];
        var rAmount = rawData[j][2];
        if (rDate instanceof Date && rDate.getFullYear() === currentYear) {
          var key = dateToKey(rDate);
          dailyMap[key] = (dailyMap[key] || 0) + Number(rAmount);
        }
      }

      var dataTotal = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, "Date")
        .addColumn(Charts.ColumnType.NUMBER, "Total");

      var dataAvg = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, "Date")
        .addColumn(Charts.ColumnType.NUMBER, "30-Day Avg");

      // Build an array of daily amounts from Jan 1 to today
      var today = new Date();
      var iteratorDate = new Date(currentYear, 0, 1);
      var dailyAmounts = []; // parallel array of spending per day
      var dayKeys = [];      // parallel array of "dd/MM" labels

      while (iteratorDate <= today) {
        var key = dateToKey(iteratorDate);
        dailyAmounts.push(dailyMap[key] || 0);
        dayKeys.push(dateToLabel(iteratorDate));
        iteratorDate.setDate(iteratorDate.getDate() + 1);
      }

      // O(N) sliding window for 30-day moving average
      var runningTotal = 0;
      var windowSum = 0;

      for (var d = 0; d < dailyAmounts.length; d++) {
        runningTotal += dailyAmounts[d];
        windowSum += dailyAmounts[d];

        // Slide the window: remove the element that fell out of the 30-day range
        if (d >= 30) {
          windowSum -= dailyAmounts[d - 30];
        }

        var windowDays = Math.min(d + 1, 30);
        var movingAvg = windowSum / windowDays;

        dataTotal.addRow([dayKeys[d], runningTotal]);
        dataAvg.addRow([dayKeys[d], movingAvg]);
      }

      chartTotalBlob = Charts.newLineChart()
        .setDataTable(dataTotal.build())
        .setTitle("Cumulative Spending (NIS)")
        .setDimensions(600, 400)
        .setOption('legend', { position: 'none' })
        .setOption('colors', ['#e63946'])
        .setOption('lineWidth', 3)
        .build().getAs('image/png');

      chartAvgBlob = Charts.newLineChart()
        .setDataTable(dataAvg.build())
        .setTitle("Last 30 Days Average (NIS)")
        .setDimensions(600, 400)
        .setOption('legend', { position: 'none' })
        .setOption('colors', ['#457b9d'])
        .setOption('lineWidth', 3)
        .build().getAs('image/png');
    }
  } catch (e) {
    console.log("Chart Error: " + e.toString());
  }

  // Get the specific URL for the Transactions tab
  var sheetUrl = ss.getUrl() + "#gid=" + transSheet.getSheetId();

  // Add an empty line for spacing, then the link
  messageLines.push("");
  messageLines.push("👉 <a href='" + sheetUrl + "'>Full Report</a>");

  // --- 5. SEND ---
  var fullMessage = messageLines.join("\n");
  recipientIds.forEach(function (chatId) {
    try {
      UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/sendMessage", {
        'method': 'post',
        'payload': { 'chat_id': chatId, 'text': fullMessage, 'parse_mode': 'HTML' }
      });
      if (chartTotalBlob && chartAvgBlob) {
        UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/sendPhoto", {
          'method': 'post',
          'payload': { 'chat_id': chatId, 'photo': chartTotalBlob, 'caption': 'Total YTD' }
        });
        UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/sendPhoto", {
          'method': 'post',
          'payload': { 'chat_id': chatId, 'photo': chartAvgBlob, 'caption': '30-Day Moving Average' }
        });
      }
    } catch (e) { console.log("Send Error: " + e.toString()); }
  });
}

function getGeminiCategory(merchantName) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    console.log("❌ Missing GEMINI_API_KEY in Script Properties");
    return "Uncategorized";
  }

  // OPTION 1: Try the "Latest" alias (often fixes the 404)
  var modelName = "gemini-2.5-flash";
  // var modelName = "gemini-pro"; // Fallback if flash fails

  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + apiKey;

  var myCategories = CATEGORIES.join(", ");

  var promptText =
    "Classify this credit card merchant (primarily in Israel unless indicated otherwise): '" + merchantName + "'\n" +
    "Into ONE of these categories: [" + myCategories + ", Uncategorized].\n" +
    "Instructions:\n" +
    "1. Use Google Search to look up the merchant name (search both the exact name and its Hebrew transliteration / Israeli business name, e.g. 'KOKPIT 67' -> 'קוקפיט 67') to verify what type of business it actually is.\n" +
    "2. Do NOT guess based on the literal meaning of words in the business name (for example, 'KOKPIT 67' is a cocktail bar/pub in Givatayim, NOT aviation/Transport).\n" +
    "3. Category rules:\n" +
    "- Restaurants, cafes, bars, pubs, cocktail bars, nightlife, wineries, bakeries, ice cream shops, fast food, food delivery ('Shipudei', 'Falafel', 'Shawarma', 'Cafe', 'Bar', 'Pub', 'Kokpit 67') -> 'Restaurants'\n" +
    "- Supermarkets, mini-markets, grocery stores, butchers, greengrocers ('Super', 'Shufersal', 'Market', 'AM:PM') -> 'Groceries'\n" +
    "- Pharmacies, drugstores, clinics, doctors, dentists ('Pharm', 'Drugstore', 'Super-Pharm', 'Be', 'Dental', 'Medical', 'Clinic') -> 'Health'\n" +
    "- 'Ikea', furniture, hardware, home goods -> 'Home Improvement'\n" +
    "- 'Bit', 'Paybox' -> 'Uncategorized' (unless context implies otherwise)\n" +
    "- If the merchant cannot be identified with confidence even after searching, return 'Uncategorized'.\n" +
    "- Return ONLY the exact category name (no explanation, no JSON, no punctuation).";

  var payload = {
    "contents": [{
      "parts": [{ "text": promptText }]
    }],
    "tools": [{
      "google_search": {}
    }],
    "generationConfig": { "temperature": 0 }
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    // Fallback without google_search tool if search grounding fails (e.g. tool quota error)
    if (responseCode !== 200) {
      console.log("⚠️ AI Search Grounding Error (" + responseCode + "), retrying without search tool...");
      delete payload.tools;
      options.payload = JSON.stringify(payload);
      response = UrlFetchApp.fetch(url, options);
      responseCode = response.getResponseCode();
      responseText = response.getContentText();
    }

    // Debug: Print error if API fails
    if (responseCode !== 200) {
      console.log("❌ AI API Error (" + responseCode + "): " + responseText);
      return "Uncategorized";
    }

    var json = JSON.parse(responseText);
    if (json.candidates && json.candidates.length > 0 && json.candidates[0].content) {
      var parts = json.candidates[0].content.parts || [];
      var answer = parts
        .filter(function (p) { return !p.thought && typeof p.text === "string"; })
        .map(function (p) { return p.text; })
        .join("\n");
      // Only accept answers from CATEGORIES (case/punctuation tolerant)
      var category = normalizeCategory(answer);
      if (category === "Uncategorized") console.log("⚠️ AI answer not in category list: " + answer);
      return category;
    }
  } catch (e) {
    console.log("❌ Script Error: " + e.toString());
  }

  return "Uncategorized";
}

function listMyModels() {
  var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    console.log("❌ Missing GEMINI_API_KEY in Script Properties");
    return;
  }
  var url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKey;

  try {
    var response = UrlFetchApp.fetch(url);
    var json = JSON.parse(response.getContentText());

    // This will print every valid model name your key can use
    var names = json.models.map(function (m) { return m.name; }).join("\n");
    console.log("✅ AVAILABLE MODELS:\n" + names);

  } catch (e) {
    console.log("❌ ERROR: " + e.toString());
  }
}

/**
 * Adds a new merchant and category to the 'Categories' sheet (Col A & B).
 * Prevents duplicates to keep the sheet clean.
 */
function addToMerchantCache(merchantName, category) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var catSheet = ss.getSheetByName("Categories"); // Or "Merchant Map" if you use that

  if (!catSheet) return;

  // 1. Check for duplicates (Don't add "McDonalds" twice)
  var lastRow = catSheet.getLastRow();
  if (lastRow > 1) {
    var data = catSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Get all existing names
    var cleanNew = normalizeMerchant(merchantName);

    for (var i = 0; i < data.length; i++) {
      var cleanExisting = normalizeMerchant(data[i][0]);
      if (cleanExisting === cleanNew) {
        console.log("Cache: Merchant '" + merchantName + "' already exists.");
        return;
      }
    }
  }

  // 2. Add to the bottom
  catSheet.appendRow([merchantName, category]);
  SpreadsheetApp.flush();
  console.log("Cache: Added '" + merchantName + "' -> '" + category + "'");
}

function testAI() {
  var result = getGeminiCategory("שיפודי סמי ובניו - רמת גן");
  console.log("Result: " + result);
}

function convertToILS(amount, currency) {
  if (currency === "ILS") return amount;

  try {
    // Free API for exchange rates (Updates daily)
    var url = "https://api.exchangerate-api.com/v4/latest/" + currency;
    var response = UrlFetchApp.fetch(url);
    var json = JSON.parse(response.getContentText());
    var rate = json.rates && json.rates.ILS;
    if (!(rate > 0)) throw new Error("No ILS rate in response for " + currency);

    var newAmount = amount * rate;
    console.log("💱 Converted " + amount + " " + currency + " -> " + newAmount.toFixed(2) + " ILS (Rate: " + rate + ")");
    return Number(newAmount.toFixed(2));
  } catch (e) {
    console.log("❌ Currency Error: " + e.toString());
    return null; // Caller flags the row as [UNCONVERTED] so it can be repaired later
  }
}

/**
 * Unit Test Runner for Google Apps Script Editor / Logger.
 */
function runUnitTests() {
  console.log("🧪 Running Receipt Extractor Unit Tests...");
  var passed = 0;
  var failed = 0;

  function assert(desc, cond) {
    if (cond) {
      console.log("  ✅ PASS: " + desc);
      passed++;
    } else {
      console.log("  ❌ FAIL: " + desc);
      failed++;
    }
  }

  // 1. Spam & Ad Filter test samples
  var testAdSample = "מבצע מיוחד 20% הנחה על כל המוצרים!";
  var testTxSample = "חשבונך חוייב בסך 150.00 ₪ עבור מבצע הרהיטים";

  assert("Ad rejected by keyword filter", testAdSample.indexOf("מבצע") > -1 && testAdSample.indexOf("חוייב") === -1);
  assert("Transaction preserved by safety word", testTxSample.indexOf("חוייב") > -1);

  // 2. Merchant Extraction test samples
  var isracardSample = "בסך 52.94 ILS בaliexpress - LUXEMBOURG.";
  var match3 = isracardSample.match(/בסך .*? ב(.*?)(?:\.|\n|$)/);
  assert("Isracard merchant extracted", match3 && match3[1] === "aliexpress - LUXEMBOURG");

  // 3. Amount Extraction test samples
  var amountText = "\u200Fעסקה בסך\u202A 99.90 ₪";
  var cleanText = amountText.replace(/[\u200B-\u200F\u202A-\u202E\u061C]/g, " ");
  var matchAmount = cleanText.match(/(?:בסך|על סך)\s*(?:₪|ILS)?\s*([\d,]+(?:\.\d+)?)/i);
  assert("Amount extracted correctly despite Unicode direction marks", matchAmount && matchAmount[1] === "99.90");

  console.log("📊 Summary: " + passed + " Passed, " + failed + " Failed.");
}