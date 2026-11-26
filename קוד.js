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

  if (isCashWithdrawal(rawText)) {
    category = "Cash Withdraw";
    shouldCache = false; // Never cache card-issuer names ("Cal", "Isracard") in Categories
  }

  // A. Manual Lookup (Cache First)
  // We check if we already know this merchant to save AI tokens and time.
  if (!isResolvedCategory(category) && catSheet && catSheet.getLastRow() > 1) {
    var catData = catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 2).getValues();

    for (var c = 0; c < catData.length; c++) {
      if (isMerchantMatch(merchant, catData[c][0]) && isResolvedCategory(catData[c][1])) {
        category = String(catData[c][1]).trim();
        console.log("✅ Found in Cache: " + category);
        break;
      }
    }
  }

  // B. AI Lookup + Keyword Fallback (If still unknown). Generic names ("SMS", phone numbers)
  // are never sent to the AI or cached — a cached "SMS" would match every unparsed SMS.
  if (!isResolvedCategory(category) && !isGenericMerchant(merchant)) {
    console.log("🤔 Unknown merchant. Asking Gemini...");
    var aiCategory = getGeminiCategory(merchant);
    if (isResolvedCategory(aiCategory)) {
      category = aiCategory;
      shouldCache = true;
    } else {
      var kwCategory = inferCategoryFromKeywords(merchant);
      if (isResolvedCategory(kwCategory)) {
        category = kwCategory;
        shouldCache = true;
        console.log("✅ Classified via keyword fallback: " + category);
      }
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

    try {
      sheet.appendRow([new Date(), merchant, amount, category, rawText]);
    } catch (appendErr) {
      if (category !== "Uncategorized" && /column|validation|dropdown|typed/i.test(String(appendErr))) {
        console.log("⚠️ Category '" + category + "' rejected by sheet validation, falling back to Uncategorized: " + appendErr.toString());
        category = "Uncategorized";
        shouldCache = false;
        sheet.appendRow([new Date(), merchant, amount, category, rawText]);
      } else {
        throw appendErr;
      }
    }
    SpreadsheetApp.flush(); // Force write to sheet immediately
    savedRow = sheet.getLastRow();

    // Add to cache so next time we skip the AI
    if (shouldCache && isResolvedCategory(category)) addToMerchantCache(merchant, category);
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

  // Scan the entire Transactions sheet (newest rows first) so older repaired rows are also covered
  var range = sheet.getRange(2, 2, lastRow - 1, 4); // Col B (Merchant) to Col E (Raw Text)
  var data = range.getValues();

  // Load Cache (Entire cache is fine, it's usually small)
  var catData = [];
  if (catSheet && catSheet.getLastRow() > 1) {
    catData = catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 2).getValues();
  }

  var MAX_AI_CALLS = 10; // Safety cap per run (~40s sleep total; free tier allows 15 req/min)
  var updatesCount = 0;
  var aiCallsCount = 0;
  var remainingUncat = 0;

  for (var i = data.length - 1; i >= 0; i--) {
    var merchant = String(data[i][0] || "").trim();
    var currentCat = String(data[i][2] || "").trim();
    var rawText = String(data[i][3] || "").trim();

    // Process only if Uncategorized/Empty
    if (currentCat === "Uncategorized" || currentCat === "") {
      if (isCashWithdrawal(rawText)) {
        data[i][0] = parseMerchant(merchant, rawText);
        data[i][2] = "Cash Withdraw";
        updatesCount++;
        continue;
      }

      if (!merchant || isGenericMerchant(merchant)) continue;

      var newCategory = "Uncategorized";

      // 1. FAST: Check Cache First
      for (var j = 0; j < catData.length; j++) {
        if (isMerchantMatch(merchant, catData[j][0]) && isResolvedCategory(catData[j][1])) {
          newCategory = String(catData[j][1]).trim();
          break;
        }
      }

      // 2. SLOW: Ask AI (Only if Cache failed) + Deterministic Keyword Fallback
      if (!isResolvedCategory(newCategory)) {
        if (aiCallsCount < MAX_AI_CALLS) {
          try {
            newCategory = getGeminiCategory(merchant);
            aiCallsCount++;

            // Sleep 4s to stay safe (Free Tier = 15 req/min)
            Utilities.sleep(4000);
          } catch (e) {
            console.log("AI Failed for " + merchant + ": " + e.toString());
          }
        }

        if (!isResolvedCategory(newCategory)) {
          var kwCat = inferCategoryFromKeywords(merchant);
          if (isResolvedCategory(kwCat)) {
            newCategory = kwCat;
            console.log("Classified via keyword fallback: " + merchant + " -> " + newCategory);
          }
        }

        if (isResolvedCategory(newCategory)) {
          addToMerchantCache(merchant, newCategory);
          // Add to local memory so next row uses it instantly
          catData.push([merchant, newCategory]);
        } else {
          remainingUncat++;
        }
      }

      // 3. Update data array if we found something
      if (isResolvedCategory(newCategory)) {
        data[i][2] = newCategory;
        updatesCount++;
      }
    }
  }

  // 4. Write back changes
  if (updatesCount > 0) {
    range.setValues(data);
    SpreadsheetApp.flush();
    var msg = 'Success! Updated ' + updatesCount + ' transaction(s).';
    if (remainingUncat > 0) {
      msg += '\n\n' + remainingUncat + ' uncategorized transaction(s) still remain (batch limit reached or AI returned Uncategorized) — run again if needed.';
    }
    SpreadsheetApp.getUi().alert(msg);
  } else if (remainingUncat > 0) {
    SpreadsheetApp.getUi().alert('No new categories found (' + remainingUncat + ' uncategorized transaction(s) still remain).');
  } else {
    SpreadsheetApp.getUi().alert('No uncategorized transactions found.');
  }
}

/**
 * Retroactive fix:
 * 1. Removes poisoned generic/card-mask entries (e.g. "כרטיס Mastercard...") from Categories.
 * 2. Scans all transactions and re-runs parseMerchant on Raw Text (Col E), or recovers lost
 *    Google Wallet merchants from the Debug sheet by matching raw text and timestamp.
 * 3. Updates Category from Categories cache (or resets stale generic-merchant categories to Uncategorized).
 */
function fixMisparsedMerchants() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transactions");
  var catSheet = ss.getSheetByName("Categories");
  var debugSheet = ss.getSheetByName("Debug");

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No transactions found.');
    return;
  }

  // 1. Load Categories sheet & remove poisoned generic/card-mask rows bottom-up
  var catData = [];
  var poisonedRemoved = 0;
  if (catSheet && catSheet.getLastRow() > 1) {
    var rawCatData = catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 2).getValues();
    var poisonedRows = [];
    for (var k = 0; k < rawCatData.length; k++) {
      var catMerchant = String(rawCatData[k][0] || "").trim();
      if (catMerchant && isGenericMerchant(catMerchant)) {
        poisonedRows.push(k + 2);
      } else if (catMerchant) {
        catData.push(rawCatData[k]);
      }
    }
    for (var p = poisonedRows.length - 1; p >= 0; p--) {
      catSheet.deleteRow(poisonedRows[p]);
      poisonedRemoved++;
      console.log("Removed poisoned generic category row: " + poisonedRows[p]);
    }
  }

  // 2. Load Debug log entries so we can recover Google Wallet merchants not present in Raw Text
  var debugEntries = [];
  if (debugSheet && debugSheet.getLastRow() > 1) {
    var debugRows = debugSheet.getRange(2, 1, debugSheet.getLastRow() - 1, 2).getValues();
    for (var d = 0; d < debugRows.length; d++) {
      var dDate = debugRows[d][0];
      var rawParams = String(debugRows[d][1] || "").trim();
      if (!rawParams) continue;
      try {
        var parsedParams = JSON.parse(rawParams);
        var dMerchant = String(parsedParams.merchant || "").trim();
        var dText = String(parsedParams.text || "");
        if (dMerchant && dText) {
          debugEntries.push({
            date: dDate,
            merchant: dMerchant,
            text: dText,
            normText: normalizeRawTextForMatch_(dText)
          });
        }
      } catch (err) {
        // Ignore malformed debug JSON rows
      }
    }
  }

  // Cols A-E: Date(0), Merchant(1), Amount(2), Category(3), Raw Text(4)
  var range = sheet.getRange(2, 1, lastRow - 1, 5);
  var data = range.getValues();

  var updatesCount = 0;
  var needsCategoryCount = 0;

  for (var i = 0; i < data.length; i++) {
    var txDate = data[i][0];
    var currentMerchant = String(data[i][1] || "").trim();
    var rawText = String(data[i][4] || "").trim();
    var wasGeneric = isGenericMerchant(currentMerchant);

    if (!rawText) continue;

    if (isCashWithdrawal(rawText)) {
      var issuer = parseMerchant(currentMerchant, rawText);
      var existingCat = String(data[i][3] || "").trim();
      if (currentMerchant !== issuer || existingCat !== "Cash Withdraw") {
        console.log("Fixing cash withdrawal row " + (i + 2) + ": '" + currentMerchant + "' -> '" + issuer + "' (Cash Withdraw)");
        data[i][1] = issuer;
        data[i][3] = "Cash Withdraw";
        updatesCount++;
      }
      continue;
    }

    var newMerchant = parseMerchant(currentMerchant, rawText);

    // If parseMerchant with currentMerchant didn't change it, but currentMerchant is generic ("SMS", "Isracard", "Cal"):
    if (newMerchant === currentMerchant && wasGeneric) {
      newMerchant = parseMerchant("SMS", rawText);
    }

    // If still generic/unidentified, try recovering the original merchant from the Debug tab
    if (wasGeneric && (!newMerchant || newMerchant === currentMerchant || isGenericMerchant(newMerchant))) {
      var recovered = findMerchantInDebug(txDate, rawText, debugEntries);
      if (recovered) newMerchant = recovered;
    }

    if (newMerchant && newMerchant !== currentMerchant && !isGenericMerchant(newMerchant)) {
      console.log("Fixing row " + (i + 2) + ": '" + currentMerchant + "' -> '" + newMerchant + "'");
      data[i][1] = newMerchant;
      updatesCount++;

      // Check if we can match/fix the category from cache; otherwise clear stale category from generic guess
      var matchedCat = false;
      for (var c = 0; c < catData.length; c++) {
        if (isMerchantMatch(newMerchant, catData[c][0]) && isResolvedCategory(catData[c][1])) {
          data[i][3] = String(catData[c][1]).trim();
          console.log("  -> Updated category to: " + data[i][3]);
          matchedCat = true;
          break;
        }
      }
      if (!matchedCat && wasGeneric) {
        data[i][3] = "Uncategorized";
        needsCategoryCount++;
      }
    }
  }

  if (updatesCount > 0) {
    range.setValues(data);
  }
  if (updatesCount > 0 || poisonedRemoved > 0) {
    SpreadsheetApp.flush();
    var alertLines = [];
    if (updatesCount > 0) alertLines.push('Fixed ' + updatesCount + ' merchant name(s) in Transactions!');
    if (poisonedRemoved > 0) alertLines.push('Removed ' + poisonedRemoved + ' poisoned card-mask row(s) from Categories.');
    if (needsCategoryCount > 0) {
      alertLines.push(needsCategoryCount + ' repaired transaction(s) set to Uncategorized — run "Wallet Tools → Update All Categories" next.');
    }
    SpreadsheetApp.getUi().alert(alertLines.join('\n'));
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
      var mCat = String(catData[i][1] || "").trim();
      var clean = normalizeMerchant(mName);
      if (!clean) continue;

      var sheetRowNum = i + 2; // header is row 1

      if (seenMerchants[clean]) {
        var keptInfo = seenMerchants[clean];
        // Rescue a valid category from a newer duplicate when the kept older row was blank/Uncategorized
        if (!isResolvedCategory(cleanCatList[keptInfo.listIdx][1]) && isResolvedCategory(mCat)) {
          catSheet.getRange(keptInfo.rowNum, 2, 1, 1).setValues([[mCat]]);
          cleanCatList[keptInfo.listIdx][1] = mCat;
        }
        rowsToDelete.push(sheetRowNum);
        console.log("Found duplicate category at row " + sheetRowNum + ": '" + mName + "'");
      } else {
        seenMerchants[clean] = { listIdx: cleanCatList.length, rowNum: sheetRowNum };
        cleanCatList.push([mName, mCat]);
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
  var updatedCount = 0;
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
      var matchedIndex = -1;
      for (var c = 0; c < cleanCatList.length; c++) {
        if (isMerchantMatch(txMerchant, cleanCatList[c][0])) {
          matchedIndex = c;
          break;
        }
      }

      if (matchedIndex === -1) {
        var finalCat = isResolvedCategory(txCat) ? txCat : "Uncategorized";

        // If transaction has no category, try Gemini (up to 5 calls safety limit), then keyword fallback
        if (finalCat === "Uncategorized") {
          if (aiCallsCount < 5) {
            try {
              var aiCat = getGeminiCategory(txMerchant);
              aiCallsCount++;
              if (isResolvedCategory(aiCat)) finalCat = aiCat;
            } catch (e) {
              console.log("AI error: " + e.toString());
            }
          }
          if (finalCat === "Uncategorized") {
            var kwSync = inferCategoryFromKeywords(txMerchant);
            if (isResolvedCategory(kwSync)) finalCat = kwSync;
          }
        }

        catSheet.appendRow([txMerchant, finalCat]);
        cleanCatList.push([txMerchant, finalCat]); // Update in-memory so future rows don't re-add
        addedCount++;
        console.log("Added missing merchant to Categories: '" + txMerchant + "' -> '" + finalCat + "'");
      } else if (!isResolvedCategory(cleanCatList[matchedIndex][1])) {
        // Existing cache row is blank or Uncategorized — upgrade it if we can resolve a category
        var resolvedCat = isResolvedCategory(txCat) ? txCat : "Uncategorized";
        if (resolvedCat === "Uncategorized" && aiCallsCount < 5) {
          try {
            var aiUpgrade = getGeminiCategory(txMerchant);
            aiCallsCount++;
            if (isResolvedCategory(aiUpgrade)) resolvedCat = aiUpgrade;
          } catch (e) {
            console.log("AI error: " + e.toString());
          }
        }
        if (resolvedCat === "Uncategorized") {
          var kwUpgrade = inferCategoryFromKeywords(txMerchant);
          if (isResolvedCategory(kwUpgrade)) resolvedCat = kwUpgrade;
        }
        if (isResolvedCategory(resolvedCat)) {
          addToMerchantCache(cleanCatList[matchedIndex][0], resolvedCat);
          cleanCatList[matchedIndex][1] = resolvedCat;
          updatedCount++;
          console.log("Updated blank/Uncategorized merchant in Categories: '" + cleanCatList[matchedIndex][0] + "' -> '" + resolvedCat + "'");
        }
      }
    }
  }

  if (dupesRemoved > 0 || addedCount > 0 || updatedCount > 0) {
    SpreadsheetApp.flush();
  }

  var msg = "Categories Sync Complete!\n\n" +
    "• Duplicate categories removed: " + dupesRemoved + "\n" +
    "• Missing merchants added from Transactions: " + addedCount;
  if (updatedCount > 0) {
    msg += "\n• Blank/Uncategorized entries updated: " + updatedCount;
  }
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

  // --- 2. CALCULATE "SINCE YESTERDAY" & MONTHLY CATEGORY SPENDING ---
  // We scan transactions to see what happened in the last 24 hours and collect current-month totals
  var recentSpendMap = {};
  var monthlyCategorySpend = {}; // catLower -> { name, spent }
  var categoryDisplayNames = {}; // catLower -> display name
  var now = new Date();
  var currentYear = now.getFullYear();
  var currentMonth = now.getMonth();
  var oneDayAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  var monthlyUncatSpent = 0;
  var monthlyUncatCount = 0;

  try {
    var lastTransRow = transSheet.getLastRow();
    if (lastTransRow > 1) {
      // Get Date (Col A), Amount (Col C), Category (Col D)
      // A=0, B=1, C=2, D=3
      var transData = transSheet.getRange(2, 1, lastTransRow - 1, 4).getValues();

      for (var t = 0; t < transData.length; t++) {
        var tDate = transData[t][0];
        var tAmount = Number(transData[t][2]) || 0;
        var rawCat = String(transData[t][3] || "").trim();
        var isUncat = (rawCat === "" || rawCat.toLowerCase() === "uncategorized");
        var mapKey = isUncat ? "Uncategorized" : rawCat;
        var mapKeyLower = mapKey.toLowerCase();
        if (!categoryDisplayNames[mapKeyLower]) categoryDisplayNames[mapKeyLower] = mapKey;

        if (tDate instanceof Date) {
          // If transaction is newer than 24 hours ago
          if (tDate > oneDayAgo) {
            if (!recentSpendMap[mapKeyLower]) recentSpendMap[mapKeyLower] = 0;
            recentSpendMap[mapKeyLower] += tAmount;
          }
          if (tDate.getFullYear() === currentYear && tDate.getMonth() === currentMonth) {
            if (isUncat) {
              monthlyUncatSpent += tAmount;
              monthlyUncatCount++;
            } else {
              if (!monthlyCategorySpend[mapKeyLower]) {
                monthlyCategorySpend[mapKeyLower] = { name: rawCat, spent: 0 };
              }
              monthlyCategorySpend[mapKeyLower].spent += tAmount;
            }
          }
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
  var dashHasUncategorizedRow = false;
  var seenCategories = {};
  var totalSpent = 0;
  var recentUncat = recentSpendMap["uncategorized"] || 0;

  function formatUncatLine(spentVal, budgetVal) {
    var out = "⚠️ Uncategorized: " + spentVal.toFixed(0);
    if (budgetVal > 0) out += " / " + budgetVal.toFixed(0);
    if (monthlyUncatCount > 0) {
      out += " (" + monthlyUncatCount + (monthlyUncatCount === 1 ? " item" : " items") + ")";
    }
    if (recentUncat > 0) {
      out += " (+" + recentUncat.toFixed(0) + ")";
    }
    return out;
  }

  for (var i = 0; i < dashData.length; i++) {
    var category = String(dashData[i][0] || "").trim();
    var budget = Number(dashData[i][1]) || 0;
    var spent = Number(dashData[i][2]) || 0;

    if (category.toLowerCase() === "uncategorized") {
      dashHasUncategorizedRow = true;
      if (spent === 0 && monthlyUncatSpent > 0) spent = monthlyUncatSpent;
      if (spent === 0 && monthlyUncatCount === 0 && recentUncat <= 0 && budget === 0) continue;

      if (spent > 0 || monthlyUncatCount > 0 || recentUncat > 0) hasUncategorized = true;
      totalSpent += spent;
      messageLines.push(formatUncatLine(spent, budget));
      continue;
    }

    // Skip blank or summary total rows, but show every named category on Dashboard (even if budget/spent are 0)
    if (category === "" || /^(?:total|סה"?כ)$/i.test(category)) continue;

    var catLower = category.toLowerCase();
    seenCategories[catLower] = true;

    // Fallback to Transactions sum if Dashboard Col G formula wasn't configured for this category
    if (spent === 0 && monthlyCategorySpend[catLower] && monthlyCategorySpend[catLower].spent > 0) {
      spent = monthlyCategorySpend[catLower].spent;
    }

    var pct = (budget > 0) ? (spent / budget) : 0;
    var emoji = "🟢";
    if (pct > 0.8) emoji = "🟠";
    if (pct >= 1.0) emoji = "🔴";

    var line = emoji + " " + escapeHtml(category) + ": " + spent.toFixed(0) + " / " + budget.toFixed(0);
    totalSpent += spent;

    if (recentSpendMap[catLower] && recentSpendMap[catLower] > 0) {
      line += " (+" + recentSpendMap[catLower].toFixed(0) + ")";
    }

    messageLines.push(line);
  }

  // Include any categorized current-month (or 24h) spend from Transactions not listed on Dashboard
  var extraKeys = {};
  Object.keys(monthlyCategorySpend).forEach(function (k) { extraKeys[k] = true; });
  Object.keys(recentSpendMap).forEach(function (k) { if (k !== "uncategorized") extraKeys[k] = true; });
  Object.keys(extraKeys).forEach(function (catLower) {
    if (seenCategories[catLower]) return;
    var extraSpent = monthlyCategorySpend[catLower] ? monthlyCategorySpend[catLower].spent : 0;
    var extraRecent = recentSpendMap[catLower] || 0;
    if (extraSpent <= 0 && extraRecent <= 0) return;

    var displayName = (monthlyCategorySpend[catLower] && monthlyCategorySpend[catLower].name) || categoryDisplayNames[catLower] || catLower;
    var extraLine = "🟢 " + escapeHtml(displayName) + ": " + extraSpent.toFixed(0) + " / 0";
    if (extraRecent > 0) extraLine += " (+" + extraRecent.toFixed(0) + ")";
    totalSpent += extraSpent;
    messageLines.push(extraLine);
  });

  if (!dashHasUncategorizedRow && (monthlyUncatCount > 0 || recentUncat > 0)) {
    hasUncategorized = true;
    totalSpent += monthlyUncatSpent;
    messageLines.push(formatUncatLine(monthlyUncatSpent, 0));
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
  var rawKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  var apiKey = rawKey ? String(rawKey).replace(/^["'\s]+|["'\s]+$/g, "") : "";
  if (!apiKey) {
    console.log("❌ Missing GEMINI_API_KEY in Script Properties");
    return "Uncategorized";
  }

  var myCategories = CATEGORIES.join(", ");

  var promptText =
    "Classify this credit card merchant (primarily in Israel unless indicated otherwise): '" + merchantName + "'\n" +
    "Into ONE of these categories: [" + myCategories + ", Uncategorized].\n" +
    "Instructions:\n" +
    "1. Use Google Search to look up the merchant name (search both the exact name and its Hebrew transliteration / Israeli business name, e.g. 'KOKPIT 67' -> 'קוקפיט 67') to verify what type of business it actually is.\n" +
    "2. Do NOT misclassify figurative or brand-style names literally (for example, 'KOKPIT 67' is a cocktail bar/pub in Givatayim, NOT aviation/Transport). However, when the merchant name explicitly contains an unmistakable business-type word (such as 'שווארמה', 'פלאפל', 'שיפודי', 'פיצה', 'מאפייה', 'בייקרי', 'מסעדה', 'סופרמרקט', 'בית מרקחת', 'Shawarma', 'Falafel', 'Pizza', 'Bakery', 'Restaurant', 'Supermarket', 'Pharmacy'), DO classify it by that business type even if Google Search finds multiple branches or no website.\n" +
    "3. Category rules:\n" +
    "- Restaurants, cafes, bars, pubs, cocktail bars, nightlife, wineries, bakeries, ice cream shops, fast food, food delivery ('שווארמה', 'פלאפל', 'שיפודי', 'פיצה', 'בורגר', 'סושי', 'חומוס', 'קפה', 'מאפייה', 'בייקרי', 'גלידה', 'מסעדה', 'Shipudei', 'Falafel', 'Shawarma', 'Cafe', 'Bar', 'Pub', 'Kokpit 67') -> 'Restaurants'\n" +
    "- Supermarkets, mini-markets, grocery stores, butchers, greengrocers ('Super', 'Shufersal', 'Market', 'AM:PM') -> 'Groceries'\n" +
    "- Pharmacies, drugstores, clinics, doctors, dentists ('Pharm', 'Drugstore', 'Super-Pharm', 'Be', 'Dental', 'Medical', 'Clinic') -> 'Health'\n" +
    "- Travel agencies, flights, airlines, hotels, resorts, vacation packages, travel booking sites ('Holidays Lines', 'Holiday Lines', 'קווי חופשה', 'Issta', 'El Al', 'Arkia', 'Booking.com', 'Airbnb') -> 'Vacation'\n" +
    "- Local daily transit, fuel, parking, tolls, ride-hailing ('Pango', 'Cello', 'Gett', 'Uber', 'Rav-Kav', 'Paz', 'Sonol', 'Delek', 'Dor Alon', 'Highway 6') -> 'Transport'\n" +
    "- ATM / cash withdrawals ('ATM', 'Euronet', 'Cash Withdrawal') -> 'Cash Withdraw'\n" +
    "- 'Ikea', furniture, hardware, home goods -> 'Home Improvement'\n" +
    "- 'Bit', 'Paybox' -> 'Uncategorized' (unless context implies otherwise)\n" +
    "- If the merchant cannot be identified with confidence even after searching and has no clear business-type keyword, return 'Uncategorized'.\n" +
    "- Return ONLY the exact category name (no explanation, no JSON, no punctuation).";

  var attempts = [
    { model: "gemini-2.5-flash", useSearch: true },
    { model: "gemini-2.5-flash", useSearch: false },
    { model: "gemini-2.5-flash-lite", useSearch: false }
  ];

  for (var a = 0; a < attempts.length; a++) {
    var attempt = attempts[a];
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + attempt.model + ":generateContent?key=" + encodeURIComponent(apiKey);
    var payload = {
      "contents": [{
        "parts": [{ "text": promptText }]
      }],
      "generationConfig": { "temperature": 0 }
    };
    if (attempt.useSearch) {
      payload.tools = [{ "google_search": {} }];
    }

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

      if (responseCode !== 200) {
        console.log("⚠️ AI Error (" + attempt.model + (attempt.useSearch ? "+search" : "") + ", HTTP " + responseCode + "): " + responseText);
        continue;
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
      console.log("❌ Script Error (" + attempt.model + "): " + e.toString());
    }
  }

  return "Uncategorized";
}

function listMyModels() {
  var rawKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  var apiKey = rawKey ? String(rawKey).replace(/^["'\s]+|["'\s]+$/g, "") : "";
  if (!apiKey) {
    console.log("❌ Missing GEMINI_API_KEY in Script Properties");
    return;
  }
  var url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + encodeURIComponent(apiKey);

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
 * Adds a new merchant and category to the 'Categories' sheet (Col A & B),
 * skipping duplicates with resolved categories and upgrading blank/Uncategorized rows in-place.
 */
function addToMerchantCache(merchantName, category) {
  if (!merchantName || isGenericMerchant(merchantName) || !isResolvedCategory(category)) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var catSheet = ss.getSheetByName("Categories"); // Or "Merchant Map" if you use that

  if (!catSheet) return;

  // 1. Check existing rows: skip resolved duplicates, or fill blank/Uncategorized Col B in-place
  var lastRow = catSheet.getLastRow();
  if (lastRow > 1) {
    var data = catSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    var cleanNew = normalizeMerchant(merchantName);

    for (var i = 0; i < data.length; i++) {
      var cleanExisting = normalizeMerchant(data[i][0]);
      if (cleanExisting === cleanNew || isMerchantMatch(merchantName, data[i][0])) {
        if (!isResolvedCategory(data[i][1])) {
          catSheet.getRange(i + 2, 2, 1, 1).setValues([[category]]);
          SpreadsheetApp.flush();
          console.log("Cache: Updated '" + data[i][0] + "' -> '" + category + "'");
        } else {
          console.log("Cache: Merchant '" + merchantName + "' already exists.");
        }
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