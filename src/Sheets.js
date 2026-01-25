/**
 * Возвращает существующий лист или создаёт новый.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  return ss.insertSheet(name);
}

/**
 * Find an existing sheet by logical key across all supported languages.
 * Does NOT depend on current language selection.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetKey One of PF_SHEET_KEYS values.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function pfFindSheetByKey_(ss, sheetKey) {
  var candidates = pfAllSheetNames_(sheetKey);
  for (var i = 0; i < candidates.length; i++) {
    var sheet = ss.getSheetByName(candidates[i]);
    if (sheet) return sheet;
  }
  return null;
}

/**
 * Find or create sheet by logical key, using i18n dictionary across languages.
 * Ensures the sheet name matches current language.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetKey One of PF_SHEET_KEYS values.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function pfFindOrCreateSheetByKey_(ss, sheetKey) {
  var lang = pfGetLanguage_();
  var desiredName = pfT_('sheet.' + sheetKey, lang);

  // Find sheet by any known localized name.
  var sheet = pfFindSheetByKey_(ss, sheetKey);

  if (!sheet) {
    sheet = ss.insertSheet(desiredName);
    return sheet;
  }

  // Rename to desired language name if needed.
  if (sheet.getName() !== desiredName) {
    sheet.setName(desiredName);
  }
  return sheet;
}

/**
 * Applies localization to known sheets + headers.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function pfApplyLocalization_(ss) {
  // Ensure settings sheet exists before reading anything else.
  pfEnsureSettingsSheet_(ss);

  // Ensure + rename known sheets.
  pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.SETTINGS);
  pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.CATEGORIES);
  pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
  pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.REPORTS);
  pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.DASHBOARD);
  pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.HELP);
  pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.BUDGETS);
  pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.RECURRING_TRANSACTIONS);
  pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.CATEGORY_RULES);

  // Headers for Transactions (first row).
  var txSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  pfEnsureHeaderRow_(txSheet, PF_TRANSACTIONS_SCHEMA);

  // Reference sheets headers.
  var accountsSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
  pfEnsureHeaderRow_(accountsSheet, PF_ACCOUNTS_SCHEMA);

  var categoriesSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.CATEGORIES);
  pfEnsureHeaderRow_(categoriesSheet, PF_CATEGORIES_SCHEMA);

  var budgetsSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.BUDGETS);
  pfEnsureHeaderRow_(budgetsSheet, PF_BUDGETS_SCHEMA);

  var recurringSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.RECURRING_TRANSACTIONS);
  pfEnsureHeaderRow_(recurringSheet, PF_RECURRING_TRANSACTIONS_SCHEMA);

  var categoryRulesSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.CATEGORY_RULES);
  pfEnsureHeaderRow_(categoryRulesSheet, PF_CATEGORY_RULES_SCHEMA);
}

/**
 * Ensure header row for a schema-driven sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {{columns: Array<{key: string}>}} schema
 */
function pfEnsureHeaderRow_(sheet, schema) {
  var lang = pfGetLanguage_();
  var headers = [];
  for (var i = 0; i < schema.columns.length; i++) {
    headers.push(pfT_('columns.' + schema.columns[i].key, lang));
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

/**
 * Ensure a filter exists on header row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} numColumns
 */
function pfEnsureFilter_(sheet, numColumns) {
  if (sheet.getFilter()) return;
  sheet.getRange(1, 1, 1, numColumns).createFilter();
}

/**
 * Upsert named range.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 */
function pfUpsertNamedRange_(ss, name, range) {
  var existing = ss.getRangeByName(name);
  if (existing) {
    // Remove existing named range.
    ss.removeNamedRange(name);
  }
  ss.setNamedRange(name, range);
}

/**
 * Safely clear sheet rows: content, formatting, and notes.
 * Attempts to delete rows, falls back to clearing content if deletion fails.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to clear
 * @param {number} startRow - First row to clear (1-based, typically 2 to skip header)
 * @param {number} numRows - Number of rows to clear
 * @returns {boolean} True if successful, false otherwise
 */
function pfClearSheetRows_(sheet, startRow, numRows) {
  if (!sheet || startRow < 1 || numRows < 1) {
    Logger.log('[SHEETS] Invalid parameters for pfClearSheetRows_: sheet=' + sheet + ', startRow=' + startRow + ', numRows=' + numRows);
    return false;
  }
  
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    Logger.log('[SHEETS] Sheet has no columns, nothing to clear');
    return true; // Nothing to clear
  }
  
  var clearRange = sheet.getRange(startRow, 1, numRows, lastCol);
  
  try {
    // Clear everything: content, formatting, and notes
    clearRange.clearContent();
    clearRange.clearFormat();
    clearRange.clearNote();
    
    // Try to delete rows
    try {
      sheet.deleteRows(startRow, numRows);
      Logger.log('[SHEETS] Successfully cleared and deleted ' + numRows + ' rows starting at ' + startRow);
      return true;
    } catch (e) {
      Logger.log('[SHEETS] WARNING: Could not delete rows, but content cleared: ' + e.toString());
      // Content is already cleared, so this is acceptable
      return true;
    }
  } catch (e) {
    Logger.log('[SHEETS] ERROR: Could not clear sheet rows: ' + e.toString());
    Logger.log('[SHEETS] Error stack: ' + (e.stack || 'No stack'));
    return false;
  }
}

