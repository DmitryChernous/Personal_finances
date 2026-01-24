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

  // Headers for Transactions (first row).
  var txSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  pfEnsureHeaderRow_(txSheet, PF_TRANSACTIONS_SCHEMA);
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

