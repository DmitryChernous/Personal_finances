/**
 * Settings storage.
 *
 * For portability, settings are stored in a dedicated sheet (not ScriptProperties).
 * The sheet is a simple key/value table: column A = key, column B = value.
 */

var PF_SETTINGS_KEYS = {
  LANGUAGE: 'Language'
};

/**
 * Get current language (ru/en).
 * @returns {string}
 */
function pfGetLanguage_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // IMPORTANT: do not call pfGetSetting_ here, because it may create/rename the
  // Settings sheet, which itself depends on language selection.
  var sheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.SETTINGS);
  if (!sheet) return PF_DEFAULT_LANG;

  var lang = pfGetSettingFromSheet_(sheet, PF_SETTINGS_KEYS.LANGUAGE);
  if (!lang) return PF_DEFAULT_LANG;

  var normalized = String(lang).trim().toLowerCase();
  if (PF_SUPPORTED_LANGS.indexOf(normalized) === -1) return PF_DEFAULT_LANG;
  return normalized;
}

/**
 * Set language and apply localization (sheet names + headers).
 * @param {string} lang
 */
function pfSetLanguage_(lang) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var normalized = String(lang || '').trim().toLowerCase();
  if (PF_SUPPORTED_LANGS.indexOf(normalized) === -1) {
    SpreadsheetApp.getUi().alert('Неизвестный язык: ' + lang);
    return;
  }

  pfSetSetting_(ss, PF_SETTINGS_KEYS.LANGUAGE, normalized);
  pfApplyLocalization_(ss);
  ss.toast('Язык установлен: ' + normalized, 'Personal finances', 5);
}

function pfSetLanguageRu() { pfSetLanguage_('ru'); }
function pfSetLanguageEn() { pfSetLanguage_('en'); }

/**
 * Ensure Settings sheet exists and has header row.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function pfEnsureSettingsSheet_(ss) {
  var sheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.SETTINGS);

  // Minimal header row: Key | Value
  var range = sheet.getRange(1, 1, 1, 2);
  var values = range.getValues();
  if (!values || !values[0] || (!values[0][0] && !values[0][1])) {
    range.setValues([['Key', 'Value']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Get setting value from Settings sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} key
 * @returns {string|null}
 */
function pfGetSetting_(ss, key) {
  var sheet = pfEnsureSettingsSheet_(ss);
  return pfGetSettingFromSheet_(sheet, key);
}

/**
 * Get setting value from a given Settings sheet (no creation/rename).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} key
 * @returns {string|null}
 */
function pfGetSettingFromSheet_(sheet, key) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      var v = values[i][1];
      return v == null ? null : String(v);
    }
  }
  return null;
}

/**
 * Upsert setting in Settings sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} key
 * @param {string} value
 */
function pfSetSetting_(ss, key, value) {
  var sheet = pfEnsureSettingsSheet_(ss);
  var lastRow = sheet.getLastRow();
  var normalizedValue = value == null ? '' : String(value);

  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === key) {
        sheet.getRange(i + 2, 2).setValue(normalizedValue);
        return;
      }
    }
  }

  sheet.appendRow([key, normalizedValue]);
}

