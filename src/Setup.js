/**
 * Setup / initialization routines.
 *
 * Goal: make `pfSetup()` idempotent and safe to rerun.
 */

var PF_NAMED_RANGES = {
  ACCOUNTS: 'PF_ACCOUNTS',
  CATEGORIES: 'PF_CATEGORIES'
};

var PF_DEFAULT_CURRENCIES = ['RUB', 'USD', 'EUR'];

var PF_SETUP_KEYS = {
  SCHEMA_VERSION: 'SchemaVersion',
  DEFAULT_CURRENCY: 'DefaultCurrency'
};

/**
 * Main setup entry.
 */
function pfRunSetup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Project-wide locale conventions (separate from UI language).
  ss.setSpreadsheetLocale('ru_RU');

  // Ensure Settings exists early (uses current language, default RU).
  pfEnsureSettingsSheet_(ss);

  // Ensure language has a value (project default is RU).
  if (!pfGetSetting_(ss, PF_SETTINGS_KEYS.LANGUAGE)) {
    pfSetSetting_(ss, PF_SETTINGS_KEYS.LANGUAGE, PF_DEFAULT_LANG);
  }

  // Track schema version and default currency.
  if (!pfGetSetting_(ss, PF_SETUP_KEYS.SCHEMA_VERSION)) {
    pfSetSetting_(ss, PF_SETUP_KEYS.SCHEMA_VERSION, String(PF_SCHEMA_VERSION));
  }
  if (!pfGetSetting_(ss, PF_SETUP_KEYS.DEFAULT_CURRENCY)) {
    pfSetSetting_(ss, PF_SETUP_KEYS.DEFAULT_CURRENCY, 'RUB');
  }

  // Create/rename sheets and set headers according to selected language.
  pfApplyLocalization_(ss);

  // Ensure Import_Raw staging sheet exists
  pfEnsureImportRawSheet_(ss);

  // Apply filters, named ranges and validations.
  pfConfigureReferenceSheets_(ss);
  pfConfigureTransactionsSheet_(ss);
  pfEnsureHelpContent_(ss);

  // Initialize Reports sheet with formulas.
  pfInitializeReports_(ss);

  // Initialize Dashboard sheet with KPI and charts.
  pfInitializeDashboard_(ss);

  SpreadsheetApp.flush();
}

function pfConfigureReferenceSheets_(ss) {
  var accountsSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
  pfEnsureFilter_(accountsSheet, PF_ACCOUNTS_SCHEMA.columns.length);

  var categoriesSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.CATEGORIES);
  pfEnsureFilter_(categoriesSheet, PF_CATEGORIES_SCHEMA.columns.length);

  // Named ranges cover full columns below header (easy to append values).
  pfUpsertNamedRange_(ss, PF_NAMED_RANGES.ACCOUNTS, accountsSheet.getRange('A2:A'));
  pfUpsertNamedRange_(ss, PF_NAMED_RANGES.CATEGORIES, categoriesSheet.getRange('A2:A'));
}

function pfConfigureTransactionsSheet_(ss) {
  var sheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  var numCols = PF_TRANSACTIONS_SCHEMA.columns.length;
  pfEnsureFilter_(sheet, numCols);

  // Formats.
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  if (dateCol) sheet.getRange(2, dateCol, sheet.getMaxRows() - 1, 1).setNumberFormat('dd.mm.yyyy');
  if (amountCol) sheet.getRange(2, amountCol, sheet.getMaxRows() - 1, 1).setNumberFormat('0.00');

  // Validations (lightweight, allow blanks where it makes sense).
  var typeCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
  var statusCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
  var currencyCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Currency');
  var accountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Account');
  var accountToCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'AccountTo');
  var categoryCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category');

  if (typeCol) {
    var ruleType = SpreadsheetApp.newDataValidation()
      .requireValueInList([PF_TRANSACTION_TYPE.EXPENSE, PF_TRANSACTION_TYPE.INCOME, PF_TRANSACTION_TYPE.TRANSFER], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, typeCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleType);
  }

  if (statusCol) {
    var ruleStatus = SpreadsheetApp.newDataValidation()
      .requireValueInList(['ok', 'needs_review', 'duplicate', 'deleted'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, statusCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleStatus);
  }

  if (currencyCol) {
    var ruleCurrency = SpreadsheetApp.newDataValidation()
      .requireValueInList(PF_DEFAULT_CURRENCIES, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, currencyCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleCurrency);
  }

  var accountsRange = ss.getRangeByName(PF_NAMED_RANGES.ACCOUNTS);
  if (accountsRange && accountCol) {
    var ruleAccount = SpreadsheetApp.newDataValidation()
      .requireValueInRange(accountsRange, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, accountCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleAccount);
  }

  if (accountsRange && accountToCol) {
    var ruleAccountTo = SpreadsheetApp.newDataValidation()
      .requireValueInRange(accountsRange, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, accountToCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleAccountTo);
  }

  var categoriesRange = ss.getRangeByName(PF_NAMED_RANGES.CATEGORIES);
  if (categoriesRange && categoryCol) {
    var ruleCategory = SpreadsheetApp.newDataValidation()
      .requireValueInRange(categoriesRange, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, categoryCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleCategory);
  }

  if (amountCol) {
    var ruleAmount = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, amountCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleAmount);
  }
}

// Help content functions moved to Help.js
// pfEnsureHelpContent_, _writeHelpContentRu_, _writeHelpContentEn_
// are now defined in Help.js to improve modularity

/**
 * @param {{columns: Array<{key: string}>}} schema
 * @param {string} key
 * @returns {number|null} 1-based index
 */
function pfColumnIndex_(schema, key) {
  for (var i = 0; i < schema.columns.length; i++) {
    if (schema.columns[i].key === key) return i + 1;
  }
  return null;
}

