/**
 * Demo data definitions and utilities.
 * 
 * Provides centralized demo data for accounts and categories.
 * Used by both Template.js (for template creation) and TestData.js (for test data generation).
 */

/**
 * Demo accounts data.
 * Format: [Name, Type, Currency, InitialBalance, IsActive]
 */
var PF_DEMO_ACCOUNTS = [
  ['Наличные', 'cash', 'RUB', 0, true],
  ['Карта', 'card', 'RUB', 0, true],
  ['Вклад', 'deposit', 'RUB', 0, true]
];

/**
 * Demo categories data.
 * Format: [Name, Type, Subcategory]
 */
var PF_DEMO_CATEGORIES = [
  ['Продукты', 'expense', ''],
  ['Транспорт', 'expense', ''],
  ['Кафе и рестораны', 'expense', ''],
  ['Здоровье', 'expense', ''],
  ['Развлечения', 'expense', ''],
  ['Зарплата', 'income', ''],
  ['Прочее', 'both', '']
];

/**
 * Get demo accounts data.
 * @returns {Array<Array>} Array of account rows
 */
function pfGetDemoAccounts_() {
  return PF_DEMO_ACCOUNTS.slice(); // Return copy
}

/**
 * Get demo categories data.
 * @returns {Array<Array>} Array of category rows
 */
function pfGetDemoCategories_() {
  return PF_DEMO_CATEGORIES.slice(); // Return copy
}

/**
 * Reset demo accounts to default values in a sheet.
 * Updates existing rows (up to number of demo accounts).
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Accounts sheet
 */
function pfResetDemoAccounts_(sheet) {
  var demoAccounts = pfGetDemoAccounts_();
  var lastRow = sheet.getLastRow();
  var numToUpdate = Math.min(demoAccounts.length, lastRow - 1);
  
  for (var i = 0; i < numToUpdate; i++) {
    var row = sheet.getRange(i + 2, 1, 1, demoAccounts[i].length);
    row.setValues([demoAccounts[i]]);
  }
  
  pfLogInfo_('Reset ' + numToUpdate + ' demo accounts', 'pfResetDemoAccounts_');
}

/**
 * Create demo accounts if sheet is empty.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Accounts sheet
 */
function pfCreateDemoAccounts_(sheet) {
  var demoAccounts = pfGetDemoAccounts_();
  
  if (sheet.getLastRow() <= 1) {
    var range = sheet.getRange(2, 1, demoAccounts.length, demoAccounts[0].length);
    range.setValues(demoAccounts);
    pfLogInfo_('Created ' + demoAccounts.length + ' demo accounts', 'pfCreateDemoAccounts_');
  }
}

/**
 * Reset demo categories to default values in a sheet.
 * Updates existing rows (up to number of demo categories).
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Categories sheet
 */
function pfResetDemoCategories_(sheet) {
  var demoCategories = pfGetDemoCategories_();
  var lastRow = sheet.getLastRow();
  var numToUpdate = Math.min(demoCategories.length, lastRow - 1);
  
  for (var i = 0; i < numToUpdate; i++) {
    var row = sheet.getRange(i + 2, 1, 1, demoCategories[i].length);
    row.setValues([demoCategories[i]]);
  }
  
  pfLogInfo_('Reset ' + numToUpdate + ' demo categories', 'pfResetDemoCategories_');
}

/**
 * Create demo categories if sheet is empty.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Categories sheet
 */
function pfCreateDemoCategories_(sheet) {
  var demoCategories = pfGetDemoCategories_();
  
  if (sheet.getLastRow() <= 1) {
    var range = sheet.getRange(2, 1, demoCategories.length, demoCategories[0].length);
    range.setValues(demoCategories);
    pfLogInfo_('Created ' + demoCategories.length + ' demo categories', 'pfCreateDemoCategories_');
  }
}
