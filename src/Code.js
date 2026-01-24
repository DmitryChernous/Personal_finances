/**
 * Точка входа: добавляем меню в Google Sheets.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Personal finances')
    .addItem('Setup (создать листы)', 'pfSetup')
    .addToUi();
}

/**
 * Минимальный setup: создаём листы-заготовки (позже расширим).
 */
function pfSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet_(ss, 'Transactions');
  getOrCreateSheet_(ss, 'Categories');
  getOrCreateSheet_(ss, 'Accounts');
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert('Готово: листы созданы (или уже существовали).');
}

