/**
 * Validation and normalization for transactions.
 *
 * Provides:
 * - Row-level validation (required fields, business rules)
 * - Normalization (trim, type conversion)
 * - Error highlighting
 * - Mark for review functionality
 */

/**
 * Validates a transaction row and returns errors.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowNum 1-based row number
 * @returns {Array<{column: string, message: string}>} Array of errors
 */
function pfValidateTransactionRow_(sheet, rowNum) {
  var errors = [];
  var row = sheet.getRange(rowNum, 1, 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues()[0];
  var typeCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
  var typeValue = typeCol ? String(row[typeCol - 1] || '').trim() : '';

  // Required fields (always).
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  if (dateCol && (!row[dateCol - 1] || row[dateCol - 1] === '')) {
    errors.push({ column: 'Date', message: 'Дата обязательна' });
  }

  if (!typeValue) {
    errors.push({ column: 'Type', message: 'Тип обязателен' });
  }

  var accountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Account');
  if (accountCol && (!row[accountCol - 1] || String(row[accountCol - 1]).trim() === '')) {
    errors.push({ column: 'Account', message: 'Счет обязателен' });
  }

  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  if (amountCol) {
    var amount = row[amountCol - 1];
    if (!amount || amount <= 0 || isNaN(amount)) {
      errors.push({ column: 'Amount', message: 'Сумма должна быть положительным числом' });
    }
  }

  var currencyCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Currency');
  if (currencyCol && (!row[currencyCol - 1] || String(row[currencyCol - 1]).trim() === '')) {
    errors.push({ column: 'Currency', message: 'Валюта обязательна' });
  }

  var sourceCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Source');
  if (sourceCol && (!row[sourceCol - 1] || String(row[sourceCol - 1]).trim() === '')) {
    errors.push({ column: 'Source', message: 'Источник обязателен' });
  }

  var statusCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
  if (statusCol && (!row[statusCol - 1] || String(row[statusCol - 1]).trim() === '')) {
    errors.push({ column: 'Status', message: 'Статус обязателен' });
  }

  // Business rules based on Type.
  if (typeValue === PF_TRANSACTION_TYPE.TRANSFER) {
    var accountToCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'AccountTo');
    if (accountToCol && (!row[accountToCol - 1] || String(row[accountToCol - 1]).trim() === '')) {
      errors.push({ column: 'AccountTo', message: 'Для перевода обязателен счет получателя' });
    }
    // Transfer should not have Category.
    var categoryCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category');
    if (categoryCol && row[categoryCol - 1] && String(row[categoryCol - 1]).trim() !== '') {
      errors.push({ column: 'Category', message: 'Перевод не должен иметь категорию' });
    }
  } else if (typeValue === 'expense' || typeValue === 'income') {
    // Expense/Income should have Category (recommended, not strict).
    var categoryCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category');
    if (categoryCol && (!row[categoryCol - 1] || String(row[categoryCol - 1]).trim() === '')) {
      // Warning, not error (optional field).
    }
  }

  return errors;
}

/**
 * Normalizes a transaction row (trim strings, set defaults).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowNum 1-based row number
 */
function pfNormalizeTransactionRow_(sheet, rowNum) {
  var row = sheet.getRange(rowNum, 1, 1, PF_TRANSACTIONS_SCHEMA.columns.length);
  var values = row.getValues()[0];
  var needsUpdate = false;
  var ss = sheet.getParent();

  // Trim string fields.
  for (var i = 0; i < PF_TRANSACTIONS_SCHEMA.columns.length; i++) {
    var col = PF_TRANSACTIONS_SCHEMA.columns[i];
    var colIndex = i + 1;
    var value = values[i];

    if (typeof value === 'string' && value !== value.trim()) {
      values[i] = value.trim();
      needsUpdate = true;
    }

    // Auto-fill defaults for new rows.
    if (rowNum > 1) {
      var sourceCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Source');
      if (col.key === 'Source' && sourceCol === colIndex && (!value || String(value).trim() === '')) {
        values[i] = PF_IMPORT_SOURCE.MANUAL;
        needsUpdate = true;
      }

      var statusCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
      if (col.key === 'Status' && statusCol === colIndex && (!value || String(value).trim() === '')) {
        values[i] = 'ok';
        needsUpdate = true;
      }

      var currencyCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Currency');
      if (col.key === 'Currency' && currencyCol === colIndex && (!value || String(value).trim() === '')) {
        var defaultCurrency = pfGetSetting_(ss, PF_SETUP_KEYS.DEFAULT_CURRENCY) || PF_DEFAULT_CURRENCY;
        values[i] = defaultCurrency;
        needsUpdate = true;
      }
    }
  }

  if (needsUpdate) {
    row.setValues([values]);
  }
}

/**
 * Highlights errors in a row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowNum 1-based row number
 * @param {Array<{column: string, message: string}>} errors
 */
function pfHighlightErrors_(sheet, rowNum, errors) {
  if (errors.length === 0) {
    // Clear highlighting if no errors.
    var row = sheet.getRange(rowNum, 1, 1, PF_TRANSACTIONS_SCHEMA.columns.length);
    row.setBackground(null);
    return;
  }

  // Highlight entire row in light red.
  var row = sheet.getRange(rowNum, 1, 1, PF_TRANSACTIONS_SCHEMA.columns.length);
  row.setBackground('#ffcccc');

  // Optionally highlight specific error columns in darker red.
  for (var i = 0; i < errors.length; i++) {
    var colIndex = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, errors[i].column);
    if (colIndex) {
      sheet.getRange(rowNum, colIndex).setBackground('#ff9999');
    }
  }
}

/**
 * Marks selected rows for review.
 * Sets Status to 'needs_review' for all selected rows in Transactions sheet.
 */
function pfMarkSelectedForReview() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  var selection = ss.getActiveRange();

  if (!selection || selection.getSheet().getName() !== sheet.getName()) {
    SpreadsheetApp.getUi().alert('Выберите строки в листе "Транзакции"');
    return;
  }

  var statusCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
  if (!statusCol) return;

  var firstRow = selection.getRow();
  var lastRow = selection.getLastRow();

  // Only process data rows (skip header).
  if (firstRow <= 1) firstRow = 2;
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('Выберите строки с данными (не заголовок)');
    return;
  }

  var statusRange = sheet.getRange(firstRow, statusCol, lastRow - firstRow + 1, 1);
  var statusValues = statusRange.getValues();
  var updated = false;

  for (var i = 0; i < statusValues.length; i++) {
    if (statusValues[i][0] !== 'needs_review') {
      statusValues[i][0] = 'needs_review';
      updated = true;
    }
  }

  if (updated) {
    statusRange.setValues(statusValues);
    SpreadsheetApp.getUi().alert('Строки помечены на проверку');
  } else {
    SpreadsheetApp.getUi().alert('Выбранные строки уже помечены на проверку');
  }
}

/**
 * Validates all rows in Transactions sheet and highlights errors.
 */
function pfValidateAllTransactions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('Нет данных для проверки');
    return;
  }

  var errorCount = 0;
  var rowsWithErrors = [];

  for (var row = 2; row <= lastRow; row++) {
    var errors = pfValidateTransactionRow_(sheet, row);
    if (errors.length > 0) {
      errorCount += errors.length;
      rowsWithErrors.push({ row: row, errors: errors });
      pfHighlightErrors_(sheet, row, errors);
    } else {
      pfHighlightErrors_(sheet, row, []);
    }
  }

  if (errorCount > 0) {
    SpreadsheetApp.getUi().alert('Найдено ошибок: ' + errorCount + ' в ' + rowsWithErrors.length + ' строках');
  } else {
    SpreadsheetApp.getUi().alert('Все строки валидны');
  }
}
