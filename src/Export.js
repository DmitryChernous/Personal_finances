/**
 * Export and Backup module.
 * Handles export of transactions and data to CSV/JSON formats,
 * and creation of backup copies.
 */

/**
 * Export transactions to CSV or JSON format.
 * @param {string} format - Format: 'csv' or 'json'
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet
 * @returns {GoogleAppsScript.Base.Blob} Blob with file content
 */
function pfExportTransactions_(format, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Cannot get spreadsheet');
  }
  
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet) {
    throw new Error('Лист транзакций не найден');
  }
  
  var lastRow = txSheet.getLastRow();
  if (lastRow <= 1) {
    throw new Error('Нет данных для экспорта');
  }
  
  var data = txSheet.getRange(1, 1, lastRow, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  if (format === 'csv') {
    return pfExportToCSV_(headers, rows);
  } else if (format === 'json') {
    return pfExportToJSON_(headers, rows);
  } else {
    throw new Error('Неверный формат: ' + format + '. Используйте "csv" или "json"');
  }
}

/**
 * Export data to CSV format.
 * @private
 * @param {Array<string>} headers - Column headers
 * @param {Array<Array>} rows - Data rows
 * @returns {GoogleAppsScript.Base.Blob} CSV blob
 */
function pfExportToCSV_(headers, rows) {
  // Determine delimiter based on locale (ru_RU uses semicolon)
  var delimiter = ';'; // ru_RU locale uses semicolon
  
  var csvLines = [];
  
  // Add headers
  var headerLine = headers.map(function(header) {
    return pfEscapeCSVField_(String(header || ''));
  }).join(delimiter);
  csvLines.push(headerLine);
  
  // Add data rows
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var csvRow = [];
    
    for (var j = 0; j < headers.length; j++) {
      var value = row[j];
      var csvValue = '';
      
      if (value instanceof Date) {
        // Format date as dd.MM.yyyy (ru_RU format)
        var day = value.getDate();
        var month = value.getMonth() + 1;
        var year = value.getFullYear();
        csvValue = String(day).padStart(2, '0') + '.' + 
                   String(month).padStart(2, '0') + '.' + 
                   String(year);
      } else if (typeof value === 'number') {
        // Format number (use dot as decimal separator for CSV compatibility)
        csvValue = String(value).replace(',', '.');
      } else if (value === null || value === undefined) {
        csvValue = '';
      } else {
        csvValue = String(value);
      }
      
      csvRow.push(pfEscapeCSVField_(csvValue));
    }
    
    csvLines.push(csvRow.join(delimiter));
  }
  
  var csvContent = csvLines.join('\n');
  var blob = Utilities.newBlob(csvContent, 'text/csv;charset=utf-8', 'transactions.csv');
  
  return blob;
}

/**
 * Escape CSV field (handle quotes and special characters).
 * @private
 * @param {string} field - Field value
 * @returns {string} Escaped field
 */
function pfEscapeCSVField_(field) {
  if (field === null || field === undefined) {
    return '';
  }
  
  var str = String(field);
  
  // If field contains delimiter, quotes, or newlines, wrap in quotes and escape quotes
  if (str.indexOf(';') !== -1 || str.indexOf(',') !== -1 || 
      str.indexOf('"') !== -1 || str.indexOf('\n') !== -1 || 
      str.indexOf('\r') !== -1) {
    // Escape quotes by doubling them
    str = str.replace(/"/g, '""');
    // Wrap in quotes
    str = '"' + str + '"';
  }
  
  return str;
}

/**
 * Export data to JSON format.
 * @private
 * @param {Array<string>} headers - Column headers
 * @param {Array<Array>} rows - Data rows
 * @returns {GoogleAppsScript.Base.Blob} JSON blob
 */
function pfExportToJSON_(headers, rows) {
  var transactions = [];
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var tx = {};
    
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      var value = row[j];
      
      // Convert Date objects to ISO strings
      if (value instanceof Date) {
        tx[header] = pfDateToISOString_(value);
      } else if (value === null || value === undefined) {
        tx[header] = null;
      } else {
        tx[header] = value;
      }
    }
    
    transactions.push(tx);
  }
  
  var jsonContent = JSON.stringify(transactions, null, 2);
  var blob = Utilities.newBlob(jsonContent, 'application/json;charset=utf-8', 'transactions.json');
  
  return blob;
}

/**
 * Export all data (Transactions, Accounts, Categories) to JSON.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet
 * @returns {GoogleAppsScript.Base.Blob} JSON blob with all data
 */
function pfExportAllData_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Cannot get spreadsheet');
  }
  
  var result = {
    exportDate: pfDateToISOString_(new Date()),
    transactions: [],
    accounts: [],
    categories: []
  };
  
  // Export Transactions
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (txSheet && txSheet.getLastRow() > 1) {
    var txLastRow = txSheet.getLastRow();
    var txData = txSheet.getRange(1, 1, txLastRow, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
    var txHeaders = txData[0];
    var txRows = txData.slice(1);
    
    for (var i = 0; i < txRows.length; i++) {
      var row = txRows[i];
      var tx = {};
      for (var j = 0; j < txHeaders.length; j++) {
        var value = row[j];
        if (value instanceof Date) {
          tx[txHeaders[j]] = pfDateToISOString_(value);
        } else {
          tx[txHeaders[j]] = value;
        }
      }
      result.transactions.push(tx);
    }
  }
  
  // Export Accounts
  var accountsSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
  if (accountsSheet && accountsSheet.getLastRow() > 1) {
    var accLastRow = accountsSheet.getLastRow();
    var accData = accountsSheet.getRange(1, 1, accLastRow, PF_ACCOUNTS_SCHEMA.columns.length).getValues();
    var accHeaders = accData[0];
    var accRows = accData.slice(1);
    
    for (var i = 0; i < accRows.length; i++) {
      var row = accRows[i];
      var acc = {};
      for (var j = 0; j < accHeaders.length; j++) {
        acc[accHeaders[j]] = row[j];
      }
      result.accounts.push(acc);
    }
  }
  
  // Export Categories
  var categoriesSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.CATEGORIES);
  if (categoriesSheet && categoriesSheet.getLastRow() > 1) {
    var catLastRow = categoriesSheet.getLastRow();
    var catData = categoriesSheet.getRange(1, 1, catLastRow, PF_CATEGORIES_SCHEMA.columns.length).getValues();
    var catHeaders = catData[0];
    var catRows = catData.slice(1);
    
    for (var i = 0; i < catRows.length; i++) {
      var row = catRows[i];
      var cat = {};
      for (var j = 0; j < catHeaders.length; j++) {
        cat[catHeaders[j]] = row[j];
      }
      result.categories.push(cat);
    }
  }
  
  var jsonContent = JSON.stringify(result, null, 2);
  var blob = Utilities.newBlob(jsonContent, 'application/json;charset=utf-8', 'all_data.json');
  
  return blob;
}

/**
 * Create backup copy of Transactions sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet
 * @returns {Object} Result: {success: boolean, message: string, sheetName: string}
 */
function pfCreateBackup_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return pfCreateErrorResponse_('Cannot get spreadsheet');
  }
  
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet) {
    return pfCreateErrorResponse_('Лист транзакций не найден');
  }
  
  try {
    // Generate backup sheet name with date
    var today = new Date();
    var dateStr = today.getFullYear() + 
      String(today.getMonth() + 1).padStart(2, '0') + 
      String(today.getDate()).padStart(2, '0');
    
    var baseName = 'Transactions_Backup_' + dateStr;
    var backupName = baseName;
    var counter = 1;
    
    // Check if sheet with this name already exists
    while (ss.getSheetByName(backupName)) {
      backupName = baseName + '_' + counter;
      counter++;
    }
    
    // Create backup sheet
    var backupSheet = ss.insertSheet(backupName);
    
    // Copy all data and formatting
    var lastRow = txSheet.getLastRow();
    var lastCol = txSheet.getLastColumn();
    
    if (lastRow > 0 && lastCol > 0) {
      var sourceRange = txSheet.getRange(1, 1, lastRow, lastCol);
      var targetRange = backupSheet.getRange(1, 1, lastRow, lastCol);
      
      // Copy values and formatting
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      
      // Copy formatting separately to ensure it's preserved
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      
      // Copy data validation
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      
      // Freeze header row if original has it
      if (txSheet.getFrozenRows() > 0) {
        backupSheet.setFrozenRows(1);
      }
      
      // Copy filter if exists
      if (txSheet.getFilter()) {
        backupSheet.getRange(1, 1, 1, lastCol).createFilter();
      }
    }
    
    pfLogInfo_('Created backup sheet: ' + backupName, 'pfCreateBackup_');
    
    return pfCreateSuccessResponse_('Резервная копия создана: ' + backupName, {
      sheetName: backupName
    });
    
  } catch (e) {
    pfLogError_(e, 'pfCreateBackup_', PF_LOG_LEVEL.ERROR);
    return pfCreateErrorResponse_('Ошибка при создании резервной копии: ' + (e.message || e.toString()));
  }
}

/**
 * Public function: Export transactions to CSV.
 * Creates file in Google Drive and shows link to user.
 */
function pfExportTransactionsCSV() {
  try {
    var blob = pfExportTransactions_('csv');
    
    // Create file in Google Drive
    var fileName = 'transactions_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.csv';
    var file = DriveApp.createFile(blob);
    file.setName(fileName);
    
    // Show link to user
    var lang = pfGetLanguage_();
    var message = lang === 'en' ? 
      'File created in Google Drive:\n' + file.getUrl() :
      'Файл создан в Google Drive:\n' + file.getUrl();
    
    SpreadsheetApp.getUi().alert(
      lang === 'en' ? 'Export Complete' : 'Экспорт завершен',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    pfLogError_(e, 'pfExportTransactionsCSV', PF_LOG_LEVEL.ERROR);
    var lang = pfGetLanguage_();
    SpreadsheetApp.getUi().alert(
      lang === 'en' ? 'Export Error' : 'Ошибка экспорта',
      e.message || e.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Public function: Export transactions to JSON.
 * Creates file in Google Drive and shows link to user.
 */
function pfExportTransactionsJSON() {
  try {
    var blob = pfExportTransactions_('json');
    
    // Create file in Google Drive
    var fileName = 'transactions_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.json';
    var file = DriveApp.createFile(blob);
    file.setName(fileName);
    
    // Show link to user
    var lang = pfGetLanguage_();
    var message = lang === 'en' ? 
      'File created in Google Drive:\n' + file.getUrl() :
      'Файл создан в Google Drive:\n' + file.getUrl();
    
    SpreadsheetApp.getUi().alert(
      lang === 'en' ? 'Export Complete' : 'Экспорт завершен',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    pfLogError_(e, 'pfExportTransactionsJSON', PF_LOG_LEVEL.ERROR);
    var lang = pfGetLanguage_();
    SpreadsheetApp.getUi().alert(
      lang === 'en' ? 'Export Error' : 'Ошибка экспорта',
      e.message || e.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Public function: Export all data to JSON.
 * Creates file in Google Drive and shows link to user.
 */
function pfExportAllDataJSON() {
  try {
    var blob = pfExportAllData_();
    
    // Create file in Google Drive
    var fileName = 'all_data_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.json';
    var file = DriveApp.createFile(blob);
    file.setName(fileName);
    
    // Show link to user
    var lang = pfGetLanguage_();
    var message = lang === 'en' ? 
      'File created in Google Drive:\n' + file.getUrl() :
      'Файл создан в Google Drive:\n' + file.getUrl();
    
    SpreadsheetApp.getUi().alert(
      lang === 'en' ? 'Export Complete' : 'Экспорт завершен',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    pfLogError_(e, 'pfExportAllDataJSON', PF_LOG_LEVEL.ERROR);
    var lang = pfGetLanguage_();
    SpreadsheetApp.getUi().alert(
      lang === 'en' ? 'Export Error' : 'Ошибка экспорта',
      e.message || e.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Public function: Create backup of Transactions sheet.
 * Shows result to user.
 */
function pfCreateBackup() {
  try {
    var result = pfCreateBackup_();
    
    var lang = pfGetLanguage_();
    if (result.success) {
      SpreadsheetApp.getUi().alert(
        lang === 'en' ? 'Backup Created' : 'Резервная копия создана',
        result.message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        lang === 'en' ? 'Backup Error' : 'Ошибка создания резервной копии',
        result.message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (e) {
    pfLogError_(e, 'pfCreateBackup', PF_LOG_LEVEL.ERROR);
    var lang = pfGetLanguage_();
    SpreadsheetApp.getUi().alert(
      lang === 'en' ? 'Backup Error' : 'Ошибка создания резервной копии',
      e.message || e.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
