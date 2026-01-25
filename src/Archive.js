/**
 * Archive module.
 * Handles archiving of old transactions to improve performance.
 */

/**
 * Create or get archive sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Archive sheet
 */
function pfCreateArchiveSheet_(ss) {
  if (!ss) {
    throw new Error('Spreadsheet is required');
  }
  
  // Check if archive sheet exists
  var archiveSheetName = 'Transactions_Archive';
  var archiveSheet = ss.getSheetByName(archiveSheetName);
  
  if (!archiveSheet) {
    // Create new archive sheet
    archiveSheet = ss.insertSheet(archiveSheetName);
    
    // Copy structure from Transactions sheet
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    if (txSheet) {
      // Copy headers
      var headers = txSheet.getRange(1, 1, 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues()[0];
      archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Copy header formatting
      var headerRange = txSheet.getRange(1, 1, 1, PF_TRANSACTIONS_SCHEMA.columns.length);
      var archiveHeaderRange = archiveSheet.getRange(1, 1, 1, PF_TRANSACTIONS_SCHEMA.columns.length);
      headerRange.copyTo(archiveHeaderRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      
      // Freeze header row
      archiveSheet.setFrozenRows(1);
      
      // Apply localization to headers
      pfEnsureHeaderRow_(archiveSheet, PF_TRANSACTIONS_SCHEMA);
    }
  }
  
  return archiveSheet;
}

/**
 * Archive old transactions (mark as deleted and copy to archive).
 * @param {Date} cutoffDate - Transactions older than this date will be archived
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet
 * @returns {Object} Statistics: {archived: number, errors: number, message: string}
 */
function pfArchiveOldTransactions_(cutoffDate, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return pfCreateErrorResponse_('Cannot get spreadsheet');
  }
  
  if (!cutoffDate || !(cutoffDate instanceof Date)) {
    return pfCreateErrorResponse_('Invalid cutoff date');
  }
  
  try {
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    if (!txSheet) {
      return pfCreateErrorResponse_('Transactions sheet not found');
    }
    
    var archiveSheet = pfCreateArchiveSheet_(ss);
    
    // Get column indices
    var dateColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
    var statusColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
    
    if (!dateColIdx || !statusColIdx) {
      return pfCreateErrorResponse_('Required columns not found');
    }
    
    // Cache lastRow before reading
    var lastRow = txSheet.getLastRow();
    if (lastRow <= 1) {
      return pfCreateSuccessResponse_('No transactions to archive', { archived: 0, errors: 0 });
    }
    
    // Read all data (skip header)
    var data = txSheet.getRange(2, 1, lastRow - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
    
    var rowsToArchive = [];
    var rowsToMarkDeleted = [];
    
    // Find rows to archive
    for (var i = 0; i < data.length; i++) {
      var rowData = data[i];
      var rowNum = i + 2; // 1-based row number (skip header)
      
      if (rowData.length <= dateColIdx) {
        continue;
      }
      
      var date = rowData[dateColIdx - 1];
      var status = rowData[statusColIdx - 1];
      
      // Skip if already deleted or archived
      if (status === PF_TRANSACTION_STATUS.DELETED) {
        continue;
      }
      
      // Check if date is before cutoff
      if (date instanceof Date && date < cutoffDate) {
        rowsToArchive.push({
          rowNum: rowNum,
          data: rowData
        });
        rowsToMarkDeleted.push(rowNum);
      }
    }
    
    if (rowsToArchive.length === 0) {
      return pfCreateSuccessResponse_('No transactions to archive', { archived: 0, errors: 0 });
    }
    
    // Copy to archive (batch operation)
    var archiveData = [];
    for (var j = 0; j < rowsToArchive.length; j++) {
      archiveData.push(rowsToArchive[j].data);
    }
    
    if (archiveData.length > 0) {
      var archiveLastRow = archiveSheet.getLastRow();
      var archiveTargetRow = archiveLastRow + 1;
      if (archiveLastRow === 0) {
        archiveTargetRow = 2; // Skip header
      }
      
      archiveSheet.getRange(archiveTargetRow, 1, archiveData.length, PF_TRANSACTIONS_SCHEMA.columns.length)
        .setValues(archiveData);
    }
    
    // Mark as deleted in Transactions sheet (batch operation)
    var deletedStatus = PF_TRANSACTION_STATUS.DELETED;
    var statusValues = [];
    for (var k = 0; k < rowsToMarkDeleted.length; k++) {
      statusValues.push([deletedStatus]);
    }
    
    if (statusValues.length > 0) {
      // Update status column in batches (to avoid timeout)
      var batchSize = 100;
      for (var b = 0; b < rowsToMarkDeleted.length; b += batchSize) {
        var batchRows = rowsToMarkDeleted.slice(b, Math.min(b + batchSize, rowsToMarkDeleted.length));
        var batchStatusValues = statusValues.slice(b, Math.min(b + batchSize, statusValues.length));
        
        for (var m = 0; m < batchRows.length; m++) {
          txSheet.getRange(batchRows[m], statusColIdx).setValue(deletedStatus);
        }
      }
    }
    
    pfLogInfo_('Archived ' + rowsToArchive.length + ' transactions older than ' + 
      Utilities.formatDate(cutoffDate, Session.getScriptTimeZone(), 'dd.MM.yyyy'), 'pfArchiveOldTransactions_');
    
    var lang = pfGetLanguage_();
    var message = lang === 'en' ? 
      'Archived ' + rowsToArchive.length + ' transactions' :
      'Заархивировано ' + rowsToArchive.length + ' транзакций';
    
    return pfCreateSuccessResponse_(message, {
      archived: rowsToArchive.length,
      errors: 0
    });
    
  } catch (e) {
    pfLogError_(e, 'pfArchiveOldTransactions_', PF_LOG_LEVEL.ERROR);
    return pfCreateErrorResponse_('Error archiving transactions: ' + (e.message || e.toString()));
  }
}

/**
 * Count transactions that would be archived (without archiving them).
 * @param {Date} cutoffDate - Cutoff date
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet
 * @returns {number} Count of transactions to archive
 */
function pfCountTransactionsToArchive_(cutoffDate, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || !cutoffDate || !(cutoffDate instanceof Date)) {
    return 0;
  }
  
  try {
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    if (!txSheet) {
      return 0;
    }
    
    var dateColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
    var statusColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
    
    if (!dateColIdx || !statusColIdx) {
      return 0;
    }
    
    var lastRow = txSheet.getLastRow();
    if (lastRow <= 1) {
      return 0;
    }
    
    var data = txSheet.getRange(2, 1, lastRow - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
    var count = 0;
    
    for (var i = 0; i < data.length; i++) {
      var rowData = data[i];
      if (rowData.length <= dateColIdx) {
        continue;
      }
      
      var date = rowData[dateColIdx - 1];
      var status = rowData[statusColIdx - 1];
      
      if (status === PF_TRANSACTION_STATUS.DELETED) {
        continue;
      }
      
      if (date instanceof Date && date < cutoffDate) {
        count++;
      }
    }
    
    return count;
  } catch (e) {
    pfLogError_(e, 'pfCountTransactionsToArchive_', PF_LOG_LEVEL.ERROR);
    return 0;
  }
}

/**
 * Restore transactions from archive (optional function).
 * @param {Date} startDate - Start date (optional)
 * @param {Date} endDate - End date (optional)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet
 * @returns {Object} Statistics: {restored: number, errors: number}
 */
function pfRestoreFromArchive_(startDate, endDate, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return pfCreateErrorResponse_('Cannot get spreadsheet');
  }
  
  try {
    var archiveSheet = ss.getSheetByName('Transactions_Archive');
    if (!archiveSheet || archiveSheet.getLastRow() <= 1) {
      return pfCreateSuccessResponse_('Archive is empty', { restored: 0, errors: 0 });
    }
    
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    if (!txSheet) {
      return pfCreateErrorResponse_('Transactions sheet not found');
    }
    
    // Get column indices
    var dateColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
    if (!dateColIdx) {
      return pfCreateErrorResponse_('Date column not found');
    }
    
    var lastRow = archiveSheet.getLastRow();
    var data = archiveSheet.getRange(2, 1, lastRow - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
    
    var rowsToRestore = [];
    
    // Find rows to restore
    for (var i = 0; i < data.length; i++) {
      var rowData = data[i];
      
      if (rowData.length <= dateColIdx) {
        continue;
      }
      
      var date = rowData[dateColIdx - 1];
      
      // Filter by date range if provided
      if (startDate && date instanceof Date && date < startDate) {
        continue;
      }
      if (endDate && date instanceof Date && date > endDate) {
        continue;
      }
      
      rowsToRestore.push(rowData);
    }
    
    if (rowsToRestore.length === 0) {
      return pfCreateSuccessResponse_('No transactions to restore', { restored: 0, errors: 0 });
    }
    
    // Copy back to Transactions sheet
    var txLastRow = txSheet.getLastRow();
    var targetRow = txLastRow + 1;
    if (txLastRow === 0) {
      targetRow = 2; // Skip header
    }
    
    txSheet.getRange(targetRow, 1, rowsToRestore.length, PF_TRANSACTIONS_SCHEMA.columns.length)
      .setValues(rowsToRestore);
    
    // Update status to 'ok' for restored transactions
    var statusColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
    if (statusColIdx) {
      var statusRange = txSheet.getRange(targetRow, statusColIdx, rowsToRestore.length, 1);
      var statusValues = [];
      for (var j = 0; j < rowsToRestore.length; j++) {
        statusValues.push([PF_TRANSACTION_STATUS.OK]);
      }
      statusRange.setValues(statusValues);
    }
    
    pfLogInfo_('Restored ' + rowsToRestore.length + ' transactions from archive', 'pfRestoreFromArchive_');
    
    var lang = pfGetLanguage_();
    var message = lang === 'en' ? 
      'Restored ' + rowsToRestore.length + ' transactions' :
      'Восстановлено ' + rowsToRestore.length + ' транзакций';
    
    return pfCreateSuccessResponse_(message, {
      restored: rowsToRestore.length,
      errors: 0
    });
    
  } catch (e) {
    pfLogError_(e, 'pfRestoreFromArchive_', PF_LOG_LEVEL.ERROR);
    return pfCreateErrorResponse_('Error restoring transactions: ' + (e.message || e.toString()));
  }
}

/**
 * Public function: Archive old transactions (called from menu).
 * Shows dialog for date selection and confirmation.
 */
function pfArchiveOldTransactions() {
  var lang = pfGetLanguage_();
  var ui = SpreadsheetApp.getUi();
  
  // Prompt for cutoff date
  var datePrompt = lang === 'en' ? 
    'Enter cutoff date (transactions older than this date will be archived).\nFormat: DD.MM.YYYY or YYYY-MM-DD' :
    'Введите дату отсечки (транзакции старше этой даты будут заархивированы).\nФормат: ДД.ММ.ГГГГ или ГГГГ-ММ-ДД';
  
  var response = ui.prompt(
    lang === 'en' ? 'Archive Old Transactions' : 'Архивирование старых транзакций',
    datePrompt,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var dateInput = response.getResponseText().trim();
  if (!dateInput) {
    ui.alert(
      lang === 'en' ? 'Error' : 'Ошибка',
      lang === 'en' ? 'Date is required' : 'Дата обязательна',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Parse date
  var cutoffDate = null;
  try {
    // Try DD.MM.YYYY format
    if (dateInput.indexOf('.') !== -1) {
      var parts = dateInput.split('.');
      if (parts.length === 3) {
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10) - 1; // Month is 0-based
        var year = parseInt(parts[2], 10);
        cutoffDate = new Date(year, month, day);
      }
    }
    // Try YYYY-MM-DD format
    else if (dateInput.indexOf('-') !== -1) {
      cutoffDate = new Date(dateInput);
    }
    // Try other formats
    else {
      cutoffDate = new Date(dateInput);
    }
    
    if (!cutoffDate || isNaN(cutoffDate.getTime())) {
      throw new Error('Invalid date');
    }
  } catch (e) {
    ui.alert(
      lang === 'en' ? 'Error' : 'Ошибка',
      lang === 'en' ? 'Invalid date format' : 'Неверный формат даты',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Count transactions to archive
  var count = pfCountTransactionsToArchive_(cutoffDate);
  if (count === 0) {
    ui.alert(
      lang === 'en' ? 'No Transactions' : 'Нет транзакций',
      lang === 'en' ? 'No transactions found to archive' : 'Не найдено транзакций для архивирования',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Confirm archiving
  var confirmMessage = lang === 'en' ? 
    'This will archive ' + count + ' transactions older than ' + 
    Utilities.formatDate(cutoffDate, Session.getScriptTimeZone(), 'dd.MM.yyyy') + 
    '.\n\nTransactions will be marked as deleted and copied to the archive sheet.\n\nContinue?' :
    'Будет заархивировано ' + count + ' транзакций старше ' + 
    Utilities.formatDate(cutoffDate, Session.getScriptTimeZone(), 'dd.MM.yyyy') + 
    '.\n\nТранзакции будут помечены как удаленные и скопированы в архивный лист.\n\nПродолжить?';
  
  var confirmResponse = ui.alert(
    lang === 'en' ? 'Confirm Archive' : 'Подтверждение архивирования',
    confirmMessage,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResponse !== ui.Button.YES) {
    return;
  }
  
  // Perform archiving
  var result = pfArchiveOldTransactions_(cutoffDate);
  
  if (result.success) {
    ui.alert(
      lang === 'en' ? 'Archive Complete' : 'Архивирование завершено',
      result.message,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      lang === 'en' ? 'Archive Error' : 'Ошибка архивирования',
      result.message,
      ui.ButtonSet.OK
    );
  }
}
