/**
 * Template generation: create a clean template from current spreadsheet.
 * 
 * This module provides functionality to prepare the spreadsheet as a reusable template
 * by clearing transaction data while preserving structure and optionally keeping
 * demo reference data.
 */

/**
 * Main entry point: prepare template (clear data, keep structure).
 * Shows confirmation dialog and clears all transaction data, optionally keeps demo categories/accounts.
 */
function pfCreateTemplate() {
  // Валидация: проверяем, что таблица существует и доступна
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не удалось получить доступ к таблице', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  } catch (e) {
    pfLogError_(e, 'pfCreateTemplate', PF_LOG_LEVEL.ERROR);
    SpreadsheetApp.getUi().alert('Ошибка', 'Ошибка доступа к таблице: ' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    'Создание шаблона',
    'Эта операция очистит все транзакции и данные, оставив только структуру таблицы.\n\n' +
    'Что будет очищено:\n' +
    '• Все транзакции (останется только шапка)\n' +
    '• Все счета (можно оставить примеры)\n' +
    '• Все категории (можно оставить примеры)\n' +
    '• Данные в отчетах и дашборде\n' +
    '• Данные импорта\n\n' +
    'Что сохранится:\n' +
    '• Структура всех листов\n' +
    '• Форматы и валидации\n' +
    '• Настройки (язык, валюта)\n\n' +
    'Продолжить?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // Ask about keeping demo data
  var keepDemoResponse = ui.alert(
    'Демо-данные',
    'Оставить примеры счетов и категорий в справочниках?\n\n' +
    'Это поможет новым пользователям понять структуру данных.',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (keepDemoResponse === ui.Button.CANCEL) {
    return;
  }
  
  var keepDemoData = (keepDemoResponse === ui.Button.YES);
  
  try {
    Logger.log('[TEMPLATE] Starting template creation, keepDemoData: ' + keepDemoData);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    pfPrepareTemplate_(ss, keepDemoData);
    
    ui.alert(
      'Шаблон готов',
      'Таблица подготовлена как шаблон.\n\n' +
      'Все транзакции и данные очищены.\n' +
      (keepDemoData ? 'Примеры счетов и категорий оставлены.\n' : 'Справочники очищены.\n') +
      'Структура и настройки сохранены.',
      ui.ButtonSet.OK
    );
    
    Logger.log('[TEMPLATE] Template creation completed successfully');
  } catch (e) {
    Logger.log('[TEMPLATE] ERROR: ' + e.toString());
    Logger.log('[TEMPLATE] Stack: ' + (e.stack || 'No stack'));
    ui.alert(
      'Ошибка',
      'Не удалось создать шаблон:\n' + e.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * Internal function: prepare template by clearing data.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {boolean} keepDemoData - Keep demo accounts and categories
 */
function pfPrepareTemplate_(ss, keepDemoData) {
  Logger.log('[TEMPLATE] Preparing template, keepDemoData: ' + keepDemoData);
  
  // 1. Clear Transactions (keep header)
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (txSheet) {
    var lastRow = txSheet.getLastRow();
    if (lastRow > 1) {
      var rowsToDelete = lastRow - 1;
      Logger.log('[TEMPLATE] Clearing ' + rowsToDelete + ' transaction rows');
      pfClearSheetRows_(txSheet, 2, rowsToDelete);
    }
  }
  
  // 2. Clear or prepare Accounts
  var accountsSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
  if (accountsSheet) {
    var lastRow = accountsSheet.getLastRow();
    if (lastRow > 1) {
      Logger.log('[TEMPLATE] Processing Accounts sheet, lastRow: ' + lastRow);
      
      if (keepDemoData) {
        // Keep only first 2-3 demo accounts, clear the rest
        var demoAccountsToKeep = 3;
        if (lastRow > demoAccountsToKeep + 1) {
          var rowsToDelete = lastRow - demoAccountsToKeep - 1;
          Logger.log('[TEMPLATE] Clearing ' + rowsToDelete + ' account rows (keeping ' + demoAccountsToKeep + ' demo)');
          pfClearSheetRows_(accountsSheet, demoAccountsToKeep + 2, rowsToDelete);
        }
        
        // Reset demo accounts to default values
        pfResetDemoAccounts_(accountsSheet);
      } else {
        // Clear all accounts
        var rowsToDelete = lastRow - 1;
        Logger.log('[TEMPLATE] Clearing all ' + rowsToDelete + ' account rows');
        pfClearSheetRows_(accountsSheet, 2, rowsToDelete);
      }
    } else if (keepDemoData) {
      // No accounts exist, create demo ones
      pfCreateDemoAccounts_(accountsSheet);
    }
  }
  
  // 3. Clear or prepare Categories
  var categoriesSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.CATEGORIES);
  if (categoriesSheet) {
    var lastRow = categoriesSheet.getLastRow();
    if (lastRow > 1) {
      Logger.log('[TEMPLATE] Processing Categories sheet, lastRow: ' + lastRow);
      
      if (keepDemoData) {
        // Keep only first 5-7 demo categories, clear the rest
        var demoCategoriesToKeep = 7;
        if (lastRow > demoCategoriesToKeep + 1) {
          var rowsToDelete = lastRow - demoCategoriesToKeep - 1;
          var lastCol = categoriesSheet.getLastColumn();
          var clearRange = categoriesSheet.getRange(demoCategoriesToKeep + 2, 1, rowsToDelete, lastCol);
          clearRange.clearContent();
          clearRange.clearFormat();
          
          try {
            categoriesSheet.deleteRows(demoCategoriesToKeep + 2, rowsToDelete);
          } catch (e) {
            Logger.log('[TEMPLATE] WARNING: Could not delete category rows: ' + e.toString());
          }
        }
        
        // Reset demo categories to default values
        pfResetDemoCategories_(categoriesSheet);
      } else {
        // Clear all categories
        var rowsToDelete = lastRow - 1;
        Logger.log('[TEMPLATE] Clearing all ' + rowsToDelete + ' category rows');
        pfClearSheetRows_(categoriesSheet, 2, rowsToDelete);
      }
    } else if (keepDemoData) {
      // No categories exist, create demo ones
      pfCreateDemoCategories_(categoriesSheet);
    }
  }
  
  // 4. Clear Import_Raw staging sheet
  var importRawSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.IMPORT_RAW);
  if (importRawSheet) {
    var lastRow = importRawSheet.getLastRow();
    if (lastRow > 1) {
      var rowsToDelete = lastRow - 1;
      Logger.log('[TEMPLATE] Clearing Import_Raw sheet: ' + rowsToDelete + ' rows');
      pfClearSheetRows_(importRawSheet, 2, rowsToDelete);
    }
  }
  
  // 5. Refresh Reports and Dashboard (they will recalculate with empty data)
  Logger.log('[TEMPLATE] Refreshing Reports and Dashboard');
  pfInitializeReports_(ss);
  pfInitializeDashboard_(ss);
  
  // 6. Update Settings (reset owner info, keep language/currency)
  Logger.log('[TEMPLATE] Updating Settings');
  var settingsSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.SETTINGS);
  if (settingsSheet) {
    // Update template creation date
    var now = new Date();
    pfSetSetting_(ss, 'TemplateCreatedAt', Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
    
    // Note: Language and DefaultCurrency are kept as they are useful defaults
  }
  
  // 7. Update named ranges (they should still work, but refresh to be sure)
  Logger.log('[TEMPLATE] Refreshing named ranges');
  pfConfigureReferenceSheets_(ss);
  
  SpreadsheetApp.flush();
  
  Logger.log('[TEMPLATE] Template preparation completed');
}

/**
 * Reset demo accounts to default values.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function pfResetDemoAccounts_(sheet) {
  var demoAccounts = [
    ['Наличные', 'cash', 'RUB', 0, true],
    ['Карта', 'card', 'RUB', 0, true],
    ['Вклад', 'deposit', 'RUB', 0, true]
  ];
  
  var lastRow = sheet.getLastRow();
  var numToUpdate = Math.min(demoAccounts.length, lastRow - 1);
  
  for (var i = 0; i < numToUpdate; i++) {
    var row = sheet.getRange(i + 2, 1, 1, demoAccounts[i].length);
    row.setValues([demoAccounts[i]]);
  }
  
  Logger.log('[TEMPLATE] Reset ' + numToUpdate + ' demo accounts');
}

/**
 * Create demo accounts if sheet is empty.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function pfCreateDemoAccounts_(sheet) {
  var demoAccounts = [
    ['Наличные', 'cash', 'RUB', 0, true],
    ['Карта', 'card', 'RUB', 0, true],
    ['Вклад', 'deposit', 'RUB', 0, true]
  ];
  
  if (sheet.getLastRow() <= 1) {
    var range = sheet.getRange(2, 1, demoAccounts.length, demoAccounts[0].length);
    range.setValues(demoAccounts);
    Logger.log('[TEMPLATE] Created ' + demoAccounts.length + ' demo accounts');
  }
}

/**
 * Reset demo categories to default values.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function pfResetDemoCategories_(sheet) {
  var demoCategories = [
    ['Продукты', 'expense', ''],
    ['Транспорт', 'expense', ''],
    ['Кафе и рестораны', 'expense', ''],
    ['Здоровье', 'expense', ''],
    ['Развлечения', 'expense', ''],
    ['Зарплата', 'income', ''],
    ['Прочее', 'both', '']
  ];
  
  var lastRow = sheet.getLastRow();
  var numToUpdate = Math.min(demoCategories.length, lastRow - 1);
  
  for (var i = 0; i < numToUpdate; i++) {
    var row = sheet.getRange(i + 2, 1, 1, demoCategories[i].length);
    row.setValues([demoCategories[i]]);
  }
  
  Logger.log('[TEMPLATE] Reset ' + numToUpdate + ' demo categories');
}

/**
 * Create demo categories if sheet is empty.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function pfCreateDemoCategories_(sheet) {
  var demoCategories = [
    ['Продукты', 'expense', ''],
    ['Транспорт', 'expense', ''],
    ['Кафе и рестораны', 'expense', ''],
    ['Здоровье', 'expense', ''],
    ['Развлечения', 'expense', ''],
    ['Зарплата', 'income', ''],
    ['Прочее', 'both', '']
  ];
  
  if (sheet.getLastRow() <= 1) {
    var range = sheet.getRange(2, 1, demoCategories.length, demoCategories[0].length);
    range.setValues(demoCategories);
    Logger.log('[TEMPLATE] Created ' + demoCategories.length + ' demo categories');
  }
}
