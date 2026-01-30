/**
 * Точка входа: добавляем меню в Google Sheets.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu(pfT_('menu.root'));

  menu.addItem(pfT_('menu.setup'), 'pfSetup');
  menu.addSeparator();
  menu.addItem(pfT_('menu.validate_all'), 'pfValidateAllTransactions');
  menu.addItem(pfT_('menu.mark_review'), 'pfMarkSelectedForReview');
  menu.addSeparator();
  menu.addItem(pfT_('menu.refresh_reports'), 'pfRefreshReports');
  menu.addItem(pfT_('menu.refresh_dashboard'), 'pfRefreshDashboard');
  menu.addItem(pfT_('menu.update_budgets'), 'pfUpdateBudgetCalculations');
  menu.addItem(pfT_('menu.create_recurring'), 'pfCreateRecurringTransactions');
  menu.addItem(pfT_('menu.apply_auto_categorization'), 'pfApplyAutoCategorizationToAll');
  menu.addSeparator();
  menu.addItem(pfT_('menu.quick_entry'), 'pfShowQuickEntry');
  menu.addItem(pfT_('menu.search_transactions'), 'pfShowSearchDialog');
  menu.addItem(pfT_('menu.import_transactions'), 'pfImportTransactions');
  menu.addItem(pfT_('menu.sync_raw_sheets'), 'pfSyncRawSheetsToTransactionsMenu');
  menu.addItem(pfT_('menu.find_duplicate'), 'pfFindDuplicateByKey');
  menu.addSeparator();
  menu.addItem(pfT_('menu.archive_old_transactions'), 'pfArchiveOldTransactions');
  menu.addSeparator();
  menu.addItem(pfT_('menu.generate_test_data'), 'pfGenerateTestData');
  menu.addItem(pfT_('menu.create_template'), 'pfCreateTemplate');
  menu.addSeparator();
  menu.addItem('Запустить тесты', 'pfRunAllTests');

  menu
    .addSubMenu(
      ui
        .createMenu(pfT_('menu.export_menu'))
        .addItem(pfT_('menu.export_transactions_csv'), 'pfExportTransactionsCSV')
        .addItem(pfT_('menu.export_transactions_json'), 'pfExportTransactionsJSON')
        .addItem(pfT_('menu.export_all_data_json'), 'pfExportAllDataJSON')
        .addSeparator()
        .addItem(pfT_('menu.create_backup'), 'pfCreateBackup')
    )
    .addSubMenu(
      ui
        .createMenu(pfT_('menu.language'))
        .addItem(pfT_('menu.lang_ru'), 'pfSetLanguageRu')
        .addItem(pfT_('menu.lang_en'), 'pfSetLanguageEn')
    )
    .addToUi();
}

/**
 * Triggered on edit. Validates and normalizes transaction rows.
 */
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var ss = e.source;
  
  // Only process Transactions sheet.
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet || sheet.getName() !== txSheet.getName()) {
    return;
  }

  var row = e.range.getRow();
  // Skip header row.
  if (row <= 1) return;

  // Normalize the row (trim, defaults).
  pfNormalizeTransactionRow_(sheet, row);

  // Validate and highlight errors (async-friendly, but may be slow on large sheets).
  // For better UX, we validate only the edited row.
  var errors = pfValidateTransactionRow_(sheet, row);
  pfHighlightErrors_(sheet, row, errors);
}

/**
 * Setup: создаём листы-заготовки и применяем базовые настройки.
 */
function pfSetup() {
  pfRunSetup_();
  SpreadsheetApp.getUi().alert('Готово: таблица инициализирована/обновлена.');
}

/**
 * Синхронизация raw-листов с листом «Транзакции». Вызывается из меню.
 */
function pfSyncRawSheetsToTransactionsMenu() {
  var result = pfSyncRawSheetsToTransactions();
  var lang = pfGetLanguage_();
  var msg = '';
  if (lang === 'en') {
    msg = 'Processed sheets: ' + result.sheetsProcessed + '\nAdded: ' + result.added + (result.skipped > 0 ? '\nAdded as "needs_review" (possible duplicates): ' + result.skipped : '');
  } else {
    msg = 'Обработано листов: ' + result.sheetsProcessed + '\nДобавлено: ' + result.added + (result.skipped > 0 ? '\nДобавлено со статусом «На проверку» (возможные дубликаты): ' + result.skipped : '');
  }
  if (result.errors && result.errors.length > 0) {
    msg += '\n\nОшибки:\n' + result.errors.join('\n');
  }
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Find duplicate transaction by deduplication key.
 * Shows dialog to enter key, then searches and highlights the transaction.
 */
function pfFindDuplicateByKey() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Поиск дубликата',
    'Введите ключ дедупликации (из столбца "Ключ дедупликации" в листе предпросмотра):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    var dedupeKey = response.getResponseText();
    
    // Sanitize input: trim and validate
    if (!dedupeKey || typeof dedupeKey !== 'string') {
      ui.alert('Ошибка', 'Ключ дедупликации не может быть пустым', ui.ButtonSet.OK);
      return;
    }
    
    dedupeKey = dedupeKey.trim();
    
    // Additional validation: check for potentially dangerous characters
    // (though in Apps Script context risk is minimal)
    if (dedupeKey.length > 200) {
      ui.alert('Ошибка', 'Ключ дедупликации слишком длинный (максимум 200 символов)', ui.ButtonSet.OK);
      return;
    }
    
    var result = pfFindDuplicateTransaction(dedupeKey);
    
    if (result.found) {
      var tx = result.transaction;
      var message = 'Найдена дублирующая транзакция:\n\n' +
                   'Строка: ' + result.rowNum + '\n' +
                   'Дата: ' + (tx.date ? Utilities.formatDate(tx.date, Session.getScriptTimeZone(), 'dd.MM.yyyy') : '') + '\n' +
                   'Тип: ' + tx.type + '\n' +
                   'Счет: ' + tx.account + '\n' +
                   'Сумма: ' + tx.amount + ' ' + tx.currency + '\n' +
                   'Описание: ' + (tx.description || '') + '\n' +
                   'Источник: ' + tx.source + '\n' +
                   (tx.sourceId ? 'ID источника: ' + tx.sourceId + '\n' : '') +
                   '\nПерейти к этой транзакции?';
      
      var goToResponse = ui.alert('Дубликат найден', message, ui.ButtonSet.YES_NO);
      
      if (goToResponse === ui.Button.YES) {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
        if (txSheet) {
          txSheet.setActiveRange(txSheet.getRange(result.rowNum, 1));
          ss.setActiveSheet(txSheet);
        }
      }
    } else {
      ui.alert('Не найдено', result.message || 'Дублирующая транзакция не найдена', ui.ButtonSet.OK);
    }
  }
}

