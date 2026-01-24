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

  menu
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

