/**
 * Точка входа: добавляем меню в Google Sheets.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu(pfT_('menu.root'));

  menu.addItem(pfT_('menu.setup'), 'pfSetup');

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
 * Setup: создаём листы-заготовки и применяем базовые настройки.
 */
function pfSetup() {
  pfRunSetup_();
  SpreadsheetApp.getUi().alert('Готово: таблица инициализирована/обновлена.');
}

