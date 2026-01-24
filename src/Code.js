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
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Project-wide locale conventions (separate from UI language).
  ss.setSpreadsheetLocale('ru_RU');

  // Ensure default language exists (project default is RU).
  if (!pfGetSetting_(ss, PF_SETTINGS_KEYS.LANGUAGE)) {
    pfSetSetting_(ss, PF_SETTINGS_KEYS.LANGUAGE, PF_DEFAULT_LANG);
  }

  // Create/rename sheets and set headers according to selected language.
  pfApplyLocalization_(ss);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert('Готово: листы созданы/обновлены, язык применён.');
}

