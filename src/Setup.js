/**
 * Setup / initialization routines.
 *
 * Goal: make `pfSetup()` idempotent and safe to rerun.
 */

var PF_NAMED_RANGES = {
  ACCOUNTS: 'PF_ACCOUNTS',
  CATEGORIES: 'PF_CATEGORIES'
};

var PF_DEFAULT_CURRENCIES = ['RUB', 'USD', 'EUR'];

var PF_SETUP_KEYS = {
  SCHEMA_VERSION: 'SchemaVersion',
  DEFAULT_CURRENCY: 'DefaultCurrency'
};

/**
 * Main setup entry.
 */
function pfRunSetup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Project-wide locale conventions (separate from UI language).
  ss.setSpreadsheetLocale('ru_RU');

  // Ensure Settings exists early (uses current language, default RU).
  pfEnsureSettingsSheet_(ss);

  // Ensure language has a value (project default is RU).
  if (!pfGetSetting_(ss, PF_SETTINGS_KEYS.LANGUAGE)) {
    pfSetSetting_(ss, PF_SETTINGS_KEYS.LANGUAGE, PF_DEFAULT_LANG);
  }

  // Track schema version and default currency.
  if (!pfGetSetting_(ss, PF_SETUP_KEYS.SCHEMA_VERSION)) {
    pfSetSetting_(ss, PF_SETUP_KEYS.SCHEMA_VERSION, String(PF_SCHEMA_VERSION));
  }
  if (!pfGetSetting_(ss, PF_SETUP_KEYS.DEFAULT_CURRENCY)) {
    pfSetSetting_(ss, PF_SETUP_KEYS.DEFAULT_CURRENCY, 'RUB');
  }

  // Create/rename sheets and set headers according to selected language.
  pfApplyLocalization_(ss);

  // Ensure Import_Raw staging sheet exists
  pfEnsureImportRawSheet_(ss);

  // Apply filters, named ranges and validations.
  pfConfigureReferenceSheets_(ss);
  pfConfigureTransactionsSheet_(ss);
  pfEnsureHelpContent_(ss);

  // Initialize Reports sheet with formulas.
  pfInitializeReports_(ss);

  // Initialize Dashboard sheet with KPI and charts.
  pfInitializeDashboard_(ss);

  SpreadsheetApp.flush();
}

function pfConfigureReferenceSheets_(ss) {
  var accountsSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
  pfEnsureFilter_(accountsSheet, PF_ACCOUNTS_SCHEMA.columns.length);

  var categoriesSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.CATEGORIES);
  pfEnsureFilter_(categoriesSheet, PF_CATEGORIES_SCHEMA.columns.length);

  // Named ranges cover full columns below header (easy to append values).
  pfUpsertNamedRange_(ss, PF_NAMED_RANGES.ACCOUNTS, accountsSheet.getRange('A2:A'));
  pfUpsertNamedRange_(ss, PF_NAMED_RANGES.CATEGORIES, categoriesSheet.getRange('A2:A'));
}

function pfConfigureTransactionsSheet_(ss) {
  var sheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  var numCols = PF_TRANSACTIONS_SCHEMA.columns.length;
  pfEnsureFilter_(sheet, numCols);

  // Formats.
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  if (dateCol) sheet.getRange(2, dateCol, sheet.getMaxRows() - 1, 1).setNumberFormat('dd.mm.yyyy');
  if (amountCol) sheet.getRange(2, amountCol, sheet.getMaxRows() - 1, 1).setNumberFormat('0.00');

  // Validations (lightweight, allow blanks where it makes sense).
  var typeCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
  var statusCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
  var currencyCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Currency');
  var accountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Account');
  var accountToCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'AccountTo');
  var categoryCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category');

  if (typeCol) {
    var ruleType = SpreadsheetApp.newDataValidation()
      .requireValueInList(['expense', 'income', 'transfer'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, typeCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleType);
  }

  if (statusCol) {
    var ruleStatus = SpreadsheetApp.newDataValidation()
      .requireValueInList(['ok', 'needs_review', 'duplicate', 'deleted'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, statusCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleStatus);
  }

  if (currencyCol) {
    var ruleCurrency = SpreadsheetApp.newDataValidation()
      .requireValueInList(PF_DEFAULT_CURRENCIES, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, currencyCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleCurrency);
  }

  var accountsRange = ss.getRangeByName(PF_NAMED_RANGES.ACCOUNTS);
  if (accountsRange && accountCol) {
    var ruleAccount = SpreadsheetApp.newDataValidation()
      .requireValueInRange(accountsRange, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, accountCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleAccount);
  }

  if (accountsRange && accountToCol) {
    var ruleAccountTo = SpreadsheetApp.newDataValidation()
      .requireValueInRange(accountsRange, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, accountToCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleAccountTo);
  }

  var categoriesRange = ss.getRangeByName(PF_NAMED_RANGES.CATEGORIES);
  if (categoriesRange && categoryCol) {
    var ruleCategory = SpreadsheetApp.newDataValidation()
      .requireValueInRange(categoriesRange, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, categoryCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleCategory);
  }

  if (amountCol) {
    var ruleAmount = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, amountCol, sheet.getMaxRows() - 1, 1).setDataValidation(ruleAmount);
  }
}

function pfEnsureHelpContent_(ss) {
  var sheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.HELP);
  // Always refresh help content to ensure it's up to date
  // Clear existing content first
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow > 0 && lastCol > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).clearContent();
    sheet.getRange(1, 1, lastRow, lastCol).clearFormat();
  }

  var lang = pfGetLanguage_();
  
  if (lang === 'en') {
    _writeHelpContentEn_(sheet);
  } else {
    _writeHelpContentRu_(sheet);
  }
  
  // Auto-resize columns
  for (var i = 1; i <= sheet.getLastColumn(); i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Write Russian help content.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function _writeHelpContentRu_(sheet) {
  var row = 1;
  
  // Title
  sheet.getRange(row, 1).setValue('Personal finances — Инструкция');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  // Quick start
  sheet.getRange(row, 1).setValue('Быстрый старт');
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;
  
  sheet.getRange(row, 1).setValue('1. Запустите "Personal finances → Setup (создать листы)" для инициализации таблицы.');
  row++;
  sheet.getRange(row, 1).setValue('2. Заполните лист "Счета" — добавьте ваши счета (карты, наличные, вклады).');
  row++;
  sheet.getRange(row, 1).setValue('3. Заполните лист "Категории" — добавьте категории доходов и расходов.');
  row++;
  sheet.getRange(row, 1).setValue('4. Добавляйте транзакции в лист "Транзакции" вручную или импортируйте из выписок.');
  row++;
  sheet.getRange(row, 1).setValue('5. Просматривайте аналитику на листах "Дашборд" и "Отчеты".');
  row += 2;
  
  // Rules
  sheet.getRange(row, 1).setValue('Правила и особенности');
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;
  
  sheet.getRange(row, 1).setValue('Сумма (Amount):');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Всегда указывается как положительное число.');
  row++;
  sheet.getRange(row, 1).setValue('• Для расходов (expense) система автоматически учитывает как отрицательное значение.');
  row++;
  sheet.getRange(row, 1).setValue('• Для доходов (income) — как положительное.');
  row++;
  sheet.getRange(row, 1).setValue('• Для переводов (transfer) сумма списывается с одного счета и добавляется на другой.');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Типы транзакций:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• expense (расход) — трата денег (покупка, оплата услуг).');
  row++;
  sheet.getRange(row, 1).setValue('• income (доход) — получение денег (зарплата, подарок).');
  row++;
  sheet.getRange(row, 1).setValue('• transfer (перевод) — перемещение денег между счетами. Укажите счет-источник (Account) и счет-получатель (AccountTo).');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Дедупликация:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• При импорте система автоматически определяет дубликаты транзакций.');
  row++;
  sheet.getRange(row, 1).setValue('• Дубликаты помечаются статусом "duplicate" и не добавляются в основную таблицу.');
  row++;
  sheet.getRange(row, 1).setValue('• Для поиска дубликата используйте меню "Personal finances → Найти дубликат (по ключу)".');
  row++;
  sheet.getRange(row, 1).setValue('• Дедупликация работает по SourceId (если есть) или по хэшу ключевых полей (дата, счет, сумма, тип).');
  row += 2;
  
  // FAQ
  sheet.getRange(row, 1).setValue('Часто задаваемые вопросы (FAQ)');
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;
  
  sheet.getRange(row, 1).setValue('Как импортировать выписку из банка?');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('1. Экспортируйте выписку в CSV формат (если банк предоставляет PDF, конвертируйте в CSV).');
  row++;
  sheet.getRange(row, 1).setValue('2. В меню выберите "Personal finances → Импорт транзакций".');
  row++;
  sheet.getRange(row, 1).setValue('3. Выберите файл, укажите счет по умолчанию и валюту.');
  row++;
  sheet.getRange(row, 1).setValue('4. Нажмите "Предпросмотр" для проверки данных.');
  row++;
  sheet.getRange(row, 1).setValue('5. Исправьте ошибки парсинга (если есть) и нажмите "Импортировать".');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Как обновить код/скрипты?');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('Если вы используете версию из GitHub:');
  row++;
  sheet.getRange(row, 1).setValue('1. Скачайте последнюю версию из репозитория.');
  row++;
  sheet.getRange(row, 1).setValue('2. Используйте clasp для синхронизации: npm run push');
  row++;
  sheet.getRange(row, 1).setValue('3. Или обновите вручную через редактор Apps Script (Расширения → Apps Script).');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Почему в отчетах/дашборде нет данных?');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Убедитесь, что в листе "Транзакции" есть данные.');
  row++;
  sheet.getRange(row, 1).setValue('• Нажмите "Personal finances → Обновить отчёты" и "Обновить дашборд".');
  row++;
  sheet.getRange(row, 1).setValue('• Проверьте, что транзакции имеют правильный формат (дата, тип, сумма).');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Как создать шаблон для другого пользователя?');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('1. В меню выберите "Personal finances → Создать шаблон".');
  row++;
  sheet.getRange(row, 1).setValue('2. Подтвердите очистку данных.');
  row++;
  sheet.getRange(row, 1).setValue('3. Выберите, оставить ли примеры счетов и категорий.');
  row++;
  sheet.getRange(row, 1).setValue('4. Скопируйте таблицу и передайте новому пользователю.');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Проблемы с правами доступа');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Убедитесь, что у вас есть права на редактирование таблицы.');
  row++;
  sheet.getRange(row, 1).setValue('• Для работы скриптов нужны права на выполнение Apps Script.');
  row++;
  sheet.getRange(row, 1).setValue('• При первом запуске меню может потребоваться авторизация Google.');
  row += 2;
  
  // Troubleshooting
  sheet.getRange(row, 1).setValue('Решение проблем');
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;
  
  sheet.getRange(row, 1).setValue('Ошибки валидации транзакций:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Проверьте обязательные поля: Дата, Тип, Счет, Сумма, Валюта.');
  row++;
  sheet.getRange(row, 1).setValue('• Убедитесь, что Счет и Категория существуют в соответствующих справочниках.');
  row++;
  sheet.getRange(row, 1).setValue('• Для переводов обязательно укажите AccountTo.');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Импорт не работает:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Проверьте формат файла (поддерживаются CSV и выписки Сбербанка).');
  row++;
  sheet.getRange(row, 1).setValue('• Убедитесь, что файл не слишком большой (рекомендуется до 2000 строк за раз).');
  row++;
  sheet.getRange(row, 1).setValue('• Проверьте логи Apps Script (Расширения → Apps Script → Выполнения).');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Формулы показывают ошибки:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Убедитесь, что локаль таблицы установлена в ru_RU (проверяется при Setup).');
  row++;
  sheet.getRange(row, 1).setValue('• Обновите отчеты и дашборд через меню.');
  row++;
  sheet.getRange(row, 1).setValue('• Проверьте, что данные в правильном формате (даты как даты, суммы как числа).');
}

/**
 * Write English help content.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function _writeHelpContentEn_(sheet) {
  var row = 1;
  
  // Title
  sheet.getRange(row, 1).setValue('Personal finances — Help');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  // Quick start
  sheet.getRange(row, 1).setValue('Quick start');
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;
  
  sheet.getRange(row, 1).setValue('1. Run "Personal finances → Setup (create sheets)" to initialize the spreadsheet.');
  row++;
  sheet.getRange(row, 1).setValue('2. Fill the "Accounts" sheet — add your accounts (cards, cash, deposits).');
  row++;
  sheet.getRange(row, 1).setValue('3. Fill the "Categories" sheet — add income and expense categories.');
  row++;
  sheet.getRange(row, 1).setValue('4. Add transactions to the "Transactions" sheet manually or import from statements.');
  row++;
  sheet.getRange(row, 1).setValue('5. View analytics on "Dashboard" and "Reports" sheets.');
  row += 2;
  
  // Rules
  sheet.getRange(row, 1).setValue('Rules and features');
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;
  
  sheet.getRange(row, 1).setValue('Amount:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Always enter as a positive number.');
  row++;
  sheet.getRange(row, 1).setValue('• For expenses, the system automatically treats it as negative.');
  row++;
  sheet.getRange(row, 1).setValue('• For income — as positive.');
  row++;
  sheet.getRange(row, 1).setValue('• For transfers, the amount is deducted from one account and added to another.');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Transaction types:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• expense — spending money (purchases, service payments).');
  row++;
  sheet.getRange(row, 1).setValue('• income — receiving money (salary, gift).');
  row++;
  sheet.getRange(row, 1).setValue('• transfer — moving money between accounts. Specify source account (Account) and destination (AccountTo).');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Deduplication:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• During import, the system automatically detects duplicate transactions.');
  row++;
  sheet.getRange(row, 1).setValue('• Duplicates are marked with "duplicate" status and not added to the main table.');
  row++;
  sheet.getRange(row, 1).setValue('• To find a duplicate, use menu "Personal finances → Find duplicate (by key)".');
  row++;
  sheet.getRange(row, 1).setValue('• Deduplication works by SourceId (if available) or by hash of key fields (date, account, amount, type).');
  row += 2;
  
  // FAQ
  sheet.getRange(row, 1).setValue('Frequently Asked Questions (FAQ)');
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;
  
  sheet.getRange(row, 1).setValue('How to import bank statement?');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('1. Export statement to CSV format (if bank provides PDF, convert to CSV).');
  row++;
  sheet.getRange(row, 1).setValue('2. In menu select "Personal finances → Import transactions".');
  row++;
  sheet.getRange(row, 1).setValue('3. Select file, specify default account and currency.');
  row++;
  sheet.getRange(row, 1).setValue('4. Click "Preview" to review data.');
  row++;
  sheet.getRange(row, 1).setValue('5. Fix parsing errors (if any) and click "Import".');
  row += 2;
  
  sheet.getRange(row, 1).setValue('How to update code/scripts?');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('If you use version from GitHub:');
  row++;
  sheet.getRange(row, 1).setValue('1. Download latest version from repository.');
  row++;
  sheet.getRange(row, 1).setValue('2. Use clasp to sync: npm run push');
  row++;
  sheet.getRange(row, 1).setValue('3. Or update manually via Apps Script editor (Extensions → Apps Script).');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Why reports/dashboard show no data?');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Make sure "Transactions" sheet has data.');
  row++;
  sheet.getRange(row, 1).setValue('• Click "Personal finances → Refresh reports" and "Refresh dashboard".');
  row++;
  sheet.getRange(row, 1).setValue('• Check that transactions have correct format (date, type, amount).');
  row += 2;
  
  sheet.getRange(row, 1).setValue('How to create template for another user?');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('1. In menu select "Personal finances → Create template".');
  row++;
  sheet.getRange(row, 1).setValue('2. Confirm data clearing.');
  row++;
  sheet.getRange(row, 1).setValue('3. Choose whether to keep example accounts and categories.');
  row++;
  sheet.getRange(row, 1).setValue('4. Copy the spreadsheet and share with new user.');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Permission issues');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Make sure you have edit permissions for the spreadsheet.');
  row++;
  sheet.getRange(row, 1).setValue('• Scripts require Apps Script execution permissions.');
  row++;
  sheet.getRange(row, 1).setValue('• First menu run may require Google authorization.');
  row += 2;
  
  // Troubleshooting
  sheet.getRange(row, 1).setValue('Troubleshooting');
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;
  
  sheet.getRange(row, 1).setValue('Transaction validation errors:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Check required fields: Date, Type, Account, Amount, Currency.');
  row++;
  sheet.getRange(row, 1).setValue('• Make sure Account and Category exist in reference sheets.');
  row++;
  sheet.getRange(row, 1).setValue('• For transfers, AccountTo is required.');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Import not working:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Check file format (CSV and Sberbank statements supported).');
  row++;
  sheet.getRange(row, 1).setValue('• Make sure file is not too large (recommended up to 2000 rows at once).');
  row++;
  sheet.getRange(row, 1).setValue('• Check Apps Script logs (Extensions → Apps Script → Executions).');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Formulas show errors:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Make sure spreadsheet locale is set to ru_RU (checked during Setup).');
  row++;
  sheet.getRange(row, 1).setValue('• Refresh reports and dashboard via menu.');
  sheet.getRange(row, 1).setValue('• Check that data is in correct format (dates as dates, amounts as numbers).');
}

/**
 * @param {{columns: Array<{key: string}>}} schema
 * @param {string} key
 * @returns {number|null} 1-based index
 */
function pfColumnIndex_(schema, key) {
  for (var i = 0; i < schema.columns.length; i++) {
    if (schema.columns[i].key === key) return i + 1;
  }
  return null;
}

