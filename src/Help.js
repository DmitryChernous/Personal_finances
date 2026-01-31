/**
 * Help sheet content management.
 * 
 * Provides functions to write and update help content in the Help sheet.
 * Separated from Setup.js to improve modularity.
 */

/**
 * Ensure Help sheet exists and has up-to-date content.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
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
  sheet.getRange(row, 1).setValue('• expense — трата денег (покупки, оплата услуг).');
  row++;
  sheet.getRange(row, 1).setValue('• income — получение денег (зарплата, подарок).');
  row++;
  sheet.getRange(row, 1).setValue('• transfer — перемещение денег между счетами. Укажите исходный счет (Счет) и получателя (Счет получателя).');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Дедупликация:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• При импорте система автоматически определяет дублирующиеся транзакции.');
  row++;
  sheet.getRange(row, 1).setValue('• Дубликаты помечаются статусом "duplicate" и не добавляются в основную таблицу.');
  row++;
  sheet.getRange(row, 1).setValue('• Для поиска дубликата используйте меню "Personal finances → Найти дубликат (по ключу)".');
  row++;
  sheet.getRange(row, 1).setValue('• Дедупликация работает по SourceId (если есть) или по хэшу ключевых полей (дата, счет, сумма, тип).');
  row += 2;
  
  // Как добавить новый счёт (raw-лист)
  sheet.getRange(row, 1).setValue('Как добавить новый счёт (raw-лист)');
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('Имя листа должно начинаться с raw (например, raw_Сбербанк, raw_Яндекс Карта). Один лист = один счёт.');
  row++;
  sheet.getRange(row, 1).setValue('В первой строке — заголовки по порядку: ДАТА, ВРЕМЯ, КАТЕГОРИЯ, ОПИСАНИЕ, СУММА, ОСТАТОК СРЕДСТВ, СЧЕТ. См. docs/RAW_SHEETS_ARCHITECTURE.md, п. 3.2.');
  row++;
  sheet.getRange(row, 1).setValue('Формат: дата dd.mm.yyyy, сумма — число (минус = расход, плюс = доход). Если колонка СЧЕТ пустая, подставится имя листа.');
  row++;
  sheet.getRange(row, 1).setValue('Пример строк данных:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('31.12.2025  16:40  Перевод СБП  Перевод для Ч.  -1500  10000  ');
  row++;
  sheet.getRange(row, 1).setValue('31.12.2025  12:44  Здоровье  YUG-FARM Shakhty  -740  9250  ');
  row++;
  sheet.getRange(row, 1).setValue('Порядок действий: создайте лист с именем raw_… → вставьте заголовки → заполните данные → меню "Personal finances → Синхронизировать с raw-листами".');
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
  
  sheet.getRange(row, 1).setValue('Почему отчеты/дашборд не показывают данные?');
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
  sheet.getRange(row, 1).setValue('4. Скопируйте таблицу и поделитесь с новым пользователем.');
  row += 2;
  
  sheet.getRange(row, 1).setValue('Проблемы с правами доступа');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('• Убедитесь, что у вас есть права на редактирование таблицы.');
  row++;
  sheet.getRange(row, 1).setValue('• Скрипты требуют прав на выполнение Apps Script.');
  row++;
  sheet.getRange(row, 1).setValue('• Первый запуск меню может потребовать авторизации Google.');
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
  sheet.getRange(row, 1).setValue('• Убедитесь, что Счет и Категория существуют в справочниках.');
  row++;
  sheet.getRange(row, 1).setValue('• Для переводов обязателен Счет получателя.');
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
  
  // How to add new account (raw sheet)
  sheet.getRange(row, 1).setValue('How to add a new account (raw sheet)');
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('Sheet name must start with raw (e.g. raw_Sberbank, raw_Yandex Card). One sheet = one account.');
  row++;
  sheet.getRange(row, 1).setValue('First row: headers in order — DATE, TIME, CATEGORY, DESCRIPTION, AMOUNT, BALANCE, ACCOUNT. See docs/RAW_SHEETS_ARCHITECTURE.md, § 3.2.');
  row++;
  sheet.getRange(row, 1).setValue('Format: date dd.mm.yyyy, amount as number (negative = expense, positive = income). If ACCOUNT column is empty, sheet name is used.');
  row++;
  sheet.getRange(row, 1).setValue('Example data rows:');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('31.12.2025  16:40  SBP Transfer  Transfer for Ch.  -1500  10000  ');
  row++;
  sheet.getRange(row, 1).setValue('31.12.2025  12:44  Health  YUG-FARM Shakhty  -740  9250  ');
  row++;
  sheet.getRange(row, 1).setValue('Steps: create a sheet named raw_… → add headers → fill data → menu "Personal finances → Sync from raw sheets".');
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
  row++;
  sheet.getRange(row, 1).setValue('• Check that data is in correct format (dates as dates, amounts as numbers).');
}
