/**
 * Test data generation for development and testing.
 *
 * Generates realistic test data for Accounts, Categories, and Transactions.
 * Useful for testing reports, dashboard, and validation features.
 */

/**
 * Generates test data for all sheets.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {boolean} clearExisting If true, clears existing data before generating.
 */
function pfGenerateTestData(clearExisting) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var lang = pfGetLanguage_();

  // Confirm if clearing existing data.
  if (clearExisting === undefined) {
    var response = ui.alert(
      lang === 'en' ? 'Generate Test Data' : 'Генерация тестовых данных',
      lang === 'en' 
        ? 'This will add test data to Accounts, Categories, and Transactions sheets.\n\nClear existing data first?'
        : 'Это добавит тестовые данные в листы Счета, Категории и Транзакции.\n\nОчистить существующие данные?',
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (response === ui.Button.CANCEL) {
      return;
    }
    clearExisting = (response === ui.Button.YES);
  }

  try {
    // Generate test data.
    pfGenerateTestAccounts_(ss, clearExisting);
    pfGenerateTestCategories_(ss, clearExisting);
    pfGenerateTestTransactions_(ss, clearExisting);

    SpreadsheetApp.flush();

    if (lang === 'en') {
      ui.alert('Test data generated successfully!');
    } else {
      ui.alert('Тестовые данные успешно сгенерированы!');
    }
  } catch (e) {
    if (lang === 'en') {
      ui.alert('Error generating test data: ' + e.toString());
    } else {
      ui.alert('Ошибка при генерации тестовых данных: ' + e.toString());
    }
  }
}

/**
 * Generates test accounts.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {boolean} clearExisting
 */
function pfGenerateTestAccounts_(ss, clearExisting) {
  var accountsSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
  
  // Ensure headers.
  pfEnsureHeaderRow_(accountsSheet, PF_ACCOUNTS_SCHEMA);

  // Cache lastRow to avoid multiple calls
  var lastRow = accountsSheet.getLastRow();
  if (clearExisting && lastRow > 1) {
    accountsSheet.deleteRows(2, lastRow - 1);
    lastRow = 1; // Reset after clearing
  }

  var accounts = [
    ['Наличные', 'cash', 'RUB', '5000', 'true', 'Наличные деньги'],
    ['Карта Сбер', 'card', 'RUB', '10000', 'true', 'Основная дебетовая карта'],
    ['Карта Тинькофф', 'card', 'RUB', '0', 'true', 'Кредитная карта'],
    ['Вклад', 'deposit', 'RUB', '50000', 'true', 'Сберегательный вклад'],
    ['Доллары', 'cash', 'USD', '500', 'true', 'Наличные доллары']
  ];

  var startRow = lastRow + 1;
  if (startRow === 2 && !clearExisting) {
    startRow = 2; // First data row.
  }

  accountsSheet.getRange(startRow, 1, accounts.length, accounts[0].length).setValues(accounts);
}

/**
 * Generates test categories.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {boolean} clearExisting
 */
function pfGenerateTestCategories_(ss, clearExisting) {
  var categoriesSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.CATEGORIES);
  
  // Ensure headers.
  pfEnsureHeaderRow_(categoriesSheet, PF_CATEGORIES_SCHEMA);

  // Cache lastRow to avoid multiple calls
  var lastRow = categoriesSheet.getLastRow();
  if (clearExisting && lastRow > 1) {
    categoriesSheet.deleteRows(2, lastRow - 1);
    lastRow = 1; // Reset after clearing
  }

  var categories = [
    // Income categories.
    ['Зарплата', '', 'income', 'true', 'Основной доход'],
    ['Премия', '', 'income', 'true', 'Премиальные выплаты'],
    ['Прочее', '', 'income', 'true', 'Прочие доходы'],
    
    // Expense categories.
    ['Еда', 'Продукты', 'expense', 'true', 'Покупка продуктов'],
    ['Еда', 'Рестораны', 'expense', 'true', 'Питание вне дома'],
    ['Транспорт', 'Общественный', 'expense', 'true', 'Метро, автобус, такси'],
    ['Транспорт', 'Бензин', 'expense', 'true', 'Заправка автомобиля'],
    ['Жильё', 'Аренда', 'expense', 'true', 'Аренда квартиры'],
    ['Жильё', 'Коммунальные', 'expense', 'true', 'ЖКХ, интернет, телефон'],
    ['Развлечения', 'Кино', 'expense', 'true', 'Походы в кино'],
    ['Развлечения', 'Рестораны', 'expense', 'true', 'Развлекательные заведения'],
    ['Здоровье', 'Врачи', 'expense', 'true', 'Медицинские услуги'],
    ['Здоровье', 'Аптека', 'expense', 'true', 'Лекарства'],
    ['Одежда', '', 'expense', 'true', 'Покупка одежды'],
    ['Образование', '', 'expense', 'true', 'Курсы, обучение']
  ];

  var startRow = lastRow + 1;
  if (startRow === 2 && !clearExisting) {
    startRow = 2; // First data row.
  }

  categoriesSheet.getRange(startRow, 1, categories.length, categories[0].length).setValues(categories);
}

/**
 * Generates test transactions for the last 3 months.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {boolean} clearExisting
 */
function pfGenerateTestTransactions_(ss, clearExisting) {
  var txSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  
  // Ensure headers.
  pfEnsureHeaderRow_(txSheet, PF_TRANSACTIONS_SCHEMA);

  // Cache lastRow to avoid multiple calls
  var lastRow = txSheet.getLastRow();
  if (clearExisting && lastRow > 1) {
    txSheet.deleteRows(2, lastRow - 1);
  }

  var today = new Date();
  var transactions = [];

  // Helper to add days to date.
  function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }

  // Helper to format date for Sheets (dd.mm.yyyy).
  function formatDate(date) {
    var day = date.getDate();
    var month = date.getMonth() + 1;
    var year = date.getFullYear();
    return day + '.' + month + '.' + year;
  }

  // Generate transactions for last 3 months.
  var startDate = new Date(today.getFullYear(), today.getMonth() - 2, 1);
  
  // Income transactions (salary on 5th of each month).
  for (var monthOffset = 0; monthOffset < 3; monthOffset++) {
    var salaryDate = new Date(today.getFullYear(), today.getMonth() - 2 + monthOffset, 5);
    if (salaryDate <= today) {
      transactions.push([
        formatDate(salaryDate),
        'income',
        'Карта Сбер',
        '',
        '80000',
        'RUB',
        'Зарплата',
        '',
        '',
        'Зарплата за месяц',
        '',
        'manual',
        '',
        'ok'
      ]);
    }
  }

  // Expense transactions - random daily expenses.
  var expenseCategories = [
    ['Еда', 'Продукты'],
    ['Еда', 'Рестораны'],
    ['Транспорт', 'Общественный'],
    ['Транспорт', 'Бензин'],
    ['Жильё', 'Коммунальные'],
    ['Развлечения', 'Кино'],
    ['Здоровье', 'Аптека'],
    ['Одежда', '']
  ];

  var expenseAmounts = [
    [500, 2000],   // Еда - Продукты
    [800, 3000],   // Еда - Рестораны
    [50, 500],     // Транспорт - Общественный
    [2000, 5000],  // Транспорт - Бензин
    [3000, 8000],  // Жильё - Коммунальные
    [500, 2000],   // Развлечения - Кино
    [300, 2000],   // Здоровье - Аптека
    [1000, 5000]   // Одежда
  ];

  var accounts = ['Наличные', 'Карта Сбер', 'Карта Тинькофф'];

  // Generate ~60 expense transactions (about 20 per month).
  for (var i = 0; i < 60; i++) {
    var daysAgo = Math.floor(Math.random() * 90); // Last 90 days.
    var txDate = addDays(startDate, daysAgo);
    if (txDate > today) continue;

    var categoryIdx = Math.floor(Math.random() * expenseCategories.length);
    var category = expenseCategories[categoryIdx];
    var amountRange = expenseAmounts[categoryIdx];
    var amount = Math.floor(Math.random() * (amountRange[1] - amountRange[0]) + amountRange[0]);
    var account = accounts[Math.floor(Math.random() * accounts.length)];

    // 5% chance of needs_review status.
    var status = Math.random() < 0.05 ? 'needs_review' : 'ok';

    transactions.push([
      formatDate(txDate),
      'expense',
      account,
      '',
      String(amount),
      'RUB',
      category[0],
      category[1] || '',
      '',
      '',
      '',
      'manual',
      '',
      status
    ]);
  }

  // Sort transactions by date (oldest first).
  transactions.sort(function(a, b) {
    var dateA = new Date(a[0].split('.').reverse().join('-'));
    var dateB = new Date(b[0].split('.').reverse().join('-'));
    return dateA - dateB;
  });

  // Add a few transfers.
  transactions.push([
    formatDate(addDays(today, -10)),
    'transfer',
    'Карта Сбер',
    'Вклад',
    '10000',
    'RUB',
    '',
    '',
    '',
    'Перевод на вклад',
    '',
    'manual',
    '',
    'ok'
  ]);

  var startRow = txSheet.getLastRow() + 1;
  if (startRow === 2 && !clearExisting) {
    startRow = 2; // First data row.
  }

  if (transactions.length > 0) {
    txSheet.getRange(startRow, 1, transactions.length, transactions[0].length).setValues(transactions);
  }

  // Apply date format to date column.
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  if (dateCol && transactions.length > 0) {
    txSheet.getRange(startRow, dateCol, transactions.length, 1).setNumberFormat('dd.mm.yyyy');
  }

  // Apply number format to amount column.
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  if (amountCol && transactions.length > 0) {
    txSheet.getRange(startRow, amountCol, transactions.length, 1).setNumberFormat('0.00');
  }
}
