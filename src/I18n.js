/**
 * i18n utilities + dictionaries.
 *
 * Language affects:
 * - sheet tab names
 * - header labels
 * - menu labels (optional)
 *
 * Locale (ru_RU) rules for formulas/number formats are documented separately:
 * see `SHEETS_LOCALE.md`.
 */

var PF_SUPPORTED_LANGS = ['ru', 'en'];
var PF_DEFAULT_LANG = 'ru';

// Stable logical keys for sheets used across codebase.
var PF_SHEET_KEYS = {
  SETTINGS: 'settings',
  TRANSACTIONS: 'transactions',
  CATEGORIES: 'categories',
  ACCOUNTS: 'accounts',
  REPORTS: 'reports',
  DASHBOARD: 'dashboard',
  HELP: 'help',
  IMPORT_RAW: 'import_raw',
  BUDGETS: 'budgets',
  RECURRING_TRANSACTIONS: 'recurring_transactions',
  CATEGORY_RULES: 'category_rules'
};

var PF_I18N = {
  ru: {
    menu: {
      root: 'Personal finances',
      setup: 'Setup (создать листы)',
      validate_all: 'Проверить все транзакции',
      mark_review: 'Пометить на проверку',
      refresh_reports: 'Обновить отчёты',
      refresh_dashboard: 'Обновить дашборд',
      generate_test_data: 'Заполнить тестовыми данными',
      import_transactions: 'Импорт транзакций',
      create_template: 'Создать шаблон',
      find_duplicate: 'Найти дубликат (по ключу)',
      update_budgets: 'Обновить бюджеты',
      create_recurring: 'Создать регулярные платежи',
      apply_auto_categorization: 'Применить автокатегоризацию',
      quick_entry: 'Быстрый ввод транзакции',
      search_transactions: 'Найти транзакции',
      export_menu: 'Экспорт',
      export_transactions_csv: 'Транзакции (CSV)',
      export_transactions_json: 'Транзакции (JSON)',
      export_all_data_json: 'Все данные (JSON)',
      create_backup: 'Создать резервную копию',
      archive_old_transactions: 'Архивировать старые транзакции',
      language: 'Язык',
      lang_ru: 'Русский',
      lang_en: 'English'
    },
    sheet: {
      settings: 'Настройки',
      transactions: 'Транзакции',
      categories: 'Категории',
      accounts: 'Счета',
      reports: 'Отчеты',
      dashboard: 'Дашборд',
      help: 'Инструкция',
      import_raw: 'Импорт (черновик)',
      budgets: 'Бюджеты',
      recurring_transactions: 'Регулярные платежи',
      category_rules: 'Правила категоризации'
    },
    columns: {
      Date: 'Дата',
      Type: 'Тип',
      Account: 'Счет',
      AccountTo: 'Счет (получатель)',
      Amount: 'Сумма',
      Currency: 'Валюта',
      Category: 'Категория',
      Subcategory: 'Подкатегория',
      CategoryType: 'Тип категории',
      AccountType: 'Тип счета',
      InitialBalance: 'Начальный баланс',
      Active: 'Активно',
      Merchant: 'Место/контрагент',
      Description: 'Комментарий',
      Tags: 'Теги',
      Source: 'Источник',
      SourceId: 'ID источника',
      Status: 'Статус',
      Period: 'Период',
      PeriodValue: 'Значение периода',
      Fact: 'Факт',
      Remaining: 'Остаток',
      PercentUsed: '% использования',
      RuleName: 'Название правила',
      Pattern: 'Паттерн',
      PatternType: 'Тип паттерна',
      Priority: 'Приоритет',
      ApplyTo: 'Применять к'
    },
    budget_status: {
      ok: 'В норме',
      warning: 'Предупреждение',
      exceeded: 'Превышен'
    },
    recurring_frequency: {
      weekly: 'Еженедельно',
      monthly: 'Ежемесячно',
      quarterly: 'Ежеквартально',
      yearly: 'Ежегодно'
    },
    pattern_type: {
      contains: 'Содержит',
      startsWith: 'Начинается с',
      endsWith: 'Заканчивается на',
      regex: 'Регулярное выражение',
      exact: 'Точное совпадение'
    },
    rule_apply_to: {
      merchant: 'Только место/контрагент',
      description: 'Только комментарий',
      both: 'Оба поля'
    }
  },
  en: {
    menu: {
      root: 'Personal finances',
      setup: 'Setup (create sheets)',
      validate_all: 'Validate all transactions',
      mark_review: 'Mark for review',
      refresh_reports: 'Refresh reports',
      refresh_dashboard: 'Refresh dashboard',
      generate_test_data: 'Generate test data',
      import_transactions: 'Import transactions',
      create_template: 'Create template',
      find_duplicate: 'Find duplicate (by key)',
      update_budgets: 'Update budgets',
      create_recurring: 'Create recurring transactions',
      apply_auto_categorization: 'Apply auto-categorization',
      quick_entry: 'Quick Transaction Entry',
      search_transactions: 'Search Transactions',
      export_menu: 'Export',
      export_transactions_csv: 'Transactions (CSV)',
      export_transactions_json: 'Transactions (JSON)',
      export_all_data_json: 'All Data (JSON)',
      create_backup: 'Create Backup',
      archive_old_transactions: 'Archive Old Transactions',
      language: 'Language',
      lang_ru: 'Русский',
      lang_en: 'English'
    },
    sheet: {
      settings: 'Settings',
      transactions: 'Transactions',
      categories: 'Categories',
      accounts: 'Accounts',
      reports: 'Reports',
      dashboard: 'Dashboard',
      help: 'Help',
      import_raw: 'Import (Staging)',
      budgets: 'Budgets',
      recurring_transactions: 'Recurring Transactions',
      category_rules: 'Category Rules'
    },
    columns: {
      Date: 'Date',
      Type: 'Type',
      Account: 'Account',
      AccountTo: 'Account To',
      Amount: 'Amount',
      Currency: 'Currency',
      Category: 'Category',
      Subcategory: 'Subcategory',
      CategoryType: 'Category type',
      AccountType: 'Account type',
      InitialBalance: 'Initial balance',
      Active: 'Active',
      Merchant: 'Merchant',
      Description: 'Description',
      Tags: 'Tags',
      Source: 'Source',
      SourceId: 'Source ID',
      Status: 'Status',
      Period: 'Period',
      PeriodValue: 'Period Value',
      Fact: 'Fact',
      Remaining: 'Remaining',
      PercentUsed: '% Used',
      Name: 'Name',
      Frequency: 'Frequency',
      DayOfMonth: 'Day of Month',
      DayOfWeek: 'Day of Week',
      StartDate: 'Start Date',
      EndDate: 'End Date',
      LastCreated: 'Last Created',
      RuleName: 'Rule Name',
      Pattern: 'Pattern',
      PatternType: 'Pattern Type',
      Priority: 'Priority',
      ApplyTo: 'Apply To'
    },
    budget_status: {
      ok: 'OK',
      warning: 'Warning',
      exceeded: 'Exceeded'
    },
    recurring_frequency: {
      weekly: 'Weekly',
      monthly: 'Monthly',
      quarterly: 'Quarterly',
      yearly: 'Yearly'
    },
    pattern_type: {
      contains: 'Contains',
      startsWith: 'Starts with',
      endsWith: 'Ends with',
      regex: 'Regular expression',
      exact: 'Exact match'
    },
    rule_apply_to: {
      merchant: 'Merchant only',
      description: 'Description only',
      both: 'Both fields'
    }
  }
};

/**
 * Get translation string by dotted path.
 * @param {string} path Example: "sheet.transactions" or "columns.Amount"
 * @param {string=} lang Optional language override.
 * @returns {string}
 */
function pfT_(path, lang) {
  var l = lang || pfGetLanguage_();
  var dict = PF_I18N[l] || PF_I18N[PF_DEFAULT_LANG];
  var fallback = PF_I18N.en;

  var value = pfGetByPath_(dict, path);
  if (value != null) return String(value);

  var fallbackValue = pfGetByPath_(fallback, path);
  if (fallbackValue != null) return String(fallbackValue);

  return path;
}

/**
 * @param {Object} obj
 * @param {string} path
 * @returns {*}
 */
function pfGetByPath_(obj, path) {
  var parts = path.split('.');
  var cur = obj;
  for (var i = 0; i < parts.length; i++) {
    if (!cur || typeof cur !== 'object') return null;
    cur = cur[parts[i]];
  }
  return cur;
}

/**
 * Returns all known names (in all supported langs) for a logical sheet key.
 * Used to find/rename sheets when switching language.
 * @param {string} sheetKey One of PF_SHEET_KEYS values.
 * @returns {string[]}
 */
function pfAllSheetNames_(sheetKey) {
  var names = [];
  for (var i = 0; i < PF_SUPPORTED_LANGS.length; i++) {
    var lang = PF_SUPPORTED_LANGS[i];
    var name = pfGetByPath_(PF_I18N[lang], 'sheet.' + sheetKey);
    if (name && names.indexOf(name) === -1) names.push(name);
  }
  return names;
}

