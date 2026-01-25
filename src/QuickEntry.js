/**
 * Quick Entry module.
 * Handles quick transaction entry via HTML sidebar:
 * - Show sidebar
 * - Load accounts/categories
 * - Add transaction from form
 * - Auto-categorization integration
 */

/**
 * Show quick entry sidebar.
 */
function pfShowQuickEntry() {
  var lang = pfGetLanguage_();
  var title = lang === 'en' ? 'Quick Transaction Entry' : 'Быстрый ввод транзакции';
  
  var html = HtmlService.createHtmlOutputFromFile('QuickEntry')
    .setTitle(title)
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Get list of accounts for quick entry dropdown.
 * @returns {Array<string>} Array of account names
 */
function pfGetAccountsForQuickEntry() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var accountsSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
    
    if (!accountsSheet) {
      return [];
    }
    
    var lastRow = accountsSheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    
    var accountColIdx = pfColumnIndex_(PF_ACCOUNTS_SCHEMA, 'Account');
    var activeColIdx = pfColumnIndex_(PF_ACCOUNTS_SCHEMA, 'Active');
    
    if (!accountColIdx) {
      return [];
    }
    
    var accounts = [];
    var data = accountsSheet.getRange(2, 1, lastRow - 1, PF_ACCOUNTS_SCHEMA.columns.length).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Check if active (default: true if empty)
      var active = true;
      if (activeColIdx) {
        var activeValue = row[activeColIdx - 1];
        if (activeValue === false || activeValue === 'false' || activeValue === 'FALSE' ||
            (typeof activeValue === 'string' && activeValue.trim().toLowerCase() === 'false')) {
          active = false;
        }
      }
      
      if (!active) {
        continue;
      }
      
      var account = String(row[accountColIdx - 1] || '').trim();
      if (account && accounts.indexOf(account) === -1) {
        accounts.push(account);
      }
    }
    
    return accounts;
  } catch (e) {
    pfLogError_(e, 'pfGetAccountsForQuickEntry', PF_LOG_LEVEL.ERROR);
    return [];
  }
}

/**
 * Get list of categories for quick entry dropdown.
 * @returns {Array<string>} Array of category names
 */
function pfGetCategoriesForQuickEntry() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var categoriesSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.CATEGORIES);
    
    if (!categoriesSheet) {
      return [];
    }
    
    var lastRow = categoriesSheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    
    var categoryColIdx = pfColumnIndex_(PF_CATEGORIES_SCHEMA, 'Category');
    var activeColIdx = pfColumnIndex_(PF_CATEGORIES_SCHEMA, 'Active');
    
    if (!categoryColIdx) {
      return [];
    }
    
    var categories = [];
    var data = categoriesSheet.getRange(2, 1, lastRow - 1, PF_CATEGORIES_SCHEMA.columns.length).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Check if active (default: true if empty)
      var active = true;
      if (activeColIdx) {
        var activeValue = row[activeColIdx - 1];
        if (activeValue === false || activeValue === 'false' || activeValue === 'FALSE' ||
            (typeof activeValue === 'string' && activeValue.trim().toLowerCase() === 'false')) {
          active = false;
        }
      }
      
      if (!active) {
        continue;
      }
      
      var category = String(row[categoryColIdx - 1] || '').trim();
      if (category && categories.indexOf(category) === -1) {
        categories.push(category);
      }
    }
    
    return categories;
  } catch (e) {
    pfLogError_(e, 'pfGetCategoriesForQuickEntry', PF_LOG_LEVEL.ERROR);
    return [];
  }
}

/**
 * Get list of subcategories for a category.
 * @param {string} category - Category name
 * @returns {Array<string>} Array of subcategory names
 */
function pfGetSubcategoriesForQuickEntry(category) {
  try {
    if (!category || String(category).trim() === '') {
      return [];
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var categoriesSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.CATEGORIES);
    
    if (!categoriesSheet) {
      return [];
    }
    
    var lastRow = categoriesSheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    
    var categoryColIdx = pfColumnIndex_(PF_CATEGORIES_SCHEMA, 'Category');
    var subcategoryColIdx = pfColumnIndex_(PF_CATEGORIES_SCHEMA, 'Subcategory');
    var activeColIdx = pfColumnIndex_(PF_CATEGORIES_SCHEMA, 'Active');
    
    if (!categoryColIdx || !subcategoryColIdx) {
      return [];
    }
    
    var subcategories = [];
    var data = categoriesSheet.getRange(2, 1, lastRow - 1, PF_CATEGORIES_SCHEMA.columns.length).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Check if active (default: true if empty)
      var active = true;
      if (activeColIdx) {
        var activeValue = row[activeColIdx - 1];
        if (activeValue === false || activeValue === 'false' || activeValue === 'FALSE' ||
            (typeof activeValue === 'string' && activeValue.trim().toLowerCase() === 'false')) {
          active = false;
        }
      }
      
      if (!active) {
        continue;
      }
      
      var rowCategory = String(row[categoryColIdx - 1] || '').trim();
      var subcategory = String(row[subcategoryColIdx - 1] || '').trim();
      
      if (rowCategory === String(category).trim() && subcategory && subcategories.indexOf(subcategory) === -1) {
        subcategories.push(subcategory);
      }
    }
    
    return subcategories;
  } catch (e) {
    pfLogError_(e, 'pfGetSubcategoriesForQuickEntry', PF_LOG_LEVEL.ERROR);
    return [];
  }
}

/**
 * Auto-categorize transaction for quick entry.
 * @param {string} merchant - Merchant name
 * @param {string} description - Transaction description
 * @returns {Object} {category: string, subcategory: string} or null
 */
function pfAutoCategorizeForQuickEntry(merchant, description) {
  try {
    var transaction = {
      merchant: merchant || '',
      description: description || '',
      category: '',
      subcategory: ''
    };
    
    var result = pfApplyCategoryRules_(transaction);
    
    if (result && result.category) {
      return {
        category: result.category,
        subcategory: result.subcategory || ''
      };
    }
    
    return null;
  } catch (e) {
    pfLogError_(e, 'pfAutoCategorizeForQuickEntry', PF_LOG_LEVEL.ERROR);
    return null;
  }
}

/**
 * Add transaction from quick entry form.
 * @param {Object} data - Form data
 * @returns {Object} Result: {success: boolean, message: string, rowNum: number}
 */
function pfAddQuickTransaction(data) {
  try {
    // Validate input
    if (!data || typeof data !== 'object') {
      return pfCreateErrorResponse_('Неверные данные формы');
    }
    
    // Validate required fields
    if (!data.date || !data.type || !data.account || !data.amount || !data.currency) {
      return pfCreateErrorResponse_('Заполните все обязательные поля');
    }
    
    // Validate amount
    var amount = Number(data.amount);
    if (isNaN(amount) || amount <= 0) {
      return pfCreateErrorResponse_('Сумма должна быть положительным числом');
    }
    
    // Validate type
    if (data.type !== PF_TRANSACTION_TYPE.EXPENSE && 
        data.type !== PF_TRANSACTION_TYPE.INCOME && 
        data.type !== PF_TRANSACTION_TYPE.TRANSFER) {
      return pfCreateErrorResponse_('Неверный тип транзакции');
    }
    
    // Validate transfer
    if (data.type === PF_TRANSACTION_TYPE.TRANSFER) {
      if (!data.accountTo || data.accountTo.trim() === '') {
        return pfCreateErrorResponse_('Для перевода обязателен счет получателя');
      }
      if (data.accountTo === data.account) {
        return pfCreateErrorResponse_('Счет получателя должен отличаться от счета отправителя');
      }
    }
    
    // Parse date
    var date = pfISOStringToDate_(data.date);
    if (!date) {
      // Try parsing as date string
      date = new Date(data.date);
      if (isNaN(date.getTime())) {
        return pfCreateErrorResponse_('Неверный формат даты');
      }
    }
    
    // Create transaction object
    var transaction = {
      date: date,
      type: data.type,
      account: String(data.account).trim(),
      accountTo: data.type === PF_TRANSACTION_TYPE.TRANSFER ? String(data.accountTo).trim() : '',
      amount: amount,
      currency: String(data.currency).trim(),
      category: data.category ? String(data.category).trim() : '',
      subcategory: data.subcategory ? String(data.subcategory).trim() : '',
      merchant: data.merchant ? String(data.merchant).trim() : '',
      description: data.description ? String(data.description).trim() : '',
      tags: '',
      source: PF_IMPORT_SOURCE.MANUAL,
      sourceId: '',
      status: PF_TRANSACTION_STATUS.OK
    };
    
    // Apply auto-categorization if category is not set
    if (!transaction.category || transaction.category === '') {
      transaction = pfApplyCategoryRules_(transaction);
    }
    
    // Get Transactions sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    if (!txSheet) {
      return pfCreateErrorResponse_('Лист транзакций не найден. Запустите Setup.');
    }
    
    // Convert transaction to row
    var row = pfTransactionDTOToRow_(transaction);
    
    // Get last row and append
    var lastRow = txSheet.getLastRow();
    var targetRow = lastRow + 1;
    
    // Write row
    txSheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    
    // Apply normalization and validation
    try {
      pfNormalizeTransactionRow_(txSheet, targetRow);
      var errors = pfValidateTransactionRow_(txSheet, targetRow);
      pfHighlightErrors_(txSheet, targetRow, errors);
      
      if (errors.length > 0) {
        // Transaction was added but has validation errors
        return pfCreateSuccessResponse_('Транзакция добавлена, но есть ошибки валидации', {
          rowNum: targetRow,
          errors: errors
        });
      }
    } catch (e) {
      pfLogWarning_('Error normalizing/validating transaction row: ' + e.toString(), 'pfAddQuickTransaction');
      // Transaction was added, but normalization/validation failed
    }
    
    return pfCreateSuccessResponse_('Транзакция успешно добавлена', {
      rowNum: targetRow
    });
    
  } catch (e) {
    pfLogError_(e, 'pfAddQuickTransaction', PF_LOG_LEVEL.ERROR);
    return pfCreateErrorResponse_('Ошибка при добавлении транзакции: ' + (e.message || e.toString()));
  }
}
