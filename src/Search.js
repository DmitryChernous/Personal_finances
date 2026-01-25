/**
 * Search and Filter module.
 * Handles search functionality for transactions.
 */

/**
 * Search transactions by text query.
 * @param {string} query - Search query
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet
 * @returns {Array<Object>} Array of found transactions with row numbers
 */
function pfSearchTransactions_(query, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return [];
  }
  
  // Normalize query
  if (!query || typeof query !== 'string') {
    return [];
  }
  
  var normalizedQuery = query.trim().toLowerCase();
  if (normalizedQuery.length === 0) {
    return [];
  }
  
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet) {
    return [];
  }
  
  var lastRow = txSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  
  // Read all data
  var data = txSheet.getRange(1, 1, lastRow, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
  var headers = data[0];
  
  // Find column indices
  var merchantCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Merchant');
  var descriptionCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Description');
  var categoryCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category');
  var subcategoryCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Subcategory');
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  var typeCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
  var accountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Account');
  
  var results = [];
  
  // Search through data rows (skip header)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowNum = i + 1; // 1-based row number
    var matchField = null;
    
    // Check Merchant field
    if (merchantCol) {
      var merchant = String(row[merchantCol - 1] || '').toLowerCase();
      if (merchant.indexOf(normalizedQuery) !== -1) {
        matchField = 'merchant';
      }
    }
    
    // Check Description field
    if (!matchField && descriptionCol) {
      var description = String(row[descriptionCol - 1] || '').toLowerCase();
      if (description.indexOf(normalizedQuery) !== -1) {
        matchField = 'description';
      }
    }
    
    // Check Category field
    if (!matchField && categoryCol) {
      var category = String(row[categoryCol - 1] || '').toLowerCase();
      if (category.indexOf(normalizedQuery) !== -1) {
        matchField = 'category';
      }
    }
    
    // Check Subcategory field
    if (!matchField && subcategoryCol) {
      var subcategory = String(row[subcategoryCol - 1] || '').toLowerCase();
      if (subcategory.indexOf(normalizedQuery) !== -1) {
        matchField = 'subcategory';
      }
    }
    
    // If match found, add to results
    if (matchField) {
      // Build transaction object
      var transaction = {
        date: row[dateCol - 1] || null,
        type: row[typeCol - 1] || '',
        account: row[accountCol - 1] || '',
        amount: row[amountCol - 1] || 0,
        merchant: merchantCol ? (row[merchantCol - 1] || '') : '',
        description: descriptionCol ? (row[descriptionCol - 1] || '') : '',
        category: categoryCol ? (row[categoryCol - 1] || '') : '',
        subcategory: subcategoryCol ? (row[subcategoryCol - 1] || '') : ''
      };
      
      results.push({
        rowNum: rowNum,
        transaction: transaction,
        matchField: matchField
      });
    }
  }
  
  return results;
}

/**
 * Filter transactions by date range, category, account, type.
 * @param {Object} filters - Filter criteria
 * @param {Date} [filters.startDate] - Start date (inclusive)
 * @param {Date} [filters.endDate] - End date (inclusive)
 * @param {string} [filters.category] - Category name
 * @param {string} [filters.account] - Account name
 * @param {string} [filters.type] - Transaction type (expense/income/transfer)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet
 * @returns {Array<Object>} Array of filtered transactions with row numbers
 */
function pfFilterTransactions_(filters, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return [];
  }
  
  filters = filters || {};
  
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet) {
    return [];
  }
  
  var lastRow = txSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  
  // Read all data
  var data = txSheet.getRange(1, 1, lastRow, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
  
  // Find column indices
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var categoryCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category');
  var accountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Account');
  var typeCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
  var merchantCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Merchant');
  var descriptionCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Description');
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  var subcategoryCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Subcategory');
  
  var results = [];
  
  // Filter through data rows (skip header)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowNum = i + 1; // 1-based row number
    var matches = true;
    
    // Filter by date range
    if (matches && dateCol && (filters.startDate || filters.endDate)) {
      var rowDate = row[dateCol - 1];
      if (rowDate instanceof Date) {
        if (filters.startDate && rowDate < filters.startDate) {
          matches = false;
        }
        if (matches && filters.endDate) {
          // End date should be inclusive, so add 1 day and compare
          var endDatePlusOne = new Date(filters.endDate);
          endDatePlusOne.setDate(endDatePlusOne.getDate() + 1);
          if (rowDate >= endDatePlusOne) {
            matches = false;
          }
        }
      } else {
        matches = false;
      }
    }
    
    // Filter by category
    if (matches && filters.category && categoryCol) {
      var rowCategory = String(row[categoryCol - 1] || '').trim();
      if (rowCategory !== filters.category) {
        matches = false;
      }
    }
    
    // Filter by account
    if (matches && filters.account && accountCol) {
      var rowAccount = String(row[accountCol - 1] || '').trim();
      if (rowAccount !== filters.account) {
        matches = false;
      }
    }
    
    // Filter by type
    if (matches && filters.type && typeCol) {
      var rowType = String(row[typeCol - 1] || '').trim();
      if (rowType !== filters.type) {
        matches = false;
      }
    }
    
    // If all filters match, add to results
    if (matches) {
      var transaction = {
        date: row[dateCol - 1] || null,
        type: row[typeCol - 1] || '',
        account: row[accountCol - 1] || '',
        amount: row[amountCol - 1] || 0,
        merchant: merchantCol ? (row[merchantCol - 1] || '') : '',
        description: descriptionCol ? (row[descriptionCol - 1] || '') : '',
        category: categoryCol ? (row[categoryCol - 1] || '') : '',
        subcategory: subcategoryCol ? (row[subcategoryCol - 1] || '') : ''
      };
      
      results.push({
        rowNum: rowNum,
        transaction: transaction
      });
    }
  }
  
  return results;
}

/**
 * Highlight search results in the Transactions sheet.
 * @param {Array<Object>} results - Search results
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet
 */
function pfHighlightSearchResults_(results, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || !results || results.length === 0) {
    return;
  }
  
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet) {
    return;
  }
  
  // Clear previous highlights (optional - could be improved to track previous highlights)
  // For now, we'll just highlight the new results
  
  // Highlight each result
  for (var i = 0; i < results.length; i++) {
    var result = results[i];
    var rowNum = result.rowNum;
    
    if (rowNum > 1 && rowNum <= txSheet.getLastRow()) {
      var row = txSheet.getRange(rowNum, 1, 1, PF_TRANSACTIONS_SCHEMA.columns.length);
      row.setBackground('#fff4cc'); // Light yellow
      
      // Add note with match information
      if (result.matchField) {
        var note = 'Найдено по полю: ' + result.matchField;
        row.setNote(note);
      }
    }
  }
  
  // Navigate to first result
  if (results.length > 0) {
    var firstRow = results[0].rowNum;
    if (firstRow > 1) {
      txSheet.setActiveRange(txSheet.getRange(firstRow, 1));
      ss.setActiveSheet(txSheet);
    }
  }
}

/**
 * Clear search highlights from Transactions sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Optional spreadsheet
 */
function pfClearSearchHighlights_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return;
  }
  
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet) {
    return;
  }
  
  var lastRow = txSheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }
  
  // Clear background and notes from all data rows
  for (var row = 2; row <= lastRow; row++) {
    var rowRange = txSheet.getRange(row, 1, 1, PF_TRANSACTIONS_SCHEMA.columns.length);
    var currentBg = rowRange.getBackground();
    // Only clear if it's the search highlight color
    if (currentBg === '#fff4cc') {
      rowRange.setBackground(null);
      rowRange.setNote('');
    }
  }
}

/**
 * Public function: Show search dialog (HTML sidebar).
 */
function pfShowSearchDialog() {
  var html = HtmlService.createHtmlOutputFromFile('SearchUI')
    .setWidth(500)
    .setHeight(600)
    .setTitle(pfT_('menu.search_transactions'));
  
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Public function: Search transactions (called from HTML).
 * @param {string} query - Search query
 * @returns {Object} Search results
 */
function pfSearchTransactions(query) {
  try {
    var results = pfSearchTransactions_(query);
    
    // Format results for display
    var formattedResults = [];
    for (var i = 0; i < results.length; i++) {
      var r = results[i];
      var tx = r.transaction;
      
      // Format date
      var dateStr = '';
      if (tx.date instanceof Date) {
        dateStr = Utilities.formatDate(tx.date, Session.getScriptTimeZone(), 'dd.MM.yyyy');
      }
      
      formattedResults.push({
        rowNum: r.rowNum,
        date: dateStr,
        type: tx.type,
        account: tx.account,
        amount: tx.amount,
        merchant: tx.merchant,
        description: tx.description,
        category: tx.category,
        subcategory: tx.subcategory,
        matchField: r.matchField
      });
    }
    
    return {
      success: true,
      count: formattedResults.length,
      results: formattedResults
    };
  } catch (e) {
    pfLogError_(e, 'pfSearchTransactions', PF_LOG_LEVEL.ERROR);
    return {
      success: false,
      error: e.message || e.toString(),
      count: 0,
      results: []
    };
  }
}

/**
 * Public function: Filter transactions (called from HTML).
 * @param {Object} filters - Filter criteria
 * @returns {Object} Filter results
 */
function pfFilterTransactions(filters) {
  try {
    var results = pfFilterTransactions_(filters);
    
    // Format results for display
    var formattedResults = [];
    for (var i = 0; i < results.length; i++) {
      var r = results[i];
      var tx = r.transaction;
      
      // Format date
      var dateStr = '';
      if (tx.date instanceof Date) {
        dateStr = Utilities.formatDate(tx.date, Session.getScriptTimeZone(), 'dd.MM.yyyy');
      }
      
      formattedResults.push({
        rowNum: r.rowNum,
        date: dateStr,
        type: tx.type,
        account: tx.account,
        amount: tx.amount,
        merchant: tx.merchant,
        description: tx.description,
        category: tx.category,
        subcategory: tx.subcategory
      });
    }
    
    return {
      success: true,
      count: formattedResults.length,
      results: formattedResults
    };
  } catch (e) {
    pfLogError_(e, 'pfFilterTransactions', PF_LOG_LEVEL.ERROR);
    return {
      success: false,
      error: e.message || e.toString(),
      count: 0,
      results: []
    };
  }
}

/**
 * Public function: Highlight and navigate to row.
 * @param {number} rowNum - Row number (1-based)
 */
function pfGoToTransactionRow(rowNum) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    
    if (!txSheet) {
      return { success: false, error: 'Transactions sheet not found' };
    }
    
    if (rowNum < 2 || rowNum > txSheet.getLastRow()) {
      return { success: false, error: 'Invalid row number' };
    }
    
    // Navigate to row
    txSheet.setActiveRange(txSheet.getRange(rowNum, 1));
    ss.setActiveSheet(txSheet);
    
    return { success: true };
  } catch (e) {
    pfLogError_(e, 'pfGoToTransactionRow', PF_LOG_LEVEL.ERROR);
    return { success: false, error: e.message || e.toString() };
  }
}

/**
 * Public function: Highlight search results.
 * @param {Array<number>} rowNums - Array of row numbers (1-based)
 */
function pfHighlightTransactionRows(rowNums) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    
    if (!txSheet || !rowNums || rowNums.length === 0) {
      return { success: false };
    }
    
    // Clear previous highlights
    pfClearSearchHighlights_(ss);
    
    // Highlight new results
    for (var i = 0; i < rowNums.length; i++) {
      var rowNum = rowNums[i];
      if (rowNum > 1 && rowNum <= txSheet.getLastRow()) {
        var row = txSheet.getRange(rowNum, 1, 1, PF_TRANSACTIONS_SCHEMA.columns.length);
        row.setBackground('#fff4cc'); // Light yellow
      }
    }
    
    // Navigate to first result
    if (rowNums.length > 0 && rowNums[0] > 1) {
      txSheet.setActiveRange(txSheet.getRange(rowNums[0], 1));
      ss.setActiveSheet(txSheet);
    }
    
    return { success: true };
  } catch (e) {
    pfLogError_(e, 'pfHighlightTransactionRows', PF_LOG_LEVEL.ERROR);
    return { success: false, error: e.message || e.toString() };
  }
}

/**
 * Public function: Clear search highlights.
 */
function pfClearSearchHighlights() {
  try {
    pfClearSearchHighlights_();
    return { success: true };
  } catch (e) {
    pfLogError_(e, 'pfClearSearchHighlights', PF_LOG_LEVEL.ERROR);
    return { success: false, error: e.message || e.toString() };
  }
}

/**
 * Public function: Get accounts for filter dropdown.
 * @returns {Array<string>} Array of account names
 */
function pfGetAccountsForSearch() {
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
    
    var accountCol = pfColumnIndex_(PF_ACCOUNTS_SCHEMA, 'Account');
    var activeCol = pfColumnIndex_(PF_ACCOUNTS_SCHEMA, 'Active');
    
    if (!accountCol) {
      return [];
    }
    
    var accounts = [];
    var data = accountsSheet.getRange(2, accountCol, lastRow - 1, 1).getValues();
    var activeData = activeCol ? accountsSheet.getRange(2, activeCol, lastRow - 1, 1).getValues() : null;
    
    for (var i = 0; i < data.length; i++) {
      var account = String(data[i][0] || '').trim();
      if (account) {
        // Check if account is active (if Active column exists)
        var isActive = true;
        if (activeData) {
          var activeValue = activeData[i][0];
          isActive = activeValue === true || activeValue === 'TRUE' || activeValue === 'true' || activeValue === '';
        }
        
        if (isActive) {
          accounts.push(account);
        }
      }
    }
    
    return accounts.sort();
  } catch (e) {
    pfLogError_(e, 'pfGetAccountsForSearch', PF_LOG_LEVEL.ERROR);
    return [];
  }
}

/**
 * Public function: Get categories for filter dropdown.
 * @returns {Array<string>} Array of category names
 */
function pfGetCategoriesForSearch() {
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
    
    var categoryCol = pfColumnIndex_(PF_CATEGORIES_SCHEMA, 'Category');
    var activeCol = pfColumnIndex_(PF_CATEGORIES_SCHEMA, 'Active');
    
    if (!categoryCol) {
      return [];
    }
    
    var categories = [];
    var data = categoriesSheet.getRange(2, categoryCol, lastRow - 1, 1).getValues();
    var activeData = activeCol ? categoriesSheet.getRange(2, activeCol, lastRow - 1, 1).getValues() : null;
    
    for (var i = 0; i < data.length; i++) {
      var category = String(data[i][0] || '').trim();
      if (category) {
        // Check if category is active (if Active column exists)
        var isActive = true;
        if (activeData) {
          var activeValue = activeData[i][0];
          isActive = activeValue === true || activeValue === 'TRUE' || activeValue === 'true' || activeValue === '';
        }
        
        if (isActive) {
          categories.push(category);
        }
      }
    }
    
    return categories.sort();
  } catch (e) {
    pfLogError_(e, 'pfGetCategoriesForSearch', PF_LOG_LEVEL.ERROR);
    return [];
  }
}
