/**
 * Reports generation and formulas.
 *
 * Generates aggregated reports on the Reports sheet using QUERY/SUMIFS formulas.
 * Formulas use ru_RU locale (semicolon as argument separator).
 */

/**
 * Gets the Transactions sheet name (localized).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {string}
 */
function pfGetTransactionsSheetName_(ss) {
  var sheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  return sheet ? sheet.getName() : pfT_('sheet.transactions');
}

/**
 * Gets column letter for a schema column key (1-based index).
 * @param {{columns: Array<{key: string}>}} schema
 * @param {string} key
 * @returns {string|null} Column letter (A, B, etc.)
 */
function pfColumnLetter_(schema, key) {
  var index = pfColumnIndex_(schema, key);
  if (!index) return null;
  // Convert 1-based index to column letter (A=1, B=2, ..., Z=26, AA=27, ...)
  var col = '';
  var num = index;
  while (num > 0) {
    var remainder = (num - 1) % 26;
    col = String.fromCharCode(65 + remainder) + col;
    num = Math.floor((num - 1) / 26);
  }
  return col;
}

/**
 * Initializes or updates the Reports sheet with formulas.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function pfInitializeReports_(ss) {
  var reportsSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.REPORTS);
  var txSheetName = pfGetTransactionsSheetName_(ss);
  var lang = pfGetLanguage_();

  // Clear existing content (but preserve structure if needed).
  if (reportsSheet.getLastRow() > 0) {
    reportsSheet.clear();
  }

  // Column letters for Transactions sheet.
  var dateCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var typeCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Type');
  var amountCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  var categoryCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Category');
  var accountCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Account');
  var statusCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Status');
  var currencyCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Currency');

  var row = 1;

  // Section 1: Summary by period (current month, current year).
  if (lang === 'en') {
    reportsSheet.getRange(row, 1).setValue('Summary by Period');
    reportsSheet.getRange(row + 1, 1).setValue('Current Month');
    reportsSheet.getRange(row + 1, 2).setValue('Income');
    reportsSheet.getRange(row + 1, 3).setValue('Expenses');
    reportsSheet.getRange(row + 1, 4).setValue('Net');
    reportsSheet.getRange(row + 2, 1).setValue('Current Year');
    reportsSheet.getRange(row + 2, 2).setValue('Income');
    reportsSheet.getRange(row + 2, 3).setValue('Expenses');
    reportsSheet.getRange(row + 2, 4).setValue('Net');
  } else {
    reportsSheet.getRange(row, 1).setValue('Сводка по периодам');
    reportsSheet.getRange(row + 1, 1).setValue('Текущий месяц');
    reportsSheet.getRange(row + 1, 2).setValue('Доходы');
    reportsSheet.getRange(row + 1, 3).setValue('Расходы');
    reportsSheet.getRange(row + 1, 4).setValue('Итого');
    reportsSheet.getRange(row + 2, 1).setValue('Текущий год');
    reportsSheet.getRange(row + 2, 2).setValue('Доходы');
    reportsSheet.getRange(row + 2, 3).setValue('Расходы');
    reportsSheet.getRange(row + 2, 4).setValue('Итого');
  }

  // Formulas for current month (income, expenses, net).
  // Use SUMIFS with date range conditions.
  // Income: SUMIFS(Amount, Type="income", Status="ok", Date >= start of month, Date <= end of month)
  if (amountCol && typeCol && statusCol && dateCol) {
    // Current month start: DATE(YEAR(TODAY());MONTH(TODAY());1)
    // Current month end: EOMONTH(TODAY();0)
    var monthStart = 'DATE(YEAR(TODAY());MONTH(TODAY());1)';
    var monthEnd = 'EOMONTH(TODAY();0)';
    
    var monthIncomeFormula = '=SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"income";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok";\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';">="&' + monthStart + ';\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';"<="&' + monthEnd + ')';
    var monthExpenseFormula = '=SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"expense";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok";\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';">="&' + monthStart + ';\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';"<="&' + monthEnd + ')';
    var monthNetFormula = '=' + reportsSheet.getRange(row + 1, 2).getA1Notation() + '-' + reportsSheet.getRange(row + 1, 3).getA1Notation();

    reportsSheet.getRange(row + 1, 2).setFormula(monthIncomeFormula);
    reportsSheet.getRange(row + 1, 3).setFormula(monthExpenseFormula);
    reportsSheet.getRange(row + 1, 4).setFormula(monthNetFormula);

    // Formulas for current year.
    var yearStart = 'DATE(YEAR(TODAY());1;1)';
    var yearEnd = 'DATE(YEAR(TODAY());12;31)';
    
    var yearIncomeFormula = '=SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"income";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok";\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';">="&' + yearStart + ';\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';"<="&' + yearEnd + ')';
    var yearExpenseFormula = '=SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"expense";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok";\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';">="&' + yearStart + ';\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';"<="&' + yearEnd + ')';
    var yearNetFormula = '=' + reportsSheet.getRange(row + 2, 2).getA1Notation() + '-' + reportsSheet.getRange(row + 2, 3).getA1Notation();

    reportsSheet.getRange(row + 2, 2).setFormula(yearIncomeFormula);
    reportsSheet.getRange(row + 2, 3).setFormula(yearExpenseFormula);
    reportsSheet.getRange(row + 2, 4).setFormula(yearNetFormula);
  }

  row += 4;

  // Section 2: Top expenses by category (current month).
  if (lang === 'en') {
    reportsSheet.getRange(row, 1).setValue('Top Expenses by Category (Current Month)');
  } else {
    reportsSheet.getRange(row, 1).setValue('Топ расходов по категориям (текущий месяц)');
  }

  // Top expenses by category (current month) - use script-based calculation instead of QUERY.
  if (categoryCol && amountCol && typeCol && statusCol && dateCol) {
    var categoryLabel = lang === 'en' ? 'Category' : 'Категория';
    var amountLabel = lang === 'en' ? 'Amount' : 'Сумма';
    
    // Set headers manually.
    reportsSheet.getRange(row + 1, 1).setValue(categoryLabel);
    reportsSheet.getRange(row + 1, 2).setValue(amountLabel);
    
    // Calculate data via script instead of QUERY formula to avoid #N/A bug.
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    if (txSheet && txSheet.getLastRow() > 1) {
      var today = new Date();
      var monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
      var monthEnd = new Date(today.getFullYear(), today.getMonth() + 1, 0);
      
      var lastRow = txSheet.getLastRow();
      if (lastRow <= 1) {
        // No data rows, skip calculation.
        return;
      }
      
      var data = txSheet.getRange(2, 1, lastRow - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
      
      // Get column indices and validate them.
      var categoryColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category');
      var amountColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
      var typeColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
      var statusColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
      var dateColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
      
      // Check if all required indices are valid.
      if (!categoryColIdx || !amountColIdx || !typeColIdx || !statusColIdx || !dateColIdx) {
        // Missing required columns, skip calculation.
        return;
      }
      
      var categoryIdx = categoryColIdx - 1;
      var amountIdx = amountColIdx - 1;
      var typeIdx = typeColIdx - 1;
      var statusIdx = statusColIdx - 1;
      var dateIdx = dateColIdx - 1;
      
      var categoryTotals = {};
      
      for (var i = 0; i < data.length; i++) {
        var rowData = data[i];
        
        // Check if row has enough columns.
        if (rowData.length <= categoryIdx || rowData.length <= amountIdx || 
            rowData.length <= typeIdx || rowData.length <= statusIdx || rowData.length <= dateIdx) {
          continue;
        }
        
        var date = rowData[dateIdx];
        var type = rowData[typeIdx];
        var status = rowData[statusIdx];
        var category = rowData[categoryIdx];
        var amount = rowData[amountIdx];
        
        // Filter: current month, expense, ok status, has category.
        if (date && date >= monthStart && date <= monthEnd && 
            type === PF_TRANSACTION_TYPE.EXPENSE && status === PF_TRANSACTION_STATUS.OK && category && String(category).trim() !== '') {
          var cat = String(category).trim();
          if (!categoryTotals[cat]) {
            categoryTotals[cat] = 0;
          }
          categoryTotals[cat] += Number(amount) || 0;
        }
      }
      
      // Convert to array and sort by amount descending.
      var result = [];
      for (var cat in categoryTotals) {
        result.push([cat, categoryTotals[cat]]);
      }
      result.sort(function(a, b) { return b[1] - a[1]; });
      result = result.slice(0, 10); // Top 10
      
      // Write results.
      if (result.length > 0) {
        reportsSheet.getRange(row + 2, 1, result.length, 2).setValues(result);
        reportsSheet.getRange(row + 2, 2, result.length, 1).setNumberFormat('#,##0.00');
      }
    }
  }

  row += 12; // Leave space for up to 10 categories + header.

  // Section 3: Monthly cashflow (last 12 months).
  // Use QUERY to group by month.
  if (lang === 'en') {
    reportsSheet.getRange(row, 1).setValue('Monthly Cashflow (Last 12 Months)');
    reportsSheet.getRange(row + 1, 1).setValue('Month');
    reportsSheet.getRange(row + 1, 2).setValue('Income');
    reportsSheet.getRange(row + 1, 3).setValue('Expenses');
    reportsSheet.getRange(row + 1, 4).setValue('Net');
  } else {
    reportsSheet.getRange(row, 1).setValue('Денежный поток по месяцам (последние 12 месяцев)');
    reportsSheet.getRange(row + 1, 1).setValue('Месяц');
    reportsSheet.getRange(row + 1, 2).setValue('Доходы');
    reportsSheet.getRange(row + 1, 3).setValue('Расходы');
    reportsSheet.getRange(row + 1, 4).setValue('Итого');
  }

  // Monthly cashflow for last 12 months - use script-based calculation.
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (txSheet) {
    // Cache lastRow to avoid multiple calls
    var lastRow = txSheet.getLastRow();
    if (lastRow > 1) {
      var today = new Date();
      var monthlyData = [];
      
      // Get all data once (cache for all months)
      var data = txSheet.getRange(2, 1, lastRow - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
      
      // Get column indices and validate them (once, outside loop)
      var amountColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
      var typeColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
      var statusColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
      var dateColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
      
      if (amountColIdx && typeColIdx && statusColIdx && dateColIdx) {
        var amountIdx = amountColIdx - 1;
        var typeIdx = typeColIdx - 1;
        var statusIdx = statusColIdx - 1;
        var dateIdx = dateColIdx - 1;
        
        // Calculate for last 12 months (including current month).
        for (var monthOffset = 11; monthOffset >= 0; monthOffset--) {
          var targetDate = new Date(today.getFullYear(), today.getMonth() - monthOffset, 1);
          var monthStart = new Date(targetDate.getFullYear(), targetDate.getMonth(), 1);
          var monthEnd = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0);
          
          var monthLabel = '';
          if (lang === 'en') {
            var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            monthLabel = monthNames[targetDate.getMonth()] + ' ' + targetDate.getFullYear();
          } else {
            var monthNamesRu = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
            monthLabel = monthNamesRu[targetDate.getMonth()] + ' ' + targetDate.getFullYear();
          }
          
          var income = 0;
          var expense = 0;
          
          for (var i = 0; i < data.length; i++) {
            var rowData = data[i];
            
            // Check if row has enough columns.
            if (rowData.length <= amountIdx || rowData.length <= typeIdx || 
                rowData.length <= statusIdx || rowData.length <= dateIdx) {
              continue;
            }
            
            var date = rowData[dateIdx];
            var type = rowData[typeIdx];
            var status = rowData[statusIdx];
            var amount = rowData[amountIdx];
            
            // Filter: current month, ok status, exclude transfers.
            if (date && date >= monthStart && date <= monthEnd && status === PF_TRANSACTION_STATUS.OK && type !== PF_TRANSACTION_TYPE.TRANSFER) {
              if (type === PF_TRANSACTION_TYPE.INCOME) {
                income += Number(amount) || 0;
              } else if (type === PF_TRANSACTION_TYPE.EXPENSE) {
                expense += Number(amount) || 0;
              }
            }
          }
          
          var net = income - expense;
          monthlyData.push([monthLabel, income, expense, net]);
        }
        
        // Write monthly data.
        if (monthlyData.length > 0) {
          reportsSheet.getRange(row + 2, 1, monthlyData.length, 4).setValues(monthlyData);
          reportsSheet.getRange(row + 2, 2, monthlyData.length, 3).setNumberFormat('#,##0.00');
        }
      }
    }
  }

  row += 14; // Leave space for 12 months + header.

  // Section 4: Account balances (if we track balances).
  if (lang === 'en') {
    reportsSheet.getRange(row, 1).setValue('Account Balances');
    reportsSheet.getRange(row + 1, 1).setValue('Account');
    reportsSheet.getRange(row + 1, 2).setValue('Balance');
  } else {
    reportsSheet.getRange(row, 1).setValue('Остатки по счетам');
    reportsSheet.getRange(row + 1, 1).setValue('Счет');
    reportsSheet.getRange(row + 1, 2).setValue('Остаток');
  }

  // Account balances calculation:
  // 1. Get initial balance from Accounts sheet
  // 2. Sum all income/expense transactions for that account
  // 3. Handle transfers (subtract from source, add to destination)
  if (accountCol && amountCol && typeCol && statusCol) {
    var accountsSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
    if (accountsSheet && accountsSheet.getLastRow() > 1) {
      // Get accounts with initial balances.
      var accountsDataRange = accountsSheet.getRange(2, 1, accountsSheet.getLastRow() - 1, PF_ACCOUNTS_SCHEMA.columns.length);
      var accountsData = accountsDataRange.getValues();
      
      // Get column indices (1-based), then convert to 0-based.
      var accountNameColIdx = pfColumnIndex_(PF_ACCOUNTS_SCHEMA, 'Account');
      var initialBalanceColIdx = pfColumnIndex_(PF_ACCOUNTS_SCHEMA, 'InitialBalance');
      var accountToCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'AccountTo');
      
      // Check if column indices are valid.
      if (accountNameColIdx && initialBalanceColIdx) {
      
      var accountNameIdx = accountNameColIdx - 1; // Convert to 0-based
      var initialBalanceIdx = initialBalanceColIdx - 1; // Convert to 0-based
      
      // Build a map of account names to initial balances.
      var initialBalances = {};
      for (var i = 0; i < accountsData.length; i++) {
        // Check if row has enough columns.
        if (accountsData[i].length <= accountNameIdx || accountsData[i].length <= initialBalanceIdx) {
          continue;
        }
        
        var accountName = accountsData[i][accountNameIdx];
        if (!accountName || String(accountName).trim() === '') continue;
        
        var initialBalance = accountsData[i][initialBalanceIdx];
        // Safely convert to number.
        var balanceNum = 0;
        if (initialBalance !== null && initialBalance !== undefined && initialBalance !== '') {
          var parsed = Number(initialBalance);
          if (!isNaN(parsed)) {
            balanceNum = parsed;
          }
        }
        initialBalances[String(accountName).trim()] = balanceNum;
      }
      
      // For each account, calculate balance using formula:
      // Balance = InitialBalance + SUMIFS(income) - SUMIFS(expenses) 
      //         - SUMIFS(transfers from this account) + SUMIFS(transfers to this account)
      var balanceRow = row + 2;
      var accountCount = 0;
      for (var accountName in initialBalances) {
        if (accountCount >= 20) break; // Limit to 20 accounts.
        
        reportsSheet.getRange(balanceRow, 1).setValue(accountName);
        
        // Get initial balance value (or 0 if not set).
        var initialBalanceValue = initialBalances[accountName];
        if (isNaN(initialBalanceValue) || initialBalanceValue === null || initialBalanceValue === undefined) {
          initialBalanceValue = 0;
        }
        
        // Income: SUMIFS for income transactions.
        var incomePart = 'SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + accountCol + '2:' + accountCol + ';"' + accountName + '";\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"income";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok")';
        
        // Expenses: SUMIFS for expense transactions.
        var expensePart = 'SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + accountCol + '2:' + accountCol + ';"' + accountName + '";\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"expense";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok")';
        
        // Transfers out: subtract when this account is the source.
        var transfersOutPart = '';
        if (accountToCol) {
          transfersOutPart = 'SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + accountCol + '2:' + accountCol + ';"' + accountName + '";\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"transfer";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok")';
        }
        
        // Transfers in: add when this account is the destination.
        var transfersInPart = '';
        if (accountToCol) {
          transfersInPart = 'SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + accountToCol + '2:' + accountToCol + ';"' + accountName + '";\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"transfer";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok")';
        }
        
        // Build final formula: InitialBalance + Income - Expenses - TransfersOut + TransfersIn
        var accountBalanceFormula = '=' + initialBalanceValue + '+' + incomePart + '-' + expensePart;
        if (transfersOutPart) {
          accountBalanceFormula += '-' + transfersOutPart;
        }
        if (transfersInPart) {
          accountBalanceFormula += '+' + transfersInPart;
        }
        
        reportsSheet.getRange(balanceRow, 2).setFormula(accountBalanceFormula);
        balanceRow++;
        accountCount++;
      }
      } // Close if (accountNameColIdx && initialBalanceColIdx)
    }
  }

  // Format number columns.
  var numFormatRange = reportsSheet.getRange(2, 2, reportsSheet.getLastRow(), 3);
  numFormatRange.setNumberFormat('#,##0.00');

  // Auto-resize columns.
  reportsSheet.autoResizeColumns(1, 4);
}

/**
 * Public function to refresh reports (called from menu).
 */
function pfRefreshReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  pfInitializeReports_(ss);
  
  // Additional flush and small delay to help with QUERY formula refresh issues.
  SpreadsheetApp.flush();
  Utilities.sleep(100); // Small delay to allow Sheets to process
  SpreadsheetApp.flush();
  
  var lang = pfGetLanguage_();
  if (lang === 'en') {
    SpreadsheetApp.getUi().alert('Reports updated');
  } else {
    SpreadsheetApp.getUi().alert('Отчёты обновлены');
  }
}
