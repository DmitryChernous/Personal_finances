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
      
      var data = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
      var categoryIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category') - 1;
      var amountIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount') - 1;
      var typeIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type') - 1;
      var statusIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status') - 1;
      var dateIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date') - 1;
      
      var categoryTotals = {};
      
      for (var i = 0; i < data.length; i++) {
        var rowData = data[i];
        var date = rowData[dateIdx];
        var type = rowData[typeIdx];
        var status = rowData[statusIdx];
        var category = rowData[categoryIdx];
        var amount = rowData[amountIdx];
        
        // Filter: current month, expense, ok status, has category.
        if (date && date >= monthStart && date <= monthEnd && 
            type === 'expense' && status === 'ok' && category && String(category).trim() !== '') {
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

  // QUERY to get monthly aggregation for last 12 months.
  // This is complex - for now, we'll use a simpler approach: generate via helper formulas or leave for future enhancement.
  // Placeholder: user can filter Transactions sheet by month manually or we implement this later.
  if (lang === 'en') {
    reportsSheet.getRange(row + 2, 1).setValue('(Use Transactions sheet filter by month)');
  } else {
    reportsSheet.getRange(row + 2, 1).setValue('(Используйте фильтр листа Транзакции по месяцам)');
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

  // For account balances, we'd need to:
  // 1. Get initial balance from Accounts sheet
  // 2. Sum all income/expense transactions for that account
  // 3. Handle transfers (subtract from source, add to destination)
  // This is complex, so for now we'll use a simple approach: sum income minus expenses per account.
  if (accountCol && amountCol && typeCol && statusCol) {
    var accountsSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
    if (accountsSheet && accountsSheet.getLastRow() > 1) {
      // Get list of accounts.
      var accountsRange = accountsSheet.getRange(2, 1, accountsSheet.getLastRow() - 1, 1);
      var accounts = accountsRange.getValues();
      
      // For each account, calculate balance: initial + income - expenses (excluding transfers).
      // Place formula in row + 2, row + 3, etc.
      var balanceRow = row + 2;
      for (var i = 0; i < accounts.length && i < 20; i++) { // Limit to 20 accounts.
        var accountName = accounts[i][0];
        if (!accountName || String(accountName).trim() === '') continue;
        
        reportsSheet.getRange(balanceRow, 1).setValue(accountName);
        
        // Balance = InitialBalance (from Accounts) + SUMIFS(income) - SUMIFS(expenses) for this account.
        // For now, just sum income - expenses (we'll add initial balance later if needed).
        var accountBalanceFormula = '=SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + accountCol + '2:' + accountCol + ';"' + accountName + '";\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"income";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok")-SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + accountCol + '2:' + accountCol + ';"' + accountName + '";\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"expense";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok")';
        reportsSheet.getRange(balanceRow, 2).setFormula(accountBalanceFormula);
        balanceRow++;
      }
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
