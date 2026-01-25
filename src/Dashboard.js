/**
 * Dashboard initialization and charts.
 *
 * Creates a visual dashboard with KPI and charts on the Dashboard sheet.
 * Uses Google Sheets Charts API for visualizations.
 */

/**
 * Initializes or updates the Dashboard sheet with KPI and charts.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function pfInitializeDashboard_(ss) {
  var dashboardSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.DASHBOARD);
  var txSheetName = pfGetTransactionsSheetName_(ss);
  var lang = pfGetLanguage_();

  // Clear existing content (but preserve charts if they exist).
  var lastRow = dashboardSheet.getLastRow();
  var lastCol = dashboardSheet.getLastColumn();
  if (lastRow > 0 && lastCol > 0) {
    // Clear only data, not charts (charts are separate objects).
    dashboardSheet.getRange(1, 1, lastRow, lastCol).clearContent();
    dashboardSheet.getRange(1, 1, lastRow, lastCol).clearFormat();
  }

  // Remove existing charts.
  var charts = dashboardSheet.getCharts();
  for (var i = 0; i < charts.length; i++) {
    dashboardSheet.removeChart(charts[i]);
  }

  var row = 1;

  // Section 1: KPI Cards (Current Month).
  if (lang === 'en') {
    dashboardSheet.getRange(row, 1).setValue('Key Metrics (Current Month)');
    dashboardSheet.getRange(row + 1, 1).setValue('Income');
    dashboardSheet.getRange(row + 1, 2).setValue('Expenses');
    dashboardSheet.getRange(row + 1, 3).setValue('Net');
    dashboardSheet.getRange(row + 1, 4).setValue('Avg Daily Expense');
  } else {
    dashboardSheet.getRange(row, 1).setValue('Ключевые показатели (текущий месяц)');
    dashboardSheet.getRange(row + 1, 1).setValue('Доходы');
    dashboardSheet.getRange(row + 1, 2).setValue('Расходы');
    dashboardSheet.getRange(row + 1, 3).setValue('Итого');
    dashboardSheet.getRange(row + 1, 4).setValue('Средний расход в день');
  }

  // KPI formulas (reuse logic from Reports).
  var dateCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var amountCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  var typeCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Type');
  var statusCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Status');

  if (amountCol && typeCol && statusCol && dateCol) {
    var monthStart = 'DATE(YEAR(TODAY());MONTH(TODAY());1)';
    var monthEnd = 'EOMONTH(TODAY();0)';
    
    // Income.
    var incomeFormula = '=SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"income";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok";\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';">="&' + monthStart + ';\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';"<="&' + monthEnd + ')';
    
    // Expenses.
    var expenseFormula = '=SUMIFS(\'' + txSheetName + '\'!' + amountCol + '2:' + amountCol + ';\'' + txSheetName + '\'!' + typeCol + '2:' + typeCol + ';"expense";\'' + txSheetName + '\'!' + statusCol + '2:' + statusCol + ';"ok";\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';">="&' + monthStart + ';\'' + txSheetName + '\'!' + dateCol + '2:' + dateCol + ';"<="&' + monthEnd + ')';
    
    // Net (Income - Expenses). Reference data row (row + 2), not header row (row + 1).
    var netFormula = '=' + dashboardSheet.getRange(row + 2, 1).getA1Notation() + '-' + dashboardSheet.getRange(row + 2, 2).getA1Notation();
    
    // Average daily expense (expenses / days in month). Reference data row (row + 2), not header row (row + 1).
    var avgDailyFormula = '=' + dashboardSheet.getRange(row + 2, 2).getA1Notation() + '/DAY(EOMONTH(TODAY();0))';

    dashboardSheet.getRange(row + 2, 1).setFormula(incomeFormula);
    dashboardSheet.getRange(row + 2, 2).setFormula(expenseFormula);
    dashboardSheet.getRange(row + 2, 3).setFormula(netFormula);
    dashboardSheet.getRange(row + 2, 4).setFormula(avgDailyFormula);

    // Format KPI values.
    var kpiRange = dashboardSheet.getRange(row + 2, 1, 1, 4);
    kpiRange.setNumberFormat('#,##0.00');
    kpiRange.setFontSize(14);
    kpiRange.setFontWeight('bold');
  }

  row += 4;

  // Section 2: Expenses by Category (Pie Chart).
  if (lang === 'en') {
    dashboardSheet.getRange(row, 1).setValue('Expenses by Category (Current Month)');
  } else {
    dashboardSheet.getRange(row, 1).setValue('Расходы по категориям (текущий месяц)');
  }

  // Create data range for pie chart (reuse QUERY from Reports).
  var categoryCol = pfColumnLetter_(PF_TRANSACTIONS_SCHEMA, 'Category');
  if (categoryCol && amountCol && typeCol && statusCol && dateCol) {
    var categoryIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category');
    var amountIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
    var typeIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
    var statusIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
    
    var categoryColQuery = 'Col' + (categoryIdx - 1);
    var amountColQuery = 'Col' + (amountIdx - 1);
    
    // Top expenses by category (current month) - use script-based calculation instead of QUERY.
    var categoryLabel = lang === 'en' ? 'Category' : 'Категория';
    var amountLabel = lang === 'en' ? 'Amount' : 'Сумма';
    
    // Set headers manually.
    dashboardSheet.getRange(row + 1, 1).setValue(categoryLabel);
    dashboardSheet.getRange(row + 1, 2).setValue(amountLabel);
    
    // Calculate data via script instead of QUERY formula to avoid #N/A bug.
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    if (txSheet) {
      // Cache lastRow to avoid multiple calls
      var lastRow = txSheet.getLastRow();
      if (lastRow <= 1) {
        // No data rows, skip calculation.
        return;
      }
      
      var today = new Date();
      var monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
      var monthEnd = new Date(today.getFullYear(), today.getMonth() + 1, 0);
      
      // Use cached lastRow
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
        
        // Check if rowData has enough columns.
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
        dashboardSheet.getRange(row + 2, 1, result.length, 2).setValues(result);
        dashboardSheet.getRange(row + 2, 2, result.length, 1).setNumberFormat('#,##0.00');
      }
    }

    // Create pie chart.
    var dataRange = dashboardSheet.getRange(row + 1, 1, 11, 2); // Header + up to 10 categories.
    var chart = dashboardSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dataRange)
      .setPosition(row + 13, 1, 0, 0)
      .setOption('title', lang === 'en' ? 'Expenses by Category' : 'Расходы по категориям')
      .setOption('legend.position', 'right')
      .setOption('pieSliceText', 'percentage')
      .build();
    dashboardSheet.insertChart(chart);
  }

  row += 25; // Leave space for chart.

  // Section 3: Monthly Trend (Line Chart) - Last 6 months.
  if (lang === 'en') {
    dashboardSheet.getRange(row, 1).setValue('Monthly Trend (Last 6 Months)');
  } else {
    dashboardSheet.getRange(row, 1).setValue('Динамика по месяцам (последние 6 месяцев)');
  }

  // For monthly trend, we'll create a helper table with months and sums.
  // This is complex with QUERY, so we'll use a simpler approach: create helper columns.
  // Placeholder for now - can be enhanced later.
  if (lang === 'en') {
    dashboardSheet.getRange(row + 1, 1).setValue('(Use Reports sheet for detailed monthly data)');
  } else {
    dashboardSheet.getRange(row + 1, 1).setValue('(Используйте лист Отчеты для детальных данных по месяцам)');
  }

  row += 3;

  // Format headers.
  var headerRanges = [
    dashboardSheet.getRange(1, 1),
    dashboardSheet.getRange(5, 1),
    dashboardSheet.getRange(30, 1)
  ];
  for (var i = 0; i < headerRanges.length; i++) {
    headerRanges[i].setFontSize(12);
    headerRanges[i].setFontWeight('bold');
  }

  // Auto-resize columns.
  dashboardSheet.autoResizeColumns(1, 4);
}

/**
 * Public function to refresh dashboard (called from menu).
 */
function pfRefreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  pfInitializeDashboard_(ss);
  SpreadsheetApp.flush();
  var lang = pfGetLanguage_();
  if (lang === 'en') {
    SpreadsheetApp.getUi().alert('Dashboard updated');
  } else {
    SpreadsheetApp.getUi().alert('Дашборд обновлён');
  }
}
