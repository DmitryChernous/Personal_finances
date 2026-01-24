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
    
    // Net.
    var netFormula = '=' + dashboardSheet.getRange(row + 1, 1).getA1Notation() + '-' + dashboardSheet.getRange(row + 1, 2).getA1Notation();
    
    // Average daily expense (expenses / days in month).
    var avgDailyFormula = '=' + dashboardSheet.getRange(row + 1, 2).getA1Notation() + '/DAY(EOMONTH(TODAY();0))';

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
    
    // Use FILTER to filter by date range and type/status, then QUERY for grouping.
    // For ru_RU locale, use Russian function names: ДАТА, ГОД, МЕСЯЦ, КОНМЕСЯЦА, СЕГОДНЯ
    // Use exact working formula structure: Col7 for Category (G), Col5 for Amount (E)
    // Labels come AFTER limit, not in SELECT
    var categoryLabel = lang === 'en' ? 'Category' : 'Категория';
    var amountLabel = lang === 'en' ? 'Amount' : 'Сумма';
    
    // Workaround for Google Sheets QUERY #N/A bug: add unique comment to force recalculation.
    var uniqueSuffix = ' '; // Tiny change to force formula refresh
    var categoriesDataFormula = '=QUERY(FILTER(\'' + txSheetName + '\'!A2:N;\'' + txSheetName + '\'!A2:A>=ДАТА(ГОД(СЕГОДНЯ());МЕСЯЦ(СЕГОДНЯ());1);\'' + txSheetName + '\'!A2:A<=КОНМЕСЯЦА(СЕГОДНЯ();0);\'' + txSheetName + '\'!B2:B="expense";\'' + txSheetName + '\'!N2:N="ok");"select Col7, sum(Col5) where Col7 is not null group by Col7 order by sum(Col5) desc limit 10 label Col7 \'' + categoryLabel + uniqueSuffix + '\', sum(Col5) \'' + amountLabel + '\'";1)';
    
    var formulaRange = dashboardSheet.getRange(row + 1, 1, 12, 2); // Up to 10 categories + header
    formulaRange.clearContent();
    formulaRange.clearFormat();
    SpreadsheetApp.flush();
    
    var targetCell = dashboardSheet.getRange(row + 1, 1);
    
    // Try multiple approaches to force formula refresh
    targetCell.setValue(categoriesDataFormula);
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    
    targetCell.setValue('');
    SpreadsheetApp.flush();
    Utilities.sleep(50);
    targetCell.setFormula(categoriesDataFormula);
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    
    // Check if still #N/A and retry
    var currentFormula = targetCell.getFormula();
    if (currentFormula && currentFormula.indexOf('#N/A') !== -1 || !currentFormula) {
      targetCell.setValue('');
      SpreadsheetApp.flush();
      Utilities.sleep(50);
      targetCell.setValue(categoriesDataFormula);
      SpreadsheetApp.flush();
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
