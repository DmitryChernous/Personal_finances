/**
 * Budgets module.
 * 
 * Handles budget planning, calculation of actuals, and status tracking.
 */

/**
 * Initialize Budgets sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function pfInitializeBudgets_(ss) {
  try {
    var sheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.BUDGETS);
    
    // Set headers
    pfEnsureHeaderRow_(sheet, PF_BUDGETS_SCHEMA);
    
    // Apply formatting
    sheet.setFrozenRows(1);
    pfEnsureFilter_(sheet, PF_BUDGETS_SCHEMA.columns.length);
    
    // Format columns
    var amountCol = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Amount');
    var factCol = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Fact');
    var remainingCol = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Remaining');
    var percentCol = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'PercentUsed');
    
    if (amountCol) {
      sheet.getRange(2, amountCol, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('#,##0.00');
    }
    if (factCol) {
      sheet.getRange(2, factCol, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('#,##0.00');
    }
    if (remainingCol) {
      sheet.getRange(2, remainingCol, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('#,##0.00');
    }
    if (percentCol) {
      sheet.getRange(2, percentCol, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('0.00%');
    }
    
    // Validations
    var periodCol = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Period');
    var categoryCol = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Category');
    var amountColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Amount');
    
    if (periodCol) {
      var rulePeriod = SpreadsheetApp.newDataValidation()
        .requireValueInList([PF_BUDGET_PERIOD.MONTH, PF_BUDGET_PERIOD.YEAR], true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, periodCol, sheet.getMaxRows() - 1, 1)
        .setDataValidation(rulePeriod);
    }
    
    if (categoryCol) {
      var categoriesRange = ss.getRangeByName(PF_NAMED_RANGES.CATEGORIES);
      if (categoriesRange) {
        var ruleCategory = SpreadsheetApp.newDataValidation()
          .requireValueInRange(categoriesRange, true)
          .setAllowInvalid(true)
          .build();
        sheet.getRange(2, categoryCol, sheet.getMaxRows() - 1, 1)
          .setDataValidation(ruleCategory);
      }
    }
    
    if (amountColIdx) {
      var ruleAmount = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(0)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, amountColIdx, sheet.getMaxRows() - 1, 1)
        .setDataValidation(ruleAmount);
    }
    
    pfLogInfo_('Budgets sheet initialized successfully', 'pfInitializeBudgets_');
  } catch (e) {
    pfLogError_(e, 'pfInitializeBudgets_', PF_LOG_LEVEL.ERROR);
    throw e;
  }
}

/**
 * Calculate actual expenses/income for a category in a period.
 * @param {string} category - Category name
 * @param {string} subcategory - Subcategory name (optional)
 * @param {string} periodValue - Period value (YYYY-MM for month, YYYY for year)
 * @param {string} period - Period type ('month' or 'year')
 * @returns {number} Actual amount
 */
function pfCalculateBudgetFact_(category, subcategory, periodValue, period) {
  // Validation
  if (!category || typeof category !== 'string') {
    pfLogWarning_('Invalid category in pfCalculateBudgetFact_: ' + category, 'pfCalculateBudgetFact_');
    return 0;
  }
  
  if (!periodValue || typeof periodValue !== 'string') {
    pfLogWarning_('Invalid periodValue in pfCalculateBudgetFact_: ' + periodValue, 'pfCalculateBudgetFact_');
    return 0;
  }
  
  if (!period || (period !== PF_BUDGET_PERIOD.MONTH && period !== PF_BUDGET_PERIOD.YEAR)) {
    pfLogWarning_('Invalid period in pfCalculateBudgetFact_: ' + period, 'pfCalculateBudgetFact_');
    return 0;
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    
    if (!txSheet) {
      pfLogWarning_('Transactions sheet not found', 'pfCalculateBudgetFact_');
      return 0;
    }
    
    // Determine date range
    var dateStart, dateEnd;
    if (period === PF_BUDGET_PERIOD.MONTH) {
      // periodValue format: YYYY-MM
      var parts = periodValue.split('-');
      if (parts.length !== 2) {
        pfLogWarning_('Invalid month format: ' + periodValue, 'pfCalculateBudgetFact_');
        return 0;
      }
      var year = parseInt(parts[0], 10);
      var month = parseInt(parts[1], 10) - 1; // Month is 0-based
      if (isNaN(year) || isNaN(month) || month < 0 || month > 11) {
        pfLogWarning_('Invalid month values: ' + periodValue, 'pfCalculateBudgetFact_');
        return 0;
      }
      dateStart = new Date(year, month, 1);
      dateEnd = new Date(year, month + 1, 0); // Last day of month
    } else if (period === PF_BUDGET_PERIOD.YEAR) {
      // periodValue format: YYYY
      var year = parseInt(periodValue, 10);
      if (isNaN(year)) {
        pfLogWarning_('Invalid year format: ' + periodValue, 'pfCalculateBudgetFact_');
        return 0;
      }
      dateStart = new Date(year, 0, 1);
      dateEnd = new Date(year, 11, 31);
    }
    
    // Get column indices
    var categoryColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category');
    var subcategoryColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Subcategory');
    var amountColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
    var typeColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
    var statusColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
    var dateColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
    
    if (!categoryColIdx || !amountColIdx || !typeColIdx || !statusColIdx || !dateColIdx) {
      pfLogWarning_('Missing required columns in Transactions schema', 'pfCalculateBudgetFact_');
      return 0;
    }
    
    // Cache lastRow and read data
    var lastRow = txSheet.getLastRow();
    if (lastRow <= 1) {
      pfLogDebug_('No transactions data (lastRow=' + lastRow + ')', 'pfCalculateBudgetFact_');
      return 0; // No data
    }
    
    var data = txSheet.getRange(2, 1, lastRow - 1, PF_TRANSACTIONS_SCHEMA.columns.length)
      .getValues();
    
    pfLogDebug_('Read ' + data.length + ' transaction rows', 'pfCalculateBudgetFact_');
    
    // Array indices (0-based)
    var categoryIdx = categoryColIdx - 1;
    var subcategoryIdx = subcategoryColIdx ? subcategoryColIdx - 1 : -1;
    var amountIdx = amountColIdx - 1;
    var typeIdx = typeColIdx - 1;
    var statusIdx = statusColIdx - 1;
    var dateIdx = dateColIdx - 1;
    
    // Sum transactions
    var total = 0;
    var categoryTrimmed = String(category).trim();
    var subcategoryTrimmed = subcategory ? String(subcategory).trim() : '';
    
    pfLogDebug_('pfCalculateBudgetFact_: category=' + categoryTrimmed + ', subcategory=' + subcategoryTrimmed + ', period=' + period + ', periodValue=' + periodValue + ', dateRange=' + dateStart + ' to ' + dateEnd, 'pfCalculateBudgetFact_');
    
    var matchedCount = 0;
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Check array bounds
      if (row.length <= categoryIdx || row.length <= amountIdx || 
          row.length <= typeIdx || row.length <= statusIdx || row.length <= dateIdx) {
        continue;
      }
      
      var rowDate = row[dateIdx];
      var rowCategory = String(row[categoryIdx] || '').trim();
      var rowSubcategory = subcategoryIdx >= 0 ? String(row[subcategoryIdx] || '').trim() : '';
      var rowType = String(row[typeIdx] || '').trim();
      var rowStatus = String(row[statusIdx] || '').trim();
      var rowAmount = Number(row[amountIdx]) || 0;
      
      // Filter
      if (!rowDate || !(rowDate instanceof Date)) {
        pfLogDebug_('Transaction ' + i + ': invalid date (type: ' + typeof rowDate + ', value: ' + rowDate + ')', 'pfCalculateBudgetFact_');
        continue;
      }
      if (rowDate < dateStart || rowDate > dateEnd) {
        pfLogDebug_('Transaction ' + i + ': date out of range (date=' + rowDate + ', range=' + dateStart + ' to ' + dateEnd + ')', 'pfCalculateBudgetFact_');
        continue;
      }
      if (rowCategory !== categoryTrimmed) {
        pfLogDebug_('Transaction ' + i + ': category mismatch (expected: "' + categoryTrimmed + '", got: "' + rowCategory + '", match: ' + (rowCategory === categoryTrimmed) + ')', 'pfCalculateBudgetFact_');
        continue;
      }
      if (subcategoryTrimmed && subcategoryTrimmed !== '' && rowSubcategory !== subcategoryTrimmed) {
        pfLogDebug_('Transaction ' + i + ': subcategory mismatch (expected: "' + subcategoryTrimmed + '", got: "' + rowSubcategory + '", budget has subcategory: ' + (subcategoryTrimmed !== '') + ')', 'pfCalculateBudgetFact_');
        continue;
      }
      if (rowStatus !== PF_TRANSACTION_STATUS.OK) {
        pfLogDebug_('Transaction ' + i + ': status not OK (expected: ' + PF_TRANSACTION_STATUS.OK + ', got: ' + rowStatus + ')', 'pfCalculateBudgetFact_');
        continue;
      }
      // Include both expenses and income (budget can be for both)
      if (rowType !== PF_TRANSACTION_TYPE.EXPENSE && rowType !== PF_TRANSACTION_TYPE.INCOME) {
        pfLogDebug_('Transaction ' + i + ': type not expense/income (expected: ' + PF_TRANSACTION_TYPE.EXPENSE + ' or ' + PF_TRANSACTION_TYPE.INCOME + ', got: ' + rowType + ')', 'pfCalculateBudgetFact_');
        continue;
      }
      
      // Log matched transaction details
      pfLogDebug_('Transaction ' + i + ' MATCHED: date=' + rowDate + ', category=' + rowCategory + ', subcategory=' + rowSubcategory + ', type=' + rowType + ', status=' + rowStatus + ', amount=' + rowAmount, 'pfCalculateBudgetFact_');
      
      // Sum (for expenses, amount is negative in some systems, but we assume positive)
      total += Math.abs(rowAmount);
      matchedCount++;
      pfLogDebug_('Transaction ' + i + ' matched: amount=' + rowAmount + ', total=' + total, 'pfCalculateBudgetFact_');
    }
    
    pfLogDebug_('pfCalculateBudgetFact_ result: total=' + total + ', matchedCount=' + matchedCount + ', totalTransactions=' + data.length, 'pfCalculateBudgetFact_');
    
    return total;
    
  } catch (e) {
    pfLogError_(e, 'pfCalculateBudgetFact_', PF_LOG_LEVEL.ERROR);
    return 0;
  }
}

/**
 * Get budget status based on plan and fact amounts.
 * @param {number} planAmount - Planned amount
 * @param {number} factAmount - Actual amount
 * @returns {string} Status (PF_BUDGET_STATUS.OK, WARNING, or EXCEEDED)
 */
function pfGetBudgetStatus_(planAmount, factAmount) {
  // Validation
  if (typeof planAmount !== 'number' || isNaN(planAmount)) {
    pfLogWarning_('Invalid planAmount: ' + planAmount, 'pfGetBudgetStatus_');
    return PF_BUDGET_STATUS.EXCEEDED;
  }
  
  if (typeof factAmount !== 'number' || isNaN(factAmount)) {
    factAmount = 0;
  }
  
  if (planAmount <= 0) {
    pfLogWarning_('planAmount <= 0: ' + planAmount, 'pfGetBudgetStatus_');
    return PF_BUDGET_STATUS.EXCEEDED;
  }
  
  if (factAmount < 0) {
    pfLogWarning_('factAmount < 0: ' + factAmount, 'pfGetBudgetStatus_');
  }
  
  // Check if exceeded
  if (factAmount > planAmount) {
    return PF_BUDGET_STATUS.EXCEEDED;
  }
  
  // Calculate remaining percentage
  var remaining = planAmount - factAmount;
  var remainingPercent = remaining / planAmount;
  
  // Check warning threshold
  if (remainingPercent <= PF_BUDGET_WARNING_THRESHOLD) {
    return PF_BUDGET_STATUS.WARNING;
  }
  
  return PF_BUDGET_STATUS.OK;
}

/**
 * Update all calculated fields in Budgets sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet (optional)
 * @returns {Object} Result with success flag and statistics
 */
function pfUpdateBudgetCalculations_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    var budgetsSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.BUDGETS);
    
    if (!budgetsSheet) {
      var error = new Error('Budgets sheet not found');
      return pfCreateErrorResponse_('Лист "Бюджеты" не найден. Запустите Setup.', 'BUDGETS_SHEET_NOT_FOUND', error);
    }
    
    var lastRow = budgetsSheet.getLastRow();
    if (lastRow <= 1) {
      pfLogInfo_('No budget data to calculate', 'pfUpdateBudgetCalculations_');
      return pfCreateSuccessResponse_('Нет данных для расчета', { updated: 0 });
    }
    
    // Get column indices
    var categoryColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Category');
    var subcategoryColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Subcategory');
    var periodColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Period');
    var periodValueColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'PeriodValue');
    var amountColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Amount');
    var factColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Fact');
    var remainingColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Remaining');
    var statusColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Status');
    var percentColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'PercentUsed');
    var activeColIdx = pfColumnIndex_(PF_BUDGETS_SCHEMA, 'Active');
    
    if (!categoryColIdx || !periodColIdx || !periodValueColIdx || !amountColIdx ||
        !factColIdx || !remainingColIdx || !statusColIdx || !percentColIdx) {
      var error = new Error('Missing required columns in Budgets schema');
      return pfCreateErrorResponse_('Отсутствуют необходимые колонки в схеме бюджетов', 'MISSING_COLUMNS', error);
    }
    
    // Cache lastRow and read all data
    var data = budgetsSheet.getRange(2, 1, lastRow - 1, PF_BUDGETS_SCHEMA.columns.length)
      .getValues();
    
    pfLogInfo_('Read ' + data.length + ' budget rows from sheet (lastRow=' + lastRow + ')', 'pfUpdateBudgetCalculations_');
    
    // Array indices (0-based)
    var categoryIdx = categoryColIdx - 1;
    var subcategoryIdx = subcategoryColIdx ? subcategoryColIdx - 1 : -1;
    var periodIdx = periodColIdx - 1;
    var periodValueIdx = periodValueColIdx - 1;
    var amountIdx = amountColIdx - 1;
    var factIdx = factColIdx - 1;
    var remainingIdx = remainingColIdx - 1;
    var statusIdx = statusColIdx - 1;
    var percentIdx = percentColIdx - 1;
    var activeIdx = activeColIdx ? activeColIdx - 1 : -1;
    
    pfLogDebug_('Column indices: category=' + categoryIdx + ', subcategory=' + subcategoryIdx + ', period=' + periodIdx + ', periodValue=' + periodValueIdx + ', amount=' + amountIdx, 'pfUpdateBudgetCalculations_');
    
    // Prepare batch update arrays
    var factValues = [];
    var remainingValues = [];
    var statusValues = [];
    var percentValues = [];
    var updatedCount = 0;
    
    // Process each budget row
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Check array bounds
      if (row.length <= categoryIdx || row.length <= periodIdx || 
          row.length <= periodValueIdx || row.length <= amountIdx) {
        factValues.push(['']);
        remainingValues.push(['']);
        statusValues.push(['']);
        percentValues.push(['']);
        continue;
      }
      
      // Check if active
      // If Active column is empty or not explicitly set to false, budget is active by default
      var active = true; // Default: active
      if (activeIdx >= 0) {
        var activeValue = row[activeIdx];
        // Only skip if explicitly set to false
        if (activeValue === false || activeValue === 'false' || activeValue === 'FALSE' || 
            (typeof activeValue === 'string' && activeValue.trim().toLowerCase() === 'false')) {
          pfLogDebug_('Skipping budget row ' + (i + 2) + ': explicitly marked as inactive', 'pfUpdateBudgetCalculations_');
          factValues.push(['']);
          remainingValues.push(['']);
          statusValues.push(['']);
          percentValues.push(['']);
          continue;
        }
        // Empty or any other value (including 'true', 'yes', etc.) means active
      }
      
      var category = String(row[categoryIdx] || '').trim();
      var subcategory = subcategoryIdx >= 0 ? String(row[subcategoryIdx] || '').trim() : '';
      var period = String(row[periodIdx] || '').trim();
      
      // Handle periodValue: Google Sheets may convert "2026-01" to Date object
      var periodValueRaw = row[periodValueIdx];
      var periodValue = '';
      if (periodValueRaw instanceof Date) {
        // Convert Date to YYYY-MM format
        var year = periodValueRaw.getFullYear();
        var month = periodValueRaw.getMonth() + 1; // getMonth() is 0-based
        periodValue = year + '-' + (month < 10 ? '0' + month : month);
        pfLogDebug_('Converted Date to periodValue: ' + periodValueRaw + ' -> ' + periodValue, 'pfUpdateBudgetCalculations_');
      } else {
        periodValue = String(periodValueRaw || '').trim();
      }
      
      var amount = Number(row[amountIdx]) || 0;
      
      pfLogDebug_('Processing budget row ' + (i + 2) + ': category=' + category + ', subcategory=' + subcategory + ', period=' + period + ', periodValue=' + periodValue + ', amount=' + amount, 'pfUpdateBudgetCalculations_');
      
      // Skip if required fields are empty
      if (!category || !period || !periodValue || amount <= 0) {
        pfLogDebug_('Skipping budget row ' + (i + 2) + ': missing required fields', 'pfUpdateBudgetCalculations_');
        factValues.push(['']);
        remainingValues.push(['']);
        statusValues.push(['']);
        percentValues.push(['']);
        continue;
      }
      
      // Calculate fact
      var fact = 0;
      try {
        fact = pfCalculateBudgetFact_(category, subcategory, periodValue, period);
        pfLogDebug_('Calculated fact for row ' + (i + 2) + ': ' + fact, 'pfUpdateBudgetCalculations_');
      } catch (e) {
        pfLogWarning_('Error calculating fact for budget row ' + (i + 2) + ': ' + e.toString(), 'pfUpdateBudgetCalculations_');
        fact = 0;
      }
      
      // Calculate remaining
      var remaining = amount - fact;
      
      // Calculate status
      var status = pfGetBudgetStatus_(amount, fact);
      
      // Calculate percent used
      var percentUsed = amount > 0 ? (fact / amount) : 0;
      
      factValues.push([fact]);
      remainingValues.push([remaining]);
      statusValues.push([status]);
      percentValues.push([percentUsed]);
      
      updatedCount++;
    }
    
    // Batch write values
    if (factValues.length > 0) {
      pfLogDebug_('Writing ' + factValues.length + ' budget calculations to sheet', 'pfUpdateBudgetCalculations_');
      budgetsSheet.getRange(2, factColIdx, factValues.length, 1).setValues(factValues);
      budgetsSheet.getRange(2, remainingColIdx, remainingValues.length, 1).setValues(remainingValues);
      budgetsSheet.getRange(2, statusColIdx, statusValues.length, 1).setValues(statusValues);
      budgetsSheet.getRange(2, percentColIdx, percentValues.length, 1).setValues(percentValues);
      
      // Format percent column
      budgetsSheet.getRange(2, percentColIdx, percentValues.length, 1).setNumberFormat('0.00%');
      
      // Apply conditional formatting (highlight rows)
      for (var i = 0; i < data.length; i++) {
        var rowNum = i + 2;
        var status = statusValues[i][0];
        
        if (status === PF_BUDGET_STATUS.EXCEEDED) {
          budgetsSheet.getRange(rowNum, 1, 1, PF_BUDGETS_SCHEMA.columns.length)
            .setBackground('#ffcccc'); // Light red
        } else if (status === PF_BUDGET_STATUS.WARNING) {
          budgetsSheet.getRange(rowNum, 1, 1, PF_BUDGETS_SCHEMA.columns.length)
            .setBackground('#fff4cc'); // Light yellow
        } else {
          budgetsSheet.getRange(rowNum, 1, 1, PF_BUDGETS_SCHEMA.columns.length)
            .setBackground(null); // Clear background
        }
      }
    }
    
    pfLogInfo_('Updated ' + updatedCount + ' budgets', 'pfUpdateBudgetCalculations_');
    return pfCreateSuccessResponse_('Обновлено бюджетов: ' + updatedCount, { updated: updatedCount });
    
  } catch (e) {
    pfLogError_(e, 'pfUpdateBudgetCalculations_', PF_LOG_LEVEL.ERROR);
    return pfCreateErrorResponse_('Ошибка при обновлении бюджетов: ' + e.toString(), 'UPDATE_ERROR', e);
  }
}

/**
 * Public function to update budget calculations (called from menu).
 */
function pfUpdateBudgetCalculations() {
  try {
    var result = pfUpdateBudgetCalculations_(SpreadsheetApp.getActiveSpreadsheet());
    
    if (result.success) {
      SpreadsheetApp.getUi().alert(result.message || 'Бюджеты обновлены успешно');
    } else {
      SpreadsheetApp.getUi().alert('Ошибка: ' + (result.message || 'Неизвестная ошибка'));
    }
  } catch (e) {
    var handled = pfHandleError_(e, 'pfUpdateBudgetCalculations', 'Ошибка при обновлении бюджетов');
    SpreadsheetApp.getUi().alert('Ошибка: ' + handled.message);
  }
}
