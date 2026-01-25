/**
 * Recurring Transactions module.
 * 
 * Handles recurring transactions (subscriptions, rent, salary, etc.):
 * - Sheet initialization
 * - Validation
 * - Transaction creation logic
 * - Next due date calculation
 */

/**
 * Initialize RecurringTransactions sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function pfInitializeRecurringTransactions_(ss) {
  try {
    var sheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.RECURRING_TRANSACTIONS);
    
    // Set headers
    pfEnsureHeaderRow_(sheet, PF_RECURRING_TRANSACTIONS_SCHEMA);
    
    // Apply formatting
    sheet.setFrozenRows(1);
    pfEnsureFilter_(sheet, PF_RECURRING_TRANSACTIONS_SCHEMA.columns.length);
    
    // Format columns
    var amountCol = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Amount');
    var startDateCol = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'StartDate');
    var endDateCol = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'EndDate');
    var lastCreatedCol = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'LastCreated');
    
    if (amountCol) {
      sheet.getRange(2, amountCol, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('#,##0.00');
    }
    if (startDateCol) {
      sheet.getRange(2, startDateCol, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('dd.MM.yyyy');
    }
    if (endDateCol) {
      sheet.getRange(2, endDateCol, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('dd.MM.yyyy');
    }
    if (lastCreatedCol) {
      sheet.getRange(2, lastCreatedCol, sheet.getMaxRows() - 1, 1)
        .setNumberFormat('dd.MM.yyyy');
    }
    
    // Validations
    var typeCol = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Type');
    var frequencyCol = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Frequency');
    var accountCol = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Account');
    var currencyCol = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Currency');
    var dayOfMonthCol = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'DayOfMonth');
    var dayOfWeekCol = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'DayOfWeek');
    var amountColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Amount');
    
    if (typeCol) {
      var ruleType = SpreadsheetApp.newDataValidation()
        .requireValueInList([PF_TRANSACTION_TYPE.EXPENSE, PF_TRANSACTION_TYPE.INCOME, PF_TRANSACTION_TYPE.TRANSFER], true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, typeCol, sheet.getMaxRows() - 1, 1)
        .setDataValidation(ruleType);
    }
    
    if (frequencyCol) {
      var ruleFrequency = SpreadsheetApp.newDataValidation()
        .requireValueInList([
          PF_RECURRING_FREQUENCY.WEEKLY,
          PF_RECURRING_FREQUENCY.MONTHLY,
          PF_RECURRING_FREQUENCY.QUARTERLY,
          PF_RECURRING_FREQUENCY.YEARLY
        ], true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, frequencyCol, sheet.getMaxRows() - 1, 1)
        .setDataValidation(ruleFrequency);
    }
    
    if (accountCol) {
      var accountsRange = ss.getRangeByName(PF_NAMED_RANGES.ACCOUNTS);
      if (accountsRange) {
        var ruleAccount = SpreadsheetApp.newDataValidation()
          .requireValueInRange(accountsRange, true)
          .setAllowInvalid(true)
          .build();
        sheet.getRange(2, accountCol, sheet.getMaxRows() - 1, 1)
          .setDataValidation(ruleAccount);
      }
    }
    
    if (currencyCol) {
      var ruleCurrency = SpreadsheetApp.newDataValidation()
        .requireValueInList(PF_SUPPORTED_CURRENCIES, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, currencyCol, sheet.getMaxRows() - 1, 1)
        .setDataValidation(ruleCurrency);
    }
    
    if (dayOfMonthCol) {
      var ruleDayOfMonth = SpreadsheetApp.newDataValidation()
        .requireNumberBetween(1, 31)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, dayOfMonthCol, sheet.getMaxRows() - 1, 1)
        .setDataValidation(ruleDayOfMonth);
    }
    
    if (dayOfWeekCol) {
      var ruleDayOfWeek = SpreadsheetApp.newDataValidation()
        .requireNumberBetween(1, 7)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, dayOfWeekCol, sheet.getMaxRows() - 1, 1)
        .setDataValidation(ruleDayOfWeek);
    }
    
    if (amountColIdx) {
      var ruleAmount = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(0)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, amountColIdx, sheet.getMaxRows() - 1, 1)
        .setDataValidation(ruleAmount);
    }
    
    pfLogInfo_('RecurringTransactions sheet initialized successfully', 'pfInitializeRecurringTransactions_');
  } catch (e) {
    pfLogError_(e, 'pfInitializeRecurringTransactions_', PF_LOG_LEVEL.ERROR);
    throw e;
  }
}

/**
 * Check if a transaction should be created for a recurring transaction.
 * @param {Object} recurringTx - Recurring transaction object (row from sheet)
 * @param {Date} [currentDate] - Current date (default: new Date())
 * @returns {boolean} True if transaction should be created
 */
function pfShouldCreateTransaction_(recurringTx, currentDate) {
  currentDate = currentDate || new Date();
  
  try {
    // Check if active
    var active = recurringTx.Active;
    if (active === false || active === 'false' || active === 'FALSE' || 
        (typeof active === 'string' && active.trim().toLowerCase() === 'false')) {
      return false;
    }
    
    // Check StartDate
    var startDate = recurringTx.StartDate;
    if (!startDate || !(startDate instanceof Date)) {
      pfLogWarning_('Invalid StartDate in recurring transaction: ' + recurringTx.Name, 'pfShouldCreateTransaction_');
      return false;
    }
    
    if (currentDate < startDate) {
      return false; // Not started yet
    }
    
    // Check EndDate
    var endDate = recurringTx.EndDate;
    if (endDate && endDate instanceof Date && currentDate > endDate) {
      return false; // Already ended
    }
    
    // Check frequency
    var frequency = String(recurringTx.Frequency || '').trim();
    var lastCreated = recurringTx.LastCreated;
    
    // If LastCreated is empty, use StartDate as base
    var baseDate = lastCreated && lastCreated instanceof Date ? lastCreated : startDate;
    
    // Calculate days/months since base date
    var daysDiff = Math.floor((currentDate - baseDate) / (1000 * 60 * 60 * 24));
    
    if (frequency === PF_RECURRING_FREQUENCY.WEEKLY) {
      // Weekly: check if >= 7 days passed since base date
      // Also check if current day of week matches DayOfWeek (if specified)
      if (daysDiff >= 7) {
        var dayOfWeek = recurringTx.DayOfWeek;
        if (dayOfWeek) {
          // Convert: our DayOfWeek (1=Monday, 7=Sunday) to JS getDay() (0=Sunday, 1=Monday, ..., 6=Saturday)
          var jsDayOfWeek = dayOfWeek === 7 ? 0 : dayOfWeek;
          if (currentDate.getDay() === jsDayOfWeek) {
            return true;
          }
          // If day doesn't match, check if we're past the scheduled day this week
          var currentDay = currentDate.getDay();
          var scheduledDay = jsDayOfWeek;
          // If we're past the scheduled day this week and it's been >= 7 days, create it
          if (currentDay > scheduledDay && daysDiff >= 7) {
            return true;
          }
        } else {
          // No DayOfWeek specified, just check if >= 7 days
          return daysDiff >= 7;
        }
      }
      return false;
    } else if (frequency === PF_RECURRING_FREQUENCY.MONTHLY) {
      // Monthly: check if day of month passed and base date was before current month
      var currentMonthStart = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
      if (baseDate < currentMonthStart) {
        var dayOfMonth = recurringTx.DayOfMonth;
        if (dayOfMonth && currentDate.getDate() >= dayOfMonth) {
          return true;
        }
      }
      return false;
    } else if (frequency === PF_RECURRING_FREQUENCY.QUARTERLY) {
      // Quarterly: check if >= 3 months passed
      var monthsDiff = (currentDate.getFullYear() - baseDate.getFullYear()) * 12 + 
                       (currentDate.getMonth() - baseDate.getMonth());
      if (monthsDiff >= 3) {
        var dayOfMonth = recurringTx.DayOfMonth;
        if (dayOfMonth && currentDate.getDate() >= dayOfMonth) {
          return true;
        }
      }
      return false;
    } else if (frequency === PF_RECURRING_FREQUENCY.YEARLY) {
      // Yearly: check if >= 12 months passed
      var monthsDiff = (currentDate.getFullYear() - baseDate.getFullYear()) * 12 + 
                       (currentDate.getMonth() - baseDate.getMonth());
      if (monthsDiff >= 12) {
        var dayOfMonth = recurringTx.DayOfMonth;
        if (dayOfMonth && currentDate.getDate() >= dayOfMonth) {
          return true;
        }
      }
      return false;
    }
    
    pfLogWarning_('Invalid frequency: ' + frequency, 'pfShouldCreateTransaction_');
    return false;
    
  } catch (e) {
    pfLogError_(e, 'pfShouldCreateTransaction_', PF_LOG_LEVEL.ERROR);
    return false;
  }
}

/**
 * Get next due date for a recurring transaction.
 * @param {Object} recurringTx - Recurring transaction object
 * @returns {Date|null} Next due date or null
 */
function pfGetNextDueDate_(recurringTx) {
  try {
    var baseDate = recurringTx.LastCreated;
    if (!baseDate || !(baseDate instanceof Date)) {
      baseDate = recurringTx.StartDate;
    }
    
    if (!baseDate || !(baseDate instanceof Date)) {
      pfLogWarning_('Invalid StartDate in recurring transaction: ' + recurringTx.Name, 'pfGetNextDueDate_');
      return null;
    }
    
    var frequency = String(recurringTx.Frequency || '').trim();
    var nextDate = new Date(baseDate);
    
    if (frequency === PF_RECURRING_FREQUENCY.WEEKLY) {
      nextDate.setDate(nextDate.getDate() + 7);
    } else if (frequency === PF_RECURRING_FREQUENCY.MONTHLY) {
      nextDate.setMonth(nextDate.getMonth() + 1);
      var dayOfMonth = recurringTx.DayOfMonth;
      if (dayOfMonth) {
        // Set day of month, handle invalid days (e.g., 31 for February)
        var lastDayOfMonth = new Date(nextDate.getFullYear(), nextDate.getMonth() + 1, 0).getDate();
        nextDate.setDate(Math.min(dayOfMonth, lastDayOfMonth));
      }
    } else if (frequency === PF_RECURRING_FREQUENCY.QUARTERLY) {
      nextDate.setMonth(nextDate.getMonth() + 3);
      var dayOfMonth = recurringTx.DayOfMonth;
      if (dayOfMonth) {
        var lastDayOfMonth = new Date(nextDate.getFullYear(), nextDate.getMonth() + 1, 0).getDate();
        nextDate.setDate(Math.min(dayOfMonth, lastDayOfMonth));
      }
    } else if (frequency === PF_RECURRING_FREQUENCY.YEARLY) {
      nextDate.setFullYear(nextDate.getFullYear() + 1);
      var dayOfMonth = recurringTx.DayOfMonth;
      if (dayOfMonth) {
        var lastDayOfMonth = new Date(nextDate.getFullYear(), nextDate.getMonth() + 1, 0).getDate();
        nextDate.setDate(Math.min(dayOfMonth, lastDayOfMonth));
      }
    } else {
      pfLogWarning_('Invalid frequency: ' + frequency, 'pfGetNextDueDate_');
      return null;
    }
    
    // Check EndDate
    var endDate = recurringTx.EndDate;
    if (endDate && endDate instanceof Date && nextDate > endDate) {
      return null; // Beyond end date
    }
    
    return nextDate;
    
  } catch (e) {
    pfLogError_(e, 'pfGetNextDueDate_', PF_LOG_LEVEL.ERROR);
    return null;
  }
}

/**
 * Create transactions for all recurring transactions in a period.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - Spreadsheet (optional)
 * @param {Date} [startDate] - Start date of period (default: start of current month)
 * @param {Date} [endDate] - End date of period (default: end of current month)
 * @returns {Object} Statistics: {created: number, skipped: number, errors: number}
 */
function pfCreateRecurringTransactions_(ss, startDate, endDate) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Default to current month
    var today = new Date();
    if (!startDate) {
      startDate = new Date(today.getFullYear(), today.getMonth(), 1);
    }
    if (!endDate) {
      endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    }
    
    var recurringSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.RECURRING_TRANSACTIONS);
    if (!recurringSheet) {
      var error = new Error('RecurringTransactions sheet not found');
      return pfCreateErrorResponse_('Лист "Регулярные платежи" не найден. Запустите Setup.', 'RECURRING_SHEET_NOT_FOUND', error);
    }
    
    var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    if (!txSheet) {
      var error = new Error('Transactions sheet not found');
      return pfCreateErrorResponse_('Лист "Транзакции" не найден. Запустите Setup.', 'TRANSACTIONS_SHEET_NOT_FOUND', error);
    }
    
    // Cache lastRow and read data
    var lastRow = recurringSheet.getLastRow();
    if (lastRow <= 1) {
      pfLogInfo_('No recurring transactions to process', 'pfCreateRecurringTransactions_');
      return pfCreateSuccessResponse_('Нет регулярных платежей для обработки', { created: 0, skipped: 0, errors: 0 });
    }
    
    // Get column indices
    var nameColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Name');
    var typeColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Type');
    var frequencyColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Frequency');
    var dayOfMonthColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'DayOfMonth');
    var dayOfWeekColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'DayOfWeek');
    var startDateColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'StartDate');
    var endDateColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'EndDate');
    var accountColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Account');
    var accountToColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'AccountTo');
    var amountColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Amount');
    var currencyColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Currency');
    var categoryColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Category');
    var subcategoryColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Subcategory');
    var merchantColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Merchant');
    var descriptionColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Description');
    var activeColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'Active');
    var lastCreatedColIdx = pfColumnIndex_(PF_RECURRING_TRANSACTIONS_SCHEMA, 'LastCreated');
    
    if (!nameColIdx || !typeColIdx || !frequencyColIdx || !startDateColIdx || 
        !accountColIdx || !amountColIdx || !currencyColIdx) {
      var error = new Error('Missing required columns in RecurringTransactions schema');
      return pfCreateErrorResponse_('Отсутствуют необходимые колонки в схеме регулярных платежей', 'MISSING_COLUMNS', error);
    }
    
    var data = recurringSheet.getRange(2, 1, lastRow - 1, PF_RECURRING_TRANSACTIONS_SCHEMA.columns.length)
      .getValues();
    
    // Array indices (0-based)
    var nameIdx = nameColIdx - 1;
    var typeIdx = typeColIdx - 1;
    var frequencyIdx = frequencyColIdx - 1;
    var dayOfMonthIdx = dayOfMonthColIdx ? dayOfMonthColIdx - 1 : -1;
    var dayOfWeekIdx = dayOfWeekColIdx ? dayOfWeekColIdx - 1 : -1;
    var startDateIdx = startDateColIdx - 1;
    var endDateIdx = endDateColIdx ? endDateColIdx - 1 : -1;
    var accountIdx = accountColIdx - 1;
    var accountToIdx = accountToColIdx ? accountToColIdx - 1 : -1;
    var amountIdx = amountColIdx - 1;
    var currencyIdx = currencyColIdx - 1;
    var categoryIdx = categoryColIdx ? categoryColIdx - 1 : -1;
    var subcategoryIdx = subcategoryColIdx ? subcategoryColIdx - 1 : -1;
    var merchantIdx = merchantColIdx ? merchantColIdx - 1 : -1;
    var descriptionIdx = descriptionColIdx ? descriptionColIdx - 1 : -1;
    var activeIdx = activeColIdx ? activeColIdx - 1 : -1;
    var lastCreatedIdx = lastCreatedColIdx ? lastCreatedColIdx - 1 : -1;
    
    var stats = { created: 0, skipped: 0, errors: 0 };
    var transactionsToAdd = [];
    var lastCreatedUpdates = [];
    
    // Process each recurring transaction
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Check array bounds
      if (row.length <= nameIdx || row.length <= typeIdx || row.length <= frequencyIdx || 
          row.length <= startDateIdx || row.length <= accountIdx || row.length <= amountIdx || 
          row.length <= currencyIdx) {
        stats.skipped++;
        continue;
      }
      
      // Build recurring transaction object
      var recurringTx = {
        Name: String(row[nameIdx] || '').trim(),
        Type: String(row[typeIdx] || '').trim(),
        Frequency: String(row[frequencyIdx] || '').trim(),
        DayOfMonth: dayOfMonthIdx >= 0 ? (row[dayOfMonthIdx] ? Number(row[dayOfMonthIdx]) : null) : null,
        DayOfWeek: dayOfWeekIdx >= 0 ? (row[dayOfWeekIdx] ? Number(row[dayOfWeekIdx]) : null) : null,
        StartDate: row[startDateIdx] instanceof Date ? row[startDateIdx] : null,
        EndDate: endDateIdx >= 0 && row[endDateIdx] instanceof Date ? row[endDateIdx] : null,
        Account: String(row[accountIdx] || '').trim(),
        AccountTo: accountToIdx >= 0 ? String(row[accountToIdx] || '').trim() : '',
        Amount: Number(row[amountIdx]) || 0,
        Currency: String(row[currencyIdx] || '').trim(),
        Category: categoryIdx >= 0 ? String(row[categoryIdx] || '').trim() : '',
        Subcategory: subcategoryIdx >= 0 ? String(row[subcategoryIdx] || '').trim() : '',
        Merchant: merchantIdx >= 0 ? String(row[merchantIdx] || '').trim() : '',
        Description: descriptionIdx >= 0 ? String(row[descriptionIdx] || '').trim() : '',
        Active: activeIdx >= 0 ? row[activeIdx] : true,
        LastCreated: lastCreatedIdx >= 0 && row[lastCreatedIdx] instanceof Date ? row[lastCreatedIdx] : null
      };
      
      // Check if should create transaction
      if (!pfShouldCreateTransaction_(recurringTx, today)) {
        stats.skipped++;
        continue;
      }
      
      // Determine transaction date
      var transactionDate = today;
      if (recurringTx.Frequency === PF_RECURRING_FREQUENCY.MONTHLY || 
          recurringTx.Frequency === PF_RECURRING_FREQUENCY.QUARTERLY || 
          recurringTx.Frequency === PF_RECURRING_FREQUENCY.YEARLY) {
        // Use day of month
        if (recurringTx.DayOfMonth) {
          var lastDayOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
          var day = Math.min(recurringTx.DayOfMonth, lastDayOfMonth);
          transactionDate = new Date(today.getFullYear(), today.getMonth(), day);
        }
      } else if (recurringTx.Frequency === PF_RECURRING_FREQUENCY.WEEKLY) {
        // Use current date (or nearest past day of week)
        transactionDate = today;
      }
      
      // Ensure date is within period
      if (transactionDate < startDate || transactionDate > endDate) {
        stats.skipped++;
        continue;
      }
      
      // Create transaction row
      try {
        var txRow = [];
        for (var j = 0; j < PF_TRANSACTIONS_SCHEMA.columns.length; j++) {
          var col = PF_TRANSACTIONS_SCHEMA.columns[j];
          var value = '';
          
          if (col.key === 'Date') {
            value = transactionDate;
          } else if (col.key === 'Type') {
            value = recurringTx.Type;
          } else if (col.key === 'Account') {
            value = recurringTx.Account;
          } else if (col.key === 'AccountTo') {
            value = recurringTx.AccountTo || '';
          } else if (col.key === 'Amount') {
            value = recurringTx.Amount;
          } else if (col.key === 'Currency') {
            value = recurringTx.Currency;
          } else if (col.key === 'Category') {
            value = recurringTx.Category || '';
          } else if (col.key === 'Subcategory') {
            value = recurringTx.Subcategory || '';
          } else if (col.key === 'Merchant') {
            value = recurringTx.Merchant || '';
          } else if (col.key === 'Description') {
            value = recurringTx.Description || recurringTx.Name;
          } else if (col.key === 'Tags') {
            value = '';
          } else if (col.key === 'Source') {
            value = PF_IMPORT_SOURCE.MANUAL;
          } else if (col.key === 'SourceId') {
            value = '';
          } else if (col.key === 'Status') {
            value = PF_TRANSACTION_STATUS.OK;
          }
          
          txRow.push(value);
        }
        
        transactionsToAdd.push(txRow);
        lastCreatedUpdates.push({ row: i + 2, date: transactionDate });
        stats.created++;
        
      } catch (e) {
        pfLogError_(e, 'pfCreateRecurringTransactions_', PF_LOG_LEVEL.ERROR);
        stats.errors++;
      }
    }
    
    // Batch write transactions
    if (transactionsToAdd.length > 0) {
      var txLastRow = txSheet.getLastRow();
      var startRow = txLastRow > 1 ? txLastRow + 1 : 2;
      txSheet.getRange(startRow, 1, transactionsToAdd.length, PF_TRANSACTIONS_SCHEMA.columns.length)
        .setValues(transactionsToAdd);
      
      // Format date column
      var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
      if (dateCol) {
        txSheet.getRange(startRow, dateCol, transactionsToAdd.length, 1)
          .setNumberFormat('dd.mm.yyyy');
      }
      
      // Format amount column
      var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
      if (amountCol) {
        txSheet.getRange(startRow, amountCol, transactionsToAdd.length, 1)
          .setNumberFormat('0.00');
      }
      
      // Update LastCreated for each recurring transaction
      for (var k = 0; k < lastCreatedUpdates.length; k++) {
        var update = lastCreatedUpdates[k];
        if (lastCreatedColIdx) {
          recurringSheet.getRange(update.row, lastCreatedColIdx).setValue(update.date);
        }
      }
    }
    
    pfLogInfo_('Created ' + stats.created + ' transactions, skipped ' + stats.skipped + ', errors ' + stats.errors, 'pfCreateRecurringTransactions_');
    return pfCreateSuccessResponse_('Создано транзакций: ' + stats.created + ', пропущено: ' + stats.skipped, stats);
    
  } catch (e) {
    pfLogError_(e, 'pfCreateRecurringTransactions_', PF_LOG_LEVEL.ERROR);
    return pfCreateErrorResponse_('Ошибка при создании регулярных транзакций: ' + e.toString(), 'CREATE_ERROR', e);
  }
}

/**
 * Public function to create recurring transactions (called from menu).
 */
function pfCreateRecurringTransactions() {
  try {
    var result = pfCreateRecurringTransactions_(SpreadsheetApp.getActiveSpreadsheet());
    
    if (result.success) {
      SpreadsheetApp.getUi().alert(result.message || 'Регулярные транзакции созданы успешно');
    } else {
      SpreadsheetApp.getUi().alert('Ошибка: ' + (result.message || 'Неизвестная ошибка'));
    }
  } catch (e) {
    var handled = pfHandleError_(e, 'pfCreateRecurringTransactions', 'Ошибка при создании регулярных транзакций');
    SpreadsheetApp.getUi().alert('Ошибка: ' + handled.message);
  }
}
