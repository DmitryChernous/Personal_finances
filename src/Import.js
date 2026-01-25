/**
 * Import framework for bank statements and transaction files.
 *
 * Architecture:
 * - DTO (Data Transfer Object) for normalized transactions
 * - Base importer interface
 * - Staging sheet (Import_Raw) for intermediate data
 * - Preview and Commit workflow
 * - Deduplication by SourceId/hash
 */

/**
 * Internal DTO for a transaction (normalized format).
 * This is the standard format all importers should produce.
 * @typedef {Object} TransactionDTO
 * @property {Date} date - Transaction date
 * @property {string} type - 'expense', 'income', or 'transfer'
 * @property {string} account - Source account name
 * @property {string} [accountTo] - Destination account (for transfers)
 * @property {number} amount - Transaction amount (always positive)
 * @property {string} currency - Currency code (RUB, USD, EUR, etc.)
 * @property {string} [category] - Category name
 * @property {string} [subcategory] - Subcategory name
 * @property {string} [merchant] - Merchant/place name
 * @property {string} [description] - Transaction description
 * @property {string} [tags] - Comma-separated tags
 * @property {string} source - Source identifier (e.g., 'import:csv', 'import:sberbank')
 * @property {string} [sourceId] - Unique ID from source (for deduplication)
 * @property {string} [rawData] - Original raw data (for debugging/review)
 * @property {Array<{field: string, message: string}>} [errors] - Parsing errors
 */

/**
 * Base importer interface.
 * All importers should implement these methods.
 * @interface
 */
var PF_IMPORTER_INTERFACE = {
  /**
   * Detect if this importer can handle the given file/data.
   * @param {Blob|string|Array<Array<*>>} data - File blob, file content, or array of rows
   * @param {string} [fileName] - Optional file name for detection
   * @returns {boolean} True if this importer can handle the data
   */
  detect: function(data, fileName) { return false; },

  /**
   * Parse the data into raw transaction objects.
   * @param {Blob|string|Array<Array<*>>} data - File blob, file content, or array of rows
   * @param {Object} [options] - Parser options (column mapping, etc.)
   * @returns {Array<Object>} Array of raw transaction objects (not yet normalized)
   */
  parse: function(data, options) { return []; },

  /**
   * Normalize raw transaction into DTO format.
   * @param {Object} rawTransaction - Raw transaction from parse()
   * @param {Object} [options] - Normalization options (default account, currency, etc.)
   * @returns {TransactionDTO} Normalized transaction DTO
   */
  normalize: function(rawTransaction, options) { return null; },

  /**
   * Generate deduplication key for a transaction.
   * @param {TransactionDTO} transaction - Normalized transaction
   * @returns {string} Unique key for deduplication
   */
  dedupeKey: function(transaction) { return ''; }
};

/**
 * Unified function to generate deduplication key from transaction data.
 * Works with both TransactionDTO objects and sheet rows.
 * 
 * @param {Object} data - Either TransactionDTO or sheet row array
 * @param {Object} [options] - Options for row-based generation
 * @param {number} [options.sourceCol] - Column index for source (1-based)
 * @param {number} [options.sourceIdCol] - Column index for sourceId (1-based)
 * @param {number} [options.dateCol] - Column index for date (1-based)
 * @param {number} [options.accountCol] - Column index for account (1-based)
 * @param {number} [options.amountCol] - Column index for amount (1-based)
 * @param {number} [options.typeCol] - Column index for type (1-based)
 * @returns {string} Deduplication key
 */
function pfGenerateDedupeKey_(data, options) {
  options = options || {};
  
  // Determine if data is TransactionDTO or row array
  var isRow = Array.isArray(data);
  
  var source, sourceId, date, account, amount, type;
  
  if (isRow) {
    // Extract from row array using column indices
    var row = data;
    source = options.sourceCol ? String(row[options.sourceCol - 1] || '').trim() : '';
    sourceId = options.sourceIdCol ? String(row[options.sourceIdCol - 1] || '').trim() : '';
    date = options.dateCol ? row[options.dateCol - 1] : null;
    account = options.accountCol ? String(row[options.accountCol - 1] || '').trim() : '';
    amount = options.amountCol ? row[options.amountCol - 1] : null;
    type = options.typeCol ? String(row[options.typeCol - 1] || '').trim() : '';
  } else {
    // Extract from TransactionDTO object
    var transaction = data;
    source = String(transaction.source || '').trim();
    sourceId = transaction.sourceId ? String(transaction.sourceId).trim() : '';
    date = transaction.date;
    account = String(transaction.account || '').trim();
    amount = transaction.amount;
    type = String(transaction.type || '').trim();
  }
  
  // Normalize sourceId
  sourceId = sourceId || '';
  source = source || '';
  
  // Use sourceId if available
  if (sourceId && sourceId !== '') {
    return source + ':' + sourceId;
  }
  
  // Fallback: hash of key fields
  var dateStr = pfFormatDateForDedupe_(date);
  
  var keyFields = [
    dateStr,
    account,
    String(amount || ''),
    type
  ].join('|');
  
  var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, keyFields).map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
  
  return (source || 'unknown') + ':' + hash;
}

/**
 * Validates a transaction DTO.
 * @param {TransactionDTO} transaction
 * @returns {Array<{field: string, message: string}>} Array of validation errors
 */
function pfValidateTransactionDTO_(transaction) {
  var errors = [];
  
  if (!transaction.date || !(transaction.date instanceof Date)) {
    errors.push({ field: 'Date', message: 'Дата обязательна и должна быть валидной' });
  }
  
  if (!transaction.type || ![PF_TRANSACTION_TYPE.EXPENSE, PF_TRANSACTION_TYPE.INCOME, PF_TRANSACTION_TYPE.TRANSFER].includes(transaction.type)) {
    errors.push({ field: 'Type', message: 'Тип должен быть expense, income или transfer' });
  }
  
  if (!transaction.account || String(transaction.account).trim() === '') {
    errors.push({ field: 'Account', message: 'Счет обязателен' });
  }
  
  if (transaction.type === 'transfer' && (!transaction.accountTo || String(transaction.accountTo).trim() === '')) {
    errors.push({ field: 'AccountTo', message: 'Для перевода обязателен счет получателя' });
  }
  
  if (!transaction.amount || transaction.amount <= 0 || isNaN(transaction.amount)) {
    errors.push({ field: 'Amount', message: 'Сумма должна быть положительным числом' });
  }
  
  if (!transaction.currency || String(transaction.currency).trim() === '') {
    errors.push({ field: 'Currency', message: 'Валюта обязательна' });
  }
  
  if (!transaction.source || String(transaction.source).trim() === '') {
    errors.push({ field: 'Source', message: 'Источник обязателен' });
  }
  
  return errors;
}

/**
 * Converts a TransactionDTO to a row array matching PF_TRANSACTIONS_SCHEMA.
 * @param {TransactionDTO} transaction
 * @returns {Array<*>} Row array with values in schema order
 */
function pfTransactionDTOToRow_(transaction) {
  var row = [];
  var schema = PF_TRANSACTIONS_SCHEMA;
  
  // Map schema keys to transaction keys (case-insensitive)
  var keyMap = {
    'Date': 'date',
    'Type': 'type',
    'Account': 'account',
    'AccountTo': 'accountTo',
    'Amount': 'amount',
    'Currency': 'currency',
    'Category': 'category',
    'Subcategory': 'subcategory',
    'Merchant': 'merchant',
    'Description': 'description',
    'Tags': 'tags',
    'Source': 'source',
    'SourceId': 'sourceId',
    'Status': 'status'
  };
  
  for (var i = 0; i < schema.columns.length; i++) {
    var col = schema.columns[i];
    var transactionKey = keyMap[col.key] || col.key.toLowerCase();
    var value = transaction[transactionKey];
    
    // Handle special cases
    if (col.key === 'Date') {
      // Date might be ISO string (from server) or Date object
      var dateObj = pfISOStringToDate_(value);
      if (dateObj) {
        row.push(dateObj);
      } else if (value instanceof Date) {
        // Already a Date object
        row.push(value);
      } else {
        // Invalid or empty date
        if (value && typeof value === 'string' && value.length > 0) {
          pfLogWarning_('Invalid date string: ' + value, 'pfTransactionDTOToRow_');
        }
        row.push('');
      }
    } else if (col.key === 'Amount' && typeof value === 'number') {
      row.push(value);
    } else if (value === null || value === undefined) {
      row.push('');
    } else {
      row.push(String(value));
    }
  }
  
  return row;
}

/**
 * Ensures Import_Raw staging sheet exists.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function pfEnsureImportRawSheet_(ss) {
  var sheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.IMPORT_RAW);
  
  // Set headers if empty
  if (sheet.getLastRow() === 0) {
    var headers = [];
    for (var i = 0; i < PF_TRANSACTIONS_SCHEMA.columns.length; i++) {
      headers.push(pfT_('columns.' + PF_TRANSACTIONS_SCHEMA.columns[i].key));
    }
    // Add extra columns for import metadata
    headers.push('Ошибки парсинга');
    headers.push('Ключ дедупликации');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * @deprecated This function is no longer used.
 * Processing logic has been moved to pfProcessDataBatch() which handles batching and progress updates.
 * Kept for reference only - can be removed in future cleanup.
 */
function pfProcessImportData_(rawData, importer, options) {
  // This function is deprecated and not used anywhere in the codebase.
  // All processing is now done via pfProcessDataBatch().
  throw new Error('pfProcessImportData_ is deprecated. Use pfProcessDataBatch() instead.');
}

/**
 * Get existing transaction deduplication keys.
 * Private function - loads all keys from Transactions sheet.
 * @returns {Object} Map of dedupeKey -> true
 */
function pfGetExistingTransactionKeys_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    pfLogWarning_('Cannot get spreadsheet in pfGetExistingTransactionKeys_', 'pfGetExistingTransactionKeys_');
    return {};
  }
  
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet) {
    return {}; // Sheet doesn't exist yet, no existing keys
  }
  
  // Cache lastRow to avoid multiple calls
  var lastRow = txSheet.getLastRow();
  if (lastRow <= 1) {
    return {}; // Only header or empty sheet
  }
  
  var keys = {};
  var sourceCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Source');
  var sourceIdCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'SourceId');
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var accountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Account');
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  var typeCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
  
  if (!sourceCol || !dateCol || !accountCol || !amountCol || !typeCol) {
    pfLogWarning_('Missing required columns in Transactions schema', 'pfGetExistingTransactionKeys_');
    return {};
  }
  
  // Optimize: only load columns we need instead of all columns
  // This significantly reduces memory and processing time for large sheets
  var numRows = lastRow - 1;
  if (numRows > 10000) {
    // For very large sheets, process in chunks to avoid timeout
    var chunkSize = 5000;
    for (var chunkStart = 0; chunkStart < numRows; chunkStart += chunkSize) {
      var chunkRows = Math.min(chunkSize, numRows - chunkStart);
      var chunkData = txSheet.getRange(2 + chunkStart, 1, chunkRows, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
      
      for (var i = 0; i < chunkData.length; i++) {
        var row = chunkData[i];
        var source = String(row[sourceCol - 1] || '').trim();
        
        // Only process rows with source
        if (!source) {
          continue;
        }
        
        // Use unified function to generate key
        var dedupeKey = pfGenerateDedupeKey_(row, {
          sourceCol: sourceCol,
          sourceIdCol: sourceIdCol,
          dateCol: dateCol,
          accountCol: accountCol,
          amountCol: amountCol,
          typeCol: typeCol
        });
        
        keys[dedupeKey] = true;
      }
    }
  } else {
    // For smaller sheets, load all at once (faster)
    var data = txSheet.getRange(2, 1, numRows, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var source = String(row[sourceCol - 1] || '').trim();
      
      // Only process rows with source
      if (!source) {
        continue;
      }
      
      // Use unified function to generate key
      var dedupeKey = pfGenerateDedupeKey_(row, {
        sourceCol: sourceCol,
        sourceIdCol: sourceIdCol,
        dateCol: dateCol,
        accountCol: accountCol,
        amountCol: amountCol,
        typeCol: typeCol
      });
      
      keys[dedupeKey] = true;
    }
  }
  
  return keys;
}

/**
 * Preview import: show what will be imported without committing.
 * @param {Array<TransactionDTO>} transactions
 * @returns {Object} Preview data for UI
 */
function pfPreviewImport_(transactions) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stagingSheet = pfEnsureImportRawSheet_(ss);
  
  // Clear existing staging data safely (content, formatting, and notes)
  var lastRow = stagingSheet.getLastRow();
  if (lastRow > 1) {
    var rowsToDelete = lastRow - 1; // Exclude header row
    if (rowsToDelete > 0) {
      pfClearSheetRows_(stagingSheet, 2, rowsToDelete);
    }
  }
  
  // Write transactions to staging sheet
  var rows = [];
  for (var i = 0; i < transactions.length; i++) {
    var tx = transactions[i];
    var row = pfTransactionDTOToRow_(tx);
    
    // Add error column
    var errorText = '';
    if (tx.errors && tx.errors.length > 0) {
      errorText = tx.errors.map(function(e) { return e.field + ': ' + e.message; }).join('; ');
    }
    row.push(errorText);
    
    // Add dedupe key
    row.push(pfGenerateDedupeKey_(tx));
    
    rows.push(row);
  }
  
  if (rows.length > 0) {
    try {
      stagingSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    } catch (e) {
      pfLogError_(e, 'pfPreviewImport_', PF_LOG_LEVEL.ERROR);
      throw e;
    }
    
    // Format error column
    var errorCol = PF_TRANSACTIONS_SCHEMA.columns.length + 1;
    var errorRange = stagingSheet.getRange(2, errorCol, rows.length, 1);
    errorRange.setFontColor('#CC0000');
    
    // Highlight rows with errors
    for (var i = 0; i < rows.length; i++) {
      if (transactions[i].errors && transactions[i].errors.length > 0) {
        stagingSheet.getRange(i + 2, 1, 1, rows[0].length).setBackground('#FFE6E6');
      } else if (transactions[i].status === PF_TRANSACTION_STATUS.DUPLICATE) {
        stagingSheet.getRange(i + 2, 1, 1, rows[0].length).setBackground('#FFFFE6');
        // Add note with dedupe key for easy lookup
        var dedupeKeyCol = rows[0].length; // Last column
        var dedupeKey = pfGenerateDedupeKey_(transactions[i]);
        var note = 'Дубликат. Ключ: ' + dedupeKey + '\nИспользуйте меню "Personal finances → Найти дубликат" для поиска оригинала.';
        stagingSheet.getRange(i + 2, dedupeKeyCol).setNote(note);
      }
    }
  }
  
  // Calculate stats
  var stats = {
    total: transactions.length,
    valid: 0,
    needsReview: 0,
    duplicates: 0
  };
  
  for (var i = 0; i < transactions.length; i++) {
    var tx = transactions[i];
    if (tx.status === PF_TRANSACTION_STATUS.OK) stats.valid++;
    else if (tx.status === PF_TRANSACTION_STATUS.NEEDS_REVIEW) stats.needsReview++;
    else if (tx.status === PF_TRANSACTION_STATUS.DUPLICATE) stats.duplicates++;
  }
  
  return {
    stats: stats,
    stagingSheetName: stagingSheet.getName()
  };
}

/**
 * Find existing transaction by deduplication key.
 * Public function for HTML Service - helps user find duplicate transactions.
 * @param {string} dedupeKey - Deduplication key (e.g., 'import:sberbank:160620251146306495')
 * @returns {Object} {found: boolean, rowNum: number, transaction: Object}
 */
function pfFindDuplicateTransaction(dedupeKey) {
  // Валидация входных данных
  if (!dedupeKey || typeof dedupeKey !== 'string' || dedupeKey.trim().length === 0) {
    return { found: false, message: 'Неверный ключ дедупликации: должен быть непустой строкой' };
  }
  
  dedupeKey = dedupeKey.trim();
  
  // Валидация формата ключа (должен содержать хотя бы одно двоеточие)
  if (dedupeKey.indexOf(':') === -1) {
    return { found: false, message: 'Неверный формат ключа дедупликации: должен содержать ":"' };
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet || txSheet.getLastRow() <= 1) {
    Logger.log('[SERVER] Transactions sheet is empty');
    return { found: false, message: 'Лист транзакций пуст' };
  }
  
  var sourceCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Source');
  var sourceIdCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'SourceId');
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var accountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Account');
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  var typeCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
  
  if (!sourceCol || !dateCol || !accountCol || !amountCol || !typeCol) {
    Logger.log('[SERVER] Missing required columns');
    return { found: false, message: 'Ошибка: не найдены необходимые колонки' };
  }
  
  var data = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
  Logger.log('[SERVER] Checking ' + data.length + ' existing transactions');
  
  // Parse dedupeKey to extract source and sourceId or hash
  var parts = dedupeKey.split(':');
  if (parts.length < 2) {
    Logger.log('[SERVER] Invalid key format');
    return { found: false, message: 'Неверный формат ключа дедупликации' };
  }
  
  var source = parts[0];
  var keySuffix = parts.slice(1).join(':'); // Everything after first ':'
  
  var checkedCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowSource = String(row[sourceCol - 1] || '').trim();
    
    // Only check rows with matching source
    if (rowSource !== source) {
      continue;
    }
    
    checkedCount++;
    var rowSourceId = sourceIdCol ? String(row[sourceIdCol - 1] || '').trim() : '';
    
    // Use unified function to generate key for this row
    var rowKey = pfGenerateDedupeKey_(row, {
      sourceCol: sourceCol,
      sourceIdCol: sourceIdCol,
      dateCol: dateCol,
      accountCol: accountCol,
      amountCol: amountCol,
      typeCol: typeCol
    });
    
    if (rowKey === dedupeKey) {
      // Found it!
      var transaction = {
        date: row[dateCol - 1],
        type: row[typeCol - 1],
        account: row[accountCol - 1],
        amount: row[amountCol - 1],
        currency: pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Currency') ? row[pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Currency') - 1] : '',
        description: pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Description') ? row[pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Description') - 1] : '',
        source: rowSource,
        sourceId: rowSourceId
      };
      
      return {
        found: true,
        rowNum: i + 2, // 1-based row number (header is row 1)
        transaction: transaction,
        message: 'Найдена дублирующая транзакция в строке ' + (i + 2) + ' листа "Транзакции"'
      };
    }
  }
  
  return { 
    found: false, 
    message: 'Дублирующая транзакция не найдена в листе "Транзакции". Проверено ' + checkedCount + ' транзакций с источником "' + source + '". Возможно, это ложное срабатывание дедупликации - транзакция будет добавлена при импорте.' 
  };
}

/**
 * Public function for HTML Service: Commit import.
 * @param {boolean} includeNeedsReview - Include transactions marked for review
 * @returns {Object} Commit result
 */
function pfCommitImport(includeNeedsReview) {
  // Валидация входных данных
  if (includeNeedsReview !== undefined && typeof includeNeedsReview !== 'boolean') {
    pfLogWarning_('includeNeedsReview should be boolean, got: ' + typeof includeNeedsReview, 'pfCommitImport');
    includeNeedsReview = Boolean(includeNeedsReview);
  }
  
  return pfCommitImport_(includeNeedsReview);
}

/**
 * Commit import: move valid transactions from staging to Transactions sheet.
 * @param {boolean} includeNeedsReview - Include transactions marked for review
 * @returns {Object} Commit result
 */
function pfCommitImport_(includeNeedsReview) {
  includeNeedsReview = includeNeedsReview || false;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return pfCreateErrorResponse_('Не удалось получить доступ к таблице');
    }
    
    var stagingSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.IMPORT_RAW);
    if (!stagingSheet) {
      return pfCreateErrorResponse_('Лист предпросмотра не найден. Сначала выполните предпросмотр импорта.');
    }
    
    // Cache lastRow to avoid multiple calls
    var stagingLastRow = stagingSheet.getLastRow();
    if (stagingLastRow <= 1) {
      return pfCreateErrorResponse_('Нет данных для импорта. Лист предпросмотра пуст.');
    }
    
    var txSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
    var numDataCols = PF_TRANSACTIONS_SCHEMA.columns.length;
    // Use cached stagingLastRow
    var data = stagingSheet.getRange(2, 1, stagingLastRow - 1, numDataCols).getValues();
    
    var statusColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
    var dateColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
    var errorCol = numDataCols + 1; // Error column is after transaction columns
    
    var rowsToAdd = [];
    var stats = {
      added: 0,
      skipped: 0,
      needsReview: 0
    };
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Check status - normalize to string
      var statusValue = statusColIdx ? String(row[statusColIdx - 1] || '').trim() : PF_TRANSACTION_STATUS.OK;
      var hasErrors = false;
      
      try {
        var errorCellValue = stagingSheet.getRange(i + 2, errorCol).getValue();
        hasErrors = errorCellValue && String(errorCellValue).trim() !== '';
      } catch (e) {
        pfLogWarning_('Could not read error cell for row ' + (i + 2) + ': ' + e.toString(), 'pfCommitImport_');
      }
      
      // Skip duplicates
      if (statusValue === PF_TRANSACTION_STATUS.DUPLICATE) {
        stats.skipped++;
        continue;
      }
      
      // Skip needs_review if not including
      if (statusValue === PF_TRANSACTION_STATUS.NEEDS_REVIEW && !includeNeedsReview) {
        stats.needsReview++;
        continue;
      }
      
      // Convert Date strings back to Date objects if needed
      var processedRow = [];
      for (var j = 0; j < row.length; j++) {
        var value = row[j];
        
        // If this is the Date column and value is a string (ISO format), convert to Date
        if (dateColIdx && j === dateColIdx - 1) {
          var dateObj = pfISOStringToDate_(value);
          if (dateObj) {
            processedRow.push(dateObj);
          } else if (value instanceof Date) {
            // Already a Date object
            processedRow.push(value);
          } else {
            // Keep as-is if can't parse (might be empty or invalid)
            if (value && typeof value === 'string' && value.length > 0) {
              pfLogWarning_('Invalid date string in row ' + (i + 2) + ': ' + value, 'pfCommitImport_');
            }
            processedRow.push(value);
          }
        } else {
          processedRow.push(value);
        }
      }
      
      // Add transaction
      rowsToAdd.push(processedRow);
      stats.added++;
    }
    
    Logger.log('[SERVER] Prepared ' + rowsToAdd.length + ' rows to add. Stats: ' + JSON.stringify(stats));
    
    if (rowsToAdd.length > 0) {
      var lastRow = txSheet.getLastRow();
      Logger.log('[SERVER] Current Transactions sheet lastRow: ' + lastRow);
      Logger.log('[SERVER] Writing ' + rowsToAdd.length + ' rows to Transactions sheet starting at row ' + (lastRow + 1));
      Logger.log('[SERVER] Row width: ' + rowsToAdd[0].length + ' columns');
      
      if (rowsToAdd[0].length !== numDataCols) {
        pfLogWarning_('Row width mismatch! Expected ' + numDataCols + ', got ' + rowsToAdd[0].length, 'pfCommitImport_');
      }
      
      try {
        var targetRange = txSheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length);
        Logger.log('[SERVER] Target range: ' + targetRange.getA1Notation());
        targetRange.setValues(rowsToAdd);
        Logger.log('[SERVER] Successfully wrote ' + rowsToAdd.length + ' rows to Transactions sheet');
        
        // Skip normalization/validation for large imports to avoid timeout
        // Data is already normalized during parsing in pfProcessDataBatch
        // Validation will happen automatically via onEdit trigger when user edits rows
        // For small imports (< 50 rows), we can still do it
        if (rowsToAdd.length < 50) {
          Logger.log('[SERVER] Small import (' + rowsToAdd.length + ' rows), running normalization/validation...');
          for (var i = 0; i < rowsToAdd.length; i++) {
            try {
              pfNormalizeTransactionRow_(txSheet, lastRow + 1 + i);
              var errors = pfValidateTransactionRow_(txSheet, lastRow + 1 + i);
              pfHighlightErrors_(txSheet, lastRow + 1 + i, errors);
            } catch (e) {
              Logger.log('[SERVER] WARNING: Error normalizing/validating row ' + (lastRow + 1 + i) + ': ' + e.toString());
            }
          }
          Logger.log('[SERVER] Completed normalization and validation');
        } else {
          Logger.log('[SERVER] Large import (' + rowsToAdd.length + ' rows), skipping normalization/validation to avoid timeout');
          Logger.log('[SERVER] Data is already normalized from parsing. Validation will occur via onEdit trigger when rows are edited.');
        }
      } catch (e) {
        pfLogError_(e, 'pfCommitImport_', PF_LOG_LEVEL.ERROR);
        pfLogWarning_('Rows to add length: ' + rowsToAdd.length, 'pfCommitImport_');
        throw new Error('Ошибка при записи транзакций: ' + (e.message || e.toString()));
      }
    }
    
    // Clear staging sheet safely (content, formatting, and notes)
    // Use cached stagingLastRow
    if (stagingLastRow > 1) {
      var rowsToDelete = stagingLastRow - 1;
      if (rowsToDelete > 0) {
        pfClearSheetRows_(stagingSheet, 2, rowsToDelete);
      }
    }
    
    return pfCreateSuccessResponse_(
      'Импортировано: ' + stats.added + ' транзакций. Пропущено: ' + stats.skipped + '. На проверку: ' + stats.needsReview,
      { stats: stats }
    );
    
  } catch (e) {
    return pfHandleError_(e, 'pfCommitImport_', 'Ошибка при импорте');
  }
}

/**
 * Step 1: Detect file format and return importer info.
 * Public function for HTML Service.
 * @param {string} fileContent - File content as string
 * @returns {Object} {importerType: string, detected: boolean}
 */
function pfDetectFileFormat(fileContent) {
  // Валидация входных данных
  if (!fileContent || typeof fileContent !== 'string' || fileContent.trim().length === 0) {
    throw new Error('fileContent must be a non-empty string');
  }
  
  if (fileContent.length > PF_IMPORT_MAX_FILE_SIZE) {
    throw new Error('File too large: ' + Math.round(fileContent.length / 1024 / 1024) + 'MB. Maximum is ' + Math.round(PF_IMPORT_MAX_FILE_SIZE / 1024 / 1024) + 'MB.');
  }
  
  // Check for PDF first (by file extension or MIME type)
  // Note: For PDF, we can't detect from content string alone, so we check fileName if available
  // In UI, we'll pass fileName separately for PDF detection
  
  // Check for PDF (if fileName is provided)
  if (typeof PF_PDF_IMPORTER !== 'undefined') {
    // For PDF, detection is based on file name or MIME type, not content
    // We'll handle PDF detection in the UI layer
  }
  
  // Ограничение размера для проверки (первые 100KB достаточно для определения формата)
  var checkContent = fileContent.length > 100000 ? fileContent.substring(0, 100000) : fileContent;
  
  if (typeof PF_SBERBANK_IMPORTER !== 'undefined' && PF_SBERBANK_IMPORTER.detect(checkContent)) {
    return { importerType: 'sberbank', detected: true };
  } else if (PF_CSV_IMPORTER.detect(checkContent)) {
    return { importerType: 'csv', detected: true };
  }
  return { importerType: null, detected: false };
}

/**
 * Step 2a: Parse PDF file (special handling for PDF).
 * Public function for HTML Service.
 * @param {string} fileContent - PDF file content as base64 string
 * @param {string} fileName - File name (for detection)
 * @param {Object} options - Parse options
 * @returns {Object} {rawData: Array, count: number, errors: Array}
 */
function pfParsePdfFile(fileContent, fileName, options) {
  options = options || {};
  
  if (!fileContent || typeof fileContent !== 'string') {
    throw new Error('fileContent must be a non-empty string (base64 encoded PDF)');
  }
  
  if (typeof PF_PDF_IMPORTER === 'undefined') {
    throw new Error('PDF importer is not available');
  }
  
  // Check if it's a PDF
  if (!PF_PDF_IMPORTER.detect(null, fileName)) {
    throw new Error('File is not a PDF: ' + fileName);
  }
  
  // Convert base64 to Blob
  try {
    var bytes = Utilities.base64Decode(fileContent);
    var blob = Utilities.newBlob(bytes, 'application/pdf', fileName);
    
    // Parse PDF
    var rawData = PF_PDF_IMPORTER.parse(blob, options);
    
    if (!Array.isArray(rawData)) {
      throw new Error('PDF parser returned invalid data: expected array, got ' + typeof rawData);
    }
    
    return {
      rawData: rawData,
      count: rawData.length,
      errors: []
    };
  } catch (parseError) {
    throw new Error('Error parsing PDF: ' + (parseError.message || parseError.toString()));
  }
}

/**
 * Step 2: Parse file content.
 * Public function for HTML Service.
 * @param {string} fileContent - File content as string
 * @param {string} importerType - 'sberbank', 'csv', or 'pdf'
 * @param {Object} options - Parse options
 * @returns {Object} {rawData: Array, count: number, errors: Array}
 */
function pfParseFileContent(fileContent, importerType, options) {
  // Валидация входных данных
  if (!fileContent || typeof fileContent !== 'string' || fileContent.trim().length === 0) {
    throw new Error('fileContent must be a non-empty string');
  }
  
  if (!importerType || typeof importerType !== 'string' || !['sberbank', 'csv', 'pdf'].includes(importerType)) {
    throw new Error('Invalid importerType: must be "sberbank", "csv", or "pdf", got: ' + String(importerType));
  }
  
  if (fileContent.length > PF_IMPORT_MAX_FILE_SIZE) {
    throw new Error('File too large: ' + Math.round(fileContent.length / 1024 / 1024) + 'MB. Maximum is ' + Math.round(PF_IMPORT_MAX_FILE_SIZE / 1024 / 1024) + 'MB.');
  }
  
  options = options || {};
  
  var importer = null;
  
  if (importerType === 'sberbank') {
    importer = PF_SBERBANK_IMPORTER;
  } else if (importerType === 'csv') {
    importer = PF_CSV_IMPORTER;
  } else if (importerType === 'pdf') {
    if (typeof PF_PDF_IMPORTER === 'undefined') {
      throw new Error('PDF importer is not available');
    }
    importer = PF_PDF_IMPORTER;
  } else {
    throw new Error('Неизвестный тип импортера: ' + importerType);
  }
  
  var rawData = [];
  var errors = [];
  
  try {
    rawData = importer.parse(fileContent, options);
  } catch (parseError) {
    errors.push('Ошибка парсинга: ' + (parseError.message || parseError.toString()));
    throw new Error('Ошибка при парсинге файла: ' + (parseError.message || parseError.toString()));
  }
  
  // Валидация результата
  if (!Array.isArray(rawData)) {
    throw new Error('Parser returned invalid data: expected array, got ' + typeof rawData);
  }
  
  return {
    rawData: rawData,
    count: rawData.length,
    errors: errors
  };
}

/**
 * Step 3: Process raw data in batches with progress tracking.
 * Public function for HTML Service.
 * Note: rawData is passed as JSON string to avoid size limits.
 * @param {string} rawDataJson - Raw data as JSON string
 * @param {string} importerType - 'sberbank', 'csv', or 'pdf'
 * @param {Object} options - Processing options
 * @param {number} batchSize - Number of items to process per batch (default: 100)
 * @param {number} startIndex - Start index for this batch (default: 0)
 * @returns {Object} {transactions: Array, stats: Object, processed: number, total: number, hasMore: boolean}
 */
function pfProcessDataBatch(rawDataJson, importerType, options, batchSize, startIndex) {
  try {
    // Валидация входных данных
    if (!rawDataJson || typeof rawDataJson !== 'string' || rawDataJson.trim().length === 0) {
      pfLogError_('rawDataJson validation failed', 'pfProcessDataBatch', PF_LOG_LEVEL.ERROR);
      throw new Error('rawDataJson must be a non-empty string');
    }
    
    if (!importerType || typeof importerType !== 'string' || !['sberbank', 'csv', 'pdf'].includes(importerType)) {
      pfLogError_('Invalid importerType: ' + importerType, 'pfProcessDataBatch', PF_LOG_LEVEL.ERROR);
      throw new Error('Invalid importerType: must be "sberbank", "csv", or "pdf", got: ' + String(importerType));
    }
    
    // Use smaller batches for PDF to avoid timeout (PDF parsing is more complex)
    batchSize = batchSize || (importerType === 'pdf' ? 50 : PF_IMPORT_BATCH_SIZE);
    if (typeof batchSize !== 'number' || isNaN(batchSize) || batchSize < 1 || batchSize > 1000) {
      pfLogError_('Invalid batchSize: ' + batchSize, 'pfProcessDataBatch', PF_LOG_LEVEL.ERROR);
      throw new Error('batchSize must be a number between 1 and 1000, got: ' + String(batchSize));
    }
    
    startIndex = startIndex || 0;
    if (typeof startIndex !== 'number' || isNaN(startIndex) || startIndex < 0) {
      pfLogWarning_('Invalid startIndex: ' + startIndex + ', using 0', 'pfProcessDataBatch');
      startIndex = 0;
    }
    
    options = options || {};
    if (typeof options !== 'object') {
      pfLogWarning_('options should be an object, got: ' + typeof options + ', using {}', 'pfProcessDataBatch');
      options = {};
    }
    
    // Get spreadsheet for auto-categorization
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Parse JSON string back to array
    // Note: rawDataJson now contains only the batch data, not the entire array
    var rawData = null;
    try {
      Logger.log('[SERVER] Parsing JSON...');
      rawData = JSON.parse(rawDataJson);
      Logger.log('[SERVER] JSON parsed successfully, array length: ' + rawData.length);
    } catch (e) {
      pfLogError_(e, 'pfProcessDataBatch', PF_LOG_LEVEL.ERROR);
      throw new Error('Invalid JSON in rawDataJson: ' + e.toString());
    }
    
    if (!Array.isArray(rawData)) {
      pfLogError_('rawData is not an array', 'pfProcessDataBatch', PF_LOG_LEVEL.ERROR);
      throw new Error('rawData must be an array');
    }
    
    // Get importer
    Logger.log('[SERVER] Getting importer...');
    var importer = null;
    if (importerType === 'sberbank') {
      importer = PF_SBERBANK_IMPORTER;
      Logger.log('[SERVER] Using Sberbank importer');
    } else if (importerType === 'csv') {
      importer = PF_CSV_IMPORTER;
      Logger.log('[SERVER] Using CSV importer');
    } else if (importerType === 'pdf') {
      if (typeof PF_PDF_IMPORTER === 'undefined') {
        throw new Error('PDF importer is not available');
      }
      importer = PF_PDF_IMPORTER;
      Logger.log('[SERVER] Using PDF importer');
    } else {
      pfLogError_('Unknown importer type', 'pfProcessDataBatch', PF_LOG_LEVEL.ERROR);
      throw new Error('Неизвестный тип импортера: ' + importerType);
    }
    
    if (!importer) {
      pfLogError_('Importer is null/undefined', 'pfProcessDataBatch', PF_LOG_LEVEL.ERROR);
      throw new Error('Importer not found');
    }
    
    var sourceName = importerType === 'sberbank' ? PF_IMPORT_SOURCE.SBERBANK : 
                     importerType === 'pdf' ? 'import:pdf' : PF_IMPORT_SOURCE.CSV;
    var transactions = [];
    var stats = {
      valid: 0,
      needsReview: 0,
      duplicates: 0,
      errors: 0
    };
    
    // Get existing keys from options (always passed from client after first load)
    // If not provided, start with empty object (no existing transactions to check)
    Logger.log('[SERVER] Getting existing keys...');
    var existingKeys = null;
    if (options._existingKeys && typeof options._existingKeys === 'object') {
      existingKeys = options._existingKeys;
      Logger.log('[SERVER] Using keys from options, count: ' + Object.keys(existingKeys).length);
    } else {
      // No keys provided - start with empty (will accumulate during processing)
      existingKeys = {};
      Logger.log('[SERVER] Starting with empty keys object');
    }
    
    Logger.log('[SERVER] Starting to process ' + rawData.length + ' items...');
    var processStartTime = new Date().getTime();
    
    // Process all items in the batch (rawData is already the batch)
    for (var i = 0; i < rawData.length; i++) {
      try {
        if (i % 50 === 0) {
          Logger.log('[SERVER] Processing item ' + i + ' of ' + rawData.length);
        }
        
        var transaction = importer.normalize(rawData[i], options);
        
        // Apply auto-categorization if category is not set
        if (!transaction.category || String(transaction.category).trim() === '') {
          transaction = pfApplyCategoryRules_(transaction, ss);
        }
        
        var dedupeKey = importer.dedupeKey(transaction);
        
        if (existingKeys[dedupeKey]) {
          transaction.status = PF_TRANSACTION_STATUS.DUPLICATE;
          stats.duplicates++;
        } else {
          existingKeys[dedupeKey] = true;
        }
        
        if (transaction.errors && transaction.errors.length > 0) {
          transaction.status = PF_TRANSACTION_STATUS.NEEDS_REVIEW;
          stats.needsReview++;
          stats.errors++;
        } else if (transaction.status === PF_TRANSACTION_STATUS.OK) {
          stats.valid++;
        }
        
        // Convert Date objects to ISO strings for JSON serialization
        // This prevents issues when passing through google.script.run
        transaction.date = pfDateToISOString_(transaction.date);
        
        transactions.push(transaction);
      } catch (e) {
        pfLogError_(e, 'pfProcessDataBatch', PF_LOG_LEVEL.ERROR);
        stats.errors++;
        transactions.push({
          date: pfDateToISOString_(new Date()), // Use ISO string instead of Date object
          type: PF_TRANSACTION_TYPE.EXPENSE,
          account: '',
          amount: 0,
          currency: PF_DEFAULT_CURRENCY,
          source: sourceName,
          status: PF_TRANSACTION_STATUS.NEEDS_REVIEW,
          errors: [{ field: 'General', message: 'Ошибка парсинга (строка ' + (i + 1) + '): ' + e.toString() }],
          rawData: JSON.stringify(rawData[i])
        });
      }
    }
    
    // Calculate processed count (startIndex + batch length)
    var processed = (options._startIndex || startIndex) + rawData.length;
    var totalCount = options._totalCount || rawData.length;
    
    var result = {
      transactions: transactions,
      stats: stats,
      processed: processed,
      total: totalCount,
      hasMore: processed < totalCount,
      existingKeys: existingKeys // Return updated keys for next batch
    };
    
    // Try to serialize to check if there are any issues
    try {
      JSON.stringify(result);
    } catch (e) {
      pfLogError_(e, 'pfProcessDataBatch', PF_LOG_LEVEL.ERROR);
      throw new Error('Failed to serialize result: ' + e.toString());
    }
    
    return result;
    
  } catch (e) {
    pfLogError_(e, 'pfProcessDataBatch', PF_LOG_LEVEL.ERROR);
    throw e;
  }
}

/**
 * Step 4: Write preview to staging sheet.
 * Public function for HTML Service.
 * Note: transactions is passed as JSON string to avoid size limits.
 * @param {string} transactionsJson - Transactions as JSON string
 * @returns {Object} Preview result
 */
function pfWritePreview(transactionsJson) {
  // Валидация входных данных
  if (!transactionsJson || typeof transactionsJson !== 'string' || transactionsJson.trim().length === 0) {
    throw new Error('transactionsJson must be a non-empty string');
  }
  
  // Parse JSON string back to array
  var transactions = null;
  try {
    transactions = JSON.parse(transactionsJson);
  } catch (e) {
    pfLogError_(e, 'pfWritePreview', PF_LOG_LEVEL.ERROR);
    throw new Error('Invalid JSON in transactionsJson: ' + e.toString());
  }
  
  if (!Array.isArray(transactions)) {
    throw new Error('transactions must be an array, got: ' + typeof transactions);
  }
  
  // Ограничение размера для безопасности
  if (transactions.length > PF_IMPORT_MAX_TRANSACTIONS) {
    throw new Error('Too many transactions: ' + transactions.length + '. Maximum is ' + PF_IMPORT_MAX_TRANSACTIONS + '.');
  }
  
  if (transactions.length === 0) {
    throw new Error('transactions array is empty');
  }
  
  return pfPreviewImport_(transactions);
}

/**
 * @deprecated This function is no longer used.
 * Import workflow is now handled client-side in ImportUI.html which calls:
 * - pfDetectFileFormat()
 * - pfParseFileContent()
 * - pfProcessDataBatch() (in batches with progress)
 * - pfWritePreview()
 * 
 * This function is kept for reference only - can be removed in future cleanup.
 */
function pfProcessFileImport_(fileContent, options) {
  // This function is deprecated and not used anywhere in the codebase.
  // Import workflow is now handled client-side with granular progress updates.
  throw new Error('pfProcessFileImport_ is deprecated. Import workflow is now handled client-side via ImportUI.html.');
}

/**
 * Get existing transaction deduplication keys.
 * Public function for HTML Service (called once before batch processing).
 * @returns {Object} Map of dedupeKey -> true
 */
function pfGetExistingTransactionKeys() {
  try {
    var keys = pfGetExistingTransactionKeys_();
    Logger.log('[SERVER] pfGetExistingTransactionKeys completed, keys count: ' + Object.keys(keys).length);
    return keys;
  } catch (e) {
    Logger.log('[SERVER] ERROR in pfGetExistingTransactionKeys: ' + e.toString());
    Logger.log('[SERVER] Error stack: ' + (e.stack || 'No stack'));
    throw e;
  }
}

/**
 * Get list of accounts for dropdown.
 * Public function for HTML Service.
 * @returns {Array<string>}
 */
function pfGetAccountsList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var accountsSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
    
    if (!accountsSheet) {
      Logger.log('[SERVER] Accounts sheet not found');
      return [];
    }
    
    if (accountsSheet.getLastRow() <= 1) {
      Logger.log('[SERVER] Accounts sheet is empty (only headers)');
      return [];
    }
    
    Logger.log('[SERVER] Accounts sheet found, lastRow: ' + accountsSheet.getLastRow());
    
    var accounts = [];
    var accountCol = pfColumnIndex_(PF_ACCOUNTS_SCHEMA, 'Account');
    
    Logger.log('[SERVER] Account column index: ' + accountCol);
    
    if (accountCol) {
      var data = accountsSheet.getRange(2, accountCol, accountsSheet.getLastRow() - 1, 1).getValues();
      Logger.log('[SERVER] Accounts data rows: ' + data.length);
      
      for (var i = 0; i < data.length; i++) {
        var account = String(data[i][0] || '').trim();
        if (account) {
          accounts.push(account);
        }
      }
    } else {
      Logger.log('[SERVER] Account column not found in schema');
    }
    
    return accounts;
  } catch (e) {
    pfLogError_(e, 'pfGetAccountsList', PF_LOG_LEVEL.ERROR);
    return [];
  }
}

/**
 * Main entry point for import workflow.
 * Shows UI for file selection and import.
 */
function pfImportTransactions() {
  // This will be called from menu, opens HTML sidebar
  var html = HtmlService.createHtmlOutputFromFile('ImportUI')
    .setTitle('Импорт транзакций')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}
