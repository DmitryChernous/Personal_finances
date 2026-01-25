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
 * Creates a deduplication key from transaction data.
 * Uses: date + account + amount + sourceId (if available) or hash of key fields.
 * @param {TransactionDTO} transaction
 * @returns {string}
 */
function pfGenerateDedupeKey_(transaction) {
  if (transaction.sourceId) {
    return transaction.source + ':' + transaction.sourceId;
  }
  
  // Fallback: hash of key fields
  var keyFields = [
    transaction.date ? Utilities.formatDate(transaction.date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    transaction.account || '',
    String(transaction.amount || ''),
    transaction.type || ''
  ].join('|');
  
  return transaction.source + ':' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, keyFields).map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
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
  
  if (!transaction.type || !['expense', 'income', 'transfer'].includes(transaction.type)) {
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
  
  for (var i = 0; i < schema.columns.length; i++) {
    var col = schema.columns[i];
    var value = transaction[col.key];
    
    // Handle special cases
    if (col.key === 'Date' && value instanceof Date) {
      row.push(value);
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
 * Process imported data: parse, normalize, validate, deduplicate.
 * @param {Array<Object>} rawData - Raw data from importer
 * @param {Object} importer - Importer object (e.g., PF_CSV_IMPORTER)
 * @param {Object} options - Import options
 * @returns {Object} {transactions: Array<TransactionDTO>, stats: Object}
 */
function pfProcessImportData_(rawData, importer, options) {
  options = options || {};
  var transactions = [];
  var stats = {
    total: rawData.length,
    valid: 0,
    needsReview: 0,
    duplicates: 0,
    errors: 0
  };
  
  // Get existing transaction keys for deduplication
  var existingKeys = pfGetExistingTransactionKeys_();
  
  for (var i = 0; i < rawData.length; i++) {
    try {
      // Normalize
      var transaction = importer.normalize(rawData[i], options);
      
      // Check for duplicates
      var dedupeKey = importer.dedupeKey(transaction);
      if (existingKeys[dedupeKey]) {
        transaction.status = 'duplicate';
        stats.duplicates++;
      } else {
        existingKeys[dedupeKey] = true;
      }
      
      // Count by status
      if (transaction.errors && transaction.errors.length > 0) {
        transaction.status = 'needs_review';
        stats.needsReview++;
        stats.errors++;
      } else if (transaction.status === 'ok') {
        stats.valid++;
      }
      
      transactions.push(transaction);
    } catch (e) {
      stats.errors++;
      // Create error transaction
      transactions.push({
        date: new Date(),
        type: 'expense',
        account: '',
        amount: 0,
        currency: 'RUB',
        source: options.source || 'import:error',
        status: 'needs_review',
        errors: [{ field: 'General', message: 'Ошибка парсинга: ' + e.toString() }],
        rawData: JSON.stringify(rawData[i])
      });
    }
  }
  
  return { transactions: transactions, stats: stats };
}

/**
 * Get existing transaction deduplication keys.
 * @returns {Object} Map of dedupeKey -> true
 */
function pfGetExistingTransactionKeys_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet || txSheet.getLastRow() <= 1) return {};
  
  var keys = {};
  var sourceCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Source');
  var sourceIdCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'SourceId');
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var accountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Account');
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  var typeCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Type');
  
  if (!sourceCol || !dateCol || !accountCol || !amountCol || !typeCol) return {};
  
  var data = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var source = row[sourceCol - 1];
    var sourceId = sourceIdCol ? row[sourceIdCol - 1] : null;
    
    if (sourceId) {
      keys[source + ':' + sourceId] = true;
    } else {
      // Generate hash key
      var date = row[dateCol - 1];
      var account = row[accountCol - 1];
      var amount = row[amountCol - 1];
      var type = row[typeCol - 1];
      
      var keyFields = [
        date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        account || '',
        String(amount || ''),
        type || ''
      ].join('|');
      
      var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, keyFields).map(function(b) {
        return ('0' + (b & 0xFF).toString(16)).slice(-2);
      }).join('');
      
      keys[(source || 'unknown') + ':' + hash] = true;
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
  
  // Clear existing staging data
  if (stagingSheet.getLastRow() > 1) {
    stagingSheet.deleteRows(2, stagingSheet.getLastRow() - 1);
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
    stagingSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    
    // Format error column
    var errorCol = PF_TRANSACTIONS_SCHEMA.columns.length + 1;
    var errorRange = stagingSheet.getRange(2, errorCol, rows.length, 1);
    errorRange.setFontColor('#CC0000');
    
    // Highlight rows with errors
    for (var i = 0; i < rows.length; i++) {
      if (transactions[i].errors && transactions[i].errors.length > 0) {
        stagingSheet.getRange(i + 2, 1, 1, rows[0].length).setBackground('#FFE6E6');
      } else if (transactions[i].status === 'duplicate') {
        stagingSheet.getRange(i + 2, 1, 1, rows[0].length).setBackground('#FFFFE6');
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
    if (tx.status === 'ok') stats.valid++;
    else if (tx.status === 'needs_review') stats.needsReview++;
    else if (tx.status === 'duplicate') stats.duplicates++;
  }
  
  return {
    stats: stats,
    stagingSheetName: stagingSheet.getName()
  };
}

/**
 * Commit import: move valid transactions from staging to Transactions sheet.
 * @param {boolean} includeNeedsReview - Include transactions marked for review
 * @returns {Object} Commit result
 */
function pfCommitImport_(includeNeedsReview) {
  includeNeedsReview = includeNeedsReview || false;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stagingSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.IMPORT_RAW);
  if (!stagingSheet || stagingSheet.getLastRow() <= 1) {
    return { success: false, message: 'Нет данных для импорта' };
  }
  
  var txSheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  var numDataCols = PF_TRANSACTIONS_SCHEMA.columns.length;
  var data = stagingSheet.getRange(2, 1, stagingSheet.getLastRow() - 1, numDataCols).getValues();
  var statusColIdx = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Status');
  var errorCol = numDataCols + 1; // Error column is after transaction columns
  
  var rowsToAdd = [];
  var stats = {
    added: 0,
    skipped: 0,
    needsReview: 0
  };
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var statusValue = statusColIdx ? row[statusColIdx - 1] : 'ok';
    var hasErrors = stagingSheet.getRange(i + 2, errorCol).getValue() !== '';
    
    // Skip duplicates
    if (statusValue === 'duplicate') {
      stats.skipped++;
      continue;
    }
    
    // Skip needs_review if not including
    if (statusValue === 'needs_review' && !includeNeedsReview) {
      stats.needsReview++;
      continue;
    }
    
    // Add transaction
    rowsToAdd.push(row);
    stats.added++;
  }
  
  if (rowsToAdd.length > 0) {
    var lastRow = txSheet.getLastRow();
    txSheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    
    // Normalize and validate added rows
    for (var i = 0; i < rowsToAdd.length; i++) {
      pfNormalizeTransactionRow_(txSheet, lastRow + 1 + i);
      var errors = pfValidateTransactionRow_(txSheet, lastRow + 1 + i);
      pfHighlightErrors_(txSheet, lastRow + 1 + i, errors);
    }
  }
  
  // Clear staging sheet
  if (stagingSheet.getLastRow() > 1) {
    stagingSheet.deleteRows(2, stagingSheet.getLastRow() - 1);
  }
  
  return {
    success: true,
    stats: stats,
    message: 'Импортировано: ' + stats.added + ' транзакций. Пропущено: ' + stats.skipped + '. На проверку: ' + stats.needsReview
  };
}

/**
 * Step 1: Detect file format and return importer info.
 * Public function for HTML Service.
 * @param {string} fileContent - File content as string
 * @returns {Object} {importerType: string, detected: boolean}
 */
function pfDetectFileFormat(fileContent) {
  if (typeof PF_SBERBANK_IMPORTER !== 'undefined' && PF_SBERBANK_IMPORTER.detect(fileContent)) {
    return { importerType: 'sberbank', detected: true };
  } else if (PF_CSV_IMPORTER.detect(fileContent)) {
    return { importerType: 'csv', detected: true };
  }
  return { importerType: null, detected: false };
}

/**
 * Step 2: Parse file content.
 * Public function for HTML Service.
 * @param {string} fileContent - File content as string
 * @param {string} importerType - 'sberbank' or 'csv'
 * @param {Object} options - Parse options
 * @returns {Object} {rawData: Array, count: number, errors: Array}
 */
function pfParseFileContent(fileContent, importerType, options) {
  options = options || {};
  var importer = null;
  
  if (importerType === 'sberbank') {
    importer = PF_SBERBANK_IMPORTER;
  } else if (importerType === 'csv') {
    importer = PF_CSV_IMPORTER;
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
 * @param {string} importerType - 'sberbank' or 'csv'
 * @param {Object} options - Processing options
 * @param {number} batchSize - Number of items to process per batch (default: 100)
 * @param {number} startIndex - Start index for this batch (default: 0)
 * @returns {Object} {transactions: Array, stats: Object, processed: number, total: number, hasMore: boolean}
 */
function pfProcessDataBatch(rawDataJson, importerType, options, batchSize, startIndex) {
  // Parse JSON string back to array
  // Note: rawDataJson now contains only the batch data, not the entire array
  var rawData = JSON.parse(rawDataJson);
  batchSize = batchSize || 100;
  startIndex = startIndex || 0; // Should be 0 since we pass only batch data
  options = options || {};
  
  var importer = null;
  if (importerType === 'sberbank') {
    importer = PF_SBERBANK_IMPORTER;
  } else if (importerType === 'csv') {
    importer = PF_CSV_IMPORTER;
  } else {
    throw new Error('Неизвестный тип импортера: ' + importerType);
  }
  
  var sourceName = importerType === 'sberbank' ? 'import:sberbank' : 'import:csv';
  var transactions = [];
  var stats = {
    valid: 0,
    needsReview: 0,
    duplicates: 0,
    errors: 0
  };
  
  // Get existing keys only once (cache it in ScriptProperties for persistence across calls)
  var existingKeys = null;
  var cacheKey = 'pf_import_existing_keys';
  var cachedKeys = PropertiesService.getScriptProperties().getProperty(cacheKey);
  if (cachedKeys) {
    try {
      existingKeys = JSON.parse(cachedKeys);
    } catch (e) {
      existingKeys = pfGetExistingTransactionKeys_();
      PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(existingKeys));
    }
  } else {
    existingKeys = pfGetExistingTransactionKeys_();
    PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(existingKeys));
  }
  
  // Process all items in the batch (rawData is already the batch)
  for (var i = 0; i < rawData.length; i++) {
    try {
      var transaction = importer.normalize(rawData[i], options);
      var dedupeKey = importer.dedupeKey(transaction);
      
      if (existingKeys[dedupeKey]) {
        transaction.status = 'duplicate';
        stats.duplicates++;
      } else {
        existingKeys[dedupeKey] = true;
      }
      
      if (transaction.errors && transaction.errors.length > 0) {
        transaction.status = 'needs_review';
        stats.needsReview++;
        stats.errors++;
      } else if (transaction.status === 'ok') {
        stats.valid++;
      }
      
      transactions.push(transaction);
    } catch (e) {
      stats.errors++;
      transactions.push({
        date: new Date(),
        type: 'expense',
        account: '',
        amount: 0,
        currency: 'RUB',
        source: sourceName,
        status: 'needs_review',
        errors: [{ field: 'General', message: 'Ошибка парсинга (строка ' + (i + 1) + '): ' + e.toString() }],
        rawData: JSON.stringify(rawData[i])
      });
    }
  }
  
  // Update cache
  PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify(existingKeys));
  
  // Calculate processed count (startIndex + batch length)
  var processed = (options._startIndex || startIndex) + rawData.length;
  var totalCount = options._totalCount || rawData.length;
  
  return {
    transactions: transactions,
    stats: stats,
    processed: processed,
    total: totalCount,
    hasMore: processed < totalCount
  };
}

/**
 * Step 4: Write preview to staging sheet.
 * Public function for HTML Service.
 * Note: transactions is passed as JSON string to avoid size limits.
 * @param {string} transactionsJson - Transactions as JSON string
 * @returns {Object} Preview result
 */
function pfWritePreview(transactionsJson) {
  // Parse JSON string back to array
  var transactions = JSON.parse(transactionsJson);
  return pfPreviewImport_(transactions);
}

/**
 * Process file import (called from UI) - simplified version that calls steps.
 * @param {string} fileContent - File content as string
 * @param {Object} options - Import options
 * @returns {Object} Preview result
 */
function pfProcessFileImport_(fileContent, options) {
  options = options || {};
  
  try {
    // Step 1: Detect format
    var formatInfo = pfDetectFileFormat(fileContent);
    if (!formatInfo.detected) {
      throw new Error('Формат файла не поддерживается. Используйте CSV файл или выписку Сбербанка.');
    }
    
    // Step 2: Parse
    var parseResult = pfParseFileContent(fileContent, formatInfo.importerType, options);
    if (parseResult.count === 0) {
      throw new Error('Файл пуст или не содержит данных для импорта');
    }
    
    // Step 3: Process all data (for now, process in one go, but can be batched)
    var allTransactions = [];
    var totalStats = {
      valid: 0,
      needsReview: 0,
      duplicates: 0,
      errors: 0
    };
    
    var batchSize = 200; // Process 200 at a time
    var processed = 0;
    
        while (processed < parseResult.rawData.length) {
      var batchResult = pfProcessDataBatch(JSON.stringify(parseResult.rawData), formatInfo.importerType, options, batchSize, processed);
      allTransactions = allTransactions.concat(batchResult.transactions);
      totalStats.valid += batchResult.stats.valid;
      totalStats.needsReview += batchResult.stats.needsReview;
      totalStats.duplicates += batchResult.stats.duplicates;
      totalStats.errors += batchResult.stats.errors;
      processed = batchResult.processed;
    }
    
    if (allTransactions.length === 0) {
      throw new Error('Не удалось обработать транзакции из файла');
    }
    
    // Step 4: Preview
    var preview = pfWritePreview(JSON.stringify(allTransactions));
    preview.stats.total = allTransactions.length;
    preview.stats.valid = totalStats.valid;
    preview.stats.needsReview = totalStats.needsReview;
    preview.stats.duplicates = totalStats.duplicates;
    
    return preview;
  } catch (error) {
    var errorMessage = error.message || error.toString();
    if (errorMessage.indexOf('timeout') !== -1 || errorMessage.indexOf('exceeded') !== -1) {
      throw new Error('Файл слишком большой или обработка заняла слишком много времени. Попробуйте разбить файл на части или уменьшить количество транзакций.');
    }
    throw error;
  }
}

/**
 * Get list of accounts for dropdown.
 * @returns {Array<string>}
 */
function pfGetAccountsList_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var accountsSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.ACCOUNTS);
  if (!accountsSheet || accountsSheet.getLastRow() <= 1) {
    return [];
  }
  
  var accounts = [];
  var accountCol = pfColumnIndex_(PF_ACCOUNTS_SCHEMA, 'Account');
  if (accountCol) {
    var data = accountsSheet.getRange(2, accountCol, accountsSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      var account = String(data[i][0] || '').trim();
      if (account) {
        accounts.push(account);
      }
    }
  }
  
  return accounts;
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
