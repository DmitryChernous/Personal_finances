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
  // Normalize sourceId - convert to string and trim
  var sourceId = transaction.sourceId ? String(transaction.sourceId).trim() : '';
  
  if (sourceId && sourceId !== '') {
    // Use sourceId if available
    var source = String(transaction.source || '').trim();
    return source + ':' + sourceId;
  }
  
  // Fallback: hash of key fields
  var dateStr = '';
  if (transaction.date) {
    // Handle both Date objects and ISO strings
    if (transaction.date instanceof Date) {
      dateStr = Utilities.formatDate(transaction.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof transaction.date === 'string') {
      // Try to parse ISO string
      try {
        var dateObj = new Date(transaction.date);
        if (!isNaN(dateObj.getTime())) {
          dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      } catch (e) {
        // Ignore
      }
    }
  }
  
  var keyFields = [
    dateStr,
    String(transaction.account || '').trim(),
    String(transaction.amount || ''),
    String(transaction.type || '').trim()
  ].join('|');
  
  var source = String(transaction.source || '').trim();
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
  Logger.log('[SERVER] pfTransactionDTOToRow_ called, transaction keys: ' + Object.keys(transaction).join(', '));
  
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
    
    Logger.log('[SERVER] Column ' + i + ': schemaKey=' + col.key + ', transactionKey=' + transactionKey + ', value=' + (value !== undefined ? String(value).substring(0, 50) : 'undefined'));
    
    // Handle special cases
    if (col.key === 'Date') {
      // Date might be ISO string (from server) or Date object
      if (value instanceof Date) {
        row.push(value);
      } else if (typeof value === 'string' && value.length > 0) {
        // Try to parse ISO string back to Date
        try {
          var dateObj = new Date(value);
          if (!isNaN(dateObj.getTime())) {
            row.push(dateObj);
          } else {
            Logger.log('[SERVER] WARNING: Invalid date string: ' + value);
            row.push('');
          }
        } catch (e) {
          Logger.log('[SERVER] ERROR parsing date: ' + e.toString());
          row.push('');
        }
      } else {
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
  
  Logger.log('[SERVER] pfTransactionDTOToRow_ returning row with ' + row.length + ' values');
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
 * Private function - loads all keys from Transactions sheet.
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
    var source = String(row[sourceCol - 1] || '').trim();
    var sourceId = sourceIdCol ? String(row[sourceIdCol - 1] || '').trim() : '';
    
    // Only process rows with source
    if (!source) {
      continue;
    }
    
    if (sourceId && sourceId !== '') {
      // Use sourceId if available
      keys[source + ':' + sourceId] = true;
    } else {
      // Generate hash key if no sourceId
      var date = row[dateCol - 1];
      var account = String(row[accountCol - 1] || '').trim();
      var amount = row[amountCol - 1];
      var type = String(row[typeCol - 1] || '').trim();
      
      var keyFields = [
        date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        account,
        String(amount || ''),
        type
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
  Logger.log('[SERVER] pfPreviewImport_ called with ' + transactions.length + ' transactions');
  
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
    
    if (i < 3) {
      Logger.log('[SERVER] Sample transaction ' + i + ' keys: ' + Object.keys(tx).join(', '));
      Logger.log('[SERVER] Sample transaction ' + i + ' data: date=' + tx.date + ', type=' + tx.type + ', account=' + tx.account + ', amount=' + tx.amount);
    }
    
    var row = pfTransactionDTOToRow_(tx);
    
    if (i < 3) {
      Logger.log('[SERVER] Sample row ' + i + ' length: ' + row.length + ', first 5 values: ' + row.slice(0, 5).join(', '));
    }
    
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
  
  Logger.log('[SERVER] Prepared ' + rows.length + ' rows for writing');
  Logger.log('[SERVER] Row length: ' + (rows.length > 0 ? rows[0].length : 0));
  
  if (rows.length > 0) {
    try {
      Logger.log('[SERVER] Writing to sheet: range=2,1,' + rows.length + ',' + rows[0].length);
      stagingSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      Logger.log('[SERVER] Successfully wrote ' + rows.length + ' rows to sheet');
    } catch (e) {
      Logger.log('[SERVER] ERROR writing to sheet: ' + e.toString());
      Logger.log('[SERVER] Error stack: ' + (e.stack || 'No stack'));
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
      } else if (transactions[i].status === 'duplicate') {
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
 * Find existing transaction by deduplication key.
 * Public function for HTML Service - helps user find duplicate transactions.
 * @param {string} dedupeKey - Deduplication key (e.g., 'import:sberbank:160620251146306495')
 * @returns {Object} {found: boolean, rowNum: number, transaction: Object}
 */
function pfFindDuplicateTransaction(dedupeKey) {
  Logger.log('[SERVER] pfFindDuplicateTransaction called with key: ' + dedupeKey);
  
  if (!dedupeKey || typeof dedupeKey !== 'string') {
    return { found: false, message: 'Неверный ключ дедупликации' };
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
  
  Logger.log('[SERVER] Looking for source: ' + source + ', suffix: ' + keySuffix);
  
  var checkedCount = 0;
  var sampleKeys = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowSource = String(row[sourceCol - 1] || '').trim();
    
    // Only check rows with matching source
    if (rowSource !== source) {
      continue;
    }
    
    checkedCount++;
    var rowSourceId = sourceIdCol ? String(row[sourceIdCol - 1] || '').trim() : '';
    
    // Check if this row matches the dedupeKey
    var rowKey = null;
    if (rowSourceId && rowSourceId !== '') {
      rowKey = rowSource + ':' + rowSourceId;
      if (checkedCount <= 5) {
        sampleKeys.push('Row ' + (i + 2) + ': ' + rowKey);
      }
    } else {
      // Generate hash key
      var date = row[dateCol - 1];
      var account = String(row[accountCol - 1] || '').trim();
      var amount = row[amountCol - 1];
      var type = String(row[typeCol - 1] || '').trim();
      
      var keyFields = [
        date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        account,
        String(amount || ''),
        type
      ].join('|');
      
      var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, keyFields).map(function(b) {
        return ('0' + (b & 0xFF).toString(16)).slice(-2);
      }).join('');
      
      rowKey = (rowSource || 'unknown') + ':' + hash;
      if (checkedCount <= 5) {
        sampleKeys.push('Row ' + (i + 2) + ': ' + rowKey + ' (hash from: ' + keyFields + ')');
      }
    }
    
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
      
      Logger.log('[SERVER] Found duplicate at row ' + (i + 2));
      return {
        found: true,
        rowNum: i + 2, // 1-based row number (header is row 1)
        transaction: transaction,
        message: 'Найдена дублирующая транзакция в строке ' + (i + 2) + ' листа "Транзакции"'
      };
    }
  }
  
  Logger.log('[SERVER] Duplicate not found. Checked ' + checkedCount + ' transactions with source "' + source + '"');
  Logger.log('[SERVER] Sample keys from existing transactions: ' + sampleKeys.join('; '));
  
  return { 
    found: false, 
    message: 'Дублирующая транзакция не найдена в листе "Транзакции". Проверено ' + checkedCount + ' транзакций с источником "' + source + '". Возможно, это ложное срабатывание дедупликации - транзакция будет добавлена при импорте.' 
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
  // Валидация входных данных
  if (!fileContent || typeof fileContent !== 'string') {
    throw new Error('fileContent must be a non-empty string');
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
 * Step 2: Parse file content.
 * Public function for HTML Service.
 * @param {string} fileContent - File content as string
 * @param {string} importerType - 'sberbank' or 'csv'
 * @param {Object} options - Parse options
 * @returns {Object} {rawData: Array, count: number, errors: Array}
 */
function pfParseFileContent(fileContent, importerType, options) {
  // Валидация входных данных
  if (!fileContent || typeof fileContent !== 'string') {
    throw new Error('fileContent must be a non-empty string');
  }
  if (!importerType || !['sberbank', 'csv'].includes(importerType)) {
    throw new Error('Invalid importerType: ' + importerType);
  }
  
  // Ограничение размера файла (50MB)
  var maxSize = 50 * 1024 * 1024; // 50MB
  if (fileContent.length > maxSize) {
    throw new Error('File too large: ' + Math.round(fileContent.length / 1024 / 1024) + 'MB. Maximum is 50MB.');
  }
  
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
 * @param {string} importerType - 'sberbank' or 'csv'
 * @param {Object} options - Processing options
 * @param {number} batchSize - Number of items to process per batch (default: 100)
 * @param {number} startIndex - Start index for this batch (default: 0)
 * @returns {Object} {transactions: Array, stats: Object, processed: number, total: number, hasMore: boolean}
 */
function pfProcessDataBatch(rawDataJson, importerType, options, batchSize, startIndex) {
  Logger.log('[SERVER] pfProcessDataBatch called at ' + new Date().toISOString());
  Logger.log('[SERVER] Parameters: importerType=' + importerType + ', batchSize=' + batchSize + ', startIndex=' + startIndex);
  Logger.log('[SERVER] rawDataJson length: ' + (rawDataJson ? rawDataJson.length : 0));
  
  try {
    // Валидация входных данных
    if (!rawDataJson || typeof rawDataJson !== 'string') {
      Logger.log('[SERVER] ERROR: rawDataJson validation failed');
      throw new Error('rawDataJson must be a non-empty string');
    }
    if (!importerType || !['sberbank', 'csv'].includes(importerType)) {
      Logger.log('[SERVER] ERROR: Invalid importerType: ' + importerType);
      throw new Error('Invalid importerType: ' + importerType);
    }
    batchSize = batchSize || 100;
    if (batchSize < 1 || batchSize > 1000) {
      Logger.log('[SERVER] ERROR: Invalid batchSize: ' + batchSize);
      throw new Error('batchSize must be between 1 and 1000');
    }
    startIndex = startIndex || 0;
    options = options || {};
    
    Logger.log('[SERVER] Validation passed');
    
    // Parse JSON string back to array
    // Note: rawDataJson now contains only the batch data, not the entire array
    var rawData = null;
    try {
      Logger.log('[SERVER] Parsing JSON...');
      rawData = JSON.parse(rawDataJson);
      Logger.log('[SERVER] JSON parsed successfully, array length: ' + rawData.length);
    } catch (e) {
      Logger.log('[SERVER] ERROR: JSON parse failed: ' + e.toString());
      throw new Error('Invalid JSON in rawDataJson: ' + e.toString());
    }
    
    if (!Array.isArray(rawData)) {
      Logger.log('[SERVER] ERROR: rawData is not an array');
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
    } else {
      Logger.log('[SERVER] ERROR: Unknown importer type');
      throw new Error('Неизвестный тип импортера: ' + importerType);
    }
    
    if (!importer) {
      Logger.log('[SERVER] ERROR: Importer is null/undefined');
      throw new Error('Importer not found');
    }
    
    var sourceName = importerType === 'sberbank' ? 'import:sberbank' : 'import:csv';
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
        
        // Convert Date objects to ISO strings for JSON serialization
        // This prevents issues when passing through google.script.run
        if (transaction.date instanceof Date) {
          transaction.date = transaction.date.toISOString();
        }
        
        transactions.push(transaction);
      } catch (e) {
        Logger.log('[SERVER] ERROR processing item ' + i + ': ' + e.toString());
        Logger.log('[SERVER] Error stack: ' + (e.stack || 'No stack'));
        stats.errors++;
        transactions.push({
          date: new Date().toISOString(), // Use ISO string instead of Date object
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
    
    var processDuration = new Date().getTime() - processStartTime;
    Logger.log('[SERVER] Processing completed in ' + processDuration + 'ms');
    Logger.log('[SERVER] Results: transactions=' + transactions.length + ', stats=' + JSON.stringify(stats));
    
    // Calculate processed count (startIndex + batch length)
    var processed = (options._startIndex || startIndex) + rawData.length;
    var totalCount = options._totalCount || rawData.length;
    
    Logger.log('[SERVER] Calculated: processed=' + processed + ', totalCount=' + totalCount + ', hasMore=' + (processed < totalCount));
    Logger.log('[SERVER] Preparing result object...');
    Logger.log('[SERVER] Result size check: transactions=' + transactions.length + ', existingKeys=' + Object.keys(existingKeys).length);
    
    // Check if result might be too large for google.script.run
    // Try to estimate size (rough approximation)
    var estimatedSize = JSON.stringify(transactions).length;
    Logger.log('[SERVER] Estimated transactions JSON size: ' + estimatedSize + ' bytes');
    
    // If result is too large, we might need to split it
    // But first, let's try to return it and see if it works
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
      var testSerialization = JSON.stringify(result);
      Logger.log('[SERVER] Result serialization test: SUCCESS, size: ' + testSerialization.length + ' bytes');
    } catch (e) {
      Logger.log('[SERVER] ERROR: Result serialization failed: ' + e.toString());
      throw new Error('Failed to serialize result: ' + e.toString());
    }
    
    Logger.log('[SERVER] Returning result...');
    Logger.log('[SERVER] pfProcessDataBatch completed successfully');
    
    return result;
    
  } catch (e) {
    Logger.log('[SERVER] FATAL ERROR in pfProcessDataBatch: ' + e.toString());
    Logger.log('[SERVER] Error stack: ' + (e.stack || 'No stack'));
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
  if (!transactionsJson || typeof transactionsJson !== 'string') {
    throw new Error('transactionsJson must be a non-empty string');
  }
  
  // Parse JSON string back to array
  var transactions = null;
  try {
    transactions = JSON.parse(transactionsJson);
  } catch (e) {
    throw new Error('Invalid JSON in transactionsJson: ' + e.toString());
  }
  
  if (!Array.isArray(transactions)) {
    throw new Error('transactions must be an array');
  }
  
  // Ограничение размера для безопасности
  if (transactions.length > 10000) {
    throw new Error('Too many transactions: ' + transactions.length + '. Maximum is 10000.');
  }
  
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
 * Get existing transaction deduplication keys.
 * Public function for HTML Service (called once before batch processing).
 * @returns {Object} Map of dedupeKey -> true
 */
function pfGetExistingTransactionKeys() {
  Logger.log('[SERVER] pfGetExistingTransactionKeys called at ' + new Date().toISOString());
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
  Logger.log('[SERVER] pfGetAccountsList called');
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
    
    Logger.log('[SERVER] pfGetAccountsList returning ' + accounts.length + ' accounts: ' + accounts.join(', '));
    return accounts;
  } catch (e) {
    Logger.log('[SERVER] ERROR in pfGetAccountsList: ' + e.toString());
    Logger.log('[SERVER] Error stack: ' + (e.stack || 'No stack'));
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
