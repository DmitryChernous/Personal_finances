/**
 * CSV importer implementation.
 *
 * Handles CSV files with flexible column mapping.
 * Supports common CSV formats from banks.
 */

/**
 * CSV importer object.
 */
var PF_CSV_IMPORTER = {
  /**
   * Detect if data is CSV format.
   * @param {Blob|string|Array<Array<*>>} data
   * @param {string} [fileName]
   * @returns {boolean}
   */
  detect: function(data, fileName) {
    if (fileName) {
      var ext = fileName.toLowerCase().split('.').pop();
      if (ext === 'csv') return true;
    }
    
    if (typeof data === 'string') {
      // Check if it looks like CSV (has commas or semicolons, multiple lines)
      var lines = data.split('\n');
      if (lines.length > 1 && (data.indexOf(',') !== -1 || data.indexOf(';') !== -1)) {
        return true;
      }
    }
    
    return false;
  },

  /**
   * Parse CSV data into raw transaction objects.
   * @param {Blob|string|Array<Array<*>>} data
   * @param {Object} [options]
   * @param {Object} [options.columnMapping] - Map of CSV column indices/names to transaction fields
   * @param {string} [options.delimiter] - CSV delimiter (default: auto-detect)
   * @param {number} [options.headerRow] - Row index with headers (0-based, default: 0)
   * @returns {Array<Object>}
   */
  parse: function(data, options) {
    options = options || {};
    var rows = [];
    
    // Convert data to array of rows
    if (data instanceof Blob) {
      // For file uploads, we'll need to handle this in the UI
      throw new Error('File blob parsing not yet implemented. Please use file content string.');
    } else if (typeof data === 'string') {
      // Parse CSV string
      var delimiter = options.delimiter || this._detectDelimiter_(data);
      var lines = data.split('\n');
      var headerRow = options.headerRow || 0;
      
      // Parse header
      var headers = this._parseCSVLine_(lines[headerRow], delimiter);
      
      // Parse data rows
      for (var i = headerRow + 1; i < lines.length; i++) {
        if (lines[i].trim() === '') continue;
        var values = this._parseCSVLine_(lines[i], delimiter);
        if (values.length === 0) continue;
        
        var rowObj = {};
        for (var j = 0; j < headers.length && j < values.length; j++) {
          rowObj[headers[j]] = values[j];
        }
        rows.push(rowObj);
      }
    } else if (Array.isArray(data)) {
      // Already array of rows
      rows = data;
    }
    
    return rows;
  },

  /**
   * Normalize raw CSV row into TransactionDTO.
   * @param {Object} rawRow
   * @param {Object} [options]
   * @param {Object} [options.columnMapping] - Map CSV column names to transaction fields
   * @param {string} [options.defaultAccount] - Default account if not in CSV
   * @param {string} [options.defaultCurrency] - Default currency (default: RUB)
   * @param {string} [options.source] - Source identifier (default: 'import:csv')
   * @returns {TransactionDTO}
   */
  normalize: function(rawRow, options) {
    options = options || {};
    var mapping = options.columnMapping || this._getDefaultMapping_(rawRow);
    var source = options.source || 'import:csv';
    var defaultCurrency = options.defaultCurrency || 'RUB';
    
    var transaction = {
      date: this._parseDate_(this._getMappedValue_(rawRow, mapping, 'date')),
      type: this._parseType_(this._getMappedValue_(rawRow, mapping, 'type'), this._getMappedValue_(rawRow, mapping, 'amount')),
      account: this._getMappedValue_(rawRow, mapping, 'account') || options.defaultAccount || '',
      accountTo: this._getMappedValue_(rawRow, mapping, 'accountTo') || '',
      amount: this._parseAmount_(this._getMappedValue_(rawRow, mapping, 'amount')),
      currency: this._getMappedValue_(rawRow, mapping, 'currency') || defaultCurrency,
      category: this._getMappedValue_(rawRow, mapping, 'category') || '',
      subcategory: this._getMappedValue_(rawRow, mapping, 'subcategory') || '',
      merchant: this._getMappedValue_(rawRow, mapping, 'merchant') || '',
      description: this._getMappedValue_(rawRow, mapping, 'description') || '',
      tags: this._getMappedValue_(rawRow, mapping, 'tags') || '',
      source: source,
      sourceId: this._getMappedValue_(rawRow, mapping, 'sourceId') || '',
      rawData: JSON.stringify(rawRow),
      errors: []
    };
    
    // Validate and collect errors
    var errors = pfValidateTransactionDTO_(transaction);
    transaction.errors = errors;
    
    // Set status based on errors
    if (errors.length > 0) {
      transaction.status = PF_TRANSACTION_STATUS.NEEDS_REVIEW;
    } else {
      transaction.status = PF_TRANSACTION_STATUS.OK;
    }
    
    return transaction;
  },

  /**
   * Generate deduplication key.
   * @param {TransactionDTO} transaction
   * @returns {string}
   */
  dedupeKey: function(transaction) {
    return pfGenerateDedupeKey_(transaction);
  },

  /**
   * Auto-detect CSV delimiter.
   * @private
   */
  _detectDelimiter_: function(csvText) {
    var firstLine = csvText.split('\n')[0];
    var commaCount = (firstLine.match(/,/g) || []).length;
    var semicolonCount = (firstLine.match(/;/g) || []).length;
    var tabCount = (firstLine.match(/\t/g) || []).length;
    
    if (tabCount > commaCount && tabCount > semicolonCount) return '\t';
    if (semicolonCount > commaCount) return ';';
    return ',';
  },

  /**
   * Parse a CSV line handling quoted fields.
   * @private
   */
  _parseCSVLine_: function(line, delimiter) {
    var result = [];
    var current = '';
    var inQuotes = false;
    
    for (var i = 0; i < line.length; i++) {
      var char = line[i];
      
      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === delimiter && !inQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    result.push(current.trim());
    
    return result;
  },

  /**
   * Get default column mapping by trying to match common column names.
   * @private
   */
  _getDefaultMapping_: function(rawRow) {
    var mapping = {};
    var keys = Object.keys(rawRow);
    
    // Common column name patterns
    var patterns = {
      date: ['дата', 'date', 'дата операции', 'дата транзакции', 'transaction date'],
      amount: ['сумма', 'amount', 'сумма операции', 'сумма транзакции', 'amount', 'сумма руб'],
      type: ['тип', 'type', 'операция', 'operation', 'приход/расход'],
      account: ['счет', 'account', 'счет отправителя', 'account from'],
      accountTo: ['счет получателя', 'account to', 'account_to', 'destination'],
      currency: ['валюта', 'currency', 'curr'],
      category: ['категория', 'category', 'кат'],
      description: ['описание', 'description', 'назначение', 'комментарий', 'comment', 'memo'],
      merchant: ['место', 'merchant', 'контрагент', 'магазин', 'store'],
      sourceId: ['id', 'source_id', 'transaction_id', 'operation_id']
    };
    
    for (var field in patterns) {
      for (var i = 0; i < keys.length; i++) {
        var key = keys[i].toLowerCase().trim();
        for (var j = 0; j < patterns[field].length; j++) {
          if (key.indexOf(patterns[field][j]) !== -1) {
            mapping[field] = keys[i];
            break;
          }
        }
        if (mapping[field]) break;
      }
    }
    
    return mapping;
  },

  /**
   * Get mapped value from raw row.
   * @private
   */
  _getMappedValue_: function(rawRow, mapping, field) {
    var mappedKey = mapping[field];
    if (!mappedKey) return null;
    return rawRow[mappedKey] || null;
  },

  /**
   * Parse date from various formats.
   * @private
   */
  _parseDate_: function(value) {
    if (!value) return null;
    if (value instanceof Date) return value;
    
    var str = String(value).trim();
    if (str === '') return null;
    
    // Try common date formats
    var formats = [
      'dd.MM.yyyy',
      'dd/MM/yyyy',
      'yyyy-MM-dd',
      'dd.MM.yy',
      'dd/MM/yy'
    ];
    
    for (var i = 0; i < formats.length; i++) {
      try {
        var date = Utilities.parseDate(str, Session.getScriptTimeZone(), formats[i]);
        if (date) return date;
      } catch (e) {
        // Try next format
      }
    }
    
    // Fallback: try JavaScript Date parsing
    try {
      var date = new Date(str);
      if (!isNaN(date.getTime())) return date;
    } catch (e) {
      // Ignore
    }
    
    return null;
  },

  /**
   * Parse transaction type from amount or explicit type field.
   * @private
   */
  _parseType_: function(typeValue, amountValue) {
    if (typeValue) {
      var type = String(typeValue).toLowerCase().trim();
      if (type === 'expense' || type === 'расход' || type === 'debit' || type === 'дебет') return 'expense';
      if (type === 'income' || type === 'доход' || type === 'credit' || type === 'кредит') return 'income';
      if (type === 'transfer' || type === 'перевод') return 'transfer';
    }
    
    // Try to infer from amount sign
    if (amountValue) {
      var amount = parseFloat(amountValue);
      if (amount < 0) return 'expense';
      if (amount > 0) return 'income';
    }
    
    return 'expense'; // Default
  },

  /**
   * Parse amount (always positive).
   * @private
   */
  _parseAmount_: function(value) {
    if (!value) return 0;
    
    var num = parseFloat(String(value).replace(/[^\d.,-]/g, '').replace(',', '.'));
    if (isNaN(num)) return 0;
    
    return Math.abs(num); // Always positive
  }
};
