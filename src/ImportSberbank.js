/**
 * Sberbank CSV importer.
 *
 * Handles specific format of Sberbank CSV exports:
 * - Headers start at row 14-17
 * - Transactions start at row 18
 * - Multi-line transactions (description spans multiple rows)
 * - Date format: dd.MM.yyyy
 * - Amount format: "1 500,00" (with spaces and comma)
 */

/**
 * Sberbank CSV importer object.
 */
var PF_SBERBANK_IMPORTER = {
  /**
   * Detect if data is Sberbank CSV format.
   * @param {Blob|string|Array<Array<*>>} data
   * @param {string} [fileName]
   * @returns {boolean}
   */
  detect: function(data, fileName) {
    if (typeof data !== 'string') return false;
    
    // Check for Sberbank-specific markers
    if (data.indexOf('Выписка по счёту') !== -1 || 
        data.indexOf('СберБанк Онлайн') !== -1 ||
        data.indexOf('ДАТА ОПЕРАЦИИ (МСК)') !== -1) {
      return true;
    }
    
    return false;
  },

  /**
   * Parse Sberbank CSV data.
   * @param {Blob|string|Array<Array<*>>} data
   * @param {Object} [options]
   * @returns {Array<Object>}
   */
  parse: function(data, options) {
    if (typeof data !== 'string') {
      throw new Error('Sberbank importer requires string data');
    }
    
    var lines = data.split('\n');
    var transactions = [];
    var currentTransaction = null;
    
    // Find all transaction sections (there can be multiple pages)
    var transactionSections = [];
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].indexOf('ДАТА ОПЕРАЦИИ (МСК)') !== -1) {
        // Transactions start 4 rows after header
        transactionSections.push(i + 4);
      }
    }
    
    if (transactionSections.length === 0) {
      throw new Error('Не найдено начало транзакций в файле Сбербанка');
    }
    
    // Parse transactions from all sections
    var allProcessed = false;
    for (var sectionIdx = 0; sectionIdx < transactionSections.length; sectionIdx++) {
      var startRow = transactionSections[sectionIdx];
      var endRow = sectionIdx < transactionSections.length - 1 
        ? transactionSections[sectionIdx + 1] - 10 // Before next section header
        : lines.length;
      
      for (var i = startRow; i < endRow; i++) {
        var line = lines[i].trim();
        
        // Skip empty lines and page breaks
        if (line === '' || 
            line.indexOf('Продолжение на следующей странице') !== -1 ||
            (line.indexOf('Страница') !== -1 && line.indexOf('из') !== -1)) {
          continue;
        }
        
        // Stop at footer
        if (line.indexOf('Для проверки подлинности') !== -1 ||
            line.indexOf('Действителен') !== -1 ||
            (line.indexOf('Выписка по счёту') !== -1 && line.indexOf('Страница') !== -1)) {
          // End of this section, but continue to next section
          if (currentTransaction) {
            transactions.push(currentTransaction);
            currentTransaction = null;
          }
          break;
        }
        
        // Parse CSV line (handle quoted fields) - ВНУТРИ ЦИКЛА
        var fields = this._parseCSVLine_(line, ',');
        
        // Check if this is a new transaction (has date in first column and amount in column 5)
        var hasDate = fields.length > 0 && this._isDate_(fields[0]);
        var hasAmount = fields.length > 4 && fields[4] && fields[4].trim() !== '';
        
        if (hasDate && hasAmount) {
          // Save previous transaction if exists
          if (currentTransaction) {
            transactions.push(currentTransaction);
          }
          
          // Start new transaction
          currentTransaction = {
            date: fields[0] || '',
            time: fields[1] || '',
            authCode: fields[2] || '',
            category: fields[3] || '', // Category is in column 4
            amount: fields[4] || '',
            balance: fields[5] || '',
            description: [] // Will collect multi-line description
          };
        } else if (currentTransaction && hasDate && !hasAmount) {
          // Line with date but no amount - continuation of description
          // Description is in column 4 (index 3)
          if (fields.length > 3 && fields[3] && fields[3].trim() !== '') {
            currentTransaction.description.push(fields[3].trim());
          }
        } else if (currentTransaction && !hasDate) {
          // Line without date - continuation of description
          // Description is in column 4 (index 3)
          if (fields.length > 3 && fields[3] && fields[3].trim() !== '') {
            currentTransaction.description.push(fields[3].trim());
          }
        }
      }
      
      // Don't forget last transaction from this section
      if (currentTransaction) {
        transactions.push(currentTransaction);
        currentTransaction = null;
      }
    }
    
    return transactions;
  },

  /**
   * Normalize Sberbank transaction to DTO.
   * @param {Object} rawRow
   * @param {Object} [options]
   * @returns {TransactionDTO}
   */
  normalize: function(rawRow, options) {
    options = options || {};
    var source = options.source || 'import:sberbank';
    var defaultCurrency = options.defaultCurrency || 'RUB';
    var defaultAccount = options.defaultAccount || '';
    
    // Parse date
    var date = this._parseDate_(rawRow.date);
    
    // Parse amount (format: "1 500,00" -> 1500.00)
    var amount = this._parseAmount_(rawRow.amount);
    
    // Determine type (all are expenses in Sberbank statements, except if amount is negative)
    var type = 'expense';
    if (amount < 0) {
      type = 'income';
      amount = Math.abs(amount);
    }
    
    // Combine description lines
    var description = rawRow.description && rawRow.description.length > 0 
      ? rawRow.description.join(' ') 
      : rawRow.category || '';
    
    // Extract merchant from description (usually first part before "RUS" or "Операция")
    var merchant = '';
    if (description) {
      var parts = description.split(/\.|RUS|Операция/);
      if (parts.length > 0) {
        merchant = parts[0].trim();
      }
    }
    
    // Generate sourceId from date + time + authCode + amount
    var sourceId = '';
    if (rawRow.date && rawRow.time && rawRow.authCode) {
      sourceId = rawRow.date.replace(/\./g, '') + rawRow.time.replace(/:/g, '') + rawRow.authCode;
    } else if (rawRow.date && rawRow.authCode) {
      sourceId = rawRow.date.replace(/\./g, '') + rawRow.authCode;
    }
    
    var transaction = {
      date: date,
      type: type,
      account: defaultAccount,
      accountTo: '',
      amount: amount,
      currency: defaultCurrency,
      category: rawRow.category || '',
      subcategory: '',
      merchant: merchant,
      description: description,
      tags: '',
      source: source,
      sourceId: sourceId,
      rawData: JSON.stringify(rawRow),
      errors: []
    };
    
    // Validate
    var errors = pfValidateTransactionDTO_(transaction);
    transaction.errors = errors;
    
    // Set status
    if (errors.length > 0) {
      transaction.status = 'needs_review';
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
   * Parse CSV line handling quoted fields.
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
   * Check if string is a date in format dd.MM.yyyy.
   * @private
   */
  _isDate_: function(str) {
    if (!str || typeof str !== 'string') return false;
    var datePattern = /^\d{2}\.\d{2}\.\d{4}$/;
    return datePattern.test(str.trim());
  },

  /**
   * Parse date from dd.MM.yyyy format.
   * @private
   */
  _parseDate_: function(value) {
    if (!value) return null;
    
    var str = String(value).trim();
    if (str === '') return null;
    
    // Format: dd.MM.yyyy
    try {
      var parts = str.split('.');
      if (parts.length === 3) {
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10) - 1; // Month is 0-based
        var year = parseInt(parts[2], 10);
        var date = new Date(year, month, day);
        if (!isNaN(date.getTime())) {
          return date;
        }
      }
    } catch (e) {
      // Ignore
    }
    
    return null;
  },

  /**
   * Parse amount from Sberbank format: "1 500,00" -> 1500.00
   * @private
   */
  _parseAmount_: function(value) {
    if (!value) return 0;
    
    var str = String(value);
    // Remove quotes if present
    str = str.replace(/^"|"$/g, '');
    // Remove spaces (thousand separators)
    str = str.replace(/\s/g, '');
    // Replace comma with dot (decimal separator)
    str = str.replace(',', '.');
    
    var num = parseFloat(str);
    if (isNaN(num)) return 0;
    
    return Math.abs(num); // Always positive (type determines income/expense)
  }
};
