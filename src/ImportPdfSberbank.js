/**
 * PDF parser for Sberbank statements.
 * 
 * Parses text extracted from Sberbank PDF statements.
 * This is a placeholder implementation that needs to be tested and refined
 * with real Sberbank PDF statements.
 */

/**
 * Sberbank PDF parser object.
 */
var PF_PDF_SBERBANK_PARSER = {
  /**
   * Detect if text is from Sberbank PDF statement.
   * @param {string} text - Extracted text from PDF
   * @returns {boolean}
   */
  detect: function(text) {
    if (!text || typeof text !== 'string') {
      return false;
    }
    
    var normalizedText = text.toLowerCase();
    
    // Check for Sberbank markers
    if (normalizedText.indexOf('сбербанк') !== -1 || 
        normalizedText.indexOf('sberbank') !== -1 ||
        normalizedText.indexOf('выписка по счёту') !== -1 ||
        normalizedText.indexOf('сбербанк онлайн') !== -1) {
      return true;
    }
    
    return false;
  },
  
  /**
   * Parse Sberbank PDF text into raw transactions.
   * @param {string} text - Extracted text from PDF
   * @param {Object} [options] - Parser options
   * @returns {Array<Object>} Array of raw transaction objects
   */
  parse: function(text, options) {
    options = options || {};
    
    if (!text || typeof text !== 'string') {
      throw new Error('Text is required for parsing');
    }
    
    if (!this.detect(text)) {
      throw new Error('Text does not appear to be from Sberbank PDF statement');
    }
    
    var transactions = [];
    var lines = text.split('\n');
    
    // Find transaction sections (similar to CSV parser)
    // Look for header "ДАТА ОПЕРАЦИИ" or similar markers
    var transactionSections = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (line.indexOf('ДАТА ОПЕРАЦИИ') !== -1 || 
          line.indexOf('Дата операции') !== -1 ||
          (line.indexOf('Дата') !== -1 && line.indexOf('Сумма') !== -1)) {
        // Transactions typically start a few lines after header
        transactionSections.push(i + 2);
      }
    }
    
    // If no sections found, try to find transactions by pattern
    if (transactionSections.length === 0) {
      // Look for lines with date and amount patterns
      for (var j = 0; j < lines.length; j++) {
        var testLine = lines[j].trim();
        var dateMatch = testLine.match(/(\d{2}\.\d{2}\.\d{4})/);
        var amountMatch = testLine.match(/([\d\s]+,\d{2})/);
        if (dateMatch && amountMatch) {
          transactionSections.push(j);
          break; // Start from first transaction found
        }
      }
    }
    
    // Patterns for parsing
    var datePattern = /(\d{2}\.\d{2}\.\d{4})/;
    var amountPattern = /([\d\s]+,\d{2})/;
    var timePattern = /(\d{2}:\d{2})/;
    
    var currentTransaction = null;
    var inTransactionSection = transactionSections.length === 0; // If no sections, parse all
    
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      
      // Check if we're entering a transaction section
      if (transactionSections.indexOf(i) !== -1) {
        inTransactionSection = true;
        continue;
      }
      
      // Stop at footer markers
      if (line.indexOf('Для проверки подлинности') !== -1 ||
          line.indexOf('Действителен') !== -1 ||
          (line.indexOf('Выписка по счёту') !== -1 && line.indexOf('Страница') !== -1) ||
          line.indexOf('Продолжение на следующей странице') !== -1) {
        if (currentTransaction) {
          transactions.push(currentTransaction);
          currentTransaction = null;
        }
        // Continue to next section if exists
        continue;
      }
      
      if (line.length === 0) {
        // Empty line might end current transaction description
        if (currentTransaction && currentTransaction.description && 
            currentTransaction.description.length > 50) {
          // Long description, might be complete
          // Continue to collect if next line has date+amount (new transaction)
        }
        continue;
      }
      
      // Skip page numbers and headers
      if ((line.indexOf('Страница') !== -1 && line.indexOf('из') !== -1) ||
          line.indexOf('Продолжение') !== -1) {
        continue;
      }
      
      // Try to find date and amount in line
      var dateMatch = line.match(datePattern);
      var amountMatch = line.match(amountPattern);
      var timeMatch = line.match(timePattern);
      
      // Check if this is a new transaction line
      // Transaction line typically has: date, time (optional), amount, description
      var isNewTransaction = false;
      
      if (dateMatch && amountMatch) {
        // Has both date and amount - likely a transaction line
        isNewTransaction = true;
      } else if (dateMatch && !currentTransaction) {
        // Has date but no current transaction - might be start of transaction
        // Check if next lines have amount
        for (var k = i + 1; k < Math.min(i + 3, lines.length); k++) {
          var nextLine = lines[k].trim();
          if (nextLine.match(amountPattern)) {
            isNewTransaction = true;
            break;
          }
        }
      }
      
      if (isNewTransaction) {
        // Save previous transaction if exists
        if (currentTransaction) {
          transactions.push(currentTransaction);
        }
        
        // Extract date
        var dateStr = dateMatch ? dateMatch[1] : '';
        
        // Extract amount (look in current line or next lines)
        var amountStr = '';
        var amountValue = 0;
        if (amountMatch) {
          amountStr = amountMatch[1];
        } else {
          // Look in next 2 lines
          for (var m = i + 1; m < Math.min(i + 3, lines.length); m++) {
            var nextLine = lines[m].trim();
            var nextAmountMatch = nextLine.match(amountPattern);
            if (nextAmountMatch) {
              amountStr = nextAmountMatch[1];
              break;
            }
          }
        }
        
        if (amountStr) {
          // Parse amount: "1 500,00" -> 1500.00
          amountValue = this._parseAmount_(amountStr);
        }
        
        // Extract time if available
        var timeStr = timeMatch ? timeMatch[1] : '';
        
        // Determine type: negative amount or keywords indicate income
        var type = 'expense';
        var lineLower = line.toLowerCase();
        if (amountValue < 0 || 
            lineLower.indexOf('зачислен') !== -1 || 
            lineLower.indexOf('зачисление') !== -1 ||
            lineLower.indexOf('пополнение') !== -1 ||
            lineLower.indexOf('возврат') !== -1 ||
            lineLower.indexOf('возврат средств') !== -1) {
          type = 'income';
          amountValue = Math.abs(amountValue);
        }
        
        // Extract description (everything except date, time, amount)
        var description = line;
        if (dateStr) {
          description = description.replace(dateStr, '').trim();
        }
        if (timeStr) {
          description = description.replace(timeStr, '').trim();
        }
        if (amountStr) {
          description = description.replace(amountStr, '').trim();
        }
        // Remove common separators
        description = description.replace(/^[\.\-\s]+/, '').trim();
        
        currentTransaction = {
          bank: 'sberbank',
          date: dateStr,
          time: timeStr,
          amount: amountValue,
          type: type,
          description: description || '',
          rawLine: line
        };
      } else if (currentTransaction) {
        // Continuation of description
        // Check if this line is part of description or a new transaction
        var hasDate = line.match(datePattern);
        var hasAmount = line.match(amountPattern);
        
        if (!hasDate && !hasAmount) {
          // Likely continuation of description
          if (currentTransaction.description) {
            currentTransaction.description += ' ' + line;
          } else {
            currentTransaction.description = line;
          }
        } else if (hasDate && hasAmount) {
          // This is actually a new transaction, save previous
          transactions.push(currentTransaction);
          // Start new transaction (will be handled in next iteration)
          currentTransaction = null;
        }
      }
    }
    
    // Don't forget last transaction
    if (currentTransaction) {
      transactions.push(currentTransaction);
    }
    
    if (transactions.length === 0) {
      throw new Error('Не удалось найти транзакции в PDF файле. Убедитесь, что файл является выпиской Сбербанка.');
    }
    
    return transactions;
  },
  
  /**
   * Parse amount from Sberbank format: "1 500,00" -> 1500.00
   * @private
   */
  _parseAmount_: function(amountStr) {
    if (!amountStr) return 0;
    
    // Remove spaces and replace comma with dot
    var cleaned = String(amountStr).replace(/\s/g, '').replace(',', '.');
    var amount = parseFloat(cleaned);
    
    if (isNaN(amount)) return 0;
    
    return amount;
  },
  
  /**
   * Normalize raw Sberbank PDF transaction to DTO.
   * @param {Object} rawTransaction - Raw transaction from parse()
   * @param {Object} [options] - Normalization options
   * @returns {TransactionDTO}
   */
  normalize: function(rawTransaction, options) {
    options = options || {};
    var source = options.source || 'import:pdf:sberbank';
    var defaultCurrency = options.defaultCurrency || 'RUB';
    var defaultAccount = options.defaultAccount || '';
    
    // Parse date
    var date = null;
    if (rawTransaction.date) {
      try {
        var dateParts = rawTransaction.date.split('.');
        if (dateParts.length === 3) {
          var day = parseInt(dateParts[0], 10);
          var month = parseInt(dateParts[1], 10) - 1; // Month is 0-based
          var year = parseInt(dateParts[2], 10);
          date = new Date(year, month, day);
        }
      } catch (e) {
        pfLogWarning_('Error parsing date: ' + rawTransaction.date, 'PF_PDF_SBERBANK_PARSER.normalize');
      }
    }
    
    // Parse amount
    var amount = Math.abs(rawTransaction.amount || 0);
    
    // Determine type
    var type = rawTransaction.type || 'expense';
    
    // Extract merchant/description from description (similar to CSV parser)
    var description = rawTransaction.description || '';
    var merchant = '';
    
    // Extract merchant from description (usually first part before "RUS" or "Операция")
    if (description) {
      var parts = description.split(/\.|RUS|Операция|операция/);
      if (parts.length > 0) {
        merchant = parts[0].trim();
        // Clean up merchant name
        merchant = merchant.replace(/[\.\-\s]{2,}/g, ' ').trim();
      }
    }
    
    // Generate sourceId (similar to CSV parser)
    var sourceId = '';
    if (rawTransaction.date && rawTransaction.time) {
      sourceId = rawTransaction.date.replace(/\./g, '') + rawTransaction.time.replace(/:/g, '');
    } else if (rawTransaction.date) {
      sourceId = rawTransaction.date.replace(/\./g, '');
    }
    if (sourceId && rawTransaction.amount) {
      sourceId += '_' + rawTransaction.amount;
    }
    if (!sourceId) {
      // Fallback: use date + amount + merchant
      sourceId = (rawTransaction.date || '') + '_' + amount + '_' + (merchant || description.substring(0, 20));
    }
    
    var transaction = {
      date: date,
      type: type,
      account: defaultAccount,
      accountTo: '',
      amount: amount,
      currency: defaultCurrency,
      category: '',
      subcategory: '',
      merchant: merchant,
      description: description,
      tags: '',
      source: source,
      sourceId: sourceId,
      rawData: JSON.stringify(rawTransaction),
      errors: []
    };
    
    // Validate
    var errors = pfValidateTransactionDTO_(transaction);
    transaction.errors = errors;
    
    return transaction;
  }
};
