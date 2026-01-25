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
    
    // TODO: Implement actual parsing logic based on real Sberbank PDF format
    // This is a placeholder that needs to be refined with real examples
    
    // Look for transaction patterns
    // Sberbank PDF typically has:
    // - Date in format dd.MM.yyyy
    // - Amount with spaces and comma (e.g., "1 500,00")
    // - Description/merchant name
    // - Transaction type (debit/credit)
    
    var datePattern = /(\d{2}\.\d{2}\.\d{4})/;
    var amountPattern = /([\d\s]+,\d{2})/;
    
    var currentTransaction = null;
    
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      
      if (line.length === 0) {
        continue;
      }
      
      // Try to find date
      var dateMatch = line.match(datePattern);
      var amountMatch = line.match(amountPattern);
      
      if (dateMatch && amountMatch) {
        // This might be a transaction line
        // Save previous transaction if exists
        if (currentTransaction) {
          transactions.push(currentTransaction);
        }
        
        // Start new transaction
        var dateStr = dateMatch[1];
        var amountStr = amountMatch[1].replace(/\s/g, '').replace(',', '.');
        var amount = parseFloat(amountStr) || 0;
        
        // Determine type (simplified - needs refinement)
        var type = 'expense';
        if (line.toLowerCase().indexOf('зачислен') !== -1 || 
            line.toLowerCase().indexOf('доход') !== -1 ||
            line.toLowerCase().indexOf('пополнение') !== -1) {
          type = 'income';
        }
        
        currentTransaction = {
          bank: 'sberbank',
          date: dateStr,
          amount: amount,
          type: type,
          description: line,
          rawLine: line
        };
      } else if (currentTransaction) {
        // Continuation of description
        if (currentTransaction.description) {
          currentTransaction.description += ' ' + line;
        } else {
          currentTransaction.description = line;
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
    
    // Extract merchant/description from description
    var description = rawTransaction.description || '';
    var merchant = '';
    
    // Try to extract merchant name (first part of description, before common separators)
    if (description) {
      var parts = description.split(/[\.\s]{2,}/); // Split on multiple dots or spaces
      if (parts.length > 0) {
        merchant = parts[0].trim();
      }
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
      sourceId: rawTransaction.date + '_' + amount + '_' + (merchant || description.substring(0, 20)),
      rawData: JSON.stringify(rawTransaction),
      errors: []
    };
    
    // Validate
    var errors = pfValidateTransactionDTO_(transaction);
    transaction.errors = errors;
    
    return transaction;
  }
};
