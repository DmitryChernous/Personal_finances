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
    
    // Find all transaction sections (similar to CSV parser - handle multiple pages)
    // Look for all headers "ДАТА ОПЕРАЦИИ (МСК)" - transactions start 2 lines after each header
    var transactionSections = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (line.indexOf('ДАТА ОПЕРАЦИИ (МСК)') !== -1 || 
          line.indexOf('ДАТА ОПЕРАЦИИ') !== -1) {
        // Transactions start 2 lines after header (skip "Дата обработки" and empty line)
        // But check if next lines are headers - if so, skip more
        var nextLineIdx = i + 1;
        var skipLines = 2;
        while (nextLineIdx < lines.length && nextLineIdx < i + 5) {
          var nextLine = lines[nextLineIdx].trim();
          if (nextLine.indexOf('Дата обработки') !== -1 ||
              nextLine.indexOf('КАТЕГОРИЯ') !== -1 ||
              nextLine.indexOf('СУММА') !== -1 ||
              nextLine.length === 0) {
            skipLines++;
            nextLineIdx++;
          } else {
            break;
          }
        }
        transactionSections.push(i + skipLines);
      }
    }
    
    // If no sections found, try to find first transaction by pattern
    if (transactionSections.length === 0) {
      for (var j = 0; j < lines.length; j++) {
        var testLine = lines[j].trim();
        // Look for pattern: date + time + code + category + amount
        if (testLine.match(/\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}\s+\d{6}/)) {
          transactionSections.push(j);
          break;
        }
      }
    }
    
    if (transactionSections.length === 0) {
      throw new Error('Не найдено начало транзакций в PDF файле');
    }
    
    // Patterns for parsing
    // Transaction line format: "31.12.2025 16:40 966521 Перевод СБП 1 500,00 96 776,18"
    var transactionLinePattern = /(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2})\s+(\d{6})\s+(.+?)\s+([\d\s]+,\d{2})\s+([\d\s]+,\d{2})/;
    var datePattern = /(\d{2}\.\d{2}\.\d{4})/;
    var amountPattern = /([\d\s]+,\d{2})/;
    
    // Process all transaction sections (handle multiple pages)
    for (var sectionIdx = 0; sectionIdx < transactionSections.length; sectionIdx++) {
      var startRow = transactionSections[sectionIdx];
      var endRow = sectionIdx < transactionSections.length - 1 
        ? transactionSections[sectionIdx + 1] - 2 // Before next section header
        : lines.length;
      
      var currentTransaction = null;
      
      for (var i = startRow; i < endRow; i++) {
        var line = lines[i].trim();
        
        // Stop at footer markers
        if (line.indexOf('Для проверки подлинности') !== -1 ||
            line.indexOf('Действителен') !== -1) {
          if (currentTransaction) {
            transactions.push(currentTransaction);
            currentTransaction = null;
          }
          break;
        }
        
        // Skip empty lines, page numbers, and section headers
        if (line.length === 0 ||
            (line.indexOf('Страница') !== -1 && line.indexOf('из') !== -1) ||
            line.indexOf('Продолжение на следующей странице') !== -1 ||
            line.indexOf('--') !== -1 ||
            line.indexOf('ДАТА ОПЕРАЦИИ') !== -1 ||
            line.indexOf('Дата обработки') !== -1 ||
            line.indexOf('КАТЕГОРИЯ') !== -1 ||
            line.indexOf('СУММА В ВАЛЮТЕ') !== -1 ||
            line.indexOf('ОСТАТОК СРЕДСТВ') !== -1 ||
            (line.indexOf('Выписка по счёту') !== -1 && line.indexOf('Страница') !== -1)) {
          // If we hit a new section header within this section, save current transaction
          if (currentTransaction && line.indexOf('ДАТА ОПЕРАЦИИ') !== -1) {
            transactions.push(currentTransaction);
            currentTransaction = null;
          }
          continue;
        }
        
        // Try to match full transaction line pattern
        // Format: "31.12.2025 16:40 966521 Перевод СБП 1 500,00 96 776,18"
        var transactionMatch = line.match(transactionLinePattern);
        
        if (transactionMatch) {
          // This is a new transaction line
          // Save previous transaction if exists
          if (currentTransaction) {
            transactions.push(currentTransaction);
          }
          
          var dateStr = transactionMatch[1]; // "31.12.2025"
          var timeStr = transactionMatch[2]; // "16:40"
          var authCode = transactionMatch[3]; // "966521"
          var category = transactionMatch[4].trim(); // "Перевод СБП"
          var amountStr = transactionMatch[5]; // "1 500,00"
          var balanceStr = transactionMatch[6]; // "96 776,18" (not used, but good to have)
          
          // Parse amount
          var amountValue = this._parseAmount_(amountStr);
          
          // Determine transaction type
          // Income indicators: "+" in category, "зачислен", "пополнение", "возврат", "заработная плата"
          var type = 'expense';
          var categoryLower = category.toLowerCase();
          if (amountValue < 0 || 
              category.indexOf('+') !== -1 || // "+50", "+350" etc. indicates income
              categoryLower.indexOf('зачислен') !== -1 || 
              categoryLower.indexOf('пополнение') !== -1 ||
              categoryLower.indexOf('возврат') !== -1) {
            type = 'income';
            amountValue = Math.abs(amountValue);
          }
          
          // Start new transaction
          currentTransaction = {
            bank: 'sberbank',
            date: dateStr,
            time: timeStr,
            authCode: authCode,
            category: category,
            amount: amountValue,
            type: type,
            description: [], // Will collect multi-line description
            rawLine: line
          };
        } else {
          // Check if this is a continuation line (has date but no time/code/amount)
          var dateMatch = line.match(datePattern);
          var amountMatch = line.match(amountPattern);
          
          if (currentTransaction) {
            if (dateMatch && !amountMatch) {
              // Line with date but no amount - continuation of description
              // Remove date from beginning if present
              var descLine = line;
              if (dateMatch.index === 0) {
                descLine = line.substring(dateMatch[0].length).trim();
              }
              if (descLine && descLine.length > 0) {
                currentTransaction.description.push(descLine);
              }
            } else if (!dateMatch && !amountMatch) {
              // Line without date or amount - continuation of description
              if (line.length > 0) {
                currentTransaction.description.push(line);
              }
            } else if (dateMatch && amountMatch) {
              // Has both date and amount but didn't match full pattern
              // Might be a new transaction with different format
              // Save current and try to parse this line as new transaction
              transactions.push(currentTransaction);
              currentTransaction = null;
              
              // Try to extract what we can
              var dateStr2 = dateMatch[1];
              var amountStr2 = amountMatch[1];
              var amountValue2 = this._parseAmount_(amountStr2);
              
              currentTransaction = {
                bank: 'sberbank',
                date: dateStr2,
                time: '',
                authCode: '',
                category: '',
                amount: amountValue2,
                type: 'expense',
                description: [line],
                rawLine: line
              };
            }
          } else if (dateMatch && amountMatch) {
            // New transaction but format doesn't match full pattern
            // Try to extract basic info
            var dateStr3 = dateMatch[1];
            var amountStr3 = amountMatch[1];
            var amountValue3 = this._parseAmount_(amountStr3);
            
            currentTransaction = {
              bank: 'sberbank',
              date: dateStr3,
              time: '',
              authCode: '',
              category: '',
              amount: amountValue3,
              type: 'expense',
              description: [line],
              rawLine: line
            };
          }
        }
      }
      
      // Don't forget last transaction from this section
      if (currentTransaction) {
        transactions.push(currentTransaction);
        currentTransaction = null;
      }
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
    
    // Determine type - check description for income indicators
    var type = rawTransaction.type || 'expense';
    
    // Combine description lines (similar to CSV parser)
    var description = '';
    if (Array.isArray(rawTransaction.description)) {
      description = rawTransaction.description.join(' ').trim();
    } else if (rawTransaction.description) {
      description = String(rawTransaction.description).trim();
    }
    
    // Use category if description is empty
    if (!description && rawTransaction.category) {
      description = rawTransaction.category;
    }
    
    // Check description for income indicators (e.g., "Заработная плата")
    var descriptionLower = description.toLowerCase();
    if (type === 'expense' && (
        descriptionLower.indexOf('заработная плата') !== -1 ||
        descriptionLower.indexOf('зарплата') !== -1 ||
        descriptionLower.indexOf('зачислен') !== -1)) {
      type = 'income';
    }
    
    // Extract merchant from description (usually first part before "RUS" or "Операция")
    var merchant = '';
    if (description) {
      // Remove common patterns that are not merchant names
      var cleanDesc = description
        .replace(/по карте\s+\*\*\*\*\d+/gi, '') // Remove "по карте ****7426"
        .replace(/\*\*\*\*\d+/g, '') // Remove "****7426"
        .replace(/операция по карте/gi, '') // Remove "Операция по карте"
        .trim();
      
      // Split by common delimiters
      var parts = cleanDesc.split(/\.|RUS|Операция|операция/);
      if (parts.length > 0) {
        merchant = parts[0].trim();
        // Clean up merchant name
        merchant = merchant.replace(/[\.\-\s]{2,}/g, ' ').trim();
        // Remove quotes if present
        merchant = merchant.replace(/^["']|["']$/g, '');
        // Remove trailing dots and spaces
        merchant = merchant.replace(/\.+$/, '').trim();
      }
      
      // If merchant is still empty or too short, try to extract from full description
      if (!merchant || merchant.length < 3) {
        // Try to find merchant name before "Операция" or "RUS"
        var match = cleanDesc.match(/^(.+?)(?:\.\s*(?:RUS|Операция)|$)/i);
        if (match && match[1]) {
          merchant = match[1].trim();
          merchant = merchant.replace(/[\.\-\s]{2,}/g, ' ').trim();
          merchant = merchant.replace(/^["']|["']$/g, '');
        }
      }
    }
    
    // Generate sourceId (similar to CSV parser)
    // Format: date + time + authCode (like CSV: date + time + authCode)
    var sourceId = '';
    if (rawTransaction.date && rawTransaction.time && rawTransaction.authCode) {
      // Full format: date + time + authCode
      sourceId = rawTransaction.date.replace(/\./g, '') + 
                 rawTransaction.time.replace(/:/g, '') + 
                 rawTransaction.authCode;
    } else if (rawTransaction.date && rawTransaction.authCode) {
      // Date + authCode (time might be missing)
      sourceId = rawTransaction.date.replace(/\./g, '') + rawTransaction.authCode;
    } else if (rawTransaction.date && rawTransaction.time) {
      // Date + time (authCode might be missing)
      sourceId = rawTransaction.date.replace(/\./g, '') + rawTransaction.time.replace(/:/g, '');
    } else if (rawTransaction.date) {
      // Only date (fallback - should be rare)
      // Try to add something unique from description or amount
      var uniquePart = '';
      if (rawTransaction.authCode) {
        uniquePart = rawTransaction.authCode;
      } else if (amount > 0) {
        uniquePart = Math.round(amount).toString();
      } else if (merchant) {
        uniquePart = merchant.substring(0, 10).replace(/\s/g, '');
      }
      sourceId = rawTransaction.date.replace(/\./g, '') + (uniquePart || '000000');
    }
    
    // Final fallback if still empty
    if (!sourceId) {
      // Use date + amount + merchant (last resort)
      sourceId = (rawTransaction.date || '00000000').replace(/\./g, '') + '_' + 
                 Math.round(amount) + '_' + 
                 (merchant || description.substring(0, 10) || 'unknown').replace(/\s/g, '');
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
