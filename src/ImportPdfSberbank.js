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
    // Also handles: "28.11.2025 11:45 647377 Прочие операции +47 330,86 141 669,66"
    // The pattern needs to handle cases where category contains "+number" before the amount
    // We match from the end: balance, then amount (with possible spaces), then category
    var transactionLinePattern = /(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2})\s+(\d{6})\s+(.+?)\s+([\d\s]+,\d{2})\s+([\d\s]+,\d{2})/;
    // Improved pattern: match from end to handle "+number" in category correctly
    // Format: date time code category amount balance
    // We'll use a more sophisticated approach: find amount and balance from the end
    var transactionLinePatternImproved = /(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2})\s+(\d{6})\s+(.+)\s+([\d\s]+,\d{2})\s+([\d\s]+,\d{2})$/;
    var datePattern = /(\d{2}\.\d{2}\.\d{4})/;
    var amountPattern = /([\d\s]+,\d{2})/;
    
    // Process all transaction sections (handle multiple pages)
    for (var sectionIdx = 0; sectionIdx < transactionSections.length; sectionIdx++) {
      var startRow = transactionSections[sectionIdx];
      var endRow = sectionIdx < transactionSections.length - 1 
        ? transactionSections[sectionIdx + 1] - 2 // Before next section header
        : lines.length;
      
      var currentTransaction = null;
      
      // DEBUG: Log first 150 lines for debugging
      var debugLineCount = 0;
      var maxDebugLines = 150;
      
      for (var i = startRow; i < endRow; i++) {
        var line = lines[i].trim();
        
        // DEBUG: Log first lines
        if (debugLineCount < maxDebugLines) {
          Logger.log('[DEBUG] Line ' + (i + 1) + ': ' + line);
          debugLineCount++;
        }
        
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
        
        // Parse transaction line using strict rules
        // Rule 1: New transaction starts with date dd.mm.yyyy
        var datePattern = /^(\d{2}\.\d{2}\.\d{4})/;
        var dateMatch = line.match(datePattern);
        
        var transactionMatch = null;
        var hasPlusInCategory = false;
        
        if (dateMatch) {
          // This is a new transaction line
          var dateStr = dateMatch[1]; // "28.11.2025"
          
          // Rule 2: Extract date (first 10 characters or until first space)
          // Already have it from regex
          
          // Rule 3: Extract time - next 5 characters after date (format HH:MM)
          // Skip date (10 chars) and space, then take 5 chars
          var timeMatch = line.substring(10).trim().match(/^(\d{2}:\d{2})/);
          var timeStr = '';
          var authCode = '';
          var categoryPart = '';
          var amountStr = '';
          var balanceStr = '';
          
          if (timeMatch) {
            timeStr = timeMatch[1]; // "11:45"
            
            // Rule 4: Extract authorization code - next 6 digits after time
            // Skip date (10) + space + time (5) + space = ~17 chars from start
            var afterTime = line.substring(10 + timeStr.length).trim();
            var authCodeMatch = afterTime.match(/^(\d{6})/);
            
            if (authCodeMatch) {
              authCode = authCodeMatch[1]; // "647377"
              
              // Rule 5: Extract category - between auth code and amount
              // Category: letters (cyrillic/latin) + special chars, no digits
              // Amount: starts with digit or "+", contains spaces, ends with ,XX
              var afterAuthCode = afterTime.substring(6).trim();
              
              // Find amount pattern: starts with digit or "+", ends with ,XX
              // Pattern: optional "+" followed by digits/spaces, then comma and two digits
              // Important: "+46 696,61" should be matched as one amount, not "+46" and "696,61"
              // Use non-greedy matching from the end to find complete amounts
              var amountPattern = /(\+?[\d\s]+,\d{2})/g;
              var amountMatches = [];
              var match;
              
              // Find all matches
              while ((match = amountPattern.exec(afterAuthCode)) !== null) {
                amountMatches.push({
                  value: match[1],
                  index: match.index
                });
              }
              
              // Sort by index to process in order
              amountMatches.sort(function(a, b) {
                return a.index - b.index;
              });
              
              // We expect two amounts: transaction amount and balance
              if (amountMatches.length >= 2) {
                amountStr = amountMatches[0].value.trim(); // First amount (may start with "+")
                balanceStr = amountMatches[1].value.trim(); // Second amount (balance)
                
                // Category is everything between auth code and first amount
                var categoryEndIndex = amountMatches[0].index;
                categoryPart = afterAuthCode.substring(0, categoryEndIndex).trim();
                
                // IMPORTANT: If amount starts with "+", it's already a complete amount
                // Don't try to combine "+number" from category with amount
                // Example: "Прочие операции +46 696,61" -> category: "Прочие операции", amount: "+46 696,61"
                // The "+46" is part of the amount, not category!
                
                // Check if category ends with "+number" pattern (old format, e.g., "+47" separate from amount)
                // This was the old format where "+47" was in category and amount was "330,86"
                // New format: amount is already "+46 696,61" (complete)
                var plusNumberMatch = categoryPart.match(/\+\s*(\d+)\s*$/);
                if (plusNumberMatch && !amountStr.startsWith('+')) {
                  // Old format: category has "+number", amount doesn't start with "+"
                  // Combine them: "+47" + "330,86" = "47 330,86"
                  hasPlusInCategory = true;
                  var plusNumber = parseInt(plusNumberMatch[1], 10);
                  categoryPart = categoryPart.replace(/\+\s*\d+\s*$/, '').trim();
                  var currentAmount = this._parseAmount_(amountStr);
                  if (currentAmount < 1000 && plusNumber > 0) {
                    var combinedAmount = plusNumber * 1000 + currentAmount;
                    var intPart = Math.floor(combinedAmount);
                    var decPart = Math.round((combinedAmount - intPart) * 100);
                    var intStr = intPart.toString();
                    var formattedInt = '';
                    for (var k = intStr.length - 1, j = 0; k >= 0; k--, j++) {
                      if (j > 0 && j % 3 === 0) {
                        formattedInt = ' ' + formattedInt;
                      }
                      formattedInt = intStr[k] + formattedInt;
                    }
                    amountStr = formattedInt + ',' + (decPart < 10 ? '0' : '') + decPart;
                  }
                } else if (amountStr.startsWith('+')) {
                  // New format: amount already has "+", mark as income
                  hasPlusInCategory = true;
                }
                
                // Extract processing date and description after balance
                var afterBalance = afterAuthCode.substring(amountMatches[1].index + amountMatches[1].value.length).trim();
                var processingDate = '';
                var descriptionStart = '';
                
                // Check if there's a date after balance (processing date)
                var processingDateMatch = afterBalance.match(/^(\d{2}\.\d{2}\.\d{4})\s+(.+)$/);
                if (processingDateMatch) {
                  processingDate = processingDateMatch[1];
                  descriptionStart = processingDateMatch[2].trim();
                } else {
                  // No processing date, description starts immediately after balance
                  descriptionStart = afterBalance;
                }
                
                transactionMatch = {
                  1: dateStr,
                  2: timeStr,
                  3: authCode,
                  4: categoryPart,
                  5: amountStr,
                  6: balanceStr,
                  7: processingDate,
                  8: descriptionStart
                };
              }
            }
          }
        }
        
        // Fallback to original pattern if new strict approach didn't work
        if (!transactionMatch) {
          transactionMatch = line.match(transactionLinePattern);
          if (transactionMatch && transactionMatch[4]) {
            hasPlusInCategory = transactionMatch[4].indexOf('+') !== -1;
            // Handle "+number" in fallback too
            if (hasPlusInCategory) {
              var originalCategory = transactionMatch[4].trim();
              var originalAmount = transactionMatch[5].trim();
              var plusNumberMatch2 = originalCategory.match(/\+\s*(\d+)\s*$/);
              if (plusNumberMatch2) {
                var plusNumber2 = parseInt(plusNumberMatch2[1], 10);
                var cleanedCategory = originalCategory.replace(/\+\s*\d+\s*$/, '').trim();
                var currentAmount2 = this._parseAmount_(originalAmount);
                if (currentAmount2 < 1000 && plusNumber2 > 0) {
                  var combinedAmount2 = plusNumber2 * 1000 + currentAmount2;
                  var intPart2 = Math.floor(combinedAmount2);
                  var decPart2 = Math.round((combinedAmount2 - intPart2) * 100);
                  var intStr2 = intPart2.toString();
                  var formattedInt2 = '';
                  for (var k2 = intStr2.length - 1, j2 = 0; k2 >= 0; k2--, j2++) {
                    if (j2 > 0 && j2 % 3 === 0) {
                      formattedInt2 = ' ' + formattedInt2;
                    }
                    formattedInt2 = intStr2[k2] + formattedInt2;
                  }
                  var correctedAmount = formattedInt2 + ',' + (decPart2 < 10 ? '0' : '') + decPart2;
                  transactionMatch[4] = cleanedCategory;
                  transactionMatch[5] = correctedAmount;
                }
              }
            }
          }
        }
        
        if (transactionMatch) {
          // This is a new transaction line
          // Save previous transaction if exists
          if (currentTransaction) {
            transactions.push(currentTransaction);
          }
          
          var dateStr = transactionMatch[1]; // "31.12.2025"
          var timeStr = transactionMatch[2]; // "16:40"
          var authCode = transactionMatch[3]; // "966521"
          var category = transactionMatch[4].trim(); // "Перевод СБП" or "Прочие операции"
          var amountStr = transactionMatch[5]; // "1 500,00" or "+46 696,61"
          var balanceStr = transactionMatch[6]; // "96 776,18" (not used, but good to have)
          var processingDate = transactionMatch[7] || ''; // "29.10.2025" (optional)
          var descriptionStart = transactionMatch[8] || ''; // Start of description from same line
          
          // Parse amount
          var amountValue = this._parseAmount_(amountStr);
          
          // Determine transaction type
          // Income indicators: "+" in original category (before cleaning), "зачислен", "пополнение", "возврат", "заработная плата"
          var type = 'expense';
          var categoryLower = category.toLowerCase();
          if (amountValue < 0 || 
              hasPlusInCategory || // "+47", "+350" etc. indicates income
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
            description: descriptionStart ? [descriptionStart] : [], // Start with description from same line
            processingDate: processingDate,
            rawLine: line
          };
        } else {
          // Check if this is a continuation line (has date but no time/code/amount)
          var dateMatch = line.match(datePattern);
          var amountMatch = line.match(amountPattern);
          
          if (currentTransaction) {
            // Check if this line looks like a full transaction line (has date, time, code, amount)
            // If so, it's likely a new transaction, not continuation
            var looksLikeFullTransaction = dateMatch && 
                                          line.match(/\d{2}:\d{2}/) && // Has time
                                          line.match(/\d{6}/) && // Has auth code
                                          amountMatch;
            
            if (looksLikeFullTransaction) {
              // This looks like a new transaction line that didn't match the pattern
              // Save current transaction and try to parse this as new
              transactions.push(currentTransaction);
              currentTransaction = null;
              
              // Try to extract what we can from this line
              var dateStr2 = dateMatch[1];
              var timeMatch = line.match(/(\d{2}:\d{2})/);
              var authCodeMatch = line.match(/(\d{6})/);
              var amountStr2 = amountMatch[1];
              var amountValue2 = this._parseAmount_(amountStr2);
              
              // Extract category (text between auth code and amount)
              var category = '';
              if (authCodeMatch && amountMatch) {
                var categoryStart = authCodeMatch.index + authCodeMatch[0].length;
                var categoryEnd = amountMatch.index;
                if (categoryEnd > categoryStart) {
                  category = line.substring(categoryStart, categoryEnd).trim();
                }
              }
              
              // Determine type
              var type2 = 'expense';
              if (category.indexOf('+') !== -1) {
                type2 = 'income';
              }
              
              currentTransaction = {
                bank: 'sberbank',
                date: dateStr2,
                time: timeMatch ? timeMatch[1] : '',
                authCode: authCodeMatch ? authCodeMatch[1] : '',
                category: category,
                amount: amountValue2,
                type: type2,
                description: [], // Will collect description from next lines
                rawLine: line
              };
            } else if (dateMatch && !amountMatch) {
              // Line with date but no amount - continuation of description
              // Remove date from beginning if present
              var descLine = line;
              if (dateMatch.index === 0) {
                descLine = line.substring(dateMatch[0].length).trim();
              }
              // Skip lines that are just "по карте ****7426" or similar
              if (descLine && descLine.length > 0 && 
                  !descLine.match(/^(по карте|операция по карте|карте)\s*\*\*\*\*\d*$/i)) {
                currentTransaction.description.push(descLine);
              }
            } else if (!dateMatch && !amountMatch) {
              // Line without date or amount - continuation of description
              // Skip lines that are just "по карте ****7426" or similar
              if (line.length > 0 && 
                  !line.match(/^(по карте|операция по карте|карте)\s*\*\*\*\*\d*$/i)) {
                currentTransaction.description.push(line);
              }
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
    
    // For income transactions, build proper description
    // If we have "Прочие операции" in category and "Заработная плата" in description
    if (type === 'income' && 
        rawTransaction.category && 
        rawTransaction.category.indexOf('Прочие операции') !== -1) {
      // Build description: "Прочие операции. Заработная плата. Операция по карте ****7426"
      var descParts = [];
      if (rawTransaction.category) {
        descParts.push(rawTransaction.category);
      }
      if (description && description.indexOf('Заработная плата') !== -1) {
        descParts.push('Заработная плата');
      }
      // Add card operation info if present
      if (description && description.indexOf('Операция по карте') !== -1) {
        var cardMatch = description.match(/Операция по карте[^*]*\*\*\*\*\d+/);
        if (cardMatch) {
          descParts.push(cardMatch[0]);
        }
      } else if (description && description.indexOf('по карте') !== -1) {
        var cardMatch2 = description.match(/по карте[^*]*\*\*\*\*\d+/);
        if (cardMatch2) {
          descParts.push('Операция ' + cardMatch2[0]);
        }
      }
      if (descParts.length > 0) {
        description = descParts.join('. ');
      }
    }
    
    // Extract merchant from description (usually first part before "RUS" or "Операция")
    var merchant = '';
    if (description) {
      // First, check if description contains a full transaction line (date + time + code + amount)
      // If so, extract merchant from the description part, not from the transaction line
      var fullTxMatch = description.match(/(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2})\s+(\d{6})\s+(.+?)\s+([\d\s]+,\d{2})/);
      if (fullTxMatch) {
        // Description contains a full transaction line - extract merchant from the part after it
        var txLineEnd = fullTxMatch.index + fullTxMatch[0].length;
        var descAfterTx = description.substring(txLineEnd).trim();
        if (descAfterTx) {
          description = descAfterTx; // Use only the part after transaction line
        } else {
          // If no description after transaction line, try to extract from category part
          var categoryPart = fullTxMatch[4].trim();
          if (categoryPart && categoryPart.indexOf('Заработная плата') === -1) {
            description = categoryPart;
          } else {
            description = ''; // Skip if it's just salary info
          }
        }
      }
      
      // Remove common patterns that are not merchant names
      var cleanDesc = description
        .replace(/по карте\s+\*\*\*\*\d+/gi, '') // Remove "по карте ****7426"
        .replace(/\*\*\*\*\d+/g, '') // Remove "****7426"
        .replace(/операция по карте/gi, '') // Remove "Операция по карте"
        .replace(/операция/gi, '') // Remove "Операция"
        .trim();
      
      // Skip if description is too short or contains only transaction metadata
      if (cleanDesc.length < 3 || 
          cleanDesc.match(/^\d{2}\.\d{2}\.\d{4}/) || // Starts with date
          cleanDesc.match(/^\d{2}:\d{2}/) || // Starts with time
          cleanDesc.match(/^\d{6}$/)) { // Just auth code
        merchant = '';
      } else {
        // Split by common delimiters
        var parts = cleanDesc.split(/\.|RUS/);
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
          if (match && match[1] && match[1].trim().length >= 3) {
            merchant = match[1].trim();
            merchant = merchant.replace(/[\.\-\s]{2,}/g, ' ').trim();
            merchant = merchant.replace(/^["']|["']$/g, '');
          }
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
