/**
 * Unit tests for Personal Finances project.
 * 
 * To run tests: Execute functions starting with "test_" from Apps Script editor.
 * Or create a test menu item that runs all tests.
 */

/**
 * Run all tests and display results.
 */
function pfRunAllTests() {
  var results = [];
  var passed = 0;
  var failed = 0;
  
  // Test Constants
  results.push(testConstants());
  
  // Test ErrorHandler
  results.push(testErrorHandler());
  
  // Test Deduplication
  results.push(testDeduplicationKeyGeneration());
  
  // Test Sheet Clearing
  results.push(testSheetClearing());
  
  // Test Input Validation
  results.push(testInputValidation());
  
  // Count results
  for (var i = 0; i < results.length; i++) {
    if (results[i].passed) {
      passed += results[i].count || 1;
    } else {
      failed += results[i].count || 1;
    }
  }
  
  // Display results
  var message = 'Тесты завершены:\n';
  message += 'Пройдено: ' + passed + '\n';
  message += 'Провалено: ' + failed + '\n\n';
  
  for (var i = 0; i < results.length; i++) {
    message += results[i].name + ': ' + (results[i].passed ? 'OK' : 'FAILED') + '\n';
    if (results[i].errors && results[i].errors.length > 0) {
      for (var j = 0; j < results[i].errors.length; j++) {
        message += '  - ' + results[i].errors[j] + '\n';
      }
    }
  }
  
  SpreadsheetApp.getUi().alert('Результаты тестов', message, SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log('=== TEST RESULTS ===');
  Logger.log(message);
}

/**
 * Test Constants.js
 */
function testConstants() {
  var errors = [];
  
  try {
    // Test transaction statuses
    if (PF_TRANSACTION_STATUS.OK !== 'ok') errors.push('PF_TRANSACTION_STATUS.OK incorrect');
    if (PF_TRANSACTION_STATUS.DUPLICATE !== 'duplicate') errors.push('PF_TRANSACTION_STATUS.DUPLICATE incorrect');
    if (PF_TRANSACTION_STATUS.NEEDS_REVIEW !== 'needs_review') errors.push('PF_TRANSACTION_STATUS.NEEDS_REVIEW incorrect');
    
    // Test transaction types
    if (PF_TRANSACTION_TYPE.EXPENSE !== 'expense') errors.push('PF_TRANSACTION_TYPE.EXPENSE incorrect');
    if (PF_TRANSACTION_TYPE.INCOME !== 'income') errors.push('PF_TRANSACTION_TYPE.INCOME incorrect');
    if (PF_TRANSACTION_TYPE.TRANSFER !== 'transfer') errors.push('PF_TRANSACTION_TYPE.TRANSFER incorrect');
    
    // Test import sources
    if (PF_IMPORT_SOURCE.MANUAL !== 'manual') errors.push('PF_IMPORT_SOURCE.MANUAL incorrect');
    if (PF_IMPORT_SOURCE.CSV !== 'import:csv') errors.push('PF_IMPORT_SOURCE.CSV incorrect');
    if (PF_IMPORT_SOURCE.SBERBANK !== 'import:sberbank') errors.push('PF_IMPORT_SOURCE.SBERBANK incorrect');
    
    // Test batch size
    if (typeof PF_IMPORT_BATCH_SIZE !== 'number' || PF_IMPORT_BATCH_SIZE !== 200) {
      errors.push('PF_IMPORT_BATCH_SIZE incorrect');
    }
    
    // Test default currency
    if (PF_DEFAULT_CURRENCY !== 'RUB') errors.push('PF_DEFAULT_CURRENCY incorrect');
    
    // Test supported currencies
    if (!Array.isArray(PF_SUPPORTED_CURRENCIES) || PF_SUPPORTED_CURRENCIES.length !== 3) {
      errors.push('PF_SUPPORTED_CURRENCIES incorrect');
    }
    
  } catch (e) {
    errors.push('Exception in testConstants: ' + e.toString());
  }
  
  return {
    name: 'Constants',
    passed: errors.length === 0,
    count: 1,
    errors: errors
  };
}

/**
 * Test ErrorHandler.js
 */
function testErrorHandler() {
  var errors = [];
  
  try {
    // Test log levels
    if (!PF_LOG_LEVEL.DEBUG || !PF_LOG_LEVEL.INFO || !PF_LOG_LEVEL.WARNING || !PF_LOG_LEVEL.ERROR) {
      errors.push('PF_LOG_LEVEL constants missing');
    }
    
    // Test error response creation
    var errorResp = pfCreateErrorResponse_('Test error');
    if (errorResp.success !== false) errors.push('pfCreateErrorResponse_ success field incorrect: expected false, got ' + errorResp.success);
    if (errorResp.message !== 'Test error') errors.push('pfCreateErrorResponse_ message incorrect');
    
    // Test success response creation
    var successResp = pfCreateSuccessResponse_('Test success', { data: 123 });
    if (!successResp.success || successResp.success !== true) errors.push('pfCreateSuccessResponse_ success field incorrect');
    if (successResp.message !== 'Test success') errors.push('pfCreateSuccessResponse_ message incorrect');
    if (successResp.data !== 123) errors.push('pfCreateSuccessResponse_ data field incorrect');
    
    // Test error handling (should not throw)
    try {
      var handled = pfHandleError_(new Error('Test'), 'testContext', 'User message');
      if (handled.success !== false) {
        errors.push('pfHandleError_ success field incorrect: expected false, got ' + handled.success);
      }
      if (handled.message !== 'User message') {
        errors.push('pfHandleError_ message incorrect: expected "User message", got "' + handled.message + '"');
      }
    } catch (e) {
      errors.push('pfHandleError_ threw exception: ' + e.toString());
    }
    
  } catch (e) {
    errors.push('Exception in testErrorHandler: ' + e.toString());
  }
  
  return {
    name: 'ErrorHandler',
    passed: errors.length === 0,
    count: 1,
    errors: errors
  };
}

/**
 * Test deduplication key generation
 */
function testDeduplicationKeyGeneration() {
  var errors = [];
  
  try {
    // Test with TransactionDTO (with sourceId)
    var tx1 = {
      date: new Date('2025-01-15'),
      type: PF_TRANSACTION_TYPE.EXPENSE,
      account: 'Test Account',
      amount: 1000,
      currency: 'RUB',
      source: PF_IMPORT_SOURCE.CSV,
      sourceId: 'test123'
    };
    
    var key1 = pfGenerateDedupeKey_(tx1);
    if (key1 !== PF_IMPORT_SOURCE.CSV + ':test123') {
      errors.push('Dedupe key with sourceId incorrect: ' + key1);
    }
    
    // Test with TransactionDTO (without sourceId - should use hash)
    var tx2 = {
      date: new Date('2025-01-15'),
      type: PF_TRANSACTION_TYPE.EXPENSE,
      account: 'Test Account',
      amount: 1000,
      currency: 'RUB',
      source: PF_IMPORT_SOURCE.CSV
      // no sourceId
    };
    
    var key2 = pfGenerateDedupeKey_(tx2);
    if (!key2 || key2.indexOf(PF_IMPORT_SOURCE.CSV + ':') !== 0) {
      errors.push('Dedupe key without sourceId format incorrect: ' + key2);
    }
    
    // Test with sheet row
    var row = [
      new Date('2025-01-15'), // Date
      PF_TRANSACTION_TYPE.EXPENSE, // Type
      'Test Account', // Account
      '', // AccountTo
      1000, // Amount
      'RUB', // Currency
      '', // Category
      '', // Subcategory
      '', // Merchant
      '', // Description
      '', // Tags
      PF_IMPORT_SOURCE.CSV, // Source
      'test456', // SourceId
      PF_TRANSACTION_STATUS.OK // Status
    ];
    
    var schema = PF_TRANSACTIONS_SCHEMA;
    var sourceCol = pfColumnIndex_(schema, 'Source');
    var sourceIdCol = pfColumnIndex_(schema, 'SourceId');
    var dateCol = pfColumnIndex_(schema, 'Date');
    var accountCol = pfColumnIndex_(schema, 'Account');
    var amountCol = pfColumnIndex_(schema, 'Amount');
    var typeCol = pfColumnIndex_(schema, 'Type');
    
    var key3 = pfGenerateDedupeKey_(row, {
      sourceCol: sourceCol,
      sourceIdCol: sourceIdCol,
      dateCol: dateCol,
      accountCol: accountCol,
      amountCol: amountCol,
      typeCol: typeCol
    });
    
    if (key3 !== PF_IMPORT_SOURCE.CSV + ':test456') {
      errors.push('Dedupe key from row with sourceId incorrect: ' + key3);
    }
    
    // Test consistency: same transaction should generate same key
    var tx3a = {
      date: new Date('2025-01-20'),
      type: PF_TRANSACTION_TYPE.INCOME,
      account: 'Account2',
      amount: 5000,
      source: PF_IMPORT_SOURCE.SBERBANK
    };
    
    var tx3b = {
      date: new Date('2025-01-20'),
      type: PF_TRANSACTION_TYPE.INCOME,
      account: 'Account2',
      amount: 5000,
      source: PF_IMPORT_SOURCE.SBERBANK
    };
    
    var key3a = pfGenerateDedupeKey_(tx3a);
    var key3b = pfGenerateDedupeKey_(tx3b);
    
    if (key3a !== key3b) {
      errors.push('Same transaction generated different keys: ' + key3a + ' vs ' + key3b);
    }
    
  } catch (e) {
    errors.push('Exception in testDeduplicationKeyGeneration: ' + e.toString());
    Logger.log('Test error: ' + e.toString());
    Logger.log('Stack: ' + (e.stack || 'No stack'));
  }
  
  return {
    name: 'Deduplication Key Generation',
    passed: errors.length === 0,
    count: 1,
    errors: errors
  };
}

/**
 * Test sheet clearing function
 */
function testSheetClearing() {
  var errors = [];
  
  try {
    // Test with invalid parameters (should handle gracefully)
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var testSheet = ss.insertSheet('TestSheet_' + new Date().getTime());
    
    try {
      // Test with invalid startRow
      var result1 = pfClearSheetRows_(testSheet, 0, 5);
      if (result1 !== false) {
        errors.push('pfClearSheetRows_ should return false for invalid startRow');
      }
      
      // Test with invalid numRows
      var result2 = pfClearSheetRows_(testSheet, 2, -1);
      if (result2 !== false) {
        errors.push('pfClearSheetRows_ should return false for invalid numRows');
      }
      
      // Test with null sheet
      var result3 = pfClearSheetRows_(null, 2, 5);
      if (result3 !== false) {
        errors.push('pfClearSheetRows_ should return false for null sheet');
      }
      
      // Test with empty sheet (should succeed)
      var result4 = pfClearSheetRows_(testSheet, 2, 0);
      // Should handle gracefully (no rows to clear)
      
    } finally {
      // Cleanup
      ss.deleteSheet(testSheet);
    }
    
  } catch (e) {
    errors.push('Exception in testSheetClearing: ' + e.toString());
  }
  
  return {
    name: 'Sheet Clearing',
    passed: errors.length === 0,
    count: 1,
    errors: errors
  };
}

/**
 * Test input validation in public functions
 */
function testInputValidation() {
  var errors = [];
  
  try {
    // Test pfDetectFileFormat validation
    try {
      pfDetectFileFormat(null);
      errors.push('pfDetectFileFormat should throw for null input');
    } catch (e) {
      // Expected
    }
    
    try {
      pfDetectFileFormat('');
      errors.push('pfDetectFileFormat should throw for empty string');
    } catch (e) {
      // Expected
    }
    
    // Test pfParseFileContent validation
    try {
      pfParseFileContent('test', null);
      errors.push('pfParseFileContent should throw for invalid importerType');
    } catch (e) {
      // Expected
    }
    
    try {
      pfParseFileContent('test', 'invalid');
      errors.push('pfParseFileContent should throw for invalid importerType');
    } catch (e) {
      // Expected
    }
    
    // Test pfWritePreview validation
    try {
      pfWritePreview(null);
      errors.push('pfWritePreview should throw for null input');
    } catch (e) {
      // Expected
    }
    
    try {
      pfWritePreview('[]');
      errors.push('pfWritePreview should throw for empty array');
    } catch (e) {
      // Expected
    }
    
    // Test pfFindDuplicateTransaction validation
    var result1 = pfFindDuplicateTransaction(null);
    if (result1.found !== false) {
      errors.push('pfFindDuplicateTransaction should return found=false for null');
    }
    
    var result2 = pfFindDuplicateTransaction('');
    if (result2.found !== false) {
      errors.push('pfFindDuplicateTransaction should return found=false for empty string');
    }
    
    var result3 = pfFindDuplicateTransaction('invalid-format');
    if (result3.found !== false) {
      errors.push('pfFindDuplicateTransaction should return found=false for invalid format');
    }
    
  } catch (e) {
    errors.push('Exception in testInputValidation: ' + e.toString());
  }
  
  return {
    name: 'Input Validation',
    passed: errors.length === 0,
    count: 1,
    errors: errors
  };
}

/**
 * Test PDF Sberbank parser with real examples
 */
function testPdfSberbankParser() {
  var errors = [];
  
  try {
    // Test cases from user examples
    var testCases = [
      {
        input: '29.10.2025 18:36 574763 Супермаркеты 204,98 97 005,99 29.10.2025 PYATEROCHKA 20477 Shakhty RUS. Операция по',
        nextLine: 'карте ****7426',
        expected: {
          date: '29.10.2025',
          time: '18:36',
          authCode: '574763',
          category: 'Супермаркеты',
          amount: '204,98',
          balance: '97 005,99',
          processingDate: '29.10.2025',
          description: 'PYATEROCHKA 20477 Shakhty RUS. Операция по карте ****7426'
        }
      },
      {
        input: '29.10.2025 18:26 571125 Супермаркеты 229,44 97 210,97 29.10.2025 CH61039 Shakhty RUS. Операция по карте ****7426',
        nextLine: null,
        expected: {
          date: '29.10.2025',
          time: '18:26',
          authCode: '571125',
          category: 'Супермаркеты',
          amount: '229,44',
          balance: '97 210,97',
          processingDate: '29.10.2025',
          description: 'CH61039 Shakhty RUS. Операция по карте ****7426'
        }
      },
      {
        input: '29.10.2025 10:16 938498 Прочие операции +46 696,61 97 440,41 29.10.2025 Заработная плата. Операция по карте ****7426',
        nextLine: null,
        expected: {
          date: '29.10.2025',
          time: '10:16',
          authCode: '938498',
          category: 'Прочие операции',
          amount: '+46 696,61',
          balance: '97 440,41',
          processingDate: '29.10.2025',
          description: 'Заработная плата. Операция по карте ****7426'
        }
      },
      {
        input: '27.10.2025 19:34 201171 Супермаркеты 449,44 50 743,80 27.10.2025 PYATEROCHKA 20477 Shakhty RUS. Операция по',
        nextLine: 'карте ****7426',
        expected: {
          date: '27.10.2025',
          time: '19:34',
          authCode: '201171',
          category: 'Супермаркеты',
          amount: '449,44',
          balance: '50 743,80',
          processingDate: '27.10.2025',
          description: 'PYATEROCHKA 20477 Shakhty RUS. Операция по карте ****7426'
        }
      },
      {
        input: '27.10.2025 19:24 401568 Супермаркеты 1 606,00 51 193,24 27.10.2025 CH61039 Shakhty RUS. Операция по карте ****7426',
        nextLine: null,
        expected: {
          date: '27.10.2025',
          time: '19:24',
          authCode: '401568',
          category: 'Супермаркеты',
          amount: '1 606,00',
          balance: '51 193,24',
          processingDate: '27.10.2025',
          description: 'CH61039 Shakhty RUS. Операция по карте ****7426'
        }
      }
    ];
    
    // Test parser
    for (var i = 0; i < testCases.length; i++) {
      var testCase = testCases[i];
      var lines = [testCase.input];
      if (testCase.nextLine) {
        lines.push(testCase.nextLine);
      }
      
      // Simulate parsing
      var text = lines.join('\n');
      var parsed = PF_PDF_SBERBANK_PARSER.parse(text, {});
      
      if (!parsed || parsed.length === 0) {
        errors.push('Test case ' + (i + 1) + ': No transactions parsed');
        continue;
      }
      
      var transaction = parsed[0];
      var expected = testCase.expected;
      
      // Check each field
      if (transaction.date !== expected.date) {
        errors.push('Test case ' + (i + 1) + ': date mismatch. Expected: "' + expected.date + '", got: "' + transaction.date + '"');
      }
      if (transaction.time !== expected.time) {
        errors.push('Test case ' + (i + 1) + ': time mismatch. Expected: "' + expected.time + '", got: "' + transaction.time + '"');
      }
      if (transaction.authCode !== expected.authCode) {
        errors.push('Test case ' + (i + 1) + ': authCode mismatch. Expected: "' + expected.authCode + '", got: "' + transaction.authCode + '"');
      }
      if (transaction.category !== expected.category) {
        errors.push('Test case ' + (i + 1) + ': category mismatch. Expected: "' + expected.category + '", got: "' + transaction.category + '"');
      }
      if (transaction.amount !== expected.amount) {
        errors.push('Test case ' + (i + 1) + ': amount mismatch. Expected: "' + expected.amount + '", got: "' + transaction.amount + '"');
      }
      
      // Check description (may be in description array)
      var description = Array.isArray(transaction.description) 
        ? transaction.description.join(' ') 
        : (transaction.description || '');
      if (description.indexOf(expected.description) === -1 && expected.description.indexOf(description) === -1) {
        errors.push('Test case ' + (i + 1) + ': description mismatch. Expected to contain: "' + expected.description + '", got: "' + description + '"');
      }
    }
    
  } catch (e) {
    errors.push('Exception in testPdfSberbankParser: ' + e.toString());
    Logger.log('Test exception: ' + e.toString());
    Logger.log('Stack: ' + (e.stack || 'No stack'));
  }
  
  return {
    name: 'PDF Sberbank Parser',
    passed: errors.length === 0,
    count: testCases ? testCases.length : 0,
    errors: errors
  };
}
