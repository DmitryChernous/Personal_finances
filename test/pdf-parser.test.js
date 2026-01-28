/**
 * Local tests for PDF Sberbank parser
 * Run with: node test/pdf-parser.test.js
 */

// Mock Google Apps Script API
global.Logger = {
  log: function(message) {
    console.log('[LOG]', message);
  }
};

// Mock other dependencies if needed
global.PF_LOG_LEVEL = {
  DEBUG: 'DEBUG',
  INFO: 'INFO',
  WARNING: 'WARNING',
  ERROR: 'ERROR'
};

global.PF_CURRENT_LOG_LEVEL = global.PF_LOG_LEVEL.WARNING;

// Simple mock for pfLogWarning_
global.pfLogWarning_ = function(message, context) {
  console.warn('[WARNING]', context, message);
};

// Load the parser (we'll need to adapt it for Node.js)
// For now, let's create a test version that doesn't depend on Apps Script

/**
 * Test cases from user examples
 */
const testCases = [
  {
    name: 'Transaction 1: Supermarket with multi-line description',
    input: [
      '29.10.2025 18:36 574763 Супермаркеты 204,98 97 005,99 29.10.2025 PYATEROCHKA 20477 Shakhty RUS. Операция по',
      'карте ****7426'
    ],
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
    name: 'Transaction 2: Supermarket single line',
    input: [
      '29.10.2025 18:26 571125 Супермаркеты 229,44 97 210,97 29.10.2025 CH61039 Shakhty RUS. Операция по карте ****7426'
    ],
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
    name: 'Transaction 3: Income with +amount',
    input: [
      '29.10.2025 10:16 938498 Прочие операции +46 696,61 97 440,41 29.10.2025 Заработная плата. Операция по карте ****7426'
    ],
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
    name: 'Transaction 4: Supermarket multi-line',
    input: [
      '27.10.2025 19:34 201171 Супермаркеты 449,44 50 743,80 27.10.2025 PYATEROCHKA 20477 Shakhty RUS. Операция по',
      'карте ****7426'
    ],
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
    name: 'Transaction 5: Supermarket with large amount',
    input: [
      '27.10.2025 19:24 401568 Супермаркеты 1 606,00 51 193,24 27.10.2025 CH61039 Shakhty RUS. Операция по карте ****7426'
    ],
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
  },
  {
    name: 'Transaction 6: 08.10.2025 supermarket with multi-line description',
    input: [
      '08.10.2025 13:31 792469 Супермаркеты 1 478,77 30 849,34 08.10.2025 PYATEROCHKA 20477 Shakhty RUS. Операция по',
      'карте ****7426'
    ],
    expected: {
      date: '08.10.2025',
      time: '13:31',
      authCode: '792469',
      category: 'Супермаркеты',
      amount: '1 478,77',
      balance: '30 849,34',
      processingDate: '08.10.2025',
      description: 'PYATEROCHKA 20477 Shakhty RUS. Операция по карте ****7426'
    }
  },
  {
    name: 'Transaction 7: 06.10.2025 SBP transfer with multi-line description',
    input: [
      '06.10.2025 21:59 986237 Перевод СБП 500,00 32 328,11 06.10.2025 Перевод для Ч. Дмитрий Вячеславович. Операция',
      'по карте ****7426'
    ],
    expected: {
      date: '06.10.2025',
      time: '21:59',
      authCode: '986237',
      category: 'Перевод СБП',
      amount: '500,00',
      balance: '32 328,11',
      processingDate: '06.10.2025',
      description: 'Перевод для Ч. Дмитрий Вячеславович. Операция по карте ****7426'
    }
  }
];

/**
 * Simple parser implementation for testing
 * This is a simplified version that focuses on parsing logic
 */
function parseTransactionLine(line, nextLine) {
  // Rule 1: New transaction starts with date dd.mm.yyyy
  const datePattern = /^(\d{2}\.\d{2}\.\d{4})/;
  const dateMatch = line.match(datePattern);
  
  if (!dateMatch) {
    console.log('  DEBUG: No date match');
    return null;
  }
  
  const dateStr = dateMatch[1];
  
  // Rule 2: Extract time - next 5 characters after date
  const timeMatch = line.substring(10).trim().match(/^(\d{2}:\d{2})/);
  if (!timeMatch) {
    console.log('  DEBUG: No time match, substring:', line.substring(10).trim());
    return null;
  }
  
  const timeStr = timeMatch[1];
  
  // Rule 3: Extract authorization code - next 6 digits after time
  // Date is 10 chars, then space, then time is 5 chars, then space, then 6 digits
  const afterDateAndTime = line.substring(10).trim(); // Skip date (10 chars)
  const afterTime = afterDateAndTime.substring(timeStr.length).trim(); // Skip time (5 chars)
  const authCodeMatch = afterTime.match(/^(\d{6})/);
  if (!authCodeMatch) {
    return null;
  }
  
  const authCode = authCodeMatch[1];
  
  // Rule 4: Extract category and amounts
  const afterAuthCode = afterTime.substring(6).trim();
  
  // Find amount pattern: starts with digit or "+", ends with ,XX
  const amountPattern = /(\+?[\d\s]+,\d{2})/g;
  const amountMatches = [];
  let match;
  
  // Reset regex lastIndex
  amountPattern.lastIndex = 0;
  while ((match = amountPattern.exec(afterAuthCode)) !== null) {
    amountMatches.push({
      value: match[1],
      index: match.index
    });
  }
  
  // Combine partial matches (e.g., "+46" and "696,61")
  // Only combine if first match is incomplete (like "+46" without comma) and next is right after
  const combinedMatches = [];
  for (let m = 0; m < amountMatches.length; m++) {
    const current = amountMatches[m];
    // Check if this looks like a partial match: starts with "+" but doesn't have a comma yet
    // AND next match is very close (0-2 chars gap) AND next doesn't start with "+"
    if (current.value.startsWith('+') && !current.value.includes(',') && m + 1 < amountMatches.length) {
      const next = amountMatches[m + 1];
      const gap = next.index - (current.index + current.value.length);
      // Only combine if gap is very small (0-2 chars, likely just a space) and next doesn't start with "+"
      if (gap >= 0 && gap <= 2 && !next.value.startsWith('+')) {
        combinedMatches.push({
          value: current.value + afterAuthCode.substring(current.index + current.value.length, next.index) + next.value,
          index: current.index
        });
        m++;
        continue;
      }
    }
    combinedMatches.push(current);
  }
  
  if (combinedMatches.length < 2) {
    return null;
  }
  
  const amountStr = combinedMatches[0].value.trim();
  const balanceStr = combinedMatches[1].value.trim();
  
  // Category is everything between auth code and first amount
  const categoryEndIndex = combinedMatches[0].index;
  let categoryPart = afterAuthCode.substring(0, categoryEndIndex).trim();
  
  // Extract processing date and description after balance
  const afterBalance = afterAuthCode.substring(combinedMatches[1].index + combinedMatches[1].value.length).trim();
  let processingDate = '';
  let descriptionStart = '';
  
  const processingDateMatch = afterBalance.match(/^(\d{2}\.\d{2}\.\d{4})\s+(.+)$/);
  if (processingDateMatch) {
    processingDate = processingDateMatch[1];
    descriptionStart = processingDateMatch[2].trim();
  } else {
    descriptionStart = afterBalance;
  }
  
  // Combine with next line if it's a continuation
  let description = descriptionStart;
  if (nextLine && !nextLine.match(/^\d{2}\.\d{2}\.\d{4}/)) {
    description = (descriptionStart + ' ' + nextLine.trim()).trim();
  }
  
  return {
    date: dateStr,
    time: timeStr,
    authCode: authCode,
    category: categoryPart,
    amount: amountStr,
    balance: balanceStr,
    processingDate: processingDate,
    description: description
  };
}

/**
 * Run tests
 */
function runTests() {
  console.log('=== PDF Parser Tests ===\n');
  
  let passed = 0;
  let failed = 0;
  const errors = [];
  
  for (let i = 0; i < testCases.length; i++) {
    const testCase = testCases[i];
    console.log(`Test ${i + 1}: ${testCase.name}`);
    
    try {
      const result = parseTransactionLine(testCase.input[0], testCase.input[1] || null);
      
      if (!result) {
        failed++;
        errors.push(`Test ${i + 1}: Parser returned null`);
        console.log('  ❌ FAILED: Parser returned null\n');
        continue;
      }
      
      const expected = testCase.expected;
      const testErrors = [];
      
      // Check each field
      if (result.date !== expected.date) {
        testErrors.push(`date: expected "${expected.date}", got "${result.date}"`);
      }
      if (result.time !== expected.time) {
        testErrors.push(`time: expected "${expected.time}", got "${result.time}"`);
      }
      if (result.authCode !== expected.authCode) {
        testErrors.push(`authCode: expected "${expected.authCode}", got "${result.authCode}"`);
      }
      if (result.category !== expected.category) {
        testErrors.push(`category: expected "${expected.category}", got "${result.category}"`);
      }
      if (result.amount !== expected.amount) {
        testErrors.push(`amount: expected "${expected.amount}", got "${result.amount}"`);
      }
      if (result.balance !== expected.balance) {
        testErrors.push(`balance: expected "${expected.balance}", got "${result.balance}"`);
      }
      if (result.processingDate !== expected.processingDate) {
        testErrors.push(`processingDate: expected "${expected.processingDate}", got "${result.processingDate}"`);
      }
      if (result.description !== expected.description) {
        testErrors.push(`description: expected "${expected.description}", got "${result.description}"`);
      }
      
      if (testErrors.length === 0) {
        passed++;
        console.log('  ✅ PASSED\n');
      } else {
        failed++;
        errors.push(`Test ${i + 1}: ${testErrors.join('; ')}`);
        console.log('  ❌ FAILED:');
        testErrors.forEach(err => console.log(`    - ${err}`));
        console.log('');
      }
    } catch (e) {
      failed++;
      errors.push(`Test ${i + 1}: Exception - ${e.message}`);
      console.log(`  ❌ FAILED: Exception - ${e.message}\n`);
      console.error(e);
    }
  }
  
  console.log('=== Test Results ===');
  console.log(`Passed: ${passed}`);
  console.log(`Failed: ${failed}`);
  
  if (errors.length > 0) {
    console.log('\nErrors:');
    errors.forEach(err => console.log(`  - ${err}`));
  }
  
  process.exit(failed > 0 ? 1 : 0);
}

// Run tests
runTests();
