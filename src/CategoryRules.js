/**
 * Category Rules module.
 * Handles automatic categorization of transactions based on rules:
 * - Sheet initialization
 * - Rule matching logic
 * - Auto-categorization for transactions
 */

/**
 * Initialize CategoryRules sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function pfInitializeCategoryRules_(ss) {
  var sheet = pfFindOrCreateSheetByKey_(ss, PF_SHEET_KEYS.CATEGORY_RULES);
  
  // Set headers
  pfEnsureHeaderRow_(sheet, PF_CATEGORY_RULES_SCHEMA);
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Apply filter
  pfEnsureFilter_(sheet, PF_CATEGORY_RULES_SCHEMA.columns.length);
  
  // Format Priority column as number
  var priorityCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'Priority');
  if (priorityCol) {
    sheet.getRange(2, priorityCol, sheet.getMaxRows() - 1, 1).setNumberFormat('0');
  }
  
  // Set data validations
  var patternTypeCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'PatternType');
  if (patternTypeCol) {
    var patternTypeRange = sheet.getRange(2, patternTypeCol, sheet.getMaxRows() - 1, 1);
    var patternTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        pfT_('pattern_type.contains'),
        pfT_('pattern_type.startsWith'),
        pfT_('pattern_type.endsWith'),
        pfT_('pattern_type.regex'),
        pfT_('pattern_type.exact')
      ], true)
      .build();
    patternTypeRange.setDataValidation(patternTypeRule);
  }
  
  var applyToCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'ApplyTo');
  if (applyToCol) {
    var applyToRange = sheet.getRange(2, applyToCol, sheet.getMaxRows() - 1, 1);
    var applyToRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        pfT_('rule_apply_to.merchant'),
        pfT_('rule_apply_to.description'),
        pfT_('rule_apply_to.both')
      ], true)
      .build();
    applyToRange.setDataValidation(applyToRule);
  }
  
  // Category validation (from Categories sheet)
  var categoryCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'Category');
  if (categoryCol) {
    var categoriesSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.CATEGORIES);
    if (categoriesSheet) {
      var categoriesRange = categoriesSheet.getRange('A2:A');
      var categoryRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(categoriesRange, true)
        .build();
      sheet.getRange(2, categoryCol, sheet.getMaxRows() - 1, 1).setDataValidation(categoryRule);
    }
  }
  
  pfLogInfo_('Initialized CategoryRules sheet', 'pfInitializeCategoryRules_');
}

/**
 * Get all active category rules from sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] Optional spreadsheet
 * @returns {Array<Object>} Array of rule objects
 */
function pfGetAllCategoryRules_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    pfLogWarning_('Cannot get spreadsheet in pfGetAllCategoryRules_', 'pfGetAllCategoryRules_');
    return [];
  }
  
  var sheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.CATEGORY_RULES);
  if (!sheet) {
    return []; // Sheet doesn't exist yet
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return []; // Only header or empty sheet
  }
  
  var rules = [];
  var ruleNameCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'RuleName');
  var patternCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'Pattern');
  var patternTypeCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'PatternType');
  var categoryCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'Category');
  var subcategoryCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'Subcategory');
  var priorityCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'Priority');
  var activeCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'Active');
  var applyToCol = pfColumnIndex_(PF_CATEGORY_RULES_SCHEMA, 'ApplyTo');
  
  if (!ruleNameCol || !patternCol || !patternTypeCol || !categoryCol) {
    pfLogWarning_('Missing required columns in CategoryRules schema', 'pfGetAllCategoryRules_');
    return [];
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, PF_CATEGORY_RULES_SCHEMA.columns.length).getValues();
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    // Check if active (default: true if empty)
    var active = true;
    if (activeCol) {
      var activeValue = row[activeCol - 1];
      if (activeValue === false || activeValue === 'false' || activeValue === 'FALSE' ||
          (typeof activeValue === 'string' && activeValue.trim().toLowerCase() === 'false')) {
        active = false;
      }
    }
    
    if (!active) {
      continue; // Skip inactive rules
    }
    
    // Get pattern type (convert from localized string to key)
    var patternTypeLocalized = String(row[patternTypeCol - 1] || '').trim();
    var patternType = '';
    if (patternTypeLocalized === pfT_('pattern_type.contains')) {
      patternType = PF_PATTERN_TYPE.CONTAINS;
    } else if (patternTypeLocalized === pfT_('pattern_type.startsWith')) {
      patternType = PF_PATTERN_TYPE.STARTS_WITH;
    } else if (patternTypeLocalized === pfT_('pattern_type.endsWith')) {
      patternType = PF_PATTERN_TYPE.ENDS_WITH;
    } else if (patternTypeLocalized === pfT_('pattern_type.regex')) {
      patternType = PF_PATTERN_TYPE.REGEX;
    } else if (patternTypeLocalized === pfT_('pattern_type.exact')) {
      patternType = PF_PATTERN_TYPE.EXACT;
    } else {
      // Try direct match (in case stored as key)
      patternType = patternTypeLocalized;
    }
    
    // Get applyTo (convert from localized string to key)
    var applyToLocalized = String(row[applyToCol - 1] || '').trim();
    var applyTo = PF_RULE_APPLY_TO.BOTH; // Default
    if (applyToLocalized === pfT_('rule_apply_to.merchant')) {
      applyTo = PF_RULE_APPLY_TO.MERCHANT;
    } else if (applyToLocalized === pfT_('rule_apply_to.description')) {
      applyTo = PF_RULE_APPLY_TO.DESCRIPTION;
    } else if (applyToLocalized === pfT_('rule_apply_to.both')) {
      applyTo = PF_RULE_APPLY_TO.BOTH;
    } else if (applyToLocalized === PF_RULE_APPLY_TO.MERCHANT || 
               applyToLocalized === PF_RULE_APPLY_TO.DESCRIPTION || 
               applyToLocalized === PF_RULE_APPLY_TO.BOTH) {
      applyTo = applyToLocalized; // Already a key
    }
    
    var rule = {
      ruleName: String(row[ruleNameCol - 1] || '').trim(),
      pattern: String(row[patternCol - 1] || '').trim(),
      patternType: patternType,
      category: String(row[categoryCol - 1] || '').trim(),
      subcategory: subcategoryCol ? String(row[subcategoryCol - 1] || '').trim() : '',
      priority: priorityCol ? (Number(row[priorityCol - 1]) || 0) : 0,
      active: active,
      applyTo: applyTo
    };
    
    // Skip rules with empty required fields
    if (!rule.ruleName || !rule.pattern || !rule.patternType || !rule.category) {
      continue;
    }
    
    rules.push(rule);
  }
  
  return rules;
}

/**
 * Match a category rule against merchant and description.
 * @param {string} merchant - Merchant name
 * @param {string} description - Transaction description
 * @param {Array<Object>} rules - Array of rule objects
 * @returns {Object|null} First matching rule or null
 */
function pfMatchCategoryRule_(merchant, description, rules) {
  if (!rules || rules.length === 0) {
    return null;
  }
  
  // Normalize inputs
  merchant = (merchant || '').trim();
  description = (description || '').trim();
  
  // Sort rules by priority (descending)
  var sortedRules = rules.slice().sort(function(a, b) {
    return (b.priority || 0) - (a.priority || 0);
  });
  
  // Try each rule
  for (var i = 0; i < sortedRules.length; i++) {
    var rule = sortedRules[i];
    
    if (!rule.active) {
      continue;
    }
    
    // Determine which fields to check
    var checkMerchant = (rule.applyTo === PF_RULE_APPLY_TO.MERCHANT || rule.applyTo === PF_RULE_APPLY_TO.BOTH);
    var checkDescription = (rule.applyTo === PF_RULE_APPLY_TO.DESCRIPTION || rule.applyTo === PF_RULE_APPLY_TO.BOTH);
    
    // Try to match pattern
    var matched = false;
    
    if (checkMerchant && merchant) {
      matched = pfMatchPattern_(merchant, rule.pattern, rule.patternType);
      if (matched) {
        return rule;
      }
    }
    
    if (checkDescription && description) {
      matched = pfMatchPattern_(description, rule.pattern, rule.patternType);
      if (matched) {
        return rule;
      }
    }
  }
  
  return null;
}

/**
 * Match a pattern against a field value.
 * @private
 * @param {string} field - Field value to check
 * @param {string} pattern - Pattern to match
 * @param {string} patternType - Type of pattern (from PF_PATTERN_TYPE)
 * @returns {boolean} True if matches
 */
function pfMatchPattern_(field, pattern, patternType) {
  if (!field || !pattern) {
    return false;
  }
  
  field = String(field).trim();
  pattern = String(pattern).trim();
  
  if (patternType === PF_PATTERN_TYPE.CONTAINS) {
    return field.toLowerCase().indexOf(pattern.toLowerCase()) !== -1;
  } else if (patternType === PF_PATTERN_TYPE.STARTS_WITH) {
    return field.toLowerCase().indexOf(pattern.toLowerCase()) === 0;
  } else if (patternType === PF_PATTERN_TYPE.ENDS_WITH) {
    var lowerField = field.toLowerCase();
    var lowerPattern = pattern.toLowerCase();
    return lowerField.length >= lowerPattern.length && 
           lowerField.indexOf(lowerPattern, lowerField.length - lowerPattern.length) !== -1;
  } else if (patternType === PF_PATTERN_TYPE.EXACT) {
    return field.toLowerCase() === pattern.toLowerCase();
  } else if (patternType === PF_PATTERN_TYPE.REGEX) {
    try {
      var regex = new RegExp(pattern, 'i');
      return regex.test(field);
    } catch (e) {
      pfLogWarning_('Invalid regex pattern: ' + pattern + ', error: ' + e.toString(), 'pfMatchPattern_');
      return false;
    }
  }
  
  return false;
}

/**
 * Apply category rules to a transaction.
 * @param {Object} transaction - TransactionDTO or transaction object
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] Optional spreadsheet
 * @returns {Object} Transaction with updated category/subcategory
 */
function pfApplyCategoryRules_(transaction, ss) {
  if (!transaction) {
    return transaction;
  }
  
  // Don't overwrite existing category
  if (transaction.category && String(transaction.category).trim() !== '') {
    return transaction;
  }
  
  // Get rules
  var rules = pfGetAllCategoryRules_(ss);
  if (rules.length === 0) {
    return transaction;
  }
  
  // Match rule
  var merchant = (transaction.merchant || '').trim();
  var description = (transaction.description || '').trim();
  var matchedRule = pfMatchCategoryRule_(merchant, description, rules);
  
  if (matchedRule) {
    transaction.category = matchedRule.category;
    if (matchedRule.subcategory) {
      transaction.subcategory = matchedRule.subcategory;
    }
    pfLogDebug_('Applied rule "' + matchedRule.ruleName + '" to transaction, category: ' + matchedRule.category, 'pfApplyCategoryRules_');
  }
  
  return transaction;
}

/**
 * Auto-categorize a single transaction (public function).
 * @param {Object} transaction - Transaction object
 * @returns {Object} Transaction with category
 */
function pfAutoCategorizeTransaction_(transaction) {
  return pfApplyCategoryRules_(transaction);
}

/**
 * Apply auto-categorization to all transactions without category.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] Optional spreadsheet
 * @returns {Object} Statistics: { processed: number, categorized: number }
 */
function pfApplyAutoCategorizationToAll_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    pfLogError_('Cannot get spreadsheet in pfApplyAutoCategorizationToAll_', 'pfApplyAutoCategorizationToAll_', PF_LOG_LEVEL.ERROR);
    return { processed: 0, categorized: 0 };
  }
  
  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet) {
    pfLogWarning_('Transactions sheet not found', 'pfApplyAutoCategorizationToAll_');
    return { processed: 0, categorized: 0 };
  }
  
  var lastRow = txSheet.getLastRow();
  if (lastRow <= 1) {
    return { processed: 0, categorized: 0 };
  }
  
  var stats = { processed: 0, categorized: 0 };
  var categoryCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Category');
  var subcategoryCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Subcategory');
  var merchantCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Merchant');
  var descriptionCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Description');
  
  if (!categoryCol || !merchantCol || !descriptionCol) {
    pfLogWarning_('Missing required columns in Transactions schema', 'pfApplyAutoCategorizationToAll_');
    return stats;
  }
  
  // Get all rules once
  var rules = pfGetAllCategoryRules_(ss);
  if (rules.length === 0) {
    pfLogInfo_('No category rules found', 'pfApplyAutoCategorizationToAll_');
    return stats;
  }
  
  // Read all data
  var data = txSheet.getRange(2, 1, lastRow - 1, PF_TRANSACTIONS_SCHEMA.columns.length).getValues();
  var updates = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var category = String(row[categoryCol - 1] || '').trim();
    
    // Skip if category already set
    if (category !== '') {
      continue;
    }
    
    stats.processed++;
    
    // Get merchant and description
    var merchant = String(row[merchantCol - 1] || '').trim();
    var description = String(row[descriptionCol - 1] || '').trim();
    
    // Match rule
    var matchedRule = pfMatchCategoryRule_(merchant, description, rules);
    
    if (matchedRule) {
      updates.push({
        row: i + 2, // 1-based row number (header is row 1)
        category: matchedRule.category,
        subcategory: matchedRule.subcategory || ''
      });
      stats.categorized++;
    }
  }
  
  // Batch update
  if (updates.length > 0) {
    for (var j = 0; j < updates.length; j++) {
      var update = updates[j];
      txSheet.getRange(update.row, categoryCol).setValue(update.category);
      if (update.subcategory && subcategoryCol) {
        txSheet.getRange(update.row, subcategoryCol).setValue(update.subcategory);
      }
    }
    pfLogInfo_('Applied auto-categorization to ' + stats.categorized + ' transactions', 'pfApplyAutoCategorizationToAll_');
  }
  
  return stats;
}

/**
 * Public function to apply auto-categorization to all transactions.
 * Shows user message with results.
 */
function pfApplyAutoCategorizationToAll() {
  try {
    var stats = pfApplyAutoCategorizationToAll_();
    var message = 'Обработано транзакций: ' + stats.processed + '\n' +
                  'Категоризировано: ' + stats.categorized;
    SpreadsheetApp.getUi().alert('Автокатегоризация', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    pfLogError_(e, 'pfApplyAutoCategorizationToAll', PF_LOG_LEVEL.ERROR);
    SpreadsheetApp.getUi().alert('Ошибка', 'Ошибка при применении автокатегоризации: ' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Create default category rules if sheet is empty.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function pfCreateDefaultCategoryRules_(ss) {
  var sheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.CATEGORY_RULES);
  if (!sheet) {
    return; // Sheet doesn't exist
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    return; // Sheet already has data
  }
  
  // Default rules (Russian merchants/descriptions)
  // Order: RuleName, Pattern, PatternType, Category, Subcategory, Priority, Active, ApplyTo
  var defaultRules = [
    // Grocery stores
    ['Пятёрочка', 'Пятёрочка', pfT_('pattern_type.contains'), 'Продукты', '', 10, true, pfT_('rule_apply_to.both')],
    ['Магнит', 'Магнит', pfT_('pattern_type.contains'), 'Продукты', '', 10, true, pfT_('rule_apply_to.both')],
    ['Перекресток', 'Перекресток', pfT_('pattern_type.contains'), 'Продукты', '', 10, true, pfT_('rule_apply_to.both')],
    ['Ашан', 'Ашан', pfT_('pattern_type.contains'), 'Продукты', '', 10, true, pfT_('rule_apply_to.both')],
    ['Лента', 'Лента', pfT_('pattern_type.contains'), 'Продукты', '', 10, true, pfT_('rule_apply_to.both')],
    
    // Restaurants and cafes
    ['Кафе', 'кафе', pfT_('pattern_type.contains'), 'Еда', 'Рестораны', 5, true, pfT_('rule_apply_to.both')],
    ['Ресторан', 'ресторан', pfT_('pattern_type.contains'), 'Еда', 'Рестораны', 5, true, pfT_('rule_apply_to.both')],
    ['Макдональдс', 'Макдональдс', pfT_('pattern_type.contains'), 'Еда', 'Рестораны', 5, true, pfT_('rule_apply_to.both')],
    ['KFC', 'KFC', pfT_('pattern_type.contains'), 'Еда', 'Рестораны', 5, true, pfT_('rule_apply_to.both')],
    
    // Transport
    ['Метро', 'Метро', pfT_('pattern_type.contains'), 'Транспорт', 'Общественный', 5, true, pfT_('rule_apply_to.both')],
    ['Автобус', 'Автобус', pfT_('pattern_type.contains'), 'Транспорт', 'Общественный', 5, true, pfT_('rule_apply_to.both')],
    ['Такси', 'Такси', pfT_('pattern_type.contains'), 'Транспорт', 'Такси', 5, true, pfT_('rule_apply_to.both')],
    ['Яндекс.Такси', 'Яндекс.Такси', pfT_('pattern_type.contains'), 'Транспорт', 'Такси', 5, true, pfT_('rule_apply_to.both')],
    ['Uber', 'Uber', pfT_('pattern_type.contains'), 'Транспорт', 'Такси', 5, true, pfT_('rule_apply_to.both')],
    
    // Health
    ['Аптека', 'Аптека', pfT_('pattern_type.contains'), 'Здоровье', 'Аптека', 5, true, pfT_('rule_apply_to.both')],
    ['36,6', '36,6', pfT_('pattern_type.contains'), 'Здоровье', 'Аптека', 5, true, pfT_('rule_apply_to.both')],
    
    // Entertainment
    ['Кино', 'Кино', pfT_('pattern_type.contains'), 'Развлечения', 'Кино', 5, true, pfT_('rule_apply_to.both')],
    ['Netflix', 'Netflix', pfT_('pattern_type.contains'), 'Развлечения', 'Подписки', 5, true, pfT_('rule_apply_to.both')],
    
    // Housing
    ['ЖКУ', 'ЖКУ', pfT_('pattern_type.contains'), 'Жильё', 'Коммунальные', 5, true, pfT_('rule_apply_to.both')],
    ['Коммунальные', 'Коммунальные', pfT_('pattern_type.contains'), 'Жильё', 'Коммунальные', 5, true, pfT_('rule_apply_to.both')],
    
    // Income
    ['Зарплата', 'Зарплата', pfT_('pattern_type.contains'), 'Зарплата', '', 10, true, pfT_('rule_apply_to.both')]
  ];
  
  // Write rules to sheet
  if (defaultRules.length > 0) {
    var range = sheet.getRange(2, 1, defaultRules.length, defaultRules[0].length);
    range.setValues(defaultRules);
    pfLogInfo_('Created ' + defaultRules.length + ' default category rules', 'pfCreateDefaultCategoryRules_');
  }
}
