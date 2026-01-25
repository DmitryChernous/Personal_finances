/**
 * Application constants.
 * 
 * Centralized location for all magic numbers, strings, and configuration values.
 * This helps avoid typos, makes code more maintainable, and allows easy changes.
 */

/**
 * Transaction status values.
 */
var PF_TRANSACTION_STATUS = {
  OK: 'ok',
  DUPLICATE: 'duplicate',
  NEEDS_REVIEW: 'needs_review',
  DELETED: 'deleted'
};

/**
 * Transaction type values.
 */
var PF_TRANSACTION_TYPE = {
  EXPENSE: 'expense',
  INCOME: 'income',
  TRANSFER: 'transfer'
};

/**
 * Import source identifiers.
 */
var PF_IMPORT_SOURCE = {
  MANUAL: 'manual',
  CSV: 'import:csv',
  SBERBANK: 'import:sberbank',
  ERROR: 'import:error'
};

/**
 * Import batch size for processing large files.
 */
var PF_IMPORT_BATCH_SIZE = 200;

/**
 * Maximum file size for import (50MB).
 */
var PF_IMPORT_MAX_FILE_SIZE = 50 * 1024 * 1024; // 50MB in bytes

/**
 * Maximum number of transactions to process in one import.
 */
var PF_IMPORT_MAX_TRANSACTIONS = 10000;

/**
 * Default currency code.
 */
var PF_DEFAULT_CURRENCY = 'RUB';

/**
 * Supported currencies.
 */
var PF_SUPPORTED_CURRENCIES = ['RUB', 'USD', 'EUR'];

/**
 * Category types.
 */
var PF_CATEGORY_TYPE = {
  EXPENSE: 'expense',
  INCOME: 'income',
  BOTH: 'both'
};

/**
 * Account types.
 */
var PF_ACCOUNT_TYPE = {
  CASH: 'cash',
  CARD: 'card',
  DEPOSIT: 'deposit',
  INVESTMENT: 'investment',
  OTHER: 'other'
};
