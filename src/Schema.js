/**
 * Central schema definitions for sheets.
 *
 * Note: This repo targets RU Google Sheets locale. See `SHEETS_LOCALE.md`.
 */

var PF_SCHEMA_VERSION = 1;

/**
 * Canonical Transactions sheet schema.
 * Headers are resolved via i18n using `key` (see `src/I18n.js`).
 */
var PF_TRANSACTIONS_SCHEMA = {
  // IMPORTANT: keep as string literal to avoid init-order issues across script files.
  sheetKey: 'transactions',
  columns: [
    { key: 'Date', required: true },
    { key: 'Type', required: true, allowed: ['expense', 'income', 'transfer'] },
    { key: 'Account', required: true },
    { key: 'AccountTo', required: false },
    { key: 'Amount', required: true },
    { key: 'Currency', required: true },
    { key: 'Category', required: false },
    { key: 'Subcategory', required: false },
    { key: 'Merchant', required: false },
    { key: 'Description', required: false },
    { key: 'Tags', required: false },
    { key: 'Source', required: true },
    { key: 'SourceId', required: false },
    { key: 'Status', required: true, allowed: ['ok', 'needs_review', 'duplicate', 'deleted'] }
  ]
};

/**
 * Accounts reference sheet schema.
 */
var PF_ACCOUNTS_SCHEMA = {
  sheetKey: 'accounts',
  columns: [
    { key: 'Account', required: true },
    { key: 'AccountType', required: false },
    { key: 'Currency', required: false },
    { key: 'InitialBalance', required: false },
    { key: 'Active', required: false },
    { key: 'Description', required: false }
  ]
};

/**
 * Categories reference sheet schema.
 */
var PF_CATEGORIES_SCHEMA = {
  sheetKey: 'categories',
  columns: [
    { key: 'Category', required: true },
    { key: 'Subcategory', required: false },
    { key: 'CategoryType', required: false },
    { key: 'Active', required: false },
    { key: 'Description', required: false }
  ]
};

/**
 * Budgets sheet schema.
 * Headers are resolved via i18n using `key` (see `src/I18n.js`).
 */
var PF_BUDGETS_SCHEMA = {
  sheetKey: 'budgets',
  columns: [
    { key: 'Category', required: true },
    { key: 'Subcategory', required: false },
    { key: 'Period', required: true, allowed: ['month', 'year'] },
    { key: 'PeriodValue', required: true },
    { key: 'Amount', required: true },
    { key: 'Fact', required: false }, // Calculated
    { key: 'Remaining', required: false }, // Calculated
    { key: 'Status', required: false }, // Calculated
    { key: 'PercentUsed', required: false }, // Calculated
    { key: 'Active', required: false },
    { key: 'Description', required: false }
  ]
};

/**
 * Recurring Transactions sheet schema.
 * Headers are resolved via i18n using `key` (see `src/I18n.js`).
 */
var PF_RECURRING_TRANSACTIONS_SCHEMA = {
  sheetKey: 'recurring_transactions',
  columns: [
    { key: 'Name', required: true },
    { key: 'Type', required: true, allowed: ['expense', 'income', 'transfer'] },
    { key: 'Frequency', required: true, allowed: ['weekly', 'monthly', 'quarterly', 'yearly'] },
    { key: 'DayOfMonth', required: false }, // 1-31, required for monthly/quarterly/yearly
    { key: 'DayOfWeek', required: false }, // 1-7 (1=Monday), required for weekly
    { key: 'StartDate', required: true },
    { key: 'EndDate', required: false },
    { key: 'Account', required: true },
    { key: 'AccountTo', required: false },
    { key: 'Amount', required: true },
    { key: 'Currency', required: true },
    { key: 'Category', required: false },
    { key: 'Subcategory', required: false },
    { key: 'Merchant', required: false },
    { key: 'Description', required: false },
    { key: 'Active', required: false },
    { key: 'LastCreated', required: false } // Date of last transaction created
  ]
};

