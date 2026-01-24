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
  sheetKey: PF_SHEET_KEYS.TRANSACTIONS,
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

