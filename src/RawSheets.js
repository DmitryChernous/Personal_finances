/**
 * Синхронизация raw-листов с листом «Транзакции».
 *
 * Листы, имя которых начинается с "raw" (без учёта регистра), считаются
 * сырыми выписками по одному счёту. Колонки: ДАТА, ВРЕМЯ, КАТЕГОРИЯ, ОПИСАНИЕ, СУММА, ОСТАТОК СРЕДСТВ, СЧЕТ.
 * См. docs/RAW_SHEETS_ARCHITECTURE.md.
 */

/** Префикс имени листа для raw-выписок */
var PF_RAW_SHEET_PREFIX = 'raw';

/** Индексы колонок на raw-листе (1-based): A=1 ДАТА, B=2 ВРЕМЯ, C=3 КАТЕГОРИЯ, D=4 ОПИСАНИЕ, E=5 СУММА, F=6 ОСТАТОК, G=7 СЧЕТ */
var PF_RAW_COL = {
  DATE: 1,
  TIME: 2,
  CATEGORY: 3,
  DESCRIPTION: 4,
  AMOUNT: 5,
  BALANCE: 6,
  ACCOUNT: 7
};

/** Количество колонок данных на raw-листе (для getRange) */
var PF_RAW_DATA_COLS = 7;

/**
 * Возвращает все листы, имя которых начинается с PF_RAW_SHEET_PREFIX (без учёта регистра).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {GoogleAppsScript.Spreadsheet.Sheet[]}
 */
function pfGetRawSheets_(ss) {
  var prefix = PF_RAW_SHEET_PREFIX.toLowerCase();
  var sheets = ss.getSheets();
  var out = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.length >= prefix.length && name.substring(0, prefix.length).toLowerCase() === prefix) {
      out.push(sheets[i]);
    }
  }
  return out;
}

/**
 * Парсит дату из строки вида dd.mm.yyyy.
 * @param {string} s
 * @returns {Date|null}
 */
function pfParseRawDate_(s) {
  if (!s || typeof s !== 'string') return null;
  var parts = s.trim().split('.');
  if (parts.length !== 3) return null;
  var day = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1;
  var year = parseInt(parts[2], 10);
  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
  var d = new Date(year, month, day);
  if (isNaN(d.getTime())) return null;
  return d;
}

/**
 * Нормализует время для SourceId: "16:40" -> "1640", "0:42" -> "0042".
 * @param {string} s
 * @returns {string}
 */
function pfNormalizeRawTime_(s) {
  if (!s || typeof s !== 'string') return '0000';
  var t = s.trim().replace(':', '');
  if (t.length === 3) t = '0' + t;
  if (t.length < 4) t = (t + '0000').substring(0, 4);
  return t;
}

/**
 * Парсит сумму из строки вида "-1 500,00" или "1 300,00".
 * @param {*} val - значение ячейки (строка или число)
 * @returns {{ amount: number, type: string }|null}
 */
function pfParseRawAmount_(val) {
  if (val === null || val === undefined || val === '') return null;
  var s = String(val).replace(/\s/g, '').replace(',', '.');
  var num = parseFloat(s);
  if (isNaN(num)) return null;
  var type = num < 0 ? 'expense' : 'income';
  return { amount: Math.abs(num), type: type };
}

/**
 * Формирует SourceId для дедупликации: дата (ddmmyyyy) + время (hhmm) + сумма (целое).
 * @param {string} dateStr - dd.mm.yyyy
 * @param {string} timeStr - HH:mm
 * @param {number} amount
 * @returns {string}
 */
function pfRawSourceId_(dateStr, timeStr, amount) {
  var d = (dateStr || '').replace(/\D/g, '');
  if (d.length !== 8) d = '00000000';
  var t = pfNormalizeRawTime_(timeStr || '');
  var a = String(Math.round(amount));
  return d + t + a;
}

/**
 * Читает данные с одного raw-листа и возвращает массив объектов в каноническом формате (для проверки дедупликации и записи).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} sheetName
 * @param {string} defaultCurrency
 * @returns {Array<{date: Date, type: string, account: string, amount: number, currency: string, category: string, description: string, source: string, sourceId: string, status: string}>}
 */
function pfReadRawSheet_(sheet, sheetName, defaultCurrency) {
  defaultCurrency = defaultCurrency || 'RUB';
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow, PF_RAW_DATA_COLS).getValues();
  var source = 'raw:' + sheetName;
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateStr = row[PF_RAW_COL.DATE - 1];
    var timeStr = row[PF_RAW_COL.TIME - 1];
    var category = row[PF_RAW_COL.CATEGORY - 1];
    var description = row[PF_RAW_COL.DESCRIPTION - 1];
    var amountVal = row[PF_RAW_COL.AMOUNT - 1];
    var accountVal = row[PF_RAW_COL.ACCOUNT - 1];

    var date = pfParseRawDate_(dateStr);
    if (!date) continue;
    var parsed = pfParseRawAmount_(amountVal);
    if (!parsed) continue;

    var account = (accountVal !== null && accountVal !== undefined && String(accountVal).trim() !== '')
      ? String(accountVal).trim()
      : sheetName;

    var sourceId = pfRawSourceId_(String(dateStr), String(timeStr || ''), parsed.amount);

    result.push({
      date: date,
      type: parsed.type,
      account: account,
      accountTo: '',
      amount: parsed.amount,
      currency: defaultCurrency,
      category: (category !== null && category !== undefined) ? String(category).trim() : '',
      subcategory: '',
      merchant: '',
      description: (description !== null && description !== undefined) ? String(description).trim() : '',
      tags: '',
      source: source,
      sourceId: sourceId,
      status: 'ok'
    });
  }
  return result;
}

/**
 * Собирает все транзакции с raw-листов и добавляет в «Транзакции» только новые (по дедупликации Source + SourceId).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] - если не передан, берётся активная таблица
 * @returns {{ added: number, skipped: number, sheetsProcessed: number, errors: string[] }}
 */
function pfSyncRawSheetsToTransactions(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var result = { added: 0, skipped: 0, sheetsProcessed: 0, errors: [] };

  var txSheet = pfFindSheetByKey_(ss, PF_SHEET_KEYS.TRANSACTIONS);
  if (!txSheet) {
    result.errors.push('Лист «Транзакции» не найден.');
    return result;
  }

  var rawSheets = pfGetRawSheets_(ss);
  if (rawSheets.length === 0) {
    result.errors.push('Нет листов с именем, начинающимся на "raw".');
    return result;
  }

  var defaultCurrency = pfGetSetting_(ss, PF_SETUP_KEYS.DEFAULT_CURRENCY) || 'RUB';

  // Загрузить существующие ключи дедупликации (Source + SourceId)
  var existingKeys = pfGetExistingTransactionKeys_();

  var allNewRows = [];
  for (var s = 0; s < rawSheets.length; s++) {
    var sheet = rawSheets[s];
    var sheetName = sheet.getName();
    try {
      var rows = pfReadRawSheet_(sheet, sheetName, defaultCurrency);
      result.sheetsProcessed++;
      for (var r = 0; r < rows.length; r++) {
        var tx = rows[r];
        var dedupeKey = tx.source + ':' + tx.sourceId;
        if (existingKeys[dedupeKey]) {
          result.skipped++;
          continue;
        }
        existingKeys[dedupeKey] = true;
        allNewRows.push(tx);
      }
    } catch (e) {
      result.errors.push('Лист "' + sheetName + '": ' + (e.message || e.toString()));
    }
  }

  if (allNewRows.length === 0) {
    return result;
  }

  // Записать новые строки в лист «Транзакции» в порядке колонок PF_TRANSACTIONS_SCHEMA
  var colOrder = ['Date', 'Type', 'Account', 'AccountTo', 'Amount', 'Currency', 'Category', 'Subcategory', 'Merchant', 'Description', 'Tags', 'Source', 'SourceId', 'Status'];
  var numCols = colOrder.length;
  var lastRow = txSheet.getLastRow();
  if (lastRow < 1) lastRow = 1;
  var startRow = lastRow + 1;
  var endRow = startRow + allNewRows.length - 1;

  var values = [];
  for (var i = 0; i < allNewRows.length; i++) {
    var tx = allNewRows[i];
    var row = [];
    for (var c = 0; c < numCols; c++) {
      var key = colOrder[c];
      var v = tx[key];
      if (key === 'Date' && v instanceof Date) {
        row.push(v);
      } else {
        row.push(v !== undefined && v !== null ? v : '');
      }
    }
    values.push(row);
  }

  txSheet.getRange(startRow, 1, endRow, numCols).setValues(values);

  // Формат даты и суммы
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  if (dateCol) txSheet.getRange(startRow, dateCol, endRow, dateCol).setNumberFormat('dd.mm.yyyy');
  if (amountCol) txSheet.getRange(startRow, amountCol, endRow, amountCol).setNumberFormat('0.00');

  result.added = allNewRows.length;
  return result;
}
