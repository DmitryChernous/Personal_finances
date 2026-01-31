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
 * Парсит дату: объект Date из Sheets или строка dd.mm.yyyy.
 * @param {Date|string} s
 * @returns {Date|null}
 */
function pfParseRawDate_(s) {
  if (s === null || s === undefined) return null;
  if (s instanceof Date && !isNaN(s.getTime())) return s;
  if (typeof s !== 'string') return null;
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
 * Принимает строку, число (доля дня из Sheets) или Date (время из ячейки).
 * @param {string|number|Date} s
 * @returns {string} "hhmm" (4 цифры)
 */
function pfNormalizeRawTime_(s) {
  if (s === null || s === undefined) return '0000';
  if (s instanceof Date && !isNaN(s.getTime())) {
    var h = s.getHours();
    var m = s.getMinutes();
    return (h < 10 ? '0' : '') + h + (m < 10 ? '0' : '') + m;
  }
  if (typeof s === 'number' && s >= 0 && s < 1) {
    var h = Math.floor(s * 24);
    var m = Math.round((s * 24 - h) * 60);
    return (h < 10 ? '0' : '') + h + (m < 10 ? '0' : '') + m;
  }
  var t = String(s).trim().replace(':', '');
  if (t.length === 3) t = '0' + t;
  if (t.length < 4) t = (t + '0000').substring(0, 4);
  return t;
}

/**
 * Парсит сумму: число из Sheets или строка вида "-1 500,00", "1 300,00".
 * @param {*} val - значение ячейки (число или строка)
 * @returns {{ amount: number, type: string }|null}
 */
function pfParseRawAmount_(val) {
  if (val === null || val === undefined || val === '') return null;
  var num;
  if (typeof val === 'number' && !isNaN(val)) {
    num = val;
  } else {
    var s = String(val).replace(/\s/g, '').replace(',', '.');
    num = parseFloat(s);
    if (isNaN(num)) return null;
  }
  var type = num < 0 ? 'expense' : 'income';
  return { amount: Math.abs(num), type: type };
}

/**
 * Префикс для SourceId, чтобы таблица всегда воспринимала значение как текст (не число).
 * Без префикса длинные числовые строки отображаются в экспоненциальной нотации (2,21E+15).
 */
var PF_RAW_SOURCE_ID_PREFIX = 'id_';

/**
 * Формирует SourceId для дедупликации: префикс + дата (ddmmyyyy) + время (hhmm) + сумма (целое) + номер строки.
 * Номер строки (rowIndex) гарантирует уникальность: разные строки raw-листа не дадут один ключ даже при одинаковых дате/времени/сумме.
 * @param {Date|string} dateVal - дата (объект Date или строка dd.mm.yyyy)
 * @param {string|number} timeStr - время HH:mm или доля дня
 * @param {number} amount
 * @param {number} [rowIndex] - номер строки на raw-листе (1-based; обычно 2, 3, …). Если задан, добавляется к id для уникальности.
 * @returns {string}
 */
function pfRawSourceId_(dateVal, timeStr, amount, rowIndex) {
  var d;
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    var day = dateVal.getDate();
    var month = dateVal.getMonth() + 1;
    var year = dateVal.getFullYear();
    d = (day < 10 ? '0' : '') + day + (month < 10 ? '0' : '') + month + year;
  } else {
    d = String(dateVal || '').replace(/\D/g, '');
    if (d.length !== 8) d = '00000000';
  }
  var t = pfNormalizeRawTime_(timeStr || '');
  var a = String(Math.round(amount));
  var base = PF_RAW_SOURCE_ID_PREFIX + d + t + a;
  if (rowIndex !== undefined && rowIndex !== null && rowIndex > 0) {
    return base + '_r' + rowIndex;
  }
  return base;
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

    var rowIndexOnSheet = i + 2;
    var sourceId = pfRawSourceId_(date, timeStr, parsed.amount, rowIndexOnSheet);

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

  // Ключи, которые уже есть на листе «Транзакции» (не меняем в процессе запуска).
  var existingOnSheet = pfGetExistingTransactionKeys_();
  // Ключи, добавленные в этом запуске (чтобы не путать «уже на листе» с «повтор в сырых данных»).
  var addedInThisRun = {};
  var duplicateCountInRun = {};

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
        var alreadyOnSheet = !!existingOnSheet[dedupeKey];
        var seenInThisRun = !!addedInThisRun[dedupeKey];

        if (alreadyOnSheet) {
          result.skipped++;
          continue; // Не добавлять дубликат в лист «Транзакции»
        }
        if (seenInThisRun) {
          // Повтор в сырых данных в этом же запуске — делаем SourceId уникальным и добавляем как обычную строку.
          var count = (duplicateCountInRun[dedupeKey] || 1) + 1;
          duplicateCountInRun[dedupeKey] = count;
          var newSourceId = tx.sourceId + '_' + count;
          tx = {
            date: tx.date,
            type: tx.type,
            account: tx.account,
            accountTo: tx.accountTo,
            amount: tx.amount,
            currency: tx.currency,
            category: tx.category,
            subcategory: tx.subcategory,
            merchant: tx.merchant,
            description: tx.description,
            tags: tx.tags,
            source: tx.source,
            sourceId: newSourceId,
            status: 'ok'
          };
          addedInThisRun[tx.source + ':' + newSourceId] = true;
        } else {
          addedInThisRun[dedupeKey] = true;
        }
        allNewRows.push(tx);
      }
    } catch (e) {
      result.errors.push('Лист "' + sheetName + '": ' + (e.message || e.toString()));
    }
  }

  if (allNewRows.length === 0) {
    return result;
  }

  // Записать новые строки в лист «Транзакции» в порядке колонок PF_TRANSACTIONS_SCHEMA.
  // Ключи объекта tx — в нижнем регистре (date, type, account, ...).
  var colOrder = ['date', 'type', 'account', 'accountTo', 'amount', 'currency', 'category', 'subcategory', 'merchant', 'description', 'tags', 'source', 'sourceId', 'status'];
  var numCols = colOrder.length;
  var values = [];
  for (var i = 0; i < allNewRows.length; i++) {
    var tx = allNewRows[i];
    var row = [];
    for (var c = 0; c < numCols; c++) {
      var key = colOrder[c];
      var v = tx[key];
      if (key === 'date' && v instanceof Date) {
        row.push(v);
      } else {
        row.push(v !== undefined && v !== null ? v : '');
      }
    }
    values.push(row);
  }

  // Запись чанками. Используем getRange(row, col, numRows, numCols), чтобы размер диапазона точно совпадал с массивом.
  var numRows = values.length;
  var lastRow = txSheet.getLastRow();
  if (lastRow < 1) lastRow = 1;
  var startRow = lastRow + 1;
  var chunkSize = 500;
  for (var offset = 0; offset < numRows; offset += chunkSize) {
    var chunk = values.slice(offset, offset + chunkSize);
    var chunkStartRow = startRow + offset;
    var numRowsInChunk = chunk.length;
    var range = txSheet.getRange(chunkStartRow, 1, numRowsInChunk, numCols);
    range.setValues(chunk);
  }

  // Формат даты, суммы и ID источника. getRange(startRow, col, numRows, 1): третий аргумент — число строк.
  // Формат применяется ко всем добавленным строкам (startRow … startRow + numRows - 1).
  var dateCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Date');
  var amountCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'Amount');
  var sourceIdCol = pfColumnIndex_(PF_TRANSACTIONS_SCHEMA, 'SourceId');
  if (dateCol) txSheet.getRange(startRow, dateCol, numRows, 1).setNumberFormat('dd.mm.yyyy');
  if (amountCol) txSheet.getRange(startRow, amountCol, numRows, 1).setNumberFormat('0.00');
  if (sourceIdCol) txSheet.getRange(startRow, sourceIdCol, numRows, 1).setNumberFormat('@');

  result.added = allNewRows.length;
  return result;
}
