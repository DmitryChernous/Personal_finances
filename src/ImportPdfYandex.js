/**
 * PDF parser for Yandex Card statements.
 *
 * Важно: структура текста Яндекс-выписки пока не изучена.
 * Этот модуль создаёт каркас парсера и логирует извлечённый текст,
 * чтобы на основе реальных логов построить точные правила.
 */

/**
 * Yandex PDF parser object.
 * Implements a subset of PF_IMPORTER_INTERFACE for PDF.
 */
var PF_PDF_YANDEX_PARSER = {
  /**
   * Detect if text is from Yandex PDF statement.
   * @param {string} text - Extracted text from PDF
   * @returns {boolean}
   */
  detect: function(text) {
    if (!text || typeof text !== 'string') {
      return false;
    }
    var normalized = text.toLowerCase();

    // Heuristics for Yandex Bank / Yandex Card.
    // Эти маркеры уточним после первой выгрузки текста.
    if (normalized.indexOf('яндекс') !== -1 ||
        normalized.indexOf('yandex') !== -1) {
      return true;
    }

    return false;
  },

  /**
   * Parse Yandex PDF text into raw transactions.
   *
   * Формат (по логам):
   *  - шапка с описанием, затем таблица:
   *    "Описание операции", "Дата и время операции", "Дата обработки",
   *    "Карта", "Сумма в валюте операции", "Сумма в валюте Договора/ЭСП".
   *  - Строки идут парами:
   *    1) строка с описанием и датой-временем:
   *       "Оплата СБП QR (YANDEX.TAXI) 02.01.2025 в 20:29"
   *    2) строка (или часть строки) с датой обработки и суммой(ами):
   *       "02.01.2025 –245,00 ₽ –245,00 ₽"
   *    Иногда в одной строке с суммами сразу несколько операций.
   *
   * Алгоритм:
   *  - читаем текст построчно;
   *  - после заголовка "Описание операции" начинаем собирать "ожидающие" операции:
   *    каждая строка вида "<описание> dd.mm.yyyy в HH:MM" -> кладём в очередь pending;
   *  - строки с суммами вида:
   *      "dd.mm.yyyy [*КАРТА] +/–XXX,YY ₽ +/–XXX,YY ₽ [dd.mm.yyyy ...]"
   *    парсим регуляркой, на каждый матч берём по одному элементу из очереди pending.
   *
   * @param {string} text - Extracted text from PDF
   * @param {Object} [options]
   * @returns {Array<Object>}
   */
  parse: function(text, options) {
    options = options || {};

    if (!this.detect(text)) {
      throw new Error('Text does not appear to be from Yandex PDF statement');
    }

    // Логируем первые ~5000 символов для анализа структуры.
    var snippet = text.substring(0, 5000);
    Logger.log('[YANDEX-PDF-TEXT] --- BEGIN ---');
    Logger.log(snippet);
    Logger.log('[YANDEX-PDF-TEXT] --- END ---');

    // Дополнительно логируем последние ~5000 символов (хвост выписки),
    // чтобы можно было анализировать последние страницы.
    var tailStart = Math.max(0, text.length - 5000);
    var tailSnippet = text.substring(tailStart);
    Logger.log('[YANDEX-PDF-TEXT-TAIL] --- BEGIN ---');
    Logger.log(tailSnippet);
    Logger.log('[YANDEX-PDF-TEXT-TAIL] --- END ---');

    var lines = String(text).split(/\r?\n/);

    // Собираем описания+даты и суммы отдельно, потом сопоставляем по дате
    // (порядок строк в OCR может не совпадать с порядком в таблице)
    var opsList = [];   // [{ description, opDate, opTime }, ...] в порядке появления
    var amountsList = []; // [{ procDate, amount }, ...] в порядке появления

    var inOperationsTable = false;
    var pendingDescOnly = [];
    var descBuffer = '';

    var headerPattern = /описание операции/i;
    var descWithDateTimePattern = /^(.*?)(\d{2}\.\d{2}\.\d{4})\s+в\s+(\d{2}:\d{2})/i;
    var amountsPattern = /(\d{2}\.\d{2}\.\d{4})\s+(?:\*\d{4}\s+)?([+\-–−]?\d[\d\s]*,\d{2})\s*₽\s+([+\-–−]?\d[\d\s]*,\d{2})\s*₽/g;

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;

      if (!inOperationsTable) {
        if (headerPattern.test(line)) inOperationsTable = true;
        continue;
      }

      // Сброс при итогах / новой таблице
      if (/исходящий остаток/i.test(line) ||
          /всего расходных операций/i.test(line) ||
          /выписка по договору/i.test(line)) {
        pendingDescOnly = [];
        descBuffer = '';
        inOperationsTable = false;
        continue;
      }

      // Описание + дата/время в одной строке
      var descMatch = line.match(descWithDateTimePattern);
      if (descMatch) {
        var descText = descMatch[1].trim();
        var opDate = descMatch[2];
        var opTime = descMatch[3];
        if (!descText && descBuffer) descText = descBuffer.trim();
        else if (descBuffer) descText = (descBuffer + ' ' + descText).trim();
        descBuffer = '';
        if (descText) {
          opsList.push({ description: descText, opDate: opDate, opTime: opTime });
        }
        continue;
      }

      // Только дата и время
      var dateTimeOnlyMatch = line.match(/^(\d{2}\.\d{2}\.\d{4})\s+в\s+(\d{2}:\d{2})/i);
      if (dateTimeOnlyMatch) {
        var dtDate = dateTimeOnlyMatch[1];
        var dtTime = dateTimeOnlyMatch[2];
        if (pendingDescOnly.length) {
          var descOnly = pendingDescOnly.shift();
          opsList.push({
            description: descOnly.description,
            opDate: dtDate,
            opTime: dtTime
          });
        } else if (descBuffer) {
          var bufferedDesc = descBuffer.trim();
          if (bufferedDesc) {
            opsList.push({
              description: bufferedDesc,
              opDate: dtDate,
              opTime: dtTime
            });
          }
        }
        descBuffer = '';
        continue;
      }

      // Заголовки страниц — пропускаем
      if (/^продолжение на следующей странице/i.test(line) ||
          /^страница \d+ из \d+/i.test(line) ||
          /^описание операции/i.test(line) ||
          /^дата и время операции/i.test(line) ||
          /^дата обработки/i.test(line) ||
          /^карта/i.test(line) ||
          /^сумма в валюте/i.test(line)) {
        continue;
      }

      // Строки описания без дат и сумм
      if (line.indexOf('₽') === -1 &&
          !/всего приходных операций/i.test(line)) {
        if (/^входящий перевод сбп/i.test(line)) {
          pendingDescOnly.push({ description: line });
        } else {
          descBuffer = descBuffer ? descBuffer + ' ' + line : line;
        }
        continue;
      }

      // Строки с суммами — только собираем (procDate, amount)
      amountsPattern.lastIndex = 0;
      var m;
      while ((m = amountsPattern.exec(line)) !== null) {
        var amountStr = m[2]
          .replace(/\s+/g, '')
          .replace(',', '.')
          .replace(/[–−]/g, '-');
        var amount = parseFloat(amountStr);
        if (!isNaN(amount)) {
          amountsList.push({ procDate: m[1], amount: amount });
        }
      }
    }

    // Сопоставление по дате: группируем описание+время и суммы по дате,
    // внутри каждой даты объединяем по порядку (1-я сумма с 1-м описанием и т.д.)
    var opsByDate = {};
    var datesOrder = [];
    for (var j = 0; j < opsList.length; j++) {
      var d = opsList[j].opDate;
      if (!opsByDate[d]) {
        opsByDate[d] = [];
        datesOrder.push(d);
      }
      opsByDate[d].push(opsList[j]);
    }
    var amountsByDate = {};
    for (var k = 0; k < amountsList.length; k++) {
      var d2 = amountsList[k].procDate;
      if (!amountsByDate[d2]) amountsByDate[d2] = [];
      amountsByDate[d2].push(amountsList[k].amount);
    }

    var rawTransactions = [];
    for (var di = 0; di < datesOrder.length; di++) {
      var dateKey = datesOrder[di];
      var opsOnDate = opsByDate[dateKey] || [];
      var amtsOnDate = amountsByDate[dateKey] || [];
      var n = Math.min(opsOnDate.length, amtsOnDate.length);
      for (var t = 0; t < n; t++) {
        var op = opsOnDate[t];
        var amount = amtsOnDate[t];
        var type = amount >= 0 ? 'income' : 'expense';
        rawTransactions.push({
          bank: 'yandex',
          date: op.opDate,
          time: op.opTime,
          description: [op.description],
          amount: amount,
          currency: 'RUB',
          type: type,
          category: '',
          sourceId: op.opDate.replace(/\D/g, '') + op.opTime.replace(/\D/g, '')
        });
      }
      // Суммы без описания (на эту дату): добавляем с датой из суммы, время 00:00
      for (var t2 = n; t2 < amtsOnDate.length; t2++) {
        var amount2 = amtsOnDate[t2];
        var type2 = amount2 >= 0 ? 'income' : 'expense';
        rawTransactions.push({
          bank: 'yandex',
          date: dateKey,
          time: '00:00',
          description: [''],
          amount: amount2,
          currency: 'RUB',
          type: type2,
          category: '',
          sourceId: dateKey.replace(/\D/g, '') + '0000'
        });
      }
    }
    // Даты, которые есть только в суммах (нет в opsList)
    for (var dateKey2 in amountsByDate) {
      if (opsByDate[dateKey2]) continue;
      var amts = amountsByDate[dateKey2];
      for (var t3 = 0; t3 < amts.length; t3++) {
        var amount3 = amts[t3];
        var type3 = amount3 >= 0 ? 'income' : 'expense';
        rawTransactions.push({
          bank: 'yandex',
          date: dateKey2,
          time: '00:00',
          description: [''],
          amount: amount3,
          currency: 'RUB',
          type: type3,
          category: '',
          sourceId: dateKey2.replace(/\D/g, '') + '0000'
        });
      }
    }

    // Сортируем по дате и времени для хронологического порядка
    rawTransactions.sort(function(a, b) {
      var dateA = a.date.replace(/(\d{2})\.(\d{2})\.(\d{4})/, '$3$2$1');
      var dateB = b.date.replace(/(\d{2})\.(\d{2})\.(\d{4})/, '$3$2$1');
      if (dateA !== dateB) return dateA < dateB ? -1 : 1;
      var timeA = (a.time || '00:00').replace(':', '');
      var timeB = (b.time || '00:00').replace(':', '');
      return timeA < timeB ? -1 : (timeA > timeB ? 1 : 0);
    });

    return rawTransactions;
  },

  /**
   * Normalize raw Yandex PDF transaction to DTO.
   *
   * @param {Object} rawTransaction
   * @param {Object} [options]
   * @returns {TransactionDTO}
   */
  normalize: function(rawTransaction, options) {
    options = options || {};
    var source = options.source || 'import:pdf:yandex';
    var defaultCurrency = options.defaultCurrency || 'RUB';
    var defaultAccount = options.defaultAccount || '';

    // Parse date (format: dd.mm.yyyy)
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
        pfLogWarning_('Error parsing date: ' + rawTransaction.date, 'PF_PDF_YANDEX_PARSER.normalize');
      }
    }

    // Parse amount - always positive (type determines expense/income)
    var amount = Math.abs(rawTransaction.amount || 0);

    // Determine type
    var type = rawTransaction.type || 'expense';

    // Combine description lines
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

    // Extract merchant from description (similar to Sberbank parser)
    var merchant = '';
    if (description) {
      // For "Входящий перевод СБП, ..." - extract sender name (first name part)
      var transferMatch = description.match(/Входящий перевод СБП[,\s]+([^,+]+?)(?:,\s*\+?\d|$)/i);
      if (transferMatch) {
        merchant = transferMatch[1].trim();
      } else {
        // For "Оплата товаров и услуг MERCHANT_NAME" pattern
        var merchantMatch = description.match(/Оплата товаров и услуг\s+(.+?)(?:\s+\d+|$)/i);
        if (merchantMatch) {
          merchant = merchantMatch[1].trim();
        } else {
          // For "Оплата СБП QR (MERCHANT)" pattern
          var qrMatch = description.match(/Оплата СБП QR\s*\(([^)]+)\)/i);
          if (qrMatch) {
            merchant = qrMatch[1].trim();
          }
        }
      }
    }

    var transaction = {
      date: date,
      type: type,
      account: rawTransaction.account || defaultAccount,
      accountTo: rawTransaction.accountTo || '',
      amount: amount,
      currency: rawTransaction.currency || defaultCurrency,
      category: rawTransaction.category || '',
      subcategory: rawTransaction.subcategory || '',
      merchant: merchant,
      description: description,
      tags: rawTransaction.tags || '',
      source: source,
      sourceId: rawTransaction.sourceId || '',
      rawData: JSON.stringify(rawTransaction),
      errors: []
    };

    var errors = pfValidateTransactionDTO_(transaction);
    transaction.errors = errors;

    return transaction;
  }
};

