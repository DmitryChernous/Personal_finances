/**
 * PDF parser for Yandex Card statements.
 * v3: последовательное сопоставление — при появлении строки с суммой сразу привязываем к последней операции с этой датой (строка описания → строка суммы в документе)
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
 * Алгоритм (v3 — последовательное сопоставление):
 *  - читаем текст построчно;
 *  - операции (описание + dd.mm.yyyy в HH:MM) кладём в очередь по дате;
 *  - при появлении строки с суммой (dd.mm.yyyy … ₽ … ₽) сразу привязываем к первой ожидающей операции с этой датой.
 *  Надёжность зависит от порядка в PDF: если в выписке строка описания идёт перед строкой суммы по каждой операции — сопоставление будет верным.
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

    // Последовательное сопоставление: при появлении строки с суммой сразу привязываем к последней
    // ожидающей операции с той же датой (как в документе: строка описания → строка суммы).
    var pendingOpsByDate = {};  // date -> [{ description, opDate, opTime }, ...] очередь по порядку
    var rawTransactions = [];
    var fallbackIndexByDate = {};  // date -> число уже выданных fallback для уникального sourceId

    var inOperationsTable = false;
    var pendingDescOnly = [];
    var descBuffer = '';

    function pushOp(opDate, opTime, description) {
      if (!pendingOpsByDate[opDate]) pendingOpsByDate[opDate] = [];
      pendingOpsByDate[opDate].push({ description: description, opDate: opDate, opTime: opTime });
    }

    function consumeOp(procDate, amount) {
      var queue = pendingOpsByDate[procDate];
      var op = queue && queue.length ? queue.shift() : null;
      if (op) {
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
      } else {
        var idx = (fallbackIndexByDate[procDate] || 0);
        fallbackIndexByDate[procDate] = idx + 1;
        var type2 = amount >= 0 ? 'income' : 'expense';
        rawTransactions.push({
          bank: 'yandex',
          date: procDate,
          time: '00:00',
          description: [''],
          amount: amount,
          currency: 'RUB',
          type: type2,
          category: '',
          sourceId: procDate.replace(/\D/g, '') + ('000' + idx).slice(-4)
        });
      }
    }

    // Убираем из описаний заголовки страниц и шапку таблицы (OCR часто вставляет их в текст)
    function cleanDesc_(s) {
      if (!s || typeof s !== 'string') return '';
      return s
        .replace(/\s*МСК\s+Дата обработки\s+МСК\s*/gi, ' ')
        .replace(/\s*Дата обработки\s+МСК\s*/gi, ' ')
        .replace(/\s*МСК\s+Оплата/gi, ' Оплата')
        .replace(/\s{2,}/g, ' ')
        .trim();
    }

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
        var descText = cleanDesc_(descMatch[1].trim());
        var opDate = descMatch[2];
        var opTime = descMatch[3];
        if (!descText && descBuffer) descText = cleanDesc_(descBuffer.trim());
        else if (descBuffer) descText = cleanDesc_(descBuffer + ' ' + descMatch[1].trim());
        descBuffer = '';
        if (descText) {
          pushOp(opDate, opTime, descText);
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
          pushOp(dtDate, dtTime, cleanDesc_(descOnly.description));
        } else if (descBuffer) {
          var bufferedDesc = cleanDesc_(descBuffer.trim());
          if (bufferedDesc) pushOp(dtDate, dtTime, bufferedDesc);
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
          pendingDescOnly.push({ description: cleanDesc_(line) });
        } else {
          var cleaned = cleanDesc_(line);
          if (cleaned) descBuffer = descBuffer ? descBuffer + ' ' + cleaned : cleaned;
        }
        continue;
      }

      // Строки с суммами — сразу сопоставляем с очередью операций по этой дате
      amountsPattern.lastIndex = 0;
      var m;
      while ((m = amountsPattern.exec(line)) !== null) {
        var amountStr = m[2]
          .replace(/\s+/g, '')
          .replace(',', '.')
          .replace(/[–−]/g, '-');
        var amount = parseFloat(amountStr);
        if (!isNaN(amount)) {
          consumeOp(m[1], amount);
        }
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

    // Determine type: по описанию надёжнее, чем по знаку суммы (OCR/парсинг может ошибаться)
    var type = rawTransaction.type || 'expense';
    if (rawTransaction.description) {
      var descStr = Array.isArray(rawTransaction.description)
        ? rawTransaction.description.join(' ').trim()
        : String(rawTransaction.description).trim();
      if (/входящий перевод\s+сбп/i.test(descStr)) type = 'income';
      else if (/оплата\s+(товаров|сбп|услуг)/i.test(descStr) || /оплата\s+сбп\s*qr/i.test(descStr)) type = 'expense';
    }

    // Combine description lines and strip header fragments (МСК Дата обработки и т.д.)
    var description = '';
    if (Array.isArray(rawTransaction.description)) {
      description = rawTransaction.description.join(' ').trim();
    } else if (rawTransaction.description) {
      description = String(rawTransaction.description).trim();
    }
    description = description
      .replace(/\s*МСК\s+Дата обработки\s+МСК\s*/gi, ' ')
      .replace(/\s*Дата обработки\s+МСК\s*/gi, ' ')
      .replace(/\s*МСК\s+Оплата/gi, ' Оплата')
      .replace(/\s{2,}/g, ' ')
      .trim();

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

