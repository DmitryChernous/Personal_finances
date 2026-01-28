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
    var rawTransactions = [];

    // Флаг, что мы дошли до таблицы с операциями
    var inOperationsTable = false;
    // Очередь ожидающих операций (описание + дата/время)
    var pendingOps = [];
    // Очередь описаний без даты/времени (например, несколько "Входящий перевод СБП"
    // подряд, после которых идут несколько строк с датой/временем)
    var pendingDescOnly = [];

    // Регексы
    var headerPattern = /описание операции/i;
    var descWithDateTimePattern = /^(.*?)(\d{2}\.\d{2}\.\d{4})\s+в\s+(\d{2}:\d{2})/i;
    // Строка с датой обработки, опциональной картой и суммами:
    //  11.01.2025 *7088 –100,00 ₽ –100,00 ₽
    //  или без карты:
    //  05.01.2025 –273,00 ₽ –273,00 ₽
    var amountsPattern = /(\d{2}\.\d{2}\.\d{4})\s+(?:\*\d{4}\s+)?([+\-–−]?\d[\d\s]*,\d{2})\s*₽\s+([+\-–−]?\d[\d\s]*,\d{2})\s*₽/g;

    // Буфер для многострочного описания до строки с датой/временем
    var descBuffer = '';

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) {
        continue;
      }

      // Ищем заголовок таблицы
      if (!inOperationsTable) {
        if (headerPattern.test(line)) {
          inOperationsTable = true;
        }
        continue;
      }

      // Линия описания с датой и временем (описание и дата в одной строке)
      var descMatch = line.match(descWithDateTimePattern);
      if (descMatch) {
        var descText = descMatch[1].trim();
        var opDate = descMatch[2];
        var opTime = descMatch[3];

        // Если до этого накапливался буфер описания (несколько строк),
        // приклеиваем его перед текущим описанием.
        if (!descText && descBuffer) {
          descText = descBuffer.trim();
        } else if (descBuffer) {
          descText = (descBuffer + ' ' + descText).trim();
        }
        descBuffer = '';

        // Иногда описание может быть пустым, но это редкий случай
        if (descText) {
          pendingOps.push({
            description: descText,
            opDate: opDate,
            opTime: opTime
          });
        }
        continue;
      }

      // Линия ТОЛЬКО с датой и временем (описание на предыдущих строках)
      var dateTimeOnlyMatch = line.match(/^(\d{2}\.\d{2}\.\d{4})\s+в\s+(\d{2}:\d{2})/i);
      if (dateTimeOnlyMatch) {
        var dtDate = dateTimeOnlyMatch[1];
        var dtTime = dateTimeOnlyMatch[2];

        // Если есть очереди отдельных описаний (например, несколько
        // "Входящий перевод СБП" подряд), связываем каждую дату/время
        // с одним элементом очереди.
        if (pendingDescOnly.length) {
          var descOnly = pendingDescOnly.shift();
          pendingOps.push({
            description: descOnly.description,
            opDate: dtDate,
            opTime: dtTime
          });
        } else if (descBuffer) {
          var bufferedDesc = descBuffer.trim();
          if (bufferedDesc) {
            pendingOps.push({
              description: bufferedDesc,
              opDate: dtDate,
              opTime: dtTime
            });
          }
        }
        // После использования буфера очищаем его
        descBuffer = '';
        continue;
      }

      // Если это строка описания без дат и сумм:
      // - для повторяющихся описаний переводов по СБП складываем
      //   каждую строку в отдельную очередь (pendingDescOnly),
      //   чтобы потом сопоставить 1:1 с датами/суммами;
      // - для остальных многострочных описаний (например,
      //   "Погашение основного долга по договору №" + номер договора)
      //   продолжаем накапливать в общем буфере descBuffer.
      if (line.indexOf('₽') === -1 &&
          !/исходящий остаток/i.test(line) &&
          !/всего расходных операций/i.test(line) &&
          !/всего приходных операций/i.test(line) &&
          !/выписка по договору/i.test(line)) {
        if (/^входящий перевод сбп/i.test(line)) {
          pendingDescOnly.push({
            description: line
          });
        } else {
          if (descBuffer) {
            descBuffer += ' ' + line;
          } else {
            descBuffer = line;
          }
        }
      }

      // Линия с суммами
      var hasAmountMatch = false;
      var m;
      while ((m = amountsPattern.exec(line)) !== null) {
        hasAmountMatch = true;
        if (!pendingOps.length) {
          // Нет соответствующей строки описания — пропускаем
          continue;
        }

        var pending = pendingOps.shift();
        var amountStr = m[2];

        // Нормализуем число: убираем пробелы, меняем запятую на точку, нормализуем минус
        var normalizedAmountStr = amountStr
          .replace(/\s+/g, '')
          .replace(',', '.')
          .replace(/[–−]/g, '-');

        var amount = parseFloat(normalizedAmountStr);
        if (isNaN(amount)) {
          continue;
        }

        var type = amount >= 0 ? 'income' : 'expense';

        rawTransactions.push({
          bank: 'yandex',
          date: pending.opDate,
          time: pending.opTime,
          description: [pending.description],
          amount: amount,
          currency: 'RUB',
          type: type,
          // Для дальнейшего улучшения можно будет выделять категорию из описания
          category: '',
          sourceId: pending.opDate.replace(/\D/g, '') + pending.opTime.replace(/\D/g, '')
        });
      }

      // Если строка с суммами не содержала ни одного матча, сбрасываем state,
      // когда доходим до итогов / новых разделов.
      if (!hasAmountMatch) {
        if (/исходящий остаток/i.test(line) ||
            /всего расходных операций/i.test(line) ||
            /выписка по договору/i.test(line)) {
          pendingOps = [];
          pendingDescOnly = [];
          descBuffer = '';
          inOperationsTable = false; // возможно начинается новая таблица
        }
      }
    }

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
      // For "Оплата товаров и услуг MERCHANT_NAME" pattern
      var merchantMatch = description.match(/Оплата товаров и услуг\s+(.+?)(?:\s+\d+|$)/i);
      if (merchantMatch) {
        merchant = merchantMatch[1].trim();
      } else {
        // For "Оплата СБП QR (MERCHANT)" pattern
        var qrMatch = description.match(/Оплата СБП QR\s*\(([^)]+)\)/i);
        if (qrMatch) {
          merchant = qrMatch[1].trim();
        } else {
          // For "Входящий перевод СБП, ..." - extract sender name
          var transferMatch = description.match(/Входящий перевод СБП[,\s]+([^,]+)/i);
          if (transferMatch) {
            merchant = transferMatch[1].trim();
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

