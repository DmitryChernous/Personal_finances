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
   * На первом этапе:
   *  - логируем фрагмент текста в логи Apps Script;
   *  - кидаем понятную ошибку, что парсер ещё не реализован.
   *
   * После того как мы увидим логи с реальным текстом выписки,
   * сюда будут добавлены реальные правила парсинга
   * (аналогично PF_PDF_SBERBANK_PARSER.parse).
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

    throw new Error('Yandex PDF parser is not implemented yet. ' +
                    'Text snippet has been logged with prefix [YANDEX-PDF-TEXT].');
  },

  /**
   * Normalize raw Yandex PDF transaction to DTO.
   *
   * Будет использовано после реализации parse().
   * Пока оставляем заготовку, похожую на fallback-нормализацию.
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

    var transaction = {
      date: rawTransaction.date || null,
      type: rawTransaction.type || 'expense',
      account: rawTransaction.account || defaultAccount,
      accountTo: rawTransaction.accountTo || '',
      amount: rawTransaction.amount || 0,
      currency: rawTransaction.currency || defaultCurrency,
      category: rawTransaction.category || '',
      subcategory: rawTransaction.subcategory || '',
      merchant: rawTransaction.merchant || '',
      description: rawTransaction.description || '',
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

