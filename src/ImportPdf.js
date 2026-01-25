/**
 * PDF importer for bank statements.
 * 
 * Handles PDF files by:
 * 1. Extracting text using Google Drive API (with OCR support)
 * 2. Detecting bank by text content
 * 3. Delegating to bank-specific parsers
 */

/**
 * Extract text from PDF file using Google Drive API.
 * @param {Blob|GoogleAppsScript.Drive.File} pdfFile - PDF file blob or Drive file
 * @param {Object} [options] - Options: {useOCR: boolean, ocrLanguage: string}
 * @returns {string} Extracted text
 */
function pfExtractTextFromPdf_(pdfFile, options) {
  options = options || {};
  var useOCR = options.useOCR !== false; // Default to true
  var ocrLanguage = options.ocrLanguage || 'ru';
  
  try {
    var fileId = null;
    var blob = null;
    
    // Handle different input types
    if (pdfFile instanceof Blob) {
      blob = pdfFile;
      // Upload to Drive temporarily
      var tempFile = DriveApp.createFile(blob);
      fileId = tempFile.getId();
    } else if (pdfFile.getId) {
      // It's a Drive File
      fileId = pdfFile.getId();
      blob = pdfFile.getBlob();
    } else {
      throw new Error('Invalid PDF file: expected Blob or Drive File');
    }
    
    // Check if file is PDF
    if (blob.getContentType() !== 'application/pdf') {
      throw new Error('File is not a PDF: ' + blob.getContentType());
    }
    
    // Extract text using Drive API with OCR
    var resource = {
      title: blob.getName() || 'temp_pdf_' + Date.now(),
      mimeType: 'application/pdf'
    };
    
    var driveOptions = {
      ocr: useOCR,
      ocrLanguage: ocrLanguage,
      fields: 'id,title'
    };
    
    // Use Drive API to convert PDF to Google Doc (with OCR)
    var doc = Drive.Files.insert(resource, blob, driveOptions);
    var docId = doc.id;
    
    // Extract text from the document
    var text = DocumentApp.openById(docId).getBody().getText();
    
    // Clean up temporary document
    try {
      DriveApp.getFileById(docId).setTrashed(true);
    } catch (e) {
      pfLogWarning_('Could not delete temporary document: ' + e.toString(), 'pfExtractTextFromPdf_');
    }
    
    // Clean up temporary file if we created it
    if (fileId && pdfFile instanceof Blob) {
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch (e) {
        pfLogWarning_('Could not delete temporary PDF file: ' + e.toString(), 'pfExtractTextFromPdf_');
      }
    }
    
    return text;
    
  } catch (e) {
    pfLogError_(e, 'pfExtractTextFromPdf_', PF_LOG_LEVEL.ERROR);
    throw new Error('Error extracting text from PDF: ' + (e.message || e.toString()));
  }
}

/**
 * Detect bank from PDF text content.
 * @param {string} text - Extracted text from PDF
 * @returns {string|null} Bank identifier ('sberbank', 'tinkoff', etc.) or null
 */
function pfDetectBankFromPdfText_(text) {
  if (!text || typeof text !== 'string') {
    return null;
  }
  
  var normalizedText = text.toLowerCase();
  
  // Check for Sberbank markers
  if (normalizedText.indexOf('сбербанк') !== -1 || 
      normalizedText.indexOf('sberbank') !== -1 ||
      normalizedText.indexOf('выписка по счёту') !== -1 ||
      normalizedText.indexOf('сбербанк онлайн') !== -1) {
    return 'sberbank';
  }
  
  // Check for Tinkoff markers
  if (normalizedText.indexOf('тинькофф') !== -1 || 
      normalizedText.indexOf('tinkoff') !== -1 ||
      normalizedText.indexOf('тинькофф банк') !== -1) {
    return 'tinkoff';
  }
  
  // Add more banks as needed
  
  return null;
}

/**
 * PDF importer object.
 * Implements PF_IMPORTER_INTERFACE.
 */
var PF_PDF_IMPORTER = {
  /**
   * Detect if data is a PDF file.
   * @param {Blob|string|Array<Array<*>>} data
   * @param {string} [fileName]
   * @returns {boolean}
   */
  detect: function(data, fileName) {
    // Check by file name
    if (fileName && fileName.toLowerCase().endsWith('.pdf')) {
      return true;
    }
    
    // Check by MIME type if it's a Blob
    if (data instanceof Blob) {
      return data.getContentType() === 'application/pdf';
    }
    
    // Can't detect PDF from string content alone (would need to check binary signature)
    // So we rely on file name or MIME type
    return false;
  },
  
  /**
   * Parse PDF file: extract text and delegate to bank-specific parser.
   * @param {Blob|string} data - PDF file blob or base64 string
   * @param {Object} [options] - Parser options
   * @returns {Array<Object>} Array of raw transaction objects
   */
  parse: function(data, options) {
    options = options || {};
    
    var blob = null;
    
    // Handle different input types
    if (data instanceof Blob) {
      blob = data;
    } else if (typeof data === 'string') {
      // Assume base64 encoded PDF
      try {
        var bytes = Utilities.base64Decode(data);
        blob = Utilities.newBlob(bytes, 'application/pdf');
      } catch (e) {
        throw new Error('Invalid PDF data: ' + (e.message || e.toString()));
      }
    } else {
      throw new Error('PDF importer requires Blob or base64 string, got: ' + typeof data);
    }
    
    // Extract text from PDF
    var extractOptions = {
      useOCR: options.useOCR !== false, // Default to true
      ocrLanguage: options.ocrLanguage || 'ru'
    };
    
    var text = pfExtractTextFromPdf_(blob, extractOptions);
    
    if (!text || text.trim().length === 0) {
      throw new Error('Could not extract text from PDF. File might be empty or corrupted.');
    }
    
    // Detect bank
    var bank = pfDetectBankFromPdfText_(text);
    
    if (!bank) {
      throw new Error('Could not detect bank from PDF content. Supported banks: Sberbank, Tinkoff.');
    }
    
    // Delegate to bank-specific parser
    var parser = null;
    if (bank === 'sberbank' && typeof PF_PDF_SBERBANK_PARSER !== 'undefined') {
      parser = PF_PDF_SBERBANK_PARSER;
    } else if (bank === 'tinkoff' && typeof PF_PDF_TINKOFF_PARSER !== 'undefined') {
      parser = PF_PDF_TINKOFF_PARSER;
    } else {
      throw new Error('Parser for bank "' + bank + '" is not available. Please ensure the parser module is loaded.');
    }
    
    // Parse using bank-specific parser
    var rawTransactions = parser.parse(text, options);
    
    return rawTransactions;
  },
  
  /**
   * Normalize raw transaction to DTO.
   * Delegates to bank-specific normalizer.
   * @param {Object} rawTransaction
   * @param {Object} [options]
   * @returns {TransactionDTO}
   */
  normalize: function(rawTransaction, options) {
    options = options || {};
    
    // Get bank from raw transaction or options
    var bank = rawTransaction.bank || options.bank;
    
    if (!bank) {
      throw new Error('Bank identifier is required for normalization');
    }
    
    // Delegate to bank-specific normalizer
    var parser = null;
    if (bank === 'sberbank' && typeof PF_PDF_SBERBANK_PARSER !== 'undefined') {
      parser = PF_PDF_SBERBANK_PARSER;
    } else if (bank === 'tinkoff' && typeof PF_PDF_TINKOFF_PARSER !== 'undefined') {
      parser = PF_PDF_TINKOFF_PARSER;
    } else {
      throw new Error('Normalizer for bank "' + bank + '" is not available');
    }
    
    if (parser.normalize) {
      return parser.normalize(rawTransaction, options);
    } else {
      // Fallback normalization
      return this._normalizeFallback_(rawTransaction, options);
    }
  },
  
  /**
   * Fallback normalization if bank-specific normalizer is not available.
   * @private
   */
  _normalizeFallback_: function(rawTransaction, options) {
    options = options || {};
    var source = options.source || 'import:pdf';
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
    
    // Validate
    var errors = pfValidateTransactionDTO_(transaction);
    transaction.errors = errors;
    
    return transaction;
  },
  
  /**
   * Generate deduplication key for transaction.
   * @param {TransactionDTO} transaction
   * @returns {string}
   */
  dedupeKey: function(transaction) {
    return pfGenerateDedupeKey_(transaction);
  }
};
