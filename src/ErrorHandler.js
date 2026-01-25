/**
 * Centralized error handling and logging.
 * 
 * Provides unified error handling, logging levels, and error response formatting.
 */

/**
 * Log levels.
 */
var PF_LOG_LEVEL = {
  DEBUG: 'DEBUG',
  INFO: 'INFO',
  WARNING: 'WARNING',
  ERROR: 'ERROR'
};

/**
 * Current log level (only ERROR and WARNING in production).
 * Set to PF_LOG_LEVEL.DEBUG for detailed debugging.
 */
var PF_CURRENT_LOG_LEVEL = PF_LOG_LEVEL.WARNING;

/**
 * Check if a log level should be logged.
 * @param {string} level - Log level (DEBUG, INFO, WARNING, ERROR)
 * @returns {boolean} True if should log
 */
function pfShouldLog_(level) {
  var levels = [PF_LOG_LEVEL.DEBUG, PF_LOG_LEVEL.INFO, PF_LOG_LEVEL.WARNING, PF_LOG_LEVEL.ERROR];
  var currentIndex = levels.indexOf(PF_CURRENT_LOG_LEVEL);
  var messageIndex = levels.indexOf(level);
  return messageIndex >= currentIndex;
}

/**
 * Centralized error logging.
 * @param {Error|string} error - Error object or error message
 * @param {string} context - Context where error occurred (function name, module, etc.)
 * @param {string} level - Log level (default: ERROR)
 */
function pfLogError_(error, context, level) {
  level = level || PF_LOG_LEVEL.ERROR;
  
  if (!pfShouldLog_(level)) {
    return; // Skip logging if level is too low
  }
  
  var errorMessage = error instanceof Error ? error.toString() : String(error);
  var errorStack = error instanceof Error ? (error.stack || 'No stack') : '';
  
  var logMessage = '[' + level + '] [' + (context || 'Unknown') + '] ' + errorMessage;
  if (errorStack && level === PF_LOG_LEVEL.ERROR) {
    logMessage += '\nStack: ' + errorStack;
  }
  
  Logger.log(logMessage);
}

/**
 * Create standardized error response object.
 * @param {string} message - Error message
 * @param {string} [code] - Error code (optional)
 * @param {Error} [error] - Original error object (optional)
 * @returns {Object} Error response {success: false, message: string, code?: string}
 */
function pfCreateErrorResponse_(message, code, error) {
  var response = {
    success: false,
    message: message || 'Произошла ошибка'
  };
  
  if (code) {
    response.code = code;
  }
  
  if (error && pfShouldLog_(PF_LOG_LEVEL.ERROR)) {
    pfLogError_(error, 'pfCreateErrorResponse_', PF_LOG_LEVEL.ERROR);
  }
  
  return response;
}

/**
 * Create standardized success response object.
 * @param {string} message - Success message
 * @param {Object} [data] - Additional data (optional)
 * @returns {Object} Success response {success: true, message: string, ...data}
 */
function pfCreateSuccessResponse_(message, data) {
  var response = {
    success: true,
    message: message || 'Операция выполнена успешно'
  };
  
  if (data) {
    for (var key in data) {
      if (data.hasOwnProperty(key)) {
        response[key] = data[key];
      }
    }
  }
  
  return response;
}

/**
 * Handle error with centralized logging and return error response.
 * @param {Error|string} error - Error object or error message
 * @param {string} context - Context where error occurred
 * @param {string} [userMessage] - User-friendly error message (optional)
 * @returns {Object} Error response {success: false, message: string}
 */
function pfHandleError_(error, context, userMessage) {
  // Log error
  pfLogError_(error, context, PF_LOG_LEVEL.ERROR);
  
  // Return user-friendly error response
  var message = userMessage || (error instanceof Error ? error.message : String(error));
  return pfCreateErrorResponse_(message, null, error);
}

/**
 * Log debug message (only if DEBUG level is enabled).
 * @param {string} message - Debug message
 * @param {string} [context] - Context (optional)
 */
function pfLogDebug_(message, context) {
  if (pfShouldLog_(PF_LOG_LEVEL.DEBUG)) {
    Logger.log('[DEBUG] [' + (context || 'Unknown') + '] ' + message);
  }
}

/**
 * Log info message (only if INFO level or lower is enabled).
 * @param {string} message - Info message
 * @param {string} [context] - Context (optional)
 */
function pfLogInfo_(message, context) {
  if (pfShouldLog_(PF_LOG_LEVEL.INFO)) {
    Logger.log('[INFO] [' + (context || 'Unknown') + '] ' + message);
  }
}

/**
 * Log warning message.
 * @param {string} message - Warning message
 * @param {string} [context] - Context (optional)
 */
function pfLogWarning_(message, context) {
  if (pfShouldLog_(PF_LOG_LEVEL.WARNING)) {
    Logger.log('[WARNING] [' + (context || 'Unknown') + '] ' + message);
  }
}
