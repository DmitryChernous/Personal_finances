/**
 * Date utility functions.
 * 
 * Provides centralized date conversion utilities for consistent handling
 * of Date objects and ISO string conversions throughout the application.
 */

/**
 * Convert Date object to ISO 8601 string.
 * Used for JSON serialization when passing data through google.script.run.
 * 
 * @param {Date} date - Date object to convert
 * @returns {string} ISO 8601 string (e.g., "2025-01-15T10:30:00.000Z")
 */
function pfDateToISOString_(date) {
  if (!date) {
    return '';
  }
  
  if (date instanceof Date) {
    return date.toISOString();
  }
  
  // If already a string, return as-is (assume it's already ISO format)
  if (typeof date === 'string') {
    return date;
  }
  
  // Try to convert to Date first
  try {
    var dateObj = new Date(date);
    if (!isNaN(dateObj.getTime())) {
      return dateObj.toISOString();
    }
  } catch (e) {
    // Ignore
  }
  
  return '';
}

/**
 * Convert ISO 8601 string to Date object.
 * Used when parsing data received from client or stored as strings.
 * 
 * @param {string} isoString - ISO 8601 string (e.g., "2025-01-15T10:30:00.000Z")
 * @returns {Date|null} Date object or null if parsing fails
 */
function pfISOStringToDate_(isoString) {
  if (!isoString || typeof isoString !== 'string' || isoString.trim().length === 0) {
    return null;
  }
  
  // If already a Date object, return as-is
  if (isoString instanceof Date) {
    return isoString;
  }
  
  try {
    var dateObj = new Date(isoString);
    if (!isNaN(dateObj.getTime())) {
      return dateObj;
    }
  } catch (e) {
    // Ignore
  }
  
  return null;
}

/**
 * Format date to yyyy-MM-dd string for use in deduplication keys.
 * 
 * @param {Date|string} date - Date object or ISO string
 * @returns {string} Formatted date string (yyyy-MM-dd) or empty string if invalid
 */
function pfFormatDateForDedupe_(date) {
  if (!date) {
    return '';
  }
  
  var dateObj = null;
  
  if (date instanceof Date) {
    dateObj = date;
  } else if (typeof date === 'string') {
    dateObj = pfISOStringToDate_(date);
  }
  
  if (!dateObj) {
    return '';
  }
  
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
