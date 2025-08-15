/**
 * Logging utility for Octagon Excel Add-in
 * This module provides centralized logging functionality with different log levels
 * to help debug issues during development and production.
 */

// Log levels
export enum LogLevel {
  DEBUG = 'DEBUG',
  INFO = 'INFO',
  WARN = 'WARN',
  ERROR = 'ERROR'
}

// Current log level - can be adjusted based on environment
let currentLogLevel: LogLevel = 
  process.env.NODE_ENV === 'production' ? LogLevel.ERROR : LogLevel.DEBUG;

// Log level priority (used to determine if a message should be logged)
const logLevelPriority: Record<LogLevel, number> = {
  [LogLevel.DEBUG]: 0,
  [LogLevel.INFO]: 1,
  [LogLevel.WARN]: 2,
  [LogLevel.ERROR]: 3
};

/**
 * Main logging function that handles all log levels
 */
export function log(level: LogLevel, message: string, data?: any): void {
  // Only log if the current level priority is less than or equal to the message level
  if (logLevelPriority[level] >= logLevelPriority[currentLogLevel]) {
    const timestamp = new Date().toISOString();
    const formattedMessage = `[${timestamp}] [${level}] ${message}`;

    switch (level) {
      case LogLevel.DEBUG:
        console.debug(formattedMessage, data || '');
        break;
      case LogLevel.INFO:
        console.info(formattedMessage, data || '');
        break;
      case LogLevel.WARN:
        console.warn(formattedMessage, data || '');
        break;
      case LogLevel.ERROR:
        console.error(formattedMessage, data || '');
        break;
    }
  }
}

/**
 * Set the current log level
 */
export function setLogLevel(level: LogLevel): void {
  currentLogLevel = level;
  log(LogLevel.INFO, `Log level set to ${level}`);
}

/**
 * Get all logs stored in session storage
 */
export function getLogs(): any[] {
  try {
    return JSON.parse(sessionStorage.getItem('octagon_logs') || '[]');
  } catch (e) {
    console.error('[Logger] Error retrieving logs from sessionStorage:', e);
    return [];
  }
}

/**
 * Clear all logs in session storage
 */
export function clearLogs(): void {
  try {
    sessionStorage.removeItem('octagon_logs');
  } catch (e) {
    // Log the error instead of silently failing
    console.error('[Logger] Error clearing logs from sessionStorage:', e);
  }
}

// Convenience methods for each log level
export const debug = (message: string, data?: any) => log(LogLevel.DEBUG, message, data);
export const info = (message: string, data?: any) => log(LogLevel.INFO, message, data);
export const warn = (message: string, data?: any) => log(LogLevel.WARN, message, data);
export const error = (message: string, data?: any) => log(LogLevel.ERROR, message, data);

// Export a default logger object
export default {
  debug,
  info,
  warn,
  error,
  setLogLevel,
  getLogs,
  clearLogs
};