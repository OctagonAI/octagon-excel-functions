/**
 * Browser and API support detection utilities
 * Used to provide compatibility warnings to users
 */

import Logger from './logger';

/**
 * Detects if the browser is Internet Explorer
 * @returns true if the browser is IE, false otherwise
 */
export function detectIE(): boolean {
  try {
    const ua = window.navigator.userAgent;
    const msie = ua.indexOf('MSIE ');
    const trident = ua.indexOf('Trident/');
    
    return (msie > 0 || trident > 0);
  } catch (error) {
    Logger.error('Error detecting browser:', error);
    return false; // Assume not IE if detection fails
  }
}

/**
 * Checks if all required APIs are supported in the current environment
 * @returns An array of issues (empty if all requirements are met)
 */
export function checkRequiredApiSupport(): string[] {
  const issues: string[] = [];
  
  try {
    // Check for SharedRuntime support
    if (!Office.context.requirements.isSetSupported('SharedRuntime', '1.1')) {
      issues.push("This add-in requires SharedRuntime support. Please update to a newer version of Office.");
    }
    
    // Check for ExcelApi support
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      issues.push("Some Excel features require Excel API 1.7 or later. Please update to a newer version of Excel.");
    }
    
    // Check for CustomFunctions support
    if (!Office.context.requirements.isSetSupported('CustomFunctions', '1.1')) {
      issues.push("Custom functions require Excel support for CustomFunctions 1.1 or later.");
    }
    
    // Check if we're in a compatible platform for our add-in
    checkPlatformCompatibility(issues);
    
  } catch (error) {
    Logger.error('Error checking API support:', error);
    issues.push("Error checking API compatibility. Some features may not work as expected.");
  }
  
  return issues;
}

/**
 * Checks platform compatibility and adds any issues to the issues array
 * @param issues Array to add platform compatibility issues to
 */
function checkPlatformCompatibility(issues: string[]): void {
  // Check if we're in a web browser
  try {
    const platform = Office.context.platform;
    const host = Office.context.host;
    
    // Log platform information for diagnostics
    Logger.info(`Running on platform: ${platform}, host: ${host}`);
    
    // Check specific platform requirements
    if (platform === Office.PlatformType.PC) {
      // Check for Windows version compatibility if needed
      // This would require additional detection code
    } else if (platform === Office.PlatformType.Mac) {
      // Check for Mac-specific issues if needed
    } else if (platform === Office.PlatformType.OfficeOnline) {
      // Check for Office Online specific issues
      
      // Example: Warn about possible streaming limitations in Excel Online
      issues.push("Excel Online may have limitations with long-running functions. Consider using Excel desktop for complex operations.");
    }
    
    // Ensure we're running in Excel
    if (host !== Office.HostType.Excel) {
      issues.push("This add-in is designed for Excel and may not function correctly in other Office applications.");
    }
    
  } catch (error) {
    Logger.error('Error detecting platform:', error);
    // Don't add to issues here as we already have a general error message
  }
}