/**
 * ============================================================
 *  FSC / WRC / WLP — Google Apps Script Backend
 *  WITH DUAL EMAIL NOTIFICATIONS
 * ============================================================
 *
 *  This file contains server-side functions for the FormApp.html
 *  web application. It handles form submissions, validates data,
 *  enforces business rules, and writes to Google Sheets.
 *
 *  EMAIL NOTIFICATION FEATURES:
 *  1. PENDING Confirmation - Sent immediately upon form submission
 *  2. TRANSMITTED Notification - Sent when STATUS → "TRANSMITTED TO KAWASAKI"
 *
 *  SETUP INSTRUCTIONS:
 *  1. Replace SPREADSHEET_ID with your Google Sheets ID
 *  2. Ensure your spreadsheet has sheets: WRC, FSC, WLP, Log
 *  3. Deploy as web app
 *  4. Install onEdit trigger (see instructions at end of file)
 * ============================================================
 */

// ════════════════════════════════════════════════════════════
//  CONSTANTS
// ════════════════════════════════════════════════════════════

/**
 * IMPORTANT — Replace with YOUR spreadsheet ID
 * Find it in your spreadsheet URL:
 * https://docs.google.com/spreadsheets/d/15rPY6eyA-mMhtAsoVag3rfbCtmTuNmbQ8DFbK6Cjaa0/edit
 */
const SPREADSHEET_ID = '15rPY6eyA-mMhtAsoVag3rfbCtmTuNmbQ8DFbK6Cjaa0';

/** Sheet names */
const SHEET_WRC = 'WRC';
const SHEET_FSC = 'FSC';
const SHEET_WLP = 'WLP';
const SHEET_LOG = 'Log';

/** Column indices for uniqueness checks (0-based) */
const WRC_COL_WRC_NO    = 0;  // Column A
const WRC_COL_ENGINE_NO = 1;  // Column B

// ════════════════════════════════════════════════════════════
//  HELPER FUNCTIONS
// ════════════════════════════════════════════════════════════

/**
 * Get or create a sheet by name
 */
function getSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sheet;
}

/**
 * Check if a value exists in a column (for uniqueness validation)
 */
function isDuplicate_(sheet, colIndex, value) {
  if (!value || String(value).trim() === '') return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const data = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  const needle = String(value).trim().toUpperCase();
  return data.some(row => String(row[0]).trim().toUpperCase() === needle);
}

/**
 * Find column index by header name (1-based)
 * Returns 0 if not found
 */
function findColumnByHeader_(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.findIndex(h => String(h).toUpperCase() === String(headerName).toUpperCase());
  return index >= 0 ? index + 1 : 0;
}

/**
 * Log errors to the Log sheet
 */
function logError_(message, wrcNo, engineNo) {
  const logSheet = getSheet_(SHEET_LOG, ['Timestamp', 'WRC No.', 'ENGINE No.', 'Message']);
  logSheet.appendRow([
    new Date(),
    wrcNo || '',
    engineNo || '',
    message
  ]);
}

// ════════════════════════════════════════════════════════════
//  WEB APP DEPLOYMENT
// ════════════════════════════════════════════════════════════

/**
 * doGet handler - serves the FormApp.html file
 * 
 * DEPLOYMENT:
 * 1. Click Deploy → New deployment
 * 2. Select type: Web app
 * 3. Execute as: Me
 * 4. Who has access: Anyone (or as needed)
 * 5. Deploy and copy the URL
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('FormApp')
    .setTitle('FSC WRC WLP Form Entry')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Show FormApp as a sidebar in the spreadsheet
 */
function showFormApp() {
  const html = HtmlService.createHtmlOutputFromFile('FormApp')
    .setTitle('FSC WRC WLP Forms');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Show FormApp as a modal dialog
 */
function showFormDialog() {
  const html = HtmlService.createHtmlOutputFromFile('FormApp')
    .setWidth(850)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'FSC WRC WLP Form Entry');
}

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 FSC WRC WLP')
    .addItem('📝 Open Form Entry', 'showFormApp')
    .addItem('🌐 Open Form Dialog', 'showFormDialog')
    .addToUi();
}

// ════════════════════════════════════════════════════════════
//  FORM SUBMISSION HANDLERS
// ════════════════════════════════════════════════════════════

/**
 * Submit WRC form data
 * 
 * Validates:
 * - Required fields
 * - Uniqueness of WRC No. and ENGINE No.
 * - Date format (yyyymmdd)
 * 
 * @param {Object} formData - Form data from client
 * @returns {Object} {success: boolean, message: string}
 */
function submitWRC(formData) {
  try {
    // Note: STATUS column already exists in sheet, just adding BRANCH and EMAIL
    const sheet = getSheet_(SHEET_WRC, [
      'WRC No.', 'ENGINE No.', 'First Name', 'MI', 'LAST NAME',
      'MUN./CITY', 'PROVINCE', 'CONTACT NO. OF CUSTOMER', 'DATE PURCHASED',
      'DEALER NAME', 'Customer ADDRESS', 'AGE', 'GENDER', 'Dealer Code', 
      'File/Image Link', 'BRANCH', 'EMAIL'
    ]);
    
    // Validate required fields
    if (!formData.wrcNo || !formData.engineNo || !formData.firstName || 
        !formData.lastName || !formData.contactNo || !formData.dealerCode ||
        !formData.branch || !formData.email) {
      return {
        success: false,
        message: 'Required fields are missing. Please fill all fields marked with *'
      };
    }
    
    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(formData.email)) {
      return {
        success: false,
        message: 'Invalid email address format.'
      };
    }
    
    // Check uniqueness of WRC No.
    if (isDuplicate_(sheet, WRC_COL_WRC_NO, formData.wrcNo)) {
      logError_('Duplicate WRC No. from web form', formData.wrcNo, formData.engineNo);
      return {
        success: false,
        message: 'WRC No. "' + formData.wrcNo + '" already exists. Please use a unique WRC number.'
      };
    }
    
    // Check uniqueness of ENGINE No.
    if (isDuplicate_(sheet, WRC_COL_ENGINE_NO, formData.engineNo)) {
      logError_('Duplicate ENGINE No. from web form', formData.wrcNo, formData.engineNo);
      return {
        success: false,
        message: 'ENGINE No. "' + formData.engineNo + '" already exists. Please use a unique engine number.'
      };
    }
    
    // Validate date format (should be yyyymmdd - 8 digits)
    if (formData.datePurchased && !/^\d{8}$/.test(formData.datePurchased)) {
      return {
        success: false,
        message: 'Invalid date format. Date should be 8 digits (yyyymmdd).'
      };
    }
    
    // Append row to WRC sheet
    // Order: A-Q (WRC No., ENGINE No., First Name, MI, LAST NAME, 
    //        MUN./CITY, PROVINCE, CONTACT NO., DATE PURCHASED, 
    //        DEALER NAME, Customer ADDRESS, AGE, GENDER, Dealer Code, 
    //        File Link, BRANCH, EMAIL)
    // Note: STATUS column exists separately and will be updated
    const newRow = sheet.appendRow([
      formData.wrcNo,              // A: WRC No.
      formData.engineNo,           // B: ENGINE No.
      formData.firstName,          // C: First Name
      formData.mi || '',           // D: MI
      formData.lastName,           // E: LAST NAME
      formData.munCity || '',      // F: MUN./CITY
      formData.province || '',     // G: PROVINCE
      formData.contactNo,          // H: CONTACT NO. OF CUSTOMER
      formData.datePurchased,      // I: DATE PURCHASED (yyyymmdd)
      formData.dealerName || '',   // J: DEALER NAME
      formData.customerAddress || '', // K: Customer ADDRESS
      formData.age || '',          // L: AGE
      formData.gender || '',       // M: GENDER
      formData.dealerCode,         // N: Dealer Code
      formData.fileUrl || '',      // O: File/Image Link
      formData.branch,             // P: BRANCH
      formData.email               // Q: EMAIL
    ]);
    
    // Update STATUS column (existing column) to PENDING
    const lastRow = sheet.getLastRow();
    const statusColIndex = findColumnByHeader_(sheet, 'STATUS');
    if (statusColIndex > 0) {
      sheet.getRange(lastRow, statusColIndex).setValue('PENDING');
    }
    
    // Send PENDING confirmation email with WRC and Engine numbers
    sendPendingConfirmationEmail(
      formData.email, 
      'WRC', 
      formData.wrcNo, 
      formData.engineNo,
      formData.branch,
      formData.dealerName
    );
    
    return {
      success: true,
      message: 'WRC registration submitted successfully! WRC No: ' + formData.wrcNo + '. A confirmation email has been sent to ' + formData.email + '.'
    };
    
  } catch (error) {
    logError_('WRC submission error: ' + error.message, formData.wrcNo || '', formData.engineNo || '');
    return {
      success: false,
      message: 'Server error: ' + error.message
    };
  }
}

/**
 * Submit FSC form data
 * 
 * Validates:
 * - Required fields
 * - Date components (month/day/year)
 * 
 * @param {Object} formData - Form data from client
 * @returns {Object} {success: boolean, message: string}
 */
function submitFSC(formData) {
  try {
    // Note: STATUS column already exists in sheet, just adding BRANCH and EMAIL
    const sheet = getSheet_(SHEET_FSC, [
      'Dealer Transmittal No.', 'Dealer/Mechanic Code', 'WRC Number',
      'Frame Number', 'Coupon Number', 'Actual Mileage',
      'Repaired Month', 'Repaired Day', 'Repaired Year', 'KSC Code', 
      'File/Image Link', 'BRANCH', 'EMAIL'
    ]);
    
    // Validate required fields
    if (!formData.dealerTransNo || !formData.dealerMechCode || !formData.wrcNumber || 
        !formData.frameNumber || !formData.actualMileage || !formData.branch || !formData.email) {
      return {
        success: false,
        message: 'Required fields are missing. Please fill all fields marked with *'
      };
    }
    
    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(formData.email)) {
      return {
        success: false,
        message: 'Invalid email address format.'
      };
    }
    
    // Validate date components
    if (!formData.repairedMonth || !formData.repairedDay || !formData.repairedYear) {
      return {
        success: false,
        message: 'Invalid repair date. Please select a valid date.'
      };
    }
    
    // Validate actual mileage is a whole number
    if (formData.actualMileage && !Number.isInteger(Number(formData.actualMileage))) {
      return {
        success: false,
        message: 'Actual mileage must be a whole number (no decimals).'
      };
    }
    
    // Validate coupon number if provided (must be 1, 2, 3, or 4)
    if (formData.couponNumber && !['1', '2', '3', '4'].includes(formData.couponNumber)) {
      return {
        success: false,
        message: 'Invalid coupon number. Must be 1, 2, 3, or 4.'
      };
    }
    
    // Append row to FSC sheet
    // Order: A-M (Dealer Transmittal No., Dealer/Mechanic Code, WRC Number,
    //        Frame Number, Coupon Number, Actual Mileage,
    //        Repaired Month, Repaired Day, Repaired Year, KSC Code, 
    //        File Link, BRANCH, EMAIL)
    // Note: STATUS column exists separately and will be updated
    const newRow = sheet.appendRow([
      formData.dealerTransNo,      // A: Dealer Transmittal No.
      formData.dealerMechCode,     // B: Dealer/Mechanic Code
      formData.wrcNumber,          // C: WRC Number
      formData.frameNumber,        // D: Frame Number
      formData.couponNumber || '', // E: Coupon Number
      formData.actualMileage,      // F: Actual Mileage
      formData.repairedMonth,      // G: Repaired Month (MM)
      formData.repairedDay,        // H: Repaired Day (DD)
      formData.repairedYear,       // I: Repaired Year (YYYY)
      formData.kscCode || '',      // J: KSC Code
      formData.fileUrl || '',      // K: File/Image Link
      formData.branch,             // L: BRANCH
      formData.email               // M: EMAIL
    ]);
    
    // Update STATUS column (existing column) to PENDING
    const lastRow = sheet.getLastRow();
    const statusColIndex = findColumnByHeader_(sheet, 'STATUS');
    if (statusColIndex > 0) {
      sheet.getRange(lastRow, statusColIndex).setValue('PENDING');
    }
    
    // Send PENDING confirmation email with WRC and Frame numbers
    sendPendingConfirmationEmail(
      formData.email, 
      'FSC', 
      formData.wrcNumber, 
      formData.frameNumber,
      formData.branch,
      formData.dealerMechCode
    );
    
    return {
      success: true,
      message: 'FSC entry submitted successfully! WRC Number: ' + formData.wrcNumber + '. A confirmation email has been sent to ' + formData.email + '.'
    };
    
  } catch (error) {
    logError_('FSC submission error: ' + error.message, formData.wrcNumber || '', '');
    return {
      success: false,
      message: 'Server error: ' + error.message
    };
  }
}

/**
 * Submit WLP form data
 * 
 * Validates:
 * - Required fields
 * - Date components (month/day/year)
 * 
 * @param {Object} formData - Form data from client
 * @returns {Object} {success: boolean, message: string}
 */
function submitWLP(formData) {
  try {
    // Note: STATUS column already exists in sheet, just adding BRANCH and EMAIL
    const sheet = getSheet_(SHEET_WLP, [
      'WRS Number', 'Repair Acknowledged Month', 'Repair Acknowledged Day',
      'Repair Acknowledged Year', 'Acknowledged By: (Customer Name)', 
      'Dealer/Mechanic Code', 'File/Image Link', 'BRANCH', 'EMAIL'
    ]);
    
    // Validate required fields
    if (!formData.wrsNumber || !formData.acknowledgedBy || !formData.dealerMechCode ||
        !formData.branch || !formData.email) {
      return {
        success: false,
        message: 'Required fields are missing. Please fill all fields marked with *'
      };
    }
    
    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(formData.email)) {
      return {
        success: false,
        message: 'Invalid email address format.'
      };
    }
    
    // Validate date components
    if (!formData.repairAckMonth || !formData.repairAckDay || !formData.repairAckYear) {
      return {
        success: false,
        message: 'Invalid acknowledgment date. Please select a valid date.'
      };
    }
    
    // Append row to WLP sheet
    // Order: A-I (WRS Number, Repair Acknowledged Month, Repair Acknowledged Day,
    //        Repair Acknowledged Year, Acknowledged By, Dealer/Mechanic Code, 
    //        File Link, BRANCH, EMAIL)
    // Note: STATUS column exists separately and will be updated
    const newRow = sheet.appendRow([
      formData.wrsNumber,        // A: WRS Number
      formData.repairAckMonth,   // B: Repair Acknowledged Month (MM)
      formData.repairAckDay,     // C: Repair Acknowledged Day (DD)
      formData.repairAckYear,    // D: Repair Acknowledged Year (YYYY)
      formData.acknowledgedBy,   // E: Acknowledged By (Customer Name)
      formData.dealerMechCode,   // F: Dealer/Mechanic Code
      formData.fileUrl || '',    // G: File/Image Link
      formData.branch,           // H: BRANCH
      formData.email             // I: EMAIL
    ]);
    
    // Update STATUS column (existing column) to PENDING
    const lastRow = sheet.getLastRow();
    const statusColIndex = findColumnByHeader_(sheet, 'STATUS');
    if (statusColIndex > 0) {
      sheet.getRange(lastRow, statusColIndex).setValue('PENDING');
    }
    
    // Send PENDING confirmation email with WRS number
    sendPendingConfirmationEmail(
      formData.email, 
      'WLP', 
      formData.wrsNumber, 
      '',  // No engine number for WLP
      formData.branch,
      formData.dealerMechCode
    );
    
    return {
      success: true,
      message: 'WLP acknowledgment submitted successfully! WRS Number: ' + formData.wrsNumber + '. A confirmation email has been sent to ' + formData.email + '.'
    };
    
  } catch (error) {
    logError_('WLP submission error: ' + error.message, '', '');
    return {
      success: false,
      message: 'Server error: ' + error.message
    };
  }
}

// ════════════════════════════════════════════════════════════
//  OPTIONAL: DYNAMIC DROPDOWN DATA
// ════════════════════════════════════════════════════════════

/**
 * Get list of provinces for dropdown population
 * @returns {Array<string>}
 */
function getProvinces() {
  return [
    'Metro Manila', 'Cebu', 'Davao', 'Rizal', 'Bulacan', 'Cavite', 'Laguna',
    'Pampanga', 'Batangas', 'Quezon', 'Pangasinan', 'Iloilo', 'Negros Occidental',
    'Leyte', 'Misamis Oriental', 'Albay', 'Cagayan', 'Isabela', 'Nueva Ecija',
    'Tarlac', 'Zambales', 'Palawan', 'Bohol', 'Samar', 'Cotabato'
  ];
}

// ═══════════════════════════════════════════════════════════════
//  TESTING & DEBUGGING UTILITIES
// ═══════════════════════════════════════════════════════════════

/**
 * TEST FUNCTION: Simulate STATUS change to test email trigger
 * 
 * HOW TO USE:
 * 1. Open Apps Script Editor
 * 2. Select this function from dropdown
 * 3. Click Run
 * 4. Check your email and Log sheet
 */
function TEST_EmailTrigger() {
  const TEST_CONFIG = {
    sheetName: 'WRC',  // Change to 'FSC' or 'WLP' to test those
    testRow: 2,        // Row to use for testing (must have email address)
    testEmail: 'kenji.devcodes@gmail.com'  // Override email for testing
  };
  
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('  EMAIL TRIGGER TEST - STARTING');
  Logger.log('═══════════════════════════════════════════════');
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(TEST_CONFIG.sheetName);
    
    if (!sheet) {
      Logger.log('✗ ERROR: Sheet not found: ' + TEST_CONFIG.sheetName);
      return;
    }
    
    // Find columns
    const statusCol = findColumnByHeader_(sheet, 'STATUS');
    const emailCol = findColumnByHeader_(sheet, 'EMAIL');
    
    Logger.log(`✓ Found STATUS column: ${statusCol}`);
    Logger.log(`✓ Found EMAIL column: ${emailCol}`);
    
    if (statusCol === 0 || emailCol === 0) {
      Logger.log('✗ ERROR: Required columns not found');
      return;
    }
    
    // Get current values
    const currentEmail = sheet.getRange(TEST_CONFIG.testRow, emailCol).getValue();
    const identifier = sheet.getRange(TEST_CONFIG.testRow, 1).getValue();
    
    Logger.log(`Row ${TEST_CONFIG.testRow} identifier: ${identifier}`);
    Logger.log(`Row ${TEST_CONFIG.testRow} email: ${currentEmail}`);
    
    // Use test email or actual email
    const targetEmail = TEST_CONFIG.testEmail || currentEmail;
    
    if (!targetEmail) {
      Logger.log('✗ ERROR: No email address available for testing');
      return;
    }
    
    // Temporarily set test email if different
    if (TEST_CONFIG.testEmail && TEST_CONFIG.testEmail !== currentEmail) {
      sheet.getRange(TEST_CONFIG.testRow, emailCol).setValue(TEST_CONFIG.testEmail);
      Logger.log(`✓ Set test email: ${TEST_CONFIG.testEmail}`);
    }
    
    // Get WRC and Engine numbers from test row
    let wrcNum = '';
    let engineNum = '';
    if (TEST_CONFIG.sheetName === 'WRC') {
      wrcNum = sheet.getRange(TEST_CONFIG.testRow, 1).getValue();  // WRC No.
      engineNum = sheet.getRange(TEST_CONFIG.testRow, 2).getValue();  // Engine No.
    } else if (TEST_CONFIG.sheetName === 'FSC') {
      const wrcCol = findColumnByHeader_(sheet, 'WRC NUMBER');
      const frameCol = findColumnByHeader_(sheet, 'FRAME NUMBER');
      if (wrcCol > 0) wrcNum = sheet.getRange(TEST_CONFIG.testRow, wrcCol).getValue();
      if (frameCol > 0) engineNum = sheet.getRange(TEST_CONFIG.testRow, frameCol).getValue();
    } else if (TEST_CONFIG.sheetName === 'WLP') {
      wrcNum = sheet.getRange(TEST_CONFIG.testRow, 1).getValue();  // WRS Number
    }
    
    // Send test email
    Logger.log(`→ Sending test email to: ${targetEmail}`);
    Logger.log(`   WRC/WRS Number: ${wrcNum}`);
    Logger.log(`   Engine/Frame Number: ${engineNum}`);
    
    const result = sendTransmittedEmail_(
      targetEmail,
      TEST_CONFIG.sheetName,
      wrcNum,
      engineNum,
      'Test Branch',
      'Test Dealer'
    );
    
    if (result.success) {
      Logger.log('═══════════════════════════════════════════════');
      Logger.log('  ✓ SUCCESS! Email sent via ' + result.method);
      Logger.log('  Check inbox: ' + targetEmail);
      Logger.log('═══════════════════════════════════════════════');
    } else {
      Logger.log('═══════════════════════════════════════════════');
      Logger.log('  ✗ FAILED: ' + result.error);
      Logger.log('═══════════════════════════════════════════════');
    }
    
    // Restore original email if changed
    if (TEST_CONFIG.testEmail && TEST_CONFIG.testEmail !== currentEmail) {
      sheet.getRange(TEST_CONFIG.testRow, emailCol).setValue(currentEmail);
      Logger.log(`✓ Restored original email: ${currentEmail}`);
    }
    
  } catch (error) {
    Logger.log('═══════════════════════════════════════════════');
    Logger.log('  ✗ TEST EXCEPTION: ' + error.message);
    Logger.log('═══════════════════════════════════════════════');
  }
}

/**
 * DIAGNOSTIC: Check trigger installation and permissions
 */
function DIAGNOSTIC_CheckTriggerSetup() {
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('  TRIGGER DIAGNOSTIC');
  Logger.log('═══════════════════════════════════════════════');
  
  // Check installed triggers
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`\nInstalled Triggers: ${triggers.length}`);
  
  let hasOnEditTrigger = false;
  triggers.forEach((trigger, index) => {
    Logger.log(`\nTrigger ${index + 1}:`);
    Logger.log(`  - Handler: ${trigger.getHandlerFunction()}`);
    Logger.log(`  - Event Type: ${trigger.getEventType()}`);
    Logger.log(`  - Source: ${trigger.getTriggerSource()}`);
    
    if (trigger.getHandlerFunction() === 'onEdit' && 
        trigger.getEventType() === ScriptApp.EventType.ON_EDIT) {
      hasOnEditTrigger = true;
    }
  });
  
  Logger.log('\n═══════════════════════════════════════════════');
  if (hasOnEditTrigger) {
    Logger.log('  ✓ onEdit INSTALLABLE trigger found');
  } else {
    Logger.log('  ✗ WARNING: No installable onEdit trigger found!');
    Logger.log('  → You MUST install trigger for emails to work');
    Logger.log('  → Go to: Triggers → + Add Trigger → onEdit');
  }
  Logger.log('═══════════════════════════════════════════════');
  
  // Check permissions
  Logger.log('\nPermission Check:');
  try {
    const testEmail = Session.getActiveUser().getEmail();
    Logger.log(`  ✓ Can access user email: ${testEmail}`);
  } catch (e) {
    Logger.log('  ✗ Cannot access user info');
  }
  
  // Check spreadsheet access
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log(`  ✓ Can access spreadsheet: ${ss.getName()}`);
  } catch (e) {
    Logger.log('  ✗ Cannot access spreadsheet: ' + e.message);
  }
  
  // Check email quota
  try {
    const quota = MailApp.getRemainingDailyQuota();
    Logger.log(`  ✓ Email quota remaining today: ${quota}`);
    if (quota === 0) {
      Logger.log('  ⚠ WARNING: Email quota exhausted!');
    }
  } catch (e) {
    Logger.log('  ? Cannot check email quota: ' + e.message);
  }
  
  Logger.log('\n═══════════════════════════════════════════════');
}

/**
 * DIAGNOSTIC: View recent logs
 */
function DIAGNOSTIC_ViewRecentLogs() {
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('  RECENT LOG ENTRIES');
  Logger.log('═══════════════════════════════════════════════');
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName('Log');
    
    if (!logSheet) {
      Logger.log('  No Log sheet found');
      return;
    }
    
    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('  Log sheet is empty');
      return;
    }
    
    // Get last 10 rows
    const startRow = Math.max(2, lastRow - 9);
    const numRows = lastRow - startRow + 1;
    const data = logSheet.getRange(startRow, 1, numRows, 6).getValues();
    
    data.forEach((row, index) => {
      Logger.log(`\n[${startRow + index}] ${row[0]}`);
      Logger.log(`  Level: ${row[1]}`);
      Logger.log(`  Sheet: ${row[2]} | Row: ${row[3]}`);
      Logger.log(`  Message: ${row[4]}`);
    });
    
    Logger.log('\n═══════════════════════════════════════════════');
    Logger.log(`  Showing last ${numRows} entries of ${lastRow - 1} total`);
    Logger.log('═══════════════════════════════════════════════');
    
  } catch (error) {
    Logger.log('  Error reading logs: ' + error.message);
  }
}

/**
 * UTILITY: Clear old log entries (keep last 100)
 */
function UTILITY_CleanupOldLogs() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName('Log');
    
    if (!logSheet) {
      Logger.log('No Log sheet found');
      return;
    }
    
    const lastRow = logSheet.getLastRow();
    if (lastRow <= 101) {
      Logger.log('Less than 100 log entries - no cleanup needed');
      return;
    }
    
    const rowsToDelete = lastRow - 101; // Keep last 100 + header
    logSheet.deleteRows(2, rowsToDelete);
    
    Logger.log(`✓ Deleted ${rowsToDelete} old log entries. Kept last 100.`);
    
  } catch (error) {
    Logger.log('Error cleaning logs: ' + error.message);
  }
}

/**
 * Get list of dealer codes
 * You can modify this to read from a "Dealers" sheet
 * @returns {Array<string>}
 */
function getDealerCodes() {
  // Example: Read from a separate sheet
  // const dealerSheet = getSheet_('Dealers');
  // const data = dealerSheet.getRange(2, 1, dealerSheet.getLastRow() - 1, 1).getValues();
  // return data.map(row => row[0]).filter(code => code);
  
  // Sample static data
  return ['D001', 'D002', 'D003', 'D004', 'D005'];
}

// ════════════════════════════════════════════════════════════
//  FILE UPLOAD HANDLER
// ════════════════════════════════════════════════════════════

/**
 * Upload file to Google Drive and return public URL
 * Creates a folder structure: FSC_WRC_WLP_Files > [FORM_TYPE]
 * @param {Object} fileData - Contains name, type, data (base64), formType
 * @returns {Object} - {success: boolean, url: string, message: string}
 */
function uploadFile(fileData) {
  try {
    // Get or create root folder
    const rootFolderName = 'FSC_WRC_WLP_Files';
    let rootFolder;
    const folders = DriveApp.getFoldersByName(rootFolderName);
    if (folders.hasNext()) {
      rootFolder = folders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }
    
    // Get or create form-specific subfolder
    const formFolderName = fileData.formType; // WRC, FSC, or WLP
    let formFolder;
    const subFolders = rootFolder.getFoldersByName(formFolderName);
    if (subFolders.hasNext()) {
      formFolder = subFolders.next();
    } else {
      formFolder = rootFolder.createFolder(formFolderName);
    }
    
    // Decode base64 data and create blob
    const bytes = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(bytes, fileData.type, fileData.name);
    
    // Create file in Drive
    const file = formFolder.createFile(blob);
    
    // Make file publicly accessible (optional - adjust based on security needs)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Return file URL
    const fileUrl = file.getUrl();
    
    return {
      success: true,
      url: fileUrl,
      message: 'File uploaded successfully',
      fileId: file.getId()
    };
    
  } catch (error) {
    Logger.log('File upload error: ' + error.toString());
    return {
      success: false,
      url: '',
      message: 'File upload failed: ' + error.toString()
    };
  }
}

// ════════════════════════════════════════════════════════════
//  AUTOMATED EMAIL NOTIFICATIONS
// ════════════════════════════════════════════════════════════

/**
 * Send PENDING confirmation email immediately after form submission
 * 
 * @param {string} email - Recipient email address
 * @param {string} formType - WRC, FSC, or WLP
 * @param {string} wrcNumber - WRC Number or WRS Number
 * @param {string} engineNumber - Engine Number or Frame Number (optional)
 * @param {string} branch - Branch name (optional)
 * @param {string} dealerName - Dealer name (optional)
 */
function sendPendingConfirmationEmail(email, formType, wrcNumber, engineNumber, branch, dealerName) {
  try {
    if (!email || String(email).trim() === '') {
      Logger.log('Cannot send pending email: empty email address');
      return;
    }
    
    const subject = 'Application Received - PENDING';
    
    // Build dynamic field rows
    let fieldRows = '';
    if (formType === 'WRC') {
      fieldRows = `
        <tr><td class="field-label">WRC Number:</td><td class="field-value">${wrcNumber || 'N/A'}</td></tr>
        <tr><td class="field-label">Engine Number:</td><td class="field-value">${engineNumber || 'N/A'}</td></tr>
      `;
    } else if (formType === 'FSC') {
      fieldRows = `
        <tr><td class="field-label">WRC Number:</td><td class="field-value">${wrcNumber || 'N/A'}</td></tr>
        <tr><td class="field-label">Frame Number:</td><td class="field-value">${engineNumber || 'N/A'}</td></tr>
      `;
    } else if (formType === 'WLP') {
      fieldRows = `
        <tr><td class="field-label">WRS Number:</td><td class="field-value">${wrcNumber || 'N/A'}</td></tr>
      `;
    }
    
    if (branch) {
      fieldRows += `<tr><td class="field-label">Branch:</td><td class="field-value">${branch}</td></tr>`;
    }
    if (dealerName) {
      fieldRows += `<tr><td class="field-label">Dealer Name:</td><td class="field-value">${dealerName}</td></tr>`;
    }
    
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 600px;
            margin: 0 auto;
            padding: 0;
            background-color: #f4f4f4;
          }
          .container {
            background-color: #ffffff;
            margin: 20px auto;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          .logo {
            text-align: center;
            padding: 20px;
            background-color: #ffffff;
          }
          .logo img {
            max-width: 200px;
            height: auto;
          }
          .header {
            background-color: #fff3cd;
            color: #856404;
            padding: 25px;
            text-align: center;
            border-bottom: 4px solid #ffc107;
          }
          .header h2 {
            margin: 0;
            font-size: 24px;
          }
          .content {
            padding: 30px;
          }
          .info-table {
            width: 100%;
            background-color: #f8f9fa;
            border-radius: 6px;
            margin: 20px 0;
            border-collapse: collapse;
          }
          .info-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #e9ecef;
          }
          .info-table tr:last-child td {
            border-bottom: none;
          }
          .field-label {
            font-weight: bold;
            color: #555;
            width: 40%;
          }
          .field-value {
            color: #333;
          }
          .status-box {
            background-color: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 15px;
            margin: 20px 0;
            border-radius: 4px;
          }
          .footer {
            background-color: #f8f9fa;
            padding: 20px;
            text-align: center;
            font-size: 0.9em;
            color: #666;
            border-top: 1px solid #ddd;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="logo">
            <img src="https://i.imgur.com/nGVvR3v.png" alt="Giant Moto Pro Logo" />
          </div>
          
          <div class="header">
            <h2>Application Status: PENDING</h2>
          </div>
          
          <div class="content">
            <p>Dear Applicant,</p>
            
            <p>Thank you for your submission to the <strong>FSC / WRC / WLP Management System</strong>.</p>
            
            <table class="info-table">
              <tr><td class="field-label">Form Type:</td><td class="field-value"><strong>${formType}</strong></td></tr>
              ${fieldRows}
              <tr><td class="field-label">Status:</td><td class="field-value"><strong>PENDING</strong></td></tr>
            </table>
            
            <div class="status-box">
              <strong>What's Next?</strong>
              <p style="margin: 10px 0 0 0;">Your application is currently being reviewed and will be processed soon. You will receive another email notification when your application is transmitted to Kawasaki.</p>
            </div>
            
            <p>Thank you for your patience and for choosing Giant Moto Pro.</p>
            
            <p>Best regards,<br>
            <strong>Giant Moto Pro Team</strong></p>
          </div>
          
          <div class="footer">
            <p><strong>This is an automated message.</strong> Please do not reply to this email.</p>
            <p>If you have questions, please contact your branch office.</p>
          </div>
        </div>
      </body>
      </html>
    `;
    
    const plainBody = `
APPLICATION RECEIVED - PENDING

Dear Applicant,

Thank you for your submission to the FSC / WRC / WLP Management System.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
APPLICATION DETAILS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Form Type: ${formType}
${formType === 'WRC' ? 'WRC Number: ' + (wrcNumber || 'N/A') : ''}
${formType === 'WRC' ? 'Engine Number: ' + (engineNumber || 'N/A') : ''}
${formType === 'FSC' ? 'WRC Number: ' + (wrcNumber || 'N/A') : ''}
${formType === 'FSC' ? 'Frame Number: ' + (engineNumber || 'N/A') : ''}
${formType === 'WLP' ? 'WRS Number: ' + (wrcNumber || 'N/A') : ''}
${branch ? 'Branch: ' + branch : ''}
${dealerName ? 'Dealer Name: ' + dealerName : ''}
Status: PENDING

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Your application is currently being reviewed and will be processed soon.
You will receive another email when your application is transmitted to Kawasaki.

Thank you for your patience and for choosing Giant Moto Pro.

Best regards,
Giant Moto Pro Team

---
This is an automated message. Please do not reply to this email.
If you have questions, please contact your branch office.
    `;
    
    // Send email using GmailApp (preferred) or MailApp
    try {
      GmailApp.sendEmail(email, subject, plainBody, {
        htmlBody: htmlBody,
        name: 'Giant Moto Pro - FSC/WRC/WLP System'
      });
      Logger.log('PENDING confirmation email sent to: ' + email);
    } catch (e) {
      // Fallback to MailApp if GmailApp fails
      MailApp.sendEmail(email, subject, plainBody, {
        htmlBody: htmlBody,
        name: 'Giant Moto Pro - FSC/WRC/WLP System'
      });
      Logger.log('PENDING confirmation email sent to: ' + email + ' (via MailApp)');
    }
    
  } catch (error) {
    Logger.log('PENDING email send error: ' + error.toString());
    logError_('PENDING email notification error: ' + error.message, wrcNumber || '', engineNumber || '');
  }
}

/**
 * Installable trigger function - automatically runs when any cell is edited
 * Monitors STATUS column changes and sends email notifications
 * 
 * TO INSTALL TRIGGER:
 * 1. Open Apps Script Editor
 * 2. Click on "Triggers" (clock icon) in left sidebar
 * 3. Click "+ Add Trigger" button
 * 4. Choose:
 *    - Function: onEdit
 *    - Event source: From spreadsheet
 *    - Event type: On edit
 * 5. Save
 */
/**
 * ═══════════════════════════════════════════════════════════════
 *  INSTALLABLE onEdit TRIGGER - STATUS CHANGE EMAIL NOTIFICATION
 *  
 *  CRITICAL: This must be installed as an INSTALLABLE trigger
 *  (not a simple trigger) to have Gmail/email permissions.
 *  
 *  Setup: Triggers → + Add Trigger → onEdit → On edit
 * ═══════════════════════════════════════════════════════════════
 */
function onEdit(e) {
  const TRIGGER_STATUS = 'TRANSMITTED TO KAWASAKI';
  const LOG_SHEET = 'Log';
  
  try {
    // ─────────────────────────────────────────────────────────────
    // 1. VALIDATE EVENT OBJECT
    // ─────────────────────────────────────────────────────────────
    if (!e || !e.range) {
      logToSheet_(LOG_SHEET, 'ERROR', 'onEdit', '', 'Invalid event object - trigger may not be properly installed');
      return;
    }
    
    // ─────────────────────────────────────────────────────────────
    // 2. GET EDIT DETAILS
    // ─────────────────────────────────────────────────────────────
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    const row = range.getRow();
    const col = range.getColumn();
    const newValue = range.getValue();
    const oldValue = e.oldValue || ''; // Previous value before edit
    
    // ─────────────────────────────────────────────────────────────
    // 3. FILTER: ONLY PROCESS WRC/FSC/WLP SHEETS
    // ─────────────────────────────────────────────────────────────
    if (!['WRC', 'FSC', 'WLP'].includes(sheetName)) {
      return; // Not a relevant sheet
    }
    
    // ─────────────────────────────────────────────────────────────
    // 4. SKIP HEADER ROW
    // ─────────────────────────────────────────────────────────────
    if (row === 1) {
      return; // Don't process header edits
    }
    
    // ─────────────────────────────────────────────────────────────
    // 5. FIND COLUMN POSITIONS DYNAMICALLY
    // ─────────────────────────────────────────────────────────────
    const statusCol = findColumnByHeader_(sheet, 'STATUS');
    const emailCol = findColumnByHeader_(sheet, 'EMAIL');
    const branchCol = findColumnByHeader_(sheet, 'BRANCH');
    
    if (statusCol === 0) {
      logToSheet_(LOG_SHEET, 'ERROR', sheetName, row, 'STATUS column not found in sheet');
      return;
    }
    
    if (emailCol === 0) {
      logToSheet_(LOG_SHEET, 'ERROR', sheetName, row, 'EMAIL column not found in sheet');
      return;
    }
    
    // ─────────────────────────────────────────────────────────────
    // 6. CHECK IF EDITED CELL IS STATUS COLUMN
    // ─────────────────────────────────────────────────────────────
    if (col !== statusCol) {
      return; // Not editing STATUS column, exit
    }
    
    // ─────────────────────────────────────────────────────────────
    // 7. CHECK IF NEW VALUE IS TARGET STATUS
    // ─────────────────────────────────────────────────────────────
    const newValueNormalized = String(newValue).trim().toUpperCase();
    if (newValueNormalized !== TRIGGER_STATUS.toUpperCase()) {
      return; // Not the status we're looking for
    }
    
    // ─────────────────────────────────────────────────────────────
    // 8. PREVENT DUPLICATE SENDS (CHECK OLD VALUE)
    // ─────────────────────────────────────────────────────────────
    const oldValueNormalized = String(oldValue).trim().toUpperCase();
    if (oldValueNormalized === TRIGGER_STATUS.toUpperCase()) {
      logToSheet_(LOG_SHEET, 'SKIPPED', sheetName, row, 
        'Status already was TRANSMITTED - prevented duplicate email send');
      return;
    }
    
    // ─────────────────────────────────────────────────────────────
    // 9. CHECK FOR DUPLICATE SEND FLAG (OPTIONAL SAFETY)
    // ─────────────────────────────────────────────────────────────
    let emailSentCol = findColumnByHeader_(sheet, 'EMAIL_SENT');
    if (emailSentCol > 0) {
      const emailSentFlag = sheet.getRange(row, emailSentCol).getValue();
      if (emailSentFlag === 'YES' || emailSentFlag === true) {
        logToSheet_(LOG_SHEET, 'SKIPPED', sheetName, row, 
          'EMAIL_SENT flag already set - prevented duplicate');
        return;
      }
    }
    
    // ─────────────────────────────────────────────────────────────
    // 10. GET IDENTIFIER AND ENGINE NUMBER BASED ON SHEET TYPE
    // ─────────────────────────────────────────────────────────────
    let wrcNumber = '';
    let engineNumber = '';
    let dealerName = '';
    
    if (sheetName === 'WRC') {
      // WRC: Column A = WRC No., Column B = Engine No., Column J = Dealer Name
      wrcNumber = String(sheet.getRange(row, 1).getValue()).trim();
      engineNumber = String(sheet.getRange(row, 2).getValue()).trim();
      const dealerCol = findColumnByHeader_(sheet, 'DEALER NAME');
      if (dealerCol > 0) {
        dealerName = String(sheet.getRange(row, dealerCol).getValue()).trim();
      }
    } else if (sheetName === 'FSC') {
      // FSC: Column C = WRC Number, Column D = Frame Number, Column B = Dealer/Mechanic Code
      const wrcCol = findColumnByHeader_(sheet, 'WRC NUMBER') || findColumnByHeader_(sheet, 'WRC Number');
      const frameCol = findColumnByHeader_(sheet, 'FRAME NUMBER') || findColumnByHeader_(sheet, 'Frame Number');
      const dealerCol = findColumnByHeader_(sheet, 'DEALER/MECHANIC CODE') || findColumnByHeader_(sheet, 'Dealer/Mechanic Code');
      
      if (wrcCol > 0) wrcNumber = String(sheet.getRange(row, wrcCol).getValue()).trim();
      if (frameCol > 0) engineNumber = String(sheet.getRange(row, frameCol).getValue()).trim();
      if (dealerCol > 0) dealerName = String(sheet.getRange(row, dealerCol).getValue()).trim();
    } else if (sheetName === 'WLP') {
      // WLP: Column A = WRS Number, Column F = Dealer/Mechanic Code
      wrcNumber = String(sheet.getRange(row, 1).getValue()).trim();  // WRS Number
      const dealerCol = findColumnByHeader_(sheet, 'DEALER/MECHANIC CODE') || findColumnByHeader_(sheet, 'Dealer/Mechanic Code');
      if (dealerCol > 0) {
        dealerName = String(sheet.getRange(row, dealerCol).getValue()).trim();
      }
    }
    
    // ─────────────────────────────────────────────────────────────
    // 11. GET EMAIL ADDRESS AND BRANCH
    // ─────────────────────────────────────────────────────────────
    const email = String(sheet.getRange(row, emailCol).getValue()).trim();
    const branch = branchCol > 0 ? String(sheet.getRange(row, branchCol).getValue()).trim() : '';
    
    // ─────────────────────────────────────────────────────────────
    // 12. VALIDATE EMAIL ADDRESS
    // ─────────────────────────────────────────────────────────────
    if (!email || email === '') {
      logToSheet_(LOG_SHEET, 'ERROR', sheetName, row, 
        `Missing email address for WRC/WRS Number: ${wrcNumber}`);
      return;
    }
    
    // Simple email validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      logToSheet_(LOG_SHEET, 'ERROR', sheetName, row, 
        `Invalid email format: ${email} for WRC/WRS Number: ${wrcNumber}`);
      return;
    }
    
    // ─────────────────────────────────────────────────────────────
    // 13. SEND EMAIL NOTIFICATION WITH WRC/ENGINE NUMBERS
    // ─────────────────────────────────────────────────────────────
    const emailResult = sendTransmittedEmail_(email, sheetName, wrcNumber, engineNumber, branch, dealerName);
    
    if (emailResult.success) {
      // Log success
      logToSheet_(LOG_SHEET, 'SUCCESS', sheetName, row, 
        `Email sent to ${email} for WRC/WRS: ${wrcNumber}, Engine/Frame: ${engineNumber || 'N/A'}`);
      
      // Set EMAIL_SENT flag if column exists
      if (emailSentCol > 0) {
        sheet.getRange(row, emailSentCol).setValue('YES');
      }
      
      Logger.log(`✓ Email sent successfully to ${email} (${sheetName} row ${row})`);
      
    } else {
      // Log failure
      logToSheet_(LOG_SHEET, 'ERROR', sheetName, row, 
        `Email send failed: ${emailResult.error} - Email: ${email}, WRC/WRS: ${wrcNumber}`);
      
      Logger.log(`✗ Email send failed for ${email}: ${emailResult.error}`);
    }
    
  } catch (error) {
    // Catch-all error handler
    const errorMsg = `onEdit exception: ${error.message} | Stack: ${error.stack}`;
    logToSheet_(LOG_SHEET, 'CRITICAL', 'onEdit', '', errorMsg);
    Logger.log('CRITICAL ERROR in onEdit: ' + error.toString());
  }
}

/**
 * ═══════════════════════════════════════════════════════════════
 *  SEND TRANSMITTED EMAIL - ISOLATED FUNCTION FOR TESTING
 * ═══════════════════════════════════════════════════════════════
 */
function sendTransmittedEmail_(email, formType, wrcNumber, engineNumber, branch, dealerName) {
  try {
    const subject = 'Application Status: TRANSMITTED TO KAWASAKI';
    
    // Build dynamic field rows
    let fieldRows = '';
    if (formType === 'WRC') {
      fieldRows = `
        <tr><td class="field-label">WRC Number:</td><td class="field-value">${wrcNumber || 'N/A'}</td></tr>
        <tr><td class="field-label">Engine Number:</td><td class="field-value">${engineNumber || 'N/A'}</td></tr>
      `;
    } else if (formType === 'FSC') {
      fieldRows = `
        <tr><td class="field-label">WRC Number:</td><td class="field-value">${wrcNumber || 'N/A'}</td></tr>
        <tr><td class="field-label">Frame Number:</td><td class="field-value">${engineNumber || 'N/A'}</td></tr>
      `;
    } else if (formType === 'WLP') {
      fieldRows = `
        <tr><td class="field-label">WRS Number:</td><td class="field-value">${wrcNumber || 'N/A'}</td></tr>
      `;
    }
    
    if (branch) {
      fieldRows += `<tr><td class="field-label">Branch:</td><td class="field-value">${branch}</td></tr>`;
    }
    if (dealerName) {
      fieldRows += `<tr><td class="field-label">Dealer Name:</td><td class="field-value">${dealerName}</td></tr>`;
    }
    
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 600px;
            margin: 0 auto;
            padding: 0;
            background-color: #f4f4f4;
          }
          .container {
            background-color: #ffffff;
            margin: 20px auto;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          .logo {
            text-align: center;
            padding: 20px;
            background-color: #ffffff;
          }
          .logo img {
            max-width: 200px;
            height: auto;
          }
          .header {
            background-color: #28a745;
            color: white;
            padding: 25px;
            text-align: center;
          }
          .header h2 {
            margin: 0;
            font-size: 24px;
          }
          .content {
            padding: 30px;
          }
          .info-table {
            width: 100%;
            background-color: #f8f9fa;
            border-radius: 6px;
            margin: 20px 0;
            border-collapse: collapse;
          }
          .info-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #e9ecef;
          }
          .info-table tr:last-child td {
            border-bottom: none;
          }
          .field-label {
            font-weight: bold;
            color: #555;
            width: 40%;
          }
          .field-value {
            color: #333;
          }
          .success-box {
            background-color: #d4edda;
            border-left: 4px solid #28a745;
            padding: 15px;
            margin: 20px 0;
            border-radius: 4px;
            color: #155724;
          }
          .footer {
            background-color: #f8f9fa;
            padding: 20px;
            text-align: center;
            font-size: 0.9em;
            color: #666;
            border-top: 1px solid #ddd;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="logo">
            <img src="https://i.imgur.com/nGVvR3v.png" alt="Giant Moto Pro Logo" />
          </div>
          
          <div class="header">
            <h2>✓ Application Status: TRANSMITTED TO KAWASAKI</h2>
          </div>
          
          <div class="content">
            <p>Dear Customer,</p>
            
            <p>We are pleased to inform you that your application has been successfully processed and <strong>transmitted to Kawasaki</strong>.</p>
            
            <table class="info-table">
              <tr><td class="field-label">Form Type:</td><td class="field-value"><strong>${formType}</strong></td></tr>
              ${fieldRows}
              <tr><td class="field-label">Status:</td><td class="field-value"><strong>TRANSMITTED TO KAWASAKI</strong></td></tr>
            </table>
            
            <div class="success-box">
              <strong>✓ What's Next?</strong>
              <p style="margin: 10px 0 0 0;">Your application is now with Kawasaki for final processing. You will be contacted if any additional information is required.</p>
            </div>
            
            <p>Thank you for your patience and for choosing Giant Moto Pro.</p>
            
            <p>Best regards,<br>
            <strong>Giant Moto Pro Team</strong></p>
          </div>
          
          <div class="footer">
            <p><strong>This is an automated notification</strong> from the FSC/WRC/WLP Management System.</p>
            <p>If you have questions, please contact your branch office.</p>
          </div>
        </div>
      </body>
      </html>
    `;
    
    const plainBody = `
STATUS UPDATE - TRANSMITTED TO KAWASAKI

Dear Customer,

We are pleased to inform you that your application has been successfully processed and transmitted to Kawasaki.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
APPLICATION DETAILS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Form Type: ${formType}
${formType === 'WRC' ? 'WRC Number: ' + (wrcNumber || 'N/A') : ''}
${formType === 'WRC' ? 'Engine Number: ' + (engineNumber || 'N/A') : ''}
${formType === 'FSC' ? 'WRC Number: ' + (wrcNumber || 'N/A') : ''}
${formType === 'FSC' ? 'Frame Number: ' + (engineNumber || 'N/A') : ''}
${formType === 'WLP' ? 'WRS Number: ' + (wrcNumber || 'N/A') : ''}
${branch ? 'Branch: ' + branch : ''}
${dealerName ? 'Dealer Name: ' + dealerName : ''}
Status: TRANSMITTED TO KAWASAKI

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Your application is now with Kawasaki for final processing. You will be contacted if any additional information is required.

Thank you for your patience and for choosing Giant Moto Pro.

Best regards,
Giant Moto Pro Team

---
This is an automated notification from the FSC/WRC/WLP Management System.
If you have questions, please contact your branch office.
    `;
    
    // Try GmailApp first, fallback to MailApp
    try {
      GmailApp.sendEmail(email, subject, plainBody, {
        htmlBody: htmlBody,
        name: 'Giant Moto Pro - FSC/WRC/WLP System'
      });
      return { success: true, method: 'GmailApp' };
    } catch (gmailError) {
      // Fallback to MailApp
      try {
        MailApp.sendEmail(email, subject, plainBody, {
          htmlBody: htmlBody,
          name: 'Giant Moto Pro - FSC/WRC/WLP System'
        });
        return { success: true, method: 'MailApp' };
      } catch (mailError) {
        return { 
          success: false, 
          error: `Both GmailApp and MailApp failed. GmailApp: ${gmailError.message}, MailApp: ${mailError.message}` 
        };
      }
    }
    
  } catch (error) {
    return { 
      success: false, 
      error: error.message 
    };
  }
}

/**
 * ═══════════════════════════════════════════════════════════════
 *  ENHANCED LOGGING TO LOG SHEET
 * ═══════════════════════════════════════════════════════════════
 */
function logToSheet_(logSheetName, level, sheetName, row, message) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(logSheetName);
    
    // Create Log sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet(logSheetName);
      logSheet.getRange(1, 1, 1, 6).setValues([[
        'Timestamp', 'Level', 'Sheet', 'Row', 'Message', 'User'
      ]]);
      logSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f0f0f0');
      logSheet.setFrozenRows(1);
    }
    
    // Append log entry
    logSheet.appendRow([
      new Date(),
      level,
      sheetName || '',
      row || '',
      message,
      Session.getActiveUser().getEmail()
    ]);
    
    // Auto-format timestamp column
    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // Color code by level
    const colors = {
      'SUCCESS': '#d4edda',
      'ERROR': '#f8d7da',
      'CRITICAL': '#ff0000',
      'SKIPPED': '#fff3cd',
      'INFO': '#d1ecf1'
    };
    if (colors[level]) {
      logSheet.getRange(lastRow, 2).setBackground(colors[level]);
    }
    
  } catch (error) {
    // Fallback to console if logging fails
    Logger.log(`LOG ERROR: ${error.message} | Original: ${level} - ${message}`);
  }
}

/**
 * Send status notification email to customer
 * 
 * @param {string} email - Recipient email address
 * @param {string} formType - WRC, FSC, or WLP
 * @param {string} identifier - WRC No. or WRS No.
 * @param {string} identifierName - Display name for identifier
 */
function sendStatusNotificationEmail(email, formType, identifier, identifierName) {
  try {
    const subject = 'Your Application Status: TRANSMITTED TO KAWASAKI';
    
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
          }
          .header {
            background-color: #f8f9fa;
            padding: 20px;
            border-left: 4px solid #007bff;
            margin-bottom: 20px;
          }
          .header h2 {
            margin: 0;
            color: #007bff;
          }
          .content {
            padding: 20px 0;
          }
          .info-box {
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 6px;
            margin: 15px 0;
          }
          .info-label {
            font-weight: bold;
            color: #555;
          }
          .footer {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            font-size: 0.9em;
            color: #666;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>Application Status Update</h2>
        </div>
        
        <div class="content">
          <p>Dear Customer,</p>
          
          <p>We are pleased to inform you that your application submitted via our <strong>FSC / WRC / WLP Management System</strong> has been successfully transmitted to Kawasaki.</p>
          
          <div class="info-box">
            <p><span class="info-label">Form Type:</span> ${formType}</p>
            <p><span class="info-label">${identifierName}:</span> ${identifier}</p>
            <p><span class="info-label">Status:</span> TRANSMITTED TO KAWASAKI</p>
          </div>
          
          <p>Your submission is now being processed by Kawasaki. You will be contacted if any additional information is required.</p>
          
          <p>Thank you for your patience and for choosing our services.</p>
        </div>
        
        <div class="footer">
          <p>This is an automated notification from the FSC/WRC/WLP Management System.</p>
          <p>If you have any questions or concerns, please contact your dealer branch.</p>
        </div>
      </body>
      </html>
    `;
    
    const plainBody = `
Application Status Update

Dear Customer,

We are pleased to inform you that your application submitted via our FSC / WRC / WLP Management System has been successfully transmitted to Kawasaki.

Form Type: ${formType}
${identifierName}: ${identifier}
Status: TRANSMITTED TO KAWASAKI

Your submission is now being processed by Kawasaki. You will be contacted if any additional information is required.

Thank you for your patience and for choosing our services.

---
This is an automated notification from the FSC/WRC/WLP Management System.
If you have any questions or concerns, please contact your dealer branch.
    `;
    
    // Send email using GmailApp (preferred) or MailApp
    try {
      GmailApp.sendEmail(email, subject, plainBody, {
        htmlBody: htmlBody,
        name: 'FSC/WRC/WLP System'
      });
    } catch (e) {
      // Fallback to MailApp if GmailApp fails
      MailApp.sendEmail(email, subject, plainBody, {
        htmlBody: htmlBody,
        name: 'FSC/WRC/WLP System'
      });
    }
    
    Logger.log('Email sent successfully to: ' + email);
    
  } catch (error) {
    Logger.log('Email send error: ' + error.toString());
    logError_('Email notification error: ' + error.message, identifier || '', '');
  }
}

