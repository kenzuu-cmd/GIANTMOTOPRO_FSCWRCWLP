/**
 * ============================================================
 *  FSC / WRC / WLP — Google Apps Script Backend
 * ============================================================
 *
 *  This file contains server-side functions for the FormApp.html
 *  web application. It handles form submissions, validates data,
 *  enforces business rules, and writes to Google Sheets.
 *
 *  SETUP INSTRUCTIONS:
 *  1. Replace SPREADSHEET_ID with your Google Sheets ID
 *  2. Ensure your spreadsheet has sheets: WRC, FSC, WLP
 *  3. Deploy as web app or use with sidebar/dialog
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
    const sheet = getSheet_(SHEET_WRC, [
      'WRC No.', 'ENGINE No.', 'First Name', 'MI', 'LAST NAME',
      'MUN./CITY', 'PROVINCE', 'CONTACT NO. OF CUSTOMER', 'DATE PURCHASED',
      'DEALER NAME', 'Customer ADDRESS', 'AGE', 'GENDER', 'Dealer Code', 'File/Image Link'
    ]);
    
    // Validate required fields
    if (!formData.wrcNo || !formData.engineNo || !formData.firstName || 
        !formData.lastName || !formData.contactNo || !formData.dealerCode) {
      return {
        success: false,
        message: 'Required fields are missing. Please fill all fields marked with *'
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
    // Order: A-O (WRC No., ENGINE No., First Name, MI, LAST NAME, 
    //        MUN./CITY, PROVINCE, CONTACT NO., DATE PURCHASED, 
    //        DEALER NAME, Customer ADDRESS, AGE, GENDER, Dealer Code, File Link)
    sheet.appendRow([
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
      formData.fileUrl || ''       // O: File/Image Link
    ]);
    
    return {
      success: true,
      message: 'WRC registration submitted successfully! WRC No: ' + formData.wrcNo
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
    const sheet = getSheet_(SHEET_FSC, [
      'Dealer Transmittal No.', 'Dealer/Mechanic Code', 'WRC Number',
      'Frame Number', 'Coupon Number', 'Actual Mileage',
      'Repaired Month', 'Repaired Day', 'Repaired Year', 'KSC Code', 'File/Image Link'
    ]);
    
    // Validate required fields
    if (!formData.dealerTransNo || !formData.dealerMechCode || !formData.wrcNumber || 
        !formData.frameNumber || !formData.actualMileage) {
      return {
        success: false,
        message: 'Required fields are missing. Please fill all fields marked with *'
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
    // Order: A-K (Dealer Transmittal No., Dealer/Mechanic Code, WRC Number,
    //        Frame Number, Coupon Number, Actual Mileage,
    //        Repaired Month, Repaired Day, Repaired Year, KSC Code, File Link)
    sheet.appendRow([
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
      formData.fileUrl || ''       // K: File/Image Link
    ]);
    
    return {
      success: true,
      message: 'FSC entry submitted successfully! WRC Number: ' + formData.wrcNumber
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
    const sheet = getSheet_(SHEET_WLP, [
      'WRS Number', 'Repair Acknowledged Month', 'Repair Acknowledged Day',
      'Repair Acknowledged Year', 'Acknowledged By: (Customer Name)', 'Dealer/Mechanic Code', 'File/Image Link'
    ]);
    
    // Validate required fields
    if (!formData.wrsNumber || !formData.acknowledgedBy || !formData.dealerMechCode) {
      return {
        success: false,
        message: 'Required fields are missing. Please fill all fields marked with *'
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
    // Order: A-G (WRS Number, Repair Acknowledged Month, Repair Acknowledged Day,
    //        Repair Acknowledged Year, Acknowledged By, Dealer/Mechanic Code, File Link)
    sheet.appendRow([
      formData.wrsNumber,        // A: WRS Number
      formData.repairAckMonth,   // B: Repair Acknowledged Month (MM)
      formData.repairAckDay,     // C: Repair Acknowledged Day (DD)
      formData.repairAckYear,    // D: Repair Acknowledged Year (YYYY)
      formData.acknowledgedBy,   // E: Acknowledged By (Customer Name)
      formData.dealerMechCode,   // F: Dealer/Mechanic Code
      formData.fileUrl || ''     // G: File/Image Link
    ]);
    
    return {
      success: true,
      message: 'WLP acknowledgment submitted successfully! WRS Number: ' + formData.wrsNumber
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
