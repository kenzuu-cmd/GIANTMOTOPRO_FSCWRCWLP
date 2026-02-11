/**
 * ============================================================
 *  FSC / WRC / WLP -- Google Apps Script Backend
 *  WITH DUAL EMAIL NOTIFICATIONS
 * ============================================================
 *
 *  Server-side functions for the FormApp.html web application.
 *  Handles form submissions, validates data, enforces business
 *  rules, and writes to Google Sheets.
 *
 *  EMAIL NOTIFICATION FEATURES:
 *  1. PENDING Confirmation  – Sent immediately upon form submission
 *  2. TRANSMITTED Notification – Sent when STATUS is changed to
 *     "TRANSMITTED TO KAWASAKI" (via installable onEdit trigger)
 *
 *  SETUP INSTRUCTIONS:
 *  1. Replace SPREADSHEET_ID with your Google Sheets ID
 *  2. Ensure your spreadsheet has sheets: WRC, FSC, WLP, Log
 *  3. Deploy as web app (Execute as: Me, Access: Anyone)
 *  4. Install onEdit trigger:
 *     – Go to Triggers (clock icon in left sidebar)
 *     – Click "+ Add Trigger"
 *     – Function: onEdit, Event source: From spreadsheet, Event: On edit
 *     – Save and authorize
 * ============================================================
 */

// ================================================================
//  CONSTANTS
// ================================================================

const SPREADSHEET_ID = '15rPY6eyA-mMhtAsoVag3rfbCtmTuNmbQ8DFbK6Cjaa0';

const SHEET_WRC = 'WRC';
const SHEET_FSC = 'FSC';
const SHEET_WLP = 'WLP';
const SHEET_LOG = 'Log';

/** Column indices for uniqueness checks (0-based, row A = 0) */
const WRC_COL_WRC_NO    = 0;  // Column A
const WRC_COL_ENGINE_NO = 1;  // Column B

/** Default Dealer Code when field is absent from UI (WLP) */
const DEFAULT_DEALER_CODE = '20551';

/**
 * Valid Dealer Codes (from Giant Moto XLSX):
 *  20052 – BACAYAN           20053 – BARILI
 *  20054 – CARMEN            20055 – CLARIN
 *  20056 – CORDOVA           20051 – LAPULAPU
 *  20057 – MINGLANILLA       20058 – SAN FERNANDO CEBU
 *  20551 – TALISAY           20059 – TOLEDO
 *  20060 – UBAY              20061 – V.RAMA
 */
const VALID_DEALER_CODES = [
  '20052','20053','20054','20055','20056','20051',
  '20057','20058','20551','20059','20060','20061'
];

// ================================================================
//  HELPER FUNCTIONS
// ================================================================

/**
 * Get or create a sheet by name.  If the sheet doesn't exist it is
 * created with the supplied header row (bold, light-grey background).
 */
function getSheet_(name, headers) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Check whether a value already exists in a column (for uniqueness).
 * @param {Sheet}  sheet    target sheet
 * @param {number} colIndex 0-based column index
 * @param {string} value    value to check
 * @return {boolean}
 */
function isDuplicate_(sheet, colIndex, value) {
  if (!value || String(value).trim() === '') return false;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var data = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  var needle = String(value).trim().toUpperCase();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === needle) return true;
  }
  return false;
}

/**
 * Return all header names from row 1 as trimmed strings.
 */
function getHeaders_(sheet) {
  if (!sheet || sheet.getLastColumn() === 0) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h).trim(); });
}

/**
 * Find a column's 1-based index by header name (case-insensitive).
 * Returns 0 if not found.
 */
function findColumnByHeader_(sheet, headerName) {
  var headers = getHeaders_(sheet);
  var target = String(headerName).toUpperCase();
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toUpperCase() === target) return i + 1;
  }
  return 0;
}

/**
 * Safely ensure a column exists in the sheet.
 * Creates at the END of the sheet if missing.
 * NEVER overwrites or shifts existing columns.
 */
function ensureColumnExists_(sheet, columnName, createIfMissing) {
  if (createIfMissing === undefined) createIfMissing = true;

  var existing = findColumnByHeader_(sheet, columnName);
  if (existing > 0) return existing;

  if (!createIfMissing) return 0;

  try {
    var headers = getHeaders_(sheet);
    var newCol = headers.length + 1;
    sheet.getRange(1, newCol).setValue(columnName);
    sheet.getRange(1, newCol).setFontWeight('bold');
    sheet.getRange(1, newCol).setBackground('#f3f3f3');
    logInfo_('Created column "' + columnName + '" at position ' + newCol, sheet.getName(), 1);
    return newCol;
  } catch (err) {
    var msg = 'Failed to create column "' + columnName + '": ' + err.message;
    logError_(msg, '', '');
    throw new Error(msg);
  }
}

/**
 * Ensure an attachment column exists, checking common name variations
 * so that legacy sheets with 'File/Image Link' etc. are handled.
 *
 * @param {Sheet}  sheet         target sheet
 * @param {string} preferredName name to create if no variation exists
 * @return {{ index: number, name: string }}  1-based index + actual header name
 */
function ensureAttachmentColumn_(sheet, preferredName) {
  preferredName = preferredName || 'Attachment URL';
  var variations = [
    'Attachment URL', 'File/Image Link', 'Image/Drive Link',
    'File Link', 'Drive Link', 'Attachment', 'File URL'
  ];
  for (var i = 0; i < variations.length; i++) {
    var idx = findColumnByHeader_(sheet, variations[i]);
    if (idx > 0) return { index: idx, name: variations[i] };
  }
  var newIdx = ensureColumnExists_(sheet, preferredName, true);
  return { index: newIdx, name: preferredName };
}

// ================================================================
//  LOGGING HELPERS
// ================================================================

function logError_(message, refId1, refId2) {
  try {
    var logSheet = getSheet_(SHEET_LOG, ['Timestamp','Level','Sheet','Row','Reference','Message']);
    logSheet.appendRow([new Date(), 'ERROR', '', '', refId1 || refId2 || '', message]);
  } catch (e) {
    Logger.log('ERROR: ' + message);
  }
}

function logInfo_(message, sheetName, rowNumber) {
  try {
    var logSheet = getSheet_(SHEET_LOG, ['Timestamp','Level','Sheet','Row','Reference','Message']);
    logSheet.appendRow([new Date(), 'INFO', sheetName || '', rowNumber || '', '', message]);
  } catch (e) {
    Logger.log('INFO: ' + message);
  }
}

// ================================================================
//  ROW MAPPING (HEADER-BASED — never hard-coded column indices)
// ================================================================

/**
 * Build a row array where each element corresponds to its header.
 * @param {Sheet}  sheet    the target sheet
 * @param {Object} fieldMap { headerName: value, … }
 */
function buildRowFromHeaders_(sheet, fieldMap) {
  var headers = getHeaders_(sheet);
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    var headerUpper = headers[i].toUpperCase();
    var matched = false;
    for (var key in fieldMap) {
      if (key.toUpperCase() === headerUpper) {
        row.push(fieldMap[key] !== undefined && fieldMap[key] !== null ? fieldMap[key] : '');
        matched = true;
        break;
      }
    }
    if (!matched) row.push('');
  }
  return row;
}

/**
 * Safely append a row with header-based mapping.
 * Ensures all required columns exist before writing.
 */
function appendMappedRow_(sheet, fieldMap, requiredColumns) {
  if (requiredColumns && requiredColumns.length > 0) {
    for (var i = 0; i < requiredColumns.length; i++) {
      ensureColumnExists_(sheet, requiredColumns[i], true);
    }
  }
  var rowData = buildRowFromHeaders_(sheet, fieldMap);
  sheet.appendRow(rowData);
  return sheet.getLastRow();
}

// ================================================================
//  DATE NORMALIZATION: MMDDYYYY (no slashes)
// ================================================================

/**
 * Normalize any reasonable date string to MMDDYYYY (8 digits, no delimiters).
 * Handles: YYYY-MM-DD, YYYYMMDD, MM/DD/YYYY, MMDDYYYY.
 * Defence-in-depth: always call server-side before writing.
 */
function normalizeDateMMDDYYYY_(dateStr) {
  if (!dateStr) return '';
  var s = String(dateStr).replace(/[\/\-]/g, '');
  if (s.length === 8) {
    var prefix = s.substring(0, 2);
    if (prefix === '19' || prefix === '20') {
      // YYYYMMDD → MMDDYYYY
      return s.substring(4, 6) + s.substring(6, 8) + s.substring(0, 4);
    }
    return s; // Already MMDDYYYY
  }
  return s; // Return as-is; validation will catch bad formats
}

// ================================================================
//  WEB APP ENTRY POINTS
// ================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('FormApp')
    .setTitle('FSC WRC WLP Form Entry')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function showFormApp() {
  var html = HtmlService.createHtmlOutputFromFile('FormApp')
    .setTitle('FSC WRC WLP Forms');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showFormDialog() {
  var html = HtmlService.createHtmlOutputFromFile('FormApp')
    .setWidth(850).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'FSC WRC WLP Form Entry');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('FSC WRC WLP')
    .addItem('Open Form (Sidebar)', 'showFormApp')
    .addItem('Open Form (Dialog)', 'showFormDialog')
    .addToUi();
}

// ================================================================
//  FORM SUBMISSION: WRC
// ================================================================

function submitWRC(formData) {
  try {
    var sheet = getSheet_(SHEET_WRC, [
      'WRC No.', 'ENGINE No.', 'First Name', 'MI', 'LAST NAME',
      'MUN./CITY', 'PROVINCE', 'CONTACT NO. OF CUSTOMER', 'DATE PURCHASED',
      'DEALER NAME', 'Customer ADDRESS', 'AGE', 'GENDER', 'Dealer Code',
      'Attachment URL', 'BRANCH', 'EMAIL', 'STATUS'
    ]);

    // ---- Validation ----
    var missing = [];
    if (!formData.wrcNo       || !String(formData.wrcNo).trim())       missing.push('WRC Number');
    if (!formData.engineNo    || !String(formData.engineNo).trim())    missing.push('Engine Number');
    if (!formData.firstName   || !String(formData.firstName).trim())   missing.push('First Name');
    if (!formData.lastName    || !String(formData.lastName).trim())    missing.push('Last Name');
    if (!formData.contactNo   || !String(formData.contactNo).trim())   missing.push('Contact Number');
    if (!formData.email       || !String(formData.email).trim())       missing.push('Email Address');
    if (!formData.branch      || !String(formData.branch).trim())      missing.push('Branch');
    if (!formData.datePurchased || !String(formData.datePurchased).trim()) missing.push('Date Purchased');
    if (!formData.dealerCode  || !String(formData.dealerCode).trim())  missing.push('Dealer Code');
    if (!formData.fileUrl     || !String(formData.fileUrl).trim())     missing.push('File/Document Upload');

    if (missing.length > 0) {
      return { success: false, message: 'Required fields missing: ' + missing.join(', ') };
    }

    // Email format
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(formData.email)) {
      return { success: false, message: 'Invalid email address format.' };
    }

    // Dealer code
    if (VALID_DEALER_CODES.indexOf(String(formData.dealerCode).trim()) === -1) {
      return { success: false, message: 'Invalid Dealer Code. Please select a valid code.' };
    }

    // Uniqueness
    if (isDuplicate_(sheet, WRC_COL_WRC_NO, formData.wrcNo)) {
      logError_('Duplicate WRC No: ' + formData.wrcNo, formData.wrcNo, formData.engineNo);
      return { success: false, message: 'WRC No. "' + formData.wrcNo + '" already exists.' };
    }
    if (isDuplicate_(sheet, WRC_COL_ENGINE_NO, formData.engineNo)) {
      logError_('Duplicate ENGINE No: ' + formData.engineNo, formData.wrcNo, formData.engineNo);
      return { success: false, message: 'ENGINE No. "' + formData.engineNo + '" already exists.' };
    }

    // Date → MMDDYYYY (defence-in-depth; client already converts)
    var normalizedDate = normalizeDateMMDDYYYY_(formData.datePurchased);
    if (!/^\d{8}$/.test(normalizedDate)) {
      return { success: false, message: 'Invalid date format. Expected 8-digit MMDDYYYY.' };
    }

    // ---- Resolve attachment column (handles legacy 'File/Image Link' sheets) ----
    var attachInfo = ensureAttachmentColumn_(sheet, 'Attachment URL');

    // ---- Build field map ----
    var fieldMap = {
      'WRC No.':                   formData.wrcNo,
      'ENGINE No.':                formData.engineNo,
      'First Name':                formData.firstName,
      'MI':                        formData.mi || '',
      'LAST NAME':                 formData.lastName,
      'MUN./CITY':                 formData.munCity || '',
      'PROVINCE':                  formData.province || '',
      'CONTACT NO. OF CUSTOMER':   formData.contactNo,
      'DATE PURCHASED':            normalizedDate,
      'DEALER NAME':               formData.dealerName || '',
      'Customer ADDRESS':          formData.customerAddress || '',
      'AGE':                       formData.age || '',
      'GENDER':                    formData.gender || '',
      'Dealer Code':               String(formData.dealerCode).trim(),
      'BRANCH':                    formData.branch,
      'EMAIL':                     formData.email,
      'STATUS':                    'PENDING'
    };
    // Use the actual attachment header name (may be legacy variant)
    fieldMap[attachInfo.name] = formData.fileUrl || '';

    // Required columns that MUST exist (created at end of sheet if missing)
    var requiredCols = [
      'WRC No.', 'ENGINE No.', 'First Name', 'LAST NAME',
      'CONTACT NO. OF CUSTOMER', 'DATE PURCHASED', 'Dealer Code',
      'BRANCH', 'EMAIL', 'STATUS'
    ];

    try {
      var lastRow = appendMappedRow_(sheet, fieldMap, requiredCols);
      logInfo_('WRC submitted – Row ' + lastRow, SHEET_WRC, lastRow);
    } catch (err) {
      logError_('WRC append failed: ' + err.message, formData.wrcNo, formData.engineNo);
      return { success: false, message: 'System error: Could not save data. Contact administrator.' };
    }

    // Send confirmation email
    sendPendingConfirmationEmail(
      formData.email, 'WRC', formData.wrcNo, formData.engineNo,
      formData.branch, formData.dealerName
    );

    return {
      success: true,
      message: 'WRC registration submitted successfully. WRC No: ' + formData.wrcNo
             + '. Confirmation email sent to ' + formData.email + '.'
    };

  } catch (err) {
    logError_('WRC exception: ' + err.message, formData.wrcNo || '', formData.engineNo || '');
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// ================================================================
//  FORM SUBMISSION: FSC
// ================================================================

function submitFSC(formData) {
  try {
    var sheet = getSheet_(SHEET_FSC, [
      'Dealer Transmittal No.', 'Dealer/Mechanic Code', 'WRC Number',
      'Frame Number', 'Coupon Number', 'Actual Mileage',
      'Repaired Month', 'Repaired Day', 'Repaired Year', 'KSC Code',
      'Dealer Code', 'Attachment URL', 'BRANCH', 'EMAIL', 'STATUS'
    ]);

    // ---- Validation ----
    var missing = [];
    if (!formData.dealerTransNo  || !String(formData.dealerTransNo).trim())  missing.push('Dealer Transmittal Number');
    if (!formData.dealerMechCode || !String(formData.dealerMechCode).trim()) missing.push('Dealer/Mechanic Code');
    if (!formData.wrcNumber      || !String(formData.wrcNumber).trim())      missing.push('WRC Number');
    if (!formData.frameNumber    || !String(formData.frameNumber).trim())    missing.push('Frame Number');
    if (!formData.actualMileage  || !String(formData.actualMileage).trim())  missing.push('Actual Mileage');
    if (!formData.repairedMonth  || !String(formData.repairedMonth).trim())  missing.push('Repaired Date (Month)');
    if (!formData.repairedDay    || !String(formData.repairedDay).trim())    missing.push('Repaired Date (Day)');
    if (!formData.repairedYear   || !String(formData.repairedYear).trim())   missing.push('Repaired Date (Year)');
    if (!formData.email          || !String(formData.email).trim())          missing.push('Email Address');
    if (!formData.branch         || !String(formData.branch).trim())         missing.push('Branch');
    if (!formData.fileUrl        || !String(formData.fileUrl).trim())        missing.push('File/Document Upload');

    if (missing.length > 0) {
      return { success: false, message: 'Required fields missing: ' + missing.join(', ') };
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(formData.email)) {
      return { success: false, message: 'Invalid email address format.' };
    }

    var dealerCode = String(formData.dealerMechCode).trim();
    if (VALID_DEALER_CODES.indexOf(dealerCode) === -1) {
      return { success: false, message: 'Invalid Dealer/Mechanic Code.' };
    }

    if (formData.actualMileage && isNaN(Number(formData.actualMileage))) {
      return { success: false, message: 'Actual mileage must be a number.' };
    }

    if (formData.couponNumber && ['1','2','3','4'].indexOf(String(formData.couponNumber)) === -1) {
      return { success: false, message: 'Coupon number must be 1, 2, 3, or 4.' };
    }

    // ---- Resolve attachment column ----
    var attachInfo = ensureAttachmentColumn_(sheet, 'Attachment URL');

    // ---- Build field map ----
    var fieldMap = {
      'Dealer Transmittal No.': formData.dealerTransNo,
      'Dealer/Mechanic Code':   dealerCode,
      'WRC Number':             formData.wrcNumber,
      'Frame Number':           formData.frameNumber,
      'Coupon Number':          formData.couponNumber || '',
      'Actual Mileage':         formData.actualMileage,
      'Repaired Month':         formData.repairedMonth,
      'Repaired Day':           formData.repairedDay,
      'Repaired Year':          formData.repairedYear,
      'KSC Code':               formData.kscCode || '',
      'Dealer Code':            dealerCode,
      'BRANCH':                 formData.branch,
      'EMAIL':                  formData.email,
      'STATUS':                 'PENDING'
    };
    fieldMap[attachInfo.name] = formData.fileUrl || '';

    var requiredCols = [
      'Dealer Transmittal No.', 'Dealer/Mechanic Code', 'WRC Number',
      'Frame Number', 'Actual Mileage', 'Repaired Month', 'Repaired Day',
      'Repaired Year', 'Dealer Code', 'BRANCH', 'EMAIL', 'STATUS'
    ];

    try {
      var lastRow = appendMappedRow_(sheet, fieldMap, requiredCols);
      logInfo_('FSC submitted – Row ' + lastRow, SHEET_FSC, lastRow);
    } catch (err) {
      logError_('FSC append failed: ' + err.message, formData.wrcNumber, '');
      return { success: false, message: 'System error: Could not save data. Contact administrator.' };
    }

    sendPendingConfirmationEmail(
      formData.email, 'FSC', formData.wrcNumber, formData.frameNumber,
      formData.branch, dealerCode
    );

    return {
      success: true,
      message: 'FSC entry submitted successfully. WRC Number: ' + formData.wrcNumber
             + '. Confirmation email sent to ' + formData.email + '.'
    };

  } catch (err) {
    logError_('FSC exception: ' + err.message, formData.wrcNumber || '', '');
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// ================================================================
//  FORM SUBMISSION: WLP
// ================================================================

function submitWLP(formData) {
  try {
    var sheet = getSheet_(SHEET_WLP, [
      'WRS Number', 'Repair Acknowledged Month', 'Repair Acknowledged Day',
      'Repair Acknowledged Year', 'Acknowledged By: (Customer Name)',
      'Dealer/Mechanic Code', 'Dealer Code',
      'Attachment URL', 'BRANCH', 'EMAIL', 'STATUS'
    ]);

    // ---- Validation ----
    var missing = [];
    if (!formData.wrsNumber      || !String(formData.wrsNumber).trim())      missing.push('WRS Number');
    if (!formData.repairAckMonth || !String(formData.repairAckMonth).trim()) missing.push('Repair Acknowledged Date (Month)');
    if (!formData.repairAckDay   || !String(formData.repairAckDay).trim())   missing.push('Repair Acknowledged Date (Day)');
    if (!formData.repairAckYear  || !String(formData.repairAckYear).trim())  missing.push('Repair Acknowledged Date (Year)');
    if (!formData.acknowledgedBy || !String(formData.acknowledgedBy).trim()) missing.push('Acknowledged By');
    if (!formData.email          || !String(formData.email).trim())          missing.push('Email Address');
    if (!formData.branch         || !String(formData.branch).trim())         missing.push('Branch');
    if (!formData.fileUrl        || !String(formData.fileUrl).trim())        missing.push('File/Document Upload');

    if (missing.length > 0) {
      return { success: false, message: 'Required fields missing: ' + missing.join(', ') };
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(formData.email)) {
      return { success: false, message: 'Invalid email address format.' };
    }

    // Default Dealer Code for WLP (no UI field)
    var dealerCode = formData.dealerMechCode
      ? String(formData.dealerMechCode).trim()
      : DEFAULT_DEALER_CODE;
    if (!dealerCode || VALID_DEALER_CODES.indexOf(dealerCode) === -1) {
      dealerCode = DEFAULT_DEALER_CODE;
    }

    // ---- Resolve attachment column ----
    var attachInfo = ensureAttachmentColumn_(sheet, 'Attachment URL');

    // ---- Build field map ----
    var fieldMap = {
      'WRS Number':                          formData.wrsNumber,
      'Repair Acknowledged Month':           formData.repairAckMonth,
      'Repair Acknowledged Day':             formData.repairAckDay,
      'Repair Acknowledged Year':            formData.repairAckYear,
      'Acknowledged By: (Customer Name)':    formData.acknowledgedBy,
      'Dealer/Mechanic Code':                dealerCode,
      'Dealer Code':                         dealerCode,
      'BRANCH':                              formData.branch,
      'EMAIL':                               formData.email,
      'STATUS':                              'PENDING'
    };
    fieldMap[attachInfo.name] = formData.fileUrl || '';

    var requiredCols = [
      'WRS Number', 'Repair Acknowledged Month', 'Repair Acknowledged Day',
      'Repair Acknowledged Year', 'Acknowledged By: (Customer Name)',
      'Dealer/Mechanic Code', 'Dealer Code', 'BRANCH', 'EMAIL', 'STATUS'
    ];

    try {
      var lastRow = appendMappedRow_(sheet, fieldMap, requiredCols);
      logInfo_('WLP submitted – Row ' + lastRow, SHEET_WLP, lastRow);
    } catch (err) {
      logError_('WLP append failed: ' + err.message, formData.wrsNumber, '');
      return { success: false, message: 'System error: Could not save data. Contact administrator.' };
    }

    sendPendingConfirmationEmail(
      formData.email, 'WLP', formData.wrsNumber, '',
      formData.branch, dealerCode
    );

    return {
      success: true,
      message: 'WLP acknowledgment submitted successfully. WRS Number: ' + formData.wrsNumber
             + '. Confirmation email sent to ' + formData.email + '.'
    };

  } catch (err) {
    logError_('WLP exception: ' + err.message, formData.wrsNumber || '', '');
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// ================================================================
//  DYNAMIC DROPDOWN DATA
// ================================================================

function getProvinces() {
  return [
    'Metro Manila', 'Cebu', 'Davao', 'Rizal', 'Bulacan', 'Cavite',
    'Laguna', 'Pampanga', 'Batangas', 'Quezon', 'Pangasinan', 'Iloilo',
    'Negros Occidental', 'Leyte', 'Misamis Oriental', 'Albay', 'Cagayan',
    'Isabela', 'Nueva Ecija', 'Tarlac', 'Zambales', 'Palawan', 'Bohol',
    'Samar', 'Cotabato'
  ];
}

function getDealerCodes() {
  return VALID_DEALER_CODES;
}

// ================================================================
//  FILE UPLOAD HANDLER
// ================================================================

function uploadFile(fileData) {
  try {
    var ROOT_FOLDER = 'FSC_WRC_WLP_Files';

    // Find or create root folder
    var rootFolder;
    var folders = DriveApp.getFoldersByName(ROOT_FOLDER);
    rootFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(ROOT_FOLDER);

    // Find or create form-type subfolder
    var subName = fileData.formType || 'OTHER';
    var formFolder;
    var subs = rootFolder.getFoldersByName(subName);
    formFolder = subs.hasNext() ? subs.next() : rootFolder.createFolder(subName);

    // Create file from base64
    var bytes = Utilities.base64Decode(fileData.data);
    var blob  = Utilities.newBlob(bytes, fileData.type, fileData.name);
    var file  = formFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      url: file.getUrl(),
      fileId: file.getId(),
      message: 'File uploaded successfully.'
    };
  } catch (err) {
    Logger.log('Upload error: ' + err.toString());
    return { success: false, url: '', message: 'Upload failed: ' + err.toString() };
  }
}

// ================================================================
//  EMAIL: PENDING CONFIRMATION
// ================================================================

/**
 * Send PENDING confirmation email immediately after form submission.
 */
function sendPendingConfirmationEmail(email, formType, primaryRef, secondaryRef, branch, dealerName) {
  try {
    if (!email || String(email).trim() === '') return;

    var subject = 'Application Received – PENDING';

    // Build reference rows for the email body
    var refRows = '';
    if (formType === 'WRC') {
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">WRC Number</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (primaryRef || 'N/A') + '</td></tr>';
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">Engine Number</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (secondaryRef || 'N/A') + '</td></tr>';
    } else if (formType === 'FSC') {
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">WRC Number</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (primaryRef || 'N/A') + '</td></tr>';
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">Frame Number</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (secondaryRef || 'N/A') + '</td></tr>';
    } else if (formType === 'WLP') {
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">WRS Number</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (primaryRef || 'N/A') + '</td></tr>';
    }
    if (branch) {
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">Branch</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + branch + '</td></tr>';
    }
    if (dealerName) {
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">Dealer</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + dealerName + '</td></tr>';
    }

    var htmlBody = '<!DOCTYPE html>'
      + '<html><head><meta charset="utf-8"></head>'
      + '<body style="font-family:Arial,sans-serif;line-height:1.6;color:#333;max-width:600px;margin:0 auto;padding:0;background:#f4f4f4">'
      + '<div style="background:#fff;margin:20px auto;border-radius:8px;overflow:hidden;box-shadow:0 2px 4px rgba(0,0,0,.1)">'
      // Logo
      + '<div style="text-align:center;padding:20px;background:#fff">'
      + '<img src="https://i.imgur.com/nGVvR3v.png" alt="Giant Moto Pro Logo" style="max-width:200px;height:auto" />'
      + '</div>'
      // Header
      + '<div style="background:#fff3cd;color:#856404;padding:25px;text-align:center;border-bottom:4px solid #ffc107">'
      + '<h2 style="margin:0;font-size:24px">Application Status: PENDING</h2>'
      + '</div>'
      // Content
      + '<div style="padding:30px">'
      + '<p>Dear Applicant,</p>'
      + '<p>Thank you for your submission to the <strong>FSC / WRC / WLP Management System</strong>.</p>'
      + '<table style="width:100%;background:#f8f9fa;border-radius:6px;margin:20px 0;border-collapse:collapse">'
      + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">Form Type</td>'
      + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef"><strong>' + formType + '</strong></td></tr>'
      + refRows
      + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;width:40%">Status</td>'
      + '<td style="color:#333;padding:12px 15px"><strong>PENDING</strong></td></tr>'
      + '</table>'
      + '<div style="background:#fff3cd;border-left:4px solid #ffc107;padding:15px;margin:20px 0;border-radius:4px">'
      + '<strong>What\'s Next?</strong>'
      + '<p style="margin:10px 0 0">Your application is currently being reviewed. You will receive another email once your application has been transmitted to Kawasaki.</p>'
      + '</div>'
      + '<p>Thank you for choosing Giant Moto Pro.</p>'
      + '<p>Best regards,<br><strong>Giant Moto Pro Team</strong></p>'
      + '</div>'
      // Footer
      + '<div style="background:#f8f9fa;padding:20px;text-align:center;font-size:.9em;color:#666;border-top:1px solid #ddd">'
      + '<p style="margin:0"><strong>This is an automated message.</strong> Please do not reply.</p>'
      + '<p style="margin:8px 0 0">Contact your branch office for questions.</p>'
      + '</div>'
      + '</div></body></html>';

    var plainBody = 'APPLICATION RECEIVED – PENDING\n\n'
      + 'Form Type: ' + formType + '\n'
      + 'Status: PENDING\n\n'
      + 'Your application is being reviewed. You will receive another email when it is transmitted to Kawasaki.\n\n'
      + 'Giant Moto Pro Team';

    try {
      GmailApp.sendEmail(email, subject, plainBody, {
        htmlBody: htmlBody,
        name: 'Giant Moto Pro – FSC/WRC/WLP System'
      });
    } catch (gmailErr) {
      MailApp.sendEmail(email, subject, plainBody, {
        htmlBody: htmlBody,
        name: 'Giant Moto Pro – FSC/WRC/WLP System'
      });
    }
  } catch (err) {
    Logger.log('PENDING email error: ' + err.toString());
    logError_('PENDING email failed: ' + err.message, primaryRef || '', secondaryRef || '');
  }
}

// ================================================================
//  EMAIL: TRANSMITTED TO KAWASAKI
// ================================================================

/**
 * Send status-update email when row is changed to TRANSMITTED TO KAWASAKI.
 */
function sendTransmittedEmail_(email, formType, primaryRef, secondaryRef, branch, dealerName) {
  try {
    var subject = 'Application Status: TRANSMITTED TO KAWASAKI';

    var refRows = '';
    if (formType === 'WRC') {
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">WRC Number</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (primaryRef || 'N/A') + '</td></tr>';
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">Engine Number</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (secondaryRef || 'N/A') + '</td></tr>';
    } else if (formType === 'FSC') {
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">WRC Number</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (primaryRef || 'N/A') + '</td></tr>';
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">Frame Number</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (secondaryRef || 'N/A') + '</td></tr>';
    } else if (formType === 'WLP') {
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">WRS Number</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (primaryRef || 'N/A') + '</td></tr>';
    }
    if (branch) {
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">Branch</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + branch + '</td></tr>';
    }
    if (dealerName) {
      refRows += '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">Dealer</td>'
               + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + dealerName + '</td></tr>';
    }

    var htmlBody = '<!DOCTYPE html>'
      + '<html><head><meta charset="utf-8"></head>'
      + '<body style="font-family:Arial,sans-serif;line-height:1.6;color:#333;max-width:600px;margin:0 auto;padding:0;background:#f4f4f4">'
      + '<div style="background:#fff;margin:20px auto;border-radius:8px;overflow:hidden;box-shadow:0 2px 4px rgba(0,0,0,.1)">'
      // Logo
      + '<div style="text-align:center;padding:20px;background:#fff">'
      + '<img src="https://i.imgur.com/nGVvR3v.png" alt="Giant Moto Pro Logo" style="max-width:200px;height:auto" />'
      + '</div>'
      // Header (green)
      + '<div style="background:#28a745;color:#fff;padding:25px;text-align:center">'
      + '<h2 style="margin:0;font-size:24px">TRANSMITTED TO KAWASAKI</h2>'
      + '</div>'
      // Content
      + '<div style="padding:30px">'
      + '<p>Dear Customer,</p>'
      + '<p>Your application has been successfully processed and <strong>transmitted to Kawasaki</strong>.</p>'
      + '<table style="width:100%;background:#f8f9fa;border-radius:6px;margin:20px 0;border-collapse:collapse">'
      + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">Form Type</td>'
      + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef"><strong>' + formType + '</strong></td></tr>'
      + refRows
      + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;width:40%">Status</td>'
      + '<td style="color:#333;padding:12px 15px"><strong style="color:#28a745">TRANSMITTED TO KAWASAKI</strong></td></tr>'
      + '</table>'
      + '<div style="background:#d4edda;border-left:4px solid #28a745;padding:15px;margin:20px 0;border-radius:4px;color:#155724">'
      + '<strong>What\'s Next?</strong>'
      + '<p style="margin:10px 0 0">Your application is now with Kawasaki for final processing. No further action is required at this time.</p>'
      + '</div>'
      + '<p>Thank you for choosing Giant Moto Pro.</p>'
      + '<p>Best regards,<br><strong>Giant Moto Pro Team</strong></p>'
      + '</div>'
      // Footer
      + '<div style="background:#f8f9fa;padding:20px;text-align:center;font-size:.9em;color:#666;border-top:1px solid #ddd">'
      + '<p style="margin:0"><strong>This is an automated notification.</strong> Please do not reply.</p>'
      + '<p style="margin:8px 0 0">Contact your branch office for questions.</p>'
      + '</div>'
      + '</div></body></html>';

    var plainBody = 'STATUS UPDATE – TRANSMITTED TO KAWASAKI\n\n'
      + 'Form Type: ' + formType + '\n'
      + 'Status: TRANSMITTED TO KAWASAKI\n\n'
      + 'Your application is now with Kawasaki for final processing.\n\n'
      + 'Giant Moto Pro Team';

    try {
      GmailApp.sendEmail(email, subject, plainBody, {
        htmlBody: htmlBody,
        name: 'Giant Moto Pro – FSC/WRC/WLP System'
      });
      return { success: true, method: 'GmailApp' };
    } catch (gmailErr) {
      try {
        MailApp.sendEmail(email, subject, plainBody, {
          htmlBody: htmlBody,
          name: 'Giant Moto Pro – FSC/WRC/WLP System'
        });
        return { success: true, method: 'MailApp' };
      } catch (mailErr) {
        return { success: false, error: 'GmailApp: ' + gmailErr.message + ' / MailApp: ' + mailErr.message };
      }
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ================================================================
//  INSTALLABLE onEdit TRIGGER
// ================================================================

/**
 * Fires when any cell is edited.  If the STATUS column on WRC, FSC,
 * or WLP is changed to "TRANSMITTED TO KAWASAKI" it sends a
 * notification email to the row's EMAIL address.
 *
 * IMPORTANT: This must be installed as an INSTALLABLE trigger,
 * not a simple onEdit, so it can send email and access services.
 */
function onEdit(e) {
  var TRIGGER_STATUS = 'TRANSMITTED TO KAWASAKI';

  try {
    if (!e || !e.range) {
      logToSheet_('Log', 'ERROR', 'onEdit', '', 'No event object');
      return;
    }

    var range     = e.range;
    var sheet     = range.getSheet();
    var sheetName = sheet.getName();
    var row       = range.getRow();
    var col       = range.getColumn();
    var newValue  = range.getValue();
    var oldValue  = e.oldValue || '';

    // Only act on WRC, FSC, or WLP sheets
    if (['WRC', 'FSC', 'WLP'].indexOf(sheetName) === -1) return;
    if (row === 1) return; // ignore header row

    // Must be the STATUS column
    var statusCol = findColumnByHeader_(sheet, 'STATUS');
    var emailCol  = findColumnByHeader_(sheet, 'EMAIL');
    var branchCol = findColumnByHeader_(sheet, 'BRANCH');

    if (statusCol === 0 || emailCol === 0) return;
    if (col !== statusCol) return;

    // Only fire for the specific status value
    var newNorm = String(newValue).trim().toUpperCase();
    if (newNorm !== TRIGGER_STATUS.toUpperCase()) return;

    var oldNorm = String(oldValue).trim().toUpperCase();
    if (oldNorm === TRIGGER_STATUS.toUpperCase()) return; // no change

    // Prevent double-send via EMAIL_SENT flag
    var emailSentCol = findColumnByHeader_(sheet, 'EMAIL_SENT');
    if (emailSentCol > 0) {
      var flag = sheet.getRange(row, emailSentCol).getValue();
      if (flag === 'YES' || flag === true) return;
    }

    // Gather data using header-based column lookups (never hardcoded)
    var email  = String(sheet.getRange(row, emailCol).getValue()).trim();
    var branch = branchCol > 0 ? String(sheet.getRange(row, branchCol).getValue()).trim() : '';

    var primaryRef   = '';
    var secondaryRef = '';
    var dealerName   = '';

    if (sheetName === 'WRC') {
      var wrcNoCol    = findColumnByHeader_(sheet, 'WRC No.');
      var engineNoCol = findColumnByHeader_(sheet, 'ENGINE No.');
      var dealerCol   = findColumnByHeader_(sheet, 'DEALER NAME');
      if (wrcNoCol > 0)    primaryRef   = String(sheet.getRange(row, wrcNoCol).getValue()).trim();
      if (engineNoCol > 0) secondaryRef = String(sheet.getRange(row, engineNoCol).getValue()).trim();
      if (dealerCol > 0)   dealerName   = String(sheet.getRange(row, dealerCol).getValue()).trim();
    } else if (sheetName === 'FSC') {
      var fscWrcCol   = findColumnByHeader_(sheet, 'WRC Number');
      var frameCol    = findColumnByHeader_(sheet, 'Frame Number');
      var mechCodeCol = findColumnByHeader_(sheet, 'Dealer/Mechanic Code');
      if (fscWrcCol > 0)   primaryRef   = String(sheet.getRange(row, fscWrcCol).getValue()).trim();
      if (frameCol > 0)    secondaryRef = String(sheet.getRange(row, frameCol).getValue()).trim();
      if (mechCodeCol > 0) dealerName   = String(sheet.getRange(row, mechCodeCol).getValue()).trim();
    } else if (sheetName === 'WLP') {
      var wrsCol       = findColumnByHeader_(sheet, 'WRS Number');
      var wlpDealerCol = findColumnByHeader_(sheet, 'Dealer/Mechanic Code');
      if (wrsCol > 0)       primaryRef = String(sheet.getRange(row, wrsCol).getValue()).trim();
      if (wlpDealerCol > 0) dealerName = String(sheet.getRange(row, wlpDealerCol).getValue()).trim();
    }

    // Validate email
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      logToSheet_('Log', 'ERROR', sheetName, row, 'Invalid or empty email: ' + email);
      return;
    }

    // Send email
    var result = sendTransmittedEmail_(email, sheetName, primaryRef, secondaryRef, branch, dealerName);

    if (result.success) {
      logToSheet_('Log', 'SUCCESS', sheetName, row,
        'Transmitted email sent to ' + email + ' via ' + result.method);
      if (emailSentCol > 0) sheet.getRange(row, emailSentCol).setValue('YES');
    } else {
      logToSheet_('Log', 'ERROR', sheetName, row, 'Email failed: ' + result.error);
    }

  } catch (err) {
    logToSheet_('Log', 'CRITICAL', 'onEdit', '', 'Exception: ' + err.message);
  }
}

// ================================================================
//  ENHANCED LOG SHEET WRITER
// ================================================================

function logToSheet_(logSheetName, level, sheetName, row, message) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName(logSheetName);
    if (!logSheet) {
      logSheet = ss.insertSheet(logSheetName);
      logSheet.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Level', 'Sheet', 'Row', 'Message', 'User']]);
      logSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f0f0f0');
      logSheet.setFrozenRows(1);
    }

    var user = '';
    try { user = Session.getActiveUser().getEmail(); } catch (ignored) {}

    logSheet.appendRow([new Date(), level, sheetName || '', row || '', message, user]);

    // Color-code the level cell
    var lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

    var colors = {
      'SUCCESS':  '#d4edda',
      'ERROR':    '#f8d7da',
      'CRITICAL': '#ff0000',
      'SKIPPED':  '#fff3cd',
      'INFO':     '#d1ecf1'
    };
    if (colors[level]) {
      logSheet.getRange(lastRow, 2).setBackground(colors[level]);
    }
  } catch (err) {
    Logger.log('logToSheet_ error: ' + err.message + ' | ' + level + ' – ' + message);
  }
}

// ================================================================
//  TESTING & DIAGNOSTICS
// ================================================================

/**
 * TEST: Send a "Transmitted" email using row 2 of WRC sheet.
 * Run manually from the Script Editor.
 */
function TEST_EmailTrigger() {
  var TEST_CONFIG = {
    sheetName: 'WRC',
    testRow: 2,
    testEmail: 'kenji.devcodes@gmail.com'  // override for testing
  };

  Logger.log('=== EMAIL TRIGGER TEST – START ===');

  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TEST_CONFIG.sheetName);
    if (!sheet) { Logger.log('Sheet "' + TEST_CONFIG.sheetName + '" not found.'); return; }

    var statusCol = findColumnByHeader_(sheet, 'STATUS');
    var emailCol  = findColumnByHeader_(sheet, 'EMAIL');
    if (statusCol === 0 || emailCol === 0) {
      Logger.log('STATUS or EMAIL column not found.');
      return;
    }

    var originalEmail = sheet.getRange(TEST_CONFIG.testRow, emailCol).getValue();
    var targetEmail   = TEST_CONFIG.testEmail || originalEmail;
    if (!targetEmail) { Logger.log('No target email.'); return; }

    // Temporarily set test email
    if (TEST_CONFIG.testEmail && TEST_CONFIG.testEmail !== originalEmail) {
      sheet.getRange(TEST_CONFIG.testRow, emailCol).setValue(TEST_CONFIG.testEmail);
    }

    var wrcNoCol    = findColumnByHeader_(sheet, 'WRC No.');
    var engineNoCol = findColumnByHeader_(sheet, 'ENGINE No.');
    var wrcNum    = wrcNoCol > 0 ? sheet.getRange(TEST_CONFIG.testRow, wrcNoCol).getValue() : 'TEST-WRC';
    var engineNum = engineNoCol > 0 ? sheet.getRange(TEST_CONFIG.testRow, engineNoCol).getValue() : 'TEST-ENGINE';

    var result = sendTransmittedEmail_(targetEmail, TEST_CONFIG.sheetName, wrcNum, engineNum, 'Test Branch', 'Test Dealer');
    Logger.log(result.success ? 'SUCCESS via ' + result.method : 'FAILED: ' + result.error);

    // Restore original email
    if (TEST_CONFIG.testEmail && TEST_CONFIG.testEmail !== originalEmail) {
      sheet.getRange(TEST_CONFIG.testRow, emailCol).setValue(originalEmail);
    }

    Logger.log('=== EMAIL TRIGGER TEST – END ===');
  } catch (err) {
    Logger.log('TEST EXCEPTION: ' + err.message);
  }
}

/**
 * DIAGNOSTIC: Check installed triggers and email quota.
 */
function DIAGNOSTIC_CheckTriggerSetup() {
  Logger.log('=== TRIGGER DIAGNOSTIC ===');

  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('Total triggers installed: ' + triggers.length);

  var hasOnEdit = false;
  triggers.forEach(function(trigger, idx) {
    var fn   = trigger.getHandlerFunction();
    var type = trigger.getEventType();
    Logger.log('  [' + (idx + 1) + '] Function: ' + fn + '  |  Event: ' + type);
    if (fn === 'onEdit' && type === ScriptApp.EventType.ON_EDIT) hasOnEdit = true;
  });

  if (hasOnEdit) {
    Logger.log('RESULT: Installable onEdit trigger found.');
  } else {
    Logger.log('WARNING: No installable onEdit trigger detected.');
    Logger.log('  -> Go to Triggers > + Add Trigger > onEdit > From spreadsheet > On edit');
  }

  try {
    Logger.log('Remaining daily email quota: ' + MailApp.getRemainingDailyQuota());
  } catch (e) {
    Logger.log('Could not check email quota: ' + e.message);
  }
}

/**
 * DIAGNOSTIC: View recent log entries (last 20).
 */
function DIAGNOSTIC_ViewRecentLogs() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName('Log');
    if (!logSheet) { Logger.log('Log sheet not found.'); return; }

    var lastRow = logSheet.getLastRow();
    if (lastRow < 2) { Logger.log('Log sheet is empty.'); return; }

    var startRow = Math.max(2, lastRow - 19);
    var data = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 6).getValues();

    Logger.log('=== RECENT LOG ENTRIES (' + data.length + ') ===');
    data.forEach(function(row, i) {
      Logger.log('[Row ' + (startRow + i) + '] '
        + row[0] + ' | ' + row[1] + ' | ' + row[2]
        + ' | Row ' + row[3] + ' | ' + row[4]);
    });
  } catch (err) {
    Logger.log('DIAGNOSTIC error: ' + err.message);
  }
}

/**
 * UTILITY: Clean up old log entries (keep last 100).
 */
function UTILITY_CleanupOldLogs() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName('Log');
    if (!logSheet) { Logger.log('No Log sheet to clean.'); return; }

    var lastRow = logSheet.getLastRow();
    var KEEP = 100;
    if (lastRow <= KEEP + 1) {
      Logger.log('Log has ' + (lastRow - 1) + ' entries (<= ' + KEEP + '). Nothing to clean.');
      return;
    }

    var deleteCount = lastRow - KEEP - 1;
    logSheet.deleteRows(2, deleteCount);
    Logger.log('Cleaned up ' + deleteCount + ' old log entries. Remaining: ' + KEEP);
  } catch (err) {
    Logger.log('Cleanup error: ' + err.message);
  }
}

