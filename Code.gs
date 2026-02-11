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

/**
 * REQUIRED: WCF Kawasaki Logo (Drive File ID)
 * This logo appears:
 * 1. At the top of WCF PDFs
 * 2. In user thank-you emails (as inline image)
 * 3. In internal branch confirmation emails
 * 
 * To set up:
 * 1. Use this file ID: 1ib5qfMi4bnYj4sfJVAAr-wFfvdFgS60X
 * 2. Ensure the file has "Anyone with link can view" permissions
 */
const WCF_HEADER_IMAGE_ID = '1ib5qfMi4bnYj4sfJVAAr-wFfvdFgS60X';

/**
 * WCF FORM Template Cell Mapping (EXACT PLACEMENT per user requirements)
 * ALL mappings verified against template structure with merged-parent detection
 * CRITICAL: Write ONLY to fillable blank cells, NEVER to labels/headers
 */
const WCF_TEMPLATE_CELLS = {
  // ── Section 1: Main identity fields (Rows 7–15) ───────────
  //  Left-side text inputs: Columns G:P within rows 7–15
  dealerName: 'G7',         // Row 7, col G - Dealer's Name
  dealerAddress: 'G8',      // Row 8, col G - Dealer's Address
  customerName: 'G9',       // Row 9, col G - Customer's Name
  customerAddress: 'G10',   // Row 10, col G - Customer's Address
  modelColor: 'G11',        // Row 11, col G - Model / Color
  frameNo: 'G12',           // Row 12, col G - Frame No.
  engineNo: 'G13',          // Row 13, col G - Engine No.
  // Row 14 is MC Usage (checkboxes, handled separately)
  kilometerRun: 'G15',      // Row 15, col G - Kilometer Run

  //  Prepared date and Job Order No. (top-right header area)
  preparedDate: 'Q9',       // Prepared date blank area Q:AA rows 9-10, write to Q9
  jobOrderNo: 'W6',         // Job Order No. blank area W:AA rows 6-15, write to W6

  //  Right-side detail fields (WRC section, rows 11–15)
  wrcNo: 'Q11',             // Row 11 right box - WRC No.
  purchaseDate: 'Q12',      // Row 12 right box - Purchase Date
  failureDate: 'Q13',       // Row 13 right box - Failure Date
  reportedDate: 'Q14',      // Row 14 right box - Reported Date
  repairDate: 'Q15',        // Row 15 right box - Repair Date

  //  MC Usage checkboxes (Row 14) - cells for checkmark placement
  mcUsageSolo: 'G14',         // SOLO checkbox cell
  mcUsageTricycle: 'K14',     // Tricycle checkbox cell
  mcUsageOthers: 'O14',       // Others checkbox cell
  mcUsageOthersText: 'P14',   // Others text input area

  // ── Section 2: Narrative boxes (Problem sections) ─────────
  //  Template has labels in rows 17, 20, 23, 26
  //  Data areas are in rows 17-18, 20-21, 23-24, 26-27
  //  Write to top-left of each blank data box (NOT label row)
  problemComplaint: 'A17',      // Rows 17-18 data box (A17 is top-left of merged data)
  probableCause: 'A20',         // Rows 20-21 data box (A20 is top-left)
  correctiveAction: 'A23',      // Rows 23-24 data box (A23 is top-left)
  suggestionRemarks: 'A26',     // Rows 26-27 data box (A26 is top-left)

  // ── Section 3: Parts recommended table ───────────────────
  //  Causal Part / Main Defective Part (Row 32)
  causalPartNo: 'A32',          // Part No. (columns A-E region)
  causalPartName: 'C32',        // Part Name (within A-E)
  causalPartQty: 'G32',         // QTY column
  //  Affected Parts (Rows 34–40) - blank-row detection
  affectedPartsStart: 34,       // First affected-parts data row
  affectedPartsEnd: 40,         // Last affected-parts data row
  affectedPartNoCol: 1,         // Column A - Part No.
  affectedPartNameCol: 3,       // Column C - Part Name
  affectedPartQtyCol: 7,        // Column G - QTY
  //  Additional columns L, M, N, O, P (if used by template)
  affectedPartColL: 12,         // Column L
  affectedPartColM: 13,         // Column M
  affectedPartColN: 14,         // Column N
  affectedPartColO: 15,         // Column O
  affectedPartColP: 16,         // Column P

  // ── Section 4: Illustration image (Q:AA box) ──────────────
  illustrationCol: 17,          // Column Q (17)
  illustrationRow: 29,          // Anchor at Q29 (top-left of Q:AA box)

  // ── Section 5: Names + signatures (rows 44–48) ────────────
  //  Three blocks: A:I (Prepared by), J:R (Acknowledged by), S:AA (Warranty Repair)
  preparedByName: 'A44',        // Printed name in A:I block
  acknowledgedByName: 'J44',    // Printed name in J:R block
  warrantyRepairName: 'S44',    // Printed name in S:AA block
  //  Signature IMAGE positions (column, row for insertImage)
  preparedBySigCol: 1,   preparedBySigRow: 46,      // A:I signature area
  acknowledgedBySigCol: 10, acknowledgedBySigRow: 46, // J:R signature area
  warrantyRepairSigCol: 19, warrantyRepairSigRow: 46, // S:AA signature area

  // ── Section 6: Claim judgment checkboxes (rows 51–56) ────
  //  Claim Judgment: Accepted / Rejected / Returned
  claimAccepted: 'B51',         // Accepted checkbox cell
  claimRejected: 'F51',         // Rejected checkbox cell
  claimReturned: 'J51',         // Returned checkbox cell
  //  Reasons of rejection (multiple checkboxes rows 52-56)
  rejectReason1: 'B52',         // First rejection reason checkbox
  rejectReason2: 'B53',
  rejectReason3: 'B54',
  rejectReason4: 'B55',
  rejectReason5: 'B56',

  // ── Other fields ──────────────────────────────────────────
  unitsSameProblem: 'A38',      // Units with same problem
  deliverTo: 'B40'              // Deliver To field
};

// ─────────────────────────────────────────────────────────
//  WCF Helper Functions: Merged-parent detection + write guards
// ─────────────────────────────────────────────────────────

/**
 * Get the top-left A1 cell of a merged range (or return original if not merged).
 * This ensures we always write to the parent cell of merged areas.
 * @param {Sheet}  sheet - The sheet to check
 * @param {string} a1    - The A1 notation cell (e.g., 'G7')
 * @return {string}      - Top-left A1 of merged range, or original a1 if not merged
 */
function getMergedTopLeftA1_(sheet, a1) {
  var range = sheet.getRange(a1);
  if (!range.isPartOfMerge()) {
    return a1;
  }
  var mergedRanges = sheet.getRange(a1).getMergedRanges();
  if (mergedRanges.length === 0) {
    return a1;
  }
  var mergedRange = mergedRanges[0];
  var topLeftRow = mergedRange.getRow();
  var topLeftCol = mergedRange.getColumn();
  return sheet.getRange(topLeftRow, topLeftCol).getA1Notation();
}

/**
 * Write a value to a cell with safety checks:
 *  1. Resolves merged-parent top-left cell
 *  2. Checks if target cell is fillable (not a static template label/header)
 *  3. Logs the write operation with field name → resolved A1 = value
 *  4. Throws error if attempting to overwrite non-fillable template content
 *
 * BLANK-ONLY WRITE GUARD: Before writing, checks if the template cell contains
 * static text. If it does AND it's not a known fillable area, throws error.
 *
 * @param {Sheet}  sheet     - The temp print sheet (NOT original template)
 * @param {string} a1        - Target A1 cell (will be resolved to merged parent)
 * @param {*}      value     - Value to write
 * @param {string} fieldName - Field name for logging/error reporting
 */
function writeIfAllowed_(sheet, a1, value, fieldName) {
  // Skip empty values
  if (value === undefined || value === null || String(value).trim() === '') {
    return;
  }

  // 1. Resolve merged parent top-left
  var targetA1 = getMergedTopLeftA1_(sheet, a1);
  var targetRange = sheet.getRange(targetA1);

  // Debug logging for merged parent resolution (especially for jobOrderNo)
  if (a1 !== targetA1) {
    Logger.log('[MERGED PARENT] Field "' + fieldName + '": ' + a1 + ' → ' + targetA1 + ' (resolved to merged top-left)');
  }

  // 2. Check if cell currently contains static template content
  //    (This check assumes the temp sheet is a fresh copy of the template)
  var existingValue = targetRange.getValue();
  var existingDisplay = targetRange.getDisplayValue();

  // Define fillable zones (rows/columns where blank input cells exist)
  var fillableZones = [
    // Job Order No. area (row 6, columns W:AA) - top-right header box
    { rowStart: 6, rowEnd: 15, colStart: 23, colEnd: 27 },  // W6:AA15 (23=W, 27=AA)
    // Main identity fields (rows 7-15, columns G:P and Q:AA)
    { rowStart: 7, rowEnd: 15, colStart: 7, colEnd: 27 },   // G to AA
    // Narrative boxes (rows 17-27, columns A:AA)
    { rowStart: 17, rowEnd: 27, colStart: 1, colEnd: 27 },  // A to AA
    // Parts table (rows 32-40, columns A:P)
    { rowStart: 32, rowEnd: 40, colStart: 1, colEnd: 16 },  // A to P
    // Names/signatures (rows 44-48, columns A:AA)
    { rowStart: 44, rowEnd: 48, colStart: 1, colEnd: 27 },  // A to AA
    // Claim judgment (rows 51-56, columns A:AA)
    { rowStart: 51, rowEnd: 56, colStart: 1, colEnd: 27 },  // A to AA
    // Other fields
    { rowStart: 38, rowEnd: 40, colStart: 1, colEnd: 27 }   // Units same problem, Deliver To
  ];

  // Check if target is within a fillable zone
  var targetRow = targetRange.getRow();
  var targetCol = targetRange.getColumn();
  var isInFillableZone = false;
  for (var i = 0; i < fillableZones.length; i++) {
    var zone = fillableZones[i];
    if (targetRow >= zone.rowStart && targetRow <= zone.rowEnd &&
        targetCol >= zone.colStart && targetCol <= zone.colEnd) {
      isInFillableZone = true;
      break;
    }
  }

  // If outside fillable zones, throw error
  if (!isInFillableZone) {
    throw new Error('[WRITE GUARD] Field "' + fieldName + '" attempted to write outside fillable zones: ' + 
                    targetA1 + ' (row=' + targetRow + ', col=' + targetCol + ')');
  }

  // If existing value is non-empty and looks like a template label (not a fillable blank),
  // we MIGHT be overwriting template content. However, since we're working on a TEMP COPY
  // and the template should have fillable cells blank, we allow the write but log it.
  if (existingDisplay && existingDisplay.trim() !== '') {
    Logger.log('[WRITE GUARD] WARNING: Overwriting non-empty cell ' + targetA1 + 
               ' (current="' + existingDisplay + '") with ' + fieldName + '="' + String(value).substring(0, 60) + '"');
  }

  // 3. Write the value
  targetRange.setValue(value);

  // 4. Log the write
  Logger.log('[WCF POPULATE] ' + fieldName + ' → ' + targetA1 + ' = ' + String(value).substring(0, 80));
}

/**
 * NOTE: With the temp-copy approach, we never clear the original template.
 * The temp sheet is deleted after PDF export. This list is kept for reference
 * only (e.g., if you ever revert to in-place editing).
 */
var WCF_TEMPLATE_CLEAR_RANGES = [
  'T6', 'T7',                    // Header dates
  'B8', 'O8', 'B9',              // Dealer info
  'B10', 'O10', 'B11',           // Customer info
  'B12', 'O12', 'B13', 'O13',    // Unit info part 1
  'B14', 'O14', 'B15', 'O15',    // Unit info part 2
  'B16', 'O16',                  // Kilometer/Repair date
  'A18', 'A21', 'A24', 'A27',    // Problem section DATA cells
  'A31', 'C31', 'G31',           // Causal parts
  'A33:G36',                     // Affected parts rows
  'A38',                         // Units same problem
  'B40',                         // Deliver to
  'A45', 'C45', 'F45'            // Signature names
];

const SHEET_WRC = 'WRC';
const SHEET_FSC = 'FSC';
const SHEET_WLP = 'WLP';
const SHEET_WCF = 'WCF';
const SHEET_WCF_FORM = 'WCF FORM';  // Template sheet for PDF export (with merged cells)
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

function doGet(e) {
  // WCF HTML preview endpoint (for layout debugging)
  if (e && e.parameter && e.parameter.wcfPreview === '1') {
    return serveWcfHtmlPreview_(e.parameter);
  }
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
//  WCF: HEADERS & ID GENERATION
// ================================================================

var WCF_HEADERS = [
  'WCF ID', 'Branch', 'Prepared Date', 'Job Order No.',
  'Dealer Name', 'Dealer Address', 'Dealer Phone',
  'Customer Name', 'Customer Address', 'Customer Phone',
  'Model', 'Color', 'Frame No.', 'Engine No.', 'WRC No.',
  'Purchase Date', 'Failure Date', 'Reported Date', 'Repair Date',
  'MC Usage', 'MC Usage Other', 'Kilometer Run',
  'Problem Complaint', 'Probable Cause', 'Corrective Action', 'Suggestion Remarks',
  'Causal Part No', 'Causal Part Name', 'Causal Part Qty', 'Part Supply Method',
  'Affected Parts', 'Units Same Problem',
  'Deliver To',
  'Prepared By Name', 'Prepared By Designation', 'Prepared By Signature URL',
  'Acknowledged By Name', 'Acknowledged By Designation', 'Acknowledged By Signature URL',
  'Warranty Repair Name', 'Warranty Repair Designation', 'Warranty Repair Signature URL',
  'Illustration URL', 'PDF URL', 'PDF File ID',
  'EMAIL', 'STATUS',
  'Created At', 'Created By', 'Email Sent At', 'Email Recipients', 'Email Status',
  'USER EMAIL', 'USER EMAIL STATUS', 'USER EMAIL ERROR', 'USER EMAIL SENT AT'
];

/**
 * Get WCF email recipients: branch email + NolanRalph + all GMPC branches
 * @param {string} branchValue - selected branch name
 * @return {string[]} array of unique email addresses
 */
function getWcfRecipients_(branchValue) {
  var recipients = [];
  
  // Always include NolanRalph
  recipients.push('NolanRalph_P@kmp.com.ph');
  
  // Always include all GMPC branch emails
  recipients.push('giantmotoprobarilibranch@gmail.com');
  recipients.push('giantmotoproinventory@gmail.com');
  recipients.push('giantmotoproopon19@gmail.com');
  recipients.push('giantmotopro_corporation@yahoo.com.ph');
  recipients.push('gmpcbacayan@gmail.com');
  recipients.push('gmpcbalamban@gmail.com');
  recipients.push('gmpccordova@gmail.com');
  recipients.push('gmpcpinamungajan22@gmail.com');
  recipients.push('gmpcsanfernando@gmail.com');
  recipients.push('gmpcsierrabullones@gmail.com');
  recipients.push('gmpctabunok@gmail.com');
  recipients.push('gmpctalibon@gmail.com');
  recipients.push('gmpctayudconsolacion@gmail.com');
  recipients.push('gmpctoledo21@gmail.com');
  recipients.push('gmpcubay@gmail.com');
  
  // De-duplicate and return
  var unique = [];
  for (var i = 0; i < recipients.length; i++) {
    if (unique.indexOf(recipients[i]) === -1) {
      unique.push(recipients[i]);
    }
  }
  return unique;
}

/** Generate unique WCF ID: WCF-YYYYMMDD-0001 */
function generateWcfId_() {
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyyMMdd');
  var sheet;
  try { sheet = getSheet_(SHEET_WCF, WCF_HEADERS); } catch (e) {
    return 'WCF-' + dateStr + '-0001';
  }
  var lastRow = sheet.getLastRow();
  var seq = 1;
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var prefix = 'WCF-' + dateStr + '-';
    for (var i = 0; i < ids.length; i++) {
      var id = String(ids[i][0]);
      if (id.indexOf(prefix) === 0) {
        var n = parseInt(id.substring(prefix.length), 10);
        if (n >= seq) seq = n + 1;
      }
    }
  }
  return 'WCF-' + dateStr + '-' + ('0000' + seq).slice(-4);
}

// ================================================================
//  WCF: FORM SUBMISSION
// ================================================================

function submitWCF(formData) {
  try {
    var sheet = getSheet_(SHEET_WCF, WCF_HEADERS);

    // ---- Validation ----
    var missing = [];
    if (!formData.branch          || !String(formData.branch).trim())          missing.push('Branch');
    if (!formData.preparedDate    || !String(formData.preparedDate).trim())    missing.push('Prepared Date');
    if (!formData.dealerName      || !String(formData.dealerName).trim())      missing.push('Dealer Name');
    if (!formData.customerName    || !String(formData.customerName).trim())    missing.push('Customer Name');
    if (!formData.model           || !String(formData.model).trim())           missing.push('Model');
    if (!formData.frameNo         || !String(formData.frameNo).trim())         missing.push('Frame No.');
    if (!formData.engineNo        || !String(formData.engineNo).trim())        missing.push('Engine No.');
    if (!formData.wrcNo           || !String(formData.wrcNo).trim())           missing.push('WRC No.');
    if (!formData.purchaseDate    || !String(formData.purchaseDate).trim())    missing.push('Purchase Date');
    if (!formData.failureDate     || !String(formData.failureDate).trim())     missing.push('Failure Date');
    if (!formData.problemComplaint|| !String(formData.problemComplaint).trim())missing.push('Problem/Complaint');
    if (!formData.email           || !String(formData.email).trim())           missing.push('Email');
    if (!formData.illustrationUrl || !String(formData.illustrationUrl).trim()) missing.push('Illustration Image');

    if (missing.length > 0) {
      return { success: false, message: 'Required fields missing: ' + missing.join(', ') };
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(formData.email)) {
      return { success: false, message: 'Invalid email address format.' };
    }

    // Generate ID
    var wcfId = generateWcfId_();

    // Normalize dates
    var prepDate   = normalizeDateMMDDYYYY_(formData.preparedDate);
    var purchDate  = normalizeDateMMDDYYYY_(formData.purchaseDate);
    var failDate   = normalizeDateMMDDYYYY_(formData.failureDate);
    var repDate    = formData.reportedDate ? normalizeDateMMDDYYYY_(formData.reportedDate) : '';
    var repairDate = formData.repairDate   ? normalizeDateMMDDYYYY_(formData.repairDate)   : '';

    // Save signatures to Drive (if provided)
    var preparedSigUrl = '';
    var acknowledgedSigUrl = '';
    var warrantySigUrl = '';
    
    try {
      if (formData.preparedBySignature) {
        preparedSigUrl = saveSignatureToDrive_(formData.preparedBySignature, wcfId, 'PreparedBy');
      }
      if (formData.acknowledgedBySignature) {
        acknowledgedSigUrl = saveSignatureToDrive_(formData.acknowledgedBySignature, wcfId, 'AcknowledgedBy');
      }
      if (formData.warrantyRepairSignature) {
        warrantySigUrl = saveSignatureToDrive_(formData.warrantyRepairSignature, wcfId, 'WarrantyRepair');
      }
    } catch (sigErr) {
      logError_('WCF signature save failed: ' + sigErr.message, wcfId, '');
    }

    // Current timestamp for audit
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd HH:mm:ss');
    var createdBy = Session.getActiveUser().getEmail() || 'System';

    // Build field map
    var fieldMap = {
      'WCF ID':                       wcfId,
      'Branch':                       formData.branch,
      'Prepared Date':                prepDate,
      'Job Order No.':                formData.jobOrderNo || '',
      'Dealer Name':                  formData.dealerName,
      'Dealer Address':               formData.dealerAddress || '',
      'Dealer Phone':                 formData.dealerPhone || '',
      'Customer Name':                formData.customerName,
      'Customer Address':             formData.customerAddress || '',
      'Customer Phone':               formData.customerPhone || '',
      'Model':                        formData.model,
      'Color':                        formData.color || '',
      'Frame No.':                    formData.frameNo,
      'Engine No.':                   formData.engineNo,
      'WRC No.':                      formData.wrcNo,
      'Purchase Date':                purchDate,
      'Failure Date':                 failDate,
      'Reported Date':                repDate,
      'Repair Date':                  repairDate,
      'MC Usage':                     formData.mcUsage || '',
      'MC Usage Other':               formData.mcUsageOther || '',
      'Kilometer Run':                formData.kilometerRun || '',
      'Problem Complaint':            formData.problemComplaint,
      'Probable Cause':               formData.probableCause || '',
      'Corrective Action':            formData.correctiveAction || '',
      'Suggestion Remarks':           formData.suggestionRemarks || '',
      'Causal Part No':               formData.causalPartNo || '',
      'Causal Part Name':             formData.causalPartName || '',
      'Causal Part Qty':              formData.causalPartQty || '',
      'Part Supply Method':           formData.partSupplyMethod || '',
      'Affected Parts':               formData.affectedParts || '',
      'Units Same Problem':           formData.unitsSameProblem || '',
      'Deliver To':                   formData.deliverTo || '',
      'Prepared By Name':             formData.preparedByName || '',
      'Prepared By Designation':      formData.preparedByDesignation || '',
      'Prepared By Signature URL':    preparedSigUrl,
      'Acknowledged By Name':         formData.acknowledgedByName || '',
      'Acknowledged By Designation':  formData.acknowledgedByDesignation || '',
      'Acknowledged By Signature URL': acknowledgedSigUrl,
      'Warranty Repair Name':         formData.warrantyRepairName || '',
      'Warranty Repair Designation':  formData.warrantyRepairDesignation || '',
      'Warranty Repair Signature URL': warrantySigUrl,
      'Illustration URL':             formData.illustrationUrl || '',
      'PDF URL':                      '',
      'PDF File ID':                  '',
      'EMAIL':                        formData.email,
      'STATUS':                       '',  // WCF: leave STATUS blank (not using status workflow)
      'Created At':                   timestamp,
      'Created By':                   createdBy,
      'Email Sent At':                '',
      'Email Recipients':             '',
      'Email Status':                 '',
      'USER EMAIL':                   (formData.email || '').trim(),
      'USER EMAIL STATUS':            '',
      'USER EMAIL ERROR':             '',
      'USER EMAIL SENT AT':           ''
    };

    // Save to sheet - capture exact row number for atomic PDF URL write-back
    var rowNumber;
    try {
      rowNumber = appendMappedRow_(sheet, fieldMap, WCF_HEADERS);
      logInfo_('WCF submitted – Row ' + rowNumber + ' – ' + wcfId, SHEET_WCF, rowNumber);
    } catch (err) {
      logError_('WCF append failed: ' + err.message, wcfId, '');
      return { success: false, message: 'System error saving data. Contact administrator.' };
    }

    // Generate PDF using WCF FORM template
    var pdfResult = { success: false };
    var pdfUrl = '';
    var pdfFileId = '';
    var userEmailStatus = 'not sent';
    
    try {
      var pdfData = {};
      for (var k in formData) pdfData[k] = formData[k];
      pdfData.wcfId       = wcfId;
      pdfData.preparedDate = prepDate;
      pdfData.purchaseDate = purchDate;
      pdfData.failureDate  = failDate;
      pdfData.reportedDate = repDate;
      pdfData.repairDate   = repairDate;
      // Add signature URLs for template population
      pdfData.preparedBySignatureUrl = preparedSigUrl;
      pdfData.acknowledgedBySignatureUrl = acknowledgedSigUrl;
      pdfData.warrantyRepairSignatureUrl = warrantySigUrl;

      pdfResult = generateWcfPdf_(pdfData);

      if (pdfResult.success && pdfResult.url && pdfResult.id) {
        pdfUrl = pdfResult.url;
        pdfFileId = pdfResult.id;
        
        // Atomic write-back: detect columns dynamically and update the exact row
        var pdfUrlCol = findColumnByHeader_(sheet, 'PDF URL');
        var pdfIdCol = findColumnByHeader_(sheet, 'PDF File ID');
        
        if (pdfUrlCol > 0) {
          sheet.getRange(rowNumber, pdfUrlCol).setValue(pdfUrl);
          logInfo_('PDF URL written to row ' + rowNumber + ', col ' + pdfUrlCol, wcfId, '');
        }
        if (pdfIdCol > 0) {
          sheet.getRange(rowNumber, pdfIdCol).setValue(pdfFileId);
          logInfo_('PDF File ID written to row ' + rowNumber + ', col ' + pdfIdCol, wcfId, '');
        }
      } else {
        logError_('PDF generation failed or incomplete', wcfId, '');
      }
    } catch (pdfErr) {
      logError_('WCF PDF exception: ' + pdfErr.message, wcfId, '');
    }

    // ---- USER THANK-YOU EMAIL (AUTOMATIC, GUARANTEED, OBSERVABLE) ----
    var userEmail = (formData.email || '').trim();
    var userEmailError = '';
    var userEmailSentAt = '';

    if (!pdfResult.success) {
      // PDF failed → cannot send email (no PDF to link)
      userEmailStatus = 'SKIPPED';
      userEmailError = 'PDF generation failed; no PDF to attach';
    } else if (!userEmail) {
      userEmailStatus = 'SKIPPED';
      userEmailError = 'Email field is empty';
    } else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(userEmail)) {
      userEmailStatus = 'SKIPPED';
      userEmailError = 'Invalid email format: ' + userEmail;
    } else {
      // Attempt send
      try {
        sendWcfUserThankYouEmail_({
          userEmail: userEmail,
          wcfId: wcfId,
          customerName: formData.customerName,
          branch: formData.branch,
          jobOrderNo: formData.jobOrderNo || '',
          pdfFileId: pdfFileId,
          pdfUrl: pdfUrl
        });
        userEmailStatus = 'SENT';
        userEmailSentAt = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd HH:mm:ss');
        logInfo_('User thank-you email SENT to: ' + userEmail, wcfId, '');
      } catch (emailErr) {
        userEmailStatus = 'FAILED';
        userEmailError = String(emailErr.message || emailErr);
        logError_('User thank-you email FAILED: ' + userEmailError, wcfId, '');
      }
    }

    // Write-back user email status columns to the SAME row (create cols if missing)
    try {
      var ueCol      = ensureColumnExists_(sheet, 'USER EMAIL', true);
      var ueStatCol  = ensureColumnExists_(sheet, 'USER EMAIL STATUS', true);
      var ueErrCol   = ensureColumnExists_(sheet, 'USER EMAIL ERROR', true);
      var ueSentCol  = ensureColumnExists_(sheet, 'USER EMAIL SENT AT', true);

      if (ueCol > 0)     sheet.getRange(rowNumber, ueCol).setValue(userEmail);
      if (ueStatCol > 0) sheet.getRange(rowNumber, ueStatCol).setValue(userEmailStatus);
      if (ueErrCol > 0)  sheet.getRange(rowNumber, ueErrCol).setValue(userEmailError);
      if (ueSentCol > 0) sheet.getRange(rowNumber, ueSentCol).setValue(userEmailSentAt);
      logInfo_('User email status written to row ' + rowNumber + ': ' + userEmailStatus, wcfId, '');
    } catch (wbErr) {
      logError_('User email status write-back failed: ' + wbErr.message, wcfId, '');
    }

    // Do NOT send internal branch email automatically
    // Internal email will be sent when user clicks "Send to Email" button in post-submit modal

    return {
      success: true,
      message: 'WCF submitted successfully. ID: ' + wcfId,
      wcfId: wcfId,
      rowNumber: rowNumber,
      pdfUrl: pdfUrl || null,
      pdfFileId: pdfFileId || null,
      previewUrl: pdfFileId ? 'https://drive.google.com/file/d/' + pdfFileId + '/preview' : null,
      downloadUrl: pdfFileId ? 'https://drive.google.com/uc?export=download&id=' + pdfFileId : null,
      branch: formData.branch,
      userEmailStatus: userEmailStatus,
      userEmailError: userEmailError,
      userEmailSentAt: userEmailSentAt
    };

  } catch (err) {
    logError_('WCF exception: ' + err.message, '', '');
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// ================================================================
//  WCF: SIGNATURE MANAGEMENT
// ================================================================

/**
 * Save a signature image (base64 PNG) to Drive
 * @param {string} base64Data - base64-encoded PNG data (with or without data:image/png;base64, prefix)
 * @param {string} wcfId - WCF ID for folder organization
 * @param {string} signatureType - e.g., 'PreparedBy', 'AcknowledgedBy', 'WarrantyRepair'
 * @return {string} Drive URL of saved signature
 */
function saveSignatureToDrive_(base64Data, wcfId, signatureType) {
  if (!base64Data) return '';
  
  // Remove data:image/png;base64, prefix if present
  var cleanData = base64Data.replace(/^data:image\/png;base64,/, '');
  
  // Decode base64 to blob
  var blob = Utilities.newBlob(Utilities.base64Decode(cleanData), 'image/png', signatureType + '_' + wcfId + '.png');
  
  // Get or create folder structure: FSC_WRC_WLP_Files/WCF/YYYY/MM/WCF-ID/
  var rootFolder;
  var rf = DriveApp.getFoldersByName('FSC_WRC_WLP_Files');
  rootFolder = rf.hasNext() ? rf.next() : DriveApp.createFolder('FSC_WRC_WLP_Files');
  
  var wcfFolder;
  var wf = rootFolder.getFoldersByName('WCF');
  wcfFolder = wf.hasNext() ? wf.next() : rootFolder.createFolder('WCF');
  
  // Create year/month subfolders
  var now = new Date();
  var year = Utilities.formatDate(now, 'Asia/Manila', 'yyyy');
  var month = Utilities.formatDate(now, 'Asia/Manila', 'MM');
  
  var yearFolder;
  var yf = wcfFolder.getFoldersByName(year);
  yearFolder = yf.hasNext() ? yf.next() : wcfFolder.createFolder(year);
  
  var monthFolder;
  var mf = yearFolder.getFoldersByName(month);
  monthFolder = mf.hasNext() ? mf.next() : yearFolder.createFolder(month);
  
  // Create WCF-specific folder
  var wcfIdFolder;
  var widf = monthFolder.getFoldersByName(wcfId);
  wcfIdFolder = widf.hasNext() ? widf.next() : monthFolder.createFolder(wcfId);
  
  // Save file
  var file = wcfIdFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}

// ================================================================
//  WCF: PDF GENERATION — HTML/CSS Renderer (replaces Sheets-based)
//  Generates pixel-perfect legal PDF from WcfPrint.html template.
//  The old Sheet-based generator is preserved as generateWcfPdfViaSheet_.
// ================================================================

// ─────────────────────────────────────────────────────────
//  HTML PDF Helpers
// ─────────────────────────────────────────────────────────

/**
 * Include an HTML file's content (for sub-templates).
 * @param {string} filename - HTML filename without extension
 * @return {string}
 */
function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Fetch the Kawasaki logo from Drive and return as base64 data URI.
 * Uses WCF_HEADER_IMAGE_ID constant.
 * @return {string} data URI or empty string on failure
 */
function getWcfLogoBase64_() {
  try {
    var file = DriveApp.getFileById(WCF_HEADER_IMAGE_ID);
    var blob = file.getBlob();
    var b64 = Utilities.base64Encode(blob.getBytes());
    return 'data:' + blob.getContentType() + ';base64,' + b64;
  } catch (e) {
    Logger.log('[WCF HTML PDF] Logo fetch failed: ' + e.message);
    return '';
  }
}

/**
 * Convert a Drive file URL to a base64 data URI.
 * @param {string} url - Google Drive file URL
 * @return {string} data URI or empty string
 */
function driveUrlToBase64_(url) {
  if (!url) return '';
  try {
    var fileId = extractDriveFileId_(url);
    if (!fileId) return '';
    var blob = DriveApp.getFileById(fileId).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    Logger.log('[WCF HTML PDF] Image fetch failed (' + url + '): ' + e.message);
    return '';
  }
}

/**
 * Escape HTML special characters for safe template rendering.
 * @param {string} text
 * @return {string}
 */
function escapeHtml_(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Format MMDDYYYY (8-digit) string to MM/DD/YYYY for PDF display.
 * Passes through other formats unchanged.
 * @param {string} mmddyyyy
 * @return {string}
 */
function formatDateForPdf_(mmddyyyy) {
  if (!mmddyyyy) return '';
  var s = String(mmddyyyy).trim();
  if (/^\d{8}$/.test(s)) {
    return s.substring(0, 2) + '/' + s.substring(2, 4) + '/' + s.substring(4, 8);
  }
  return s;
}

/**
 * Get or create WCF Drive folder: FSC_WRC_WLP_Files/WCF/YYYY/MM/wcfId/
 * @param {string} wcfId
 * @return {Folder}
 */
function getOrCreateWcfFolder_(wcfId) {
  var rf = DriveApp.getFoldersByName('FSC_WRC_WLP_Files');
  var rootFolder = rf.hasNext() ? rf.next() : DriveApp.createFolder('FSC_WRC_WLP_Files');

  var wf = rootFolder.getFoldersByName('WCF');
  var wcfFolder = wf.hasNext() ? wf.next() : rootFolder.createFolder('WCF');

  var now = new Date();
  var year = Utilities.formatDate(now, 'Asia/Manila', 'yyyy');
  var month = Utilities.formatDate(now, 'Asia/Manila', 'MM');

  var yf = wcfFolder.getFoldersByName(year);
  var yearFolder = yf.hasNext() ? yf.next() : wcfFolder.createFolder(year);

  var mf = yearFolder.getFoldersByName(month);
  var monthFolder = mf.hasNext() ? mf.next() : yearFolder.createFolder(month);

  var widf = monthFolder.getFoldersByName(wcfId);
  return widf.hasNext() ? widf.next() : monthFolder.createFolder(wcfId);
}

// ─────────────────────────────────────────────────────────
//  generateWcfPdfFromHtml_  (primary WCF PDF generator)
// ─────────────────────────────────────────────────────────

/**
 * Generate WCF PDF from HTML/CSS template (WcfPrint.html).
 * Replaces the fragile Sheets-based approach.
 *
 * Steps:
 *  1. Fetch logo + images as base64
 *  2. Parse affected parts JSON
 *  3. Build HtmlService template with all data
 *  4. Evaluate to HTML string
 *  5. Convert HTML → PDF blob via Utilities
 *  6. Save to Drive in standard folder structure
 *  7. Return {success, url, id}
 *
 * @param {string} wcfId  - WCF document ID
 * @param {object} d      - Full form data payload
 * @return {object} {success, url, id}
 */
function generateWcfPdfFromHtml_(wcfId, d) {
  Logger.log('[WCF HTML PDF] ── START ── ' + wcfId);

  try {
    // 1. Fetch all images as base64 data URIs
    var logoBase64 = getWcfLogoBase64_();
    var illustrationBase64 = driveUrlToBase64_(d.illustrationUrl);
    var prepSigBase64 = driveUrlToBase64_(d.preparedBySignatureUrl);
    var ackSigBase64 = driveUrlToBase64_(d.acknowledgedBySignatureUrl);
    var warSigBase64 = driveUrlToBase64_(d.warrantyRepairSignatureUrl);

    // 2. Parse affected parts
    var affectedParts = [];
    try {
      if (d.affectedParts) affectedParts = JSON.parse(d.affectedParts);
    } catch (e) {
      Logger.log('[WCF HTML PDF] Affected parts parse error: ' + e.message);
    }

    // 3. Parse rejection reasons
    var rejectReasons = [];
    try {
      if (d.rejectReasons) {
        rejectReasons = Array.isArray(d.rejectReasons) ? d.rejectReasons : JSON.parse(d.rejectReasons);
      }
    } catch (e) {}

    // 4. Build template
    var tpl = HtmlService.createTemplateFromFile('WcfPrint');
    tpl.wcfId = wcfId;
    tpl.logoBase64 = logoBase64;
    tpl.illustrationBase64 = illustrationBase64;
    tpl.prepSigBase64 = prepSigBase64;
    tpl.ackSigBase64 = ackSigBase64;
    tpl.warSigBase64 = warSigBase64;
    tpl.affectedParts = affectedParts;
    tpl.data = d;

    // Template helper: format date
    tpl.formatDate = formatDateForPdf_;

    // Template helper: MC Usage checkbox
    tpl.mcUsageChecked = function(type) {
      var usage = String(d.mcUsage || '').toUpperCase();
      if (type === 'OTHERS') {
        return (usage !== '' && usage.indexOf('SOLO') < 0 && usage.indexOf('TRICYCLE') < 0) ? '✓' : '';
      }
      return usage.indexOf(type) >= 0 ? '✓' : '';
    };

    // Template helper: claim judgment checkbox
    tpl.claimChecked = function(type) {
      return String(d.claimJudgment || '').toUpperCase().indexOf(type) >= 0 ? '✓' : '';
    };

    // Template helper: rejection reason checkbox
    tpl.rejectReasonChecked = function(idx) {
      for (var i = 0; i < rejectReasons.length; i++) {
        if (String(rejectReasons[i]) === String(idx)) return '✓';
      }
      return '';
    };

    // 5. Evaluate template → HTML string
    var htmlContent = tpl.evaluate().getContent();
    Logger.log('[WCF HTML PDF] HTML rendered, length=' + htmlContent.length);

    // 6. Convert HTML → PDF blob
    var pdfBlob = Utilities.newBlob(htmlContent, 'text/html', 'WCF_' + wcfId + '.html')
      .getAs('application/pdf');
    pdfBlob.setName('WCF_' + wcfId + '.pdf');
    Logger.log('[WCF HTML PDF] PDF blob created: ' + pdfBlob.getName());

    // 7. Save to Drive
    var folder = getOrCreateWcfFolder_(wcfId);
    var pdfFile = folder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log('[WCF HTML PDF] ── DONE ── saved: ' + pdfFile.getUrl());
    return {
      success: true,
      url: pdfFile.getUrl(),
      id: pdfFile.getId()
    };

  } catch (err) {
    Logger.log('[WCF HTML PDF] ERROR: ' + err.message);
    Logger.log('[WCF HTML PDF] Stack: ' + (err.stack || ''));
    throw err;
  }
}

// ─────────────────────────────────────────────────────────
//  WCF HTML Preview Endpoint (for layout debugging)
// ─────────────────────────────────────────────────────────

/**
 * Serve WCF HTML preview in browser for layout inspection.
 * Usage: deploy URL + ?wcfPreview=1&wcfId=WCF-20260211-0001
 * If no wcfId provided, renders with sample/placeholder data.
 *
 * @param {object} params - URL parameters
 * @return {HtmlOutput}
 */
function serveWcfHtmlPreview_(params) {
  var wcfId = params.wcfId || 'WCF-PREVIEW-0000';
  var d = {};

  // Try to load real data from sheet if wcfId provided
  if (params.wcfId) {
    try {
      var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_WCF);
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        var headers = data[0];
        for (var i = 1; i < data.length; i++) {
          if (data[i][0] === params.wcfId) {
            for (var j = 0; j < headers.length; j++) {
              d[String(headers[j]).replace(/[^a-zA-Z0-9]/g, '')] = data[i][j];
            }
            // Map to expected field names
            d.dealerName = data[i][headers.indexOf('Dealer Name')] || '';
            d.dealerAddress = data[i][headers.indexOf('Dealer Address')] || '';
            d.dealerPhone = data[i][headers.indexOf('Dealer Phone')] || '';
            d.customerName = data[i][headers.indexOf('Customer Name')] || '';
            d.customerAddress = data[i][headers.indexOf('Customer Address')] || '';
            d.customerPhone = data[i][headers.indexOf('Customer Phone')] || '';
            d.model = data[i][headers.indexOf('Model')] || '';
            d.color = data[i][headers.indexOf('Color')] || '';
            d.frameNo = data[i][headers.indexOf('Frame No.')] || '';
            d.engineNo = data[i][headers.indexOf('Engine No.')] || '';
            d.wrcNo = data[i][headers.indexOf('WRC No.')] || '';
            d.branch = data[i][headers.indexOf('Branch')] || '';
            d.preparedDate = data[i][headers.indexOf('Prepared Date')] || '';
            d.jobOrderNo = data[i][headers.indexOf('Job Order No.')] || '';
            d.purchaseDate = data[i][headers.indexOf('Purchase Date')] || '';
            d.failureDate = data[i][headers.indexOf('Failure Date')] || '';
            d.reportedDate = data[i][headers.indexOf('Reported Date')] || '';
            d.repairDate = data[i][headers.indexOf('Repair Date')] || '';
            d.mcUsage = data[i][headers.indexOf('MC Usage')] || '';
            d.mcUsageOther = data[i][headers.indexOf('MC Usage Other')] || '';
            d.kilometerRun = data[i][headers.indexOf('Kilometer Run')] || '';
            d.problemComplaint = data[i][headers.indexOf('Problem Complaint')] || '';
            d.probableCause = data[i][headers.indexOf('Probable Cause')] || '';
            d.correctiveAction = data[i][headers.indexOf('Corrective Action')] || '';
            d.suggestionRemarks = data[i][headers.indexOf('Suggestion Remarks')] || '';
            d.causalPartNo = data[i][headers.indexOf('Causal Part No')] || '';
            d.causalPartName = data[i][headers.indexOf('Causal Part Name')] || '';
            d.causalPartQty = data[i][headers.indexOf('Causal Part Qty')] || '';
            d.partSupplyMethod = data[i][headers.indexOf('Part Supply Method')] || '';
            d.affectedParts = data[i][headers.indexOf('Affected Parts')] || '';
            d.unitsSameProblem = data[i][headers.indexOf('Units Same Problem')] || '';
            d.deliverTo = data[i][headers.indexOf('Deliver To')] || '';
            d.preparedByName = data[i][headers.indexOf('Prepared By Name')] || '';
            d.acknowledgedByName = data[i][headers.indexOf('Acknowledged By Name')] || '';
            d.warrantyRepairName = data[i][headers.indexOf('Warranty Repair Name')] || '';
            d.illustrationUrl = data[i][headers.indexOf('Illustration URL')] || '';
            d.preparedBySignatureUrl = data[i][headers.indexOf('Prepared By Signature URL')] || '';
            d.acknowledgedBySignatureUrl = data[i][headers.indexOf('Acknowledged By Signature URL')] || '';
            d.warrantyRepairSignatureUrl = data[i][headers.indexOf('Warranty Repair Signature URL')] || '';
            break;
          }
        }
      }
    } catch (e) {
      Logger.log('[WCF Preview] Data load error: ' + e.message);
    }
  }

  // If no real data, use sample placeholders
  if (!d.dealerName) {
    d = {
      branch: 'TALISAY', dealerName: 'Sample Dealer Inc.', dealerAddress: '123 Main St, Cebu City',
      dealerPhone: '032-123-4567', customerName: 'Juan Dela Cruz', customerAddress: '456 Sample Ave, Talisay City',
      customerPhone: '0917-123-4567', model: 'BARAKO II', color: 'BLACK', frameNo: 'FRAME-SAMPLE-001',
      engineNo: 'ENGINE-SAMPLE-001', wrcNo: 'WRC-SAMPLE', preparedDate: '02112026',
      jobOrderNo: 'JO-SAMPLE-001', purchaseDate: '01152026', failureDate: '02012026',
      reportedDate: '02052026', repairDate: '02102026', mcUsage: 'Solo', mcUsageOther: '',
      kilometerRun: '5000', problemComplaint: 'Engine overheating during prolonged use. Customer reports vibration at idle.',
      probableCause: 'Thermostat malfunction. Coolant level below minimum.', correctiveAction: 'Replaced thermostat. Topped up coolant. Adjusted idle speed.',
      suggestionRemarks: 'Recommend follow-up inspection after 500km.', causalPartNo: '49054-0034',
      causalPartName: 'THERMOSTAT', causalPartQty: '1', partSupplyMethod: 'FOC',
      affectedParts: JSON.stringify([{partNo:'92055-0188', partName:'RING,O', qty:'2'},{partNo:'49016-0014', partName:'VALVE-ASSY', qty:'1'}]),
      unitsSameProblem: '3', deliverTo: 'KMPC HO', preparedByName: 'JOHN DOE',
      acknowledgedByName: 'JANE SMITH', warrantyRepairName: 'MIKE SANTOS',
      illustrationUrl: '', preparedBySignatureUrl: '', acknowledgedBySignatureUrl: '', warrantyRepairSignatureUrl: ''
    };
  }

  // Build template (same logic as generateWcfPdfFromHtml_ but return HTML output)
  var logoBase64 = getWcfLogoBase64_();
  var illustrationBase64 = driveUrlToBase64_(d.illustrationUrl);
  var prepSigBase64 = driveUrlToBase64_(d.preparedBySignatureUrl);
  var ackSigBase64 = driveUrlToBase64_(d.acknowledgedBySignatureUrl);
  var warSigBase64 = driveUrlToBase64_(d.warrantyRepairSignatureUrl);

  var affectedParts = [];
  try { if (d.affectedParts) affectedParts = JSON.parse(d.affectedParts); } catch (e) {}

  var rejectReasons = [];
  try { if (d.rejectReasons) rejectReasons = Array.isArray(d.rejectReasons) ? d.rejectReasons : JSON.parse(d.rejectReasons); } catch (e) {}

  var tpl = HtmlService.createTemplateFromFile('WcfPrint');
  tpl.wcfId = wcfId;
  tpl.logoBase64 = logoBase64;
  tpl.illustrationBase64 = illustrationBase64;
  tpl.prepSigBase64 = prepSigBase64;
  tpl.ackSigBase64 = ackSigBase64;
  tpl.warSigBase64 = warSigBase64;
  tpl.affectedParts = affectedParts;
  tpl.data = d;
  tpl.formatDate = formatDateForPdf_;
  tpl.mcUsageChecked = function(type) {
    var usage = String(d.mcUsage || '').toUpperCase();
    if (type === 'OTHERS') return (usage !== '' && usage.indexOf('SOLO') < 0 && usage.indexOf('TRICYCLE') < 0) ? '✓' : '';
    return usage.indexOf(type) >= 0 ? '✓' : '';
  };
  tpl.claimChecked = function(type) {
    return String(d.claimJudgment || '').toUpperCase().indexOf(type) >= 0 ? '✓' : '';
  };
  tpl.rejectReasonChecked = function(idx) {
    for (var i = 0; i < rejectReasons.length; i++) {
      if (String(rejectReasons[i]) === String(idx)) return '✓';
    }
    return '';
  };

  return tpl.evaluate()
    .setTitle('WCF Preview: ' + wcfId)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ================================================================
//  WCF: LEGACY PDF GENERATION (Sheet-based) — Preserved as fallback
//  Renamed from generateWcfPdf_ to generateWcfPdfViaSheet_
//  NEVER writes to the original WCF FORM template.
//  NEVER calls merge/unmerge/breakApart/mergeAcross.
// ================================================================

/**
 * Create a temporary copy of the WCF FORM template sheet.
 * NOTE: Only used by the LEGACY Sheet-based PDF generator (generateWcfPdfViaSheet_).
 * @param {Spreadsheet} ss  - The spreadsheet
 * @param {string}      wcfId - Used for the temp sheet name
 * @return {Sheet} The temporary sheet (caller must delete it later)
 */
function createWcfPrintSheet_(ss, wcfId) {
  var templateSheet = ss.getSheetByName(SHEET_WCF_FORM);
  if (!templateSheet) {
    throw new Error('WCF FORM template sheet not found. Ensure "' + SHEET_WCF_FORM + '" exists.');
  }
  var safeName = ('WCF_PRINT_' + wcfId).substring(0, 100);
  var tempSheet = templateSheet.copyTo(ss).setName(safeName);
  Logger.log('[WCF PDF] Created temp sheet: ' + safeName + '  gid=' + tempSheet.getSheetId());
  return tempSheet;
}

/**
 * Export one sheet as a PDF blob (legal, portrait).
 * @param {Spreadsheet} ss
 * @param {number}      gid - The sheet's getSheetId()
 * @param {string}      filename - e.g. 'WCF-20260211-0001.pdf'
 * @return {Blob}
 */
function exportSheetToPdf_(ss, gid, filename) {
  var exportUrl = ss.getUrl().replace(/\/edit.*$/, '')
    + '/export?format=pdf'
    + '&gid=' + gid
    + '&size=legal&portrait=true&scale=4&gridlines=false'
    + '&sheetnames=false&printtitle=false&pagenumbers=false&fzr=false'
    + '&top_margin=0.3&bottom_margin=0.3&left_margin=0.4&right_margin=0.4';

  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF export returned HTTP ' + response.getResponseCode());
  }
  return response.getBlob().setName(filename);
}

/**
 * Delete a temporary sheet (best-effort, non-fatal).
 * @param {Spreadsheet} ss
 * @param {Sheet}       sheet
 */
function cleanupTempSheet_(ss, sheet) {
  if (!sheet) return;
  try {
    Logger.log('[WCF PDF] Deleting temp sheet: ' + sheet.getName());
    ss.deleteSheet(sheet);
  } catch (e) {
    Logger.log('[WCF PDF] Could not delete temp sheet: ' + e.message);
  }
}

/**
 * Find the first blank row in a range (for parts table blank-row rule).
 * @param {Sheet}  sheet
 * @param {number} startRow - 1-based first data row
 * @param {number} endRow   - 1-based last data row
 * @param {number} keyColumn - 1-based column to check for emptiness
 * @return {number} Row number, or -1 if the table is full
 */
function findNextBlankRow_(sheet, startRow, endRow, keyColumn) {
  for (var row = startRow; row <= endRow; row++) {
    var val = sheet.getRange(row, keyColumn).getValue();
    if (!val || String(val).trim() === '') return row;
  }
  return -1; // Table full
}

// ─────────────────────────────────────────────────────────
//  Main orchestrator
// ─────────────────────────────────────────────────────────

/**
 * Generate WCF PDF — MAIN ENTRY POINT.
 * Routes to HTML-based renderer (primary) with Sheet-based fallback.
 *
 * @param {object} d - Form data object (must include d.wcfId)
 * @return {object} {success, url, id}
 */
function generateWcfPdf_(d) {
  // ── Primary: HTML/CSS-based PDF (no Sheets dependency) ──
  try {
    Logger.log('[WCF PDF] Using HTML/CSS renderer for: ' + d.wcfId);
    return generateWcfPdfFromHtml_(d.wcfId, d);
  } catch (htmlErr) {
    Logger.log('[WCF PDF] HTML renderer failed: ' + htmlErr.message + '. Falling back to Sheet-based.');
  }

  // ── Fallback: Sheet-based PDF (legacy) ──
  return generateWcfPdfViaSheet_(d);
}

/**
 * LEGACY: Generate WCF PDF via Sheets temp-copy pipeline.
 *   lock → copy template → remove stale images → populate → export → save → delete temp → unlock
 *
 * @param {object} d - Form data object
 * @return {object} {success, url, id}
 */
function generateWcfPdfViaSheet_(d) {
  var lock = LockService.getDocumentLock();
  var lockAcquired = false;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tempSheet = null;

  try {
    lockAcquired = lock.tryLock(30000);
    if (!lockAcquired) {
      throw new Error('Could not acquire lock for PDF generation. Please try again.');
    }

    Logger.log('[WCF PDF] ── START ── ' + d.wcfId);
    Logger.log('[WCF PDF] Safety: this path uses ZERO merge/unmerge/breakApart calls.');

    // 1. Copy template → temp sheet
    tempSheet = createWcfPrintSheet_(ss, d.wcfId);

    // 2. Remove ALL images from the temp copy (clean slate for this submission)
    var oldImages = tempSheet.getImages();
    for (var oi = 0; oi < oldImages.length; oi++) {
      try { oldImages[oi].remove(); } catch (e) {}
    }

    // 3. Populate temp sheet (single-cell writes ONLY, safety-checked)
    populateWcfPrintSheet_(tempSheet, d);

    // 4. Apply print formatting (wrapping, row heights, alignment) — temp sheet only
    formatWcfPrintSheet_(tempSheet);

    // 5. Flush + wait for images
    SpreadsheetApp.flush();
    Utilities.sleep(1500);

    // 6. Export PDF
    var pdfBlob = exportSheetToPdf_(ss, tempSheet.getSheetId(), 'WCF-' + d.wcfId + '.pdf');
    Logger.log('[WCF PDF] PDF blob generated: ' + pdfBlob.getName());

    // 6. Save to Drive:  FSC_WRC_WLP_Files / WCF / YYYY / MM / WCF-ID /
    var rootFolder;
    var rf = DriveApp.getFoldersByName('FSC_WRC_WLP_Files');
    rootFolder = rf.hasNext() ? rf.next() : DriveApp.createFolder('FSC_WRC_WLP_Files');

    var wcfFolder;
    var wf = rootFolder.getFoldersByName('WCF');
    wcfFolder = wf.hasNext() ? wf.next() : rootFolder.createFolder('WCF');

    var now = new Date();
    var year = Utilities.formatDate(now, 'Asia/Manila', 'yyyy');
    var month = Utilities.formatDate(now, 'Asia/Manila', 'MM');

    var yearFolder;
    var yf = wcfFolder.getFoldersByName(year);
    yearFolder = yf.hasNext() ? yf.next() : wcfFolder.createFolder(year);

    var monthFolder;
    var mf = yearFolder.getFoldersByName(month);
    monthFolder = mf.hasNext() ? mf.next() : yearFolder.createFolder(month);

    var wcfIdFolder;
    var widf = monthFolder.getFoldersByName(d.wcfId);
    wcfIdFolder = widf.hasNext() ? widf.next() : monthFolder.createFolder(d.wcfId);

    var pdfFile = wcfIdFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log('[WCF PDF] ── DONE ── saved: ' + pdfFile.getUrl());
    return { success: true, url: pdfFile.getUrl(), id: pdfFile.getId() };

  } catch (err) {
    Logger.log('[WCF PDF] ERROR: ' + err.message);
    throw err;
  } finally {
    // 7. Always delete temp sheet + release lock
    cleanupTempSheet_(ss, tempSheet);
    if (lockAcquired) lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────
//  WCF Print Formatting (temp sheet only)
// ─────────────────────────────────────────────────────────

/**
 * Apply print-optimized formatting to the WCF temp sheet to prevent text cutoffs.
 * This ensures readable PDFs with proper wrapping, row heights, and alignment.
 * 
 * IMPORTANT: Runs ONLY on temp sheet copies (never on original WCF FORM template).
 * 
 * Fixes applied:
 *  1. Text wrapping + row heights for narrative boxes and long address fields
 *  2. Vertical alignment (middle) for single-line identity fields
 *  3. Horizontal alignment for specific fields (job order, dates, etc.)
 * 
 * @param {Sheet} sheet - The temp WCF print sheet (WCF_PRINT_<wcfId>)
 */
function formatWcfPrintSheet_(sheet) {
  // Safety check: never format the original template
  if (sheet.getName() === SHEET_WCF_FORM) {
    throw new Error('Safety stop: attempted to format the original WCF FORM template. Aborting.');
  }

  try {
    Logger.log('[WCF FORMAT] ═══ START PRINT FORMATTING ═══');
    Logger.log('[WCF FORMAT] Target sheet: ' + sheet.getName());

    // ═══════════════════════════════════════════════════════════
    //  1. NARRATIVE BOXES: Wrapping + increased row heights
    // ═══════════════════════════════════════════════════════════
    //  These boxes contain multi-line user text input and need wrapping
    //  to display long text without cutoffs
    
    var narrativeBoxes = [
      { name: 'Problem Complaint',     rows: [17, 18], cols: 'A:AA' },
      { name: 'Probable Cause',        rows: [20, 21], cols: 'A:AA' },
      { name: 'Corrective Action',     rows: [23, 24], cols: 'A:AA' },
      { name: 'Suggestion/Remarks',    rows: [26, 27], cols: 'A:AA' }
    ];

    for (var i = 0; i < narrativeBoxes.length; i++) {
      var box = narrativeBoxes[i];
      for (var r = 0; r < box.rows.length; r++) {
        var rowNum = box.rows[r];
        var range = sheet.getRange(box.cols + rowNum);
        
        // Enable wrapping for the entire row span of the narrative box
        range.setWrap(true);
        
        // Increase row height to accommodate wrapped text (adjust based on typical content)
        // Legal PDF layout can handle taller rows; prioritize readability
        sheet.setRowHeight(rowNum, 55);
        
        Logger.log('[WCF FORMAT] ' + box.name + ' row ' + rowNum + ': wrap=true, height=55');
      }
    }

    // ═══════════════════════════════════════════════════════════
    //  2. ADDRESS FIELDS: Wrapping for long addresses
    // ═══════════════════════════════════════════════════════════
    //  Dealer and Customer addresses can be long; enable wrapping
    
    var addressRows = [
      { name: 'Dealer Name',        row: 7,  cols: 'G:P' },
      { name: 'Dealer Address',     row: 8,  cols: 'G:P' },
      { name: 'Customer Name',      row: 9,  cols: 'G:P' },
      { name: 'Customer Address',   row: 10, cols: 'G:P' }
    ];

    for (var j = 0; j < addressRows.length; j++) {
      var addr = addressRows[j];
      var addrRange = sheet.getRange(addr.cols + addr.row);
      addrRange.setWrap(true);
      
      // Increase row height slightly for addresses (moderate increase)
      sheet.setRowHeight(addr.row, 28);
      
      Logger.log('[WCF FORMAT] ' + addr.name + ' row ' + addr.row + ': wrap=true, height=28');
    }

    // ═══════════════════════════════════════════════════════════
    //  3. VERTICAL ALIGNMENT: Middle for single-line fields
    // ═══════════════════════════════════════════════════════════
    //  Prevents "looks unfilled" issue caused by vertical clipping
    
    var singleLineFields = [
      { name: 'Identity Fields',       range: 'G7:P15' },   // Main left-side fields
      { name: 'Right-Side WRC Fields', range: 'Q11:AA15' }, // WRC no, dates
      { name: 'Job Order & Prep Date', range: 'Q9:AA10' },  // Prepared date area
      { name: 'Parts Causal Row',      range: 'A32:P32' },  // Causal part row
      { name: 'Parts Affected Rows',   range: 'A34:P40' },  // Affected parts rows
      { name: 'Signature Names',       range: 'A44:AA44' }, // Name fields row 44
      { name: 'Units Same Problem',    range: 'A38:P38' },  // Units field
      { name: 'Deliver To',            range: 'A40:P40' }   // Deliver to field
    ];

    for (var k = 0; k < singleLineFields.length; k++) {
      var field = singleLineFields[k];
      var fieldRange = sheet.getRange(field.range);
      fieldRange.setVerticalAlignment('middle');
      
      Logger.log('[WCF FORMAT] ' + field.name + ' (' + field.range + '): vertical=middle');
    }

    // ═══════════════════════════════════════════════════════════
    //  4. HORIZONTAL ALIGNMENT: Specific field adjustments
    // ═══════════════════════════════════════════════════════════
    //  Job Order No., dates, and certain fields benefit from left or center alignment
    
    var alignments = [
      { name: 'Job Order No.',    range: 'W6:AA15', align: 'left' },   // Job Order box
      { name: 'Dates (right)',    range: 'Q11:AA15', align: 'left' },  // Dates on right side
      { name: 'WRC No.',          range: 'Q11:AA11', align: 'left' },  // WRC No. field
      { name: 'Parts Table',      range: 'A32:P40', align: 'left' }    // Parts table
    ];

    for (var m = 0; m < alignments.length; m++) {
      var alignment = alignments[m];
      var alignRange = sheet.getRange(alignment.range);
      alignRange.setHorizontalAlignment(alignment.align);
      
      Logger.log('[WCF FORMAT] ' + alignment.name + ' (' + alignment.range + '): horizontal=' + alignment.align);
    }

    // ═══════════════════════════════════════════════════════════
    //  5. NO-WRAP for fields that must stay single-line
    // ═══════════════════════════════════════════════════════════
    //  Job Order, WRC No, Frame/Engine No should not wrap
    
    var noWrapFields = [
      { name: 'Job Order No.',  range: 'W6:AA6' },
      { name: 'WRC No.',        range: 'Q11:AA11' },
      { name: 'Frame No.',      range: 'G12:P12' },
      { name: 'Engine No.',     range: 'G13:P13' },
      { name: 'Model/Color',    range: 'G11:P11' }
    ];

    for (var n = 0; n < noWrapFields.length; n++) {
      var noWrap = noWrapFields[n];
      var noWrapRange = sheet.getRange(noWrap.range);
      noWrapRange.setWrap(false);
      
      Logger.log('[WCF FORMAT] ' + noWrap.name + ' (' + noWrap.range + '): wrap=false');
    }

    Logger.log('[WCF FORMAT] ═══ FORMATTING COMPLETE ═══');
    Logger.log('[WCF FORMAT] Summary:');
    Logger.log('[WCF FORMAT]   - Narrative boxes (4): wrap + row height 55');
    Logger.log('[WCF FORMAT]   - Address fields (4): wrap + row height 28');
    Logger.log('[WCF FORMAT]   - Vertical alignment: middle (8 ranges)');
    Logger.log('[WCF FORMAT]   - Horizontal alignment: left (4 ranges)');
    Logger.log('[WCF FORMAT]   - No-wrap fields: 5 single-line fields');

  } catch (err) {
    Logger.log('[WCF FORMAT] ERROR: ' + err.message);
    Logger.log('[WCF FORMAT] Stack: ' + err.stack);
    throw new Error('Failed to format print sheet: ' + err.message);
  }
}

// ─────────────────────────────────────────────────────────
//  Populate (single-cell writes only, merge-safe)
// ─────────────────────────────────────────────────────────

/**
 * Populate the WCF temp print sheet with form data.
 * Uses single-cell writes with merged-parent detection and fillable-zone guards.
 * 
 * @param {Sheet}  sheet - Temp print sheet (WCF_PRINT_<wcfId>)
 * @param {object} d     - Form data object
 */

/**
 * Populate a TEMP COPY of WCF FORM with form data.
 *
 * NEW ARCHITECTURE (per user requirements):
 *  • Throws immediately if the sheet IS the original template.
 *  • Uses writeIfAllowed_() for ALL writes (merged-parent detection + fillable-zone guard).
 *  • Every setValue is logged with fieldKey → resolved A1 = value.
 *  • NEVER calls merge / unmerge / breakApart.
 *  • Problem sections: writes ONLY user data to DATA cells (not label rows).
 *  • Parts table: uses findNextBlankRow_() so rows are never overwritten.
 *  • MC Usage checkboxes: places ✓ in appropriate checkbox cell.
 *  • Claim judgment checkboxes: places ✓ in selected judgment/reason cells.
 *  • Illustration image: embedded in Q:AA box (must show in PDF).
 *
 * @param {Sheet}  sheet - The temporary print sheet (NOT WCF FORM!)
 * @param {object} d     - Form data
 */
function populateWcfPrintSheet_(sheet, d) {
  // ── Safety assertion ──────────────────────────────────────────
  if (sheet.getName() === SHEET_WCF_FORM) {
    throw new Error('Safety stop: attempted to write to the original WCF FORM template. Aborting.');
  }

  var cells = WCF_TEMPLATE_CELLS;

  try {
    Logger.log('[WCF POPULATE] ═══ START POPULATION ═══');
    Logger.log('[WCF POPULATE] All writes use merged-parent detection + fillable-zone guards.');

    // ══════════════════════════════════════════════════════════════
    //  Section 1: Main identity fields (Rows 7–15)
    // ══════════════════════════════════════════════════════════════

    // Left-side text inputs (columns G:P within rows 7–15)
    writeIfAllowed_(sheet, cells.dealerName,      d.dealerName,      'dealerName');
    writeIfAllowed_(sheet, cells.dealerAddress,   d.dealerAddress,   'dealerAddress');
    writeIfAllowed_(sheet, cells.customerName,    d.customerName,    'customerName');
    writeIfAllowed_(sheet, cells.customerAddress, d.customerAddress, 'customerAddress');
    
    // Model / Color combined
    var modelColor = (d.model || '') + (d.color ? ' / ' + d.color : '');
    writeIfAllowed_(sheet, cells.modelColor, modelColor, 'modelColor');
    
    writeIfAllowed_(sheet, cells.frameNo,      d.frameNo,      'frameNo');
    writeIfAllowed_(sheet, cells.engineNo,     d.engineNo,     'engineNo');
    writeIfAllowed_(sheet, cells.kilometerRun, d.kilometerRun, 'kilometerRun');

    // Prepared date and Job Order No. (top-right header area)
    writeIfAllowed_(sheet, cells.preparedDate, d.preparedDate, 'preparedDate');
    writeIfAllowed_(sheet, cells.jobOrderNo,   d.jobOrderNo,   'jobOrderNo');

    // Right-side detail fields (WRC section, rows 11–15)
    writeIfAllowed_(sheet, cells.wrcNo,        d.wrcNo,        'wrcNo');
    writeIfAllowed_(sheet, cells.purchaseDate, d.purchaseDate, 'purchaseDate');
    writeIfAllowed_(sheet, cells.failureDate,  d.failureDate,  'failureDate');
    writeIfAllowed_(sheet, cells.reportedDate, d.reportedDate, 'reportedDate');
    writeIfAllowed_(sheet, cells.repairDate,   d.repairDate,   'repairDate');

    // MC Usage checkboxes (Row 14)
    // Place ✓ in the appropriate checkbox cell based on d.mcUsage value
    if (d.mcUsage) {
      var mcUsageUpper = String(d.mcUsage).toUpperCase();
      if (mcUsageUpper.indexOf('SOLO') >= 0) {
        writeIfAllowed_(sheet, cells.mcUsageSolo, '✓', 'mcUsageSolo');
      } else if (mcUsageUpper.indexOf('TRICYCLE') >= 0) {
        writeIfAllowed_(sheet, cells.mcUsageTricycle, '✓', 'mcUsageTricycle');
      } else {
        writeIfAllowed_(sheet, cells.mcUsageOthers, '✓', 'mcUsageOthers');
        if (d.mcUsageOther) {
          writeIfAllowed_(sheet, cells.mcUsageOthersText, d.mcUsageOther, 'mcUsageOthersText');
        }
      }
    }

    // ══════════════════════════════════════════════════════════════
    //  Section 2: Narrative boxes (Problem sections)
    // ══════════════════════════════════════════════════════════════
    //  Template has labels; write ONLY user data to blank data cells

    writeIfAllowed_(sheet, cells.problemComplaint,  d.problemComplaint,  'problemComplaint');
    writeIfAllowed_(sheet, cells.probableCause,     d.probableCause,     'probableCause');
    writeIfAllowed_(sheet, cells.correctiveAction,  d.correctiveAction,  'correctiveAction');
    writeIfAllowed_(sheet, cells.suggestionRemarks, d.suggestionRemarks, 'suggestionRemarks');

    // ══════════════════════════════════════════════════════════════
    //  Section 3: Parts recommended table
    // ══════════════════════════════════════════════════════════════

    // Causal Part / Main Defective Part (Row 32)
    writeIfAllowed_(sheet, cells.causalPartNo,   d.causalPartNo,   'causalPartNo');
    writeIfAllowed_(sheet, cells.causalPartName, d.causalPartName, 'causalPartName');
    writeIfAllowed_(sheet, cells.causalPartQty,  d.causalPartQty,  'causalPartQty');

    // Affected Parts (Rows 34–40) - blank-row detection
    var affectedParts = [];
    try { if (d.affectedParts) affectedParts = JSON.parse(d.affectedParts); } catch (e) {}

    for (var i = 0; i < affectedParts.length; i++) {
      var part = affectedParts[i];
      var blankRow = findNextBlankRow_(sheet, cells.affectedPartsStart, cells.affectedPartsEnd, cells.affectedPartNoCol);
      if (blankRow < 0) {
        Logger.log('[WCF POPULATE] Affected-parts table full (rows 34-40) — cannot write part ' + (i + 1));
        break;
      }
      // Write part data to blank row (columns A, C, G for Part No, Name, Qty)
      if (part.partNo)   sheet.getRange(blankRow, cells.affectedPartNoCol).setValue(part.partNo);
      if (part.partName) sheet.getRange(blankRow, cells.affectedPartNameCol).setValue(part.partName);
      if (part.qty)      sheet.getRange(blankRow, cells.affectedPartQtyCol).setValue(part.qty);
      
      // Additional columns L, M, N, O, P (if used by template and data provided)
      // These are placeholders - adjust based on actual data structure
      // if (part.fieldL) sheet.getRange(blankRow, cells.affectedPartColL).setValue(part.fieldL);
      
      Logger.log('[WCF POPULATE] affectedPart[' + i + '] → row ' + blankRow + 
                 ' | partNo=' + (part.partNo || '') + ' | partName=' + (part.partName || ''));
    }

    // Units with same problem
    writeIfAllowed_(sheet, cells.unitsSameProblem, d.unitsSameProblem, 'unitsSameProblem');

    // ══════════════════════════════════════════════════════════════
    //  Section 4: Illustration image (Q:AA box)
    // ══════════════════════════════════════════════════════════════
    //  Must be embedded and visible in exported PDF

    if (d.illustrationUrl) {
      try {
        var fileId = extractDriveFileId_(d.illustrationUrl);
        if (fileId) {
          var imgBlob = DriveApp.getFileById(fileId).getBlob();
          var img = sheet.insertImage(imgBlob, cells.illustrationCol, cells.illustrationRow);
          
          // Scale to fit Q:AA box (approx 11 columns × ~10 rows)
          // Adjust maxW/maxH based on actual template box size
          var maxW = 400, maxH = 300;
          var scale = Math.min(maxW / img.getWidth(), maxH / img.getHeight(), 1);
          if (scale < 1) {
            img.setWidth(Math.round(img.getWidth() * scale));
            img.setHeight(Math.round(img.getHeight() * scale));
          }
          Logger.log('[WCF POPULATE] Illustration inserted at Q' + cells.illustrationRow + 
                     ' | size=' + img.getWidth() + 'x' + img.getHeight());
        }
      } catch (imgErr) {
        Logger.log('[WCF POPULATE] Illustration insert failed: ' + imgErr.message);
      }
    }

    // Deliver To
    writeIfAllowed_(sheet, cells.deliverTo, d.deliverTo, 'deliverTo');

    // ══════════════════════════════════════════════════════════════
    //  Section 5: Names + signatures (rows 44–48)
    // ══════════════════════════════════════════════════════════════
    //  Three blocks: A:I (Prepared by), J:R (Acknowledged by), S:AA (Warranty Repair)

    writeIfAllowed_(sheet, cells.preparedByName,     d.preparedByName,     'preparedByName');
    writeIfAllowed_(sheet, cells.acknowledgedByName, d.acknowledgedByName, 'acknowledgedByName');
    writeIfAllowed_(sheet, cells.warrantyRepairName, d.warrantyRepairName, 'warrantyRepairName');

    // Signature IMAGES (inserted into signature areas)
    insertSignatureImages_(sheet, d);

    // ══════════════════════════════════════════════════════════════
    //  Section 6: Claim judgment checkboxes (rows 51–56)
    // ══════════════════════════════════════════════════════════════
    //  Place ✓ in selected judgment and rejection reason cells

    // Claim Judgment: Accepted / Rejected / Returned
    if (d.claimJudgment) {
      var judgment = String(d.claimJudgment).toUpperCase();
      if (judgment.indexOf('ACCEPTED') >= 0) {
        writeIfAllowed_(sheet, cells.claimAccepted, '✓', 'claimAccepted');
      } else if (judgment.indexOf('REJECTED') >= 0) {
        writeIfAllowed_(sheet, cells.claimRejected, '✓', 'claimRejected');
      } else if (judgment.indexOf('RETURNED') >= 0) {
        writeIfAllowed_(sheet, cells.claimReturned, '✓', 'claimReturned');
      }
    }

    // Reasons of rejection (if rejection reasons are provided as array or flags)
    // Adjust based on actual data structure - assuming d.rejectReasons is an array of reason indices
    if (d.rejectReasons && Array.isArray(d.rejectReasons)) {
      for (var j = 0; j < d.rejectReasons.length; j++) {
        var reasonIdx = d.rejectReasons[j];
        var reasonCell = cells['rejectReason' + reasonIdx];
        if (reasonCell) {
          writeIfAllowed_(sheet, reasonCell, '✓', 'rejectReason' + reasonIdx);
        }
      }
    }

    Logger.log('[WCF POPULATE] ═══ POPULATION COMPLETE ═══');

  } catch (err) {
    Logger.log('[WCF POPULATE] ERROR: ' + err.message);
    Logger.log('[WCF POPULATE] Stack: ' + err.stack);
    throw new Error('Failed to populate print sheet: ' + err.message);
  }
}

/**
 * Insert signature PNG images into the temp print sheet.
 * @param {Sheet}  sheet - Temp print sheet
 * @param {object} d     - Data with *SignatureUrl fields
 */
function insertSignatureImages_(sheet, d) {
  var cells = WCF_TEMPLATE_CELLS;
  var signatures = [
    { url: d.preparedBySignatureUrl,     col: cells.preparedBySigCol,     row: cells.preparedBySigRow,     label: 'PreparedBy' },
    { url: d.acknowledgedBySignatureUrl, col: cells.acknowledgedBySigCol, row: cells.acknowledgedBySigRow, label: 'AcknowledgedBy' },
    { url: d.warrantyRepairSignatureUrl, col: cells.warrantyRepairSigCol, row: cells.warrantyRepairSigRow, label: 'WarrantyRepair' }
  ];

  for (var i = 0; i < signatures.length; i++) {
    var sig = signatures[i];
    if (!sig.url) continue;
    try {
      var fileId = extractDriveFileId_(sig.url);
      if (fileId) {
        var blob = DriveApp.getFileById(fileId).getBlob();
        var img = sheet.insertImage(blob, sig.col, sig.row);
        img.setWidth(100);
        img.setHeight(40);
        Logger.log('[WCF POPULATE] Signature ' + sig.label + ' → col=' + sig.col + ' row=' + sig.row);
      }
    } catch (sigErr) {
      Logger.log('[WCF POPULATE] Signature ' + sig.label + ' insert failed: ' + sigErr.message);
    }
  }
}

// ─────────────────────────────────────────────────────────
//  Diagnostic: run this manually to verify cell mapping
// ─────────────────────────────────────────────────────────

/**
 * DEBUG OVERLAY REPORT: Run from Apps Script editor → inspectWcfFormTemplate_()
 * 
 * Programmatically detects merged parent ranges and logs final resolved write cells.
 * Verifies that all mapped cells are within fillable zones and not overwriting template labels.
 * 
 * Output format for each field:
 *   field name → requested target (e.g., "preparedDate → Q9") → resolved merged parent top-left → final written A1 → current value
 * 
 * Use this to confirm:
 *  1. All DATA cells are blank (ready to receive input)
 *  2. All LABEL cells contain expected section labels
 *  3. Merged ranges are correctly detected
 *  4. No cell mapping points to a label/header that should not be overwritten
 */
function inspectWcfFormTemplate_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tpl = ss.getSheetByName(SHEET_WCF_FORM);
  if (!tpl) { 
    Logger.log('ERROR: WCF FORM sheet not found'); 
    return; 
  }

  Logger.log('');
  Logger.log('═══════════════════════════════════════════════════════════════');
  Logger.log('  WCF FORM Template Cell Mapping Debug Overlay Report');
  Logger.log('═══════════════════════════════════════════════════════════════');
  Logger.log('');

  var cells = WCF_TEMPLATE_CELLS;

  // Helper to inspect one cell mapping
  function inspectCell(fieldName, requestedA1) {
    if (!requestedA1) {
      Logger.log('⚠ ' + fieldName + ' → NOT MAPPED');
      return;
    }
    
    var range = tpl.getRange(requestedA1);
    var resolvedA1 = getMergedTopLeftA1_(tpl, requestedA1);
    var value = range.getValue();
    var display = range.getDisplayValue();
    var merged = range.isPartOfMerge();
    var row = range.getRow();
    var col = range.getColumn();
    
    var status = '✓';
    if (display && display.trim() !== '') {
      // Cell has content - might be a label or pre-filled value
      status = '⚠ NON-EMPTY';
    }
    
    Logger.log(status + ' ' + fieldName.padEnd(25) + ' → ' + requestedA1.padEnd(6) + 
               ' | resolved=' + resolvedA1.padEnd(6) + 
               ' | merged=' + String(merged).padEnd(5) + 
               ' | row=' + String(row).padStart(2) + ' col=' + String(col).padStart(2) + 
               ' | value="' + String(display).substring(0, 40) + '"');
  }

  // ═══ Section 1: Main identity fields (Rows 7–15) ═══
  Logger.log('');
  Logger.log('─── SECTION 1: Main Identity Fields (Rows 7-15) ───');
  inspectCell('dealerName',        cells.dealerName);
  inspectCell('dealerAddress',     cells.dealerAddress);
  inspectCell('customerName',      cells.customerName);
  inspectCell('customerAddress',   cells.customerAddress);
  inspectCell('modelColor',        cells.modelColor);
  inspectCell('frameNo',           cells.frameNo);
  inspectCell('engineNo',          cells.engineNo);
  inspectCell('kilometerRun',      cells.kilometerRun);
  inspectCell('preparedDate',      cells.preparedDate);
  inspectCell('jobOrderNo',        cells.jobOrderNo);
  inspectCell('wrcNo',             cells.wrcNo);
  inspectCell('purchaseDate',      cells.purchaseDate);
  inspectCell('failureDate',       cells.failureDate);
  inspectCell('reportedDate',      cells.reportedDate);
  inspectCell('repairDate',        cells.repairDate);

  // ═══ MC Usage checkboxes ═══
  Logger.log('');
  Logger.log('─── MC Usage Checkboxes (Row 14) ───');
  inspectCell('mcUsageSolo',       cells.mcUsageSolo);
  inspectCell('mcUsageTricycle',   cells.mcUsageTricycle);
  inspectCell('mcUsageOthers',     cells.mcUsageOthers);
  inspectCell('mcUsageOthersText', cells.mcUsageOthersText);

  // ═══ Section 2: Narrative boxes (Problem sections) ═══
  Logger.log('');
  Logger.log('─── SECTION 2: Narrative Boxes (Rows 17-27) ───');
  Logger.log('NOTE: Template should have LABELS in separate rows/cells.');
  Logger.log('      DATA cells (mapped below) should be BLANK.');
  inspectCell('problemComplaint',  cells.problemComplaint);
  inspectCell('probableCause',     cells.probableCause);
  inspectCell('correctiveAction',  cells.correctiveAction);
  inspectCell('suggestionRemarks', cells.suggestionRemarks);

  // ═══ Section 3: Parts table ═══
  Logger.log('');
  Logger.log('─── SECTION 3: Parts Recommended Table ───');
  inspectCell('causalPartNo',      cells.causalPartNo);
  inspectCell('causalPartName',    cells.causalPartName);
  inspectCell('causalPartQty',     cells.causalPartQty);
  Logger.log('Affected parts rows: ' + cells.affectedPartsStart + '-' + cells.affectedPartsEnd + 
             ' | cols: NoCol=' + cells.affectedPartNoCol + 
             ' NameCol=' + cells.affectedPartNameCol + 
             ' QtyCol=' + cells.affectedPartQtyCol);

  // ═══ Section 4: Illustration ═══
  Logger.log('');
  Logger.log('─── SECTION 4: Illustration Image ───');
  Logger.log('Illustration anchor: col=' + cells.illustrationCol + ' (Q) row=' + cells.illustrationRow);

  // ═══ Section 5: Names + Signatures ═══
  Logger.log('');
  Logger.log('─── SECTION 5: Names + Signatures (Rows 44-48) ───');
  inspectCell('preparedByName',     cells.preparedByName);
  inspectCell('acknowledgedByName', cells.acknowledgedByName);
  inspectCell('warrantyRepairName', cells.warrantyRepairName);
  Logger.log('Signature positions:');
  Logger.log('  PreparedBy     → col=' + cells.preparedBySigCol     + ' row=' + cells.preparedBySigRow);
  Logger.log('  AcknowledgedBy → col=' + cells.acknowledgedBySigCol + ' row=' + cells.acknowledgedBySigRow);
  Logger.log('  WarrantyRepair → col=' + cells.warrantyRepairSigCol + ' row=' + cells.warrantyRepairSigRow);

  // ═══ Section 6: Claim judgment checkboxes ═══
  Logger.log('');
  Logger.log('─── SECTION 6: Claim Judgment Checkboxes (Rows 51-56) ───');
  inspectCell('claimAccepted',     cells.claimAccepted);
  inspectCell('claimRejected',     cells.claimRejected);
  inspectCell('claimReturned',     cells.claimReturned);
  inspectCell('rejectReason1',     cells.rejectReason1);
  inspectCell('rejectReason2',     cells.rejectReason2);
  inspectCell('rejectReason3',     cells.rejectReason3);
  inspectCell('rejectReason4',     cells.rejectReason4);
  inspectCell('rejectReason5',     cells.rejectReason5);

  // ═══ Other fields ═══
  Logger.log('');
  Logger.log('─── OTHER FIELDS ───');
  inspectCell('unitsSameProblem',  cells.unitsSameProblem);
  inspectCell('deliverTo',         cells.deliverTo);

  // ═══ Summary ═══
  Logger.log('');
  Logger.log('═══════════════════════════════════════════════════════════════');
  Logger.log('  INSPECTION COMPLETE');
  Logger.log('═══════════════════════════════════════════════════════════════');
  Logger.log('');
  Logger.log('ACTION REQUIRED:');
  Logger.log('  • ✓ marks = cell is ready (blank or correctly mapped)');
  Logger.log('  • ⚠ NON-EMPTY = cell contains text (verify it is NOT a fillable blank)');
  Logger.log('  • If DATA cells show labels or wrong content, update WCF_TEMPLATE_CELLS');
  Logger.log('  • Run a test submission to verify all fields land in correct positions');
  Logger.log('');
}

/**
 * Send USER thank-you email after WCF submission (AUTOMATIC per fix.md)
 * Professional customer-friendly email with inline logo and PDF action buttons
 * @param {object} params - {userEmail, wcfId, customerName, branch, jobOrderNo, pdfFileId, pdfUrl}
 */
function sendWcfUserThankYouEmail_(params) {
  if (!params.userEmail || !params.wcfId) {
    Logger.log('User thank-you email skipped: missing email or wcfId');
    return;
  }

  var logoBlob = null;
  var logoHtml = '';
  
  // Fetch logo for inline embedding
  try {
    if (WCF_HEADER_IMAGE_ID && String(WCF_HEADER_IMAGE_ID).trim() !== '') {
      logoBlob = DriveApp.getFileById(WCF_HEADER_IMAGE_ID).getBlob().setName('kawasaki-logo');
      logoHtml = '<img src="cid:logo" alt="Kawasaki Logo" style="width:200px; height:auto; margin-bottom:20px;" />';
    }
  } catch (logoErr) {
    Logger.log('User email logo fetch failed: ' + logoErr.message);
  }

  var previewUrl = params.pdfFileId ? 'https://drive.google.com/file/d/' + params.pdfFileId + '/preview' : params.pdfUrl;
  var downloadUrl = params.pdfFileId ? 'https://drive.google.com/uc?export=download&id=' + params.pdfFileId : params.pdfUrl;

  var htmlBody = '<!DOCTYPE html>' +
    '<html><head><meta charset="UTF-8"></head><body style="font-family:Arial,sans-serif; color:#333; line-height:1.6; max-width:600px; margin:0 auto;">' +
    '<div style="text-align:center; padding:30px 20px; background:#f8f9fa; border-bottom:4px solid #007bff;">' +
    logoHtml +
    '<h1 style="margin:10px 0; color:#007bff; font-size:24px;">Thank You for Your Submission!</h1>' +
    '</div>' +
    '<div style="padding:30px 20px; background:#ffffff;">' +
    '<p style="font-size:16px;">Dear <strong>' + (params.customerName || 'Valued Customer') + '</strong>,</p>' +
    '<p style="font-size:14px;">Thank you for submitting your <strong>Warranty Claim Form (WCF)</strong>. We have received your request and are reviewing it.</p>' +
    '<div style="background:#f0f8ff; border-left:4px solid #007bff; padding:15px; margin:20px 0;">' +
    '<p style="margin:5px 0; font-size:14px;"><strong>WCF ID:</strong> ' + params.wcfId + '</p>';
  
  if (params.branch) {
    htmlBody += '<p style="margin:5px 0; font-size:14px;"><strong>Branch:</strong> ' + params.branch + '</p>';
  }
  
  if (params.jobOrderNo) {
    htmlBody += '<p style="margin:5px 0; font-size:14px;"><strong>Job Order No.:</strong> ' + params.jobOrderNo + '</p>';
  }
  
  htmlBody += '</div>' +
    '<p style="font-size:14px; margin-top:25px;">Your Warranty Claim Form is available as a PDF. You can preview or download it using the buttons below:</p>' +
    '<div style="text-align:center; margin:30px 0;">';

  if (previewUrl) {
    htmlBody += '<a href="' + previewUrl + '" target="_blank" ' +
      'style="display:inline-block; margin:10px; padding:12px 30px; background:#007bff; color:#ffffff; ' +
      'text-decoration:none; border-radius:5px; font-weight:bold; font-size:14px;">📄 Preview PDF</a>';
  }
  
  if (downloadUrl) {
    htmlBody += '<a href="' + downloadUrl + '" target="_blank" ' +
      'style="display:inline-block; margin:10px; padding:12px 30px; background:#28a745; color:#ffffff; ' +
      'text-decoration:none; border-radius:5px; font-weight:bold; font-size:14px;">⬇️ Download PDF</a>';
  }

  htmlBody += '</div>' +
    '<p style="font-size:14px; margin-top:25px;">If you have any questions or need assistance, please contact us at:</p>' +
    '<p style="font-size:14px; margin:5px 0;"><strong>Email:</strong> KawasakiWarranty@kmp.com.ph</p>' +
    '<p style="font-size:14px; margin:5px 0;"><strong>Phone:</strong> KMPC Trunkline (02) 8842-31-40 to 43</p>' +
    '<p style="font-size:14px; margin:5px 0; margin-bottom:25px;"><strong>Visayas:</strong> Local 5290 | <strong>Mindanao:</strong> Local 5291</p>' +
    '<p style="font-size:13px; color:#666; margin-top:30px;">Thank you for choosing Kawasaki!</p>' +
    '</div>' +
    '<div style="padding:20px; text-align:center; background:#f8f9fa; border-top:1px solid #ddd; font-size:12px; color:#666;">' +
    '<p style="margin:5px 0;">Kawasaki Motors (Phils.) Corporation</p>' +
    '<p style="margin:5px 0;">Km. 24 East Service Rd., Cupang, Muntinlupa City</p>' +
    '<p style="margin:5px 0;">Email: KawasakiWarranty@kmp.com.ph</p>' +
    '</div>' +
    '</body></html>';

  var mailOptions = {
    to: params.userEmail,
    subject: '✓ Warranty Claim Form Received - ' + params.wcfId,
    htmlBody: htmlBody
  };

  // Add inline logo if available
  if (logoBlob) {
    mailOptions.inlineImages = { logo: logoBlob };
  }

  try {
    MailApp.sendEmail(mailOptions);
    Logger.log('User thank-you email sent to: ' + params.userEmail);
  } catch (mailErr) {
    Logger.log('Failed to send user thank-you email: ' + mailErr.message);
    throw mailErr;
  }
}

/**
 * Send WCF email manually (called from post-submit modal)
 * This sends to INTERNAL recipients (branch + NolanRalph + GMPC branches)
 * @param {string} wcfId - The WCF ID to send email for
 * @return {object} {success, recipients, message}
 */
function sendWcfEmailManual(wcfId) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_WCF);
    if (!sheet) {
      return { success: false, message: 'WCF sheet not found' };
    }

    // Find the row with this WCF ID
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var wcfIdCol = headers.indexOf('WCF ID');
    var rowIndex = -1;
    var rowData = null;

    for (var i = 1; i < data.length; i++) {
      if (data[i][wcfIdCol] === wcfId) {
        rowIndex = i + 1; // 1-based
        rowData = data[i];
        break;
      }
    }

    if (!rowData) {
      return { success: false, message: 'WCF ID not found: ' + wcfId };
    }

    // Extract data from row
    var formData = {
      branch: rowData[headers.indexOf('Branch')] || '',
      jobOrderNo: rowData[headers.indexOf('Job Order No.')] || '',
      dealerName: rowData[headers.indexOf('Dealer Name')] || '',
      customerName: rowData[headers.indexOf('Customer Name')] || '',
      engineNo: rowData[headers.indexOf('Engine No.')] || '',
      wrcNo: rowData[headers.indexOf('WRC No.')] || ''
    };
    var pdfUrl = rowData[headers.indexOf('PDF URL')] || '';

    // Get recipients
    var recipients = getWcfRecipients_(formData.branch);
    if (!recipients || recipients.length === 0) {
      return { success: false, message: 'No email recipients configured' };
    }

    // Send email using existing function
    sendWcfConfirmation_(wcfId, formData, pdfUrl);

    // Update audit columns in sheet
    var emailSentAtCol = headers.indexOf('Email Sent At');
    var emailRecipientsCol = headers.indexOf('Email Recipients');
    var emailStatusCol = headers.indexOf('Email Status');

    var timestamp = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd HH:mm:ss');
    
    if (emailSentAtCol >= 0) {
      sheet.getRange(rowIndex, emailSentAtCol + 1).setValue(timestamp);
    }
    if (emailRecipientsCol >= 0) {
      sheet.getRange(rowIndex, emailRecipientsCol + 1).setValue(recipients.join(', '));
    }
    if (emailStatusCol >= 0) {
      sheet.getRange(rowIndex, emailStatusCol + 1).setValue('SENT');
    }

    return {
      success: true,
      recipients: recipients,
      message: 'Email sent successfully to ' + recipients.length + ' recipient(s)'
    };

  } catch (err) {
    logError_('WCF email manual send failed: ' + err.message, wcfId || '', '');
    return { success: false, message: 'Error sending email: ' + err.message };
  }
}

/**
 * Send WCF confirmation email to branch + NolanRalph + GMPC branches (NOT to user email)
 */
function sendWcfConfirmation_(wcfId, formData, pdfUrl) {
  var recipients = getWcfRecipients_(formData.branch);
  if (!recipients || recipients.length === 0) return;
  
  var subject = 'WCF Submitted: ' + wcfId + ' - ' + (formData.branch || 'N/A');
  
  var htmlBody = '<!DOCTYPE html>'
    + '<html><head><meta charset="utf-8"></head>'
    + '<body style="font-family:Arial,sans-serif;line-height:1.6;color:#333;max-width:600px;margin:0 auto;padding:0;background:#f4f4f4">'
    + '<div style="background:#fff;margin:20px auto;border-radius:8px;overflow:hidden;box-shadow:0 2px 4px rgba(0,0,0,.1)">'
    + '<div style="text-align:center;padding:20px;background:#fff">'
    + '<img src="https://i.imgur.com/nGVvR3v.png" alt="Giant Moto Pro Logo" style="max-width:200px;height:auto" />'
    + '</div>'
    + '<div style="background:#d1fae5;color:#065f46;padding:25px;text-align:center;border-bottom:4px solid:#10b981">'
    + '<h2 style="margin:0;font-size:24px">New WCF Submission</h2>'
    + '</div>'
    + '<div style="padding:30px">'
    + '<p>A new Warranty Claim Form (WCF) has been submitted.</p>'
    + '<table style="width:100%;background:#f8f9fa;border-radius:6px;margin:20px 0;border-collapse:collapse">'
    + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef;width:40%">WCF ID</td>'
    + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef"><strong>' + wcfId + '</strong></td></tr>'
    + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef">Branch</td>'
    + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (formData.branch || 'N/A') + '</td></tr>'
    + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef">Job Order No.</td>'
    + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (formData.jobOrderNo || 'N/A') + '</td></tr>'
    + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef">Dealer</td>'
    + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (formData.dealerName || 'N/A') + '</td></tr>'
    + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef">Customer</td>'
    + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (formData.customerName || 'N/A') + '</td></tr>'
    + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef">Engine No.</td>'
    + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (formData.engineNo || 'N/A') + '</td></tr>'
    + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px;border-bottom:1px solid #e9ecef">WRC No.</td>'
    + '<td style="color:#333;padding:12px 15px;border-bottom:1px solid #e9ecef">' + (formData.wrcNo || 'N/A') + '</td></tr>'
    + '<tr><td style="font-weight:bold;color:#555;padding:12px 15px">PDF Link</td>'
    + '<td style="color:#333;padding:12px 15px">' + (pdfUrl ? '<a href="' + pdfUrl + '" style="color:#10b981;text-decoration:none;font-weight:bold">View PDF</a>' : 'Pending') + '</td></tr>'
    + '</table>'
    + '<p>Please review the submission and take appropriate action.</p>'
    + '<p>Best regards,<br><strong>FSC/WRC/WLP Management System</strong></p>'
    + '</div>'
    + '<div style="background:#f8f9fa;padding:20px;text-align:center;font-size:.9em;color:#666;border-top:1px solid #ddd">'
    + '<p style="margin:0"><strong>This is an automated message.</strong></p>'
    + '</div>'
    + '</div>'
    + '</body></html>';
  
  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * Extract Google Drive file ID from a URL.
 */
function extractDriveFileId_(url) {
  if (!url) return null;
  var m = String(url).match(/[-\w]{25,}/);
  return m ? m[0] : null;
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

    // Only act on WRC, FSC, WLP, or WCF sheets
    if (['WRC', 'FSC', 'WLP', 'WCF'].indexOf(sheetName) === -1) return;
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
    } else if (sheetName === 'WCF') {
      var wcfIdCol     = findColumnByHeader_(sheet, 'WCF ID');
      var wcfEngCol    = findColumnByHeader_(sheet, 'Engine No.');
      var wcfDealerCol = findColumnByHeader_(sheet, 'Dealer Name');
      if (wcfIdCol > 0)     primaryRef   = String(sheet.getRange(row, wcfIdCol).getValue()).trim();
      if (wcfEngCol > 0)    secondaryRef = String(sheet.getRange(row, wcfEngCol).getValue()).trim();
      if (wcfDealerCol > 0) dealerName   = String(sheet.getRange(row, wcfDealerCol).getValue()).trim();
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

