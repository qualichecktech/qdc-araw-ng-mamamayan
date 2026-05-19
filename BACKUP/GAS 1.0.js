// ============================================================
// YAKAP Pre-Registration — Google Apps Script
// QualiCheck Diagnostic Clinic
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Open your Google Sheet
// 2. Extensions → Apps Script → paste this entire file
// 3. Save (Ctrl+S), then run formatSheet() ONCE manually:
//      Run → Run function → formatSheet
// 4. Deploy → New Deployment → Web App
//      Execute as: Me  |  Who has access: Anyone
// 5. Copy the Web App URL → paste in admin.html as SHEET_API_URL
//
// EXTRA UTILITIES (run manually anytime):
//   reformatAllRows() — re-applies row formatting to all data rows
//   formatSheet()     — full reset of headers, widths, rules, etc.
// ============================================================

const SHEET_NAME = 'Registrations';
const TOTAL_COLS = 19;

// ── SECURITY: Secret token ───────────────────────────────────
// Must match SUBMIT_SECRET in index.html
// Change this to any strong passphrase you like.
const SECRET_TOKEN = 'QDC-YAKAP-2026';

// ─── Column map (1-indexed) ──────────────────────────────────
const COL = {
  TIMESTAMP:         1,
  TRACKING:          2,
  LAST_NAME:         3,
  FIRST_NAME:        4,
  MIDDLE_NAME:       5,
  SUFFIX:            6,
  EMAIL:             7,
  BIRTHDATE:         8,
  AGE:               9,
  SEX:               10,
  CONTACT:           11,
  ADDRESS:           12,
  PHILHEALTH_ID:     13,
  DATE_REGISTERED:   14,
  PROVIDER:          15,   // admin-editable
  QDC_STATUS:        16,   // admin-editable
  REGISTRATION_DATE: 17,   // admin-editable
  COMPANY:           18,   // from form
  QDC_PROGRAM_DATE:  19,   // from form
};

// ─── Brand colors ────────────────────────────────────────────
const COLOR = {
  PURPLE:       '#6F2DBD',
  PURPLE_DARK:  '#4A0E8F',
  PURPLE_LIGHT: '#EDE0FF',
  BLUE:         '#0057A8',
  BLUE_LIGHT:   '#D6E8F9',
  YELLOW:       '#FFD100',
  YELLOW_LIGHT: '#FFF8CC',
  TEAL:         '#0A5C6B',
  TEAL_LIGHT:   '#D4EEF2',
  WHITE:        '#FFFFFF',
  LIGHT_BG:     '#F3F6FB',
  ROW_ALT:      '#F8F6FF',
  MUTED:        '#5C7280',
  GREEN:        '#16A34A',
  GREEN_LIGHT:  '#DCFCE7',
  RED:          '#DC2626',
  RED_LIGHT:    '#FEE2E2',
};

// ─── Header labels ───────────────────────────────────────────
const HEADERS = [
  'Timestamp', 'Tracking #', 'Last Name', 'First Name', 'Middle Name',
  'Suffix', 'Email', 'Date of Birth', 'Age', 'Sex', 'Contact', 'Address',
  'PhilHealth ID', 'Date Registered',
  'Provider', 'QDC Status', 'Admin Reg. Date',
  'Company', 'QDC Program Date',
];

// ─── Column widths (px) ──────────────────────────────────────
const COL_WIDTHS = [
  165, // Timestamp
  140, // Tracking #
  120, // Last Name
  120, // First Name
  110, // Middle Name
  70,  // Suffix
  200, // Email
  110, // Date of Birth
  55,  // Age
  65,  // Sex
  120, // Contact
  240, // Address
  140, // PhilHealth ID
  140, // Date Registered
  130, // Provider
  140, // QDC Status
  140, // Admin Reg. Date
  180, // Company
  140, // QDC Program Date
];


// ════════════════════════════════════════════════════════════
//  SECURITY HELPERS
// ════════════════════════════════════════════════════════════

// ── Rate limit settings ──────────────────────────────────────
const RATE_LIMIT_MAX        = 3;    // max submissions per IP per window
const RATE_LIMIT_WINDOW_MS  = 60 * 60 * 1000; // 1-hour window

// Checks the secret token from POST body or GET parameter.
function _isAuthorized(secret) {
  return secret === SECRET_TOKEN;
}

// Rate limiter using PropertiesService (persists across requests).
// Returns true if the IP is allowed, false if it is over the limit.
function _checkRateLimit(ip) {
  const store = PropertiesService.getScriptProperties();
  const key   = 'rl_' + ip.replace(/[^a-zA-Z0-9]/g, '_'); // safe key
  const now   = Date.now();

  let record;
  try {
    record = JSON.parse(store.getProperty(key) || 'null');
  } catch (e) {
    record = null;
  }

  // Reset if window has expired or no prior record
  if (!record || (now - record.windowStart) > RATE_LIMIT_WINDOW_MS) {
    record = { windowStart: now, count: 1 };
    store.setProperty(key, JSON.stringify(record));
    return true; // allowed
  }

  // Within the window — check count
  if (record.count >= RATE_LIMIT_MAX) {
    return false; // blocked
  }

  // Increment and allow
  record.count += 1;
  store.setProperty(key, JSON.stringify(record));
  return true;
}

// Server-side validation of submitted patient data.
// Returns null if valid, or an error message string if invalid.
function _validatePayload(data) {
  if (!data.lastName  || data.lastName.trim().length < 1)  return 'Missing last name.';
  if (!data.firstName || data.firstName.trim().length < 1) return 'Missing first name.';
  if (!data.birthdate)                                      return 'Missing birthdate.';
  if (!data.sex || !['male','female'].includes(data.sex.toLowerCase())) return 'Invalid sex value.';
  if (!data.contact)                                        return 'Missing contact number.';
  if (!/^09\d{9}$/.test(data.contact.replace(/\s/g, '')))  return 'Invalid contact number — must be 09XXXXXXXXX.';
  if (!data.address  || data.address.trim().length < 5)    return 'Missing or too-short address.';
  if (!data.philhealthId)                                   return 'Missing PhilHealth ID.';
  if (!/^\d{2}-\d{9}-\d$/.test(data.philhealthId.trim()))  return 'Invalid PhilHealth ID format (expected ##-#########-#).';
  if (!data.tracking || !/^QC-\d{4}-\d{6}$/.test(data.tracking)) return 'Invalid tracking number format.';
  return null;
}


// ════════════════════════════════════════════════════════════
//  formatSheet() — run manually once after pasting this code
// ════════════════════════════════════════════════════════════
function formatSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  // 1. Tab color
  sheet.setTabColor(COLOR.PURPLE);

  // 2. Headers
  _writeHeaders(sheet);

  // 3. Column widths
  COL_WIDTHS.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 4. Row heights
  sheet.setRowHeight(1, 42);
  const lastDataRow = Math.max(sheet.getLastRow(), 2);
  if (lastDataRow > 1) sheet.setRowHeightsForced(2, lastDataRow - 1, 32);

  // 5. Freeze header row + Tracking # column
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // 6. Alternating row banding
  _applyBanding(sheet);

  // 7. Conditional formatting rules
  _applyConditionalFormatting(sheet);

  // 8. Data validation dropdowns
  _applyDataValidation(sheet);

  // 9. Admin column styling
  _styleAdminColumns(sheet);

  // 10. Sheet-wide formatting defaults
  const maxRows  = sheet.getMaxRows();
  const dataRows = maxRows - 1;

  // Default: no wrap, middle-aligned
  sheet.getRange(2, 1, dataRows, TOTAL_COLS)
    .setWrap(false)
    .setVerticalAlignment('middle')
    .setFontSize(10);

  // Address column: wrap text
  sheet.getRange(2, COL.ADDRESS, dataRows, 1).setWrap(true);

  // Number formats
  sheet.getRange(2, COL.TIMESTAMP,         dataRows, 1).setNumberFormat('MMM d, yyyy  h:mm am/pm');
  sheet.getRange(2, COL.BIRTHDATE,         dataRows, 1).setNumberFormat('MMM d, yyyy');
  sheet.getRange(2, COL.DATE_REGISTERED,   dataRows, 1).setNumberFormat('MMM d, yyyy');
  sheet.getRange(2, COL.REGISTRATION_DATE, dataRows, 1).setNumberFormat('MMM d, yyyy');
  sheet.getRange(2, COL.QDC_PROGRAM_DATE, dataRows, 1).setNumberFormat('MMM d, yyyy');
  sheet.getRange(2, COL.QDC_PROGRAM_DATE, dataRows, 1).setHorizontalAlignment('center');
  sheet.getRange(2, COL.COMPANY, dataRows, 1).setHorizontalAlignment('left');
  sheet.getRange(2, COL.AGE,               dataRows, 1).setNumberFormat('0');

  // Horizontal alignment
  [COL.AGE, COL.SEX, COL.SUFFIX, COL.QDC_STATUS, COL.PROVIDER, COL.REGISTRATION_DATE].forEach(c => {
    sheet.getRange(2, c, dataRows, 1).setHorizontalAlignment('center');
  });
  sheet.getRange(2, COL.QDC_STATUS, dataRows, 1).setFontWeight('bold');

  // 11. Warning protection on patient data columns
  _protectAdminColumns(sheet);

  SpreadsheetApp.flush();

  try {
    SpreadsheetApp.getUi().alert(
      '✅ Sheet formatted!\n\nYAKAP Registrations sheet is ready.\n' +
      'Next step: Deploy → New Deployment → Web App, then copy the URL into admin.html.'
    );
  } catch (e) {
    Logger.log('formatSheet() complete.');
  }
}


// ════════════════════════════════════════════════════════════
//  reformatAllRows() — re-formats every existing data row
// ════════════════════════════════════════════════════════════
function reformatAllRows() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Sheet "' + SHEET_NAME + '" not found.'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No data rows found.'); return; }

  for (let r = 2; r <= lastRow; r++) {
    _formatNewRow(sheet, r);
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('✅ Reformatted ' + (lastRow - 1) + ' row(s).');
}


// ════════════════════════════════════════════════════════════
//  PRIVATE HELPERS
// ════════════════════════════════════════════════════════════

function _writeHeaders(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, TOTAL_COLS);
  headerRange.setValues([HEADERS]);

  sheet.getRange(1, 1, 1, COL.DATE_REGISTERED)
    .setBackground(COLOR.PURPLE)
    .setFontColor(COLOR.WHITE)
    .setFontWeight('bold')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setBorder(
      false, false, true, false, false, false,
      COLOR.PURPLE_DARK, SpreadsheetApp.BorderStyle.SOLID_THICK
    );

  sheet.getRange(1, COL.PROVIDER, 1, 3)
    .setBackground(COLOR.BLUE)
    .setFontColor(COLOR.WHITE)
    .setFontWeight('bold')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setBorder(
      false, true, true, false, false, false,
      COLOR.PURPLE_DARK, SpreadsheetApp.BorderStyle.SOLID_THICK
    );

  sheet.getRange(1, COL.TRACKING, 1, 1)
    .setBackground(COLOR.PURPLE_DARK)
    .setFontSize(12);
}


function _applyBanding(sheet) {
  sheet.getBandings().forEach(b => b.remove());

  const lastRow = Math.max(sheet.getMaxRows(), 100);
  const range   = sheet.getRange(1, 1, lastRow, TOTAL_COLS);

  try {
    range.applyRowBanding(SpreadsheetApp.BandingTheme.CUSTOM, true, false)
      .setHeaderRowColor(COLOR.PURPLE)
      .setFirstRowColor(COLOR.WHITE)
      .setSecondRowColor(COLOR.ROW_ALT);
  } catch (e1) {
    try {
      const banding = range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
      banding.setHeaderRowColor(COLOR.PURPLE);
      banding.setFirstRowColor(COLOR.WHITE);
      banding.setSecondRowColor(COLOR.ROW_ALT);
    } catch (e2) {
      Logger.log('Banding API unavailable — applying manual alternating rows.');
      _manualAlternatingRows(sheet);
    }
  }
}

function _manualAlternatingRows(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  for (let r = 2; r <= lastRow; r++) {
    const bg = (r % 2 === 0) ? COLOR.WHITE : COLOR.ROW_ALT;
    sheet.getRange(r, 1, 1, TOTAL_COLS).setBackground(bg);
  }
}


function _applyConditionalFormatting(sheet) {
  sheet.clearConditionalFormatRules();
  const lastRow = sheet.getMaxRows();
  const rules   = [];

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$P2="QDC-REGISTERED"')
      .setBackground(COLOR.GREEN_LIGHT)
      .setFontColor(COLOR.GREEN)
      .setRanges([sheet.getRange(2, COL.QDC_STATUS, lastRow - 1, 1)])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$P2="Not Registered"')
      .setBackground(COLOR.RED_LIGHT)
      .setFontColor(COLOR.RED)
      .setRanges([sheet.getRange(2, COL.QDC_STATUS, lastRow - 1, 1)])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$O2="With Provider"')
      .setBackground(COLOR.TEAL_LIGHT)
      .setFontColor(COLOR.TEAL)
      .setRanges([sheet.getRange(2, COL.PROVIDER, lastRow - 1, 1)])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$O2="No Provider"')
      .setBackground(COLOR.LIGHT_BG)
      .setFontColor(COLOR.MUTED)
      .setRanges([sheet.getRange(2, COL.PROVIDER, lastRow - 1, 1)])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()
      .setBackground(COLOR.PURPLE_LIGHT)
      .setFontColor(COLOR.PURPLE)
      .setRanges([sheet.getRange(2, COL.TRACKING, lastRow - 1, 1)])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()
      .setBackground(COLOR.YELLOW_LIGHT)
      .setRanges([sheet.getRange(2, COL.REGISTRATION_DATE, lastRow - 1, 1)])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}


function _applyDataValidation(sheet) {
  const lastRow = sheet.getMaxRows();

  sheet.getRange(2, COL.PROVIDER, lastRow - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['With Provider', 'No Provider'], true)
      .setAllowInvalid(false)
      .setHelpText('Select provider status for this patient.')
      .build()
  );

  sheet.getRange(2, COL.QDC_STATUS, lastRow - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Not Registered', 'QDC-REGISTERED'], true)
      .setAllowInvalid(false)
      .setHelpText('Set the QDC registration status.')
      .build()
  );

  sheet.getRange(2, COL.SEX, lastRow - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Male', 'Female'], true)
      .setAllowInvalid(true)
      .build()
  );
}


function _styleAdminColumns(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 2);

  sheet.getRange(2, COL.PROVIDER, lastRow - 1, 3)
    .setBackground(COLOR.BLUE_LIGHT);

  sheet.getRange(1, COL.PROVIDER, lastRow, 1)
    .setBorder(
      null, true, null, null, null, null,
      COLOR.PURPLE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );

  sheet.getRange(2, COL.QDC_STATUS,        lastRow - 1, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(2, COL.PROVIDER,          lastRow - 1, 1).setHorizontalAlignment('center');
  sheet.getRange(2, COL.REGISTRATION_DATE, lastRow - 1, 1).setHorizontalAlignment('center');
}


function _protectAdminColumns(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  sheet.getRange(2, 1, sheet.getMaxRows() - 1, COL.DATE_REGISTERED)
    .protect()
    .setDescription('Patient data — managed by Web App. Edit admin columns (O–Q) directly here if needed.')
    .setWarningOnly(true);
}


function _formatNewRow(sheet, rowNum) {
  sheet.setRowHeight(rowNum, 32);

  sheet.getRange(rowNum, 1, 1, COL.DATE_REGISTERED)
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(false);

  sheet.getRange(rowNum, COL.ADDRESS, 1, 1).setWrap(true);

  sheet.getRange(rowNum, COL.TIMESTAMP,         1, 1).setNumberFormat('MMM d, yyyy  h:mm am/pm');
  sheet.getRange(rowNum, COL.BIRTHDATE,         1, 1).setNumberFormat('MMM d, yyyy');
  sheet.getRange(rowNum, COL.DATE_REGISTERED,   1, 1).setNumberFormat('MMM d, yyyy');
  sheet.getRange(rowNum, COL.REGISTRATION_DATE, 1, 1).setNumberFormat('MMM d, yyyy');
  sheet.getRange(rowNum, COL.AGE,               1, 1).setNumberFormat('0');

  [COL.AGE, COL.SEX, COL.SUFFIX].forEach(c => {
    sheet.getRange(rowNum, c, 1, 1).setHorizontalAlignment('center');
  });

  sheet.getRange(rowNum, COL.PROVIDER, 1, 3)
    .setBackground(COLOR.BLUE_LIGHT)
    .setFontSize(10)
    .setVerticalAlignment('middle');

  sheet.getRange(rowNum, COL.QDC_STATUS, 1, 1)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  sheet.getRange(rowNum, COL.PROVIDER, 1, 1).setHorizontalAlignment('center');

  sheet.getRange(rowNum, COL.REGISTRATION_DATE, 1, 1)
    .setHorizontalAlignment('center')
    .setNumberFormat('MMM d, yyyy');

  sheet.getRange(rowNum, COL.COMPANY, 1, 1)
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');

  sheet.getRange(rowNum, COL.QDC_PROGRAM_DATE, 1, 1)
    .setNumberFormat('MMM d, yyyy')
    .setHorizontalAlignment('center')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  sheet.getRange(rowNum, COL.PROVIDER, 1, 1)
    .setBorder(
      null, true, null, null, null, null,
      COLOR.PURPLE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
}


// ════════════════════════════════════════════════════════════
//  clearRateLimits() — run manually to unblock all IPs
//  Useful if a legit user gets stuck or after a test session.
// ════════════════════════════════════════════════════════════
function clearRateLimits() {
  const store = PropertiesService.getScriptProperties();
  const all   = store.getProperties();
  let cleared = 0;
  for (const key in all) {
    if (key.startsWith('rl_')) {
      store.deleteProperty(key);
      cleared++;
    }
  }
  try {
    SpreadsheetApp.getUi().alert('✅ Cleared ' + cleared + ' rate-limit record(s).');
  } catch (e) {
    Logger.log('Cleared ' + cleared + ' rate-limit record(s).');
  }
}

// ════════════════════════════════════════════════════════════
//  doPost — receive a new pre-registration from index.html
// ════════════════════════════════════════════════════════════
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);

    // ── SECURITY: reject requests without the correct secret ──
    if (!_isAuthorized(data.secret)) {
      lock.releaseLock();
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'Unauthorized.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── SECURITY: rate limiting per IP ────────────────────────
    //const clientIp = e.parameter['X-Forwarded-For'] || 'unknown';
    //if (!_checkRateLimit(clientIp)) {
    //  lock.releaseLock();
    //  return ContentService
    //    .createTextOutput(JSON.stringify({ success: false, error: 'Too many submissions. Please try again later.' }))
    //    .setMimeType(ContentService.MimeType.JSON);
    //}

    // ── SECURITY: server-side validation ──────────────────────
    const validationError = _validatePayload(data);
    if (validationError) {
      lock.releaseLock();
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: validationError }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      formatSheet();
    }

    const now = new Date();
    const row = [
      now,
      data.tracking        || '',
      data.lastName.trim() || '',
      data.firstName.trim()|| '',
      data.middleName      || '',
      data.suffix          || '',
      data.email           || '',
      data.birthdate       || '',
      data.age             || '',
      data.sex             || '',
      data.contact         || '',
      data.address         || '',
      data.philhealthId    || '',
      data.dateRegistered  || '',
      '',                         // Provider — admin fills
      'Not Registered',           // Default QDC Status
      '',                         // Admin Registration Date
      data.company        || '',  // Company
      data.qdcProgramDate || '',  // QDC Program Date
    ];

    sheet.appendRow(row);
    _formatNewRow(sheet, sheet.getLastRow());

    lock.releaseLock();

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, tracking: data.tracking }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    lock.releaseLock();
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function _formatDateCell(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  }
  return String(val);
}

// ════════════════════════════════════════════════════════════
//  doGet — admin list or update
// ════════════════════════════════════════════════════════════
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    // ── SECURITY: reject requests without the correct secret ──
    const secret = e.parameter.secret || '';
    if (!_isAuthorized(secret)) {
      output.setContent(JSON.stringify({ success: false, error: 'Unauthorized.' }));
      return output;
    }

    const action = e.parameter.action || '';
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    let sheet    = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      output.setContent(JSON.stringify({ success: true, records: [] }));
      return output;
    }

    // ── LIST ───────────────────────────────────────────────
    if (action === 'list') {
      const data    = sheet.getDataRange().getValues();
      const records = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[COL.TRACKING - 1]) continue;
        records.push({
          row:             i + 1,
          tracking:        row[COL.TRACKING - 1],
          lastName:        row[COL.LAST_NAME - 1],
          firstName:       row[COL.FIRST_NAME - 1],
          middleName:      row[COL.MIDDLE_NAME - 1],
          suffix:          row[COL.SUFFIX - 1],
          email:           row[COL.EMAIL - 1],
          birthdate:       row[COL.BIRTHDATE - 1],
          age:             row[COL.AGE - 1],
          sex:             row[COL.SEX - 1],
          contact:         row[COL.CONTACT - 1],
          address:         row[COL.ADDRESS - 1],
          philhealthId:    row[COL.PHILHEALTH_ID - 1],
          dateRegistered:  row[COL.DATE_REGISTERED - 1],
          provider:        row[COL.PROVIDER - 1],
          qdcStatus:       row[COL.QDC_STATUS - 1],
          adminRegDate:    _formatDateCell(row[COL.REGISTRATION_DATE - 1]),
          company:         row[COL.COMPANY - 1],
          qdcProgramDate:  row[COL.QDC_PROGRAM_DATE - 1],
          timestamp:       row[COL.TIMESTAMP - 1],
        });
      }
      output.setContent(JSON.stringify({ success: true, records }));
      return output;
    }

    // ── UPDATE ─────────────────────────────────────────────
    if (action === 'update') {
      const rowNum    = parseInt(e.parameter.row, 10);
      const provider  = e.parameter.provider  || '';
      const qdcStatus = e.parameter.qdcStatus || '';
      const adminDate = e.parameter.adminDate || '';

      if (!rowNum || rowNum < 2) throw new Error('Invalid row number');

      sheet.getRange(rowNum, COL.PROVIDER).setValue(provider);
      sheet.getRange(rowNum, COL.QDC_STATUS).setValue(qdcStatus);
      sheet.getRange(rowNum, COL.REGISTRATION_DATE).setValue(adminDate);

      sheet.getRange(rowNum, COL.QDC_STATUS, 1, 1)
        .setHorizontalAlignment('center').setFontWeight('bold');
      sheet.getRange(rowNum, COL.PROVIDER, 1, 1)
        .setHorizontalAlignment('center');

      output.setContent(JSON.stringify({ success: true }));
      return output;
    }

    output.setContent(JSON.stringify({ success: false, error: 'Unknown action' }));
    return output;

  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.message }));
    return output;
  }
}