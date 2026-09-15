/**
 * ============================================================================
 * SANJIVANI VIKAS FOUNDATION × UCO BANK — BC / BCA APPOINTMENT SYSTEM
 * High-Performance, Ultra-Reliable Google Apps Script Backend (Code.gs)
 * ============================================================================
 */

// ── CONFIGURATION & CONSTANTS ────────────────────────────────────────────────
var SPREADSHEET_ID = ''; // Leave blank to auto-bind active Google Spreadsheet

var SHEET_APPOINTMENTS   = 'BC_APPOINTMENTS';
var SHEET_SETTINGS       = 'SETTINGS';
var SHEET_BRANCH_MASTER  = 'BRANCH_MASTER';

var FIXED_RECIPIENT_1 = 'svf.samar@gmail.com';
var FIXED_RECIPIENT_2 = 'sanjivani.uco@gmail.com';

var LOGO_URL_SVF = 'https://res.cloudinary.com/date69bba/image/upload/v1775886016/Poster_Maker_Bg_Remover_22022024_084705_xwusdq.png';
var LOGO_URL_UCO = 'https://res.cloudinary.com/date69bba/image/upload/v1776667693/uco-bank-logo_1_i04gtr.png';

// ── HTTP GET ROUTER (JSONP & JSON API) ───────────────────────────────────────
function doGet(e) {
  var params = (e && e.parameter) || {};
  var callback = params.callback || params.jsonp;

  if (!params.api) {
    return HtmlService.createHtmlOutput(
      '<h3>SVF x UCO Bank BC Appointment System API</h3><p>Status: Online & Operational</p>'
    );
  }

  try {
    var payload = null;
    if (params.payload) {
      try {
        payload = typeof params.payload === 'string' ? JSON.parse(params.payload) : params.payload;
      } catch (err) {
        return apiResponse_({ success: false, error: 'Invalid payload JSON: ' + err.message }, callback);
      }
    }
    return handleApiRequest_(params.api, payload, callback);
  } catch (err) {
    return apiResponse_({ success: false, error: err.message }, callback);
  }
}

// ── HTTP POST ROUTER ─────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      try {
        body = JSON.parse(e.postData.contents);
      } catch (pErr) {
        body = {};
      }
    }
    var action = body.action || (e && e.parameter && e.parameter.api);
    var payload = body.payload || (e && e.parameter && e.parameter.payload);

    if (typeof payload === 'string') {
      try { payload = JSON.parse(payload); } catch (e) {}
    }

    return handleApiRequest_(action, payload, null);
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

// ── API ACTION HANDLER ───────────────────────────────────────────────────────
function handleApiRequest_(action, payload, callback) {
  if (action === 'getSettings') {
    return apiResponse_(getSettings(), callback);
  }
  if (action === 'getMasterData') {
    return apiResponse_(getMasterData(), callback);
  }
  if (action === 'submitApplication') {
    return apiResponse_(submitApplication(payload || {}), callback);
  }
  if (action === 'testMail') {
    return apiResponse_(testMail_(), callback);
  }
  if (action === 'testConnection') {
    return apiResponse_({ success: true, data: testConnection() }, callback);
  }
  return apiResponse_({ success: false, error: 'Unknown API action: ' + action }, callback);
}

// ── RESPONSE FORMATTERS ──────────────────────────────────────────────────────
function apiResponse_(data, callback) {
  if (callback) {
    return jsonpResponse_(data, callback);
  }
  return jsonResponse_(data);
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpResponse_(data, callback) {
  var safeCallback = String(callback).replace(/[^\w.$]/g, '');
  if (!safeCallback) safeCallback = 'callback';
  return ContentService
    .createTextOutput(safeCallback + '(' + JSON.stringify(data) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ── SPREADSHEET HELPERS ──────────────────────────────────────────────────────
function getSpreadsheet_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(sheetName) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initSheetHeaders_(sheet, sheetName);
  }
  return sheet;
}

function initSheetHeaders_(sheet, sheetName) {
  if (sheetName === SHEET_APPOINTMENTS) {
    var headers = [
      'SINo', 'BCPartner', 'Bank', 'State', 'Region', 'District', 'Block', 'Branch', 'BranchCode',
      'VillageName', 'VillageCode', 'CSP Name', 'ContactNumber', 'RequestedBy', 'ApprovedBy',
      'NoofOperators', 'Cluster', 'TolName', 'TolContactNo', 'TolMailID',
      'BolName', 'BolContactNo', 'BolMailID', 'SAName', 'SAContactNo', 'SAMailID',
      'UploadPurpose', 'Fathers Name', 'Aadhaar No', 'Other ID Type', 'ID No',
      'Agent CIF No', 'Settlement Account No', 'Saving Account No', 'PAN No',
      'Alternate Contact No', 'Address', 'Pin Code', 'Education', 'DOJ', 'DOB', 'Gender',
      'SHG Member Y/N', 'Physically Challenged', 'Caste(General/OBC/SC/ST)',
      'IIBF Certificate', 'Certificate Date', 'Bank Mitra Any other activity',
      'Replaced agent details', 'Network Service Provider', 'Mail I\'d',
      'SubmissionDateTime', 'ReferenceID'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheetName === SHEET_SETTINGS) {
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
    sheet.appendRow(['BCPartner', 'SANJIVANI VIKAS FOUNDATION']);
    sheet.appendRow(['Bank', 'UCO BANK']);
    sheet.appendRow(['ReferencePrefix', 'SVF-UCO']);
    sheet.appendRow(['ReferenceCounter', '1']);
    sheet.setFrozenRows(1);
  } else if (sheetName === SHEET_BRANCH_MASTER) {
    sheet.getRange(1, 1, 1, 4).setValues([['State', 'Zone', 'BranchCode', 'BranchName']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

// ── READ SETTINGS ────────────────────────────────────────────────────────────
function getSettings() {
  try {
    var sheet = getSheet_(SHEET_SETTINGS);
    var rows = sheet.getDataRange().getValues();
    var data = {};
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0]) {
        data[String(rows[i][0]).trim()] = String(rows[i][1] || '').trim();
      }
    }
    return { success: true, data: data };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── READ BRANCH MASTER ───────────────────────────────────────────────────────
function getMasterData() {
  try {
    var sheet = getSheet_(SHEET_BRANCH_MASTER);
    var rows = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] && rows[i][2]) {
        result.push([
          String(rows[i][0]).trim(),
          String(rows[i][1] || '').trim(),
          String(rows[i][2] || '').trim(),
          String(rows[i][3] || '').trim()
        ]);
      }
    }
    return { success: true, data: result };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── SUBMIT APPLICATION ───────────────────────────────────────────────────────
function submitApplication(formData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, error: 'Server busy. Please try submitting again in a moment.' };
  }

  try {
    var settingsSheet = getSheet_(SHEET_SETTINGS);
    var appointmentsSheet = getSheet_(SHEET_APPOINTMENTS);

    // Fetch Settings
    var settingsRes = getSettings();
    var settings = settingsRes.data || {};
    var bcPartner = settings['BCPartner'] || 'SANJIVANI VIKAS FOUNDATION';
    var bank = settings['Bank'] || 'UCO BANK';
    var prefix = settings['ReferencePrefix'] || 'SVF-UCO';

    // Thread-safe Reference ID Generation
    var refId = generateReferenceId_(settingsSheet, prefix);
    var subDT = getSubmissionDateTime_();

    function toUpper_(v) { return v ? String(v).trim().toUpperCase() : ''; }
    function toStr_(v) { return v ? String(v).trim() : ''; }

    var iibfVal = (formData.iibfCertified === 'Yes' && formData.iibfCertificate)
      ? toUpper_(formData.iibfCertificate)
      : 'NO';

    var lastRow = appointmentsSheet.getLastRow();
    var siNo = lastRow > 0 ? lastRow : 1;

    var row = [
      siNo,
      bcPartner,
      bank,
      toUpper_(formData.state),
      toUpper_(formData.zone),
      toUpper_(formData.district),
      toUpper_(formData.block),
      toUpper_(formData.branch),
      toStr_(formData.branchCode),
      toUpper_(formData.villageName),
      toStr_(formData.villageCode),
      toUpper_(formData.cspName),
      toStr_(formData.contactNumber),
      '', '', '', '', '', '', '', '', '', '', '', '', '', // Hidden columns
      toUpper_(formData.uploadPurpose),
      toUpper_(formData.fathersName),
      toStr_(formData.aadhaarNo),
      toUpper_(formData.otherIdType),
      toUpper_(formData.otherIdNo),
      toStr_(formData.agentCifNo),
      toStr_(formData.settlementAccount),
      toStr_(formData.savingAccount),
      toUpper_(formData.panNo),
      toStr_(formData.altContactNo),
      toUpper_(formData.address),
      toStr_(formData.pinCode),
      toUpper_(formData.education),
      toStr_(formData.doj),
      toStr_(formData.dob),
      toUpper_(formData.gender),
      toUpper_(formData.shgMember),
      toUpper_(formData.physicallyChallenged),
      toUpper_(formData.caste),
      iibfVal,
      toStr_(formData.certificateDate),
      'NO',
      toUpper_(formData.replacedAgent),
      toUpper_(formData.networkProvider),
      toStr_(formData.mailId),
      subDT,
      refId
    ];

    appointmentsSheet.appendRow(row);

    // Format account number columns as text to preserve leading zeros
    var newRow = appointmentsSheet.getLastRow();
    if (formData.settlementAccount) {
      appointmentsSheet.getRange(newRow, 33).setNumberFormat('@').setValue(toStr_(formData.settlementAccount));
    }
    if (formData.savingAccount) {
      appointmentsSheet.getRange(newRow, 34).setNumberFormat('@').setValue(toStr_(formData.savingAccount));
    }

    // Trigger Email Delivery Pipeline
    var mailResult = generateAndEmailPdf_(formData, refId, siNo, subDT);

    return {
      success: true,
      referenceId: refId,
      siNo: siNo,
      submissionDateTime: subDT,
      mailSent: mailResult ? mailResult.success : false,
      mailError: mailResult ? mailResult.error : ''
    };

  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ── REFERENCE ID GENERATOR ───────────────────────────────────────────────────
function generateReferenceId_(settingsSheet, prefix) {
  var rows = settingsSheet.getDataRange().getValues();
  var counterRowIdx = -1;
  var counterVal = 1;
  var yearMonth = getYearMonthPrefix_();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'ReferenceCounter') {
      counterRowIdx = i + 1;
      counterVal = parseInt(rows[i][1], 10) || 1;
      break;
    }
  }

  var formattedCounter = ('000000' + counterVal).slice(-6);
  var refId = prefix + '-' + yearMonth + '-' + formattedCounter;

  if (counterRowIdx > 0) {
    settingsSheet.getRange(counterRowIdx, 2).setValue(counterVal + 1);
  } else {
    settingsSheet.appendRow(['ReferenceCounter', 2]);
  }

  return refId;
}

function getYearMonthPrefix_() {
  var d = new Date();
  var y = d.getFullYear();
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  return y + m;
}

function getSubmissionDateTime_() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss');
}

// ── RECIPIENT LIST BUILDER ───────────────────────────────────────────────────
function buildRecipientList_(officeMailIdFieldValue) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var seen = {};
  var result = [];

  function addIfValid(addr) {
    var a = (addr || '').trim().toLowerCase();
    if (a && emailRegex.test(a) && !seen[a]) {
      seen[a] = true;
      result.push(a);
    }
  }

  addIfValid(FIXED_RECIPIENT_1);
  addIfValid(FIXED_RECIPIENT_2);

  if (officeMailIdFieldValue) {
    var parts = String(officeMailIdFieldValue).split(',');
    parts.forEach(function(p) { addIfValid(p); });
  }

  return result.join(',');
}

// ── FAST & FAILURE-FREE EMAIL PIPELINE ───────────────────────────────────────
function generateAndEmailPdf_(formData, refId, siNo, subDT) {
  try {
    var recipients = buildRecipientList_(formData.officeMailId);
    if (!recipients) return { success: false, error: 'No valid recipient email addresses.' };

    var d = formData;
    function esc(v) {
      return v ? String(v).trim().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;') : '';
    }

    var agentName  = d.cspName ? String(d.cspName).trim() : 'N/A';
    var branchName = d.branch ? String(d.branch).trim() : 'N/A';
    var zoneName   = d.zone ? String(d.zone).trim() : 'N/A';
    var distName   = d.district ? String(d.district).trim() : 'N/A';
    var contactNo  = d.contactNumber ? String(d.contactNumber).trim() : 'N/A';

    var emailSubject = 'BC/BCA Appointment Application — ' + agentName + ' (' + branchName + ', ' + zoneName + ') — ' + refId;

    var plainBody =
      'A new BC/BCA appointment application has been submitted successfully.\n\n' +
      '• Candidate Name: ' + agentName + '\n' +
      '• Branch: ' + branchName + ' (' + (d.branchCode || '') + ')\n' +
      '• Zone: ' + zoneName + ' / ' + (d.state || '') + '\n' +
      '• District: ' + distName + '\n' +
      '• Contact: ' + contactNo + '\n' +
      '• Reference ID: ' + refId + '\n' +
      '• Serial No: #' + siNo + '\n' +
      '• Date/Time: ' + subDT;

    var htmlEmailBody =
      '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
      '<body style="margin:0;padding:20px;background-color:#f8fafc;font-family:sans-serif;">' +
      '<div style="max-width:680px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e2e8f0;box-shadow:0 4px 6px rgba(0,0,0,0.05);">' +
      '<div style="background:#1a5c1a;padding:20px;text-align:center;color:#ffffff;border-bottom:4px solid #d9531e;">' +
      '<div style="font-size:18px;font-weight:900;letter-spacing:0.04em;">SANJIVANI VIKAS FOUNDATION</div>' +
      '<div style="font-size:12px;font-weight:700;color:#a7f3d0;">BC / BCA Appointment Application System</div>' +
      '</div>' +
      '<div style="padding:24px;color:#0f172a;line-height:1.6;">' +
      '<h3 style="color:#1a5c1a;margin-top:0;">New Application Submitted Successfully</h3>' +
      '<table width="100%" cellpadding="8" cellspacing="0" style="border-collapse:collapse;margin:16px 0;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;">' +
      '<tr><td width="35%" style="border-bottom:1px solid #e2e8f0;font-weight:bold;color:#64748b;">Candidate Name</td><td style="border-bottom:1px solid #e2e8f0;font-weight:bold;color:#0f172a;">' + esc(agentName) + '</td></tr>' +
      '<tr><td style="border-bottom:1px solid #e2e8f0;font-weight:bold;color:#64748b;">Branch & Code</td><td style="border-bottom:1px solid #e2e8f0;color:#1a5c1a;font-weight:bold;">' + esc(branchName) + ' (' + esc(d.branchCode) + ')</td></tr>' +
      '<tr><td style="border-bottom:1px solid #e2e8f0;font-weight:bold;color:#64748b;">Zone & State</td><td style="border-bottom:1px solid #e2e8f0;">' + esc(zoneName) + ' / ' + esc(d.state) + '</td></tr>' +
      '<tr><td style="border-bottom:1px solid #e2e8f0;font-weight:bold;color:#64748b;">District & Block</td><td style="border-bottom:1px solid #e2e8f0;">' + esc(distName) + ' / ' + esc(d.block) + '</td></tr>' +
      '<tr><td style="border-bottom:1px solid #e2e8f0;font-weight:bold;color:#64748b;">Contact Number</td><td style="border-bottom:1px solid #e2e8f0;font-weight:bold;">' + esc(contactNo) + '</td></tr>' +
      '<tr><td style="border-bottom:1px solid #e2e8f0;font-weight:bold;color:#64748b;">Reference ID</td><td style="border-bottom:1px solid #e2e8f0;font-weight:bold;color:#1a5c1a;font-family:monospace;">' + esc(refId) + '</td></tr>' +
      '<tr><td style="font-weight:bold;color:#64748b;">Submission Date</td><td>' + esc(subDT) + '</td></tr>' +
      '</table>' +
      '</div>' +
      '<div style="background:#f1f5f9;padding:12px;text-align:center;font-size:11px;color:#64748b;">Sanjivani Vikas Foundation × UCO Bank</div>' +
      '</div></body></html>';

    var mailSent = false;
    var errDetail = '';

    // Primary: MailApp
    try {
      MailApp.sendEmail({
        to: recipients,
        subject: emailSubject,
        body: plainBody,
        htmlBody: htmlEmailBody,
        name: 'SVF UCO BC Appointment System'
      });
      mailSent = true;
    } catch (mErr) {
      // Fallback: GmailApp
      try {
        GmailApp.sendEmail(recipients, emailSubject, plainBody, {
          htmlBody: htmlEmailBody,
          name: 'SVF UCO BC Appointment System'
        });
        mailSent = true;
      } catch (gErr) {
        errDetail = 'MailApp: ' + mErr.message + ' | GmailApp: ' + gErr.message;
      }
    }

    return { success: mailSent, error: errDetail };

  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── DIAGNOSTIC TEST ROUTINES ─────────────────────────────────────────────────
function testMail_() {
  try {
    var recipients = buildRecipientList_('');
    var testSubject = 'Diagnostic Test Email — ' + new Date().toISOString();
    var testBody = 'This is a diagnostic email from the SVF x UCO Bank BC Appointment System.\n\nRecipients: ' + recipients;

    var mailSent = false;
    var errDetail = '';

    try {
      MailApp.sendEmail({ to: recipients, subject: testSubject, body: testBody, name: 'SVF UCO System Test' });
      mailSent = true;
    } catch (e1) {
      try {
        GmailApp.sendEmail(recipients, testSubject, testBody, { name: 'SVF UCO System Test' });
        mailSent = true;
      } catch (e2) {
        errDetail = 'MailApp: ' + e1.message + ' | GmailApp: ' + e2.message;
      }
    }

    return { success: mailSent, recipients: recipients, error: errDetail };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function testConnection() {
  var results = {};
  try {
    var ss = getSpreadsheet_();
    results.spreadsheetName = ss.getName();

    var settings = getSettings();
    results.settings = settings.success ? 'OK (' + Object.keys(settings.data).length + ' keys)' : 'FAIL: ' + settings.error;

    var master = getMasterData();
    results.branchMaster = master.success ? 'OK (' + master.data.length + ' branches)' : 'FAIL: ' + master.error;

    var appt = getSheet_(SHEET_APPOINTMENTS);
    results.appointmentsSheet = 'OK (' + appt.getLastRow() + ' rows)';

  } catch (e) {
    results.error = e.message;
  }
  return results;
}

function testEmailPdf() {
  var fakeData = {
    state: 'BIHAR', zone: 'PATNA', branch: 'PATNA MAIN', branchCode: 'UCBA0000001',
    district: 'PATNA', block: 'PATNA', villageName: 'KANKARBAGH', villageCode: '800026',
    cspName: 'TEST CANDIDATE', fathersName: 'TEST FATHER', contactNumber: '9955999077',
    officeMailId: 'svf.samar@gmail.com'
  };
  var subDT = getSubmissionDateTime_();
  var res = generateAndEmailPdf_(fakeData, 'SVF-UCO-TEST-000001', 1, subDT);
  Logger.log('testEmailPdf Result: ' + JSON.stringify(res));
}