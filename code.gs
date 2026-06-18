// ============================================================
// SVF UCO BANK – BC/BCA APPOINTMENT SYSTEM
// Code.gs — Main Apps Script Backend
// Sanjivani Vikas Foundation × UCO Bank
// ============================================================

var SPREADSHEET_ID = ''; // Leave blank — script uses the bound spreadsheet
var SHEET_APPOINTMENTS = 'BC_APPOINTMENTS';
var SHEET_BRANCH_MASTER = 'BRANCH_MASTER';
var SHEET_SETTINGS      = 'SETTINGS';

// ── Helpers ──────────────────────────────────────────────────

function getSpreadsheet_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found. Please check your spreadsheet setup.');
  return sheet;
}

// ── doGet ────────────────────────────────────────────────────

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.api) {
    var payload = null;
    if (params.payload) {
      try {
        payload = JSON.parse(params.payload);
      } catch (err) {
        return apiResponse_({ success: false, error: 'Invalid payload JSON: ' + err.message }, params.callback);
      }
    }
    return handleApiRequest_(params.api, payload, params.callback);
  }

  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('BC / BCA Appointment Form – SVF × UCO Bank')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0');
}

function doPost(e) {
  try {
    var body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};
    return handleApiRequest_(body.action, body.payload, null);
  } catch (err) {
    return jsonResponse_({ success: false, error: err.message });
  }
}

function handleApiRequest_(action, payload, callback) {
  try {
    if (action === 'getSettings') return apiResponse_(getSettings(), callback);
    if (action === 'getMasterData') return apiResponse_(getMasterData(), callback);
    if (action === 'submitApplication') return apiResponse_(submitApplication(payload || {}), callback);
    return apiResponse_({ success: false, error: 'Unknown API action: ' + action }, callback);
  } catch (err) {
    return apiResponse_({ success: false, error: err.message }, callback);
  }
}

function apiResponse_(data, callback) {
  if (callback) return jsonpResponse_(data, callback);
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

// ── include() for partials ───────────────────────────────────

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── getSettings ──────────────────────────────────────────────

function getSettings() {
  try {
    var sheet  = getSheet_(SHEET_SETTINGS);
    var data   = sheet.getDataRange().getValues();
    var result = {};
    data.forEach(function(row) {
      if (row[0]) result[String(row[0]).trim()] = String(row[1]).trim();
    });
    return { success: true, data: result };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── getMasterData ─────────────────────────────────────────────
// Returns all branch rows in one call.
// Client JS does all filtering — no repeated server calls.
// Column mapping: B=STATE, C=ZONE, D=BRNCH_ID, E=BRANCH

function getMasterData() {
  try {
    var sheet  = getSheet_(SHEET_BRANCH_MASTER);
    var data   = sheet.getDataRange().getValues();
    var rows   = [];
    // Skip header row (row 0)
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var state    = String(r[1] || '').trim();  // Col B
      var zone     = String(r[2] || '').trim();  // Col C
      var branchId = String(r[3] || '').trim();  // Col D
      var branch   = String(r[4] || '').trim();  // Col E
      if (state && zone && branch) {
        rows.push([state, zone, branchId, branch]);
      }
    }
    return { success: true, data: rows };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── generateReferenceId_ ─────────────────────────────────────
// Called INSIDE submitApplication() while lock is held.

function generateReferenceId_(settingsSheet, prefix) {
  var data    = settingsSheet.getDataRange().getValues();
  var counter = 1;
  var counterRow = -1;

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'ReferenceCounter') {
      counter    = parseInt(data[i][1], 10) || 1;
      counterRow = i + 1; // 1-indexed sheet row
      break;
    }
  }

  var today  = new Date();
  var yyyy   = today.getFullYear();
  var mm     = String(today.getMonth() + 1).padStart(2, '0');
  var dd     = String(today.getDate()).padStart(2, '0');
  var serial = String(counter).padStart(6, '0');
  var refId  = prefix + '-' + yyyy + mm + dd + '-' + serial;

  // Increment counter in sheet
  if (counterRow > 0) {
    settingsSheet.getRange(counterRow, 2).setValue(counter + 1);
  }

  return refId;
}

// ── getSubmissionDateTime_ ───────────────────────────────────

function getSubmissionDateTime_() {
  var now = new Date();
  var dd  = String(now.getDate()).padStart(2, '0');
  var mm  = String(now.getMonth() + 1).padStart(2, '0');
  var yy  = now.getFullYear();
  var hh  = String(now.getHours()).padStart(2, '0');
  var mi  = String(now.getMinutes()).padStart(2, '0');
  var ss  = String(now.getSeconds()).padStart(2, '0');
  return dd + '/' + mm + '/' + yy + ' ' + hh + ':' + mi + ':' + ss;
}

// ── submitApplication ────────────────────────────────────────
// Full column order matches BC_APPOINTMENTS headers exactly.

function submitApplication(formData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // wait up to 15s
  } catch (e) {
    return { success: false, error: 'Server busy. Please wait a moment and try again.' };
  }

  try {
    var settingsSheet     = getSheet_(SHEET_SETTINGS);
    var appointmentsSheet = getSheet_(SHEET_APPOINTMENTS);

    // Read settings
    var settingsData   = settingsSheet.getDataRange().getValues();
    var settingsMap    = {};
    settingsData.forEach(function(r) {
      if (r[0]) settingsMap[String(r[0]).trim()] = String(r[1]).trim();
    });

    var bcPartner = settingsMap['BCPartner'] || 'Sanjivani Vikas Foundation';
    var bank      = settingsMap['Bank']      || 'UCO Bank';
    var prefix    = settingsMap['ReferencePrefix'] || 'SVF-UCO';

    // Generate Reference ID (counter incremented here inside lock)
    var refId    = generateReferenceId_(settingsSheet, prefix);
    var subDT    = getSubmissionDateTime_();

    // Helper: uppercase a string field, return '' if blank
    function toUpper_(val) { return val ? String(val).trim().toUpperCase() : ''; }
    // Helper: return value as plain string (no case change) — for numbers/codes
    function toStr_(val) { return val ? String(val).trim() : ''; }

    // Determine IIBF value: save certificate number or 'NO'
    var iibfValue = (formData.iibfCertified === 'Yes' && formData.iibfCertificate)
      ? toUpper_(formData.iibfCertificate)
      : 'NO';

    // Get last SINo
    var lastRow = appointmentsSheet.getLastRow();
    var siNo    = lastRow > 1 ? lastRow - 1 : 1; // Header is row 1

    // Build row — EXACT column order from BC_APPOINTMENTS
    // SINo | BCPartner | Bank | State | Region | District | Block | Branch | BranchCode |
    // VillageName | VillageCode | CSP Name | ContactNumber | RequestedBy | ApprovedBy |
    // NoofOperators | Cluster | TolName | TolContactNo | TolMailID |
    // BolName | BolContactNo | BolMailID | SAName | SAContactNo | SAMailID |
    // UploadPurpose | Fathers Name | Aadhaar No | Other ID Type | ID No |
    // Agent CIF No | Settlement Account No | Saving Account No | PAN No |
    // Alternate Contact No | Address | Pin Code | Education | DOJ | DOB | Gender |
    // SHG Member Y/N | Physically Challenged | Caste(General/OBC/SC/ST) |
    // IIBF Certificate | Certificate Date | Bank Mitra Any other activity |
    // Replaced agent details | Network Service Provider | Mail I'd |
    // SubmissionDateTime | ReferenceID

    var row = [
      siNo,                                          // SINo
      bcPartner,                                     // BCPartner
      bank,                                          // Bank
      toUpper_(formData.state),                      // State
      toUpper_(formData.zone),                       // Region
      toUpper_(formData.district),                   // District
      toUpper_(formData.block),                      // Block
      toUpper_(formData.branch),                     // Branch
      toStr_(formData.branchCode),                   // BranchCode
      toUpper_(formData.villageName),                // VillageName
      toStr_(formData.villageCode),                  // VillageCode
      toUpper_(formData.cspName),                    // CSP Name
      toStr_(formData.contactNumber),                // ContactNumber
      '',                                            // RequestedBy (hidden)
      '',                                            // ApprovedBy (hidden)
      '',                                            // NoofOperators (hidden)
      '',                                            // Cluster (hidden)
      '',                                            // TolName (hidden)
      '',                                            // TolContactNo (hidden)
      '',                                            // TolMailID (hidden)
      '',                                            // BolName (hidden)
      '',                                            // BolContactNo (hidden)
      '',                                            // BolMailID (hidden)
      '',                                            // SAName (hidden)
      '',                                            // SAContactNo (hidden)
      '',                                            // SAMailID (hidden)
      toUpper_(formData.uploadPurpose),              // UploadPurpose
      toUpper_(formData.fathersName),                // Fathers Name
      toStr_(formData.aadhaarNo),                    // Aadhaar No
      toUpper_(formData.otherIdType),                // Other ID Type
      toUpper_(formData.otherIdNo),                  // ID No
      toStr_(formData.agentCifNo),                   // Agent CIF No
      toStr_(formData.settlementAccount),            // Settlement Account No (text — see format below)
      toStr_(formData.savingAccount),                // Saving Account No (text — see format below)
      toUpper_(formData.panNo),                      // PAN No
      toStr_(formData.altContactNo),                 // Alternate Contact No
      toUpper_(formData.address),                    // Address
      toStr_(formData.pinCode),                      // Pin Code
      toUpper_(formData.education),                  // Education
      toStr_(formData.doj),                          // DOJ
      toStr_(formData.dob),                          // DOB
      toUpper_(formData.gender),                     // Gender
      toUpper_(formData.shgMember),                  // SHG Member Y/N
      toUpper_(formData.physicallyChallenged),        // Physically Challenged
      toUpper_(formData.caste),                      // Caste
      iibfValue,                                     // IIBF Certificate
      toStr_(formData.certificateDate),              // Certificate Date
      toUpper_(formData.bankMitraActivity),          // Bank Mitra Any other activity
      toUpper_(formData.replacedAgent),              // Replaced agent details
      toUpper_(formData.networkProvider),            // Network Service Provider
      toStr_(formData.mailId),                       // Mail I'd (email stays as-is)
      subDT,                                         // SubmissionDateTime
      refId                                          // ReferenceID
    ];

    appointmentsSheet.appendRow(row);

    // ── Force text format on account number columns so leading zeros are preserved ──
    // Column 33 = Settlement Account No, Column 34 = Saving Account No (1-indexed)
    var newRow = appointmentsSheet.getLastRow();
    if (formData.settlementAccount) {
      appointmentsSheet.getRange(newRow, 33).setNumberFormat('@').setValue(toStr_(formData.settlementAccount));
    }
    if (formData.savingAccount) {
      appointmentsSheet.getRange(newRow, 34).setNumberFormat('@').setValue(toStr_(formData.savingAccount));
    }

    return {
      success: true,
      referenceId: refId,
      siNo: siNo,
      submissionDateTime: subDT
    };

  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ── testConnection ────────────────────────────────────────────
// Run this manually from Apps Script editor to verify setup.

function testConnection() {
  var results = {};
  try {
    var ss = getSpreadsheet_();
    results.spreadsheetName = ss.getName();

    var settings = getSettings();
    results.settings = settings.success ? 'OK — ' + Object.keys(settings.data).length + ' keys found' : 'FAIL: ' + settings.error;

    var master = getMasterData();
    results.branchMaster = master.success ? 'OK — ' + master.data.length + ' branches loaded' : 'FAIL: ' + master.error;

    var appt = getSheet_(SHEET_APPOINTMENTS);
    results.appointmentsSheet = 'OK — ' + appt.getLastRow() + ' rows (including header)';

    Logger.log(JSON.stringify(results, null, 2));
  } catch (e) {
    Logger.log('testConnection ERROR: ' + e.message);
  }
  return results;
}