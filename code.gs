// ============================================================
// SVF UCO BANK – BC/BCA APPOINTMENT SYSTEM
// Code.gs — Main Apps Script Backend
// Sanjivani Vikas Foundation × UCO Bank
// ============================================================

var SPREADSHEET_ID = ''; // Leave blank — script uses the bound spreadsheet
var SHEET_APPOINTMENTS = 'BC_APPOINTMENTS';
var SHEET_BRANCH_MASTER = 'BRANCH_MASTER';
var SHEET_SETTINGS      = 'SETTINGS';

// ── Email / PDF Constants ────────────────────────────────────────
// TODO: Replace the two placeholder values below with real email addresses
// before deploying. These are the fixed recipients on every submission.
var FIXED_RECIPIENT_1 = 'svf.samar@gmail.com'; // TODO: set real address
var FIXED_RECIPIENT_2 = 'sanjivani.uco@gmail.com'; // TODO: set real address
var LOGO_URL_SVF = 'https://res.cloudinary.com/date69bba/image/upload/v1775886016/Poster_Maker_Bg_Remover_22022024_084705_xwusdq.png';
var LOGO_URL_UCO = 'https://res.cloudinary.com/date69bba/image/upload/v1776667693/uco-bank-logo_1_i04gtr.png';

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
    if (action === 'testMail') return apiResponse_(testMail_(), callback);
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

// ── getLogoBase64_ ───────────────────────────────────────────────
// Fetches a logo URL via UrlFetchApp, base64-encodes the bytes, and returns
// a data URI string (e.g. 'data:image/png;base64,....').
// Results are cached in CacheService (TTL 6 h, ~21600 s).
// Apps Script cache values are capped at 100 KB per key; if the encoded
// logo exceeds that, the result is fetched fresh every time rather than
// throwing — the PDF will still render correctly, just without caching.
// Fallback behaviour: if UrlFetchApp fails (e.g. a typo in the URL), the
// function returns an empty string so the PDF renders without that logo
// instead of letting the error surface to the end user.

function getLogoBase64_(url) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'logo_b64_' + url.replace(/[^a-zA-Z0-9]/g, '_').slice(-80);
  var cached = cache.get(cacheKey);
  if (cached) return cached;

  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('getLogoBase64_: non-200 response for ' + url);
      return '';
    }
    var bytes = response.getContent();
    var b64   = Utilities.base64Encode(bytes);
    // Detect MIME type from Content-Type header; fall back to image/png
    var ct    = response.getHeaders()['Content-Type'] || 'image/png';
    var mime  = ct.split(';')[0].trim() || 'image/png';
    var dataUri = 'data:' + mime + ';base64,' + b64;

    // Cache only if within Apps Script's 100 KB per-key limit
    if (dataUri.length <= 100000) {
      cache.put(cacheKey, dataUri, 21600); // TTL = 6 hours
    }
    return dataUri;
  } catch (err) {
    // UrlFetchApp failure (bad URL, network error, etc.) — log and degrade
    // gracefully: the submission already succeeded at this point, so we
    // never let a logo fetch error break anything.
    Logger.log('getLogoBase64_ failed for ' + url + ': ' + err.message);
    return '';
  }
}

// ── buildRecipientList_ ──────────────────────────────────────────
// Builds a comma-joined recipient string for MailApp.sendEmail().
// Always starts with the two fixed recipients, then appends any valid
// addresses from officeMailIdFieldValue (comma-separated input).
// De-duplicates case-insensitively. Never uses formData.mailId — that
// is the BC/applicant's own email and is unrelated to this feature.

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

  // Fixed recipients first
  addIfValid(FIXED_RECIPIENT_1);
  addIfValid(FIXED_RECIPIENT_2);

  // Parse comma-separated officeMailId field
  if (officeMailIdFieldValue) {
    var parts = String(officeMailIdFieldValue).split(',');
    parts.forEach(function(p) { addIfValid(p); });
  }

  return result.join(',');
}

// ── generateAndEmailPdf_ ─────────────────────────────────────────
// Generates a PDF of the filled application form and emails it to the
// recipient list built by buildRecipientList_().
// This function has its own try/catch — any failure is logged but never
// propagated to the caller, so a PDF/email problem can never cause
// submitApplication() to return success:false or throw.

function generateAndEmailPdf_(formData, refId, siNo, subDT, isTest) {
  try {
    Logger.log('generateAndEmailPdf_: starting for refId=' + refId + ' isTest=' + isTest);

    // ── Step 1: Build recipient list ─────────────────────────────
    var recipients = buildRecipientList_(formData.officeMailId);
    Logger.log('generateAndEmailPdf_: recipients="' + recipients + '"');
    if (!recipients) {
      Logger.log('generateAndEmailPdf_: no valid recipients — skipping send');
      return;
    }

    // ── Step 2: Set logo URLs ─────────────────────────────────────
    Logger.log('generateAndEmailPdf_: using logo URLs');
    var logoSvf = LOGO_URL_SVF;
    var logoUco = LOGO_URL_UCO;
    Logger.log('generateAndEmailPdf_: logoSvf=' + logoSvf);

    // ── Step 3: Build PDF HTML inline — unified design system ───
    var d = formData;

    // HTML-escape a plain string value
    function esc(v) {
      return v ? String(v).trim()
        .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;') : '';
    }
    // Mask Aadhaar: show only last 4 digits
    function mask(v) {
      var s = v ? String(v).trim() : '';
      return s ? '\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022' + esc(s.slice(-4)) : '';
    }

    function section(title, pairs) {
      var rows = pairs.filter(function(p){ return p[1]; });
      if (!rows.length) return '';

      // Section Header Banner — Light green background (#e8f5e9) with 4px terracotta left border (#d9531e) and bold uppercase #1a5c1a text
      var h = '<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px;margin-bottom:0;background:#e8f5e9;border-left:4px solid #d9531e;border-radius:3px 3px 0 0;page-break-inside:avoid;break-inside:avoid;page-break-after:avoid;">' +
        '<tr><td style="padding:4px 10px;">' +
          '<span style="font-size:10px;font-weight:900;color:#1a5c1a;text-transform:uppercase;letter-spacing:.05em;">' +
          esc(title) + '</span>' +
        '</td></tr></table>';

      // Clean 2-column grid table with compact padding for single-page fit
      h += '<table width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #e2e8f0;margin-bottom:8px;page-break-inside:auto;">';
      for (var i = 0; i < rows.length; i += 2) {
        h += '<tr style="page-break-inside:avoid;break-inside:avoid;">';
        // Left cell
        h += '<td width="50%" style="border:1px solid #e2e8f0;padding:4px 8px;vertical-align:top;background:#ffffff;page-break-inside:avoid;break-inside:avoid;">' +
               '<div style="font-size:7.5px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:.04em;margin-bottom:1px;">' + esc(rows[i][0]) + '</div>' +
               '<div style="font-size:10.5px;font-weight:800;color:#0f172a;word-break:break-all;">' + rows[i][1] + '</div>' +
             '</td>';
        // Right cell (or empty cell)
        if (i + 1 < rows.length) {
          h += '<td width="50%" style="border:1px solid #e2e8f0;padding:4px 8px;vertical-align:top;background:#ffffff;page-break-inside:avoid;break-inside:avoid;">' +
                 '<div style="font-size:7.5px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:.04em;margin-bottom:1px;">' + esc(rows[i+1][0]) + '</div>' +
                 '<div style="font-size:10.5px;font-weight:800;color:#0f172a;word-break:break-all;">' + rows[i+1][1] + '</div>' +
               '</td>';
        } else {
          h += '<td width="50%" style="border:1px solid #e2e8f0;padding:4px 8px;background:#ffffff;">&nbsp;</td>';
        }
        h += '</tr>';
      }
      h += '</table>';
      return h;
    }

    // Logo image tag (uses fetched logoSvf if available, or fallback text)
    var flowerLogoHtml = logoSvf
      ? '<img src="' + logoSvf + '" alt="SVF Logo" style="width:55px;height:55px;display:block;margin:0 auto;">'
      : '<div style="font-size:16px;color:#1a5c1a;font-weight:bold;">SVF</div>';

    // Full Sanjivani Vikas Foundation logo for diagonal watermark
    var logoFullSvf = 'https://res.cloudinary.com/date69bba/image/upload/v1774602823/SanjivaniVikasLogo_new_2_hpwra4.png';

    // Diagonal logo watermark centered across PDF page
    var watermarkSrc = logoFullSvf || logoSvf;
    var watermarkHtml =
      '<div style="position:fixed;top:32%;left:5%;width:90%;text-align:center;transform:rotate(-25deg);opacity:0.08;pointer-events:none;z-index:0;">' +
        (watermarkSrc ? '<img src="' + watermarkSrc + '" style="max-width:440px;width:80%;height:auto;display:block;margin:0 auto;" />' : '') +
        (isTest ? '<div style="color:#d9531e;font-size:52px;font-weight:900;letter-spacing:0.1em;margin-top:8px;">TEST DATA</div>' : '') +
      '</div>';

    var html =
      '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
      '<style>' +
        '@page{margin:15px;}' +
        'body{font-family:Arial,Helvetica,sans-serif;font-size:10.5px;color:#0f172a;margin:0;padding:0;background:#fff;}' +
        'td,th,tr,div{box-sizing:border-box;}' +
        'tr,td{page-break-inside:avoid !important;break-inside:avoid !important;}' +
      '</style></head>' +
      '<body style="margin:0;padding:0;position:relative;">' +

      watermarkHtml +

      // ══════════════════════════════════════════════════════════════
      // HEADER — Primary Green Title (#1a5c1a) + Dark Grey Text (#475569)
      // ══════════════════════════════════════════════════════════════
      '<table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;margin-bottom:0;">' +
        '<tr>' +
          '<td width="75" style="padding:6px 6px;vertical-align:middle;text-align:center;">' +
            flowerLogoHtml +
          '</td>' +
          '<td style="padding:6px 10px;vertical-align:middle;text-align:center;">' +
            '<div style="font-size:22px;font-weight:900;letter-spacing:0.02em;color:#1a5c1a;margin-bottom:2px;">SANJIVANI VIKAS FOUNDATION</div>' +
            '<div style="font-size:8.5px;font-weight:bold;color:#475569;margin-bottom:1px;">Behind P.N.B. Mahatma Gandhi Nagar, Kankarbagh, Patna &ndash; 800026 (Bihar)</div>' +
            '<div style="font-size:8px;font-weight:bold;color:#475569;margin-bottom:1px;">Regd. under Societies Regn. Act, 21 of 1860 &bull; Regn. No. 497/2004-05</div>' +
            '<div style="font-size:8px;font-weight:bold;color:#475569;margin-bottom:1px;">Ph: 0612-2360350 / 9955999077 / 8969803111 / 9934011735 &bull; sanjivanivf@gmail.com &bull; sanjivani.foundation</div>' +
            '<div style="font-size:8px;font-weight:bold;color:#64748b;">ISO Certified &bull; PAN: AABAS9473R &bull; GSTIN: 10AABAS9473R1ZW</div>' +
          '</td>' +
        '</tr>' +
      '</table>' +

      // Document title strip with Terracotta Accent Border Line (#d9531e) replacing old gold rule
      '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1a5c1a;border-top:3px solid #d9531e;margin-bottom:10px;">' +
        '<tr>' +
          '<td style="padding:6px 12px;vertical-align:middle;">' +
            '<div style="font-size:12px;font-weight:bold;color:#ffffff;letter-spacing:.04em;text-transform:uppercase;">' +
              'BC / BCA APPOINTMENT APPLICATION &nbsp;&bull;&nbsp; UCO BANK' +
            '</div>' +
          '</td>' +
          '<td style="padding:6px 12px;vertical-align:middle;text-align:right;">' +
            '<span style="background:#e8f5e9;color:#1a5c1a;border:1px solid #a7f3d0;font-size:8.5px;font-weight:800;padding:2px 8px;border-radius:12px;letter-spacing:0.04em;text-transform:uppercase;">' +
              '&#10003; SUBMITTED' +
            '</span>' +
          '</td>' +
        '</tr>' +
      '</table>' +

      // ── Form content wrapper ──────────────────────────────────────
      '<div style="padding:0 10px 10px 10px;">' +

        // 4 Summary Cards styled as light bordered cards (#f8fafc bg, #e2e8f0 border)
        '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:10px;page-break-inside:avoid;break-inside:avoid;">' +
          '<tr>' +
            '<td width="22%" style="padding:2px;">' +
              '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:4px;padding:5px 8px;">' +
                '<div style="font-size:7px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.04em;margin-bottom:1px;">BANK</div>' +
                '<div style="font-size:11px;font-weight:900;color:#1a5c1a;">UCO Bank</div>' +
              '</div>' +
            '</td>' +
            '<td width="28%" style="padding:2px;">' +
              '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:4px;padding:5px 8px;">' +
                '<div style="font-size:7px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.04em;margin-bottom:1px;">REFERENCE ID</div>' +
                '<div style="font-size:11px;font-weight:900;color:#1a5c1a;letter-spacing:0.02em;">' + esc(refId) + '</div>' +
              '</div>' +
            '</td>' +
            '<td width="20%" style="padding:2px;">' +
              '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:4px;padding:5px 8px;">' +
                '<div style="font-size:7px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.04em;margin-bottom:1px;">SERIAL NO.</div>' +
                '<div style="font-size:11px;font-weight:900;color:#1a5c1a;">#' + esc(String(siNo)) + '</div>' +
              '</div>' +
            '</td>' +
            '<td width="30%" style="padding:2px;">' +
              '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:4px;padding:5px 8px;">' +
                '<div style="font-size:7px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.04em;margin-bottom:1px;">SUBMITTED ON</div>' +
                '<div style="font-size:10.5px;font-weight:900;color:#1a5c1a;">' + esc(subDT) + '</div>' +
              '</div>' +
            '</td>' +
          '</tr>' +
        '</table>' +

        section('DETAILS OF LOCATION (SSA / NON-SSA)', [
          ['State', esc(d.state)], ['Zone', esc(d.zone)], ['Link Branch', esc(d.branch)],
          ['Branch Code', esc(d.branchCode)], ['District', esc(d.district)],
          ['Block / Sub District', esc(d.block)], ['Village / Gram Panchayat', esc(d.villageName)], ['Village Code', esc(d.villageCode)]
        ]) +
        section('PERMANENT ADDRESS & BANK MITRA DETAILS', [
          ['CSP / Bank Mitra Name', esc(d.cspName)], ["Father's Name", esc(d.fathersName)],
          ['Mobile No.', esc(d.contactNumber)], ['Alt. Contact No.', esc(d.altContactNo)],
          ['Date of Birth', esc(d.dob)], ['Gender', esc(d.gender)],
          ['Permanent Address', esc(d.address)], ['PIN Code', esc(d.pinCode)], ['Caste Category', esc(d.caste)]
        ]) +
        section('KYC & IDENTIFICATION DETAILS', [
          ['Aadhaar No.', mask(d.aadhaarNo)], ['PAN No.', esc(d.panNo)],
          ['Other ID Type', esc(d.otherIdType)], ['Other ID No.', esc(d.otherIdNo)]
        ]) +
        section('BANK ACCOUNT DETAILS', [
          ['Agent CIF No.', esc(d.agentCifNo)], ['Settlement A/C No.', esc(d.settlementAccount)], ['Saving A/C No.', esc(d.savingAccount)]
        ]) +
        section('PROFESSIONAL & OTHER DETAILS', [
          ['Education Qualification', esc(d.education)], ['Date of Joining (DOJ)', esc(d.doj)],
          ['Upload Purpose', esc(d.uploadPurpose)], ['SHG Member', esc(d.shgMember)],
          ['Physically Challenged', esc(d.physicallyChallenged)],
          ['Network Provider', esc(d.networkProvider)], ['Email ID', esc(d.mailId)],
          ['Bank Mitra Activity', esc(d.bankMitraActivity)], ['Replaced Agent Name', esc(d.replacedAgent)]
        ]) +
        section('IIBF CERTIFICATION DETAILS', [
          ['IIBF Certified', esc(d.iibfCertified)],
          ['Certificate No.', esc(d.iibfCertificate)], ['Certificate Date', esc(d.certificateDate)]
        ]) +

        // Single clean page footer line (no duplicate stray elements)
        '<table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #e2e8f0;padding-top:6px;margin-top:10px;">' +
          '<tr>' +
            '<td style="font-size:8px;color:#64748b;">' +
              'Sanjivani Vikas Foundation &bull; BC/BCA Appointment Application &bull; UCO Bank &bull; Generated ' + esc(subDT) +
            '</td>' +
            '<td style="font-size:8px;color:#64748b;text-align:right;">' +
              'Page 1' +
            '</td>' +
          '</tr>' +
        '</table>' +

      '</div>' +
      '</body></html>';

    // ── Step 4: Convert HTML string to PDF blob (with graceful fallback) ──
    var safeName = (d.cspName ? String(d.cspName).trim() : 'Candidate').replace(/[^a-zA-Z0-9_\-]/g, '_');
    var safeZone = (d.zone ? String(d.zone).trim() : 'Zone').replace(/[^a-zA-Z0-9_\-]/g, '_');
    var pdfFilename = 'BC-Appointment-' + safeName + '-' + safeZone + '.pdf';
    var attachments = [];

    try {
      var htmlOutput = HtmlService.createHtmlOutput(html);
      var pdfBlob = htmlOutput.getAs('application/pdf').setName(pdfFilename);
      attachments.push(pdfBlob);
      Logger.log('generateAndEmailPdf_: PDF created successfully, size=' + pdfBlob.getBytes().length);
    } catch (pdfErr) {
      Logger.log('generateAndEmailPdf_: PDF creation skipped (' + pdfErr.message + '), sending HTML email directly');
    }

    // ── Step 5: Send email with Candidate, Branch, and Zone details ──
    var agentName  = d.cspName ? String(d.cspName).trim() : 'N/A';
    var branchName = d.branch ? String(d.branch).trim() : 'N/A';
    var zoneName   = d.zone ? String(d.zone).trim() : 'N/A';
    var distName   = d.district ? String(d.district).trim() : 'N/A';
    var contactNo  = d.contactNumber ? String(d.contactNumber).trim() : 'N/A';

    var emailSubject = 'BC/BCA Appointment Application — ' + agentName + ' (' + branchName + ', ' + zoneName + ') — ' + refId;

    var plainBody =
      'A new BC/BCA appointment application has been submitted.\n\n' +
      '• Candidate Name: ' + agentName + '\n' +
      '• Branch: ' + branchName + '\n' +
      '• Zone: ' + zoneName + '\n' +
      '• District: ' + distName + '\n' +
      '• Contact: ' + contactNo + '\n\n' +
      '• Reference ID: ' + refId + '\n' +
      '• Serial No.: ' + siNo + '\n' +
      '• Submitted On: ' + subDT;

    var htmlEmailBody =
      '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
      '<body style="margin:0;padding:20px;background-color:#f1f5f9;font-family:sans-serif;">' +
      '<div style="max-width:700px;margin:0 auto;background:#fff;padding:24px;border-radius:10px;border:1px solid #cbd5e1;">' +
      '<div style="background:#1a5c1a;color:#fff;padding:16px;text-align:center;border-radius:8px 8px 0 0;font-size:18px;font-weight:bold;">SANJIVANI VIKAS FOUNDATION — BC APPOINTMENT</div>' +
      '<div style="padding:20px;line-height:1.6;color:#0f172a;">' +
      '<h3>New Application Received</h3>' +
      '<ul>' +
      '<li><b>Candidate Name:</b> ' + esc(agentName) + '</li>' +
      '<li><b>Branch & Code:</b> ' + esc(branchName) + ' (' + esc(d.branchCode) + ')</li>' +
      '<li><b>Zone & State:</b> ' + esc(zoneName) + ' / ' + esc(d.state) + '</li>' +
      '<li><b>District:</b> ' + esc(distName) + '</li>' +
      '<li><b>Contact:</b> ' + esc(contactNo) + '</li>' +
      '<li><b>Reference ID:</b> ' + esc(refId) + '</li>' +
      '<li><b>Submitted On:</b> ' + esc(subDT) + '</li>' +
      '</ul>' +
      '</div></div></body></html>';

    Logger.log('generateAndEmailPdf_: sending email to ' + recipients);
    var mailSuccess = false;
    var mailErrDetail = '';

    try {
      MailApp.sendEmail({
        to: recipients,
        subject: emailSubject,
        body: plainBody,
        htmlBody: htmlEmailBody,
        name: 'SVF UCO BC Appointment System',
        attachments: attachments
      });
      mailSuccess = true;
      Logger.log('generateAndEmailPdf_: email sent successfully via MailApp');
    } catch (mErr) {
      Logger.log('generateAndEmailPdf_: MailApp failed: ' + mErr.message + ', trying GmailApp fallback...');
      try {
        GmailApp.sendEmail(recipients, emailSubject, plainBody, {
          htmlBody: htmlEmailBody,
          name: 'SVF UCO BC Appointment System',
          attachments: attachments
        });
        mailSuccess = true;
        Logger.log('generateAndEmailPdf_: email sent successfully via GmailApp fallback');
      } catch (gErr) {
        mailErrDetail = 'MailApp: ' + mErr.message + ' | GmailApp: ' + gErr.message;
        Logger.log('generateAndEmailPdf_: BOTH MailApp and GmailApp failed: ' + mailErrDetail);
      }
    }
    return { success: mailSuccess, error: mailErrDetail };

  } catch (err) {
    Logger.log('generateAndEmailPdf_ FAILED at step: ' + err.message + ' | stack: ' + err.stack);
    return { success: false, error: err.message };
  }
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

    // Get last SINo (if lastRow is 1, it's just headers, so next is 1. If lastRow is 2, next is 2)
    var lastRow = appointmentsSheet.getLastRow();
    var siNo    = lastRow > 0 ? lastRow : 1;

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

    // Generate and email the filled-form PDF. This runs inside the lock
    // (so the refId/siNo are stable), but the function's own try/catch
    // means any failure here can never affect the success:true return below.
    generateAndEmailPdf_(formData, refId, siNo, subDT, false);

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

// ── testEmailPdf ──────────────────────────────────────────────────
// Run this function directly from the Apps Script editor (Run > testEmailPdf)
// to test the email + PDF pipeline without needing to submit the form.
// IMPORTANT: set TEST_RECIPIENT below to a real inbox you can check.

function testEmailPdf() {
  var TEST_RECIPIENT = 'svf.samar@gmail.com'; // <-- change this before running

  var fakeFormData = {
    state: 'TEST STATE', zone: 'TEST ZONE', branch: 'TEST BRANCH',
    branchCode: 'BR001', district: 'TEST DISTRICT', block: 'TEST BLOCK',
    villageName: 'TEST VILLAGE', villageCode: '12345',
    cspName: 'TEST CSP', fathersName: 'TEST FATHER',
    contactNumber: '9876543210', altContactNo: '9876543211',
    dob: '1990-01-01', gender: 'Male',
    address: '123 TEST STREET', pinCode: '400001', caste: 'General',
    aadhaarNo: '123456789012', panNo: 'ABCDE1234F',
    otherIdType: 'Voter ID', otherIdNo: 'TEST123',
    agentCifNo: '987654', settlementAccount: '001234567890',
    savingAccount: '009876543210',
    education: 'Graduate', doj: '2024-01-15',
    uploadPurpose: 'NEW', shgMember: 'No',
    physicallyChallenged: 'No', networkProvider: 'Jio',
    mailId: 'testbc@example.com',
    bankMitraActivity: 'NA', replacedAgent: 'NA',
    iibfCertified: 'Yes', iibfCertificate: 'IIBF202600001',
    certificateDate: '2025-06-01',
    officeMailId: TEST_RECIPIENT  // the office email routes to your test inbox
  };

  Logger.log('testEmailPdf: starting');
  var currentSubDT = getSubmissionDateTime_();
  generateAndEmailPdf_(fakeFormData, 'SVF-UCO-TEST-000001', 1, currentSubDT, true);
  Logger.log('testEmailPdf: done — check Apps Script Logs and your inbox at ' + TEST_RECIPIENT);
}

function testMail_() {
  try {
    var recipients = buildRecipientList_('');
    var testSubject = 'Test Email from BC Appointment System — ' + new Date().toISOString();
    var testBody = 'This is a diagnostic test email from Sanjivani Vikas Foundation BC Appointment System.\n\nRecipients: ' + recipients;

    var mailSent = false;
    var errDetail = '';

    try {
      MailApp.sendEmail({
        to: recipients,
        subject: testSubject,
        body: testBody,
        name: 'SVF UCO BC Appointment System'
      });
      mailSent = true;
    } catch (e1) {
      try {
        GmailApp.sendEmail(recipients, testSubject, testBody, { name: 'SVF UCO BC Appointment System' });
        mailSent = true;
      } catch (e2) {
        errDetail = 'MailApp: ' + e1.message + ' | GmailApp: ' + e2.message;
      }
    }

    return {
      success: mailSent,
      recipients: recipients,
      error: errDetail
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}