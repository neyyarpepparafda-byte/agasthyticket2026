// ====== Admin Credentials (use Script Properties for security) ======
function getAdminCredentials() {
  const props = PropertiesService.getScriptProperties();
  return {
    username: props.getProperty('ADMIN_USERNAME') || 'admin',
    password: props.getProperty('ADMIN_PASSWORD') || 'password123'
  };
}

function checkAdminCredentials(username, password) {
  const creds = getAdminCredentials();
  return username === creds.username && password === creds.password;
}

// ====== Config Helpers ======
function getConfigValue(key) {
  try {
    key = (key || '').toString().trim();
    if (!key) return '';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = ss.getSheetByName('Config');
    if (!config) return '';
    const data = config.getRange(1, 1, config.getLastRow(), 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if ((data[i][0] + '').toString().trim().toLowerCase() === key.toLowerCase()) {
        return (data[i][1] || '').toString();
      }
    }
    return '';
  } catch (e) {
    return '';
  }
}

function setConfigValue(key, value) {
  try {
    key = (key || '').toString().trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let config = ss.getSheetByName('Config');
    if (!config) config = ss.insertSheet('Config');
    const lastRow = Math.max(1, config.getLastRow());
    const data = config.getRange(1, 1, lastRow, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if ((data[i][0] + '').toString().trim().toLowerCase() === key.toLowerCase()) {
        config.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
    config.appendRow([key, value]);
  } catch (e) {
    // ignore
  }
}

// ====== Utility Functions ======
function formatDateDDMMYYYY(date) {
  if (!date) return "";
  if (typeof date === "string" && date.includes("/")) return date;
  if (typeof date === "string" && date.includes("-")) date = new Date(date);
  if (Object.prototype.toString.call(date) === "[object Date]" && !isNaN(date)) {
    const dd = String(date.getDate()).padStart(2, "0");
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const yyyy = date.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }
  return "";
}

function formatDateDDMMYY(date) {
  if (!date) return "";
  if (typeof date === "string" && date.includes("/")) {
    const [dd, mm, yyyy] = date.split("/");
    return `${dd}${mm}${yyyy.slice(-2)}`;
  }
  if (typeof date === "string" && date.includes("-")) date = new Date(date);
  if (Object.prototype.toString.call(date) === "[object Date]" && !isNaN(date)) {
    const dd = String(date.getDate()).padStart(2, "0");
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const yy = String(date.getFullYear()).slice(-2);
    return `${dd}${mm}${yy}`;
  }
  return "";
}

function getColIdx(sheet, header) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(header) + 1;
}

function ensureColumn(sheet, header) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!headers.includes(header)) {
    sheet.insertColumnAfter(headers.length);
    sheet.getRange(1, headers.length + 1).setValue(header);
  }
}

// Robust date normalization: accepts Date objects, dd/mm/yyyy, dd-mm-yyyy, yyyy-mm-dd, and strings with time
function normalizeToDDMMYYYY(input) {
  if (input === null || input === undefined || input === '') return '';
  try {
    // If Date object
    if (Object.prototype.toString.call(input) === "[object Date]" && !isNaN(input)) {
      return formatDateDDMMYYYY(input);
    }

    // If it's a number or numeric string, allow it to be handled by Date constructor as fallback
    var s = ('' + input).toString().trim();
    if (!s) return '';

    // Remove time portion if present
    var datePart = s.split(' ')[0];

    // If contains '/'
    if (datePart.indexOf('/') !== -1) {
      var parts = datePart.split('/');
      if (parts.length === 3) {
        var dd = Number(parts[0]), mm = Number(parts[1]), yyyy = Number(parts[2]);
        if (!isNaN(dd) && !isNaN(mm) && !isNaN(yyyy)) {
          return `${String(dd).padStart(2,'0')}/${String(mm).padStart(2,'0')}/${yyyy}`;
        }
      }
    }

    // If contains '-'
    if (datePart.indexOf('-') !== -1) {
      var dashParts = datePart.split('-');
      if (dashParts.length === 3) {
        // If first part is 4 chars -> assume YYYY-MM-DD
        if (dashParts[0].length === 4) {
          var yyyy2 = Number(dashParts[0]), mm2 = Number(dashParts[1]), dd2 = Number(dashParts[2]);
          if (!isNaN(dd2) && !isNaN(mm2) && !isNaN(yyyy2)) {
            return `${String(dd2).padStart(2,'0')}/${String(mm2).padStart(2,'0')}/${yyyy2}`;
          }
        }
        // If last part is 4 chars -> assume DD-MM-YYYY
        if (dashParts[2].length === 4) {
          var dd3 = Number(dashParts[0]), mm3 = Number(dashParts[1]), yyyy3 = Number(dashParts[2]);
          if (!isNaN(dd3) && !isNaN(mm3) && !isNaN(yyyy3)) {
            return `${String(dd3).padStart(2,'0')}/${String(mm3).padStart(2,'0')}/${yyyy3}`;
          }
        }
      }
    }

    // Final fallback: attempt Date parsing
    var parsed = new Date(s);
    if (!isNaN(parsed)) return formatDateDDMMYYYY(parsed);
  } catch (e) {
    // ignore
  }
  return '';
}

// Convert numbers (or numeric cells) to plain full-digit strings (no exponential)
function toPlainNumberString(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'string') return val;
  if (typeof val === 'number') {
    try {
      return Utilities.formatString('%.0f', val);
    } catch (e) {
      return String(Math.trunc(val));
    }
  }
  return String(val);
}

// ====== Logo Helpers ======
function getLogoFileIdFromConfig() {
  try {
    return (getConfigValue('LogoFileId') || '').toString().trim();
  } catch (e) {
    return '';
  }
}

function getLogoDataUriFromConfig() {
  try {
    return (getConfigValue('LogoDataUri') || '').toString().trim();
  } catch (e) {
    return '';
  }
}

function generateAndStoreLogoDataUri(fileId) {
  try {
    if (!fileId) throw new Error('Missing fileId');
    const file = DriveApp.getFileById(fileId);
    if (!file) throw new Error('Drive file not found');
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const contentType = blob.getContentType() || 'image/png';
    const dataUrl = 'data:' + contentType + ';base64,' + base64;
    setConfigValue('LogoDataUri', dataUrl);
    return dataUrl;
  } catch (e) {
    return '';
  }
}

// ====== Receipts Folder Helper ======
function getReceiptsFolder() {
  try {
    var folderId = '';
    if (typeof getConfigValue === 'function') {
      folderId = (getConfigValue('ReceiptsFolderId') || '').toString().trim();
    }
    if (!folderId) {
      var props = PropertiesService.getScriptProperties();
      folderId = (props.getProperty('RECEIPTS_FOLDER_ID') || '').toString().trim();
    }
    if (folderId) {
      try {
        return DriveApp.getFolderById(folderId);
      } catch (e) {
        // invalid id, fall through to create
      }
    }
    var folderName = 'Agasthyarkoodam-Receipts';
    var folder = DriveApp.createFolder(folderName);
    var newId = folder.getId();
    try {
      if (typeof setConfigValue === 'function') {
        setConfigValue('ReceiptsFolderId', newId);
      } else {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var conf = ss.getSheetByName('Config');
        if (!conf) conf = ss.insertSheet('Config');
        conf.appendRow(['ReceiptsFolderId', newId]);
      }
    } catch (e) {
      // ignore persist error
    }
    try {
      PropertiesService.getScriptProperties().setProperty('RECEIPTS_FOLDER_ID', newId);
    } catch (e) {}
    return folder;
  } catch (e) {
    return DriveApp.getRootFolder();
  }
}

// ====== Email Authorization ======
function authorizeEmail(email, maxTrekkers, allowedDate) {
  try {
    email = (email || '').trim().toLowerCase();
    if (!email) return { success: false, message: 'Invalid email.' };
    maxTrekkers = Number(maxTrekkers) || 10;
    if (maxTrekkers < 1) maxTrekkers = 1;
    if (maxTrekkers > 10) maxTrekkers = 10;

    const allowedDateNormalized = normalizeToDDMMYYYY(allowedDate);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('AuthorizedEmails');
    if (!sheet) {
      sheet = ss.insertSheet('AuthorizedEmails');
      sheet.getRange(1,1,1,3).setValues([['Authorized Email','MaxTrekkers','AllowedDate']]);
    }

    const headers = sheet.getRange(1, 1, 1, Math.max(3, sheet.getLastColumn())).getValues()[0];
    if (!headers.includes('Authorized Email')) sheet.getRange(1,1).setValue('Authorized Email');
    if (!headers.includes('MaxTrekkers')) sheet.getRange(1,2).setValue('MaxTrekkers');
    if (!headers.includes('AllowedDate')) {
      if (sheet.getLastColumn() < 3) {
        sheet.getRange(1,3).setValue('AllowedDate');
      } else {
        sheet.insertColumnAfter(2);
        sheet.getRange(1,3).setValue('AllowedDate');
      }
    }

    const lastRow = sheet.getLastRow();
    const rowsCount = Math.max(0, lastRow - 1);
    const existingRows = rowsCount > 0 ? sheet.getRange(2, 1, rowsCount, Math.max(3, sheet.getLastColumn())).getValues() : [];
    const existingLower = existingRows.map(r => (r[0] + '').toLowerCase());
    if (existingLower.includes(email)) {
      const idx = existingLower.indexOf(email);
      const rowNum = idx + 2;
      sheet.getRange(rowNum, 2).setValue(maxTrekkers);
      sheet.getRange(rowNum, 3).setValue(allowedDateNormalized);
      return { success: true, message: `${email} already authorized; updated MaxTrekkers to ${maxTrekkers} and AllowedDate.` };
    }

    if (sheet.getLastColumn() < 3) {
      sheet.getRange(1,1,1,1).setValue('Authorized Email');
      if (sheet.getLastColumn() < 2) {
        sheet.insertColumnAfter(1);
        sheet.getRange(1,2).setValue('MaxTrekkers');
      }
      sheet.insertColumnAfter(2);
      sheet.getRange(1,3).setValue('AllowedDate');
    }
    sheet.appendRow([email, maxTrekkers, allowedDateNormalized]);

    let webAppUrl = '';
    try {
      webAppUrl = (getConfigValue('WebAppUrl') || '').toString().trim();
    } catch (e) {}
    if (!webAppUrl) {
      const props = PropertiesService.getScriptProperties();
      webAppUrl = props.getProperty('WEB_APP_URL') || '';
    }

    if (webAppUrl) {
      const sep = webAppUrl.indexOf('?') === -1 ? '?' : '&';
      webAppUrl = `${webAppUrl}${sep}role=user&email=${encodeURIComponent(email)}`;
    }

    const bodyHtml = [
      `<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px;">`,
      `<h2 style="color: #2E7D32;">Agasthyarkoodam Trekking Registration</h2>`,
      `<p>Your group email (<strong>${email}</strong>) has been authorized for registration.</p>`,
      `<p>Maximum allowed trekkers for your group: <strong>${maxTrekkers}</strong>.</p>`,
      allowedDateNormalized ? `<p>Assigned trek date: <strong>${allowedDateNormalized}</strong></p>` : '',
      `<hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">`,
      webAppUrl ? `<p><strong>ശ്രദ്ധിക്കുക:</strong> ലിങ്ക് തുറക്കുന്നതിന് മുൻപ് ചുവടെ പറയുന്നതിന്റ[narrative truncated in mail for brevity]</p>` : '',
      `<div style="background-color: #f1f8e9; border: 1px solid #c5e1a5; padding: 15px; border-radius: 8px; margin: 20px 0;">`,
      `<h3 style="margin-top: 0; color: #2E7D32; border-bottom: 1px solid #a5d6a7; padding-bottom: 5px;">Bank Account Details</h3>`,
      `<table style="width: 100%; border-collapse: collapse;">`,
      `<tr><td style="padding: 5px 0; width: 140px; font-weight: bold;">Account No:</td><td style="padding: 5px 0; font-family: monospace; font-size: 16px;">0503073000000828</td></tr>`,
      `<tr><td style="padding: 5px 0; font-weight: bold;">IFSC Code:</td><td style="padding: 5px 0; font-family: monospace; font-size: 16px;">SIBL0000503</td></tr>`,
      `<tr><td style="padding: 5px 0; font-weight: bold;">Account Name:</td><td style="padding: 5px 0;">AGASTHYADHANAM</td></tr>`,
      `<tr><td style="padding: 5px 0; font-weight: bold;">Branch:</td><td style="padding: 5px 0;">SIB Sasthamangalam Branch</td></tr>`,
      `<tr><td style="padding: 5px 0; font-weight: bold;">Amount:</td><td style="padding: 5px 0;">₹5,000 per person (Local), ₹10,000 per person (Foreigner)</td></tr>`,
      `</table>`,
      `</div>`,
      webAppUrl ? `<div style="text-align:center; margin:20px 0;"><a href="${webAppUrl}" style="background-color:#2E7D32;color:white;padding:12px 25px;text-decoration:none;border-radius:5px;font-weight:bold;">Open Registration</a></div>` : '',
      `<br><p style="color:#555;">Regards,<br><strong>Thiruvananthapuram Wildlife Division</strong></p>`,
      `</div>`
    ].join('');

    MailApp.sendEmail({
      to: email,
      subject: "Agasthyarkoodam Trekking Registration - Access Link",
      htmlBody: bodyHtml
    });

    return { success: true, message: `Email ${email} authorized with max ${maxTrekkers}.` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ====== User Authorization & Status ======
function userLoginStatus(email) {
  email = (email || '').trim().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const authSheet = ss.getSheetByName('AuthorizedEmails');
  let allowedCount = 10;
  let allowedDate = '';

  if (!authSheet) {
    return { status: "unauthorized", message: "Email not authorized. Please contact the office." };
  }

  const last = authSheet.getLastRow();
  if (last > 1) {
    const lastCol = Math.max(1, authSheet.getLastColumn());
    const vals = authSheet.getRange(2, 1, last - 1, lastCol).getValues();
    const found = vals.find(r => (r[0] + '').toLowerCase() === email);

    if (!found) {
      return { status: "unauthorized", message: "Email not authorized. Please contact the office." };
    }

    if (found.length >= 2 && found[1] !== '' && !isNaN(Number(found[1]))) {
      allowedCount = Math.max(1, Math.min(10, Number(found[1])));
    } else {
      allowedCount = 10;
    }

    if (found.length >= 3 && (found[2] || '').toString().trim()) {
      allowedDate = normalizeToDDMMYYYY(found[2].toString().trim());
    } else {
      allowedDate = '';
    }
  } else {
    return { status: "unauthorized", message: "Email not authorized. Please contact the office." };
  }

  // Check registrations
  const regSheet = ss.getSheetByName('Registrations');
  if (!regSheet) {
    return { status: "not_registered", message: "Authorized. Please register.", allowedCount: allowedCount, allowedDate: allowedDate };
  }

  const regRows = Math.max(0, regSheet.getLastRow() - 1);
  if (regRows > 0) {
    const headers = regSheet.getRange(1, 1, 1, regSheet.getLastColumn()).getValues()[0];
    const data = regSheet.getRange(2, 1, regRows, regSheet.getLastColumn()).getValues();

    let ticketIssuedFound = false;
    let pendingFound = false;
    let pendingTrekDate = '';

    for (let i = 0; i < data.length; i++) {
      const rowEmail = (data[i][headers.indexOf('Email')] + '').toLowerCase();

      if (rowEmail === email) {
        const status = (data[i][headers.indexOf('Status')] + '').trim();
        const trekDate = data[i][headers.indexOf('Trek Date')];
        const trekDateDMY = formatDateDDMMYYYY(trekDate);

        // Check for Ticket Issued
        if (status === "Ticket Issued") {
          ticketIssuedFound = true;
          return {
            status: "ticket_issued",
            message: "Your Ticket is already issued, Please check your Authorized email",
            trekDate: trekDateDMY,
            allowedCount: allowedCount,
            allowedDate: allowedDate
          };
        }

        // Track pending status
        if (status === "Pending" && !pendingFound) {
          pendingFound = true;
          pendingTrekDate = trekDateDMY;
        }
      }
    }

    // Return pending status if found
    if (pendingFound) {
      return {
        status: "processing",
        message: "Your Application is under processing, Ticket will be issued soon",
        trekDate: pendingTrekDate,
        allowedCount: allowedCount,
        allowedDate: allowedDate
      };
    }
  }

  // No registration found
  return { status: "not_registered", message: "Authorized. Please register.", allowedCount: allowedCount, allowedDate: allowedDate };
}

// ====== Submit Registration with Receipt Upload ======
function submitRegistration(email, trekDate, amount, paymentId, trekkers, receiptObj, submissionId, emergencyContact) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server busy. Please try again in a moment.' };
  }

  try {
    email = (email || '').toString().trim().toLowerCase();
    if (!email) return { success: false, message: 'Invalid email.' };

    // Emergency contact is required (group-level)
    emergencyContact = (emergencyContact || '').toString().trim();
    var normalizedEmergency = emergencyContact.replace(/\s+/g, '');
    var digitsOnly = normalizedEmergency.replace(/[^\d]/g, '');
    if (!emergencyContact || digitsOnly.length < 7) {
      return { success: false, message: 'Emergency contact is required and must be a valid phone number.' };
    }

    if (!Array.isArray(trekkers) || trekkers.length < 1 || trekkers.length > 10) {
      return { success: false, message: 'Number of trekkers must be between 1 and 10.' };
    }

    for (var i = 0; i < trekkers.length; i++) {
      var ageNum = Number(trekkers[i].age || trekkers[i].Age || 0);
      if (isNaN(ageNum) || ageNum < 14) {
        return { success: false, message: 'Trekker ' + (i + 1) + ' must be at least 14 years old.' };
      }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Registrations');
    if (!sheet) {
      sheet = ss.insertSheet('Registrations');
      sheet.getRange(1,1,1,18).setValues([[
        'Email','Trek Date','Amount Remitted','Payment ID','No. of Trekkers','Status','Timestamp',
        'Trekker Name','Gender','Age','ID Type','ID Number','Mobile Number','Emergency Contact','Is Foreigner','Ticket No','SubmissionId','PaymentReceiptFileId'
      ]]);
    } else {
      var headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
      if (headers.indexOf('SubmissionId') === -1) {
        sheet.insertColumnAfter(headers.length);
        sheet.getRange(1, headers.length + 1).setValue('SubmissionId');
        headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
      }
      if (headers.indexOf('PaymentReceiptFileId') === -1) {
        sheet.insertColumnAfter(sheet.getLastColumn());
        sheet.getRange(1, sheet.getLastColumn()).setValue('PaymentReceiptFileId');
      }
      headers = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getValues()[0];
      if (headers.indexOf('Emergency Contact') === -1) {
        sheet.insertColumnAfter(headers.length);
        sheet.getRange(1, headers.length + 1).setValue('Emergency Contact');
      }
      headers = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getValues()[0];
      if (headers.indexOf('Is Foreigner') === -1) {
        sheet.insertColumnAfter(headers.length);
        sheet.getRange(1, headers.length + 1).setValue('Is Foreigner');
      }
    }

    var data = sheet.getDataRange().getValues();
    var headerRow = data[0] || [];
    var emailIdx = headerRow.indexOf('Email');
    var trekDateIdx = headerRow.indexOf('Trek Date');
    var submissionIdIdx = headerRow.indexOf('SubmissionId');

    var trekDateDMY = formatDateDDMMYYYY(trekDate);

    if (submissionId && submissionIdIdx !== -1) {
      for (var r = 1; r < data.length; r++) {
        if ((data[r][submissionIdIdx] || '') === submissionId) {
          return { success: true, message: 'Duplicate submission ignored (already processed).' };
        }
      }
    }

    for (var j = 1; j < data.length; j++) {
      var rowEmail = (data[j][emailIdx] || '').toString().toLowerCase();
      var rowDate = formatDateDDMMYYYY(data[j][trekDateIdx]);
      if (rowEmail === email && rowDate === trekDateDMY) {
        return { success: false, message: 'This email is already registered for the selected trek date.' };
      }
    }

    var localFee = Number(getFeePerPerson()) || 0;
    // read foreigner fee from Config; fallback to localFee
    var foreignFee = Number(getConfigValue('FeePerPersonForeigner')) || localFee;
    var expected = 0;
    for (var ti = 0; ti < trekkers.length; ti++) {
      var t = trekkers[ti];
      var isF = !!t.isForeigner;
      expected += isF ? foreignFee : localFee;
    }

    if (Number(amount) !== expected) {
      return { success: false, message: 'Amount mismatch. Expected ' + expected + ' for provided trekkers and categories.' };
    }

    var receiptFileId = '';
    if (receiptObj && receiptObj.base64) {
      try {
        var parentFolder = getReceiptsFolder();
        var blob = Utilities.newBlob(Utilities.base64Decode(receiptObj.base64), receiptObj.mime || 'application/octet-stream', receiptObj.name || ('receipt-' + Date.now()));
        var file = parentFolder.createFile(blob);
        receiptFileId = file.getId();
      } catch (saveErr) {
        return { success: false, message: 'Failed to save receipt file. Please try again.' };
      }
    } else {
      return { success: false, message: 'Payment receipt is required.' };
    }

    var now = new Date();
    headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var receiptColIdx = headerRow.indexOf('PaymentReceiptFileId');
    var submissionColIdx = headerRow.indexOf('SubmissionId');
    var emergencyColIdx = headerRow.indexOf('Emergency Contact');
    var isForeignerColIdx = headerRow.indexOf('Is Foreigner');

    for (var k = 0; k < trekkers.length; k++) {
      var t = trekkers[k];
      var row = [];
      row[headerRow.indexOf('Email')] = email;
      row[headerRow.indexOf('Trek Date')] = trekDateDMY;
      row[headerRow.indexOf('Amount Remitted')] = amount;
      row[headerRow.indexOf('Payment ID')] = paymentId;
      row[headerRow.indexOf('No. of Trekkers')] = trekkers.length;
      row[headerRow.indexOf('Status')] = 'Pending';
      row[headerRow.indexOf('Timestamp')] = now;
      row[headerRow.indexOf('Trekker Name')] = t.name;
      row[headerRow.indexOf('Gender')] = t.gender;
      row[headerRow.indexOf('Age')] = t.age;
      row[headerRow.indexOf('ID Type')] = t.idType;
      row[headerRow.indexOf('ID Number')] = t.idNumber;
      row[headerRow.indexOf('Mobile Number')] = t.mobile;
      if (emergencyColIdx !== -1) row[emergencyColIdx] = emergencyContact;
      if (isForeignerColIdx !== -1) row[isForeignerColIdx] = t.isForeigner ? 'Yes' : 'No';
      if (submissionColIdx !== -1) row[submissionColIdx] = submissionId || '';
      if (receiptColIdx !== -1) row[receiptColIdx] = receiptFileId || '';
      while (row.length < headerRow.length) row.push('');
      sheet.appendRow(row);
    }

    return { success: true, message: 'Registration submitted; payment receipt uploaded. Admin will verify.' };

  } catch (err) {
    return { success: false, message: (err && err.message) ? err.message : 'Unknown error' };
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }
}

// ====== Admin Dashboard: List Registrations ======
function getAllRegistrations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Registrations');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return [];

  const headers = data[0].map(h => (h || '').toString());
  const emailIdx = headers.indexOf('Email');
  const trekDateIdx = headers.indexOf('Trek Date');
  const statusIdx = headers.indexOf('Status');
  const timestampIdx = headers.indexOf('Timestamp');

  if (emailIdx === -1 || trekDateIdx === -1) {
    return [];
  }

  const groupMap = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = (row[emailIdx] || '').toString().trim();
    if (!email) continue;
    const trekDateRaw = row[trekDateIdx];
    const trekDateDMY = formatDateDDMMYYYY(trekDateRaw);
    const status = ((statusIdx !== -1 && row[statusIdx]) || '').toString().trim();
    const tsVal = (timestampIdx !== -1 && row[timestampIdx]) ? row[timestampIdx] : '';
    let tsDate = null;
    if (tsVal instanceof Date && !isNaN(tsVal)) tsDate = tsVal;
    else if (typeof tsVal === 'string' && tsVal.trim()) {
      const s = tsVal.toString().trim();
      if (s.match(/^\d{1,2}\/\d{1,2}\/\d{4}/)) {
        const parts = s.split(' ')[0].split('/');
        const dd = Number(parts[0]), mm = Number(parts[1]), yyyy = Number(parts[2]);
        tsDate = new Date(yyyy, mm - 1, dd);
      } else {
        const parsed = new Date(s);
        if (!isNaN(parsed)) tsDate = parsed;
      }
    }

    const key = email.toLowerCase() + '||' + trekDateDMY;
    if (!groupMap[key]) {
      groupMap[key] = {
        Email: email,
        'Trek Date': trekDateDMY,
        'No. of Trekkers': 0,
        Status: status || 'Pending',
        _timestamps: []
      };
    }
    groupMap[key]['No. of Trekkers']++;
    if ((status || '').toString().toLowerCase() === 'ticket issued') {
      groupMap[key].Status = 'Ticket Issued';
    }
    if (tsDate) groupMap[key]._timestamps.push(tsDate);
  }

  let groups = Object.values(groupMap);

  groups.forEach(g => {
    if (g._timestamps && g._timestamps.length > 0) {
      g._timestamps.sort((a, b) => a - b);
      g._regTime = g._timestamps[0];
    } else {
      g._regTime = null;
    }
    delete g._timestamps;
  });

  groups.sort((a, b) => {
    const at = a._regTime ? a._regTime.getTime() : 0;
    const bt = b._regTime ? b._regTime.getTime() : 0;
    return at - bt;
  });

  for (let i = 0; i < groups.length; i++) {
    groups[i].SerialNo = i + 1;
    delete groups[i]._regTime;
  }

  return groups.reverse();
}

// ====== Admin: View Group Details ======
function getGroupDetails(email, date) {
  email = (email || '').toString().trim().toLowerCase();
  if (!email) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Registrations');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return null;
  const headers = data[0].map(h => (h || '').toString());
  const emailIdx = headers.findIndex(h => (h || '').toLowerCase() === 'email');
  const trekDateIdx = headers.findIndex(h => (h || '').toLowerCase() === 'trek date' || (h || '').toLowerCase() === 'trekdate');
  const receiptIdx = headers.findIndex(h => (h || '').toLowerCase() === 'paymentreceiptfileid' || (h || '').toLowerCase() === 'payment receipt fileid' || (h || '').toLowerCase() === 'payment receip');

  const trekDateDMY = formatDateDDMMYYYY(date);

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const rowEmail = (data[i][emailIdx] || '').toString().toLowerCase();
    const rowDate = formatDateDDMMYYYY(data[i][trekDateIdx]);
    if (rowEmail === email && rowDate === trekDateDMY) {
      rows.push(data[i]);
    }
  }
  if (rows.length === 0) return null;

  const firstRow = rows[0];
  const groupDetails = {
    Email: firstRow[emailIdx],
    'Trek Date': trekDateDMY,
    'Amount Remitted': firstRow[headers.findIndex(h => (h || '').toLowerCase() === 'amount remitted' || (h || '').toLowerCase() === 'amount')],
    'Payment ID': toPlainNumberString(firstRow[headers.findIndex(h => (h || '').toLowerCase() === 'payment id' || (h || '').toLowerCase() === 'paymentid')]) || '',
    'Ticket No': firstRow[headers.findIndex(h => (h || '').toLowerCase() === 'ticket no' || (h || '').toLowerCase() === 'ticketno')] || '',
    Status: firstRow[headers.findIndex(h => (h || '').toLowerCase() === 'status')] || '',
    PaymentReceiptFileId: receiptIdx !== -1 ? (firstRow[receiptIdx] || '') : ''
  };

  if (groupDetails.PaymentReceiptFileId) {
    const fid = groupDetails.PaymentReceiptFileId.toString().trim();
    groupDetails.ReceiptDownloadUrl = `https://drive.google.com/uc?export=download&id=${fid}`;
    groupDetails.ReceiptViewUrl = `https://drive.google.com/file/d/${fid}/view`;
  } else {
    groupDetails.ReceiptDownloadUrl = '';
    groupDetails.ReceiptViewUrl = '';
  }

  const emergencyIdx = headers.findIndex(h => (h || '').toLowerCase() === 'emergency contact');
  groupDetails['Emergency Contact'] = emergencyIdx !== -1 ? (firstRow[emergencyIdx] || '') : '';

  const nameIdx = headers.findIndex(h => (h || '').toLowerCase() === 'trekker name' || (h || '').toLowerCase() === 'name');
  const genderIdx = headers.findIndex(h => (h || '').toLowerCase() === 'gender');
  const ageIdx = headers.findIndex(h => (h || '').toLowerCase() === 'age');
  const idTypeIdx = headers.findIndex(h => (h || '').toLowerCase() === 'id type' || (h || '').toLowerCase() === 'photo id type');
  const idNumberIdx = headers.findIndex(h => (h || '').toLowerCase() === 'id number' || (h || '').toLowerCase() === 'idnumber');
  const mobileIdx = headers.findIndex(h => (h || '').toLowerCase() === 'mobile number' || (h || '').toLowerCase() === 'mobile');
  const isForeignerIdx = headers.findIndex(h => (h || '').toLowerCase() === 'is foreigner');

  groupDetails.Trekkers = rows.map(row => ({
    Name: nameIdx !== -1 ? row[nameIdx] : '',
    Gender: genderIdx !== -1 ? row[genderIdx] : '',
    Age: ageIdx !== -1 ? row[ageIdx] : '',
    'Photo ID Type': idTypeIdx !== -1 ? row[idTypeIdx] : '',
    'ID Number': idNumberIdx !== -1 ? toPlainNumberString(row[idNumberIdx]) : '',
    'Mobile Number': mobileIdx !== -1 ? toPlainNumberString(row[mobileIdx]) : '',
    'Is Foreigner': isForeignerIdx !== -1 ? ( (row[isForeignerIdx] || '').toString().trim() === 'Yes' ) : false
  }));

  return groupDetails;
}

// ====== Admin: Edit Trekker Details ======
function updateTrekkerDetails(email, trekDate, index, editedData) {
  try {
    email = (email || '').trim().toLowerCase();
    if (!email) return { success: false, message: 'Invalid email.' };
    const trekDateDMY = formatDateDDMMYYYY(trekDate);
    if (!trekDateDMY) return { success: false, message: 'Invalid trek date.' };

    if (editedData && editedData.Age !== undefined && editedData.Age !== null && editedData.Age !== '') {
      const ageNum = Number(editedData.Age);
      if (isNaN(ageNum) || ageNum < 14) {
        return { success: false, message: 'Age must be a number and at least 14 years.' };
      }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Registrations');
    if (!sheet) return { success: false, message: 'Registrations sheet not found.' };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: 'No registration data.' };
    const headers = data[0];

    const emailIdx = headers.indexOf('Email');
    const trekDateIdx = headers.indexOf('Trek Date');
    const nameIdx = headers.indexOf('Trekker Name');
    const genderIdx = headers.indexOf('Gender');
    const ageIdx = headers.indexOf('Age');
    const idTypeIdx = headers.indexOf('ID Type');
    const idNumberIdx = headers.indexOf('ID Number');
    const mobileIdx = headers.indexOf('Mobile Number');

    if (emailIdx === -1 || trekDateIdx === -1) {
      return { success: false, message: 'Required columns not found.' };
    }

    const groupRowNums = [];
    for (let i = 1; i < data.length; i++) {
      if ((data[i][emailIdx] + '').toLowerCase() === email &&
          formatDateDDMMYYYY(data[i][trekDateIdx]) === trekDateDMY) {
        groupRowNums.push(i);
      }
    }

    if (groupRowNums.length === 0) {
      return { success: false, message: 'Group not found for the provided email and trek date.' };
    }
    if (index < 0 || index >= groupRowNums.length) {
      return { success: false, message: 'Invalid trekker index.' };
    }

    const rowIdx = groupRowNums[index];
    const rowNumber = rowIdx + 1;

    if (nameIdx !== -1) sheet.getRange(rowNumber, nameIdx + 1).setValue(editedData.Name || '');
    if (genderIdx !== -1) sheet.getRange(rowNumber, genderIdx + 1).setValue(editedData.Gender || '');
    if (ageIdx !== -1) sheet.getRange(rowNumber, ageIdx + 1).setValue(editedData.Age || '');
    if (idTypeIdx !== -1) sheet.getRange(rowNumber, idTypeIdx + 1).setValue(editedData['Photo ID Type'] || '');
    if (idNumberIdx !== -1) sheet.getRange(rowNumber, idNumberIdx + 1).setValue(editedData['ID Number'] || '');
    if (mobileIdx !== -1) sheet.getRange(rowNumber, mobileIdx + 1).setValue(editedData['Mobile Number'] || '');

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ====== Admin: Update Payment ID ======
function updatePaymentId(email, trekDate, paymentId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server busy. Try again.' };
  }
  try {
    email = (email || '').toString().trim().toLowerCase();
    if (!email) return { success: false, message: 'Invalid email.' };
    var dateDMY = formatDateDDMMYYYY(trekDate);
    if (!dateDMY) return { success: false, message: 'Invalid trek date.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Registrations');
    if (!sheet) return { success: false, message: 'Registrations sheet not found.' };

    var data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return { success: false, message: 'No registration data.' };

    var headers = data[0].map(h => (h || '').toString());
    var emailIdx = headers.findIndex(h => (h || '').toLowerCase() === 'email');
    var trekDateIdx = headers.findIndex(h => (h || '').toLowerCase() === 'trek date' || (h || '').toLowerCase() === 'trekdate');
    var paymentIdIdx = headers.findIndex(h => (h || '').toLowerCase() === 'payment id' || (h || '').toLowerCase() === 'paymentid');

    if (emailIdx === -1 || trekDateIdx === -1 || paymentIdIdx === -1) {
      return { success: false, message: 'Required columns not found.' };
    }

    var updated = false;
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][emailIdx] || '').toString().toLowerCase();
      var rowDate = formatDateDDMMYYYY(data[i][trekDateIdx]);
      if (rowEmail === email && rowDate === dateDMY) {
        sheet.getRange(i + 1, paymentIdIdx + 1).setValue(paymentId || '');
        updated = true;
      }
    }

    if (!updated) return { success: false, message: 'Group not found for the provided email and trek date.' };
    return { success: true, message: 'Payment ID updated.' };
  } catch (err) {
    return { success: false, message: (err && err.message) ? err.message : 'Unknown error' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ====== Ticket Management ======
function getNextTicketNo(issueDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Registrations');
  ensureColumn(sheet, 'Ticket No');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ticketNoIdx = headers.indexOf('Ticket No');
  const statusIdx = headers.indexOf('Status');
  const year = issueDate.getFullYear();
  let maxSeq = 0;
  for (let i = 1; i < data.length; i++) {
    const ticketNo = data[i][ticketNoIdx];
    const status = (data[i][statusIdx] || "").toString();
    if (ticketNo && status === "Ticket Issued") {
      const parts = ticketNo.toString().split("/");
      if (parts.length === 2 && parts[1].length === 6) {
        const yy = parts[1].slice(4, 6);
        const fullYear = Number("20" + yy);
        if (fullYear === year) {
          const seq = parseInt(parts[0], 10);
          if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
        }
      }
    }
  }
  const nextSeq = maxSeq + 1;
  const seqStr = String(nextSeq).padStart(4, "0");
  const dmy = formatDateDDMMYY(issueDate);
  return `${seqStr}/${dmy}`;
}

function assignTicketNoToGroup(email, trekDate, issueDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Registrations');
  ensureColumn(sheet, 'Ticket No');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailIdx = headers.indexOf('Email');
  const trekDateIdx = headers.indexOf('Trek Date');
  const ticketNoIdx = headers.indexOf('Ticket No');
  const statusIdx = headers.indexOf('Status');
  const trekDateDMY = formatDateDDMMYYYY(trekDate);
  let rows = [];
  for (let i = 1; i < data.length; i++) {
    if (
      (data[i][emailIdx] + '').toLowerCase() === email &&
      formatDateDDMMYYYY(data[i][trekDateIdx]) === trekDateDMY
    ) {
      rows.push(i);
    }
  }
  if (rows.length === 0) throw new Error("Group not found.");
  let ticketNo = data[rows[0]][ticketNoIdx];
  if (!ticketNo) {
    ticketNo = getNextTicketNo(issueDate);
    for (const row of rows) {
      sheet.getRange(row + 1, ticketNoIdx + 1).setValue(ticketNo);
      if ((data[row][statusIdx] || "").toLowerCase() !== "ticket issued") {
        sheet.getRange(row + 1, statusIdx + 1).setValue("Ticket Issued");
      }
    }
  }
  return ticketNo;
}

function getImageDataUrlFromDrive(driveFileId) {
  try {
    if (!driveFileId) return '';
    const blob = DriveApp.getFileById(driveFileId).getBlob();
    const contentType = blob.getContentType() || 'image/png';
    const base64 = Utilities.base64Encode(blob.getBytes());
    return 'data:' + contentType + ';base64,' + base64;
  } catch (e) {
    return '';
  }
}

function getFeePerPerson() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');
    if (!configSheet) return 0;
    const lastRow = configSheet.getLastRow();
    if (lastRow < 1) return 0;
    const data = configSheet.getRange(1, 1, lastRow, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if ((data[i][0] + '').trim() === 'FeePerPerson') {
        const val = Number((data[i][1] + '').toString().trim());
        return isNaN(val) ? 0 : val;
      }
    }
    return 0;
  } catch (e) {
    return 0;
  }
}

function generateQrDataUrl(text, size) {
  try {
    text = text || '';
    size = size || 300;
    const encoded = encodeURIComponent(text);
    const providers = [
      'https://api.qrserver.com/v1/create-qr-code/?size=' + size + 'x' + size + '&data=' + encoded,
      'https://chart.googleapis.com/chart?chs=' + size + 'x' + size + '&cht=qr&chld=L|1&chl=' + encoded
    ];
    for (let i = 0; i < providers.length; i++) {
      try {
        const resp = UrlFetchApp.fetch(providers[i], { muteHttpExceptions: true });
        const code = resp.getResponseCode ? resp.getResponseCode() : 0;
        if (code === 200) {
          const blob = resp.getBlob();
          const base64 = Utilities.base64Encode(blob.getBytes());
          const contentType = blob.getContentType() || 'image/png';
          return 'data:' + contentType + ';base64,' + base64;
        }
      } catch (innerErr) {
        continue;
      }
    }
    return '';
  } catch (e) {
    return '';
  }
}

function issueTicketAndEmail(email, date) {
  email = (email || '').trim().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Registrations');
  ensureColumn(sheet, 'Ticket No');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const trekDateDMY = formatDateDDMMYYYY(date);

  const groupRows = [];
  const emailIdx = headers.indexOf('Email');
  const trekDateIdx = headers.indexOf('Trek Date');
  for (let i = 1; i < data.length; i++) {
    if (
      (data[i][emailIdx] + '').toLowerCase() === email &&
      formatDateDDMMYYYY(data[i][trekDateIdx]) === trekDateDMY
    ) {
      groupRows.push({ i, row: data[i] });
    }
  }
  if (groupRows.length === 0) {
    return { success: false, message: 'Registration not found.' };
  }

  const issueDate = new Date();
  const ticketNo = assignTicketNoToGroup(email, trekDateDMY, issueDate);

  const firstRow = groupRows[0].row;
  const amountRemitted = firstRow[headers.indexOf('Amount Remitted')];
  const paymentId = toPlainNumberString(firstRow[headers.indexOf('Payment ID')]);
  const emergencyIdx = headers.findIndex(h => (h || '').toLowerCase() === 'emergency contact');
  const emergencyContact = emergencyIdx !== -1 ? (firstRow[emergencyIdx] || '') : '';

  const groupDetails = {
    trekDate: trekDateDMY,
    amountRemitted: amountRemitted,
    paymentId: paymentId,
    ticketNo: ticketNo,
    email: email,
    emergencyContact: emergencyContact
  };
  const trekkers = groupRows.map(({ row }) => ({
    name: row[headers.indexOf('Trekker Name')],
    gender: row[headers.indexOf('Gender')],
    age: row[headers.indexOf('Age')],
    idType: row[headers.indexOf('ID Type')],
    idNumber: toPlainNumberString(row[headers.indexOf('ID Number')]),
    mobile: toPlainNumberString(row[headers.indexOf('Mobile Number')]),
    isForeigner: (function(){
      var idx = headers.indexOf('Is Foreigner');
      return idx !== -1 ? ((row[idx]||'').toString().trim() === 'Yes') : false;
    })()
  }));

  try {
    let logoDataUri = getLogoDataUriFromConfig();
    if (!logoDataUri) {
      const LOGO_FILE_ID = getLogoFileIdFromConfig();
      logoDataUri = LOGO_FILE_ID ? getImageDataUrlFromDrive(LOGO_FILE_ID) : '';
    }

    const qrDataUrl = generateQrDataUrl(ticketNo, 300);

    const htmlTemplate = HtmlService.createTemplateFromFile('Ticket');
    htmlTemplate.groupDetails = groupDetails;
    htmlTemplate.trekkers = trekkers;
    htmlTemplate.logoDataUrl = logoDataUri;
    htmlTemplate.qrDataUrl = qrDataUrl;
    const ticketHtml = htmlTemplate.evaluate().getContent();

    const pdfBlob = Utilities.newBlob(ticketHtml, 'text/html')
      .getAs('application/pdf')
      .setName(`Agasthyarkoodam_Ticket_${groupDetails.ticketNo || 'ticket'}.pdf`);

    const emailHtmlBody = createEmailBody(groupDetails, trekkers);

    MailApp.sendEmail({
      to: email,
      subject: `✅ Agasthyarkoodam Trekking Ticket - ${groupDetails.ticketNo}`,
      htmlBody: emailHtmlBody,
      attachments: [pdfBlob]
    });

    return { success: true, message: `Ticket issued and emailed to ${email}!` };
  } catch (e) {
    return { success: false, message: `Failed to issue ticket: ${e.message}` };
  }
}

function createEmailBody(groupDetails, trekkers) {
  var html = `
 <!DOCTYPE html>
 <html>
 <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto;">
   <div style="text-align: center; padding: 20px 0; border-bottom: 3px solid #0b60a5;">
     <h1 style="color: #0b60a5; margin: 10px 0;">KERALA FOREST & WILDLIFE DEPARTMENT</h1>
     <h3 style="color: #1b5e20; margin: 5px 0;">Agasthyarkoodam Trekking 2026</h3>
     <h2 style="color: #2563eb; margin: 15px 0;">BOOKING CONFIRMED</h2>
   </div>

   <div style="background: #f8f9fa; padding: 20px; margin: 20px 0; border-radius: 5px;">
     <h3 style="color: #2563eb; margin-top: 0;">📋 Booking Details</h3>
     <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
       <tr>
         <td style="padding: 10px; border-bottom: 1px solid #ddd; font-weight: bold; width: 40%;">Ticket Number:</td>
         <td style="padding: 10px; border-bottom: 1px solid #ddd;"><strong>${groupDetails.ticketNo}</strong></td>
       </tr>
       <tr>
         <td style="padding: 10px; border-bottom: 1px solid #ddd; font-weight: bold;">Trek Date:</td>
         <td style="padding: 10px; border-bottom: 1px solid #ddd;">${groupDetails.trekDate}</td>
       </tr>
       <tr>
         <td style="padding: 10px; border-bottom: 1px solid #ddd; font-weight: bold;">Number of Trekkers:</td>
         <td style="padding: 10px; border-bottom: 1px solid #ddd;">${trekkers.length}</td>
       </tr>
       <tr>
         <td style="padding: 10px; border-bottom: 1px solid #ddd; font-weight: bold;">Amount Paid:</td>
         <td style="padding: 10px; border-bottom: 1px solid #ddd;">₹ ${groupDetails.amountRemitted}</td>
       </tr>
       <tr>
         <td style="padding: 10px; font-weight: bold;">Payment ID:</td>
         <td style="padding: 10px;">${groupDetails.paymentId || 'N/A'}</td>
       </tr>
     </table>
   </div>

   <div style="background: #e8f5e9; padding: 20px; border-radius: 5px; margin: 20px 0;">
     <h3 style="color: #1b5e20; margin-top: 0;">✅ What's Next?</h3>
     <ol style="margin: 15px 0; padding-left: 20px;">
       <li style="margin-bottom: 10px;"><strong>Download the PDF</strong> attached to this email</li>
       <li style="margin-bottom: 10px;"><strong>Print all 5 pages</strong> on A4 paper</li>
       <li style="margin-bottom: 10px;"><strong>Complete these documents:</strong>
         <ul style="margin: 5px 0; padding-left: 20px;">
           <li>Medical Certificate (Page 3)</li>
           <li>Affidavit (Page 4)</li>
           <li>Consent Form for minors (Page 5)</li>
         </ul>
       </li>
       <li style="margin-bottom: 10px;"><strong>Bring to trek:</strong>
         <ul style="margin: 5px 0; padding-left: 20px;">
           <li>Printed ticket</li>
           <li>Original photo ID</li>
           <li>Medical certificate</li>
           <li>Signed affidavit</li>
         </ul>
       </li>
     </ol>
   </div>

   <div style="background: #fff3cd; padding: 15px; border: 1px solid #ffeaa7; border-radius: 3px; margin: 20px 0;">
     <h4 style="color: #856404; margin-top: 0;">⏰ Important Notes:</h4>
     <p style="margin: 10px 0;"><strong>Reporting Time:</strong> 7:00 AM</p>
     <p style="margin: 10px 0;"><strong>Reporting Place:</strong> Forest Picket Station, Bonaccord</p>
     <p style="margin: 10px 0;"><strong>Entry Requirements:</strong> Valid only with signed affidavit</p>
   </div>

   <div style="text-align: center; margin: 30px 0;">
     <div style="background: #2563eb; color: white; padding: 15px 30px; display: inline-block; border-radius: 4px;">
       <p style="margin: 0; font-size: 16px;"><strong>Your ticket PDF is attached 📎</strong></p>
     </div>
   </div>

   <div style="background: #e3f2fd; padding: 15px; border-radius: 5px; margin: 20px 0; font-size: 13px;">
     <h4 style="color: #1565c0; margin-top: 0;">📞 Contact Information</h4>
     <p style="margin: 5px 0;"><strong>Wildlife Warden Office:</strong></p>
     <p style="margin: 5px 0;">📍 PTP Nagar, Thiruvananthapuram</p>
     <p style="margin: 5px 0;">📞 0471-2360762 / 8129782524</p>
     <p style="margin: 5px 0;">📧 wildlife.tvm@kerala.gov.in</p>
   </div>

   <div style="font-size: 12px; color: #666; border-top: 1px solid #eee; padding-top: 15px; margin-top: 20px; text-align: center;">
     <p><strong>Kerala FOREST & WILDLIFE DEPARTMENT</strong><br>
     Thiruvananthapuram Wildlife Division</p>
     <p style="color: #1b5e20; margin-top: 10px; font-style: italic;">അഗസ്ത്യാർകൂടം സീസൺ ട്രക്കിംഗ് - 2026</p>
     <p style="margin-top: 10px;">SAVE FOREST • SAVE WATER • SAVE WILDLIFE</p>
   </div>
 </body>
 </html>
   `;
  return html;
}

function userTicketLogin(email) {
  email = email.trim().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Registrations');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailIdx = headers.indexOf('Email');
  const statusIdx = headers.indexOf('Status');
  const trekDateIdx = headers.indexOf('Trek Date');
  const ticketNoIdx = headers.indexOf('Ticket No');
  let groupFound = null;
  for (let i = 1; i < data.length; i++) {
    if (
      (data[i][emailIdx] + '').toLowerCase() === email &&
      (data[i][statusIdx] || '').toLowerCase() === 'ticket issued'
    ) {
      groupFound = {
        email: email,
        trekDate: formatDateDDMMYYYY(data[i][trekDateIdx]),
        ticketNo: data[i][ticketNoIdx]
      };
      break;
    }
  }
  if (!groupFound) return { success: false, message: "Ticket not issued or email not found." };
  let groupRows = [];
  for (let i = 1; i < data.length; i++) {
    if (
      (data[i][emailIdx] + '').toLowerCase() === groupFound.email &&
      formatDateDDMMYYYY(data[i][trekDateIdx]) === groupFound.trekDate
    ) {
      groupRows.push(data[i]);
    }
  }
  return {
    success: true,
    group: {
      Email: groupFound.email,
      'Trek Date': groupFound.trekDate,
      'Ticket No': groupFound.ticketNo,
      Trekkers: groupRows.map(row => ({
        Name: row[headers.indexOf('Trekker Name')],
        Gender: row[headers.indexOf('Gender')],
        Age: row[headers.indexOf('Age')],
        'Photo ID Type': row[headers.indexOf('ID Type')],
        'ID Number': toPlainNumberString(row[headers.indexOf('ID Number')]),
        'Mobile Number': toPlainNumberString(row[headers.indexOf('Mobile Number')]),
        'Is Foreigner': (function(){
          var idx = headers.indexOf('Is Foreigner');
          return idx !== -1 ? ((row[idx]||'').toString().trim() === 'Yes') : false;
        })()
      }))
    }
  };
}

function getAllIssuedTickets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Registrations');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailIdx = headers.indexOf('Email');
  const trekDateIdx = headers.indexOf('Trek Date');
  const ticketNoIdx = headers.indexOf('Ticket No');
  const statusIdx = headers.indexOf('Status');
  const groupMap = {};
  for (let i = 1; i < data.length; i++) {
    if ((data[i][statusIdx] || '').toLowerCase() === 'ticket issued') {
      const email = (data[i][emailIdx] + '').toLowerCase();
      const trekDate = formatDateDDMMYYYY(data[i][trekDateIdx]);
      const key = email + '|' + trekDate;
      if (!groupMap[key]) {
        groupMap[key] = {
          Email: email,
          'Trek Date': trekDate,
          'Ticket No': data[i][ticketNoIdx]
        };
      }
    }
  }
  return Object.values(groupMap);
}

function getIssuedTrekkersReport(dateFromStr, dateToStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Registrations');
  if (!sheet) return { trekkers: [], totalAmount: 0 };

  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return { trekkers: [], totalAmount: 0 };

  const headers = data[0];
  const statusIdx = headers.indexOf('Status');
  const trekDateIdx = headers.indexOf('Trek Date');
  const ticketNoIdx = headers.indexOf('Ticket No');
  const nameIdx = headers.indexOf('Trekker Name');
  const genderIdx = headers.indexOf('Gender');
  const ageIdx = headers.indexOf('Age');
  const idTypeIdx = headers.indexOf('ID Type');
  const idNumberIdx = headers.indexOf('ID Number');
  const mobileIdx = headers.indexOf('Mobile Number');
  const amountRemittedIdx = headers.indexOf('Amount Remitted');
  const emailIdx = headers.indexOf('Email');

  function parseDMY(s) {
    if (!s) return null;
    const parts = (s + '').split("/");
    if (parts.length !== 3) return null;
    const dd = Number(parts[0]), mm = Number(parts[1]), yyyy = Number(parts[2]);
    if (isNaN(dd) || isNaN(mm) || isNaN(yyyy)) return null;
    return new Date(yyyy, mm - 1, dd);
  }
  const dateFrom = parseDMY(dateFromStr);
  const dateTo = parseDMY(dateToStr);

  let results = [];
  let groupSet = {};
  let totalAmount = 0;

  for (let i = 1; i < data.length; i++) {
    const status = (data[i][statusIdx] || "").toString().toLowerCase();
    const trekDateStr = formatDateDDMMYYYY(data[i][trekDateIdx]);
    if (status === "ticket issued" && trekDateStr) {
      let trekDate = parseDMY(trekDateStr);
      if (!trekDate || isNaN(trekDate.getTime())) continue;
      if (trekDate >= dateFrom && trekDate <= dateTo) {
        const groupKey = (data[i][emailIdx] + '').toLowerCase() + "||" + trekDateStr;
        let amt = Number(data[i][amountRemittedIdx]) || 0;
        if (!groupSet[groupKey]) {
          totalAmount += amt;
          groupSet[groupKey] = true;
        }
        results.push({
          TicketNo: data[i][ticketNoIdx] || '',
          Name: data[i][nameIdx] || '',
          Gender: data[i][genderIdx] || '',
          Age: data[i][ageIdx] || '',
          IDType: data[i][idTypeIdx] || '',
          IDNumber: toPlainNumberString(data[i][idNumberIdx]) || '',
          Mobile: toPlainNumberString(data[i][mobileIdx]) || '',
          TrekDate: trekDateStr,
          AmountRemitted: amt
        });
      }
    }
  }

  function ticketSeq(tn) {
    if (!tn) return { seq: 0, raw: (tn || '') };
    const s = (tn + '').toString();
    const parts = s.split('/');
    const n = parseInt(parts[0], 10);
    if (!isNaN(n)) return { seq: n, raw: s };
    return { seq: 0, raw: s };
  }

  results.sort(function(a, b) {
    const da = parseDMY(a.TrekDate);
    const db = parseDMY(b.TrekDate);
    const ta = da ? da.getTime() : 0;
    const tb = db ? db.getTime() : 0;
    if (ta !== tb) return ta - tb;

    const sa = ticketSeq(a.TicketNo);
    const sb = ticketSeq(b.TicketNo);
    if (sa.seq !== sb.seq) return sa.seq - sb.seq;
    return (sa.raw || '').localeCompare(sb.raw || '');
  });

  for (let i = 0; i < results.length; i++) {
    results[i].SlNo = i + 1;
  }

  return { trekkers: results, totalAmount: totalAmount };
}

function apiDownloadTrekkersReportPDF(dateFrom, dateTo) {
  const reportData = getIssuedTrekkersReport(dateFrom, dateTo);
  const template = HtmlService.createTemplateFromFile('TrekkersReport');
  template.trekkers = reportData.trekkers;
  template.totalAmount = reportData.totalAmount;
  template.dateFrom = dateFrom;
  template.dateTo = dateTo;
  const html = template.evaluate().getContent();
  const pdfBlob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  return Utilities.base64Encode(pdfBlob.getBytes());
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Agasthyarkoodam Trekking 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function apiDownloadTicketPDF(email, trekDate) {
  const groupData = getGroupDetails(email, trekDate);
  if (!groupData || !groupData['Ticket No']) return "";
  const ticketDetails = {
    ticketNo: groupData['Ticket No'],
    trekDate: groupData['Trek Date'],
    email: groupData.Email,
    amountRemitted: groupData['Amount Remitted'],
    paymentId: groupData['Payment ID'],
    emergencyContact: groupData['Emergency Contact'] || ''
  };
  const trekkers = groupData.Trekkers.map(trekker => ({
    name: trekker.Name,
    gender: trekker.Gender,
    age: trekker.Age,
    idType: trekker['Photo ID Type'],
    idNumber: trekker['ID Number'],
    mobile: trekker['Mobile Number'],
    isForeigner: trekker['Is Foreigner']
  }));

  let logoDataUri = getLogoDataUriFromConfig();
  if (!logoDataUri) {
    const LOGO_FILE_ID = getLogoFileIdFromConfig();
    logoDataUri = LOGO_FILE_ID ? getImageDataUrlFromDrive(LOGO_FILE_ID) : '';
  }
  const qrDataUrl = generateQrDataUrl(ticketDetails.ticketNo, 300);

  const htmlTemplate = HtmlService.createTemplateFromFile('Ticket');
  htmlTemplate.groupDetails = ticketDetails;
  htmlTemplate.trekkers = trekkers;
  htmlTemplate.logoDataUrl = logoDataUri;
  htmlTemplate.qrDataUrl = qrDataUrl;
  const html = htmlTemplate.evaluate().getContent();
  const pdfBlob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  return Utilities.base64Encode(pdfBlob.getBytes());
}

// ====== API Wrappers (Exposed to Client) ======
function apiUserTicketLogin(email) {
  return userTicketLogin(email);
}

function apiAdminGetGroupTicket(email, trekDate) {
  return getGroupDetails(email, trekDate);
}

function apiGetAllIssuedTickets() {
  return getAllIssuedTickets();
}

function apiGetIssuedTrekkersReport(dateFrom, dateTo) {
  return getIssuedTrekkersReport(dateFrom, dateTo);
}

function apiUserLoginStatus(email) {
  return userLoginStatus(email);
}

function apiGetFeePerPerson() {
  return getFeePerPerson();
}

function apiUpdatePaymentId(email, trekDate, paymentId) {
  return updatePaymentId(email, trekDate, paymentId);
}

// ====== Admin: Delete Group ======
function deleteGroup(email, trekDate) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server busy. Try again.' };
  }
  try {
    email = (email || '').toString().trim().toLowerCase();
    if (!email) return { success: false, message: 'Invalid email.' };
    var dateDMY = formatDateDDMMYYYY(trekDate);
    if (!dateDMY) return { success: false, message: 'Invalid trek date.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Registrations');
    if (!sheet) return { success: false, message: 'Registrations sheet not found.' };

    var data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return { success: false, message: 'No registration data.' };

    var headers = data[0].map(h => (h || '').toString());
    var emailIdx = headers.findIndex(h => (h || '').toLowerCase() === 'email');
    var trekDateIdx = headers.findIndex(h => (h || '').toLowerCase() === 'trek date' || (h || '').toLowerCase() === 'trekdate');
    if (emailIdx === -1 || trekDateIdx === -1) return { success: false, message: 'Required columns not found.' };

    var rowsToDelete = [];
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][emailIdx] || '').toString().toLowerCase();
      var rowDate = formatDateDDMMYYYY(data[i][trekDateIdx]);
      if (rowEmail === email && rowDate === dateDMY) {
        rowsToDelete.push(i + 1);
      }
    }

    if (rowsToDelete.length === 0) return { success: false, message: 'No registration rows found for this group.' };

    rowsToDelete.reverse().forEach(function(r) {
      try { sheet.deleteRow(r); } catch (e) {}
    });

    return { success: true, message: 'Group registrations deleted.' };
  } catch (err) {
    return { success: false, message: (err && err.message) ? err.message : 'Unknown error' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ====== Admin: Get All Authorized Emails ======
function getAllAuthorizedEmails() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('AuthorizedEmails');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return [];
    var headers = data[0].map(h => (h || '').toString());
    var emailIdx = headers.findIndex(h => (h || '').toLowerCase() === 'authorized email' || (h || '').toLowerCase() === 'email');
    var maxIdx = headers.findIndex(h => (h || '').toLowerCase() === 'maxtrekkers' || (h || '').toLowerCase() === 'max trekkers');
    var dateIdx = headers.findIndex(h => (h || '').toLowerCase() === 'alloweddate' || (h || '').toLowerCase() === 'allowed date');

    var out = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var item = {
        Email: emailIdx !== -1 ? (row[emailIdx] || '').toString() : (row[0] || '').toString(),
        MaxTrekkers: maxIdx !== -1 ? row[maxIdx] : (row[1] || ''),
        AllowedDate: dateIdx !== -1 ? formatDateDDMMYYYY(row[dateIdx]) : (row[2] ? formatDateDDMMYYYY(row[2]) : '')
      };
      out.push(item);
    }
    return out;
  } catch (e) {
    return [];
  }
}

// ====== Admin: Delete Authorized Email ======
function deleteAuthorizedEmail(email) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server busy. Try again.' };
  }
  try {
    email = (email || '').toString().trim().toLowerCase();
    if (!email) return { success: false, message: 'Invalid email.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('AuthorizedEmails');
    if (!sheet) return { success: false, message: 'AuthorizedEmails sheet not found.' };

    var data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return { success: false, message: 'No authorized emails.' };
    var headers = data[0].map(h => (h || '').toString());
    var emailIdx = headers.findIndex(h => (h || '').toLowerCase() === 'authorized email' || (h || '').toLowerCase() === 'email');

    if (emailIdx === -1) emailIdx = 0;

    var found = false;
    for (var i = data.length - 1; i >= 1; i--) {
      var rowEmail = (data[i][emailIdx] || '').toString().toLowerCase();
      if (rowEmail === email) {
        sheet.deleteRow(i + 1);
        found = true;
      }
    }
    if (!found) return { success: false, message: 'Email not found.' };
    return { success: true, message: 'Authorized email deleted.' };
  } catch (err) {
    return { success: false, message: (err && err.message) ? err.message : 'Unknown error' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ====== Admin: Get Yet to Register Emails ======
function getYetToRegisterEmails() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var authSheet = ss.getSheetByName('AuthorizedEmails');
    var regSheet = ss.getSheetByName('Registrations');

    if (!authSheet) return [];

    var authData = authSheet.getDataRange().getValues();
    if (!authData || authData.length <= 1) return [];

    var authHeaders = authData[0].map(h => (h || '').toString());
    var emailIdx = authHeaders.findIndex(h => (h || '').toLowerCase() === 'authorized email' || (h || '').toLowerCase() === 'email');
    var maxIdx = authHeaders.findIndex(h => (h || '').toLowerCase() === 'maxtrekkers' || (h || '').toLowerCase() === 'max trekkers');
    var dateIdx = authHeaders.findIndex(h => (h || '').toLowerCase() === 'alloweddate' || (h || '').toLowerCase() === 'allowed date');

    if (emailIdx === -1) emailIdx = 0;
    if (maxIdx === -1) maxIdx = 1;
    if (dateIdx === -1) dateIdx = 2;

    var registeredEmails = new Set();
    if (regSheet) {
      var regData = regSheet.getDataRange().getValues();
      if (regData && regData.length > 1) {
        var regHeaders = regData[0].map(h => (h || '').toString());
        var regEmailIdx = regHeaders.findIndex(h => (h || '').toLowerCase() === 'email');
        if (regEmailIdx !== -1) {
          for (var i = 1; i < regData.length; i++) {
            var email = (regData[i][regEmailIdx] || '').toString().toLowerCase().trim();
            if (email) {
              registeredEmails.add(email);
            }
          }
        }
      }
    }

    var out = [];
    for (var i = 1; i < authData.length; i++) {
      var row = authData[i];
      var email = (row[emailIdx] || '').toString().toLowerCase().trim();

      if (email && !registeredEmails.has(email)) {
        var allowedDateStr = dateIdx !== -1 ? formatDateDDMMYYYY(row[dateIdx]) : '';
        var item = {
          Email: email,
          MaxTrekkers: maxIdx !== -1 ? row[maxIdx] : '',
          AllowedDate: allowedDateStr,
          AllowedDateSort: parseDateDMY(allowedDateStr)
        };
        out.push(item);
      }
    }

    out.sort(function(a, b) {
      var dateA = a.AllowedDateSort ? a.AllowedDateSort.getTime() : 0;
      var dateB = b.AllowedDateSort ? b.AllowedDateSort.getTime() : 0;
      return dateA - dateB;
    });

    out.forEach(function(item) {
      delete item.AllowedDateSort;
    });

    return out;
  } catch (e) {
    return [];
  }
}

// ====== Helper: Parse DD/MM/YYYY to Date ======
function parseDateDMY(dateStr) {
  if (!dateStr) return null;
  try {
    var parts = (dateStr + '').split('/');
    if (parts.length !== 3) return null;
    var dd = Number(parts[0]);
    var mm = Number(parts[1]);
    var yyyy = Number(parts[2]);
    if (isNaN(dd) || isNaN(mm) || isNaN(yyyy)) return null;
    return new Date(yyyy, mm - 1, dd);
  } catch (e) {
    return null;
  }
}

// ====== Admin: Delete Yet to Register Email ======
function deleteYetToRegisterEmail(email) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server busy. Try again.' };
  }
  try {
    email = (email || '').toString().trim().toLowerCase();
    if (!email) return { success: false, message: 'Invalid email.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var authSheet = ss.getSheetByName('AuthorizedEmails');
    if (!authSheet) return { success: false, message: 'AuthorizedEmails sheet not found.' };

    var data = authSheet.getDataRange().getValues();
    if (!data || data.length <= 1) return { success: false, message: 'No authorized emails.' };

    var headers = data[0].map(h => (h || '').toString());
    var emailIdx = headers.findIndex(h => (h || '').toLowerCase() === 'authorized email' || (h || '').toLowerCase() === 'email');

    if (emailIdx === -1) emailIdx = 0;

    var found = false;
    for (var i = data.length - 1; i >= 1; i--) {
      var rowEmail = (data[i][emailIdx] || '').toString().toLowerCase();
      if (rowEmail === email) {
        authSheet.deleteRow(i + 1);
        found = true;
      }
    }

    if (!found) return { success: false, message: 'Email not found in authorized list.' };
    return { success: true, message: 'Authorization deleted successfully.' };
  } catch (err) {
    return { success: false, message: (err && err.message) ? err.message : 'Unknown error' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ====== Download Summary Report PDF ======
function apiDownloadSummaryReportPDF(dateFrom, dateTo) {
  const reportData = getSummaryReport(dateFrom, dateTo);
  const summary = reportData.summary;

  let tableHtml = `
    <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
      <thead>
        <tr style="background: #1e40af; color: white;">
          <th style="border: 1px solid #ddd; padding: 10px; text-align: left;">Date</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: center;">No. of Trekkers</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: center;">Male</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: center;">Female</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: center;">Others</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: right;">Amount (₹)</th>
          <th style="border: 1px solid #ddd; padding: 10px; text-align: center;">No of Groups</th>
        </tr>
      </thead>
      <tbody>
  `;

  summary.forEach((row, idx) => {
    const isTotal = row.Date === 'TOTAL';
    const bgColor = isTotal ? '#f0f0f0' : (idx % 2 === 0 ? '#ffffff' : '#f8f9ff');
    const fontWeight = isTotal ? 'bold' : 'normal';

    tableHtml += `
      <tr style="background: ${bgColor};">
        <td style="border: 1px solid #ddd; padding: 10px; font-weight: ${fontWeight};">${row.Date}</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-weight: ${fontWeight};">${row['No. of Trekkers']}</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-weight: ${fontWeight};">${row.Male}</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-weight: ${fontWeight};">${row.Female}</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-weight: ${fontWeight};">${row.Others}</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: right; font-weight: ${fontWeight};">₹ ${row.Amount.toLocaleString('en-IN')}</td>
        <td style="border: 1px solid #ddd; padding: 10px; text-align: center; font-weight: ${fontWeight};">${row['No of Groups']}</td>
      </tr>
    `;
  });

  tableHtml += `</tbody></table>`;

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>Summary Report</title>
      <meta charset="utf-8">
    </head>
    <body style="font-family: Arial, sans-serif; margin: 20px;">
      <div style="text-align: center; margin-bottom: 30px;">
        <h1 style="color: #0b60a5; margin: 0;">Agasthyarkoodam Trekking - Summary Report</h1>
        <h3 style="color: #666; margin: 5px 0;">Kerala Forest & Wildlife Department</h3>
        <p style="color: #999; margin-top: 10px;">
          Report Period: ${dateFrom} to ${dateTo}
        </p>
      </div>

      ${tableHtml}

      <div style="margin-top: 40px; text-align: center; color: #666; font-size: 12px;">
        <p>Generated on: ${new Date().toLocaleString()}</p>
        <p>Thiruvananthapuram Wildlife Division</p>
      </div>
    </body>
    </html>
  `;

  const pdfBlob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  return Utilities.base64Encode(pdfBlob.getBytes());
}

// ====== STATISTICS FUNCTIONS ======
function getStatistics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const regSheet = ss.getSheetByName('Registrations');

    if (!regSheet) {
      return {
        totalRegistrations: 0,
        pendingApprovals: 0,
        ticketsIssued: 0,
        totalRevenue: 0,
        statusBreakdown: {},
        trekDateBreakdown: {},
        genderDistribution: {},
        ageDistribution: {},
        averageTrkkersPerGroup: 0,
        totalTrekkers: 0,
        revenueByDate: {}
      };
    }

    const data = regSheet.getDataRange().getValues();
    if (data.length <= 1) return null;

    const headers = data[0].map(h => (h || '').toString());
    const emailIdx = headers.indexOf('Email');
    const trekDateIdx = headers.indexOf('Trek Date');
    const statusIdx = headers.indexOf('Status');
    const amountIdx = headers.indexOf('Amount Remitted');
    const genderIdx = headers.indexOf('Gender');
    const ageIdx = headers.indexOf('Age');
    const noOfTrkkersIdx = headers.indexOf('No. of Trekkers');

    let stats = {
      totalRegistrations: 0,
      pendingApprovals: 0,
      ticketsIssued: 0,
      totalRevenue: 0,
      statusBreakdown: { 'Pending': 0, 'Ticket Issued': 0 },
      trekDateBreakdown: {},
      genderDistribution: { 'Male': 0, 'Female': 0, 'Other': 0 },
      ageDistribution: { '14-20': 0, '21-30': 0, '31-40': 0, '41-50': 0, '50+': 0 },
      totalTrekkers: 0,
      revenueByDate: {},
      registrationsByDay: {}
    };

    const processedGroups = new Set();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      const email = (row[emailIdx] || '').toString().toLowerCase();
      const trekDate = formatDateDDMMYYYY(row[trekDateIdx]);
      const groupKey = email + '||' + trekDate;

      if (!processedGroups.has(groupKey)) {
        processedGroups.add(groupKey);
        stats.totalRegistrations++;

        const status = (row[statusIdx] || '').toString().trim();
        if (status === 'Pending') {
          stats.pendingApprovals++;
          stats.statusBreakdown['Pending']++;
        } else if (status === 'Ticket Issued') {
          stats.ticketsIssued++;
          stats.statusBreakdown['Ticket Issued']++;
        }

        const amount = Number(row[amountIdx]) || 0;
        stats.totalRevenue += amount;

        if (!stats.revenueByDate[trekDate]) {
          stats.revenueByDate[trekDate] = 0;
        }
        stats.revenueByDate[trekDate] += amount;

        if (!stats.trekDateBreakdown[trekDate]) {
          stats.trekDateBreakdown[trekDate] = 0;
        }
        stats.trekDateBreakdown[trekDate]++;

        const noOfTrekkers = Number(row[noOfTrkkersIdx]) || 1;
        stats.totalTrekkers += noOfTrekkers;
      }

      const gender = (row[genderIdx] || '').toString().trim();
      if (gender && stats.genderDistribution.hasOwnProperty(gender)) {
        stats.genderDistribution[gender]++;
      }

      const age = Number(row[ageIdx]) || 0;
      if (age > 0) {
        if (age >= 14 && age <= 20) stats.ageDistribution['14-20']++;
        else if (age >= 21 && age <= 30) stats.ageDistribution['21-30']++;
        else if (age >= 31 && age <= 40) stats.ageDistribution['31-40']++;
        else if (age >= 41 && age <= 50) stats.ageDistribution['41-50']++;
        else if (age > 50) stats.ageDistribution['50+']++;
      }
    }

    stats.averageTrkkersPerGroup = stats.totalRegistrations > 0
      ? (stats.totalTrekkers / stats.totalRegistrations).toFixed(2)
      : 0;

    return stats;
  } catch (e) {
    Logger.log('Statistics Error: ' + e.message);
    return null;
  }
}

// ====== NEW: Summary Report Functions ======
function getSummaryReport(dateFromStr, dateToStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Registrations');
  if (!sheet) return { summary: [], totalAmount: 0 };

  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return { summary: [], totalAmount: 0 };

  const headers = data[0];
  const statusIdx = headers.indexOf('Status');
  const trekDateIdx = headers.indexOf('Trek Date');
  const amountRemittedIdx = headers.indexOf('Amount Remitted');
  const emailIdx = headers.indexOf('Email');
  const genderIdx = headers.indexOf('Gender');

  function parseDMY(s) {
    if (!s) return null;
    const parts = (s + '').split("/");
    if (parts.length !== 3) return null;
    const dd = Number(parts[0]), mm = Number(parts[1]), yyyy = Number(parts[2]);
    if (isNaN(dd) || isNaN(mm) || isNaN(yyyy)) return null;
    return new Date(yyyy, mm - 1, dd);
  }

  const dateFrom = parseDMY(dateFromStr);
  const dateTo = parseDMY(dateToStr);

  const processedGroupsForAmount = new Set();
  const summaryMap = {};

  for (let i = 1; i < data.length; i++) {
    const status = (data[i][statusIdx] || "").toString().toLowerCase();
    const trekDateStr = formatDateDDMMYYYY(data[i][trekDateIdx]);

    if (status === "ticket issued" && trekDateStr) {
      let trekDate = parseDMY(trekDateStr);
      if (!trekDate || isNaN(trekDate.getTime())) continue;
      if (trekDate < dateFrom || trekDate > dateTo) continue;

      const gender = (data[i][genderIdx] || "").toString().trim();
      const email = (data[i][emailIdx] || "").toString().toLowerCase();
      const dateKey = trekDateStr;
      const groupKey = `${email}||${trekDateStr}`;

      if (!summaryMap[dateKey]) {
        summaryMap[dateKey] = {
          date: trekDateStr,
          trekkers: 0,
          male: 0,
          female: 0,
          other: 0,
          amount: 0,
          groups: new Set()
        };
      }

      summaryMap[dateKey].trekkers += 1;

      if (gender.toLowerCase() === 'male') {
        summaryMap[dateKey].male += 1;
      } else if (gender.toLowerCase() === 'female') {
        summaryMap[dateKey].female += 1;
      } else {
        summaryMap[dateKey].other += 1;
      }

      if (!processedGroupsForAmount.has(groupKey)) {
        const amount = Number(data[i][amountRemittedIdx]) || 0;
        summaryMap[dateKey].amount += amount;
        processedGroupsForAmount.add(groupKey);
      }

      summaryMap[dateKey].groups.add(groupKey);
    }
  }

  let summary = Object.values(summaryMap).map(item => ({
    Date: item.date,
    'No. of Trekkers': item.trekkers,
    'Male': item.male,
    'Female': item.female,
    'Others': item.other,
    'Amount': item.amount,
    'No of Groups': item.groups.size
  }));

  summary.sort((a, b) => {
    const parseDate = (dateStr) => {
      const parts = dateStr.split('/');
      return new Date(parts[2], parts[1] - 1, parts[0]);
    };
    return parseDate(a.Date) - parseDate(b.Date);
  });

  const totalTrekkers = summary.reduce((sum, item) => sum + item['No. of Trekkers'], 0);
  const totalMale = summary.reduce((sum, item) => sum + item.Male, 0);
  const totalFemale = summary.reduce((sum, item) => sum + item.Female, 0);
  const totalOther = summary.reduce((sum, item) => sum + item.Others, 0);
  const totalAmount = summary.reduce((sum, item) => sum + item.Amount, 0);
  const totalGroups = summary.reduce((sum, item) => sum + item['No of Groups'], 0);

  summary.push({
    Date: 'TOTAL',
    'No. of Trekkers': totalTrekkers,
    'Male': totalMale,
    'Female': totalFemale,
    'Others': totalOther,
    'Amount': totalAmount,
    'No of Groups': totalGroups
  });

  return { summary, totalAmount };
}

// ====== API Wrapper for Summary Report ======
function apiGetSummaryReport(dateFrom, dateTo) {
  return getSummaryReport(dateFrom, dateTo);
}
// ====== API WRAPPER ======
function apiGetStatistics() {
  return getStatistics();
}

// ====== API Wrapper for Yet to Register ======
function apiGetYetToRegisterEmails() {
  return getYetToRegisterEmails();
}

function apiDeleteYetToRegisterEmail(email) {
  return deleteYetToRegisterEmail(email);
}
