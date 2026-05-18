const SPREADSHEET_ID = "1oup4Beadx0z_oXGpPzRZn0e2HySoPaFMG7wxLerOcSU";
const USERS_SHEET_NAME = "NFC REGISTERED";
const ADMIN_SHEET_NAME = "ADMIN";
const TIMEZONE = "Asia/Manila";
const WORK_SCHEDULE_SHEET = "WORK SCHEDULE";

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Arrange sheets in the default order
function arrangeSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var desiredOrder = [
    'ADMIN',
    'NFC REGISTERED',
    '201 FILES',
    '201 FILES ARCHIVED',
    'SERVICE CREDITS',
    'CREDIT HISTORY',
    'LEAVE REQUESTS',
    'SCANNER SETTINGS',
    'WORK SCHEDULE',
    'NOTIFICATIONS'
  ];
  for (var i = 0; i < desiredOrder.length; i++) {
    var sheet = ss.getSheetByName(desiredOrder[i]);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1);
    }
  }
}

// Check if email exists in the NFC REGISTERED sheet
function checkIsRegisteredTeacher(email) {
  if (!email) return false;
  const sheet = getSpreadsheet().getSheetByName(USERS_SHEET_NAME);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const normalizedEmail = email.toString().trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    var rowEmail = (data[i][3] || '').toString().trim().toLowerCase();
    if (rowEmail === normalizedEmail) return true;
  }
  return false;
}

// Serve the correct page based on user role
function doGet(e) {
  const page = e && e.parameter && e.parameter.page ? e.parameter.page : null;

  if (page === "scanner") {
    return HtmlService.createTemplateFromFile('Nfc_Scanner')
      .evaluate()
      .setTitle("TeachTap - NFC Scanner")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  var email = getCurrentEmail();

  // 1) Check if email is in ADMIN sheet → Admin Dashboard
  if (email && checkIsAdmin(email)) {
    return HtmlService.createTemplateFromFile('Admin')
      .evaluate()
      .setTitle("TeachTap - Admin Dashboard")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // 2) Check if email is in NFC REGISTERED sheet → Teacher Dashboard
  if (email && checkIsRegisteredTeacher(email)) {
    return HtmlService.createTemplateFromFile('Teacher')
      .evaluate()
      .setTitle("TeachTap - Teacher Dashboard")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // 3) Not found in either sheet → show Not Registered page
  var template = HtmlService.createTemplateFromFile('NotRegistered');
  template.userEmail = email || 'Unknown';
  return template.evaluate()
    .setTitle("TeachTap - Not Registered")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Check if email exists in the ADMIN sheet
function checkIsAdmin(email) {
  if (!email) return false;
  const sheet = getSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const normalizedEmail = email.toString().trim().toLowerCase();
  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
      const cellVal = data[i][j] ? data[i][j].toString().trim().toLowerCase() : "";
      if (cellVal === normalizedEmail) return true;
    }
  }
  return false;
}

// Get current user email - prioritize getActiveUser (actual accessing user)
function getCurrentEmail() {
  // First try Session.getActiveUser - returns the actual person accessing the web app
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) return email.toString().trim().toLowerCase();
  } catch(e) {}

  // Fallback to identity token (note: in "Execute as: Me" mode this returns the owner's email)
  try {
    var token = ScriptApp.getIdentityToken();
    if (token) {
      var parts = token.split('.');
      var payload = parts[1].replace(/-/g, '+').replace(/_/g, '/');
      while (payload.length % 4 !== 0) payload += '=';
      var decoded = Utilities.newBlob(Utilities.base64Decode(payload)).getDataAsString();
      var claims = JSON.parse(decoded);
      if (claims.email) return claims.email.toString().trim().toLowerCase();
    }
  } catch(e) {}

  return "";
}

function getCurrentUserRole() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch(e) {}
  return { email: email };
}

function getLogoutUrl() {
  return ScriptApp.getService().getUrl();
}

// Format ADMIN PIN column as plain text and pad with leading zeros
function formatAdminPinColumn() {
  var sheet = getSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var pinRange = sheet.getRange(2, 2, lastRow - 1, 1);
  pinRange.setNumberFormat('@');
  var values = pinRange.getValues();
  for (var i = 0; i < values.length; i++) {
    var pin = (values[i][0] || '').toString().trim();
    while (pin.length < 6) pin = '0' + pin;
    values[i][0] = pin;
  }
  pinRange.setValues(values);
}

// Get admin settings for the current logged-in admin
function getAdminSettings() {
  var email = getCurrentEmail();
  if (!email) return { email: '' };
  var sheet = getSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return { email: '' };
  var data = sheet.getDataRange().getValues();
  var normalizedEmail = email.toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = (data[i][0] || '').toString().trim().toLowerCase();
    if (rowEmail === normalizedEmail) {
      return { email: data[i][0].toString().trim() };
    }
  }
  return { email: email };
}

// Update admin settings (email and/or PIN)
function updateAdminSettings(newEmail, currentPin, newPin) {
  var email = getCurrentEmail();
  if (!email) return { success: false, message: 'Unable to identify admin email.' };
  if (!checkIsAdmin(email)) return { success: false, message: 'Not authorized as admin.' };
  var sheet = getSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return { success: false, message: 'Admin sheet not found.' };
  var data = sheet.getDataRange().getValues();
  var normalizedEmail = email.toString().trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var rowEmail = (data[i][0] || '').toString().trim().toLowerCase();
    if (rowEmail === normalizedEmail) {
      // Verify current PIN
      var storedPin = (data[i][1] || '').toString().trim();
      while (storedPin.length < 6) storedPin = '0' + storedPin;
      var enteredPin = currentPin.toString().trim();
      while (enteredPin.length < 6) enteredPin = '0' + enteredPin;
      if (storedPin !== enteredPin) {
        return { success: false, message: 'Current PIN is incorrect.' };
      }
      // Update email
      if (newEmail && newEmail.trim()) {
        sheet.getRange(i + 1, 1).setValue(newEmail.trim());
      }
      // Update PIN if provided
      if (newPin && newPin.toString().trim().length === 6) {
        var pinCell = sheet.getRange(i + 1, 2);
        pinCell.setNumberFormat('@');
        pinCell.setValue(newPin.toString().trim());
      }
      return { success: true, message: 'Admin settings updated successfully.' };
    }
  }
  return { success: false, message: 'Admin account not found.' };
}

// Forgot PIN - send admin PIN to admin's email
function forgotAdminPin() {
  var email = getCurrentEmail();
  if (!email) return { success: false, message: 'Unable to identify admin email.' };
  if (!checkIsAdmin(email)) return { success: false, message: 'Not authorized as admin.' };
  var sheet = getSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return { success: false, message: 'Admin sheet not found.' };
  var data = sheet.getDataRange().getValues();
  var normalizedEmail = email.toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = (data[i][0] || '').toString().trim().toLowerCase();
    if (rowEmail === normalizedEmail) {
      var storedPin = (data[i][1] || '').toString().trim();
      while (storedPin.length < 6) storedPin = '0' + storedPin;
      try {
        MailApp.sendEmail({
          to: data[i][0].toString().trim(),
          subject: 'TeachTap - Your Admin PIN Recovery',
          htmlBody: '<div style="font-family:Poppins,Arial,sans-serif;max-width:400px;margin:0 auto;padding:24px;">' +
            '<h2 style="color:#1e293b;margin-bottom:8px;">TeachTap Admin PIN Recovery</h2>' +
            '<p style="color:#64748b;">You requested your admin PIN. Here it is:</p>' +
            '<div style="background:#f1f5f9;border-radius:12px;padding:20px;text-align:center;margin:16px 0;">' +
            '<span style="font-size:32px;font-weight:700;letter-spacing:8px;color:#1e293b;">' + storedPin + '</span></div>' +
            '<p style="color:#94a3b8;font-size:12px;">If you did not request this, please change your PIN immediately.</p></div>'
        });
        return { success: true, message: 'Your PIN has been sent to your email (' + data[i][0].toString().trim() + ').' };
      } catch(e) {
        return { success: false, message: 'Failed to send email: ' + e.message };
      }
    }
  }
  return { success: false, message: 'Admin email not found in the system.' };
}

// Forgot PIN - send teacher PIN reset notification
function forgotTeacherPin(email) {
  if (!email) return { success: false, message: 'Email is required.' };
  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var normalizedEmail = email.toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = (data[i][3] || '').toString().trim().toLowerCase();
    if (rowEmail === normalizedEmail) {
      var name = (data[i][1] || '').toString().trim();
      var storedPin = (data[i][4] || '').toString().trim();
      if (!storedPin) {
        return { success: false, message: 'No PIN has been set yet. Please log in and set your PIN first.' };
      }
      while (storedPin.length < 4) storedPin = '0' + storedPin;
      try {
        MailApp.sendEmail({
          to: data[i][3].toString().trim(),
          subject: 'TeachTap - Your Teacher PIN Recovery',
          htmlBody: '<div style="font-family:Poppins,Arial,sans-serif;max-width:400px;margin:0 auto;padding:24px;">' +
            '<h2 style="color:#1e293b;margin-bottom:8px;">TeachTap Teacher PIN Recovery</h2>' +
            '<p style="color:#64748b;">Hi ' + name + ',</p>' +
            '<p style="color:#64748b;">You requested your teacher PIN. Here it is:</p>' +
            '<div style="background:#f1f5f9;border-radius:12px;padding:20px;text-align:center;margin:16px 0;">' +
            '<span style="font-size:32px;font-weight:700;letter-spacing:8px;color:#1e293b;">' + storedPin + '</span></div>' +
            '<p style="color:#94a3b8;font-size:12px;">If you did not request this, please change your PIN immediately.</p></div>'
        });
        return { success: true, message: 'Your PIN has been sent to your email (' + data[i][3].toString().trim() + ').' };
      } catch(e) {
        return { success: false, message: 'Failed to send email: ' + e.message };
      }
    }
  }
  return { success: false, message: 'Email not found in the system. Please check your email address.' };
}

// Verify admin 6-digit PIN from the ADMIN sheet
function verifyAdminPin(pin) {
  var email = getCurrentEmail();
  if (!email) return { success: false, message: 'Unable to identify admin email.' };
  if (!checkIsAdmin(email)) return { success: false, message: 'Not authorized as admin.' };
  var sheet = getSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return { success: false, message: 'Admin sheet not found.' };
  var data = sheet.getDataRange().getValues();
  var normalizedEmail = email.toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = (data[i][0] || '').toString().trim().toLowerCase();
    if (rowEmail === normalizedEmail) {
      var storedPin = (data[i][1] || '').toString().trim();
      // Pad stored PIN with leading zeros to 6 digits (Sheets may strip leading zeros)
      while (storedPin.length < 6) storedPin = '0' + storedPin;
      var enteredPin = pin.toString().trim();
      while (enteredPin.length < 6) enteredPin = '0' + enteredPin;
      if (storedPin === enteredPin) {
        return { success: true };
      } else {
        return { success: false, message: 'Incorrect PIN. Please try again.' };
      }
    }
  }
  return { success: false, message: 'Admin PIN not found for this email.' };
}

// Format cell value as 12-hour AM/PM time
function formatTime12h(cellValue) {
  if (!cellValue) return "";
  if (cellValue instanceof Date) {
    var sheetTz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(cellValue, sheetTz, "h:mm:ss a");
  }
  var str = cellValue.toString().trim();
  if (/[AaPp][Mm]/.test(str)) return str;
  var match = str.match(/(\d{1,2}):(\d{2})(:(\d{2}))?/);
  if (match) {
    var h = parseInt(match[1]);
    var m = match[2];
    var s = match[4] || "00";
    var ampm = h >= 12 ? "PM" : "AM";
    if (h === 0) h = 12;
    else if (h > 12) h = h - 12;
    return h + ":" + m + ":" + s + " " + ampm;
  }
  return str;
}

// Compare date values
function isSameDate(cellValue, targetDateStr) {
  if (!cellValue) return false;
  let formattedDate;
  if (cellValue instanceof Date) {
    formattedDate = Utilities.formatDate(cellValue, TIMEZONE, "yyyy-MM-dd");
  } else {
    const str = cellValue.toString();
    const match = str.match(/(\d{4}-\d{2}-\d{2})/);
    if (match) formattedDate = match[1];
    else return false;
  }
  return formattedDate === targetDateStr;
}

// Convert image to base64 data URI (bypasses CORS/CSP in sandboxed iframe)
function getProfileImageBase64(imageInput) {
  if (!imageInput) return "";
  var cellText = imageInput.toString().trim();
  if (!cellText) return "";

  var imgFormula = cellText.match(/=IMAGE\s*\(\s*["'](.*?)["']/i);
  if (imgFormula && imgFormula[1]) cellText = imgFormula[1];

  var fileId = null;
  var m1 = cellText.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m1) fileId = m1[1];
  var m2 = cellText.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) fileId = m2[1];
  if (!fileId && cellText.match(/^[a-zA-Z0-9_-]{25,}$/)) fileId = cellText;

  if (fileId) {
    try {
      var file = DriveApp.getFileById(fileId);
      var blob = file.getBlob();
      var contentType = blob.getContentType();
      var b64 = Utilities.base64Encode(blob.getBytes());
      return "data:" + contentType + ";base64," + b64;
    } catch (err) {
      return "https://lh3.googleusercontent.com/d/" + fileId;
    }
  }

  if (cellText.match(/^https?:\/\//i)) {
    try {
      var response = UrlFetchApp.fetch(cellText, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        var cType = response.getHeaders()["Content-Type"] || "image/jpeg";
        var bytes = response.getContent();
        return "data:" + cType + ";base64," + Utilities.base64Encode(bytes);
      }
    } catch (err) {}
    return cellText;
  }

  return "";
}

// Remove duplicate IN records across daily sheets
function removeDuplicateInRecords() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  let totalDeleted = 0;
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sheetName)) return;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    const seenUids = new Set();
    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const ts = row[0];
      if (!ts || !isSameDate(ts, sheetName)) continue;
      const uid = row[1] ? row[1].toString().trim() : "";
      if (!uid) continue;
      if (row[3] === "IN") {
        if (seenUids.has(uid)) { rowsToDelete.push(i + 1); }
        else { seenUids.add(uid); }
      }
    }
    for (let i = rowsToDelete.length - 1; i >= 0; i--) { sheet.deleteRow(rowsToDelete[i]); }
    if (rowsToDelete.length > 0) totalDeleted += rowsToDelete.length;
  });
  return "Done. Deleted " + totalDeleted + " duplicate IN record(s).";
}

// Record attendance (IN/OUT) and return scan result with image
function recordAttendance(uid) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    if (!uid || uid.toString().trim() === "") {
      return { success: false, message: "Empty card UID" };
    }
    uid = uid.toString().trim();

    const userDetails = getUserDetailsByUID(uid);
    if (!userDetails) {
      return { success: false, message: "NFC " + uid + " not registered." };
    }
    const employeeName = userDetails.name;
    var profileImageB64 = "";
    if (userDetails.imageRaw) {
      try { profileImageB64 = getProfileImageBase64(userDetails.imageRaw); } catch(e) {}
    }

    const todayStr = getTodayDateString();
    const dailySheet = getOrCreateDailySheet(todayStr);
    const dailyData = dailySheet.getDataRange().getValues();

    let dailyRowIndex = -1;
    let dailyExistingRow = null;
    const duplicateRowsToDelete = [];

    for (let i = 1; i < dailyData.length; i++) {
      const row = dailyData[i];
      const ts = row[0];
      if (!ts || !isSameDate(ts, todayStr)) continue;
      const rowUid = row[1] ? row[1].toString().trim() : "";
      if (rowUid === uid) {
        if (dailyRowIndex !== -1) { duplicateRowsToDelete.push(i + 1); }
        else { dailyRowIndex = i + 1; dailyExistingRow = row; }
      }
    }

    if (duplicateRowsToDelete.length > 0) {
      for (let i = duplicateRowsToDelete.length - 1; i >= 0; i--) {
        dailySheet.deleteRow(duplicateRowsToDelete[i]);
      }
    }

    const hasIn = dailyExistingRow ? dailyExistingRow[3] === "IN" : false;
    const hasOut = dailyExistingRow ? dailyExistingRow[6] === "OUT" : false;
    let eventType = null;
    let errorMessage = "";

    if (!hasIn) { eventType = "IN"; }
    else if (hasIn && !hasOut) { eventType = "OUT"; }
    else if (hasIn && hasOut) { errorMessage = employeeName + " already completed IN and OUT today."; }
    else { errorMessage = "Unexpected attendance state."; }

    if (eventType === null) {
      return { success: false, message: errorMessage || "Attendance not allowed." };
    }

    const now = new Date();
    const timestamp = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd h:mm:ss a");
    const timeOnly = Utilities.formatDate(now, TIMEZONE, "h:mm:ss a");
    let status = "";
    let late = false;
    const hour = parseInt(Utilities.formatDate(now, TIMEZONE, "H"));
    const minute = parseInt(Utilities.formatDate(now, TIMEZONE, "m"));
    const currentTotalMinutes = hour * 60 + minute;

    // Get work schedule for today (default 7:00 AM - 5:00 PM)
    var ws = getWorkScheduleForDate(todayStr);
    var timeInTotalMinutes = ws.timeInHour * 60 + ws.timeInMinute;
    var timeOutTotalMinutes = ws.timeOutHour * 60 + ws.timeOutMinute;

    if (eventType === "IN") {
      if (currentTotalMinutes >= timeInTotalMinutes) { status = "Late"; late = true; }
      else { status = "Early In"; }
    } else {
      if (currentTotalMinutes < timeOutTotalMinutes) { status = "Early Out"; }
      else { status = "On Time Out"; }
    }

    if (dailyRowIndex === -1) {
      if (eventType === "IN") {
        var newRow = dailySheet.getLastRow() + 1;
        var newRange = dailySheet.getRange(newRow, 1, 1, 9);
        newRange.setNumberFormat("@");
        newRange.setValues([[timestamp, uid, employeeName, "IN", timeOnly, status, "", "", ""]]);
      } else {
        return { success: false, message: "No morning IN record found for today." };
      }
    } else {
      if (eventType === "IN") {
        var inRange = dailySheet.getRange(dailyRowIndex, 4, 1, 3);
        inRange.setNumberFormat("@");
        inRange.setValues([["IN", timeOnly, status]]);
      } else {
        var outRange = dailySheet.getRange(dailyRowIndex, 7, 1, 3);
        outRange.setNumberFormat("@");
        outRange.setValues([["OUT", timeOnly, status]]);
      }
    }

    var ampm = hour < 12 ? 'AM' : 'PM';
    const actionText = (eventType === "IN") ? "checked IN (" + ampm + ")" : "checked OUT (" + ampm + ")";

    // Time-based greeting
    var greeting = "";
    if (hour < 12) greeting = "Good Morning";
    else if (hour < 18) greeting = "Good Afternoon";
    else greeting = "Good Evening";

    // Birthday check from 201 FILES sheet (only read columns K, L, O)
    var isBirthday = false;
    try {
      var ss201 = getSpreadsheet();
      var sheet201 = ss201.getSheetByName('201 FILES');
      if (sheet201 && sheet201.getLastRow() >= 2) {
        var lastRow201 = sheet201.getLastRow() - 1;
        var names201 = sheet201.getRange(2, 11, lastRow201, 2).getValues();
        var dobs201 = sheet201.getRange(2, 15, lastRow201, 1).getValues();
        var nameParts = employeeName.toLowerCase().split(/[,\s]+/).filter(function(p){ return p.length > 0; });
        var todayMonth = parseInt(Utilities.formatDate(now, TIMEZONE, "M"));
        var todayDay = parseInt(Utilities.formatDate(now, TIMEZONE, "d"));
        for (var b = 0; b < names201.length; b++) {
          var ln201 = (names201[b][0] || '').toString().trim().toLowerCase();
          var fn201 = (names201[b][1] || '').toString().trim().toLowerCase();
          var nameMatch = (ln201 && nameParts.indexOf(ln201) !== -1) || (fn201 && nameParts.indexOf(fn201) !== -1);
          if (nameMatch) {
            var dobVal = dobs201[b][0];
            if (dobVal) {
              var dob = (dobVal instanceof Date) ? dobVal : new Date(dobVal.toString().trim());
              if (dob && !isNaN(dob.getTime()) && (dob.getMonth() + 1) === todayMonth && dob.getDate() === todayDay) {
                isBirthday = true;
              }
            }
            break;
          }
        }
      }
    } catch(e) {}

    var firstName = employeeName.split(/[,\s]+/)[0] || employeeName;
    var fullMessage = greeting + ", " + firstName + "! " + employeeName + " " + actionText + " at " + timeOnly;
    if (isBirthday) {
      fullMessage = "Happy Birthday, " + firstName + "! " + fullMessage;
    }

    // Queue email + credit check to run async (non-blocking) via PropertiesService
    try {
      var emailData = JSON.stringify({
        email: userDetails.email || '',
        employeeName: employeeName,
        firstName: firstName,
        eventType: eventType,
        ampm: ampm,
        timeOnly: timeOnly,
        status: status,
        todayStr: todayStr,
        timestamp: now.getTime()
      });
      PropertiesService.getScriptProperties().setProperty('PENDING_ATT_EMAIL', emailData);
      // Fire-and-forget: create a 1-second trigger to send email in background
      ScriptApp.newTrigger('processAttendanceEmailQueue').timeBased().after(1000).create();
    } catch(queueErr) {
      // Fallback: send email synchronously if queue fails
      try {
        sendAttendanceEmail_(userDetails.email, employeeName, firstName, eventType, ampm, timeOnly, status, now);
      } catch(e2) {}
      try { checkAndGrantCreditsForDate(todayStr, employeeName); } catch(e3) {}
    }

    return {
      success: true,
      message: fullMessage,
      employeeName: employeeName,
      eventType: (eventType === "IN" ? ampm + " IN" : ampm + " OUT"),
      timeRecorded: timeOnly,
      imageUrl: profileImageB64,
      late: late,
      status: status,
      greeting: greeting,
      isBirthday: isBirthday
    };
  } catch (error) {
    console.error("recordAttendance error:", error);
    return { success: false, message: "Server error: " + error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Background email processor - called by time trigger after recordAttendance returns
function processAttendanceEmailQueue() {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty('PENDING_ATT_EMAIL');
    props.deleteProperty('PENDING_ATT_EMAIL');
    // Clean up this trigger
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'processAttendanceEmailQueue') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    if (!raw) return;
    var d = JSON.parse(raw);
    var now = new Date(d.timestamp);
    sendAttendanceEmail_(d.email, d.employeeName, d.firstName, d.eventType, d.ampm, d.timeOnly, d.status, now);
    try { checkAndGrantCreditsForDate(d.todayStr, d.employeeName); } catch(e) {}
  } catch(err) {
    console.error('processAttendanceEmailQueue error:', err);
  }
}

// Send attendance confirmation email (used by both sync fallback and async queue)
function sendAttendanceEmail_(teacherEmail, employeeName, firstName, eventType, ampm, timeOnly, status, now) {
  if (!teacherEmail) return;
  try {
    var dateFormatted = Utilities.formatDate(now, TIMEZONE, "MMMM d, yyyy");
    var dayOfWeek = Utilities.formatDate(now, TIMEZONE, "EEEE");
    var emailSubject = 'TeachTap - Attendance ' + (eventType === 'IN' ? 'Time In' : 'Time Out') + ' Confirmation';
    var statusColor = (status === 'Late' || status === 'Early Out') ? '#ef4444' : '#22c55e';
    var eventLabel = eventType === 'IN' ? 'Time In (' + ampm + ')' : 'Time Out (' + ampm + ')';
    var emailBody = '<div style="font-family:Poppins,Arial,sans-serif;max-width:480px;margin:0 auto;padding:0;">' +
      '<div style="background:#2563eb;padding:20px 24px;border-radius:12px 12px 0 0;">' +
      '<h2 style="color:#fff;margin:0;font-size:20px;">TeachTap Attendance Confirmation</h2></div>' +
      '<div style="background:#f8fafc;padding:24px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 12px 12px;">' +
      '<p style="color:#334155;font-size:15px;margin:0 0 16px;">Hi <strong>' + firstName + '</strong>,</p>' +
      '<p style="color:#334155;font-size:14px;margin:0 0 16px;">Your attendance has been recorded successfully.</p>' +
      '<table style="width:100%;border-collapse:collapse;margin:0 0 16px;">' +
      '<tr><td style="padding:8px 12px;color:#64748b;font-size:13px;border-bottom:1px solid #e2e8f0;">Name</td>' +
      '<td style="padding:8px 12px;color:#1e293b;font-size:13px;font-weight:600;border-bottom:1px solid #e2e8f0;">' + employeeName + '</td></tr>' +
      '<tr><td style="padding:8px 12px;color:#64748b;font-size:13px;border-bottom:1px solid #e2e8f0;">Date</td>' +
      '<td style="padding:8px 12px;color:#1e293b;font-size:13px;font-weight:600;border-bottom:1px solid #e2e8f0;">' + dayOfWeek + ', ' + dateFormatted + '</td></tr>' +
      '<tr><td style="padding:8px 12px;color:#64748b;font-size:13px;border-bottom:1px solid #e2e8f0;">Event</td>' +
      '<td style="padding:8px 12px;color:#1e293b;font-size:13px;font-weight:600;border-bottom:1px solid #e2e8f0;">' + eventLabel + '</td></tr>' +
      '<tr><td style="padding:8px 12px;color:#64748b;font-size:13px;border-bottom:1px solid #e2e8f0;">Time</td>' +
      '<td style="padding:8px 12px;color:#1e293b;font-size:13px;font-weight:600;border-bottom:1px solid #e2e8f0;">' + timeOnly + '</td></tr>' +
      '<tr><td style="padding:8px 12px;color:#64748b;font-size:13px;">Status</td>' +
      '<td style="padding:8px 12px;font-size:13px;font-weight:600;"><span style="background:' + statusColor + ';color:#fff;padding:2px 10px;border-radius:12px;font-size:12px;">' + status + '</span></td></tr>' +
      '</table>' +
      '<p style="color:#94a3b8;font-size:11px;margin:0;text-align:center;">This is an automated confirmation from TeachTap NFC Attendance System.</p>' +
      '</div></div>';
    MailApp.sendEmail({ to: teacherEmail, subject: emailSubject, htmlBody: emailBody });
  } catch(emailErr) {
    console.error('Attendance email error:', emailErr);
  }
}

// Get user details by NFC UID (columns: A=UID, B=Name, C=Image, D=Email, E=PIN)
function getUserDetailsByUID(uid) {
  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const trimmedUid = uid.toString().trim();

  for (let i = 1; i < data.length; i++) {
    const rowUid = data[i][0] ? data[i][0].toString().trim() : "";
    if (rowUid === trimmedUid) {
      const name = data[i][1] ? data[i][1].toString().trim() : null;
      const imageRaw = data[i][2] ? data[i][2].toString().trim() : "";
      var email = data[i][3] ? data[i][3].toString().trim() : "";
      return { name: name, imageRaw: imageRaw, email: email };
    }
  }
  return null;
}

// Look up teacher by email
function getTeacherByEmail(email) {
  if (!email) return { found: false };
  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const normalizedEmail = email.toString().trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowEmail = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
    if (rowEmail === normalizedEmail) {
      const uid = data[i][0] ? data[i][0].toString().trim() : "";
      const name = data[i][1] ? data[i][1].toString().trim() : "";
      const imageRaw = data[i][2] ? data[i][2].toString().trim() : "";
      var pin = data[i][4] ? data[i][4].toString().trim() : "";
      if (pin && /^\d+$/.test(pin)) pin = pin.padStart(4, '0');
      var imageB64 = "";
      try { imageB64 = getProfileImageBase64(imageRaw); } catch(e) {}
      return {
        found: true,
        uid: uid,
        name: name,
        imageUrl: imageB64,
        hasPin: /^\d{4}$/.test(pin),
        row: i + 1
      };
    }
  }
  return { found: false };
}

// Check registration by email (public wrapper for Teacher.html)
function checkTeacherByEmail(email) {
  if (!email) return { registered: false, email: '' };
  var result = getTeacherByEmail(email);
  if (!result.found) return { registered: false, email: email };
  return {
    registered: true,
    email: email,
    name: result.name,
    hasPin: result.hasPin,
    imageUrl: result.imageUrl
  };
}

// Check current user's teacher status
function checkTeacherStatus() {
  var email = getCurrentEmail();
  if (!email) return { registered: false, email: "" };
  var result = getTeacherByEmail(email);
  if (!result.found) return { registered: false, email: email };
  return {
    registered: true,
    email: email,
    name: result.name,
    hasPin: result.hasPin,
    imageUrl: result.imageUrl
  };
}

// Verify teacher PIN
function verifyTeacherPin(pin) {
  var email = getCurrentEmail();
  if (!email) return { success: false, message: "Unable to get your email." };
  if (!pin) return { success: false, message: "PIN is required." };
  pin = pin.toString().trim();

  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const normalizedEmail = email.toString().trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowEmail = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
    if (rowEmail === normalizedEmail) {
      var storedPin = data[i][4] ? data[i][4].toString().trim() : "";
      if (storedPin && /^\d+$/.test(storedPin)) storedPin = storedPin.padStart(4, '0');
      if (storedPin === pin) {
        var imageRaw = data[i][2] ? data[i][2].toString().trim() : "";
        var imageB64 = "";
        try { imageB64 = getProfileImageBase64(imageRaw); } catch(e) {}
        return {
          success: true,
          name: data[i][1] ? data[i][1].toString().trim() : "",
          uid: data[i][0] ? data[i][0].toString().trim() : "",
          imageUrl: imageB64
        };
      } else {
        return { success: false, message: "Incorrect PIN." };
      }
    }
  }
  return { success: false, message: "Your email is not registered." };
}

// Set teacher PIN (first time setup)
function setTeacherPin(newPin) {
  var email = getCurrentEmail();
  if (!email) return { success: false, message: "Unable to get your email." };
  return setTeacherPinByEmail(email, newPin);
}

// Set PIN by email
function setTeacherPinByEmail(email, newPin) {
  if (!email) return { success: false, message: "Email is required." };
  if (!newPin || !/^\d{4}$/.test(newPin.toString().trim())) {
    return { success: false, message: "PIN must be exactly 4 digits." };
  }
  newPin = newPin.toString().trim();
  var normalizedEmail = email.toString().trim().toLowerCase();

  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowEmail = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
    if (rowEmail === normalizedEmail) {
      var cell = sheet.getRange(i + 1, 5);
      cell.setNumberFormat('@');
      cell.setValue(newPin);
      return { success: true, message: "PIN set successfully." };
    }
  }
  return { success: false, message: "Email is not registered." };
}

// Change teacher PIN (requires old PIN)
function changeTeacherPin(oldPin, newPin) {
  var email = getCurrentEmail();
  if (!email) return { success: false, message: "Unable to get your email." };
  return changeTeacherPinByEmail(email, oldPin, newPin);
}

// Change PIN by email
function changeTeacherPinByEmail(email, oldPin, newPin) {
  if (!email) return { success: false, message: "Email is required." };
  if (!oldPin || !/^\d{4}$/.test(oldPin.toString().trim())) {
    return { success: false, message: "Current PIN must be 4 digits." };
  }
  if (!newPin || !/^\d{4}$/.test(newPin.toString().trim())) {
    return { success: false, message: "New PIN must be exactly 4 digits." };
  }
  oldPin = oldPin.toString().trim();
  newPin = newPin.toString().trim();
  var normalizedEmail = email.toString().trim().toLowerCase();

  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowEmail = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
    if (rowEmail === normalizedEmail) {
      const storedPin = data[i][4] ? data[i][4].toString().trim().padStart(4, '0') : "";
      if (storedPin !== oldPin) {
        return { success: false, message: "Current PIN is incorrect." };
      }
      var cell = sheet.getRange(i + 1, 5);
      cell.setNumberFormat('@');
      cell.setValue(newPin);
      return { success: true, message: "PIN changed successfully." };
    }
  }
  return { success: false, message: "Email is not registered." };
}

// Get teacher attendance for a specific date (by email)
function getMyAttendanceByDate(dateStr) {
  var email = getCurrentEmail();
  if (!email) return null;
  var teacherInfo = getTeacherByEmail(email);
  if (!teacherInfo.found) return null;

  var teacherName = teacherInfo.name;
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(dateStr);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const normalizedName = teacherName.toString().trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowName = row[2] ? row[2].toString().trim().toLowerCase() : "";
    if (rowName === normalizedName) {
      return {
        date: dateStr,
        name: row[2] ? row[2].toString() : "",
        amStatus: row[3] ? row[3].toString() : "",
        amTime: formatTime12h(row[4]),
        amLabel: row[5] ? row[5].toString() : "",
        pmStatus: row[6] ? row[6].toString() : "",
        pmTime: formatTime12h(row[7]),
        pmLabel: row[8] ? row[8].toString() : ""
      };
    }
  }
  return null;
}

// Get teacher monthly summary (by email)
function getMyMonthlySummary(yearMonth) {
  var email = getCurrentEmail();
  if (!email) return [];
  var teacherInfo = getTeacherByEmail(email);
  if (!teacherInfo.found) return [];

  var teacherName = teacherInfo.name;
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const normalizedName = teacherName.toString().trim().toLowerCase();
  const records = [];

  sheets.forEach(function(sheet) {
    const sheetName = sheet.getName();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sheetName)) return;
    if (!sheetName.startsWith(yearMonth)) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowName = row[2] ? row[2].toString().trim().toLowerCase() : "";
      if (rowName === normalizedName) {
        records.push({
          date: sheetName,
          amStatus: row[3] ? row[3].toString() : "",
          amTime: formatTime12h(row[4]),
          amLabel: row[5] ? row[5].toString() : "",
          pmStatus: row[6] ? row[6].toString() : "",
          pmTime: formatTime12h(row[7]),
          pmLabel: row[8] ? row[8].toString() : ""
        });
      }
    }
  });
  records.sort(function(a, b) { return b.date.localeCompare(a.date); });
  return records;
}

// Get admin dashboard data (counts + today's records)
function getAdminDashboardData() {
  const todayStr = getTodayDateString();
  const ss = getSpreadsheet();
  const usersSheet = getOrCreateSheet(USERS_SHEET_NAME);
  const usersData = usersSheet.getDataRange().getValues();
  const totalRegistered = Math.max(0, usersData.length - 1);

  let todaySheet = ss.getSheetByName(todayStr);
  let presentCount = 0, lateCount = 0, outCount = 0;
  let todayRecords = [];

  if (todaySheet) {
    const todayData = todaySheet.getDataRange().getValues();
    for (let i = 1; i < todayData.length; i++) {
      const row = todayData[i];
      if (!row[0] || !isSameDate(row[0], todayStr)) continue;
      const hasIn = row[3] === "IN";
      const hasOut = row[6] === "OUT";
      const amLabel = row[5] ? row[5].toString() : "";
      if (hasIn) presentCount++;
      if (amLabel === "Late") lateCount++;
      if (hasOut) outCount++;
      todayRecords.push({
        name: row[2] ? row[2].toString() : "",
        amTime: formatTime12h(row[4]),
        amLabel: amLabel,
        pmTime: formatTime12h(row[7]),
        pmLabel: row[8] ? row[8].toString() : ""
      });
    }
  }

  // Build present/absent/late/earlyOut name lists
  var presentNames = todayRecords.map(function(r) { return r.name; });
  var lateNames = todayRecords.filter(function(r) { return r.amLabel === 'Late'; }).map(function(r) { return r.name; });
  var earlyOutNames = todayRecords.filter(function(r) { return r.pmLabel === 'Early Out'; }).map(function(r) { return r.name; });
  var earlyOutCount = earlyOutNames.length;
  var allNames = [];
  for (let j = 1; j < usersData.length; j++) {
    var n = usersData[j][1] ? usersData[j][1].toString().trim() : "";
    if (n) allNames.push(n);
  }
  var absentNames = allNames.filter(function(n) { return presentNames.indexOf(n) === -1; });

  return {
    date: todayStr,
    totalRegistered: totalRegistered,
    presentCount: presentCount,
    lateCount: lateCount,
    outCount: outCount,
    earlyOutCount: earlyOutCount,
    absentCount: Math.max(0, totalRegistered - presentCount),
    records: todayRecords,
    presentNames: presentNames,
    absentNames: absentNames,
    lateNames: lateNames,
    earlyOutNames: earlyOutNames
  };
}

// Get attendance records for a specific date
function getAttendanceByDate(dateStr) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(dateStr);
  if (!sheet) return { date: dateStr, records: [] };
  const data = sheet.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    records.push({
      name: row[2] ? row[2].toString() : "",
      amTime: formatTime12h(row[4]),
      amLabel: row[5] ? row[5].toString() : "",
      pmTime: formatTime12h(row[7]),
      pmLabel: row[8] ? row[8].toString() : ""
    });
  }
  return { date: dateStr, records: records };
}

// Get all registered users (includes email)
function getRegisteredUsers() {
  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const uid = data[i][0] ? data[i][0].toString().trim() : "";
    const name = data[i][1] ? data[i][1].toString().trim() : "";
    const imageRaw = data[i][2] ? data[i][2].toString().trim() : "";
    const email = data[i][3] ? data[i][3].toString().trim() : "";
    if (uid) {
      var b64 = "";
      try { b64 = getProfileImageBase64(imageRaw); } catch(e) {}
      users.push({ uid: uid, name: name, imageUrl: b64, imageRaw: imageRaw, email: email, row: i + 1 });
    }
  }
  return users;
}

// Register new NFC user (admin only)
function addRegisteredUser(uid, name, imageUrl, email) {
  const adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: "Unauthorized" };
  if (!uid || !name) return { success: false, message: "UID and Name are required." };
  uid = uid.toString().trim();
  name = name.toString().trim();
  const existing = getUserDetailsByUID(uid);
  if (existing) return { success: false, message: "UID " + uid + " is already registered to " + existing.name };
  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  sheet.appendRow([uid, name, imageUrl || "", email || "", ""]);
  return { success: true, message: name + " registered successfully." };
}

// Delete registered user (admin only)
function deleteRegisteredUser(uid) {
  const email = getCurrentEmail();
  if (!checkIsAdmin(email)) return { success: false, message: "Unauthorized" };
  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const trimmedUid = uid.toString().trim();
  for (let i = 1; i < data.length; i++) {
    const rowUid = data[i][0] ? data[i][0].toString().trim() : "";
    if (rowUid === trimmedUid) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "User removed." };
    }
  }
  return { success: false, message: "UID not found." };
}

// Update registered user (admin only, preserves PIN)
function updateRegisteredUser(originalUid, newUid, newName, newImageUrl, newEmail) {
  const adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: "Unauthorized" };
  if (!newUid || !newName) return { success: false, message: "UID and Name are required." };

  newUid = newUid.toString().trim();
  newName = newName.toString().trim();
  originalUid = originalUid.toString().trim();

  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  if (newUid !== originalUid) {
    for (let i = 1; i < data.length; i++) {
      const rowUid = data[i][0] ? data[i][0].toString().trim() : "";
      if (rowUid === newUid) {
        return { success: false, message: "UID " + newUid + " is already used by " + (data[i][1] || "another user") };
      }
    }
  }

  for (let i = 1; i < data.length; i++) {
    const rowUid = data[i][0] ? data[i][0].toString().trim() : "";
    if (rowUid === originalUid) {
      sheet.getRange(i + 1, 1).setValue(newUid);
      sheet.getRange(i + 1, 2).setValue(newName);
      sheet.getRange(i + 1, 3).setValue(newImageUrl || "");
      sheet.getRange(i + 1, 4).setValue(newEmail || "");
      return { success: true, message: newName + " updated successfully." };
    }
  }
  return { success: false, message: "Original UID not found." };
}

// Reset teacher PIN (admin only)
function resetTeacherPin(uid) {
  const adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: "Unauthorized" };

  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const trimmedUid = uid.toString().trim();

  for (let i = 1; i < data.length; i++) {
    const rowUid = data[i][0] ? data[i][0].toString().trim() : "";
    if (rowUid === trimmedUid) {
      var cell = sheet.getRange(i + 1, 5);
      cell.setNumberFormat('@');
      cell.setValue("");
      var name = data[i][1] ? data[i][1].toString().trim() : "User";
      return { success: true, message: name + "'s PIN has been reset. They will set a new one on next login." };
    }
  }
  return { success: false, message: "UID not found." };
}

// Get all attendance date sheets
function getAttendanceDates() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const dates = [];
  sheets.forEach(function(sheet) {
    const name = sheet.getName();
    if (/^\d{4}-\d{2}-\d{2}$/.test(name)) dates.push(name);
  });
  dates.sort(function(a, b) { return b.localeCompare(a); });
  return dates;
}

// Get teacher record by date and name (for teacher dashboard)
function getTeacherRecordsByDate(dateStr, teacherName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(dateStr);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const normalizedName = teacherName.toString().trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowName = row[2] ? row[2].toString().trim().toLowerCase() : "";
    if (rowName === normalizedName) {
      return {
        date: dateStr,
        name: row[2] ? row[2].toString() : "",
        amStatus: row[3] ? row[3].toString() : "",
        amTime: formatTime12h(row[4]),
        amLabel: row[5] ? row[5].toString() : "",
        pmStatus: row[6] ? row[6].toString() : "",
        pmTime: formatTime12h(row[7]),
        pmLabel: row[8] ? row[8].toString() : ""
      };
    }
  }
  return null;
}

// Get teacher monthly summary by name
function getTeacherMonthlySummary(teacherName, yearMonth) {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const normalizedName = teacherName.toString().trim().toLowerCase();
  const records = [];
  sheets.forEach(function(sheet) {
    const sheetName = sheet.getName();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sheetName)) return;
    if (!sheetName.startsWith(yearMonth)) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowName = row[2] ? row[2].toString().trim().toLowerCase() : "";
      if (rowName === normalizedName) {
        records.push({
          date: sheetName,
          amStatus: row[3] ? row[3].toString() : "",
          amTime: formatTime12h(row[4]),
          amLabel: row[5] ? row[5].toString() : "",
          pmStatus: row[6] ? row[6].toString() : "",
          pmTime: formatTime12h(row[7]),
          pmLabel: row[8] ? row[8].toString() : ""
        });
      }
    }
  });
  records.sort(function(a, b) { return b.date.localeCompare(a.date); });
  return records;
}

// Get today's status for a UID
function getTodayStatusForUID(uid) {
  try {
    const todayStr = getTodayDateString();
    const sheet = getOrCreateDailySheet(todayStr);
    const data = sheet.getDataRange().getValues();
    const trimmedUid = uid.toString().trim();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const ts = row[0];
      if (!ts || !isSameDate(ts, todayStr)) continue;
      const rowUid = row[1] ? row[1].toString().trim() : "";
      if (rowUid !== trimmedUid) continue;
      const hasIn = row[3] === "IN";
      const hasOut = row[6] === "OUT";
      let lastEvent = null;
      if (hasOut) lastEvent = "OUT";
      else if (hasIn) lastEvent = "IN";
      return { exists: true, lastEvent: lastEvent };
    }
    return { exists: false, lastEvent: null };
  } catch (error) {
    return { exists: false, lastEvent: null };
  }
}

function getServerTimezoneHint() { return TIMEZONE; }

// Get or create a named sheet
function getOrCreateSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName, ss.getSheets().length);
    if (sheetName === USERS_SHEET_NAME) {
      sheet.getRange(1, 1, 1, 5).setValues([["Card UID", "Employee Name", "Profile Image", "Email", "PIN"]]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// Get or create daily attendance sheet
function getOrCreateDailySheet(dateString) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(dateString);
  if (!sheet) {
    sheet = ss.insertSheet(dateString, ss.getSheets().length);
    const headers = [["TIMESTAMP", "NFC UID", "FULLNAME", "AM", "AM TIME-IN", "AM STATUS", "PM", "PM TIME-OUT", "PM STATUS"]];
    sheet.getRange(1, 1, 1, 9).setValues(headers);
    sheet.setFrozenRows(1);
    // Set text format for timestamp and time columns to prevent auto-conversion
    var maxRows = sheet.getMaxRows();
    sheet.getRange(1, 1, maxRows, 1).setNumberFormat('@'); // TIMESTAMP col
    sheet.getRange(1, 2, maxRows, 1).setNumberFormat('@'); // NFC UID col
    sheet.getRange(1, 5, maxRows, 1).setNumberFormat('@'); // AM TIME-IN col
    sheet.getRange(1, 8, maxRows, 1).setNumberFormat('@'); // PM TIME-OUT col
    const fullRange = sheet.getRange(1, 1, maxRows, 9);
    fullRange.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#f3f3f3");
  }
  return sheet;
}

function getTodayDateString() {
  return Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
}

function getServerDateTime() {
  var now = new Date();
  return {
    date: Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd"),
    time: Utilities.formatDate(now, TIMEZONE, "h:mm:ss a"),
    dateFormatted: Utilities.formatDate(now, TIMEZONE, "MMMM d, yyyy"),
    dayOfWeek: Utilities.formatDate(now, TIMEZONE, "EEEE")
  };
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// Validate teacher login with email and PIN
function validateTeacherLogin(email, pin) {
  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][3]).trim().toLowerCase();
    var rowPin = String(data[i][4]).trim();
    if (rowPin && /^\d+$/.test(rowPin)) rowPin = rowPin.padStart(4, '0');

    if (rowEmail === email.trim().toLowerCase() && rowPin === pin.toString().trim()) {
      var imageRaw = String(data[i][2]).trim();
      var imageB64 = '';
      if (imageRaw) {
        try { imageB64 = getProfileImageBase64(imageRaw); } catch(e) {}
      }
      return {
        success: true,
        name: String(data[i][1]).trim(),
        uid: String(data[i][0]).trim(),
        imageUrl: imageB64,
        message: 'Login successful'
      };
    }
  }

  // Check if email exists but PIN is wrong
  for (var j = 1; j < data.length; j++) {
    var checkEmail = String(data[j][3]).trim().toLowerCase();
    if (checkEmail === email.trim().toLowerCase()) {
      var checkPin = String(data[j][4]).trim();
      if (!checkPin) {
        return { success: false, message: 'No PIN set for this account. Please contact your administrator.' };
      }
      return { success: false, message: 'Invalid PIN. Please try again.' };
    }
  }

  return { success: false, message: 'Email not found. Please check your email address.' };
}

// Search registered users by name or NFC UID
function searchRegisteredUsers(query) {
  const sheet = getOrCreateSheet(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  var results = [];
  var q = query.trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var uid = String(data[i][0]).trim();
    var name = String(data[i][1]).trim();
    var imageUrl = String(data[i][2]).trim();
    var email = String(data[i][3]).trim();

    if (!q || uid.toLowerCase().indexOf(q) !== -1 || name.toLowerCase().indexOf(q) !== -1) {
      if (imageUrl && imageUrl.includes('drive.google.com/file/d/')) {
        var fileId = imageUrl.match(/\/d\/([^\/]+)/);
        if (fileId) {
          imageUrl = 'https://drive.google.com/uc?id=' + fileId[1] + '&export=view';
        }
      }
      results.push({
        uid: uid,
        name: name,
        imageUrl: imageUrl,
        imageRaw: String(data[i][2]).trim(),
        email: email
      });
    }
  }

  return results;
}

// ===== 201 FILES =====
var FILES_SHEET_NAME = '201 FILES';
var ARCHIVED_FILES_SHEET_NAME = '201 FILES ARCHIVED';

// Column mapping (0-based)
var COL_201 = {
  EDDIS: 0, DISTRICT: 1, SCHOOL: 2, LEVEL: 3, SHS_TRACK_STRAND: 4,
  ITEM_NUMBER: 5, POSITION_TITLE: 6, SG: 7, STEP: 8, EMPLOYEE: 9,
  LASTNAME: 10, FIRSTNAME: 11, MIDDLE_NAME: 12, EXT: 13,
  DATE_OF_BIRTH: 14, DATE_OF_ORIGINAL_APT: 15, DATE_OF_LAST_PROMOTION: 16,
  COURSE_MAJOR_COLLEGE: 17, COURSE_MAJOR_GRADUATE: 18,
  PHILHEALTH: 19, GSIS_BP: 20, PAGIBIG: 21, TIN: 22,
  COMPLETE_ADDRESS: 23, ELIGIBILITY: 24, GENDER: 25
};

// Get all 201 files
function getAll201Files() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found: ' + FILES_SHEET_NAME };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 26);
    var values = dataRange.getValues();
    var teachers = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      teachers.push({
        rowIndex: i + 2,
        eddis: row[COL_201.EDDIS] || '',
        district: row[COL_201.DISTRICT] || '',
        school: row[COL_201.SCHOOL] || '',
        level: row[COL_201.LEVEL] || '',
        shsTrackStrand: row[COL_201.SHS_TRACK_STRAND] || '',
        itemNumber: row[COL_201.ITEM_NUMBER] || '',
        positionTitle: row[COL_201.POSITION_TITLE] || '',
        sg: row[COL_201.SG] || '',
        step: row[COL_201.STEP] || '',
        employee: row[COL_201.EMPLOYEE] || '',
        lastName: row[COL_201.LASTNAME] || '',
        firstName: row[COL_201.FIRSTNAME] || '',
        middleName: row[COL_201.MIDDLE_NAME] || '',
        ext: row[COL_201.EXT] || '',
        dateOfBirth: formatDate201(row[COL_201.DATE_OF_BIRTH]),
        dateOfOriginalApt: formatDate201(row[COL_201.DATE_OF_ORIGINAL_APT]),
        dateOfLastPromotion: formatDate201(row[COL_201.DATE_OF_LAST_PROMOTION]),
        courseMajorCollege: row[COL_201.COURSE_MAJOR_COLLEGE] || '',
        courseMajorGraduate: row[COL_201.COURSE_MAJOR_GRADUATE] || '',
        philhealth: row[COL_201.PHILHEALTH] || '',
        gsisBp: row[COL_201.GSIS_BP] || '',
        pagibig: row[COL_201.PAGIBIG] || '',
        tin: row[COL_201.TIN] || '',
        completeAddress: row[COL_201.COMPLETE_ADDRESS] || '',
        eligibility: row[COL_201.ELIGIBILITY] || '',
        gender: row[COL_201.GENDER] || ''
      });
    }
    return { success: true, data: teachers };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Get single 201 file by row index
function get201File(rowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    var row = sheet.getRange(rowIndex, 1, 1, 26).getValues()[0];
    var teacher = {
      rowIndex: rowIndex,
      eddis: row[COL_201.EDDIS] || '',
      district: row[COL_201.DISTRICT] || '',
      school: row[COL_201.SCHOOL] || '',
      level: row[COL_201.LEVEL] || '',
      shsTrackStrand: row[COL_201.SHS_TRACK_STRAND] || '',
      itemNumber: row[COL_201.ITEM_NUMBER] || '',
      positionTitle: row[COL_201.POSITION_TITLE] || '',
      sg: row[COL_201.SG] || '',
      step: row[COL_201.STEP] || '',
      employee: row[COL_201.EMPLOYEE] || '',
      lastName: row[COL_201.LASTNAME] || '',
      firstName: row[COL_201.FIRSTNAME] || '',
      middleName: row[COL_201.MIDDLE_NAME] || '',
      ext: row[COL_201.EXT] || '',
      dateOfBirth: formatDate201(row[COL_201.DATE_OF_BIRTH]),
      dateOfOriginalApt: formatDate201(row[COL_201.DATE_OF_ORIGINAL_APT]),
      dateOfLastPromotion: formatDate201(row[COL_201.DATE_OF_LAST_PROMOTION]),
      courseMajorCollege: row[COL_201.COURSE_MAJOR_COLLEGE] || '',
      courseMajorGraduate: row[COL_201.COURSE_MAJOR_GRADUATE] || '',
      philhealth: row[COL_201.PHILHEALTH] || '',
      gsisBp: row[COL_201.GSIS_BP] || '',
      pagibig: row[COL_201.PAGIBIG] || '',
      tin: row[COL_201.TIN] || '',
      completeAddress: row[COL_201.COMPLETE_ADDRESS] || '',
      eligibility: row[COL_201.ELIGIBILITY] || '',
      gender: row[COL_201.GENDER] || ''
    };
    return { success: true, data: teacher };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Update 201 file
function update201File(rowIndex, teacherData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    var rowData = [
      teacherData.eddis || '', teacherData.district || '', teacherData.school || '',
      teacherData.level || '', teacherData.shsTrackStrand || '', teacherData.itemNumber || '',
      teacherData.positionTitle || '', teacherData.sg || '', teacherData.step || '',
      teacherData.employee || '', teacherData.lastName || '', teacherData.firstName || '',
      teacherData.middleName || '', teacherData.ext || '',
      teacherData.dateOfBirth || '', teacherData.dateOfOriginalApt || '', teacherData.dateOfLastPromotion || '',
      teacherData.courseMajorCollege || '', teacherData.courseMajorGraduate || '',
      teacherData.philhealth || '', teacherData.gsisBp || '', teacherData.pagibig || '',
      teacherData.tin || '', teacherData.completeAddress || '',
      teacherData.eligibility || '', teacherData.gender || ''
    ];
    sheet.getRange(rowIndex, 1, 1, 26).setValues([rowData]);
    return { success: true, message: 'Teacher 201 file updated successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Add new 201 file
function add201File(teacherData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    var rowData = [
      teacherData.eddis || '', teacherData.district || '', teacherData.school || '',
      teacherData.level || '', teacherData.shsTrackStrand || '', teacherData.itemNumber || '',
      teacherData.positionTitle || '', teacherData.sg || '', teacherData.step || '',
      teacherData.employee || '', teacherData.lastName || '', teacherData.firstName || '',
      teacherData.middleName || '', teacherData.ext || '',
      teacherData.dateOfBirth || '', teacherData.dateOfOriginalApt || '', teacherData.dateOfLastPromotion || '',
      teacherData.courseMajorCollege || '', teacherData.courseMajorGraduate || '',
      teacherData.philhealth || '', teacherData.gsisBp || '', teacherData.pagibig || '',
      teacherData.tin || '', teacherData.completeAddress || '',
      teacherData.eligibility || '', teacherData.gender || ''
    ];
    sheet.appendRow(rowData);
    return { success: true, message: 'New teacher 201 file added successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Archive 201 file (move to archived sheet instead of deleting)
function archive201File(rowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    // Get the row data before removing
    var rowData = sheet.getRange(rowIndex, 1, 1, 26).getValues()[0];
    // Get or create archived sheet
    var archivedSheet = ss.getSheetByName(ARCHIVED_FILES_SHEET_NAME);
    if (!archivedSheet) {
      archivedSheet = ss.insertSheet(ARCHIVED_FILES_SHEET_NAME);
      // Copy headers from main sheet
      var headers = sheet.getRange(1, 1, 1, 26).getValues()[0];
      archivedSheet.getRange(1, 1, 1, 26).setValues([headers]);
    }
    // Append the row to archived sheet
    archivedSheet.appendRow(rowData);
    // Delete from main sheet
    sheet.deleteRow(rowIndex);
    return { success: true, message: 'Teacher 201 file archived successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Restore archived 201 file (move back to main sheet)
function restore201File(rowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var archivedSheet = ss.getSheetByName(ARCHIVED_FILES_SHEET_NAME);
    if (!archivedSheet) return { success: false, error: 'Archived sheet not found' };
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Main sheet not found' };
    // Get the row data
    var rowData = archivedSheet.getRange(rowIndex, 1, 1, 26).getValues()[0];
    // Append to main sheet
    sheet.appendRow(rowData);
    // Delete from archived sheet
    archivedSheet.deleteRow(rowIndex);
    return { success: true, message: 'Teacher 201 file restored successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Get all archived 201 files
function getArchived201Files() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ARCHIVED_FILES_SHEET_NAME);
    if (!sheet) return { success: true, data: [] };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 26);
    var values = dataRange.getValues();
    var teachers = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      teachers.push({
        rowIndex: i + 2,
        eddis: row[COL_201.EDDIS] || '',
        district: row[COL_201.DISTRICT] || '',
        school: row[COL_201.SCHOOL] || '',
        level: row[COL_201.LEVEL] || '',
        shsTrackStrand: row[COL_201.SHS_TRACK_STRAND] || '',
        itemNumber: row[COL_201.ITEM_NUMBER] || '',
        positionTitle: row[COL_201.POSITION_TITLE] || '',
        sg: row[COL_201.SG] || '',
        step: row[COL_201.STEP] || '',
        employee: row[COL_201.EMPLOYEE] || '',
        lastName: row[COL_201.LASTNAME] || '',
        firstName: row[COL_201.FIRSTNAME] || '',
        middleName: row[COL_201.MIDDLE_NAME] || '',
        ext: row[COL_201.EXT] || '',
        dateOfBirth: formatDate201(row[COL_201.DATE_OF_BIRTH]),
        dateOfOriginalApt: formatDate201(row[COL_201.DATE_OF_ORIGINAL_APT]),
        dateOfLastPromotion: formatDate201(row[COL_201.DATE_OF_LAST_PROMOTION]),
        courseMajorCollege: row[COL_201.COURSE_MAJOR_COLLEGE] || '',
        courseMajorGraduate: row[COL_201.COURSE_MAJOR_GRADUATE] || '',
        philhealth: row[COL_201.PHILHEALTH] || '',
        gsisBp: row[COL_201.GSIS_BP] || '',
        pagibig: row[COL_201.PAGIBIG] || '',
        tin: row[COL_201.TIN] || '',
        completeAddress: row[COL_201.COMPLETE_ADDRESS] || '',
        eligibility: row[COL_201.ELIGIBILITY] || '',
        gender: row[COL_201.GENDER] || ''
      });
    }
    return { success: true, data: teachers };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Format date for 201 files
function formatDate201(dateValue) {
  if (!dateValue) return '';
  if (dateValue instanceof Date) {
    var options = { year: 'numeric', month: 'long', day: 'numeric' };
    return dateValue.toLocaleDateString('en-US', options);
  }
  return String(dateValue);
}

// Get teacher 201 file by email (for teacher dashboard)
function getTeacher201ByEmail(email) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };

    // Match email to teacher name via NFC REGISTERED
    var usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
    var teacherName = '';
    if (usersSheet) {
      var usersData = usersSheet.getDataRange().getValues();
      for (var u = 1; u < usersData.length; u++) {
        var userEmail = String(usersData[u][3] || '').trim().toLowerCase();
        if (userEmail === email.trim().toLowerCase()) {
          teacherName = String(usersData[u][1] || '').trim();
          break;
        }
      }
    }

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, error: 'No data found' };

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var lastName = String(row[COL_201.LASTNAME] || '').trim();
      var firstName = String(row[COL_201.FIRSTNAME] || '').trim();
      var fullName = firstName + ' ' + lastName;

      if (teacherName && (
        fullName.toLowerCase().includes(teacherName.toLowerCase()) ||
        teacherName.toLowerCase().includes(lastName.toLowerCase()) && lastName.length > 1
      )) {
        return {
          success: true,
          data: {
            rowIndex: i + 1,
            eddis: row[COL_201.EDDIS] || '',
            district: row[COL_201.DISTRICT] || '',
            school: row[COL_201.SCHOOL] || '',
            level: row[COL_201.LEVEL] || '',
            shsTrackStrand: row[COL_201.SHS_TRACK_STRAND] || '',
            itemNumber: row[COL_201.ITEM_NUMBER] || '',
            positionTitle: row[COL_201.POSITION_TITLE] || '',
            sg: row[COL_201.SG] || '',
            step: row[COL_201.STEP] || '',
            employee: row[COL_201.EMPLOYEE] || '',
            lastName: lastName,
            firstName: firstName,
            middleName: row[COL_201.MIDDLE_NAME] || '',
            ext: row[COL_201.EXT] || '',
            dateOfBirth: formatDate201(row[COL_201.DATE_OF_BIRTH]),
            dateOfOriginalApt: formatDate201(row[COL_201.DATE_OF_ORIGINAL_APT]),
            dateOfLastPromotion: formatDate201(row[COL_201.DATE_OF_LAST_PROMOTION]),
            courseMajorCollege: row[COL_201.COURSE_MAJOR_COLLEGE] || '',
            courseMajorGraduate: row[COL_201.COURSE_MAJOR_GRADUATE] || '',
            philhealth: row[COL_201.PHILHEALTH] || '',
            gsisBp: row[COL_201.GSIS_BP] || '',
            pagibig: row[COL_201.PAGIBIG] || '',
            tin: row[COL_201.TIN] || '',
            completeAddress: row[COL_201.COMPLETE_ADDRESS] || '',
            eligibility: row[COL_201.ELIGIBILITY] || '',
            gender: row[COL_201.GENDER] || ''
          }
        };
      }
    }

    return { success: false, error: 'No 201 file found for this account. Please contact your administrator.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Get profile image by matching name in NFC REGISTERED
function get201ProfileImageByName(lastName, firstName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!usersSheet) return '';
    var data = usersSheet.getDataRange().getValues();
    var searchLast = (lastName || '').toString().trim().toLowerCase();
    var searchFirst = (firstName || '').toString().trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      var regName = (data[i][1] || '').toString().trim().toLowerCase();
      if (searchLast && searchLast.length > 1 && regName.indexOf(searchLast) !== -1) {
        var imageRaw = (data[i][2] || '').toString().trim();
        if (imageRaw) {
          try { return getProfileImageBase64(imageRaw); } catch(e) { return ''; }
        }
        return '';
      }
    }
    if (searchFirst && searchFirst.length > 1) {
      for (var j = 1; j < data.length; j++) {
        var regName2 = (data[j][1] || '').toString().trim().toLowerCase();
        if (regName2.indexOf(searchFirst) !== -1) {
          var imageRaw2 = (data[j][2] || '').toString().trim();
          if (imageRaw2) {
            try { return getProfileImageBase64(imageRaw2); } catch(e) { return ''; }
          }
          return '';
        }
      }
    }
    return '';
  } catch(e) { return ''; }
}

// ===== SERVICE CREDITS =====
var SC_SHEET = 'SERVICE CREDITS';
var SC_HISTORY_SHEET = 'CREDIT HISTORY';
var LR_SHEET = 'LEAVE REQUESTS';
var NOTIF_SHEET = 'NOTIFICATIONS';

// Normalize name for credit matching (handles "LastName, FirstName" vs "FirstName LastName")
function normalizeName(name) {
  if (!name) return '';
  var s = name.toString().trim().toLowerCase().replace(/\s+/g, ' ');
  if (s.indexOf(',') !== -1) {
    var parts = s.split(',');
    s = parts[1].trim() + ' ' + parts[0].trim();
  }
  return s;
}

// Arrange sheet tabs in correct order
function arrangeSheetTabs() {
  var ss = getSpreadsheet();
  var order = [ADMIN_SHEET_NAME, USERS_SHEET_NAME, FILES_SHEET_NAME, SC_SHEET, SC_HISTORY_SHEET, LR_SHEET];
  var pos = 0;
  for (var i = 0; i < order.length; i++) {
    var sheet = ss.getSheetByName(order[i]);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(pos + 1);
      pos++;
    }
  }
}

// Format cell value to readable date string
function formatSCDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var m = val.getMonth() + 1;
    var d = val.getDate();
    var y = val.getFullYear();
    return String(m).padStart(2,'0') + '/' + String(d).padStart(2,'0') + '/' + y;
  }
  var s = val.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var parts = s.split('-');
    return parts[1] + '/' + parts[2] + '/' + parts[0];
  }
  return s;
}

// Ensure service credit sheets exist with headers
function getOrCreateSCSheet(name, headers) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    try { arrangeSheetTabs(); } catch(e) {}
  }
  return sheet;
}

// Normalize name without reversing comma-separated parts
function normalizeNameFlat(name) {
  if (!name) return '';
  return name.toString().trim().toLowerCase().replace(/[,]+/g, ' ').replace(/\s+/g, ' ').trim();
}

// Get all teachers with 201 files and their credit balances, plus NFC registered teachers with credits
function getTeachersWithCredits() {
  var scSheet = getOrCreateSCSheet(SC_SHEET, ['Email', 'TeacherName', 'TotalPoints']);
  var scData = scSheet.getDataRange().getValues();
  // Build credit maps with both normalization strategies
  var creditMap = {};
  var creditMapFlat = {};
  for (var i = 1; i < scData.length; i++) {
    var rawName = (scData[i][1] || '').toString().trim();
    if (!rawName) continue;
    var pts = parseFloat(scData[i][2]) || 0;
    var key = normalizeName(rawName);
    var keyFlat = normalizeNameFlat(rawName);
    if (key) creditMap[key] = (creditMap[key] || 0) + pts;
    if (keyFlat) creditMapFlat[keyFlat] = (creditMapFlat[keyFlat] || 0) + pts;
  }
  function lookupCredits(name) {
    var k1 = normalizeName(name);
    if (creditMap[k1]) return creditMap[k1];
    var k2 = normalizeNameFlat(name);
    if (creditMapFlat[k2]) return creditMapFlat[k2];
    return 0;
  }
  var addedKeys = {};
  var addedFlat = {};
  var teachers = [];
  function isAdded(name) {
    return addedKeys[normalizeName(name)] || addedFlat[normalizeNameFlat(name)];
  }
  function markAdded(name) {
    addedKeys[normalizeName(name)] = true;
    addedFlat[normalizeNameFlat(name)] = true;
  }
  // Add teachers from 201 FILES
  var result = getAll201Files();
  if (result.success && result.data) {
    result.data.forEach(function(t) {
      var fullName = ((t.firstName || '') + ' ' + (t.lastName || '')).trim();
      if (!fullName) return;
      if (isAdded(fullName)) return;
      var credits = lookupCredits(fullName);
      teachers.push({ name: fullName, firstName: t.firstName, lastName: t.lastName, employee: t.employee, position: t.positionTitle, school: t.school, credits: credits });
      markAdded(fullName);
    });
  }
  // Add NFC registered teachers NOT already in list
  try {
    var nfcSheet = getSpreadsheet().getSheetByName(USERS_SHEET_NAME);
    if (nfcSheet && nfcSheet.getLastRow() >= 2) {
      var nfcData = nfcSheet.getDataRange().getValues();
      for (var n = 1; n < nfcData.length; n++) {
        var nfcName = (nfcData[n][1] || '').toString().trim();
        if (!nfcName) continue;
        if (isAdded(nfcName)) continue;
        var credits = lookupCredits(nfcName);
        teachers.push({ name: nfcName, firstName: '', lastName: '', employee: '', position: 'NFC Registered', school: '', credits: credits });
        markAdded(nfcName);
      }
    }
  } catch(e) { console.error('Error adding NFC teachers:', e); }
  // Also add any SERVICE CREDITS entries not yet included
  for (var s = 1; s < scData.length; s++) {
    var scName = (scData[s][1] || '').toString().trim();
    if (!scName) continue;
    if (isAdded(scName)) continue;
    var pts = parseFloat(scData[s][2]) || 0;
    teachers.push({ name: scName, firstName: '', lastName: '', employee: '', position: '', school: '', credits: pts });
    markAdded(scName);
  }
  return { success: true, data: teachers };
}

// Add service credit points to a teacher
function addServiceCredit(teacherName, points, reason, dateStr) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  points = parseFloat(points);
  if (!points || points <= 0) return { success: false, message: 'Points must be greater than 0' };
  if (!teacherName) return { success: false, message: 'Teacher name is required' };

  var scSheet = getOrCreateSCSheet(SC_SHEET, ['Email', 'TeacherName', 'TotalPoints']);
  var scData = scSheet.getDataRange().getValues();
  var found = false;
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < scData.length; i++) {
    if (normalizeName(scData[i][1]) === nameKey) {
      var current = parseFloat(scData[i][2]) || 0;
      scSheet.getRange(i + 1, 3).setValue(current + points);
      found = true;
      break;
    }
  }
  if (!found) {
    scSheet.appendRow(['', teacherName.trim(), points]);
  }

  // Log to history
  var histSheet = getOrCreateSCSheet(SC_HISTORY_SHEET, ['Date', 'TeacherName', 'Points', 'Reason', 'AddedBy', 'Timestamp']);
  histSheet.appendRow([dateStr || getTodayDateString(), teacherName.trim(), points, reason || '', adminEmail, new Date()]);

  // Send email + dashboard notification
  var scData2 = scSheet.getDataRange().getValues();
  var totalNow = 0;
  for (var j = 1; j < scData2.length; j++) {
    if (normalizeName(scData2[j][1]) === nameKey) { totalNow = parseFloat(scData2[j][2]) || 0; break; }
  }
  sendTransactionNotification(
    teacherName.trim(),
    'service_credit',
    'Service Credits Added',
    points + ' service credit point(s) have been added to your account. Reason: ' + (reason || 'N/A') + '. Your new total is ' + totalNow + ' point(s).',
    'TeachTap - Service Credits Added',
    [
      { label: 'Teacher', value: teacherName.trim() },
      { label: 'Points Added', value: '+' + points },
      { label: 'Reason', value: reason || 'N/A' },
      { label: 'New Total', value: totalNow + ' point(s)' },
      { label: 'Added By', value: adminEmail }
    ]
  );

  return { success: true, message: points + ' point(s) added to ' + teacherName };
}

// Get credit history for a specific teacher
function getTeacherCreditHistory(teacherName) {
  var histSheet = getOrCreateSCSheet(SC_HISTORY_SHEET, ['Date', 'TeacherName', 'Points', 'Reason', 'AddedBy', 'Timestamp']);
  var data = histSheet.getDataRange().getValues();
  var history = [];
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < data.length; i++) {
    if (normalizeName(data[i][1]) === nameKey) {
      history.push({
        date: formatSCDate(data[i][0]),
        points: parseFloat(data[i][2]) || 0,
        reason: (data[i][3] || '').toString(),
        addedBy: (data[i][4] || '').toString()
      });
    }
  }
  return history;
}

// Update service credit points (set total) for a teacher
function updateServiceCreditPoints(teacherName, newTotal, reason) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  newTotal = parseFloat(newTotal);
  if (isNaN(newTotal) || newTotal < 0) return { success: false, message: 'Invalid point value' };
  if (!teacherName) return { success: false, message: 'Teacher name is required' };
  if (!reason) return { success: false, message: 'Reason is required' };

  var scSheet = getOrCreateSCSheet(SC_SHEET, ['Email', 'TeacherName', 'TotalPoints']);
  var scData = scSheet.getDataRange().getValues();
  var found = false;
  var nameKey = normalizeName(teacherName);
  var oldTotal = 0;
  for (var i = 1; i < scData.length; i++) {
    if (normalizeName(scData[i][1]) === nameKey) {
      oldTotal = parseFloat(scData[i][2]) || 0;
      scSheet.getRange(i + 1, 3).setValue(newTotal);
      found = true;
      break;
    }
  }
  if (!found) {
    scSheet.appendRow(['', teacherName.trim(), newTotal]);
    oldTotal = 0;
  }

  var diff = newTotal - oldTotal;
  var histSheet = getOrCreateSCSheet(SC_HISTORY_SHEET, ['Date', 'TeacherName', 'Points', 'Reason', 'AddedBy', 'Timestamp']);
  histSheet.appendRow([getTodayDateString(), teacherName.trim(), diff, reason + ' (set to ' + newTotal + ')', adminEmail, new Date()]);

  // Send email + dashboard notification
  sendTransactionNotification(
    teacherName.trim(),
    'service_credit_update',
    'Service Credits Updated',
    'Your service credits have been updated from ' + oldTotal + ' to ' + newTotal + ' point(s). Reason: ' + reason,
    'TeachTap - Service Credits Updated',
    [
      { label: 'Teacher', value: teacherName.trim() },
      { label: 'Previous Total', value: oldTotal + ' point(s)' },
      { label: 'New Total', value: newTotal + ' point(s)' },
      { label: 'Change', value: (diff >= 0 ? '+' : '') + diff + ' point(s)' },
      { label: 'Reason', value: reason },
      { label: 'Updated By', value: adminEmail }
    ]
  );

  return { success: true, message: 'Credits updated to ' + newTotal + ' for ' + teacherName };
}

// Submit a leave request (teacher)
function submitLeaveRequest(teacherName, days, reason, startDate, endDate, returnDate) {
  if (!teacherName || !days || !reason || !startDate || !endDate || !returnDate) {
    return { success: false, message: 'All fields are required' };
  }
  days = parseFloat(days);
  if (days <= 0) return { success: false, message: 'Days must be greater than 0' };

  // Check balance
  var scSheet = getOrCreateSCSheet(SC_SHEET, ['Email', 'TeacherName', 'TotalPoints']);
  var scData = scSheet.getDataRange().getValues();
  var balance = 0;
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < scData.length; i++) {
    if (normalizeName(scData[i][1]) === nameKey) {
      balance = parseFloat(scData[i][2]) || 0;
      break;
    }
  }
  if (days > balance) return { success: false, message: 'Insufficient credits. You have ' + balance + ' day(s) available.' };

  var lrSheet = getOrCreateSCSheet(LR_SHEET, ['RequestID', 'TeacherName', 'Days', 'Reason', 'StartDate', 'EndDate', 'ReturnDate', 'Status', 'AdminComment', 'Timestamp']);
  var reqId = 'LR-' + new Date().getTime();
  lrSheet.appendRow([reqId, teacherName.trim(), days, reason, startDate, endDate, returnDate, 'Pending', '', new Date()]);

  // Send email + dashboard notification for submission confirmation
  sendTransactionNotification(
    teacherName.trim(),
    'leave_request',
    'Leave Request Submitted',
    'Your leave request for ' + days + ' day(s) has been submitted and is pending admin approval. Reason: ' + reason,
    'TeachTap - Leave Request Submitted',
    [
      { label: 'Teacher', value: teacherName.trim() },
      { label: 'Days Requested', value: days + ' day(s)' },
      { label: 'Reason', value: reason },
      { label: 'Start Date', value: startDate },
      { label: 'End Date', value: endDate },
      { label: 'Return Date', value: returnDate },
      { label: 'Status', value: '<span style="background:#f59e0b;color:#fff;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:600">Pending</span>' }
    ]
  );

  return { success: true, message: 'Leave request submitted successfully' };
}

// Get pending leave requests (admin)
function getPendingLeaveRequests() {
  var lrSheet = getOrCreateSCSheet(LR_SHEET, ['RequestID', 'TeacherName', 'Days', 'Reason', 'StartDate', 'EndDate', 'ReturnDate', 'Status', 'AdminComment', 'Timestamp']);
  var data = lrSheet.getDataRange().getValues();
  var requests = [];
  for (var i = 1; i < data.length; i++) {
    requests.push({
      rowIndex: i + 1,
      requestId: (data[i][0] || '').toString(),
      teacherName: (data[i][1] || '').toString(),
      days: parseFloat(data[i][2]) || 0,
      reason: (data[i][3] || '').toString(),
      startDate: formatSCDate(data[i][4]),
      endDate: formatSCDate(data[i][5]),
      returnDate: formatSCDate(data[i][6]),
      status: (data[i][7] || '').toString(),
      comment: (data[i][8] || '').toString(),
      timestamp: formatSCDate(data[i][9])
    });
  }
  return requests;
}

// Approve a leave request (admin)
function approveLeaveRequest(rowIndex, comment) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  if (!comment || !comment.trim()) return { success: false, message: 'Comment is required' };

  var lrSheet = getOrCreateSCSheet(LR_SHEET, ['RequestID', 'TeacherName', 'Days', 'Reason', 'StartDate', 'EndDate', 'ReturnDate', 'Status', 'AdminComment', 'Timestamp']);
  var row = lrSheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
  var teacherName = (row[1] || '').toString().trim();
  var days = parseFloat(row[2]) || 0;

  if ((row[7] || '').toString() !== 'Pending') return { success: false, message: 'Request is no longer pending' };

  // Deduct credits
  var scSheet = getOrCreateSCSheet(SC_SHEET, ['Email', 'TeacherName', 'TotalPoints']);
  var scData = scSheet.getDataRange().getValues();
  var nameKey = normalizeName(teacherName);
  var deducted = false;
  for (var i = 1; i < scData.length; i++) {
    if (normalizeName(scData[i][1]) === nameKey) {
      var current = parseFloat(scData[i][2]) || 0;
      scSheet.getRange(i + 1, 3).setValue(Math.max(0, current - days));
      deducted = true;
      break;
    }
  }

  // Update status
  lrSheet.getRange(rowIndex, 8).setValue('Approved');
  lrSheet.getRange(rowIndex, 9).setValue(comment.trim());

  // Send notification to teacher
  sendCreditRequestNotification(teacherName, 'Approved', comment.trim(), days);

  return { success: true, message: 'Request approved. ' + days + ' day(s) deducted from ' + teacherName };
}

// Decline a leave request (admin)
function declineLeaveRequest(rowIndex, comment) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  if (!comment || !comment.trim()) return { success: false, message: 'Comment is required' };

  var lrSheet = getOrCreateSCSheet(LR_SHEET, ['RequestID', 'TeacherName', 'Days', 'Reason', 'StartDate', 'EndDate', 'ReturnDate', 'Status', 'AdminComment', 'Timestamp']);
  var row = lrSheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
  if ((row[7] || '').toString() !== 'Pending') return { success: false, message: 'Request is no longer pending' };

  lrSheet.getRange(rowIndex, 8).setValue('Declined');
  lrSheet.getRange(rowIndex, 9).setValue(comment.trim());

  // Send notification to teacher
  var teacherName = (row[1] || '').toString().trim();
  var days = parseFloat(row[2]) || 0;
  sendCreditRequestNotification(teacherName, 'Declined', comment.trim(), days);

  return { success: true, message: 'Request declined' };
}

// Get teacher's own leave requests (teacher dashboard)
function getMyLeaveRequests(teacherName) {
  var lrSheet = getOrCreateSCSheet(LR_SHEET, ['RequestID', 'TeacherName', 'Days', 'Reason', 'StartDate', 'EndDate', 'ReturnDate', 'Status', 'AdminComment', 'Timestamp']);
  var data = lrSheet.getDataRange().getValues();
  var requests = [];
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < data.length; i++) {
    if (normalizeName(data[i][1]) === nameKey) {
      requests.push({
        requestId: (data[i][0] || '').toString(),
        days: parseFloat(data[i][2]) || 0,
        reason: (data[i][3] || '').toString(),
        startDate: formatSCDate(data[i][4]),
        endDate: formatSCDate(data[i][5]),
        returnDate: formatSCDate(data[i][6]),
        status: (data[i][7] || '').toString(),
        comment: (data[i][8] || '').toString(),
        timestamp: formatSCDate(data[i][9])
      });
    }
  }
  return requests;
}

// Get teacher's credit balance (teacher dashboard)
function getMyServiceCredits(teacherName) {
  var scSheet = getOrCreateSCSheet(SC_SHEET, ['Email', 'TeacherName', 'TotalPoints']);
  var data = scSheet.getDataRange().getValues();
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < data.length; i++) {
    if (normalizeName(data[i][1]) === nameKey) {
      return parseFloat(data[i][2]) || 0;
    }
  }
  return 0;
}

// ===== NOTIFICATIONS =====

// Look up teacher email by name from NFC REGISTERED sheet
function getTeacherEmailByName(teacherName) {
  if (!teacherName) return '';
  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < data.length; i++) {
    var rowName = normalizeName(data[i][1]);
    if (rowName === nameKey) {
      return (data[i][3] || '').toString().trim();
    }
  }
  return '';
}

// Send email and create dashboard notification when admin approves/declines
function sendCreditRequestNotification(teacherName, status, comment, days) {
  var teacherEmail = getTeacherEmailByName(teacherName);
  var now = new Date();
  var timestamp = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd h:mm:ss a");

  // Store notification for dashboard alert
  var notifSheet = getOrCreateSCSheet(NOTIF_SHEET, ['NotifID', 'TeacherName', 'Type', 'Title', 'Message', 'Read', 'Timestamp']);
  var notifId = 'N-' + now.getTime();
  var title = 'Leave Request ' + status;
  var message = 'Your leave request for ' + days + ' day(s) has been ' + status.toLowerCase() + '. Admin comment: ' + comment;
  notifSheet.appendRow([notifId, teacherName.trim(), 'credit_request', title, message, 'No', now]);

  // Send email notification
  if (teacherEmail) {
    try {
      var subject = 'TeachTap - Leave Request ' + status;
      var statusColor = status === 'Approved' ? '#16a34a' : '#dc2626';
      var statusIcon = status === 'Approved' ? '&#10004;' : '&#10008;';
      var htmlBody = '<div style="font-family:Poppins,Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden">'
        + '<div style="background:linear-gradient(135deg,#1e40af,#3b82f6);padding:24px 28px;text-align:center">'
        + '<h1 style="color:#fff;margin:0;font-size:22px">TeachTap</h1>'
        + '<p style="color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:13px">Attendance System</p></div>'
        + '<div style="padding:28px">'
        + '<div style="text-align:center;margin-bottom:20px">'
        + '<div style="display:inline-block;background:' + statusColor + ';color:#fff;border-radius:50%;width:48px;height:48px;line-height:48px;font-size:24px">' + statusIcon + '</div></div>'
        + '<h2 style="text-align:center;color:#1e293b;margin:0 0 8px;font-size:18px">Leave Request ' + status + '</h2>'
        + '<p style="text-align:center;color:#64748b;margin:0 0 20px;font-size:14px">Your credit request has been reviewed</p>'
        + '<div style="background:#f8fafc;border-radius:8px;padding:16px;margin-bottom:16px">'
        + '<table style="width:100%;font-size:14px;color:#334155">'
        + '<tr><td style="padding:6px 0;font-weight:600">Teacher:</td><td style="padding:6px 0">' + teacherName + '</td></tr>'
        + '<tr><td style="padding:6px 0;font-weight:600">Days Requested:</td><td style="padding:6px 0">' + days + '</td></tr>'
        + '<tr><td style="padding:6px 0;font-weight:600">Status:</td><td style="padding:6px 0"><span style="background:' + statusColor + ';color:#fff;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:600">' + status + '</span></td></tr>'
        + '<tr><td style="padding:6px 0;font-weight:600">Admin Comment:</td><td style="padding:6px 0">' + comment + '</td></tr>'
        + '<tr><td style="padding:6px 0;font-weight:600">Date:</td><td style="padding:6px 0">' + timestamp + '</td></tr>'
        + '</table></div>';
      if (status === 'Approved') {
        htmlBody += '<p style="color:#16a34a;font-size:13px;text-align:center"><strong>' + days + ' day(s)</strong> have been deducted from your service credits.</p>';
      }
      htmlBody += '</div>'
        + '<div style="background:#f1f5f9;padding:16px;text-align:center;font-size:12px;color:#94a3b8">'
        + 'This is an automated message from TeachTap NFC Attendance System</div></div>';

      var plainBody = 'TeachTap - Leave Request ' + status + '\n\n'
        + 'Teacher: ' + teacherName + '\n'
        + 'Days Requested: ' + days + '\n'
        + 'Status: ' + status + '\n'
        + 'Admin Comment: ' + comment + '\n'
        + 'Date: ' + timestamp + '\n';
      if (status === 'Approved') {
        plainBody += '\n' + days + ' day(s) have been deducted from your service credits.';
      }

      GmailApp.sendEmail(teacherEmail, subject, plainBody, { htmlBody: htmlBody });
    } catch (e) {
      // Fallback to MailApp if GmailApp fails
      try {
        MailApp.sendEmail({
          to: teacherEmail,
          subject: 'TeachTap - Leave Request ' + status,
          body: 'Your leave request for ' + days + ' day(s) has been ' + status.toLowerCase() + '. Admin comment: ' + comment
        });
      } catch (e2) {
        console.error('Email send failed:', e2);
      }
    }
  }
}

// Send email + dashboard notification for any transaction
function sendTransactionNotification(teacherName, type, title, message, emailSubject, emailDetails) {
  var teacherEmail = getTeacherEmailByName(teacherName);
  var now = new Date();
  var timestamp = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd h:mm:ss a");

  // Store notification for dashboard alert
  var notifSheet = getOrCreateSCSheet(NOTIF_SHEET, ['NotifID', 'TeacherName', 'Type', 'Title', 'Message', 'Read', 'Timestamp']);
  var notifId = 'N-' + now.getTime();
  notifSheet.appendRow([notifId, teacherName.trim(), type, title, message, 'No', now]);

  // Send email notification
  if (teacherEmail) {
    try {
      var detailRows = '';
      var plainDetails = '';
      for (var k = 0; k < emailDetails.length; k++) {
        detailRows += '<tr><td style="padding:6px 0;font-weight:600">' + emailDetails[k].label + ':</td><td style="padding:6px 0">' + emailDetails[k].value + '</td></tr>';
        plainDetails += emailDetails[k].label + ': ' + emailDetails[k].value + '\n';
      }
      var htmlBody = '<div style="font-family:Poppins,Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden">'
        + '<div style="background:linear-gradient(135deg,#1e40af,#3b82f6);padding:24px 28px;text-align:center">'
        + '<h1 style="color:#fff;margin:0;font-size:22px">TeachTap</h1>'
        + '<p style="color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:13px">Attendance System</p></div>'
        + '<div style="padding:28px">'
        + '<h2 style="text-align:center;color:#1e293b;margin:0 0 16px;font-size:18px">' + title + '</h2>'
        + '<div style="background:#f8fafc;border-radius:8px;padding:16px;margin-bottom:16px">'
        + '<table style="width:100%;font-size:14px;color:#334155">' + detailRows
        + '<tr><td style="padding:6px 0;font-weight:600">Date:</td><td style="padding:6px 0">' + timestamp + '</td></tr>'
        + '</table></div></div>'
        + '<div style="background:#f1f5f9;padding:16px;text-align:center;font-size:12px;color:#94a3b8">'
        + 'This is an automated message from TeachTap NFC Attendance System</div></div>';

      var plainBody = emailSubject + '\n\n' + plainDetails + 'Date: ' + timestamp + '\n';

      GmailApp.sendEmail(teacherEmail, emailSubject, plainBody, { htmlBody: htmlBody });
    } catch (e) {
      try {
        MailApp.sendEmail({ to: teacherEmail, subject: emailSubject, body: message });
      } catch (e2) {
        console.error('Email send failed:', e2);
      }
    }
  }
}

// Get unread notifications for a teacher (dashboard alerts)
function getTeacherNotifications(teacherName) {
  var notifSheet = getOrCreateSCSheet(NOTIF_SHEET, ['NotifID', 'TeacherName', 'Type', 'Title', 'Message', 'Read', 'Timestamp']);
  var data = notifSheet.getDataRange().getValues();
  var notifications = [];
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < data.length; i++) {
    if (normalizeName(data[i][1]) === nameKey) {
      notifications.push({
        notifId: (data[i][0] || '').toString(),
        type: (data[i][2] || '').toString(),
        title: (data[i][3] || '').toString(),
        message: (data[i][4] || '').toString(),
        read: (data[i][5] || '').toString() === 'Yes',
        timestamp: formatSCDate(data[i][6])
      });
    }
  }
  notifications.reverse();
  return notifications;
}

// Mark a notification as read
function markNotificationRead(notifId) {
  var notifSheet = getOrCreateSCSheet(NOTIF_SHEET, ['NotifID', 'TeacherName', 'Type', 'Title', 'Message', 'Read', 'Timestamp']);
  var data = notifSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString() === notifId) {
      notifSheet.getRange(i + 1, 6).setValue('Yes');
      return { success: true };
    }
  }
  return { success: false };
}

// Mark all notifications as read for a teacher
function markAllNotificationsRead(teacherName) {
  var notifSheet = getOrCreateSCSheet(NOTIF_SHEET, ['NotifID', 'TeacherName', 'Type', 'Title', 'Message', 'Read', 'Timestamp']);
  var data = notifSheet.getDataRange().getValues();
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < data.length; i++) {
    if (normalizeName(data[i][1]) === nameKey && (data[i][5] || '').toString() !== 'Yes') {
      notifSheet.getRange(i + 1, 6).setValue('Yes');
    }
  }
  return { success: true };
}

// ===== SCANNER SETTINGS & AUTO CREDIT GRANTING =====
var SCANNER_SETTINGS_SHEET = 'SCANNER SETTINGS';

function getScannerSettings() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SCANNER_SETTINGS_SHEET);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var settings = [];
  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][0];
    var dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, TIMEZONE, 'yyyy-MM-dd');
    } else {
      dateStr = (dateVal || '').toString().trim();
    }
    settings.push({
      rowIndex: i + 1,
      date: dateStr,
      creditAmount: parseFloat(data[i][1]) || 0,
      reason: (data[i][2] || '').toString(),
      createdBy: (data[i][3] || '').toString(),
      createdAt: (data[i][4] || '').toString(),
      active: (data[i][5] || '').toString() === 'Yes'
    });
  }
  return settings;
}

function toggleScannerSetting(rowIndex, active) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SCANNER_SETTINGS_SHEET);
  if (!sheet) return { success: false, message: 'Sheet not found' };
  sheet.getRange(rowIndex, 6).setValue(active ? 'Yes' : 'No');
  
  if (active) {
    var data = sheet.getDataRange().getValues();
    var row = data[rowIndex - 1];
    var dateVal = row[0];
    var dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, TIMEZONE, 'yyyy-MM-dd');
    } else {
      dateStr = (dateVal || '').toString().trim();
    }
    var creditAmount = parseFloat(row[1]) || 0;
    var reason = (row[2] || '').toString();
    try { retroactiveGrantCredits(dateStr, creditAmount, reason); } catch(e) {}
  }
  
  return { success: true, message: active ? 'Setting enabled' : 'Setting disabled' };
}

function deleteScannerSetting(rowIndex) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SCANNER_SETTINGS_SHEET);
  if (!sheet) return { success: false, message: 'Sheet not found' };
  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Setting deleted' };
}

function saveScannerSettings(dateStr, creditAmount, reason) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  if (!dateStr || !creditAmount) return { success: false, message: 'Date and credit amount required' };

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SCANNER_SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SCANNER_SETTINGS_SHEET);
    sheet.appendRow(['Date', 'CreditAmount', 'Reason', 'CreatedBy', 'CreatedAt', 'Active']);
  }
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    var existDate = '';
    if (data[i][0] instanceof Date) {
      existDate = Utilities.formatDate(data[i][0], TIMEZONE, 'yyyy-MM-dd');
    } else {
      existDate = (data[i][0] || '').toString().trim();
    }
    if (existDate === dateStr) {
      sheet.getRange(i + 1, 2).setValue(parseFloat(creditAmount));
      sheet.getRange(i + 1, 3).setValue(reason || '');
      sheet.getRange(i + 1, 4).setValue(adminEmail);
      sheet.getRange(i + 1, 5).setValue(new Date());
      sheet.getRange(i + 1, 6).setValue('Yes');
      found = true;
      break;
    }
  }
  if (!found) {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1).setNumberFormat('@').setValue(dateStr);
    sheet.getRange(newRow, 2).setValue(parseFloat(creditAmount));
    sheet.getRange(newRow, 3).setValue(reason || '');
    sheet.getRange(newRow, 4).setValue(adminEmail);
    sheet.getRange(newRow, 5).setValue(new Date());
    sheet.getRange(newRow, 6).setValue('Yes');
  }

  // Retroactively grant credits for this date
  try {
    retroactiveGrantCredits(dateStr, parseFloat(creditAmount), reason || '');
  } catch(e) { console.error('Retroactive grant error:', e); }

  return { success: true, message: 'Settings saved for ' + dateStr };
}

function checkAndGrantCreditsForDate(dateStr, teacherName) {
  var ss = getSpreadsheet();
  var settingsSheet = ss.getSheetByName(SCANNER_SETTINGS_SHEET);
  if (!settingsSheet || settingsSheet.getLastRow() < 2) return;

  var settingsData = settingsSheet.getDataRange().getValues();
  for (var i = 1; i < settingsData.length; i++) {
    var settingDate = '';
    if (settingsData[i][0] instanceof Date) {
      settingDate = Utilities.formatDate(settingsData[i][0], TIMEZONE, 'yyyy-MM-dd');
    } else {
      settingDate = (settingsData[i][0] || '').toString().trim();
    }
    var active = (settingsData[i][5] || '').toString().trim();
    if (settingDate === dateStr && active === 'Yes') {
      var creditAmount = parseFloat(settingsData[i][1]) || 0;
      var reason = (settingsData[i][2] || '').toString();
      if (creditAmount <= 0) continue;

      // Check if already granted for this date+teacher in CREDIT HISTORY
      var histSheet = getOrCreateSCSheet(SC_HISTORY_SHEET, ['Date', 'TeacherName', 'Points', 'Reason', 'AddedBy', 'Timestamp']);
      var histData = histSheet.getDataRange().getValues();
      var alreadyGranted = false;
      var teacherKey = normalizeName(teacherName);
      var teacherKeyFlat = normalizeNameFlat(teacherName);
      for (var h = 1; h < histData.length; h++) {
        var hDate = '';
        if (histData[h][0] instanceof Date) {
          hDate = Utilities.formatDate(histData[h][0], TIMEZONE, 'yyyy-MM-dd');
        } else {
          hDate = (histData[h][0] || '').toString().trim();
        }
        if (hDate === dateStr && (normalizeName(histData[h][1]) === teacherKey || normalizeNameFlat(histData[h][1]) === teacherKeyFlat)) {
          alreadyGranted = true;
          break;
        }
      }
      if (alreadyGranted) continue;

      // Grant the credit
      var scSheet = getOrCreateSCSheet(SC_SHEET, ['Email', 'TeacherName', 'TotalPoints']);
      var scData = scSheet.getDataRange().getValues();
      var foundSC = false;
      for (var s = 1; s < scData.length; s++) {
        if (normalizeName(scData[s][1]) === teacherKey || normalizeNameFlat(scData[s][1]) === teacherKeyFlat) {
          var current = parseFloat(scData[s][2]) || 0;
          scSheet.getRange(s + 1, 3).setValue(current + creditAmount);
          foundSC = true;
          break;
        }
      }
      if (!foundSC) {
        scSheet.appendRow(['', teacherName.trim(), creditAmount]);
      }
      histSheet.appendRow([dateStr, teacherName.trim(), creditAmount, reason || 'Auto-scan credit', 'System', new Date()]);
      break;
    }
  }
}

function retroactiveGrantCredits(dateStr, creditAmount, reason) {
  var ss = getSpreadsheet();
  var dailySheet = ss.getSheetByName(dateStr);
  if (!dailySheet || dailySheet.getLastRow() < 2) return;

  var dailyData = dailySheet.getDataRange().getValues();
  var processed = {};
  for (var i = 1; i < dailyData.length; i++) {
    var teacherName = (dailyData[i][2] || '').toString().trim();
    var uid = (dailyData[i][1] || '').toString().trim();
    var eventType = (dailyData[i][3] || '').toString().trim();
    if (!teacherName || eventType !== 'IN') continue;
    var nameKey = normalizeName(teacherName);
    if (processed[nameKey]) continue;
    processed[nameKey] = true;
    try {
      checkAndGrantCreditsForDate(uid, teacherName, dateStr);
    } catch(e) {}
  }
}

// ===== WORK SCHEDULE FUNCTIONS =====

// Helper: extract HH:MM string from a cell value (handles Date objects, strings, and numeric fractions)
function extractTimeStr(val) {
  if (val === null || val === undefined || val === '') return '';
  // Handle numeric fractional day (e.g., 0.291667 = 7:00 AM, 0.708333 = 5:00 PM)
  if (typeof val === 'number') {
    if (val >= 0 && val <= 1) {
      var totalMin = Math.round(val * 24 * 60);
      var hh = Math.floor(totalMin / 60);
      var mm = totalMin % 60;
      return (hh < 10 ? '0' : '') + hh + ':' + (mm < 10 ? '0' : '') + mm;
    }
    return val.toString();
  }
  // Handle Date objects (Google Sheets auto-converts time strings to Date)
  if (val instanceof Date) {
    return Utilities.formatDate(val, TIMEZONE, 'HH:mm');
  }
  var s = val.toString().trim();
  // If it looks like HH:MM already, return as-is
  if (/^\d{1,2}:\d{2}$/.test(s)) return s;
  // Try to parse from longer date string
  var m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) return m[1] + ':' + m[2];
  return s;
}

// Get work schedule for a specific date (returns hours/minutes for time in and time out)
// Default: 7:00 AM (7:00) - 5:00 PM (17:00)
function getWorkScheduleForDate(dateStr) {
  var defaults = { timeInHour: 7, timeInMinute: 0, timeOutHour: 17, timeOutMinute: 0 };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(WORK_SCHEDULE_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return defaults;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowDate = '';
      if (data[i][0] instanceof Date) {
        rowDate = Utilities.formatDate(data[i][0], TIMEZONE, 'yyyy-MM-dd');
      } else {
        rowDate = (data[i][0] || '').toString().trim();
      }
      if (rowDate === dateStr) {
        var timeIn = extractTimeStr(data[i][1]);
        var timeOut = extractTimeStr(data[i][2]);
        var inParts = timeIn.split(':');
        var outParts = timeOut.split(':');
        if (inParts.length >= 2 && outParts.length >= 2) {
          return {
            timeInHour: parseInt(inParts[0]) || 7,
            timeInMinute: parseInt(inParts[1]) || 0,
            timeOutHour: parseInt(outParts[0]) || 17,
            timeOutMinute: parseInt(outParts[1]) || 0
          };
        }
      }
    }
  } catch(e) {
    console.error('getWorkScheduleForDate error:', e);
  }
  return defaults;
}

// Get all work schedule settings (for admin UI)
function getWorkScheduleSettings() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(WORK_SCHEDULE_SHEET);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var settings = [];
  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][0];
    var dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, TIMEZONE, 'yyyy-MM-dd');
    } else {
      dateStr = (dateVal || '').toString().trim();
    }
    settings.push({
      rowIndex: i + 1,
      date: dateStr,
      timeIn: extractTimeStr(data[i][1]),
      timeOut: extractTimeStr(data[i][2]),
      createdBy: (data[i][3] || '').toString(),
      createdAt: (data[i][4] || '').toString()
    });
  }
  return settings;
}

// Save work schedule for a specific date
function saveWorkSchedule(dateStr, timeIn, timeOut) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  if (!dateStr || !timeIn || !timeOut) return { success: false, message: 'Date, Time In, and Time Out are required' };

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(WORK_SCHEDULE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(WORK_SCHEDULE_SHEET);
    sheet.appendRow(['Date', 'TimeIn', 'TimeOut', 'CreatedBy', 'CreatedAt']);
    // Set columns A, B, C to plain text to prevent auto-conversion of dates/times
    sheet.getRange(1, 1, sheet.getMaxRows(), 3).setNumberFormat('@');
  }
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    var existDate = '';
    if (data[i][0] instanceof Date) {
      existDate = Utilities.formatDate(data[i][0], TIMEZONE, 'yyyy-MM-dd');
    } else {
      existDate = (data[i][0] || '').toString().trim();
    }
    if (existDate === dateStr) {
      sheet.getRange(i + 1, 1).setNumberFormat('@').setValue(dateStr);
      sheet.getRange(i + 1, 2).setNumberFormat('@').setValue(timeIn);
      sheet.getRange(i + 1, 3).setNumberFormat('@').setValue(timeOut);
      sheet.getRange(i + 1, 4).setValue(adminEmail);
      sheet.getRange(i + 1, 5).setValue(new Date());
      found = true;
      break;
    }
  }
  if (!found) {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1).setNumberFormat('@').setValue(dateStr);
    sheet.getRange(newRow, 2).setNumberFormat('@').setValue(timeIn);
    sheet.getRange(newRow, 3).setNumberFormat('@').setValue(timeOut);
    sheet.getRange(newRow, 4).setValue(adminEmail);
    sheet.getRange(newRow, 5).setValue(new Date());
  }
  return { success: true, message: 'Work schedule saved for ' + dateStr + ' (' + formatTime12hStr(timeIn) + ' - ' + formatTime12hStr(timeOut) + ')' };
}

// Delete work schedule entry
function deleteWorkSchedule(rowIndex) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(WORK_SCHEDULE_SHEET);
  if (!sheet) return { success: false, message: 'Sheet not found' };
  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Work schedule deleted' };
}

// Get today's work schedule (for NFC Scanner footer display)
function getTodayWorkSchedule() {
  var todayStr = getTodayDateString();
  var ws = getWorkScheduleForDate(todayStr);
  return {
    timeIn: formatTime12hStr(padZero(ws.timeInHour) + ':' + padZero(ws.timeInMinute)),
    timeOut: formatTime12hStr(padZero(ws.timeOutHour) + ':' + padZero(ws.timeOutMinute)),
    isCustom: !(ws.timeInHour === 7 && ws.timeInMinute === 0 && ws.timeOutHour === 17 && ws.timeOutMinute === 0)
  };
}

// Format 24h time string (HH:MM) to 12h format (h:MM AM/PM)
function formatTime12hStr(timeStr) {
  if (!timeStr) return '';
  var parts = timeStr.split(':');
  if (parts.length < 2) return timeStr;
  var h = parseInt(parts[0]);
  var m = parseInt(parts[1]);
  var ampm = h >= 12 ? 'PM' : 'AM';
  if (h === 0) h = 12;
  else if (h > 12) h = h - 12;
  return h + ':' + padZero(m) + ' ' + ampm;
}

// Pad number with leading zero
function padZero(n) {
  return (n < 10 ? '0' : '') + n;
}

// ===== UPLOAD ATTENDANCE FEATURE =====

// Find column index by header name (case-insensitive)
function findColumnIndex(headerRow, name) {
  var normalizedName = name.toString().trim().toLowerCase();
  for (var i = 0; i < headerRow.length; i++) {
    if ((headerRow[i] || '').toString().trim().toLowerCase() === normalizedName) return i;
  }
  return -1;
}

// Process attendance upload from client-side parsed data
function processAttendanceUpload(dataRows, dateStr) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized: Admin access required.' };
  if (!dataRows || dataRows.length < 2) return { success: false, message: 'No data rows found in uploaded file.' };
  if (!dateStr) return { success: false, message: 'Date is required.' };

  var header = dataRows[0];
  var nameCol = findColumnIndex(header, 'Employee Name');
  if (nameCol === -1) nameCol = findColumnIndex(header, 'Name');
  if (nameCol === -1) return { success: false, message: 'Column "Employee Name" or "Name" not found in header row.' };

  var amTimeCol = findColumnIndex(header, 'AM Time-In');
  if (amTimeCol === -1) amTimeCol = findColumnIndex(header, 'AM Time');
  var amStatusCol = findColumnIndex(header, 'AM Status');
  var pmTimeCol = findColumnIndex(header, 'PM Time-Out');
  if (pmTimeCol === -1) pmTimeCol = findColumnIndex(header, 'PM Time');
  var pmStatusCol = findColumnIndex(header, 'PM Status');

  var ss = getSpreadsheet();
  var dailySheet = ss.getSheetByName(dateStr);
  if (!dailySheet) {
    dailySheet = ss.insertSheet(dateStr);
    dailySheet.appendRow(['Timestamp', 'UID', 'Employee Name', 'IN', 'AM Time', 'AM Status', 'OUT', 'PM Time', 'PM Status']);
    dailySheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  }

  var usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  var usersData = usersSheet ? usersSheet.getDataRange().getValues() : [];
  function findUidByName(empName) {
    var normalized = empName.toString().trim().toLowerCase();
    for (var i = 1; i < usersData.length; i++) {
      if ((usersData[i][1] || '').toString().trim().toLowerCase() === normalized) {
        return (usersData[i][0] || '').toString().trim();
      }
    }
    return '';
  }

  var existingData = dailySheet.getDataRange().getValues();
  var existingNames = {};
  for (var e = 1; e < existingData.length; e++) {
    var eName = (existingData[e][2] || '').toString().trim().toLowerCase();
    if (eName) existingNames[eName] = e + 1;
  }

  var added = 0, updated = 0, skipped = 0;
  for (var r = 1; r < dataRows.length; r++) {
    var row = dataRows[r];
    var empName = (row[nameCol] || '').toString().trim();
    if (!empName) { skipped++; continue; }

    var amTime = amTimeCol !== -1 ? (row[amTimeCol] || '').toString().trim() : '';
    var amStatus = amStatusCol !== -1 ? (row[amStatusCol] || '').toString().trim() : '';
    var pmTime = pmTimeCol !== -1 ? (row[pmTimeCol] || '').toString().trim() : '';
    var pmStatus = pmStatusCol !== -1 ? (row[pmStatusCol] || '').toString().trim() : '';
    var uid = findUidByName(empName);
    var timestamp = dateStr + ' ' + (amTime || '12:00:00 AM');
    var hasIn = amTime ? 'IN' : '';
    var hasOut = pmTime ? 'OUT' : '';

    var existingRow = existingNames[empName.toLowerCase()];
    if (existingRow) {
      if (hasIn) {
        dailySheet.getRange(existingRow, 4).setValue(hasIn);
        dailySheet.getRange(existingRow, 5).setNumberFormat('@').setValue(amTime);
        dailySheet.getRange(existingRow, 6).setValue(amStatus);
      }
      if (hasOut) {
        dailySheet.getRange(existingRow, 7).setValue(hasOut);
        dailySheet.getRange(existingRow, 8).setNumberFormat('@').setValue(pmTime);
        dailySheet.getRange(existingRow, 9).setValue(pmStatus);
      }
      updated++;
    } else {
      var newRow = dailySheet.getLastRow() + 1;
      var newRange = dailySheet.getRange(newRow, 1, 1, 9);
      newRange.setNumberFormat('@');
      newRange.setValues([[timestamp, uid, empName, hasIn, amTime, amStatus, hasOut, pmTime, pmStatus]]);
      existingNames[empName.toLowerCase()] = newRow;
      added++;
    }
  }

  return {
    success: true,
    message: 'Upload complete for ' + dateStr + ': ' + added + ' added, ' + updated + ' updated, ' + skipped + ' skipped.'
  };
}

// ===== DATE RANGE ATTENDANCE =====

// Get attendance records for a date range (startDate to endDate, inclusive)
function getAttendanceByDateRange(startDate, endDate) {
  var ss = getSpreadsheet();
  var sheets = ss.getSheets();
  var result = [];

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sheetName)) return;
    if (sheetName < startDate || sheetName > endDate) return;

    var data = sheet.getDataRange().getValues();
    var records = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      records.push({
        name: row[2] ? row[2].toString() : '',
        amTime: formatTime12h(row[4]),
        amLabel: row[5] ? row[5].toString() : '',
        pmTime: formatTime12h(row[7]),
        pmLabel: row[8] ? row[8].toString() : ''
      });
    }
    if (records.length > 0) {
      result.push({ date: sheetName, records: records });
    }
  });

  result.sort(function(a, b) { return a.date.localeCompare(b.date); });
  return result;
}

// ===== MONTHLY MONITORING DATA =====

// Get monthly monitoring/analytics data for admin dashboard
function getMonthlyMonitoringData(yearMonth) {
  var ss = getSpreadsheet();
  var sheets = ss.getSheets();
  var usersSheet = getOrCreateSheet(USERS_SHEET_NAME);
  var usersData = usersSheet.getDataRange().getValues();

  var allTeachers = [];
  for (var j = 1; j < usersData.length; j++) {
    var n = usersData[j][1] ? usersData[j][1].toString().trim() : '';
    if (n) allTeachers.push(n);
  }

  var teacherStats = {};
  allTeachers.forEach(function(name) {
    teacherStats[name] = {
      present: 0, absent: 0, late: 0, earlyOut: 0,
      lateDetails: [], earlyOutDetails: [], absentDates: []
    };
  });

  var dailyData = [];
  var workingDays = 0;

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sheetName)) return;
    if (!sheetName.startsWith(yearMonth)) return;

    workingDays++;
    var data = sheet.getDataRange().getValues();
    var presentOnDay = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var name = row[2] ? row[2].toString().trim() : '';
      if (!name) continue;

      presentOnDay[name] = true;

      if (!teacherStats[name]) {
        teacherStats[name] = {
          present: 0, absent: 0, late: 0, earlyOut: 0,
          lateDetails: [], earlyOutDetails: [], absentDates: []
        };
      }

      teacherStats[name].present++;

      var amLabel = row[5] ? row[5].toString() : '';
      var pmLabel = row[8] ? row[8].toString() : '';

      if (amLabel === 'Late') {
        teacherStats[name].late++;
        teacherStats[name].lateDetails.push({
          date: sheetName,
          time: formatTime12h(row[4])
        });
      }
      if (pmLabel === 'Early Out') {
        teacherStats[name].earlyOut++;
        teacherStats[name].earlyOutDetails.push({
          date: sheetName,
          time: formatTime12h(row[7])
        });
      }
    }

    var presentCount = Object.keys(presentOnDay).length;
    var absentCount = 0;
    allTeachers.forEach(function(name) {
      if (!presentOnDay[name]) {
        teacherStats[name].absent++;
        teacherStats[name].absentDates.push(sheetName);
        absentCount++;
      }
    });

    dailyData.push({
      date: sheetName,
      present: presentCount,
      absent: absentCount,
      total: allTeachers.length
    });
  });

  dailyData.sort(function(a, b) { return a.date.localeCompare(b.date); });

  var absentList = [];
  var lateList = [];
  var earlyOutList = [];

  Object.keys(teacherStats).forEach(function(name) {
    var stats = teacherStats[name];
    if (stats.absent > 0) {
      absentList.push({ name: name, count: stats.absent, dates: stats.absentDates });
    }
    if (stats.late > 0) {
      lateList.push({ name: name, count: stats.late, details: stats.lateDetails });
    }
    if (stats.earlyOut > 0) {
      earlyOutList.push({ name: name, count: stats.earlyOut, details: stats.earlyOutDetails });
    }
  });

  absentList.sort(function(a, b) { return b.count - a.count; });
  lateList.sort(function(a, b) { return b.count - a.count; });
  earlyOutList.sort(function(a, b) { return b.count - a.count; });

  var totalAbsences = absentList.reduce(function(sum, item) { return sum + item.count; }, 0);
  var totalLate = lateList.reduce(function(sum, item) { return sum + item.count; }, 0);
  var totalEarlyOut = earlyOutList.reduce(function(sum, item) { return sum + item.count; }, 0);

  return {
    yearMonth: yearMonth,
    totalTeachers: allTeachers.length,
    workingDays: workingDays,
    totalAbsences: totalAbsences,
    totalLate: totalLate,
    totalEarlyOut: totalEarlyOut,
    absentList: absentList,
    lateList: lateList,
    earlyOutList: earlyOutList,
    dailyData: dailyData
  };
}
