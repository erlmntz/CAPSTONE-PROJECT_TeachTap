/**
 * Config.gs
 * Global constants, spreadsheet access, and shared utility functions.
 */

const SPREADSHEET_ID = "1oup4Beadx0z_oXGpPzRZn0e2HySoPaFMG7wxLerOcSU";
const USERS_SHEET_NAME = "NFC REGISTERED";
const ADMIN_SHEET_NAME = "ADMIN";
const TIMEZONE = "Asia/Manila";

var FILES_SHEET_NAME = '201 FILES';
var ARCHIVED_FILES_SHEET_NAME = '201 FILES ARCHIVED';
var SC_SHEET = 'SERVICE CREDITS';
var SC_HISTORY_SHEET = 'CREDIT HISTORY';
var LR_SHEET = 'LEAVE REQUESTS';
var NOTIF_SHEET = 'NOTIFICATIONS';
var SCANNER_SETTINGS_SHEET = 'SCANNER SETTINGS';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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

function getServerTimezoneHint() {
  return TIMEZONE;
}

function getLogoutUrl() {
  return ScriptApp.getService().getUrl();
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
  var formattedDate;
  if (cellValue instanceof Date) {
    formattedDate = Utilities.formatDate(cellValue, TIMEZONE, "yyyy-MM-dd");
  } else {
    var str = cellValue.toString();
    var match = str.match(/(\d{4}-\d{2}-\d{2})/);
    if (match) formattedDate = match[1];
    else return false;
  }
  return formattedDate === targetDateStr;
}

// Convert image to base64 data URI
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

// Get or create a named sheet
function getOrCreateSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
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
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(dateString);
  if (!sheet) {
    sheet = ss.insertSheet(dateString, ss.getSheets().length);
    var headers = [["TIMESTAMP", "NFC UID", "FULLNAME", "AM", "AM TIME-IN", "AM STATUS", "PM", "PM TIME-OUT", "PM STATUS"]];
    sheet.getRange(1, 1, 1, 9).setValues(headers);
    sheet.setFrozenRows(1);
    var fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), 9);
    fullRange.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#f3f3f3");
  }
  return sheet;
}

// Normalize name for credit matching
function normalizeName(name) {
  if (!name) return '';
  var s = name.toString().trim().toLowerCase().replace(/\s+/g, ' ');
  if (s.indexOf(',') !== -1) {
    var parts = s.split(',');
    s = parts[1].trim() + ' ' + parts[0].trim();
  }
  return s;
}

// Normalize name without reversing comma-separated parts
function normalizeNameFlat(name) {
  if (!name) return '';
  return name.toString().trim().toLowerCase().replace(/[,]+/g, ' ').replace(/\s+/g, ' ').trim();
}

// Format date for service credits
function formatSCDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var m = val.getMonth() + 1;
    var d = val.getDate();
    var y = val.getFullYear();
    return String(m).padStart(2, '0') + '/' + String(d).padStart(2, '0') + '/' + y;
  }
  var s = val.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var parts = s.split('-');
    return parts[1] + '/' + parts[2] + '/' + parts[0];
  }
  return s;
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

// Ensure service credit sheets exist with headers
function getOrCreateSCSheet(name, headers) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    try { arrangeSheetTabs(); } catch (e) {}
  }
  return sheet;
}

// Arrange sheets in the default order
function arrangeSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var desiredOrder = [
    'ADMIN', 'NFC REGISTERED', '201 FILES', '201 FILES ARCHIVED',
    'SERVICE CREDITS', 'CREDIT HISTORY', 'LEAVE REQUESTS',
    'SCANNER SETTINGS', 'NOTIFICATIONS'
  ];
  for (var i = 0; i < desiredOrder.length; i++) {
    var sheet = ss.getSheetByName(desiredOrder[i]);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1);
    }
  }
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
