/**
 * ScannerSettingsFunctions.gs
 * Scanner settings and auto credit granting.
 */

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
    sheet.appendRow([dateStr, parseFloat(creditAmount), reason || '', adminEmail, new Date(), 'Yes']);
  }

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
      checkAndGrantCreditsForDate(dateStr, teacherName);
    } catch(e) {}
  }
}
