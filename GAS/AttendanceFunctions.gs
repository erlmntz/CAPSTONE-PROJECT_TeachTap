/**
 * AttendanceFunctions.gs
 * Attendance recording, retrieval, and dashboard data.
 */

// Record attendance (IN/OUT) and return scan result with image
function recordAttendance(uid) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    if (!uid || uid.toString().trim() === "") {
      return { success: false, message: "Empty card UID" };
    }
    uid = uid.toString().trim();

    var userDetails = getUserDetailsByUID(uid);
    if (!userDetails) {
      return { success: false, message: "NFC " + uid + " not registered." };
    }
    var employeeName = userDetails.name;
    var profileImageB64 = "";
    if (userDetails.imageRaw) {
      try { profileImageB64 = getProfileImageBase64(userDetails.imageRaw); } catch(e) {}
    }

    var todayStr = getTodayDateString();
    var dailySheet = getOrCreateDailySheet(todayStr);
    var dailyData = dailySheet.getDataRange().getValues();

    var dailyRowIndex = -1;
    var dailyExistingRow = null;
    var duplicateRowsToDelete = [];

    for (var i = 1; i < dailyData.length; i++) {
      var row = dailyData[i];
      var ts = row[0];
      if (!ts || !isSameDate(ts, todayStr)) continue;
      var rowUid = row[1] ? row[1].toString().trim() : "";
      if (rowUid === uid) {
        if (dailyRowIndex !== -1) { duplicateRowsToDelete.push(i + 1); }
        else { dailyRowIndex = i + 1; dailyExistingRow = row; }
      }
    }

    if (duplicateRowsToDelete.length > 0) {
      for (var d = duplicateRowsToDelete.length - 1; d >= 0; d--) {
        dailySheet.deleteRow(duplicateRowsToDelete[d]);
      }
    }

    var hasIn = dailyExistingRow ? dailyExistingRow[3] === "IN" : false;
    var hasOut = dailyExistingRow ? dailyExistingRow[6] === "OUT" : false;
    var eventType = null;
    var errorMessage = "";

    if (!hasIn) { eventType = "IN"; }
    else if (hasIn && !hasOut) { eventType = "OUT"; }
    else if (hasIn && hasOut) { errorMessage = employeeName + " already completed IN and OUT today."; }
    else { errorMessage = "Unexpected attendance state."; }

    if (eventType === null) {
      return { success: false, message: errorMessage || "Attendance not allowed." };
    }

    var now = new Date();
    var timestamp = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd h:mm:ss a");
    var timeOnly = Utilities.formatDate(now, TIMEZONE, "h:mm:ss a");
    var status = "";
    var late = false;
    var hour = parseInt(Utilities.formatDate(now, TIMEZONE, "H"));

    if (eventType === "IN") {
      if (hour >= 7) { status = "Late"; late = true; }
      else { status = "Early In"; }
    } else {
      if (hour < 16) { status = "Early Out"; }
      else { status = "On Time Out"; }
    }

    if (dailyRowIndex === -1) {
      if (eventType === "IN") {
        dailySheet.appendRow([timestamp, uid, employeeName, "IN", timeOnly, status, "", "", ""]);
        var newLastRow = dailySheet.getLastRow();
        dailySheet.getRange(newLastRow, 1).setNumberFormat("@");
        dailySheet.getRange(newLastRow, 5).setNumberFormat("@");
      } else {
        return { success: false, message: "No morning IN record found for today." };
      }
    } else {
      if (eventType === "IN") {
        dailySheet.getRange(dailyRowIndex, 4).setValue("IN");
        dailySheet.getRange(dailyRowIndex, 5).setNumberFormat("@").setValue(timeOnly);
        dailySheet.getRange(dailyRowIndex, 6).setValue(status);
      } else {
        dailySheet.getRange(dailyRowIndex, 7).setValue("OUT");
        dailySheet.getRange(dailyRowIndex, 8).setNumberFormat("@").setValue(timeOnly);
        dailySheet.getRange(dailyRowIndex, 9).setValue(status);
      }
    }

    var actionText = (eventType === "IN") ? "checked IN (AM)" : "checked OUT (PM)";

    var greeting = "";
    if (hour < 12) greeting = "Good Morning";
    else if (hour < 18) greeting = "Good Afternoon";
    else greeting = "Good Evening";

    // Birthday check from 201 FILES sheet
    var isBirthday = false;
    try {
      var ss201 = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet201 = ss201.getSheetByName('201 FILES');
      if (sheet201 && sheet201.getLastRow() >= 2) {
        var all201 = sheet201.getRange(2, 1, sheet201.getLastRow() - 1, 26).getValues();
        var nameParts = employeeName.toLowerCase().split(/[,\s]+/).filter(function(p){ return p.length > 0; });
        for (var b = 0; b < all201.length; b++) {
          var ln201 = (all201[b][10] || '').toString().trim().toLowerCase();
          var fn201 = (all201[b][11] || '').toString().trim().toLowerCase();
          var nameMatch = false;
          if (ln201 && nameParts.indexOf(ln201) !== -1) nameMatch = true;
          if (!nameMatch && fn201 && nameParts.indexOf(fn201) !== -1) nameMatch = true;
          if (nameMatch) {
            var dobVal = all201[b][14];
            if (dobVal) {
              var dob = (dobVal instanceof Date) ? dobVal : new Date(dobVal.toString().trim());
              if (dob && !isNaN(dob.getTime())) {
                var todayMonth = parseInt(Utilities.formatDate(now, TIMEZONE, "M"));
                var todayDay = parseInt(Utilities.formatDate(now, TIMEZONE, "d"));
                if ((dob.getMonth() + 1) === todayMonth && dob.getDate() === todayDay) {
                  isBirthday = true;
                }
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

    // Send email confirmation to teacher
    try {
      var teacherEmail = userDetails.email;
      if (teacherEmail) {
        var dateFormatted = Utilities.formatDate(now, TIMEZONE, "MMMM d, yyyy");
        var dayOfWeek = Utilities.formatDate(now, TIMEZONE, "EEEE");
        var emailSubject = 'TeachTap - Attendance ' + (eventType === 'IN' ? 'Time In' : 'Time Out') + ' Confirmation';
        var statusColor = (status === 'Late' || status === 'Early Out') ? '#ef4444' : '#22c55e';
        var eventLabel = eventType === 'IN' ? 'Time In (AM)' : 'Time Out (PM)';
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
      }
    } catch(emailErr) {
      console.error('Attendance email error:', emailErr);
    }

    // Auto-grant service credits based on SCANNER SETTINGS
    try {
      checkAndGrantCreditsForDate(todayStr, employeeName);
    } catch(creditErr) {
      console.error('Auto-credit error:', creditErr);
    }

    return {
      success: true,
      message: fullMessage,
      employeeName: employeeName,
      eventType: eventType === "IN" ? "AM IN" : "PM OUT",
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

// Get user details by NFC UID
function getUserDetailsByUID(uid) {
  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var trimmedUid = uid.toString().trim();

  for (var i = 1; i < data.length; i++) {
    var rowUid = data[i][0] ? data[i][0].toString().trim() : "";
    if (rowUid === trimmedUid) {
      var name = data[i][1] ? data[i][1].toString().trim() : null;
      var imageRaw = data[i][2] ? data[i][2].toString().trim() : "";
      var email = data[i][3] ? data[i][3].toString().trim() : "";
      return { name: name, imageRaw: imageRaw, email: email };
    }
  }
  return null;
}

// Get admin dashboard data (counts + today's records)
function getAdminDashboardData() {
  var todayStr = getTodayDateString();
  var ss = getSpreadsheet();
  var usersSheet = getOrCreateSheet(USERS_SHEET_NAME);
  var usersData = usersSheet.getDataRange().getValues();
  var totalRegistered = Math.max(0, usersData.length - 1);

  var todaySheet = ss.getSheetByName(todayStr);
  var presentCount = 0, lateCount = 0, outCount = 0;
  var todayRecords = [];

  if (todaySheet) {
    var todayData = todaySheet.getDataRange().getValues();
    for (var i = 1; i < todayData.length; i++) {
      var row = todayData[i];
      if (!row[0] || !isSameDate(row[0], todayStr)) continue;
      var hasIn = row[3] === "IN";
      var hasOut = row[6] === "OUT";
      var amLabel = row[5] ? row[5].toString() : "";
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

  var presentNames = todayRecords.map(function(r) { return r.name; });
  var lateNames = todayRecords.filter(function(r) { return r.amLabel === 'Late'; }).map(function(r) { return r.name; });
  var earlyOutNames = todayRecords.filter(function(r) { return r.pmLabel === 'Early Out'; }).map(function(r) { return r.name; });
  var earlyOutCount = earlyOutNames.length;
  var allNames = [];
  for (var j = 1; j < usersData.length; j++) {
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
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(dateStr);
  if (!sheet) return { date: dateStr, records: [] };
  var data = sheet.getDataRange().getValues();
  var records = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
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

// Get all attendance date sheets
function getAttendanceDates() {
  var ss = getSpreadsheet();
  var sheets = ss.getSheets();
  var dates = [];
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (/^\d{4}-\d{2}-\d{2}$/.test(name)) dates.push(name);
  });
  dates.sort(function(a, b) { return b.localeCompare(a); });
  return dates;
}

// Get today's status for a UID
function getTodayStatusForUID(uid) {
  try {
    var todayStr = getTodayDateString();
    var sheet = getOrCreateDailySheet(todayStr);
    var data = sheet.getDataRange().getValues();
    var trimmedUid = uid.toString().trim();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var ts = row[0];
      if (!ts || !isSameDate(ts, todayStr)) continue;
      var rowUid = row[1] ? row[1].toString().trim() : "";
      if (rowUid !== trimmedUid) continue;
      var hasIn = row[3] === "IN";
      var hasOut = row[6] === "OUT";
      var lastEvent = null;
      if (hasOut) lastEvent = "OUT";
      else if (hasIn) lastEvent = "IN";
      return { exists: true, lastEvent: lastEvent };
    }
    return { exists: false, lastEvent: null };
  } catch (error) {
    return { exists: false, lastEvent: null };
  }
}

// Remove duplicate IN records across daily sheets
function removeDuplicateInRecords() {
  var ss = getSpreadsheet();
  var sheets = ss.getSheets();
  var totalDeleted = 0;
  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sheetName)) return;
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    var seenUids = {};
    var rowsToDelete = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var ts = row[0];
      if (!ts || !isSameDate(ts, sheetName)) continue;
      var uid = row[1] ? row[1].toString().trim() : "";
      if (!uid) continue;
      if (row[3] === "IN") {
        if (seenUids[uid]) { rowsToDelete.push(i + 1); }
        else { seenUids[uid] = true; }
      }
    }
    for (var r = rowsToDelete.length - 1; r >= 0; r--) { sheet.deleteRow(rowsToDelete[r]); }
    if (rowsToDelete.length > 0) totalDeleted += rowsToDelete.length;
  });
  return "Done. Deleted " + totalDeleted + " duplicate IN record(s).";
}
