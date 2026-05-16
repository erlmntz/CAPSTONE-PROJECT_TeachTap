/**
 * ServiceCreditsFunctions.gs
 * Service credits management: add, update, history, and teacher credit lookup.
 */

// Get all teachers with 201 files and their credit balances, plus NFC registered teachers with credits
function getTeachersWithCredits() {
  var scSheet = getOrCreateSCSheet(SC_SHEET, ['Email', 'TeacherName', 'TotalPoints']);
  var scData = scSheet.getDataRange().getValues();
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

  var histSheet = getOrCreateSCSheet(SC_HISTORY_SHEET, ['Date', 'TeacherName', 'Points', 'Reason', 'AddedBy', 'Timestamp']);
  histSheet.appendRow([dateStr || getTodayDateString(), teacherName.trim(), points, reason || '', adminEmail, new Date()]);

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

  return { success: true, message: 'Credits updated to ' + newTotal + ' for ' + teacherName };
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
