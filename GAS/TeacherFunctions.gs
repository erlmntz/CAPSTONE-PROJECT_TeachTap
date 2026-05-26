/**
 * TeacherFunctions.gs
 * Teacher-specific functions: login, PIN, profile, attendance lookup.
 */

// Look up teacher by email
function getTeacherByEmail(email) {
  if (!email) return { found: false };
  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var normalizedEmail = email.toString().trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var rowEmail = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
    if (rowEmail === normalizedEmail) {
      var uid = data[i][0] ? data[i][0].toString().trim() : "";
      var name = data[i][1] ? data[i][1].toString().trim() : "";
      var imageRaw = data[i][2] ? data[i][2].toString().trim() : "";
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

  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var normalizedEmail = email.toString().trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var rowEmail = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
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

  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowEmail = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
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

  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowEmail = data[i][3] ? data[i][3].toString().trim().toLowerCase() : "";
    if (rowEmail === normalizedEmail) {
      var storedPin = data[i][4] ? data[i][4].toString().trim().padStart(4, '0') : "";
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

// Get teacher attendance for a specific date (by email)
function getMyAttendanceByDate(dateStr) {
  var email = getCurrentEmail();
  if (!email) return null;
  var teacherInfo = getTeacherByEmail(email);
  if (!teacherInfo.found) return null;

  var teacherName = teacherInfo.name;
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(dateStr);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var normalizedName = teacherName.toString().trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowName = row[2] ? row[2].toString().trim().toLowerCase() : "";
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
  var ss = getSpreadsheet();
  var sheets = ss.getSheets();
  var normalizedName = teacherName.toString().trim().toLowerCase();
  var records = [];

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sheetName)) return;
    if (!sheetName.startsWith(yearMonth)) return;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowName = row[2] ? row[2].toString().trim().toLowerCase() : "";
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

// Get teacher record by date and name
function getTeacherRecordsByDate(dateStr, teacherName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(dateStr);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var normalizedName = teacherName.toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowName = row[2] ? row[2].toString().trim().toLowerCase() : "";
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
  var ss = getSpreadsheet();
  var sheets = ss.getSheets();
  var normalizedName = teacherName.toString().trim().toLowerCase();
  var records = [];
  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sheetName)) return;
    if (!sheetName.startsWith(yearMonth)) return;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowName = row[2] ? row[2].toString().trim().toLowerCase() : "";
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

// Get teacher 201 file by email (for teacher dashboard)
function getTeacher201ByEmail(email) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };

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
