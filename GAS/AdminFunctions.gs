/**
 * AdminFunctions.gs
 * Admin-specific functions: settings, PIN management, admin list management.
 */

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
      var storedPin = (data[i][1] || '').toString().trim();
      while (storedPin.length < 6) storedPin = '0' + storedPin;
      var enteredPin = currentPin.toString().trim();
      while (enteredPin.length < 6) enteredPin = '0' + enteredPin;
      if (storedPin !== enteredPin) {
        return { success: false, message: 'Current PIN is incorrect.' };
      }
      if (newEmail && newEmail.trim()) {
        sheet.getRange(i + 1, 1).setValue(newEmail.trim());
      }
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

// Get list of all admins
function getAdminList() {
  var email = getCurrentEmail();
  if (!checkIsAdmin(email)) return { success: false, message: 'Unauthorized' };
  var sheet = getSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return { success: false, message: 'Admin sheet not found.' };
  var data = sheet.getDataRange().getValues();
  var admins = [];
  for (var i = 1; i < data.length; i++) {
    var adminEmail = (data[i][0] || '').toString().trim();
    if (adminEmail) {
      admins.push({
        email: adminEmail,
        row: i + 1
      });
    }
  }
  return { success: true, admins: admins };
}

// Add a new admin (email + 6-digit PIN)
function addNewAdmin(newAdminEmail, newAdminPin) {
  var email = getCurrentEmail();
  if (!checkIsAdmin(email)) return { success: false, message: 'Unauthorized' };
  if (!newAdminEmail || !newAdminEmail.trim()) return { success: false, message: 'Email is required.' };
  if (!newAdminPin || !/^\d{6}$/.test(newAdminPin.toString().trim())) {
    return { success: false, message: 'PIN must be exactly 6 digits.' };
  }

  var normalizedNew = newAdminEmail.toString().trim().toLowerCase();
  if (checkIsAdmin(normalizedNew)) {
    return { success: false, message: 'This email is already an admin.' };
  }

  var sheet = getSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return { success: false, message: 'Admin sheet not found.' };

  var pinCell = sheet.getRange(sheet.getLastRow() + 1, 2);
  sheet.appendRow([newAdminEmail.trim(), newAdminPin.toString().trim()]);
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 2).setNumberFormat('@');
  sheet.getRange(lastRow, 2).setValue(newAdminPin.toString().trim());

  return { success: true, message: newAdminEmail.trim() + ' has been added as admin.' };
}

// Remove an admin by row index
function removeAdmin(rowIndex) {
  var email = getCurrentEmail();
  if (!checkIsAdmin(email)) return { success: false, message: 'Unauthorized' };

  var sheet = getSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return { success: false, message: 'Admin sheet not found.' };

  var data = sheet.getDataRange().getValues();
  if (rowIndex < 2 || rowIndex > data.length) {
    return { success: false, message: 'Invalid row index.' };
  }

  var removedEmail = (data[rowIndex - 1][0] || '').toString().trim().toLowerCase();
  var currentEmail = email.toString().trim().toLowerCase();

  if (removedEmail === currentEmail) {
    return { success: false, message: 'You cannot remove yourself as admin.' };
  }

  var adminCount = 0;
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim()) adminCount++;
  }
  if (adminCount <= 1) {
    return { success: false, message: 'Cannot remove the last admin.' };
  }

  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Admin removed successfully.' };
}
