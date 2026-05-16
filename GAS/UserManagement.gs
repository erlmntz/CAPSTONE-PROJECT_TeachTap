/**
 * UserManagement.gs
 * User registration, deletion, updates, PIN reset, and search.
 */

// Get all registered users (includes email)
function getRegisteredUsers() {
  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var uid = data[i][0] ? data[i][0].toString().trim() : "";
    var name = data[i][1] ? data[i][1].toString().trim() : "";
    var imageRaw = data[i][2] ? data[i][2].toString().trim() : "";
    var email = data[i][3] ? data[i][3].toString().trim() : "";
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
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: "Unauthorized" };
  if (!uid || !name) return { success: false, message: "UID and Name are required." };
  uid = uid.toString().trim();
  name = name.toString().trim();
  var existing = getUserDetailsByUID(uid);
  if (existing) return { success: false, message: "UID " + uid + " is already registered to " + existing.name };
  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  sheet.appendRow([uid, name, imageUrl || "", email || "", ""]);
  return { success: true, message: name + " registered successfully." };
}

// Delete registered user (admin only)
function deleteRegisteredUser(uid) {
  var email = getCurrentEmail();
  if (!checkIsAdmin(email)) return { success: false, message: "Unauthorized" };
  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var trimmedUid = uid.toString().trim();
  for (var i = 1; i < data.length; i++) {
    var rowUid = data[i][0] ? data[i][0].toString().trim() : "";
    if (rowUid === trimmedUid) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "User removed." };
    }
  }
  return { success: false, message: "UID not found." };
}

// Update registered user (admin only, preserves PIN)
function updateRegisteredUser(originalUid, newUid, newName, newImageUrl, newEmail) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: "Unauthorized" };
  if (!newUid || !newName) return { success: false, message: "UID and Name are required." };

  newUid = newUid.toString().trim();
  newName = newName.toString().trim();
  originalUid = originalUid.toString().trim();

  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  if (newUid !== originalUid) {
    for (var i = 1; i < data.length; i++) {
      var rowUid = data[i][0] ? data[i][0].toString().trim() : "";
      if (rowUid === newUid) {
        return { success: false, message: "UID " + newUid + " is already used by " + (data[i][1] || "another user") };
      }
    }
  }

  for (var i = 1; i < data.length; i++) {
    var rowUid = data[i][0] ? data[i][0].toString().trim() : "";
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
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: "Unauthorized" };

  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var trimmedUid = uid.toString().trim();

  for (var i = 1; i < data.length; i++) {
    var rowUid = data[i][0] ? data[i][0].toString().trim() : "";
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

// Search registered users by name, UID, or email
function searchRegisteredUsers(query) {
  if (!query) return [];
  var sheet = getOrCreateSheet(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var q = query.toString().trim().toLowerCase();
  var results = [];

  for (var i = 1; i < data.length; i++) {
    var uid = (data[i][0] || '').toString().trim();
    var name = (data[i][1] || '').toString().trim();
    var email = (data[i][3] || '').toString().trim();

    if (uid.toLowerCase().indexOf(q) !== -1 ||
        name.toLowerCase().indexOf(q) !== -1 ||
        email.toLowerCase().indexOf(q) !== -1) {
      var imageRaw = (data[i][2] || '').toString().trim();
      var imageUrl = '';
      try { imageUrl = getProfileImageBase64(imageRaw); } catch(e) {}
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
