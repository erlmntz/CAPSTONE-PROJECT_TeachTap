/**
 * Auth.gs
 * Authentication and authorization functions.
 */

// Get current user email via identity token (works with "Execute as: Me")
function getCurrentEmail() {
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

  try {
    var email = Session.getActiveUser().getEmail();
    if (email) return email.toString().trim().toLowerCase();
  } catch(e) {}

  return "";
}

// Check if email exists in the ADMIN sheet
function checkIsAdmin(email) {
  if (!email) return false;
  var sheet = getSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return false;
  var data = sheet.getDataRange().getValues();
  var normalizedEmail = email.toString().trim().toLowerCase();
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      var cellVal = data[i][j] ? data[i][j].toString().trim().toLowerCase() : "";
      if (cellVal === normalizedEmail) return true;
    }
  }
  return false;
}

// Get current user role
function getCurrentUserRole() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch(e) {}
  return { email: email };
}
