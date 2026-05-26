/**
 * Code.gs
 * Main entry point — serves the correct page based on user role.
 *
 * All backend logic is organized into separate files:
 *   Config.gs              — Constants, spreadsheet access, utility helpers
 *   Auth.gs                — Authentication and authorization
 *   AdminFunctions.gs      — Admin settings, PIN, admin list management
 *   TeacherFunctions.gs    — Teacher login, PIN, profile, attendance lookup
 *   AttendanceFunctions.gs — Attendance recording and dashboard data
 *   UserManagement.gs      — NFC user registration, updates, search
 *   Files201Functions.gs   — 201 Files CRUD and archiving
 *   ServiceCreditsFunctions.gs — Service credits management
 *   LeaveRequestsFunctions.gs  — Leave request submit/approve/decline
 *   NotificationsFunctions.gs  — Notifications and email alerts
 *   ScannerSettingsFunctions.gs — Scanner settings and auto credit granting
 */

// Serve the correct page based on user role
function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : null;

  if (page === "scanner") {
    return HtmlService.createTemplateFromFile('Nfc_Scanner')
      .evaluate()
      .setTitle("TeachTap - NFC Scanner")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  var email = getCurrentEmail();
  var isAdmin = email && checkIsAdmin(email);

  if (isAdmin) {
    return HtmlService.createTemplateFromFile('Admin')
      .evaluate()
      .setTitle("TeachTap - Admin Dashboard")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  } else {
    return HtmlService.createTemplateFromFile('Teacher')
      .evaluate()
      .setTitle("TeachTap - Teacher Dashboard")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }
}

// Include external HTML/CSS files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
