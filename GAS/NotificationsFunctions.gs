/**
 * NotificationsFunctions.gs
 * Notification management: send, retrieve, and mark as read.
 */

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

  var notifSheet = getOrCreateSCSheet(NOTIF_SHEET, ['NotifID', 'TeacherName', 'Type', 'Title', 'Message', 'Read', 'Timestamp']);
  var notifId = 'N-' + now.getTime();
  var title = 'Leave Request ' + status;
  var message = 'Your leave request for ' + days + ' day(s) has been ' + status.toLowerCase() + '. Admin comment: ' + comment;
  notifSheet.appendRow([notifId, teacherName.trim(), 'credit_request', title, message, 'No', now]);

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
