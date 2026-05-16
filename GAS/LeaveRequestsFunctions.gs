/**
 * LeaveRequestsFunctions.gs
 * Leave request management: submit, approve, decline, and retrieval.
 */

// Submit a leave request (teacher)
function submitLeaveRequest(teacherName, days, reason, startDate, endDate, returnDate) {
  if (!teacherName || !days || !reason || !startDate || !endDate || !returnDate) {
    return { success: false, message: 'All fields are required' };
  }
  days = parseFloat(days);
  if (days <= 0) return { success: false, message: 'Days must be greater than 0' };

  var scSheet = getOrCreateSCSheet(SC_SHEET, ['Email', 'TeacherName', 'TotalPoints']);
  var scData = scSheet.getDataRange().getValues();
  var balance = 0;
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < scData.length; i++) {
    if (normalizeName(scData[i][1]) === nameKey) {
      balance = parseFloat(scData[i][2]) || 0;
      break;
    }
  }
  if (days > balance) return { success: false, message: 'Insufficient credits. You have ' + balance + ' day(s) available.' };

  var lrSheet = getOrCreateSCSheet(LR_SHEET, ['RequestID', 'TeacherName', 'Days', 'Reason', 'StartDate', 'EndDate', 'ReturnDate', 'Status', 'AdminComment', 'Timestamp']);
  var reqId = 'LR-' + new Date().getTime();
  lrSheet.appendRow([reqId, teacherName.trim(), days, reason, startDate, endDate, returnDate, 'Pending', '', new Date()]);
  return { success: true, message: 'Leave request submitted successfully' };
}

// Get pending leave requests (admin)
function getPendingLeaveRequests() {
  var lrSheet = getOrCreateSCSheet(LR_SHEET, ['RequestID', 'TeacherName', 'Days', 'Reason', 'StartDate', 'EndDate', 'ReturnDate', 'Status', 'AdminComment', 'Timestamp']);
  var data = lrSheet.getDataRange().getValues();
  var requests = [];
  for (var i = 1; i < data.length; i++) {
    requests.push({
      rowIndex: i + 1,
      requestId: (data[i][0] || '').toString(),
      teacherName: (data[i][1] || '').toString(),
      days: parseFloat(data[i][2]) || 0,
      reason: (data[i][3] || '').toString(),
      startDate: formatSCDate(data[i][4]),
      endDate: formatSCDate(data[i][5]),
      returnDate: formatSCDate(data[i][6]),
      status: (data[i][7] || '').toString(),
      comment: (data[i][8] || '').toString(),
      timestamp: formatSCDate(data[i][9])
    });
  }
  return requests;
}

// Approve a leave request (admin)
function approveLeaveRequest(rowIndex, comment) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  if (!comment || !comment.trim()) return { success: false, message: 'Comment is required' };

  var lrSheet = getOrCreateSCSheet(LR_SHEET, ['RequestID', 'TeacherName', 'Days', 'Reason', 'StartDate', 'EndDate', 'ReturnDate', 'Status', 'AdminComment', 'Timestamp']);
  var row = lrSheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
  var teacherName = (row[1] || '').toString().trim();
  var days = parseFloat(row[2]) || 0;

  if ((row[7] || '').toString() !== 'Pending') return { success: false, message: 'Request is no longer pending' };

  var scSheet = getOrCreateSCSheet(SC_SHEET, ['Email', 'TeacherName', 'TotalPoints']);
  var scData = scSheet.getDataRange().getValues();
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < scData.length; i++) {
    if (normalizeName(scData[i][1]) === nameKey) {
      var current = parseFloat(scData[i][2]) || 0;
      scSheet.getRange(i + 1, 3).setValue(Math.max(0, current - days));
      break;
    }
  }

  lrSheet.getRange(rowIndex, 8).setValue('Approved');
  lrSheet.getRange(rowIndex, 9).setValue(comment.trim());

  sendCreditRequestNotification(teacherName, 'Approved', comment.trim(), days);

  return { success: true, message: 'Request approved. ' + days + ' day(s) deducted from ' + teacherName };
}

// Decline a leave request (admin)
function declineLeaveRequest(rowIndex, comment) {
  var adminEmail = getCurrentEmail();
  if (!checkIsAdmin(adminEmail)) return { success: false, message: 'Unauthorized' };
  if (!comment || !comment.trim()) return { success: false, message: 'Comment is required' };

  var lrSheet = getOrCreateSCSheet(LR_SHEET, ['RequestID', 'TeacherName', 'Days', 'Reason', 'StartDate', 'EndDate', 'ReturnDate', 'Status', 'AdminComment', 'Timestamp']);
  var row = lrSheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
  if ((row[7] || '').toString() !== 'Pending') return { success: false, message: 'Request is no longer pending' };

  lrSheet.getRange(rowIndex, 8).setValue('Declined');
  lrSheet.getRange(rowIndex, 9).setValue(comment.trim());

  var teacherName = (row[1] || '').toString().trim();
  var days = parseFloat(row[2]) || 0;
  sendCreditRequestNotification(teacherName, 'Declined', comment.trim(), days);

  return { success: true, message: 'Request declined' };
}

// Get teacher's own leave requests (teacher dashboard)
function getMyLeaveRequests(teacherName) {
  var lrSheet = getOrCreateSCSheet(LR_SHEET, ['RequestID', 'TeacherName', 'Days', 'Reason', 'StartDate', 'EndDate', 'ReturnDate', 'Status', 'AdminComment', 'Timestamp']);
  var data = lrSheet.getDataRange().getValues();
  var requests = [];
  var nameKey = normalizeName(teacherName);
  for (var i = 1; i < data.length; i++) {
    if (normalizeName(data[i][1]) === nameKey) {
      requests.push({
        requestId: (data[i][0] || '').toString(),
        days: parseFloat(data[i][2]) || 0,
        reason: (data[i][3] || '').toString(),
        startDate: formatSCDate(data[i][4]),
        endDate: formatSCDate(data[i][5]),
        returnDate: formatSCDate(data[i][6]),
        status: (data[i][7] || '').toString(),
        comment: (data[i][8] || '').toString(),
        timestamp: formatSCDate(data[i][9])
      });
    }
  }
  return requests;
}
