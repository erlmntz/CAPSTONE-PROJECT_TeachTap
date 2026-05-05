// ==================== CONFIGURATION ====================
var SPREADSHEET_ID = '1oup4Beadx0z_oXGpPzRZn0e2HySoPaFMG7wxLerOcSU';
var SHEET_NAME = '201 FILES';

// Column mapping (0-based index for the array)
var COL = {
  EDDIS: 0,
  DISTRICT: 1,
  SCHOOL: 2,
  LEVEL: 3,
  SHS_TRACK_STRAND: 4,
  ITEM_NUMBER: 5,
  POSITION_TITLE: 6,
  SG: 7,
  STEP: 8,
  EMPLOYEE: 9,
  LASTNAME: 10,
  FIRSTNAME: 11,
  MIDDLE_NAME: 12,
  EXT: 13,
  DATE_OF_BIRTH: 14,
  DATE_OF_ORIGINAL_APT: 15,
  DATE_OF_LAST_PROMOTION: 16,
  COURSE_MAJOR_COLLEGE: 17,
  COURSE_MAJOR_GRADUATE: 18,
  PHILHEALTH: 19,
  GSIS_BP: 20,
  PAGIBIG: 21,
  TIN: 22,
  COMPLETE_ADDRESS: 23,
  ELIGIBILITY: 24,
  GENDER: 25
};

// ==================== WEB APP ENTRY POINT ====================
function doGet(e) {
  var page = e.parameter.page || 'landing';
  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('HTML/admin')
      .setTitle('TeachTap - Admin Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('HTML/dashboard')
      .setTitle('TeachTap | K-6 Staff Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('HTML/landing')
    .setTitle('TeachTap - Maguinao Elementary School')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== 201 FILES - GET ALL TEACHERS ====================
function getAll201Files() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Sheet not found: ' + SHEET_NAME };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: [] };
    }
    
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 26);
    var values = dataRange.getValues();
    
    var teachers = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      teachers.push({
        rowIndex: i + 2, // actual row number in sheet (1-based, skip header)
        eddis: row[COL.EDDIS] || '',
        district: row[COL.DISTRICT] || '',
        school: row[COL.SCHOOL] || '',
        level: row[COL.LEVEL] || '',
        shsTrackStrand: row[COL.SHS_TRACK_STRAND] || '',
        itemNumber: row[COL.ITEM_NUMBER] || '',
        positionTitle: row[COL.POSITION_TITLE] || '',
        sg: row[COL.SG] || '',
        step: row[COL.STEP] || '',
        employee: row[COL.EMPLOYEE] || '',
        lastName: row[COL.LASTNAME] || '',
        firstName: row[COL.FIRSTNAME] || '',
        middleName: row[COL.MIDDLE_NAME] || '',
        ext: row[COL.EXT] || '',
        dateOfBirth: formatDate(row[COL.DATE_OF_BIRTH]),
        dateOfOriginalApt: formatDate(row[COL.DATE_OF_ORIGINAL_APT]),
        dateOfLastPromotion: formatDate(row[COL.DATE_OF_LAST_PROMOTION]),
        courseMajorCollege: row[COL.COURSE_MAJOR_COLLEGE] || '',
        courseMajorGraduate: row[COL.COURSE_MAJOR_GRADUATE] || '',
        philhealth: row[COL.PHILHEALTH] || '',
        gsisBp: row[COL.GSIS_BP] || '',
        pagibig: row[COL.PAGIBIG] || '',
        tin: row[COL.TIN] || '',
        completeAddress: row[COL.COMPLETE_ADDRESS] || '',
        eligibility: row[COL.ELIGIBILITY] || '',
        gender: row[COL.GENDER] || ''
      });
    }
    
    return { success: true, data: teachers };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==================== 201 FILES - GET SINGLE TEACHER ====================
function get201File(rowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Sheet not found: ' + SHEET_NAME };
    }
    
    var row = sheet.getRange(rowIndex, 1, 1, 26).getValues()[0];
    
    var teacher = {
      rowIndex: rowIndex,
      eddis: row[COL.EDDIS] || '',
      district: row[COL.DISTRICT] || '',
      school: row[COL.SCHOOL] || '',
      level: row[COL.LEVEL] || '',
      shsTrackStrand: row[COL.SHS_TRACK_STRAND] || '',
      itemNumber: row[COL.ITEM_NUMBER] || '',
      positionTitle: row[COL.POSITION_TITLE] || '',
      sg: row[COL.SG] || '',
      step: row[COL.STEP] || '',
      employee: row[COL.EMPLOYEE] || '',
      lastName: row[COL.LASTNAME] || '',
      firstName: row[COL.FIRSTNAME] || '',
      middleName: row[COL.MIDDLE_NAME] || '',
      ext: row[COL.EXT] || '',
      dateOfBirth: formatDate(row[COL.DATE_OF_BIRTH]),
      dateOfOriginalApt: formatDate(row[COL.DATE_OF_ORIGINAL_APT]),
      dateOfLastPromotion: formatDate(row[COL.DATE_OF_LAST_PROMOTION]),
      courseMajorCollege: row[COL.COURSE_MAJOR_COLLEGE] || '',
      courseMajorGraduate: row[COL.COURSE_MAJOR_GRADUATE] || '',
      philhealth: row[COL.PHILHEALTH] || '',
      gsisBp: row[COL.GSIS_BP] || '',
      pagibig: row[COL.PAGIBIG] || '',
      tin: row[COL.TIN] || '',
      completeAddress: row[COL.COMPLETE_ADDRESS] || '',
      eligibility: row[COL.ELIGIBILITY] || '',
      gender: row[COL.GENDER] || ''
    };
    
    return { success: true, data: teacher };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==================== 201 FILES - UPDATE TEACHER ====================
function update201File(rowIndex, teacherData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Sheet not found: ' + SHEET_NAME };
    }
    
    var rowData = [
      teacherData.eddis || '',
      teacherData.district || '',
      teacherData.school || '',
      teacherData.level || '',
      teacherData.shsTrackStrand || '',
      teacherData.itemNumber || '',
      teacherData.positionTitle || '',
      teacherData.sg || '',
      teacherData.step || '',
      teacherData.employee || '',
      teacherData.lastName || '',
      teacherData.firstName || '',
      teacherData.middleName || '',
      teacherData.ext || '',
      teacherData.dateOfBirth || '',
      teacherData.dateOfOriginalApt || '',
      teacherData.dateOfLastPromotion || '',
      teacherData.courseMajorCollege || '',
      teacherData.courseMajorGraduate || '',
      teacherData.philhealth || '',
      teacherData.gsisBp || '',
      teacherData.pagibig || '',
      teacherData.tin || '',
      teacherData.completeAddress || '',
      teacherData.eligibility || '',
      teacherData.gender || ''
    ];
    
    sheet.getRange(rowIndex, 1, 1, 26).setValues([rowData]);
    
    return { success: true, message: 'Teacher 201 file updated successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==================== 201 FILES - ADD NEW TEACHER ====================
function add201File(teacherData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Sheet not found: ' + SHEET_NAME };
    }
    
    var rowData = [
      teacherData.eddis || '',
      teacherData.district || '',
      teacherData.school || '',
      teacherData.level || '',
      teacherData.shsTrackStrand || '',
      teacherData.itemNumber || '',
      teacherData.positionTitle || '',
      teacherData.sg || '',
      teacherData.step || '',
      teacherData.employee || '',
      teacherData.lastName || '',
      teacherData.firstName || '',
      teacherData.middleName || '',
      teacherData.ext || '',
      teacherData.dateOfBirth || '',
      teacherData.dateOfOriginalApt || '',
      teacherData.dateOfLastPromotion || '',
      teacherData.courseMajorCollege || '',
      teacherData.courseMajorGraduate || '',
      teacherData.philhealth || '',
      teacherData.gsisBp || '',
      teacherData.pagibig || '',
      teacherData.tin || '',
      teacherData.completeAddress || '',
      teacherData.eligibility || '',
      teacherData.gender || ''
    ];
    
    sheet.appendRow(rowData);
    
    return { success: true, message: 'New teacher 201 file added successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==================== 201 FILES - DELETE TEACHER ====================
function delete201File(rowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      return { success: false, error: 'Sheet not found: ' + SHEET_NAME };
    }
    
    sheet.deleteRow(rowIndex);
    
    return { success: true, message: 'Teacher 201 file deleted successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==================== 201 FILES - GET BY EMAIL (Teacher Dashboard) ====================
function getTeacher201ByEmail(email) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    
    var usersSheet = ss.getSheetByName('NFC REGISTERED');
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
      var lastName = String(row[COL.LASTNAME] || '').trim();
      var firstName = String(row[COL.FIRSTNAME] || '').trim();
      var fullName = firstName + ' ' + lastName;
      
      if (teacherName && (
        fullName.toLowerCase().includes(teacherName.toLowerCase()) ||
        teacherName.toLowerCase().includes(lastName.toLowerCase()) && lastName.length > 1
      )) {
        return {
          success: true,
          data: {
            rowIndex: i + 1,
            eddis: row[COL.EDDIS] || '',
            district: row[COL.DISTRICT] || '',
            school: row[COL.SCHOOL] || '',
            level: row[COL.LEVEL] || '',
            shsTrackStrand: row[COL.SHS_TRACK_STRAND] || '',
            itemNumber: row[COL.ITEM_NUMBER] || '',
            positionTitle: row[COL.POSITION_TITLE] || '',
            sg: row[COL.SG] || '',
            step: row[COL.STEP] || '',
            employee: row[COL.EMPLOYEE] || '',
            lastName: lastName,
            firstName: firstName,
            middleName: row[COL.MIDDLE_NAME] || '',
            ext: row[COL.EXT] || '',
            dateOfBirth: formatDate(row[COL.DATE_OF_BIRTH]),
            dateOfOriginalApt: formatDate(row[COL.DATE_OF_ORIGINAL_APT]),
            dateOfLastPromotion: formatDate(row[COL.DATE_OF_LAST_PROMOTION]),
            courseMajorCollege: row[COL.COURSE_MAJOR_COLLEGE] || '',
            courseMajorGraduate: row[COL.COURSE_MAJOR_GRADUATE] || '',
            philhealth: row[COL.PHILHEALTH] || '',
            gsisBp: row[COL.GSIS_BP] || '',
            pagibig: row[COL.PAGIBIG] || '',
            tin: row[COL.TIN] || '',
            completeAddress: row[COL.COMPLETE_ADDRESS] || '',
            eligibility: row[COL.ELIGIBILITY] || '',
            gender: row[COL.GENDER] || ''
          }
        };
      }
    }
    
    return { success: false, error: 'No 201 file found for this account.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==================== HELPER FUNCTIONS ====================
function formatDate(dateValue) {
  if (!dateValue) return '';
  if (dateValue instanceof Date) {
    var options = { year: 'numeric', month: 'long', day: 'numeric' };
    return dateValue.toLocaleDateString('en-US', options);
  }
  return String(dateValue);
}
