/**
 * Files201Functions.gs
 * 201 Files management: CRUD, archiving, and profile image lookup.
 */

// Column mapping (0-based)
var COL_201 = {
  EDDIS: 0, DISTRICT: 1, SCHOOL: 2, LEVEL: 3, SHS_TRACK_STRAND: 4,
  ITEM_NUMBER: 5, POSITION_TITLE: 6, SG: 7, STEP: 8, EMPLOYEE: 9,
  LASTNAME: 10, FIRSTNAME: 11, MIDDLE_NAME: 12, EXT: 13,
  DATE_OF_BIRTH: 14, DATE_OF_ORIGINAL_APT: 15, DATE_OF_LAST_PROMOTION: 16,
  COURSE_MAJOR_COLLEGE: 17, COURSE_MAJOR_GRADUATE: 18,
  PHILHEALTH: 19, GSIS_BP: 20, PAGIBIG: 21, TIN: 22,
  COMPLETE_ADDRESS: 23, ELIGIBILITY: 24, GENDER: 25
};

// Get all 201 files
function getAll201Files() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found: ' + FILES_SHEET_NAME };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 26);
    var values = dataRange.getValues();
    var teachers = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      teachers.push({
        rowIndex: i + 2,
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
        lastName: row[COL_201.LASTNAME] || '',
        firstName: row[COL_201.FIRSTNAME] || '',
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
      });
    }
    return { success: true, data: teachers };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Get single 201 file by row index
function get201File(rowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    var row = sheet.getRange(rowIndex, 1, 1, 26).getValues()[0];
    var teacher = {
      rowIndex: rowIndex,
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
      lastName: row[COL_201.LASTNAME] || '',
      firstName: row[COL_201.FIRSTNAME] || '',
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
    };
    return { success: true, data: teacher };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Update 201 file
function update201File(rowIndex, teacherData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    var rowData = [
      teacherData.eddis || '', teacherData.district || '', teacherData.school || '',
      teacherData.level || '', teacherData.shsTrackStrand || '', teacherData.itemNumber || '',
      teacherData.positionTitle || '', teacherData.sg || '', teacherData.step || '',
      teacherData.employee || '', teacherData.lastName || '', teacherData.firstName || '',
      teacherData.middleName || '', teacherData.ext || '',
      teacherData.dateOfBirth || '', teacherData.dateOfOriginalApt || '', teacherData.dateOfLastPromotion || '',
      teacherData.courseMajorCollege || '', teacherData.courseMajorGraduate || '',
      teacherData.philhealth || '', teacherData.gsisBp || '', teacherData.pagibig || '',
      teacherData.tin || '', teacherData.completeAddress || '',
      teacherData.eligibility || '', teacherData.gender || ''
    ];
    sheet.getRange(rowIndex, 1, 1, 26).setValues([rowData]);
    return { success: true, message: 'Teacher 201 file updated successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Add new 201 file
function add201File(teacherData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    var rowData = [
      teacherData.eddis || '', teacherData.district || '', teacherData.school || '',
      teacherData.level || '', teacherData.shsTrackStrand || '', teacherData.itemNumber || '',
      teacherData.positionTitle || '', teacherData.sg || '', teacherData.step || '',
      teacherData.employee || '', teacherData.lastName || '', teacherData.firstName || '',
      teacherData.middleName || '', teacherData.ext || '',
      teacherData.dateOfBirth || '', teacherData.dateOfOriginalApt || '', teacherData.dateOfLastPromotion || '',
      teacherData.courseMajorCollege || '', teacherData.courseMajorGraduate || '',
      teacherData.philhealth || '', teacherData.gsisBp || '', teacherData.pagibig || '',
      teacherData.tin || '', teacherData.completeAddress || '',
      teacherData.eligibility || '', teacherData.gender || ''
    ];
    sheet.appendRow(rowData);
    return { success: true, message: 'New teacher 201 file added successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Archive 201 file (move to archived sheet instead of deleting)
function archive201File(rowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    var rowData = sheet.getRange(rowIndex, 1, 1, 26).getValues()[0];
    var archivedSheet = ss.getSheetByName(ARCHIVED_FILES_SHEET_NAME);
    if (!archivedSheet) {
      archivedSheet = ss.insertSheet(ARCHIVED_FILES_SHEET_NAME);
      var headers = sheet.getRange(1, 1, 1, 26).getValues()[0];
      archivedSheet.getRange(1, 1, 1, 26).setValues([headers]);
    }
    archivedSheet.appendRow(rowData);
    sheet.deleteRow(rowIndex);
    return { success: true, message: 'Teacher 201 file archived successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Restore archived 201 file (move back to main sheet)
function restore201File(rowIndex) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var archivedSheet = ss.getSheetByName(ARCHIVED_FILES_SHEET_NAME);
    if (!archivedSheet) return { success: false, error: 'Archived sheet not found' };
    var sheet = ss.getSheetByName(FILES_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Main sheet not found' };
    var rowData = archivedSheet.getRange(rowIndex, 1, 1, 26).getValues()[0];
    sheet.appendRow(rowData);
    archivedSheet.deleteRow(rowIndex);
    return { success: true, message: 'Teacher 201 file restored successfully.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Get all archived 201 files
function getArchived201Files() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(ARCHIVED_FILES_SHEET_NAME);
    if (!sheet) return { success: true, data: [] };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };
    var dataRange = sheet.getRange(2, 1, lastRow - 1, 26);
    var values = dataRange.getValues();
    var teachers = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      teachers.push({
        rowIndex: i + 2,
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
        lastName: row[COL_201.LASTNAME] || '',
        firstName: row[COL_201.FIRSTNAME] || '',
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
      });
    }
    return { success: true, data: teachers };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Get profile image by matching name in NFC REGISTERED
function get201ProfileImageByName(lastName, firstName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!usersSheet) return '';
    var data = usersSheet.getDataRange().getValues();
    var searchLast = (lastName || '').toString().trim().toLowerCase();
    var searchFirst = (firstName || '').toString().trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      var regName = (data[i][1] || '').toString().trim().toLowerCase();
      if (searchLast && searchLast.length > 1 && regName.indexOf(searchLast) !== -1) {
        var imageRaw = (data[i][2] || '').toString().trim();
        if (imageRaw) {
          try { return getProfileImageBase64(imageRaw); } catch(e) { return ''; }
        }
        return '';
      }
    }
    if (searchFirst && searchFirst.length > 1) {
      for (var j = 1; j < data.length; j++) {
        var regName2 = (data[j][1] || '').toString().trim().toLowerCase();
        if (regName2.indexOf(searchFirst) !== -1) {
          var imageRaw2 = (data[j][2] || '').toString().trim();
          if (imageRaw2) {
            try { return getProfileImageBase64(imageRaw2); } catch(e) { return ''; }
          }
          return '';
        }
      }
    }
    return '';
  } catch(e) { return ''; }
}
