/**
 * ATTENDANCE MANAGER — Google Apps Script Web App
 * -------------------------------------------------
 * Each class gets its own Google Spreadsheet (created automatically) with:
 *   - "Students"              : RollNo | Name
 *   - "Attendance - Mon YYYY" : RollNo | Name | <date1> | <date2> | ...
 *   - "Tests - Mon YYYY"      : RollNo | Name | <test1 header> | <test2 header> | ...
 * Each test column header encodes "Subject - dd-Mon-yyyy - NNM" so a test
 * can be found again later to edit marks or regenerate its message.
 *
 * The Class -> Spreadsheet mapping is stored in Script Properties, so it
 * persists across edits/redeployments and needs no separate "master" sheet.
 *
 * SETUP:
 * 1. Paste this file into Code.gs of a new Apps Script project.
 * 2. Create a new HTML file named exactly "Index" and paste Index.html into it.
 * 3. Deploy > New deployment > Web app (Execute as: Me).
 * 4. Open the deployment URL and authorize Drive + Sheets + Docs access.
 *
 * IMPORTANT: After editing this code, create a NEW DEPLOYMENT VERSION
 * (Deploy > Manage deployments > pencil icon > Version: New version > Deploy)
 * for changes to take effect on the live web app URL.
 */

const CLASS_MAP_KEY = 'CLASS_MAP';
const FOLDER_NAME = 'Attendance Manager - Class Sheets';
const ATTENDANCE_SHEET_PREFIX = 'Attendance - ';
const TEST_SHEET_PREFIX = 'Tests - ';
const MONTH_NAMES_ = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const FULL_MONTH_NAMES_ = ['January','February','March','April','May','June','July',
  'August','September','October','November','December'];

// Parses a 'yyyy-mm-dd' string (what <input type="date"> always sends) into parts.
// This never depends on the browser's locale.
function parseIsoDate_(isoDate) {
  const parts = (isoDate || '').trim().split('-');
  if (parts.length !== 3) throw new Error('Invalid date: ' + isoDate);
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10); // 1-12
  const day = parseInt(parts[2], 10);
  if (!year || !month || !day || month < 1 || month > 12) {
    throw new Error('Invalid date: ' + isoDate);
  }
  return { year: year, month: month, day: day };
}

// Builds the "dd-Mon-yyyy" text used as the attendance column header.
function formatDisplayDate_(isoDate) {
  const d = parseIsoDate_(isoDate);
  return String(d.day).padStart(2, '0') + '-' + MONTH_NAMES_[d.month - 1] + '-' + d.year;
}

// Builds a "25 August 2026" style date, used in the attendance WhatsApp message.
function formatFullDate_(isoDate) {
  const d = parseIsoDate_(isoDate);
  return d.day + ' ' + FULL_MONTH_NAMES_[d.month - 1] + ' ' + d.year;
}

// Builds a "06/09/2026" style date, used in the test-result WhatsApp message.
function formatSlashDate_(isoDate) {
  const d = parseIsoDate_(isoDate);
  return String(d.day).padStart(2, '0') + '/' + String(d.month).padStart(2, '0') + '/' + d.year;
}

// Builds "dd/mm" — used for exam dates on the performance report PDF.
function formatDdMm_(isoDate) {
  const d = parseIsoDate_(isoDate);
  return String(d.day).padStart(2, '0') + '/' + String(d.month).padStart(2, '0');
}

// Builds the monthly attendance sheet name, e.g. "Attendance - Sep 2026".
function getMonthSheetNameFromIso_(isoDate) {
  const d = parseIsoDate_(isoDate);
  return ATTENDANCE_SHEET_PREFIX + MONTH_NAMES_[d.month - 1] + ' ' + d.year;
}

// Builds the monthly test-marks sheet name, e.g. "Tests - Sep 2026".
function getTestMonthSheetNameFromIso_(isoDate) {
  const d = parseIsoDate_(isoDate);
  return TEST_SHEET_PREFIX + MONTH_NAMES_[d.month - 1] + ' ' + d.year;
}

// Safely turns a header cell's value back into "dd-Mon-yyyy" text, even if
// Sheets silently auto-converted it into a real Date object.
function headerToDisplayString_(headerValue) {
  if (Object.prototype.toString.call(headerValue) === '[object Date]') {
    return String(headerValue.getDate()).padStart(2, '0') + '-' +
      MONTH_NAMES_[headerValue.getMonth()] + '-' + headerValue.getFullYear();
  }
  return String(headerValue).trim();
}

// Builds a test column header, e.g. "BIOLOGY (UNIT 2) - 06-Sep-2026 - 55M".
function buildTestHeader_(subject, isoDate, totalMarks) {
  return subject.trim() + ' - ' + formatDisplayDate_(isoDate) + ' - ' + totalMarks + 'M';
}

// Parses a test column header back into { subject, displayDate, isoDate, totalMarks }.
// Returns null if the header doesn't match the expected pattern.
function parseTestHeader_(headerValue) {
  const headerStr = headerToDisplayString_(headerValue);
  const m = headerStr.match(/^(.*) - (\d{2})-([A-Za-z]{3})-(\d{4}) - (\d+)M$/);
  if (!m) return null;
  const monthIdx = MONTH_NAMES_.findIndex(function (mn) { return mn.toLowerCase() === m[3].toLowerCase(); });
  if (monthIdx === -1) return null;
  return {
    subject: m[1].trim(),
    displayDate: m[2] + '-' + m[3] + '-' + m[4],
    isoDate: m[4] + '-' + String(monthIdx + 1).padStart(2, '0') + '-' + m[2],
    totalMarks: parseInt(m[5], 10)
  };
}

// Splits the combined "SUBJECT (CHAPTER)" string (built by the front end)
// back into its two parts, e.g. for the performance report table.
function splitSubjectChapter_(combined) {
  const m = combined.match(/^(.*?)\s*\(([^)]*)\)\s*$/);
  if (m) return { subject: m[1].trim(), chapter: m[2].trim() };
  return { subject: combined.trim(), chapter: '' };
}

// Formats a fraction (0-100) as "NN.N%".
function formatPercent_(n) {
  return (Math.round(n * 10) / 10).toFixed(1) + '%';
}

// ---------------------------------------------------------------------
// INTERNAL HELPERS
// ---------------------------------------------------------------------
function getClassMap_() {
  const raw = PropertiesService.getScriptProperties().getProperty(CLASS_MAP_KEY);
  return raw ? JSON.parse(raw) : [];
}

function saveClassMap_(map) {
  PropertiesService.getScriptProperties().setProperty(CLASS_MAP_KEY, JSON.stringify(map));
}

function findClass_(map, className) {
  const target = (className || '').trim().toLowerCase();
  return map.find(function (c) { return c.name.toLowerCase() === target; });
}

function getOrCreateFolder_() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

function getClassOrThrow_(map, className) {
  const cls = findClass_(map, className);
  if (!cls) throw new Error('Class "' + className + '" was not found. It may have been renamed or deleted.');
  return cls;
}

function getAllAttendanceSheets_(ss) {
  return ss.getSheets().filter(function (s) {
    return s.getName().indexOf(ATTENDANCE_SHEET_PREFIX) === 0;
  });
}

function getAllTestSheets_(ss) {
  return ss.getSheets().filter(function (s) {
    return s.getName().indexOf(TEST_SHEET_PREFIX) === 0;
  });
}

// Returns every roster-tracking sheet (attendance AND test sheets) so
// addStudent/deleteStudent can keep both kinds in sync with "Students".
function getAllDataSheets_(ss) {
  return getAllAttendanceSheets_(ss).concat(getAllTestSheets_(ss));
}

// Sorts monthly sheets ("Attendance - Mon YYYY" / "Tests - Mon YYYY")
// into chronological order, regardless of the order they were created in.
function sortMonthSheets_(sheets, prefix) {
  return sheets.slice().sort(function (a, b) {
    const pa = parseMonthSheetName_(a.getName(), prefix);
    const pb = parseMonthSheetName_(b.getName(), prefix);
    if (pa.year !== pb.year) return pa.year - pb.year;
    return pa.monthIndex - pb.monthIndex;
  });
}

// Extracts { monthIndex, year } from a "<prefix>Mon YYYY" sheet name.
function parseMonthSheetName_(sheetName, prefix) {
  const label = sheetName.substring(prefix.length).trim().split(' ');
  const monthIndex = MONTH_NAMES_.findIndex(function (mn) { return mn === label[0]; });
  return { monthIndex: monthIndex, year: parseInt(label[1], 10) };
}

// Returns the 1-based row number of a roll number within a roster sheet,
// or -1 if the roll number isn't present.
function findStudentRow_(sheet, rollNo) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const rolls = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < rolls.length; i++) {
    if (String(rolls[i][0]).trim() === String(rollNo).trim()) return i + 2;
  }
  return -1;
}

// Gets (or lazily creates) the monthly attendance sheet for the given date.
// A brand-new month sheet is pre-populated with the current student roster
// so it never starts out of sync with the Students sheet.
function getOrCreateMonthSheet_(ss, isoDate) {
  const sheetName = getMonthSheetNameFromIso_(isoDate);
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, 2).setValues([['RollNo', 'Name']]);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.getRange('A1:B1').setFontWeight('bold');
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 200);

  const studentsSheet = ss.getSheetByName('Students');
  const lastRow = studentsSheet.getLastRow();
  if (lastRow > 1) {
    const data = studentsSheet.getRange(2, 1, lastRow - 1, 2).getValues()
      .filter(function (r) { return r[0] !== '' && r[0] !== null; });
    if (data.length) {
      sheet.getRange(2, 1, data.length, 2).setValues(data);
    }
  }
  return sheet;
}

// Gets (or lazily creates) the monthly test-marks sheet for the given date.
function getOrCreateTestMonthSheet_(ss, isoDate) {
  const sheetName = getTestMonthSheetNameFromIso_(isoDate);
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, 2).setValues([['RollNo', 'Name']]);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.getRange('A1:B1').setFontWeight('bold');
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 200);

  const studentsSheet = ss.getSheetByName('Students');
  const lastRow = studentsSheet.getLastRow();
  if (lastRow > 1) {
    const data = studentsSheet.getRange(2, 1, lastRow - 1, 2).getValues()
      .filter(function (r) { return r[0] !== '' && r[0] !== null; });
    if (data.length) {
      sheet.getRange(2, 1, data.length, 2).setValues(data);
    }
  }
  return sheet;
}

// ---------------------------------------------------------------------
// WEB APP ENTRY POINT
// ---------------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Synapse Academy - Attendance Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------------
// CLASS MANAGEMENT
// ---------------------------------------------------------------------
function addClass(className) {
  className = (className || '').trim();
  if (!className) throw new Error('Class name cannot be empty.');
  if (className.length > 80) throw new Error('Class name is too long.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const map = getClassMap_();
    if (findClass_(map, className)) {
      throw new Error('A class named "' + className + '" already exists.');
    }

    const ss = SpreadsheetApp.create(className + ' - Attendance');
    const file = DriveApp.getFileById(ss.getId());

    try {
      getOrCreateFolder_().addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (e) {
      // Non-fatal — file still exists and is usable even if it stays in root
    }

    const studentsSheet = ss.getSheets()[0];
    studentsSheet.setName('Students');
    studentsSheet.getRange(1, 1, 1, 2).setValues([['RollNo', 'Name']]);
    studentsSheet.setFrozenRows(1);
    studentsSheet.getRange('A1:B1').setFontWeight('bold');
    studentsSheet.setColumnWidth(1, 80);
    studentsSheet.setColumnWidth(2, 200);

    const entry = { name: className, sheetId: ss.getId(), sheetUrl: ss.getUrl() };
    map.push(entry);
    saveClassMap_(map);

    return entry;
  } finally {
    lock.releaseLock();
  }
}

function getClassList() {
  return getClassMap_().sort(function (a, b) { return a.name.localeCompare(b.name); });
}

function deleteClass(className) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const map = getClassMap_();
    const idx = map.findIndex(function (c) { return c.name.toLowerCase() === className.trim().toLowerCase(); });
    if (idx === -1) throw new Error('Class not found.');
    const removed = map.splice(idx, 1)[0];
    saveClassMap_(map);
    return removed;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------
// STUDENT MANAGEMENT
// ---------------------------------------------------------------------
function addStudent(className, rollNo, name) {
  rollNo = (rollNo || '').toString().trim();
  name = (name || '').toString().trim();
  if (!rollNo || !name) throw new Error('Roll number and name are both required.');
  if (rollNo.length > 20) throw new Error('Roll number is too long.');

  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(cls.sheetId);
    const studentsSheet = ss.getSheetByName('Students');
    const lastRow = studentsSheet.getLastRow();

    if (lastRow > 1) {
      const existingRolls = studentsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < existingRolls.length; i++) {
        if (String(existingRolls[i][0]).trim() === rollNo) {
          throw new Error('Roll number ' + rollNo + ' already exists in ' + className + '.');
        }
      }
    }

    studentsSheet.appendRow([rollNo, name]);

    getAllDataSheets_(ss).forEach(function (sheet) {
      sheet.appendRow([rollNo, name]);
    });

    return { rollNo: rollNo, name: name };
  } finally {
    lock.releaseLock();
  }
}

function getStudents(className) {
  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);

  const ss = SpreadsheetApp.openById(cls.sheetId);
  const studentsSheet = ss.getSheetByName('Students');
  const lastRow = studentsSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = studentsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return data
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) { return { rollNo: String(r[0]), name: String(r[1]) }; })
    .sort(function (a, b) {
      const na = parseFloat(a.rollNo), nb = parseFloat(b.rollNo);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return a.rollNo.localeCompare(b.rollNo);
    });
}

function deleteStudent(className, rollNo) {
  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);
  rollNo = String(rollNo).trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(cls.sheetId);
    const sheetsToClean = [ss.getSheetByName('Students')].concat(getAllDataSheets_(ss));

    sheetsToClean.forEach(function (sheet) {
      if (!sheet) return;
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      const rolls = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = rolls.length - 1; i >= 0; i--) {
        if (String(rolls[i][0]).trim() === rollNo) {
          sheet.deleteRow(i + 2);
          break;
        }
      }
    });
    return true;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------
// ATTENDANCE
// ---------------------------------------------------------------------
/**
 * records: [{ rollNo: '1', status: 'Present' | 'Absent' }, ...]
 * isoDate: e.g. '2026-09-04' — determines both the column AND which
 *          monthly sheet ("Attendance - Sep 2026") data is written to.
 */
function markAttendance(className, isoDate, records) {
  isoDate = (isoDate || '').trim();
  if (!isoDate) throw new Error('Date is required.');
  if (!records || !records.length) throw new Error('No attendance records were provided.');

  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);
  const displayDate = formatDisplayDate_(isoDate);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(cls.sheetId);
    const sheet = getOrCreateMonthSheet_(ss, isoDate);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2) throw new Error('No students found in this class yet. Add students first.');

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    let dateCol = -1;
    for (let c = 2; c < headers.length; c++) {
      if (headerToDisplayString_(headers[c]) === displayDate) { dateCol = c + 1; break; }
    }
    if (dateCol === -1) {
      dateCol = lastCol + 1;
      sheet.getRange(1, dateCol).setNumberFormat('@').setValue(displayDate).setFontWeight('bold');
    }

    const rollColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const rowByRoll = {};
    rollColumn.forEach(function (r, i) { rowByRoll[String(r[0]).trim()] = i + 2; });

    const statusValues = [];
    for (let r = 2; r <= lastRow; r++) statusValues.push(['']);

    let marked = 0;
    records.forEach(function (rec) {
      const row = rowByRoll[String(rec.rollNo).trim()];
      if (row) {
        statusValues[row - 2][0] = rec.status === 'Present' ? 'P' : 'A';
        marked++;
      }
    });

    sheet.getRange(2, dateCol, lastRow - 1, 1).setValues(statusValues);
    sheet.autoResizeColumn(dateCol);

    return { date: displayDate, marked: marked, sheet: sheet.getName() };
  } finally {
    lock.releaseLock();
  }
}

/** Returns { rollNo: 'Present' | 'Absent' | '' } for prefilling the UI on a given date */
function getAttendanceForDate(className, isoDate) {
  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);
  isoDate = (isoDate || '').trim();
  const displayDate = formatDisplayDate_(isoDate);

  const ss = SpreadsheetApp.openById(cls.sheetId);
  const sheetName = getMonthSheetNameFromIso_(isoDate);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 3) return {};

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let dateCol = -1;
  for (let c = 2; c < headers.length; c++) {
    if (headerToDisplayString_(headers[c]) === displayDate) { dateCol = c + 1; break; }
  }
  if (dateCol === -1) return {};

  const rolls = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const statuses = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();

  const result = {};
  for (let i = 0; i < rolls.length; i++) {
    const roll = String(rolls[i][0]).trim();
    if (!roll) continue;
    const s = statuses[i][0];
    result[roll] = s === 'P' ? 'Present' : (s === 'A' ? 'Absent' : '');
  }
  return result;
}

/** Names of every monthly attendance sheet that exists so far, most recent first. */
function getAttendanceMonths(className) {
  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);
  const ss = SpreadsheetApp.openById(cls.sheetId);
  return getAllAttendanceSheets_(ss)
    .map(function (s) { return s.getName().replace(ATTENDANCE_SHEET_PREFIX, ''); })
    .reverse();
}

// ---------------------------------------------------------------------
// TEST MARKS
// ---------------------------------------------------------------------
/**
 * records: [{ rollNo: '1', marks: 46 }, { rollNo: '2', marks: 'AB' }, ...]
 * marks may be a number (0..totalMarks) or the string 'AB' for a student
 * who was absent for the test.
 *
 * If a test with the same subject + date already exists for this class,
 * its column is reused (so re-saving updates marks / total marks instead
 * of creating a duplicate column).
 */
function recordTestMarks(className, isoDate, subject, totalMarks, records) {
  isoDate = (isoDate || '').trim();
  subject = (subject || '').trim();
  totalMarks = parseInt(totalMarks, 10);

  if (!isoDate) throw new Error('Date is required.');
  if (!subject) throw new Error('Subject is required.');
  if (!totalMarks || totalMarks <= 0) throw new Error('Total marks must be a positive number.');
  if (!records || !records.length) throw new Error('No marks were provided.');

  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(cls.sheetId);
    const sheet = getOrCreateTestMonthSheet_(ss, isoDate);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2) throw new Error('No students found in this class yet. Add students first.');

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const targetKey = subject.toLowerCase() + '|' + formatDisplayDate_(isoDate);

    let testCol = -1;
    for (let c = 2; c < headers.length; c++) {
      const parsed = parseTestHeader_(headers[c]);
      if (parsed && (parsed.subject.toLowerCase() + '|' + parsed.displayDate) === targetKey) {
        testCol = c + 1;
        break;
      }
    }
    if (testCol === -1) testCol = lastCol + 1;

    const headerText = buildTestHeader_(subject, isoDate, totalMarks);
    sheet.getRange(1, testCol).setNumberFormat('@').setValue(headerText).setFontWeight('bold');

    const rollColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const rowByRoll = {};
    rollColumn.forEach(function (r, i) { rowByRoll[String(r[0]).trim()] = i + 2; });

    const marksValues = [];
    for (let r = 2; r <= lastRow; r++) marksValues.push(['']);

    let marked = 0;
    records.forEach(function (rec) {
      const row = rowByRoll[String(rec.rollNo).trim()];
      if (!row) return;
      let val = rec.marks;
      if (typeof val === 'string' && val.trim().toUpperCase() === 'AB') {
        val = 'AB';
      } else {
        const n = parseFloat(val);
        val = isNaN(n) ? '' : n;
      }
      if (val !== '') {
        marksValues[row - 2][0] = val;
        marked++;
      }
    });

    sheet.getRange(2, testCol, lastRow - 1, 1).setValues(marksValues);
    sheet.autoResizeColumn(testCol);

    return {
      sheetName: sheet.getName(),
      columnIndex: testCol,
      subject: subject,
      displayDate: formatDisplayDate_(isoDate),
      totalMarks: totalMarks,
      marked: marked
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns { marks: { rollNo: number|'AB' }, totalMarks: number|null } for
 * prefilling the UI when re-opening an already-saved test.
 */
function getMarksForTest(className, isoDate, subject) {
  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);
  isoDate = (isoDate || '').trim();
  subject = (subject || '').trim();

  const ss = SpreadsheetApp.openById(cls.sheetId);
  const sheetName = getTestMonthSheetNameFromIso_(isoDate);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { marks: {}, totalMarks: null };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 3) return { marks: {}, totalMarks: null };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const targetKey = subject.toLowerCase() + '|' + formatDisplayDate_(isoDate);

  let testCol = -1;
  let totalMarks = null;
  for (let c = 2; c < headers.length; c++) {
    const parsed = parseTestHeader_(headers[c]);
    if (parsed && (parsed.subject.toLowerCase() + '|' + parsed.displayDate) === targetKey) {
      testCol = c + 1;
      totalMarks = parsed.totalMarks;
      break;
    }
  }
  if (testCol === -1) return { marks: {}, totalMarks: null };

  const rolls = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const marksVals = sheet.getRange(2, testCol, lastRow - 1, 1).getValues();

  const result = {};
  for (let i = 0; i < rolls.length; i++) {
    const roll = String(rolls[i][0]).trim();
    if (!roll) continue;
    result[roll] = marksVals[i][0];
  }
  return { marks: result, totalMarks: totalMarks };
}

/**
 * Every test entered for a class, most recently added first. Each entry
 * carries sheetName + columnIndex so the client can pass them straight to
 * generateTestMessage without re-typing date/subject.
 */
function getTestsForClass(className) {
  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);
  const ss = SpreadsheetApp.openById(cls.sheetId);
  const sheets = getAllTestSheets_(ss);

  const tests = [];
  sheets.forEach(function (sheet) {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 3) return;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (let c = 2; c < headers.length; c++) {
      const parsed = parseTestHeader_(headers[c]);
      if (parsed) {
        tests.push({
          sheetName: sheet.getName(),
          columnIndex: c + 1,
          subject: parsed.subject,
          displayDate: parsed.displayDate,
          isoDate: parsed.isoDate,
          totalMarks: parsed.totalMarks
        });
      }
    }
  });

  return tests.reverse();
}

/**
 * Builds the "*RESULT DECLARATION*" WhatsApp message for one already-saved
 * test, identified by sheetName + columnIndex. The student with the
 * highest numeric score (ties included) gets a trailing "**"; absent
 * students show "AB"; students with no marks entered yet show "--".
 */
function generateTestMessage(className, sheetName, columnIndex) {
  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);
  const ss = SpreadsheetApp.openById(cls.sheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('That test could not be found (sheet missing).');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  columnIndex = parseInt(columnIndex, 10);
  if (lastRow < 2 || columnIndex < 3 || columnIndex > lastCol) {
    throw new Error('That test could not be found.');
  }

  const headerValue = sheet.getRange(1, columnIndex).getValue();
  const parsed = parseTestHeader_(headerValue);
  if (!parsed) throw new Error("That test's column header could not be read.");

  const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const marksVals = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues();

  const rows = [];
  let highest = -Infinity;
  for (let i = 0; i < names.length; i++) {
    const name = String(names[i][0]).trim();
    if (!name) continue;
    const raw = marksVals[i][0];
    let display, numeric = null;
    if (raw === '' || raw === null || raw === undefined) {
      display = '--';
    } else if (String(raw).trim().toUpperCase() === '*ABSENT*') {
      display = '*ABSENT*';
    } else {
      numeric = parseFloat(raw);
      display = isNaN(numeric) ? String(raw) : numeric;
      if (!isNaN(numeric) && numeric > highest) highest = numeric;
    }
    rows.push({ name: name, display: display, numeric: numeric });
  }

  const lines = rows.map(function (r, i) {
    let marksText = String(r.display);
    if (r.numeric !== null && r.numeric === highest) marksText += '**';
    return (i + 1) + ') ' + r.name + ': ' + marksText;
  });

  const message =
    '*RESULT DECLARATION*\n' +
    'DATE: ' + formatSlashDate_(parsed.isoDate) + '\n\n' +
    'SUBJECT: ' + parsed.subject.toUpperCase() + ' *(' + parsed.totalMarks + ' MARKS)*\n\n' +
    lines.join('\n');

  return {
    message: message,
    subject: parsed.subject,
    displayDate: parsed.displayDate,
    totalMarks: parsed.totalMarks,
    studentCount: rows.length
  };
}

// ---------------------------------------------------------------------
// MULTI-CLASS (COMBINED BATCH) TESTS
// ---------------------------------------------------------------------
function getStudentsForClasses(classNames) {
  if (!classNames || !classNames.length) throw new Error('Select at least one class.');
  const out = [];
  classNames.forEach(function (cn) {
    getStudents(cn).forEach(function (s) {
      out.push({ className: cn, rollNo: s.rollNo, name: s.name });
    });
  });
  return out;
}

/** Returns { marks: { "Class::roll": value }, totalMarks } across all given classes. */
function getMarksForTestMulti(classNames, isoDate, subject) {
  const marks = {};
  let totalMarks = null;
  classNames.forEach(function (cn) {
    const res = getMarksForTest(cn, isoDate, subject);
    if (res.totalMarks && !totalMarks) totalMarks = res.totalMarks;
    Object.keys(res.marks).forEach(function (roll) {
      marks[cn + '::' + roll] = res.marks[roll];
    });
  });
  return { marks: marks, totalMarks: totalMarks };
}

/**
 * records: [{ className, rollNo, marks }, ...] — grouped by class and saved
 * into each class's own spreadsheet.
 */
function recordTestMarksMulti(classNames, isoDate, subject, totalMarks, records) {
  if (!records || !records.length) throw new Error('No marks were provided.');
  let marked = 0;
  classNames.forEach(function (cn) {
    const recs = records
      .filter(function (r) { return r.className === cn; })
      .map(function (r) { return { rollNo: r.rollNo, marks: r.marks }; });
    if (!recs.length) return;
    const res = recordTestMarks(cn, isoDate, subject, totalMarks, recs);
    marked += res.marked;
  });
  return { subject: subject, marked: marked };
}

/** Every distinct test (subject + date) found in any of the given classes. */
function getTestsForClasses(classNames) {
  const seen = {};
  const tests = [];
  classNames.forEach(function (cn) {
    getTestsForClass(cn).forEach(function (t) {
      const key = t.subject.toLowerCase() + '|' + t.isoDate;
      if (seen[key]) return;
      seen[key] = true;
      tests.push({
        subject: t.subject,
        displayDate: t.displayDate,
        isoDate: t.isoDate,
        totalMarks: t.totalMarks
      });
    });
  });
  tests.sort(function (a, b) { return b.isoDate.localeCompare(a.isoDate); });
  return tests;
}

/** Builds ONE "RESULT DECLARATION" message covering every selected class. */
function generateCombinedTestMessage(classNames, isoDate, subject) {
  isoDate = (isoDate || '').trim();
  subject = (subject || '').trim();
  const targetKey = subject.toLowerCase() + '|' + formatDisplayDate_(isoDate);

  const rows = [];
  let highest = -Infinity;
  let totalMarks = null;

  classNames.forEach(function (cn) {
    const cls = getClassOrThrow_(getClassMap_(), cn);
    const ss = SpreadsheetApp.openById(cls.sheetId);
    const sheet = ss.getSheetByName(getTestMonthSheetNameFromIso_(isoDate));

    let testCol = -1;
    let source = ss.getSheetByName('Students'); // fallback: names only
    if (sheet && sheet.getLastColumn() >= 3) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      for (let c = 2; c < headers.length; c++) {
        const p = parseTestHeader_(headers[c]);
        if (p && (p.subject.toLowerCase() + '|' + p.displayDate) === targetKey) {
          testCol = c + 1;
          if (!totalMarks) totalMarks = p.totalMarks;
          break;
        }
      }
      if (testCol !== -1) source = sheet;
    }

    const lastRow = source.getLastRow();
    if (lastRow < 2) return;
    const names = source.getRange(2, 2, lastRow - 1, 1).getValues();
    const marksVals = testCol !== -1
      ? source.getRange(2, testCol, lastRow - 1, 1).getValues()
      : null;

    for (let i = 0; i < names.length; i++) {
      const name = String(names[i][0]).trim();
      if (!name) continue;
      const raw = marksVals ? marksVals[i][0] : '';
      let display, numeric = null;
      if (raw === '' || raw === null || raw === undefined) {
        display = '--';
      } else if (String(raw).trim().toUpperCase() === 'AB') {
        display = 'AB';
      } else {
        numeric = parseFloat(raw);
        display = isNaN(numeric) ? String(raw) : numeric;
        if (!isNaN(numeric) && numeric > highest) highest = numeric;
      }
      rows.push({ name: name, display: display, numeric: numeric });
    }
  });

  if (!totalMarks) throw new Error('That test could not be found in the selected classes.');

  const lines = rows.map(function (r, i) {
    let marksText = String(r.display);
    if (r.numeric !== null && r.numeric === highest) marksText += '**';
    return (i + 1) + ') ' + r.name + ': ' + marksText;
  });

  const message =
    '*RESULT DECLARATION*\n' +
    'DATE: ' + formatSlashDate_(isoDate) + '\n\n' +
    'SUBJECT: ' + subject.toUpperCase() + ' *(' + totalMarks + ' MARKS)*\n\n' +
    lines.join('\n');

  return { message: message, studentCount: rows.length, totalMarks: totalMarks };
}

// ---------------------------------------------------------------------
// STUDENT PERFORMANCE REPORT (PDF)
// ---------------------------------------------------------------------
/**
 * Gathers every test result for one student, grouped by month in
 * chronological order. Each month also carries its obtained/max totals,
 * matching the "MONTH TOTAL" rows on the sample marksheet.
 */
function collectTestResultsForStudent_(ss, rollNo) {
  const testSheets = sortMonthSheets_(getAllTestSheets_(ss), TEST_SHEET_PREFIX);
  const months = [];

  testSheets.forEach(function (sheet) {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 3) return;

    const row = findStudentRow_(sheet, rollNo);
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const monthInfo = parseMonthSheetName_(sheet.getName(), TEST_SHEET_PREFIX);
    const exams = [];

    for (let c = 2; c < headers.length; c++) {
      const parsed = parseTestHeader_(headers[c]);
      if (!parsed) continue;
      const split = splitSubjectChapter_(parsed.subject);

      const rawMark = row !== -1 ? sheet.getRange(row, c + 1).getValue() : '';

      let obtained = 0;
      let display = '--';
      let isAbsent = false;
      if (rawMark !== '' && rawMark !== null && rawMark !== undefined) {
        if (String(rawMark).trim().toUpperCase() === 'AB') {
          isAbsent = true;
          display = 'AB';
        } else {
          const n = parseFloat(rawMark);
          if (!isNaN(n)) { obtained = n; display = n; }
        }
      }

      exams.push({
        isoDate: parsed.isoDate,
        subject: split.subject,
        chapter: split.chapter,
        totalMarks: parsed.totalMarks,
        obtained: obtained,
        display: display,
        isAbsent: isAbsent
      });
    }

    exams.sort(function (a, b) { return a.isoDate.localeCompare(b.isoDate); });

    if (exams.length) {
      const totalObtained = exams.reduce(function (sum, e) { return sum + e.obtained; }, 0);
      const totalMax = exams.reduce(function (sum, e) { return sum + e.totalMarks; }, 0);
      months.push({
        label: FULL_MONTH_NAMES_[monthInfo.monthIndex] + ' ' + monthInfo.year,
        exams: exams,
        totalObtained: totalObtained,
        totalMax: totalMax,
        percentage: totalMax > 0 ? (totalObtained / totalMax * 100) : 0
      });
    }
  });

  return months;
}

/**
 * Present / absent / total marked days for one student, across every
 * monthly attendance sheet for the class.
 */
function collectAttendanceForStudent_(ss, rollNo) {
  const attSheets = sortMonthSheets_(getAllAttendanceSheets_(ss), ATTENDANCE_SHEET_PREFIX);
  let present = 0;
  let absent = 0;

  attSheets.forEach(function (sheet) {
    const row = findStudentRow_(sheet, rollNo);
    const lastCol = sheet.getLastColumn();
    if (row === -1 || lastCol < 3) return;

    const values = sheet.getRange(row, 3, 1, lastCol - 2).getValues()[0];
    values.forEach(function (v) {
      const s = String(v).trim().toUpperCase();
      if (s === 'P') present++;
      else if (s === 'A') absent++;
    });
  });

  const total = present + absent;
  return {
    present: present,
    absent: absent,
    total: total,
    percentage: total > 0 ? (present / total * 100) : 0
  };
}

// Bolds every cell in a table row (TableRow has no editAsText() of its own).
function boldTableRow_(row) {
  for (let i = 0; i < row.getNumCells(); i++) {
    row.getCell(i).editAsText().setBold(true);
  }
}

// Best-effort academic session label (e.g. "2026-27"), assuming an
// April-March session — matches the sample marksheet's "2026-27".
function getAcademicSessionLabel_() {
  const now = new Date();
  const y = now.getFullYear();
  const startYear = now.getMonth() >= 3 ? y : y - 1; // April = month index 3
  return startYear + '-' + String((startYear + 1) % 100).padStart(2, '0');
}

/**
 * Builds the "SYNAPSE ACADEMY — Individual Student Marksheet" performance
 * report PDF for one student: every test result grouped by month (mirroring
 * the sample marksheet layout) plus an attendance summary. A temporary
 * Google Doc is used to render the PDF and is trashed immediately after —
 * nothing is left behind in Drive. Returns a base64-encoded PDF for the
 * front end to download directly.
 */
function generatePerformanceReportPdf(className, rollNo) {
  const map = getClassMap_();
  const cls = getClassOrThrow_(map, className);
  rollNo = String(rollNo).trim();
  if (!rollNo) throw new Error('Select a student first.');

  const ss = SpreadsheetApp.openById(cls.sheetId);
  const studentsSheet = ss.getSheetByName('Students');
  const studentRow = findStudentRow_(studentsSheet, rollNo);
  if (studentRow === -1) throw new Error('Student with roll number ' + rollNo + ' was not found in ' + className + '.');
  const studentName = String(studentsSheet.getRange(studentRow, 2).getValue()).trim();

  const months = collectTestResultsForStudent_(ss, rollNo);
  const attendance = collectAttendanceForStudent_(ss, rollNo);

  const grandObtained = months.reduce(function (sum, m) { return sum + m.totalObtained; }, 0);
  const grandMax = months.reduce(function (sum, m) { return sum + m.totalMax; }, 0);
  const grandPercentage = grandMax > 0 ? (grandObtained / grandMax * 100) : 0;

  const doc = DocumentApp.create('TEMP - ' + studentName + ' Performance Report - ' + new Date().getTime());
  const docId = doc.getId();

  try {
    const body = doc.getBody();
    body.setMarginTop(36).setMarginBottom(36).setMarginLeft(50).setMarginRight(50);

    const title = body.appendParagraph('SYNAPSE ACADEMY');
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    title.editAsText().setBold(true).setFontSize(18);

    const tagline = body.appendParagraph('Bridging Students With Success');
    tagline.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    tagline.editAsText().setItalic(true).setFontSize(11);

    const subtitle = body.appendParagraph('INDIVIDUAL STUDENT PERFORMANCE REPORT');
    subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    subtitle.editAsText().setBold(true).setFontSize(13);

    const monthRangeLabel = months.length
      ? (months[0].label + (months.length > 1 ? ' - ' + months[months.length - 1].label : ''))
      : 'No tests recorded yet';
    const sessionLine = body.appendParagraph(
      'Academic Session: ' + getAcademicSessionLabel_() +
      ' | Class: ' + className +
      ' | Period: ' + monthRangeLabel
    );
    sessionLine.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sessionLine.editAsText().setFontSize(10);

    body.appendParagraph(' ');
    const nameLine = body.appendParagraph('Student Name: ' + studentName + '     Roll No: ' + rollNo + '     Class: ' + className);
    nameLine.editAsText().setBold(true);
    body.appendParagraph(' ');

    if (!months.length) {
      body.appendParagraph('No test results have been recorded for this student yet.').editAsText().setItalic(true);
      body.appendParagraph(' ');
    }

    months.forEach(function (m) {
      const heading = body.appendParagraph('TEST RESULTS — ' + m.label.toUpperCase());
      heading.editAsText().setBold(true).setFontSize(12);

      const tableData = [['EXAM', 'SUBJECT & CHAPTER', 'MARKS OBTAINED', 'MAXIMUM MARKS', 'PERCENTAGE (%)']];
      m.exams.forEach(function (e, idx) {
        const dateTag = formatDdMm_(e.isoDate);
        const subjectCell = e.chapter
          ? (e.subject + '\n' + e.chapter + ' (' + dateTag + ')')
          : (e.subject + ' (' + dateTag + ')');
        const examPct = e.isAbsent ? 'AB' : formatPercent_(e.totalMarks > 0 ? (e.obtained / e.totalMarks * 100) : 0);
        tableData.push([
          'Exam ' + (idx + 1),
          subjectCell,
          String(e.display),
          String(e.totalMarks),
          examPct
        ]);
      });
      tableData.push([
        '',
        m.label.split(' ')[0].toUpperCase() + ' TOTAL',
        String(m.totalObtained),
        String(m.totalMax),
        formatPercent_(m.percentage)
      ]);

      const table = body.appendTable(tableData);
      boldTableRow_(table.getRow(0));
      boldTableRow_(table.getRow(table.getNumRows() - 1));
      body.appendParagraph(' ');
    });

    const overallHeading = body.appendParagraph('OVERALL RESULT');
    overallHeading.editAsText().setBold(true).setFontSize(12);
    const overallLine = body.appendParagraph(grandObtained + ' / ' + grandMax + '     ' + formatPercent_(grandPercentage));
    overallLine.editAsText().setBold(true).setFontSize(14);
    body.appendParagraph('Grand Total Percentage').editAsText().setItalic(true).setFontSize(9);
    body.appendParagraph(' ');

    const attHeading = body.appendParagraph('ATTENDANCE SUMMARY');
    attHeading.editAsText().setBold(true).setFontSize(12);
    const attTable = body.appendTable([
      ['TOTAL DAYS', 'PRESENT', 'ABSENT', 'ATTENDANCE %'],
      [String(attendance.total), String(attendance.present), String(attendance.absent), formatPercent_(attendance.percentage)]
    ]);
    boldTableRow_(attTable.getRow(0));
    body.appendParagraph(' ');

    const footer = body.appendParagraph('This is a computer-generated document issued by Synapse Academy, Ahmedabad, Gujarat.');
    footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    footer.editAsText().setItalic(true).setFontSize(8.5);

    doc.saveAndClose();

    const pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF);
    const fileName = (studentName.replace(/[^A-Za-z0-9 _-]/g, '').trim() || rollNo) + '_Performance_Report.pdf';
    pdfBlob.setName(fileName);

    return { fileName: fileName, base64: Utilities.base64Encode(pdfBlob.getBytes()) };
  } finally {
    DriveApp.getFileById(docId).setTrashed(true);
  }
}

// ---------------------------------------------------------------------
// WHATSAPP MESSAGE GENERATOR (ATTENDANCE)
// ---------------------------------------------------------------------
// Looks up attendance for one class on one date.
// Returns { marked: bool, absentees: [names] }
function getAttendanceInfoForClass_(cls, isoDate) {
  const displayDate = formatDisplayDate_(isoDate);
  const ss = SpreadsheetApp.openById(cls.sheetId);
  const sheetName = getMonthSheetNameFromIso_(isoDate);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { marked: false, absentees: [] };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 3) return { marked: false, absentees: [] };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let dateCol = -1;
  for (let c = 2; c < headers.length; c++) {
    if (headerToDisplayString_(headers[c]) === displayDate) { dateCol = c + 1; break; }
  }
  if (dateCol === -1) return { marked: false, absentees: [] };

  const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const statuses = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();

  const absentees = [];
  for (let i = 0; i < names.length; i++) {
    const nm = String(names[i][0]).trim();
    if (!nm) continue;
    if (statuses[i][0] === 'A') absentees.push(nm);
  }
  return { marked: true, absentees: absentees };
}

// Builds the exact WhatsApp-ready message text for one class
function buildWhatsAppMessage_(className, isoDate, info) {
  const fullDate = formatFullDate_(isoDate);

  if (!info.marked) {
    return (
      '*' + className + '*\n' +
      '⚠️ Attendance has not been marked yet for today\'s *LECTURE* on ' + fullDate + '.'
    );
  }

  if (info.absentees.length) {
    const listText = info.absentees
      .map(function (n, i) { return (i + 1) + '. ' + n; })
      .join('\n');

    return (
      '*' + className + '*\n' +
      '📢 *Attendance Update – ' + fullDate + '*\n\n' +
      'The following students were *Absent* for today\'s *LECTURE*:\n\n' +
      listText + '\n\n' +
      'Parents are kindly requested to ensure regular attendance. \n\n' +
      '~*Synapse Academy*'
    );
  }

  return (
    '*' + className + '*\n' +
    '📢 *Attendance Update – ' + fullDate + '*\n\n' +
    'All students were *Present* today\'s *LECTURE*. 🎉\n\n' +
    '~*Synapse Academy*'
  );
}

/** Returns one WhatsApp-ready message per class for the given date. */
function getWhatsAppMessages(isoDate) {
  isoDate = (isoDate || '').trim();
  if (!isoDate) throw new Error('Date is required.');

  const map = getClassMap_();
  return map.map(function (cls) {
    const info = getAttendanceInfoForClass_(cls, isoDate);
    return {
      className: cls.name,
      message: buildWhatsAppMessage_(cls.name, isoDate, info),
      marked: info.marked,
      absentCount: info.absentees.length
    };
  });
}