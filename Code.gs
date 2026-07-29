const SHEET_NAME = 'Daily Log';
const SUMMARY_SHEET_NAME = 'Summary';
const ARCHIVE_SHEET_NAME = 'History';

const YEAR_LEVELS = ['1st Year', '2nd Year', '3rd Year', '4th Year'];

const AM_PM_CUTOFF_HOUR = 12;
const END_OF_DAY_HOUR = 21;

const STUDENTS_BY_COURSE = {
  'BSIT': {
    '1st Year': [],
    '2nd Year': [],
    '3rd Year': [
      'Gustilo Alex Jr. Reyes',
      'Rivera Luegi S.',
      'Sioson,Eritch Jhun U.',
      'Arsaga, Shyane Brent',
      'Alava, Raven C.',
      'Villena, Jonathan Andrew A.',
      'Onglatco Prince C.',
      'Galvez, Annie Jenn',
      'Cabanas, Mary Joy',
      'Operio, Shanelle',
      'Espura, Ezza',
      'Ochida, Jaden Angelie'
    ],
    '4th Year': []
  },
  'BSN': {
    '1st Year': [],
    '2nd Year': [],
    '3rd Year': ['Espinola, Irish Nicole V.'],
    '4th Year': []
  },
  'BEED': {
    '1st Year': [],
    '2nd Year': [],
    '3rd Year': [
      'Senope,  Rose Ann G.',
      'Entredicho, Ana Pauleen N.',
      'Malaga, Louie Chanele B.',
      'Justine Mae Operiano',
      'Robelyn Dejuan',
      'Jessa Marie Bedia',
      'Alleah June Albuera',
      'Rheazel Baylosis ',
      'Omillo, Mary Angel',
      'Rivera, Ma. Nicole'
      ],
    '4th Year': []
  },
  'BSCRIM': {
    '1st Year': [],
    '2nd Year': [
      'Bantayanon, Stephen O.',
      'Cerbas, Harvey Dwight',
      'Juanillo, Denver Alien B.',
      'Deciembre, Remuel S.',
      'Tuslo, Mark Anthony J.',
      'Urbano, Jonas S.'
    ],
    '3rd Year': [],
    '4th Year': []
  }
};

// Column map (1-based) for Sheet1 (Daily Log):
// 1 Name | 2 Course | 3 Year | 4 Date | 5 AM In | 6 AM Out | 7 AM Hours
// 8 PM In | 9 PM Out | 10 PM Hours | 11 Total Hours | 12 Status | 13 Remarks

// Column map (1-based) for Sheet3 (History) — note: NO Date column here,
// each date instead has its own "Date: ..." header block above the table:
// 1 Name | 2 Course | 3 Year | 4 AM In | 5 AM Out | 6 AM Hours
// 7 PM In | 8 PM Out | 9 PM Hours | 10 Total Hours | 11 Status | 12 Remarks

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Student Assistant Time Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- Menu (Approval Workflow) ---

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Attendance Admin')
    .addItem('Approve Selected Row(s)', 'approveSelectedRows')
    .addItem('Reject Selected Row(s)', 'rejectSelectedRows')
    .addSeparator()
    .addItem('Recalculate All Summaries', 'recalcAllSummaries')
    .addItem('Archive Yesterday\'s Records Now', 'archiveOldRows')
    .addToUi();
}

function approveSelectedRows() {
  setStatusForSelectedRows_('Approved');
}

function rejectSelectedRows() {
  setStatusForSelectedRows_('Rejected');
}


function setStatusForSelectedRows_(status) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();

  if (sheetName !== SHEET_NAME && sheetName !== ARCHIVE_SHEET_NAME) {
    ui.alert('Please select rows in the "' + SHEET_NAME + '" or "' + ARCHIVE_SHEET_NAME + '" sheet first.');
    return;
  }

  const selection = activeSheet.getActiveRange();
  if (!selection) {
    ui.alert('No rows selected.');
    return;
  }

  const startRow = Math.max(selection.getRow(), 2);
  const endRow = selection.getLastRow();

  if (endRow < 2) {
    ui.alert('Please select data rows (not just the header).');
    return;
  }

  
  const statusCol = (sheetName === SHEET_NAME) ? 12 : 11;

  let count = 0;
  for (let row = startRow; row <= endRow; row++) {
    const name = activeSheet.getRange(row, 1).getValue();
    if (!isRealDataRow_(name)) continue; // laktawan ang "Date: ..." at table header rows

    const course = activeSheet.getRange(row, 2).getValue();
    const year = activeSheet.getRange(row, 3).getValue();

    activeSheet.getRange(row, statusCol).setValue(status);
    updateSummary_(name, course, year);
    count++;
  }

  if (count === 0) {
    ui.alert('Wala pang napiling valid na row. Siguraduhing naka-select ka sa mga row na may pangalan ng estudyante (hindi yung "Date:" o header row).');
    return;
  }

  ui.alert(status + ' applied to ' + count + ' row(s).');
}


function isRealDataRow_(cellA) {
  if (!cellA) return false;
  if (cellA === 'Student Assistant') return false;
  if (typeof cellA === 'string' && cellA.indexOf('Date: ') === 0) return false;
  return true;
}

function recalcAllSummaries() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const seen = {};

  for (let i = 1; i < data.length; i++) {
    const name = data[i][0], course = data[i][1], year = data[i][2];
    if (!name) continue;
    const key = name + '|' + course + '|' + year;
    if (!seen[key]) {
      seen[key] = true;
      updateSummary_(name, course, year);
    }
  }

  
  const archiveSheet = getArchiveSheet();
  const archiveData = archiveSheet.getDataRange().getValues();
  for (let i = 0; i < archiveData.length; i++) {
    const row = archiveData[i];
    const cellA = row[0];
    if (!isRealDataRow_(cellA)) continue;

    const name = row[0], course = row[1], year = row[2];
    const key = name + '|' + course + '|' + year;
    if (!seen[key]) {
      seen[key] = true;
      updateSummary_(name, course, year);
    }
  }

  SpreadsheetApp.getUi().alert('Summary recalculated.');
}

// --- Sheet setup ---

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getRange('A1').getValue() === '') {
    sheet.getRange('A1:M1').setValues([[
      'Student Assistant', 'Course', 'Year Level', 'Date',
      'AM Time In', 'AM Time Out', 'AM Hours',
      'PM Time In', 'PM Time Out', 'PM Hours',
      'Total Hours', 'Status', 'Remarks'
    ]]);
    sheet.getRange('A1:M1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SUMMARY_SHEET_NAME);
  }
  if (sheet.getRange('A1').getValue() === '') {
    sheet.getRange('A1:D1').setValues([['Student Assistant', 'Course', 'Year Level', 'Total Hours (Approved)']]);
    sheet.getRange('A1:D1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getArchiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ARCHIVE_SHEET_NAME);
  }
  return sheet;
}

// --- Validation helpers ---

function validateCourse_(course) {
  return typeof course === 'string' && STUDENTS_BY_COURSE.hasOwnProperty(course);
}

function validateYear_(course, year) {
  if (!validateCourse_(course)) return false;
  return typeof year === 'string' && STUDENTS_BY_COURSE[course].hasOwnProperty(year);
}

function validateName_(course, year, name) {
  if (!validateYear_(course, year)) return false;
  if (typeof name !== 'string') return false;
  const trimmed = name.trim();
  return STUDENTS_BY_COURSE[course][year].indexOf(trimmed) !== -1;
}

// --- Data for frontend ---

function getCourseList() {
  return Object.keys(STUDENTS_BY_COURSE);
}

function getYearLevels() {
  return YEAR_LEVELS;
}

function getStudentsByCourseYear(course, year) {
  if (!validateYear_(course, year)) return [];
  return STUDENTS_BY_COURSE[course][year];
}

// --- Date/row helpers ---

function getTodayDateStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

function findTodayRow_(sheet, name, course, year) {
  const data = sheet.getDataRange().getValues();
  const todayStr = getTodayDateStr_();
  const tz = Session.getScriptTimeZone();

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === name && data[i][1] === course && data[i][2] === year) {
      const cellDate = data[i][3];
      const cellDateStr = (cellDate instanceof Date)
        ? Utilities.formatDate(cellDate, tz, 'MM/dd/yyyy')
        : cellDate;

      if (cellDateStr === todayStr) {
        return { rowIndex: i + 1, rowData: data[i] };
      }
    }
  }
  return null;
}

// --- Core actions ---

function getStatus(course, year, name) {
  try {
    if (!validateName_(course, year, name)) {
      return { success: false, message: 'Invalid selection.', status: 'out' };
    }

    const sheet = getSheet();
    const found = findTodayRow_(sheet, name, course, year);

    if (!found) {
      return { success: true, status: 'out', session: null, message: 'No time-in yet today.' };
    }

    const r = found.rowData;
    const amIn = r[4], amOut = r[5], pmIn = r[7], pmOut = r[8];

    if (amIn && !amOut) return { success: true, status: 'in', session: 'AM' };
    if (pmIn && !pmOut) return { success: true, status: 'in', session: 'PM' };
    if (amIn && amOut && pmIn && pmOut) {
      return { success: true, status: 'done', session: null, message: 'Both AM and PM sessions completed today.' };
    }
    if (amIn && amOut && !pmIn) {
      return { success: true, status: 'out', session: null, message: 'AM session done. Ready for PM time-in.' };
    }
    return { success: true, status: 'out', session: null };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message, status: 'out' };
  }
}

function timeIn(course, year, name) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      return { success: false, message: 'System is busy, please try again.' };
    }
    if (!validateName_(course, year, name)) {
      return { success: false, message: 'Invalid selection.' };
    }

    const sheet = getSheet();
    const todayStr = getTodayDateStr_();
    const now = new Date();
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/dd/yyyy hh:mm a');

    const currentHour = now.getHours() + now.getMinutes() / 60;
    const isAMSession = currentHour < AM_PM_CUTOFF_HOUR;
    const session = isAMSession ? 'AM' : 'PM';

    const found = findTodayRow_(sheet, name, course, year);

    if (!found) {
      const lastRow = sheet.getLastRow() + 1;
      const rowData = [name, course, year, todayStr, '', '', '', '', '', '', '', '', ''];
      if (isAMSession) {
        rowData[4] = timeStr; // AM In
      } else {
        rowData[7] = timeStr; // PM In
      }
      sheet.getRange(lastRow, 1, 1, 13).setValues([rowData]);
      return { success: true, message: name + ' timed in (' + session + ') at ' + timeStr, status: 'in', session: session };
    }

    const rowIndex = found.rowIndex;
    const r = found.rowData;
    const amIn = r[4], amOut = r[5], pmIn = r[7], pmOut = r[8];

    if (isAMSession) {
      if (amIn && !amOut) {
        return { success: false, message: name + ' is already timed in (AM session).' };
      }
      if (amIn && amOut) {
        return { success: false, message: name + ' has already completed the AM session today.' };
      }
      sheet.getRange(rowIndex, 5).setValue(timeStr);
      return { success: true, message: name + ' timed in (AM) at ' + timeStr, status: 'in', session: 'AM' };
    } else {
      if (pmIn && !pmOut) {
        return { success: false, message: name + ' is already timed in (PM session).' };
      }
      if (pmIn && pmOut) {
        return { success: false, message: name + ' has already completed the PM session today.' };
      }
      sheet.getRange(rowIndex, 8).setValue(timeStr);
      return { success: true, message: name + ' timed in (PM) at ' + timeStr, status: 'in', session: 'PM' };
    }

  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function timeOut(course, year, name) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      return { success: false, message: 'System is busy, please try again.' };
    }
    if (!validateName_(course, year, name)) {
      return { success: false, message: 'Invalid selection.' };
    }

    const sheet = getSheet();
    const found = findTodayRow_(sheet, name, course, year);
    if (!found) {
      return { success: false, message: 'No active time-in found for today.' };
    }

    const rowIndex = found.rowIndex;
    const r = found.rowData;
    const amIn = r[4], amOut = r[5], pmIn = r[7], pmOut = r[8];
    const now = new Date();
    const timeOutStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/dd/yyyy hh:mm a');

    let session = null;
    if (amIn && !amOut) session = 'AM';
    else if (pmIn && !pmOut) session = 'PM';

    if (!session) {
      return { success: false, message: 'No active time-in found for today.' };
    }

    const inTime = new Date(session === 'AM' ? amIn : pmIn);
    let hours = (now - inTime) / (1000 * 60 * 60);
    if (hours < 0) hours = 0;

    if (session === 'AM') {
      sheet.getRange(rowIndex, 6).setValue(timeOutStr);
      sheet.getRange(rowIndex, 7).setValue(hours.toFixed(2));
    } else {
      sheet.getRange(rowIndex, 9).setValue(timeOutStr);
      sheet.getRange(rowIndex, 10).setValue(hours.toFixed(2));
    }

    const amH = parseFloat(sheet.getRange(rowIndex, 7).getValue()) || 0;
    const pmH = parseFloat(sheet.getRange(rowIndex, 10).getValue()) || 0;
    sheet.getRange(rowIndex, 11).setValue((amH + pmH).toFixed(2));

    // Always Pending after time-out — adviser must approve manually via Attendance Admin menu
    // (Approve/Reject menu ngayon ay gumagana kahit sa "History" sheet na, kaya
    // hindi na mawawala ang pagkakataon ma-approve kahit late na makapag-review.)
    sheet.getRange(rowIndex, 12).setValue('Pending');

    return {
      success: true,
      message: name + ' timed out (' + session + '). Hours: ' + hours.toFixed(2),
      status: 'out',
      session: session
    };

  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function updateSummary_(name, course, year) {
  let total = 0;

  // Approved hours mula sa Sheet1 (ngayong araw)
  const logSheet = getSheet();
  const data = logSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name && data[i][1] === course && data[i][2] === year && data[i][11] === 'Approved') {
      total += parseFloat(data[i][10]) || 0;
    }
  }

  // Approved hours mula sa Sheet3 (archive)
  const archived = getArchiveRecordsForStudent_(name, course, year);
  archived.forEach(function (r) {
    if (r.status === 'Approved') {
      total += parseFloat(r.totalHours) || 0;
    }
  });

  const summarySheet = getSummarySheet();
  const lastRow = summarySheet.getLastRow();
  let found = false;

  for (let i = 2; i <= lastRow; i++) {
    if (
      summarySheet.getRange(i, 1).getValue() === name &&
      summarySheet.getRange(i, 2).getValue() === course &&
      summarySheet.getRange(i, 3).getValue() === year
    ) {
      summarySheet.getRange(i, 4).setValue(total.toFixed(2));
      found = true;
      break;
    }
  }

  if (!found) {
    const nextRow = summarySheet.getLastRow() + 1;
    summarySheet.getRange(nextRow, 1, 1, 4).setValues([[name, course, year, total.toFixed(2)]]);
  }
}

// --- Attendance History ---

function getAttendanceHistory(course, year, name) {
  if (!validateName_(course, year, name)) return [];

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const history = [];

  function fmtDateTime_(val) {
    if (!val) return '';
    if (val instanceof Date) return Utilities.formatDate(val, tz, 'MM/dd/yyyy hh:mm a');
    return val;
  }

  function fmtDateOnly_(val) {
    if (!val) return '';
    if (val instanceof Date) return Utilities.formatDate(val, tz, 'MM/dd/yyyy');
    return val;
  }

  // Records ng ngayong araw (Sheet1)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name && data[i][1] === course && data[i][2] === year) {
      history.push({
        date: fmtDateOnly_(data[i][3]),
        amIn: fmtDateTime_(data[i][4]),
        amOut: fmtDateTime_(data[i][5]),
        amHours: data[i][6],
        pmIn: fmtDateTime_(data[i][7]),
        pmOut: fmtDateTime_(data[i][8]),
        pmHours: data[i][9],
        totalHours: data[i][10],
        status: data[i][11] || 'Pending',
        remarks: data[i][12]
      });
    }
  }

  // Archived records (Sheet3)
  const archived = getArchiveRecordsForStudent_(name, course, year);
  archived.forEach(function (r) {
    history.push({
      date: r.date,
      amIn: r.amIn, amOut: r.amOut, amHours: r.amHours,
      pmIn: r.pmIn, pmOut: r.pmOut, pmHours: r.pmHours,
      totalHours: r.totalHours,
      status: r.status || 'Pending',
      remarks: r.remarks
    });
  });

  history.sort(function (a, b) {
    return new Date(b.date) - new Date(a.date);
  });

  return history;
}

// --- Archive (Sheet3) ---

// Inililipat ang mga rows sa Sheet1 na HINDI petsa ngayon, papunta sa Sheet3 (grouped by date)
function archiveOldRows() {
  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(10000);

    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const todayStr = getTodayDateStr_();
    const tz = Session.getScriptTimeZone();

    const rowsToKeep = [data[0]];
    const rowsToArchive = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[0];
      if (!name) continue;

      const cellDate = row[3];
      const cellDateStr = (cellDate instanceof Date)
        ? Utilities.formatDate(cellDate, tz, 'MM/dd/yyyy')
        : cellDate;

      if (cellDateStr === todayStr) {
        rowsToKeep.push(row);
      } else {
        rowsToArchive.push({ dateStr: cellDateStr, row: row });
      }
    }

    if (rowsToArchive.length === 0) {
      return;
    }

    const groups = {};
    rowsToArchive.forEach(function (item) {
      if (!groups[item.dateStr]) groups[item.dateStr] = [];
      groups[item.dateStr].push(item.row);
    });

    const sortedDates = Object.keys(groups).sort(function (a, b) {
      return new Date(a) - new Date(b);
    });

    const archiveSheet = getArchiveSheet();
    archiveSheet.setFrozenColumns(0); // I-unfreeze muna bago mag-merge, para hindi mag-error
    const NAVY = '#1F2F6F';
    const headerRow = ['Student Assistant', 'Course', 'Year Level',
      'AM Time In', 'AM Time Out', 'AM Hours',
      'PM Time In', 'PM Time Out', 'PM Hours',
      'Total Hours', 'Status', 'Remarks'];

    sortedDates.forEach(function (dateStr) {
      const lastRow = archiveSheet.getLastRow();
      const startRow = lastRow === 0 ? 1 : lastRow + 2;

      archiveSheet.getRange(startRow, 1, 1, headerRow.length).merge();
      archiveSheet.getRange(startRow, 1).setValue('Date: ' + dateStr)
        .setBackground(NAVY).setFontColor('#FFFFFF').setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      archiveSheet.setRowHeight(startRow, 30);

      const tableHeaderRow = startRow + 1;
      archiveSheet.getRange(tableHeaderRow, 1, 1, headerRow.length).setValues([headerRow]);
      archiveSheet.getRange(tableHeaderRow, 1, 1, headerRow.length)
        .setFontWeight('bold').setBackground('#E8ECF2').setHorizontalAlignment('center');

      const dataRows = groups[dateStr].map(function (r) {
        return [r[0], r[1], r[2], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12]];
      });
      archiveSheet.getRange(tableHeaderRow + 1, 1, dataRows.length, headerRow.length).setValues(dataRows);

      const dataBlockRange = archiveSheet.getRange(tableHeaderRow + 1, 1, dataRows.length, headerRow.length);
      dataBlockRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    });

    const widths = [170, 70, 90, 130, 130, 75, 130, 130, 75, 90, 100, 160];
    widths.forEach(function (w, i) { archiveSheet.setColumnWidth(i + 1, w); });

    // Ilagay ang dropdown (Pending/Approved/Rejected) at conditional formatting
    // sa buong Status column (K) ng History sheet, para pwede na ring i-approve
    // dito nang direkta anumang oras — hindi lang sa Daily Log.
    applyArchiveStatusFormatting_(archiveSheet);

    sheet.clearContents();
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    sheet.getRange('A1:M1').setFontWeight('bold');

  } catch (err) {
    console.error('Archive error: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

// Naglalagay ng dropdown validation + conditional formatting sa Status column (K)
// ng History sheet, mula row 2 hanggang sa dulo ng laman.
function applyArchiveStatusFormatting_(archiveSheet) {
  const lastRow = Math.max(archiveSheet.getLastRow(), 2);
  const statusRange = archiveSheet.getRange(2, 11, lastRow - 1, 1); // Column K

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Rejected'], true)
    .setAllowInvalid(true) // payagan ang blangko sa "Date:"/header rows
    .build();
  statusRange.setDataValidation(rule);

  const existingRules = archiveSheet.getConditionalFormatRules();
  const nonStatusRules = existingRules.filter(function (r) {
    const ranges = r.getRanges();
    return !(ranges.length === 1 && ranges[0].getColumn() === 11);
  });

  const GREEN = '#2E5E20';
  const newRules = nonStatusRules.slice();
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Approved')
    .setBackground('#EEF9F0').setFontColor(GREEN).setBold(true)
    .setRanges([statusRange]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pending')
    .setBackground('#FFF9E8').setFontColor('#8A6D00').setBold(true)
    .setRanges([statusRange]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Rejected')
    .setBackground('#FFF1F1').setFontColor('#C0392B').setBold(true)
    .setRanges([statusRange]).build());
  archiveSheet.setConditionalFormatRules(newRules);
}

// Binabasa ang Sheet3 (na naka-table-per-date) para sa isang estudyante
function getArchiveRecordsForStudent_(name, course, year) {
  const archiveSheet = getArchiveSheet();
  const data = archiveSheet.getDataRange().getValues();
  const records = [];
  let currentDate = null;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const cellA = row[0];

    if (typeof cellA === 'string' && cellA.indexOf('Date: ') === 0) {
      currentDate = cellA.replace('Date: ', '').trim();
      continue;
    }
    if (cellA === 'Student Assistant') continue;
    if (!cellA) continue;

    if (row[0] === name && row[1] === course && row[2] === year) {
      records.push({
        date: currentDate,
        amIn: row[3], amOut: row[4], amHours: row[5],
        pmIn: row[6], pmOut: row[7], pmHours: row[8],
        totalHours: row[9], status: row[10], remarks: row[11]
      });
    }
  }
  return records;
}

function createDailyArchiveTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'archiveOldRows') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('archiveOldRows')
    .timeBased()
    .atHour(0)          // mag-a-archive sa may 12 AM - 1 AM
    .everyDays(1)
    .create();
}

// --- Auto time-out (Lunch 12PM / End of day 9PM) ---

function checkAutoTimeOuts() {
  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(10000);

    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const tz = Session.getScriptTimeZone();

    for (let i = 1; i < data.length; i++) {
      const name = data[i][0];
      if (!name) continue;

      const amIn = data[i][4], amOut = data[i][5];
      const pmIn = data[i][7], pmOut = data[i][8];
      const rowIndex = i + 1;
      let didClose = false;

      // AM session open
      if (amIn && !amOut) {
        const inDate = new Date(amIn);
        const noonCutoff = new Date(inDate.getFullYear(), inDate.getMonth(), inDate.getDate(), 12, 0, 0);
        const fivePmCutoff = new Date(inDate.getFullYear(), inDate.getMonth(), inDate.getDate(), END_OF_DAY_HOUR, 0, 0);

        let cutoff = null, remark = '';
        if (now >= fivePmCutoff) { cutoff = fivePmCutoff; remark = 'Auto Time Out (End of Day)'; }
        else if (now >= noonCutoff) { cutoff = noonCutoff; remark = 'Auto Time Out (Lunch Break)'; }

        if (cutoff) {
          let hours = (cutoff - inDate) / (1000 * 60 * 60);
          if (hours < 0) hours = 0;
          const outStr = Utilities.formatDate(cutoff, tz, 'MM/dd/yyyy hh:mm a');
          sheet.getRange(rowIndex, 6).setValue(outStr);
          sheet.getRange(rowIndex, 7).setValue(hours.toFixed(2));
          sheet.getRange(rowIndex, 12).setValue('Pending');
          sheet.getRange(rowIndex, 13).setValue(remark);
          didClose = true;
        }
      }

      // PM session open
      if (pmIn && !pmOut) {
        const inDate = new Date(pmIn);
        const fivePmCutoff = new Date(inDate.getFullYear(), inDate.getMonth(), inDate.getDate(), END_OF_DAY_HOUR, 0, 0);

        if (now >= fivePmCutoff) {
          let hours = (fivePmCutoff - inDate) / (1000 * 60 * 60);
          if (hours < 0) hours = 0;
          const outStr = Utilities.formatDate(fivePmCutoff, tz, 'MM/dd/yyyy hh:mm a');
          sheet.getRange(rowIndex, 9).setValue(outStr);
          sheet.getRange(rowIndex, 10).setValue(hours.toFixed(2));
          sheet.getRange(rowIndex, 12).setValue('Pending');
          sheet.getRange(rowIndex, 13).setValue('Auto Time Out (End of Day)');
          didClose = true;
        }
      }

      if (didClose) {
        const amH = parseFloat(sheet.getRange(rowIndex, 7).getValue()) || 0;
        const pmH = parseFloat(sheet.getRange(rowIndex, 10).getValue()) || 0;
        sheet.getRange(rowIndex, 11).setValue((amH + pmH).toFixed(2));
        // Not added to summary yet — stays Pending until adviser approves via the sheet menu.
      }
    }
  } catch (err) {
    console.error('Auto timeout error: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

function createAutoTimeOutTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'checkAutoTimeOuts') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('checkAutoTimeOuts')
    .timeBased()
    .everyMinutes(5)
    .create();
}

// --- Auto-run when Status dropdown is changed directly in the sheet ---
// Gumagana na ito ngayon sa "Daily Log" AT sa "History" sheet, kaya kahit
// direktang i-type/i-dropdown ng adviser ang "Approved"/"Rejected" sa History,
// awtomatikong mag-a-update ang Summary.
function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    if (sheetName !== SHEET_NAME && sheetName !== ARCHIVE_SHEET_NAME) return;

    const col = e.range.getColumn();
    const row = e.range.getRow();
    const statusCol = (sheetName === SHEET_NAME) ? 12 : 11;
    if (col !== statusCol || row < 2) return;

    const name = sheet.getRange(row, 1).getValue();
    if (!isRealDataRow_(name)) return;

    const course = sheet.getRange(row, 2).getValue();
    const year = sheet.getRange(row, 3).getValue();

    updateSummary_(name, course, year);
  } catch (err) {
    console.error('onEdit error: ' + err.message);
  }
}

// --- One-time visual setup (run manually once from the Apps Script editor) ---
function setupSheetDesign() {
  const sheet = getSheet();
  const summarySheet = getSummarySheet();
  const NAVY = '#1F2F6F';
  const GOLD = '#FFC700';
  const GREEN = '#2E5E20';

  // Header styling
  sheet.getRange('A1:M1')
    .setBackground(NAVY)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 34);

  summarySheet.getRange('A1:D1')
    .setBackground(NAVY)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  summarySheet.setRowHeight(1, 34);

  // Column widths
  const widths = [170, 70, 90, 90, 130, 130, 75, 130, 130, 75, 90, 100, 160];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });

  // Alternating row banding
  const dataRange = sheet.getRange('A2:M1000');
  const existingBandings = sheet.getBandings();
  existingBandings.forEach(function (b) { b.remove(); });
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  // Dropdown validation on Status column (L2:L1000)
  const statusRange = sheet.getRange('L2:L1000');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Rejected'], true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(rule);

  // Conditional formatting for Status column
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Approved')
    .setBackground('#EEF9F0').setFontColor(GREEN).setBold(true)
    .setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pending')
    .setBackground('#FFF9E8').setFontColor('#8A6D00').setBold(true)
    .setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Rejected')
    .setBackground('#FFF1F1').setFontColor('#C0392B').setBold(true)
    .setRanges([statusRange]).build());
  sheet.setConditionalFormatRules(rules);

  // Freeze header + first column
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  summarySheet.setFrozenRows(1);

  // --- Sheet3 (Archive) design ---
  // Hindi na natin ini-freeze ang column dito dahil may mga merged
  // "Date: ..." header row (mula A hanggang L) sa History sheet — kapag
  // sinubukang i-freeze ang isang column lang habang may merged cell na
  // sumasaklaw pa sa iba, mag-e-error ("can't freeze columns which
  // contain only part of a merged cell").
  const archiveSheet = getArchiveSheet();
  archiveSheet.setColumnWidths(1, 12, 120);
  archiveSheet.setColumnWidth(1, 170);

  // Approve/Reject dropdown + colors sa History sheet, para pwede na rin
  // itong gamitin anumang oras — hindi na kailangang i-approve lahat
  // bago mag-12 midnight.
  applyArchiveStatusFormatting_(archiveSheet);

  SpreadsheetApp.getUi().alert('Sheet design applied! Status column now has an Approve/Reject dropdown in both "Daily Log" and "History" sheets.');
}