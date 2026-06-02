//**************************************************
var SALARY_REMARKS_BACKUP_PARENT_FOLDER_ID = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_BACKUP_PARENT_FOLDER_ID);
var SALARY_REMARKS_BACKUP_SUBFOLDER_NAME = CONFIG.SALARY_REMARKS_SUBFOLDER_NAME;

function getSalaryRemarksBackupFolder_() {
  var hrDataFolder = getHrDataFolder_();
  var folders = hrDataFolder.getFoldersByName(SALARY_REMARKS_BACKUP_SUBFOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return hrDataFolder.createFolder(SALARY_REMARKS_BACKUP_SUBFOLDER_NAME);
}

function getSalaryRemarksBackupFileName_(year, month) {
  return "Draft_Calculation_remarks_" + year + "-" + pad2_(Number(month) + 1) + ".json";
}

function readSalaryRemarksBackupFromDrive_(year, month) {
  var folder = getSalaryRemarksBackupFolder_();
  var fileName = getSalaryRemarksBackupFileName_(year, month);
  var files = folder.getFilesByName(fileName);

  if (!files.hasNext()) {
    return {};
  }

  var file = files.next();
  var text = file.getBlob().getDataAsString();
  if (!text || String(text).trim() === "") {
    return {};
  }

  var parsed = JSON.parse(text);

  if (parsed.remarks) {
    return parsed.remarks;
  }

  return parsed;
}

//************************************************
function writeSalaryRemarksBackupToDrive_(year, month, remarksBackup) {
  var folder = getSalaryRemarksBackupFolder_();
  var fileName = getSalaryRemarksBackupFileName_(year, month);
  var files = folder.getFilesByName(fileName);

  var monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  var payload = {
    year: Number(year),
    month: Number(month) + 1,      // human month: April = 4
    monthIndex: Number(month),     // script month: April = 3
    monthName: monthNames[Number(month)],
    updatedAt: new Date().toISOString(),
    remarks: remarksBackup || {}
  };

  var json = JSON.stringify(payload, null, 2);

  if (files.hasNext()) {
    var file = files.next();
    file.setContent(json);
    return file.getUrl();
  }

  var newFile = folder.createFile(fileName, json, MimeType.PLAIN_TEXT);
  return newFile.getUrl();
}

function setSalaryVisibleMonthYear_(sheet, year, month) {
  if (!sheet) return;

  sheet.getRange(1, 1).setNote(JSON.stringify({
    year: Number(year),
    month: Number(month) + 1,      // human month
    monthIndex: Number(month)      // script month
  }));
}

function getSalaryVisibleMonthYear_(sheet) {
  if (sheet) {
    var note = sheet.getRange(1, 1).getNote();

    if (note) {
      try {
        var obj = JSON.parse(note);

        if (obj.year !== undefined) {
          if (obj.monthIndex !== undefined) {
            return {
              year: Number(obj.year),
              month: Number(obj.monthIndex)
            };
          }

          if (obj.month !== undefined) {
            return {
              year: Number(obj.year),
              month: Number(obj.month) - 1
            };
          }
        }
      } catch (err) {}
    }
  }

  var my = getAttendanceMonthYearKey();
  return {
    year: Number(my.year),
    month: Number(my.month)
  };
}

//═════════════════════════════════════════════════════════════════════════════════════════════════════════
function saveSalaryRemarksData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SALARY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  var headers = data[0];

  var empIdIdx = getColIndex(headers, [
    "Employee ID",
    "Employee Id",
    "Employee HRM",
    "Emp ID",
    "EmpID"
  ]);

  var remarksIdx = getColIndex(headers, ["Remarks"]);

  if (empIdIdx === -1 || remarksIdx === -1) {
    console.log("saveSalaryRemarksData skipped: Employee ID or Remarks column missing");
    return;
  }

  var my = getSalaryVisibleMonthYear_(sheet);
  var year = my.year;
  var month = my.month;
  var saved = readSalaryRemarksBackupFromDrive_(year, month);
  for (var i = 1; i < data.length; i++) {
    var empId = String(data[i][empIdIdx] || "").trim();
    if (!empId) continue;

    saved[empId] = valueOrBlankSalaryRemarks_(data[i][remarksIdx]);
  }

  var url = writeSalaryRemarksBackupToDrive_(year, month, saved);
  console.log("Salary Remarks JSON backup saved for " + year + "-" + pad2_(month + 1) +" employees: " + Object.keys(saved).length +" file: " + url);
}

//******************************************************
var FINAL_SS_BACKUP_PARENT_FOLDER_ID = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_BACKUP_PARENT_FOLDER_ID);
var FINAL_SS_BACKUP_SUBFOLDER_NAME = CONFIG.FINAL_SS_BACKUP_SUBFOLDER_NAME;
function getFinalSSBackupFolder_() {
  var hrDataFolder = getHrDataFolder_();
  var folders = hrDataFolder.getFoldersByName(FINAL_SS_BACKUP_SUBFOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return hrDataFolder.createFolder(FINAL_SS_BACKUP_SUBFOLDER_NAME);
}

// month is zero-based: Jan=0, Feb=1, Mar=2
function getFinalSSBackupFileName_(year, month) {
  return "Final_Employee_SS_" + year + "-" + pad2_(Number(month) + 1) + ".json";
}

//*******************************************
function readFinalSSBackupFromDrive_(year, month) {
  var folder = getFinalSSBackupFolder_();
  var fileName = getFinalSSBackupFileName_(year, month);
  var files = folder.getFilesByName(fileName);

  if (!files.hasNext()) {
    return {};
  }

  var file = files.next();
  var text = file.getBlob().getDataAsString();

  if (!text || String(text).trim() === "") {
    return {};
  }

  var parsed = JSON.parse(text);
  // New format
  if (parsed.employees) {
    return parsed.employees;
  }
  // Fallback if old direct object format is ever used
  return parsed;
}

//**********************************************
function writeFinalSSBackupToDrive_(year, month, employeesBackup) {
  var folder = getFinalSSBackupFolder_();
  var fileName = getFinalSSBackupFileName_(year, month);
  var files = folder.getFilesByName(fileName);

  var monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  var payload = {
    year: Number(year),
    month: Number(month) + 1,      // human month: April = 4
    monthIndex: Number(month),     // script month: April = 3
    monthName: monthNames[Number(month)],
    updatedAt: new Date().toISOString(),
    employees: employeesBackup || {}
};

  var json = JSON.stringify(payload, null, 2);

  if (files.hasNext()) {
    var file = files.next();
    file.setContent(json);
    return file.getUrl();
  }

  var newFile = folder.createFile(fileName, json, MimeType.PLAIN_TEXT);
  return newFile.getUrl();
}

function setFinalSSVisibleMonthYear_(sheet, year, month) {
  if (!sheet) return;
  sheet.getRange(1, 1).setNote(JSON.stringify({
  year: Number(year),
  month: Number(month) + 1,      // human month: April = 4
  monthIndex: Number(month)      // script month: April = 3
}));
}

//**********************************************
function getFinalSSVisibleMonthYear_(sheet) {
  var props = PropertiesService.getScriptProperties();
  if (sheet) {
    var note = sheet.getRange(1, 1).getNote();
    if (note) {
      try {
        var obj = JSON.parse(note);
        if (obj.year !== undefined) {
        if (obj.monthIndex !== undefined) {
          return {
            year: Number(obj.year),
            month: Number(obj.monthIndex) // use script month internally
          };
        }

        if (obj.month !== undefined) {
          return {
            year: Number(obj.year),
            month: Number(obj.month) - 1 // convert human month back to script month
          };
        }
      }
      } catch (err) {}
    }
  }

  // Fallback only for first run when old sheet has no note
  var selectedMonth = props.getProperty("SELECTED_MONTH");
  var selectedYear = props.getProperty("SELECTED_YEAR");

  if (selectedMonth !== null && selectedYear !== null) {
    return {
      year: Number(selectedYear),
      month: Number(selectedMonth)
    };
  }

  var today = new Date();
  return {
    year: today.getFullYear(),
    month: today.getMonth()
  };
}

// ═════════════════════════════════saveFinalSSData══════════════════════════════════════════════════
function saveFinalSSData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.FINAL_EMPLOYEE_SS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  var hasRealData = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || "").trim() !== "") {
      hasRealData = true;
      break;
    }
  }

  if (!hasRealData) {
    // console.log("saveFinalSSData skipped: no employee data");
    return;
  }

  var headers = data[0];

  var empIdCol = -1;
  for (var h = 0; h < headers.length; h++) {
    var hn = String(headers[h]).toLowerCase().replace(/\s+/g, "");
    if (hn === "employeeid") {
      empIdCol = h;
      break;
    }
  }

  if (empIdCol === -1) return;

  var my = getFinalSSVisibleMonthYear_(sheet);
  var year = my.year;
  var month = my.month;
  var saved = readFinalSSBackupFromDrive_(year, month);
  for (var r = 1; r < data.length; r++) {
    var empId = String(data[r][empIdCol] || "").trim();
    if (!empId) continue;

    if (!saved[empId]) {
      saved[empId] = {
        oneTime: "",
        arrears: "",
        referral: "",
        tds: "",
        advance: "",
        status: "",
        remarks: ""
      };
    }

    var oneTimeVal = valueOrBlankFinalSS_(data[r][12]);
    var arrearsVal = valueOrBlankFinalSS_(data[r][13]);
    var referralVal = valueOrBlankFinalSS_(data[r][15]);
    var tdsVal = valueOrBlankFinalSS_(data[r][18]);
    var advanceVal = valueOrBlankFinalSS_(data[r][19]);
    var statusVal = valueOrBlankFinalSS_(data[r][21]);
    var remarksVal = valueOrBlankFinalSS_(data[r][22]);

    if (oneTimeVal !== "") saved[empId].oneTime = oneTimeVal;
    if (arrearsVal !== "") saved[empId].arrears = arrearsVal;
    if (referralVal !== "") saved[empId].referral = referralVal;
    if (tdsVal !== "") saved[empId].tds = tdsVal;
    if (advanceVal !== "") saved[empId].advance = advanceVal;
    if (statusVal !== "") saved[empId].status = statusVal;
    if (remarksVal !== "") saved[empId].remarks = remarksVal;
      }

  var url = writeFinalSSBackupToDrive_(year, month, saved);
  console.log("Final Employee SS JSON backup saved for " +year + "-" + pad2_(month + 1) +" employees: " + Object.keys(saved).length +" file: "
);
}

// ═════════════════════════════════ Final SS — cell edit auto-save ══════════════════════════════
function handleFinalSSEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.FINAL_EMPLOYEE_SS_SHEET) return;

  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();
  var startRow = e.range.getRow();
  var endRow = e.range.getLastRow();

  if (startRow === 1 && endRow === 1) return;
  if (startRow < 2) startRow = 2;

  var permanentCols = [13, 14, 16, 19, 20, 22, 23];

  var colsToProcess = [];
  for (var c = startCol; c <= endCol; c++) {
    if (permanentCols.indexOf(c) !== -1) {
      colsToProcess.push(c);
    }
  }

  if (colsToProcess.length === 0) return;

  var numRows = endRow - startRow + 1;
  var empIds = sheet.getRange(startRow, 2, numRows, 1).getValues();

  var colData = {};
  for (var ci = 0; ci < colsToProcess.length; ci++) {
    var col = colsToProcess[ci];

    colData[col] = {
      formulas: sheet.getRange(startRow, col, numRows, 1).getFormulas(),
      values: sheet.getRange(startRow, col, numRows, 1).getValues()
    };
  }

  var my = getFinalSSVisibleMonthYear_(sheet);
  var year = my.year;
  var month = my.month;

  var backup = readFinalSSBackupFromDrive_(year, month);

  for (var r = 0; r < numRows; r++) {
    var empId = String(empIds[r][0] || "").trim();
    if (!empId) continue;

    if (!backup[empId]) {
      backup[empId] = {
          oneTime: "",
          arrears: "",
          referral: "",
          tds: "",
          advance: "",
          status: "",
          remarks: ""
        };
    }

    for (var x = 0; x < colsToProcess.length; x++) {
      var colNum = colsToProcess[x];
      var formula = colData[colNum].formulas[r][0];
      var val = colData[colNum].values[r][0];
      var toSave = formula !== "" ? formula : valueOrBlankFinalSS_(val);

      if (colNum === 13) backup[empId].oneTime = toSave;
      else if (colNum === 14) backup[empId].arrears = toSave;
      else if (colNum === 16) backup[empId].referral = toSave;
      else if (colNum === 19) backup[empId].tds = toSave;
      else if (colNum === 20) backup[empId].advance = toSave;
      else if (colNum === 22) backup[empId].status = toSave;
      else if (colNum === 23) backup[empId].remarks = toSave;
    }
  }

  writeFinalSSBackupToDrive_(year, month, backup);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✔ Final Employee SS saved to Drive JSON for " + year + "-" + pad2_(month + 1),
    "Auto-saved",
    2
  );
}

function pad2_(n) {
  return String(n).padStart(2, "0");
}


function valueOrBlankSalaryRemarks_(v) {
  if (v === null || v === undefined || v === "") {
    return "";
  }
  return v;
}

function valueOrBlankFinalSS_(v) {
  if (v === null || v === undefined || v === "") {
    return "";
  }
  return v; // keeps 0 as 0
}







































