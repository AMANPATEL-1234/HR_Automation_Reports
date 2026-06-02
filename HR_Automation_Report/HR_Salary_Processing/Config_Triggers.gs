// ════════════════════════════GLOBAL CONFIG — change sheet names here only════════════════════════════════
var CONFIG = {
  ATTENDANCE_SHEET        : "Attendance Record-Zoho",
  SALARY_SHEET            : "Draft Calculation",
  SALARY_PF_SHEET         : "Salary and PF Details",
  REIMBURSEMENT_SHEET     : "Reimbursement-Zoho",
  COMPENSATION_SHEET      : "Overtime Compensation-Zoho",
  FINAL_EMPLOYEE_SS_SHEET : "Salary Sheet",
  INTERN_DATA_SHEET       : "Intern Stipend Details",
  INTERN_ATTENDANCE_SHEET : "Intern Attendance Record-FloTrack",
  FINAL_INTERNS_SHEET     : "Stipend Sheet",
  FINAL_SS_BACKUP_SUBFOLDER_NAME      : "Salary Sheet",
  FINAL_INTERNS_BACKUP_SUBFOLDER_NAME : "Stipend Sheet",
  REIMBURSEMENT_IMAGE_SUBFOLDER_NAME  : "Reimbursement-Zoho",
  SALARY_REMARKS_SUBFOLDER_NAME       : "Draft Calculation",
  PROP_EMP_REGISTRY       : "EMP_REGISTRY",
  PROP_EMP_ORDER_CTR      : "EMP_ORDER_CTR",
  PROP_ATT_SNAPSHOT       : "ATT_SNAPSHOT",
  PROP_BACKUP_PARENT_FOLDER_ID: "Hr_backup_files"
};


function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("onSheetEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  ScriptApp.newTrigger("onSheetChange")
    .forSpreadsheet(ss)
    .onChange()
    .create();
}

// ════════onEdit ROUTER—only handles Salary & PF Details cell-level edits════════════
function onSheetEdit(e) {
  if (!e || !e.range) return;
  var sheetName = e.range.getSheet().getName();
  if (sheetName === CONFIG.SALARY_PF_SHEET) {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(8000)) return;
    try {
      handleSalaryPFEdit(e);
    } catch (err) {
      console.error("onSheetEdit error (PF): " + err.message);
    } finally {
      lock.releaseLock();
    }
  }
  //────── Route: Final Employee SS permanent columns ───────
  if (sheetName === CONFIG.FINAL_EMPLOYEE_SS_SHEET) {
    var lock2 = LockService.getScriptLock();
    if (!lock2.tryLock(8000)) return;
    try {
      handleFinalSSEdit(e);
    } catch (err) {
      console.error("onSheetEdit error (FinalSS): " + err.message);
    } finally {
      lock2.releaseLock();
    }
  }
  //*****************************************
  if (sheetName === CONFIG.SALARY_SHEET) {
  var lockSal = LockService.getScriptLock();
  if (!lockSal.tryLock(8000)) return;
  try {
    handleSalaryRemarksEdit(e);
  } catch (err) {
    console.error("onSheetEdit error (SalaryRemarks): " + err.message);
  } finally {
    lockSal.releaseLock();
  }
}
  //*****************************************
  if (sheetName === CONFIG.FINAL_INTERNS_SHEET) {
  var lock3 = LockService.getScriptLock();
  if (!lock3.tryLock(8000)) return;
  try {
    handleFinalInternsEdit(e);
  } catch (err) {
    console.error("onSheetEdit error (FinalInterns): " + err.message);
  } finally {
    lock3.releaseLock();
  }
}
}

// ═══════════════════════════════════════════════════════════════════════════════════════════════════════════════
function onSheetChange(e) {
  if (!e) return;
  var changeType = e.changeType || "";
  if (changeType !== "INSERT_ROW" &&
      changeType !== "DELETE_ROW" &&
      changeType !== "OTHER") return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;
  try {
    syncAttendanceChanges();
  } catch (err) {
    console.error("onSheetChange error: " + err.message);
  } finally {
    lock.releaseLock();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("HR All Automation")
    .addItem("Set Month & Year", "showMonthYearPicker")
    .addItem("Intern Sheet Selection", "showInternSheetPicker")
    .addItem("Run Employee Master", "master")
    .addItem("Run Intern Master", "internMaster")
    .addToUi();
}