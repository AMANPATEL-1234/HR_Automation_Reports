function syncAttendanceChanges() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var attSheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if (!attSheet) return;
  var lastRow          = attSheet.getLastRow();
  var currentEmployees = {};
  if (lastRow >= 2) {
    var attData = attSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < attData.length; i++) {
      var id   = String(attData[i][0] || "").trim();
      var name = String(attData[i][1] || "").trim();
      if (id) currentEmployees[id] = name;
    }
  }

  var props    = PropertiesService.getScriptProperties();
  var raw      = props.getProperty(CONFIG.PROP_ATT_SNAPSHOT);
  var snapshot = raw ? JSON.parse(raw) : {};
  var addedIds   = [];
  var deletedIds = [];
  for (var cid in currentEmployees) {
    if (!snapshot.hasOwnProperty(cid)) addedIds.push(cid);
  }
  for (var sid in snapshot) {
    if (!currentEmployees.hasOwnProperty(sid)) deletedIds.push(sid);
  }

  if (addedIds.length === 0 && deletedIds.length === 0) return; 
  // Process deletions first (cleaner state before additions)
  if (deletedIds.length > 0) {
    deletedIds.forEach(function(empId) {
      removeEmployeeFromAllSheets(ss, empId);
    });
  }

  // Process additions
  if (addedIds.length > 0) {
    addedIds.forEach(function(empId) {
      addEmployeeToAllSheets(ss, empId, currentEmployees[empId]);
    });
  }

  SpreadsheetApp.flush();
  // Save updated snapshot
  props.setProperty(CONFIG.PROP_ATT_SNAPSHOT, JSON.stringify(currentEmployees));
  var msg = [];
  if (addedIds.length   > 0) msg.push("Added: "   + addedIds.join(", "));
  if (deletedIds.length > 0) msg.push("Removed: " + deletedIds.join(", "));
  ss.toast(msg.join(" | "), "Employee Sync ✔", 5);
}


//═════Reimbursement & Compensation are Zoho-sourced,rows appear when master()fetches from Zoho,not via manual attendance edits═════
function addEmployeeToAllSheets(ss, empId, empName) {
  var salSheet = ss.getSheetByName(CONFIG.SALARY_SHEET);
  if (salSheet && salSheet.getLastRow() >= 1) {
    var salLastRow = salSheet.getLastRow();
    var salAlready = false;
    if (salLastRow >= 2) {
      var salIds = salSheet.getRange(2, 2, salLastRow - 1, 1).getValues();
      for (var s = 0; s < salIds.length; s++) {
        if (String(salIds[s][0] || "").trim() === empId) { salAlready = true; break; }
      }
    }
    if (!salAlready) {
      var newRow = salLastRow + 1;
      salSheet.getRange(newRow, 1).setValue("");       
      salSheet.getRange(newRow, 2).setValue(empId);    
      salSheet.getRange(newRow, 3).setValue(empName);  
      applySalaryFormulas(salSheet, newRow);            
    }
  }
}

//REMOVE employee from ALL dependent sheets,Sheets: Salary & PF Details,Salary sheet,Reimbursement,Compensation
// ═══════════════════════════════════════════════════════════════════════════
function removeEmployeeFromAllSheets(ss, empId) {
  deleteRowsByEmpId(ss, CONFIG.SALARY_SHEET,        empId,  1); 
  deleteRowsByEmpId(ss, CONFIG.REIMBURSEMENT_SHEET, empId, -1);
  deleteRowsByEmpId(ss, CONFIG.COMPENSATION_SHEET,  empId, -1);
}

//DELETE rows matching empId in a sheet->empColIdx: 0-based column index, or -1 to auto-detect via header lookup
// ═══════════════════════════════════════════════════════════════════════════
function deleteRowsByEmpId(ss, sheetName, empId, empColIdx) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  if (empColIdx === -1) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    empColIdx = getColIndex(headers, [
      "Employee ID", "Employee Id", "Employee HRM", "EmpID", "Emp ID"
    ]);
    if (empColIdx === -1) return;
  }

  var colData  = sheet.getRange(2, empColIdx + 1, lastRow - 1, 1).getValues();
  var toDelete = [];
  for (var r = 0; r < colData.length; r++) {
    var id = String(colData[r][0] || "").trim();
    var numId    = empId.replace(/\D/g, "");
    var numRowId = id.replace(/\D/g, "");
    if (id === empId || (numId !== "" && numRowId === numId)) {
      toDelete.push(r + 2); // 1-based row (header = row 1)
    }
  }
  // Delete bottom-up so row numbers don't shift
  toDelete.sort(function(a, b) { return b - a; });
  toDelete.forEach(function(rn) { sheet.deleteRow(rn); });
}
// ═════════════════════════════════════════════════════════════════════════════════════════
function initSnapshot() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var attSheet=ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if (!attSheet){ss.toast("Attendance sheet not found!","Error",5);return;}
  var lastRow=attSheet.getLastRow(),snapshot={};
  if (lastRow>=2){
    var data=attSheet.getRange(2,1,lastRow-1,2).getValues();
    data.forEach(function(row){var id=String(row[0]||"").trim(),name=String(row[1]||"").trim();if(id)snapshot[id]=name;});
  }
  PropertiesService.getScriptProperties().setProperty(CONFIG.PROP_ATT_SNAPSHOT,JSON.stringify(snapshot));
  ss.toast("Snapshot seeded: "+Object.keys(snapshot).length+" employees ✔","initSnapshot",4);
}

//Run from Apps Script UI for a manual forced sync
function manualSync() {
  var lock=LockService.getScriptLock();
  if (!lock.tryLock(15000))return;
  try{syncAttendanceChanges();}finally{lock.releaseLock();}
}


// ═════════════════════════════════ EMPLOYEE REGISTRY═══════════════════════════════════════════════
function loadRegistry() {
  var raw = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_EMP_REGISTRY);
  return raw ? JSON.parse(raw) : {};
}
function saveRegistry(registry) {
  PropertiesService.getScriptProperties().setProperty(CONFIG.PROP_EMP_REGISTRY, JSON.stringify(registry));
}
function getNextOrder() {
  var props = PropertiesService.getScriptProperties();
  var ctr   = parseInt(props.getProperty(CONFIG.PROP_EMP_ORDER_CTR) || "0", 10);
  ctr++;
  props.setProperty(CONFIG.PROP_EMP_ORDER_CTR, String(ctr));
  return ctr;
}





















