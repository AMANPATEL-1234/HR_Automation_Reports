// ═════════════════════════════════BUILD SALARY SHEET═════════════════════════════════════════════════
function buildSalarySheet(attendanceSheetName, reimbursementSheetName, compensationSheetName) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var attSheet=ss.getSheetByName(attendanceSheetName);
  var reimbSheet=ss.getSheetByName(reimbursementSheetName);
  var compSheet=ss.getSheetByName(compensationSheetName);
  var pfSheet=ss.getSheetByName(CONFIG.SALARY_PF_SHEET);
  if (!attSheet) throw new Error("Sheet not found: "+attendanceSheetName);
  if (!pfSheet)  throw new Error("Sheet not found: "+CONFIG.SALARY_PF_SHEET);

  var outName = CONFIG.SALARY_SHEET;
  var sheet   = ss.getSheetByName(outName);

  var props = PropertiesService.getScriptProperties();
  var salaryMY = getAttendanceMonthYearKey();
  var salaryBackupKey = salaryMY.key;

// Mark Salary sheet as now belonging to this attendance month.
  props.setProperty("SALARY_ACTIVE_MONTH", String(salaryMY.month));
  props.setProperty("SALARY_ACTIVE_YEAR", String(salaryMY.year));

  if (!sheet){sheet=ss.insertSheet(outName);}else{sheet.clearContents();sheet.clearFormats();}
  var headers=["Entities","Employee ID","Employee Name","Expected Payable Day(s)","Weekends","Mandatory Off(s)","Expected Worked Day(s)","Total Present","Present","On Duty","Total Paid Leaves","Earned Leaves","Sick Leaves","Casual Leaves","Compensatory Off","Optional WFH","Maternity Leaves","Paternity Leaves","Unpaid Leaves","Total Payable Days","Days before joining / after leaving","Monthly Fixed","Compensation request","Reimbursement","ESI/PF","Remarks"];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  var attVals=attSheet.getDataRange().getValues();
  var pfVals=pfSheet.getDataRange().getValues();
  var compVals=compSheet?compSheet.getDataRange().getValues():[];

  var reimbVals=reimbSheet?reimbSheet.getDataRange().getValues():[];
  if (attVals.length<2) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "⚠️ No attendance data — salary sheet not built.",
      "Warning", 6
    );
    return {sheetName: outName, rowsWritten: 0};
  }
  if (pfVals.length<2)  throw new Error("Salary and PF Details sheet has no data");

  function nk(s){return String(s||"").toLowerCase().replace(/\s+/g,"").replace(/[^a-z0-9]/g,"");}
  function gci(h,names){var n=h.map(function(x){return nk(x);});for(var i=0;i<names.length;i++){var x=n.indexOf(nk(names[i]));if(x!==-1)return x;}return -1;}
  function pn(v){if(v===null||v===undefined||v==="")return 0;if(typeof v==="number")return v;var n=parseFloat(String(v).replace(/,/g,"").trim());return isNaN(n)?0:n;}
  function toObjs(vals){if(!vals||vals.length<2)return[];var h=vals[0];return vals.slice(1).map(function(r){var o={};h.forEach(function(c,i){o[String(c).trim()]=r[i];});return o;});}
  function sumByEmp(rows,idN,amtN){
    var out={};
    rows.forEach(function(obj){
      var keys=Object.keys(obj);
      var ik=keys.find(function(k){return idN.some(function(n){return nk(k)===nk(n);});});
      var ak=keys.find(function(k){return amtN.some(function(n){return nk(k)===nk(n);});});
      if(!ik||!ak)return;
      var id=String(obj[ik]||"").trim();if(!id)return;
      out[id]=(out[id]||0)+pn(obj[ak]);
    });
    return out;
  }
  function firstAmountByEmp(rows, idN, amtN) {
  var out = {};
  var seen = {};
  rows.forEach(function(obj) {
    var keys = Object.keys(obj);
    var ik = keys.find(function(k) {
      return idN.some(function(n) { return nk(k) === nk(n); });
    });

    var ak = keys.find(function(k) {
      return amtN.some(function(n) { return nk(k) === nk(n); });
    });

    if (!ik || !ak) return;
    // Only take approved reimbursement rows
    var approvalStatus = String(
      obj["Approval Status"] ||
      obj["ApprovalStatus"] ||
      obj["approval status"] ||
      ""
    ).trim().toLowerCase();
    if (approvalStatus !== "approved") {
      return;
    }

    var id = String(obj[ik] || "").trim();
    if (!id) return;
    var payDate = String(
      obj["Date of Payment"] ||
      obj["Date of payment"] ||
      ""
    ).trim();

    var amount = String(obj[ak] || "").trim();
    var dedupKey = id + "|" + payDate + "|" + amount;

    if (seen[dedupKey]) return;
    seen[dedupKey] = true;

    out[id] = (out[id] || 0) + pn(obj[ak]);
  });

  return out;
}

  var uniq=[],seen={};
  for (var i=1;i<attVals.length;i++){
    var eid=String(attVals[i][0]||"").trim(),en=String(attVals[i][1]||"").trim();
    if (!eid||seen[eid])continue;seen[eid]=true;uniq.push([eid,en]);
  }
  if (uniq.length===0) throw new Error("No employees found in Attendance sheet");
  sheet.getRange(2,2,uniq.length,2).setValues(uniq);

  var ph=pfVals[0];
  var pei=gci(ph,["Entities","Entity"]),pid=gci(ph,["Employee ID","Employee Id","Employee HRM"]);
  var pmi=gci(ph,["Monthly Fixed"]),ppi=gci(ph,["PF"]);
  if (pei===-1||pid===-1||pmi===-1||ppi===-1) throw new Error("Salary and PF Details must have: Entities, Employee ID, Monthly Fixed, PF");

  var salMap={};
  for (var r=1;r<pfVals.length;r++){
    var row=pfVals[r],id=String(row[pid]||"").trim();
    if (!id)continue;
    salMap[id]={entity:row[pei]||"",monthlyFixed:pn(row[pmi]),pf:pn(row[ppi])};
  }

  var compByEmp=sumByEmp(toObjs(compVals),["Employee HRM","Employee ID","Employee Id"],["Total Amount","Amount","Compensation Amount"]);
  var reimbByEmp = firstAmountByEmp(
  toObjs(reimbVals),
  ["Employee HRM", "Employee ID", "Employee Id"],
  ["Total Amount"]
);
  var rc=uniq.length,ca=[],cv=[],cx=[],caa=[],bl=[];
  for (var j=0;j<rc;j++){
    var sid=uniq[j][0],sal=salMap[sid]||{};
    ca.push([sal.entity||""]);cv.push([sal.monthlyFixed||0]);
    cx.push([compByEmp[sid]||0]);caa.push([reimbByEmp[sid]||0]);bl.push([""]);
  }
  sheet.getRange(2,1,rc,1).setValues(ca);
  sheet.getRange(2,22,rc,1).setValues(cv);
  sheet.getRange(2,24,rc,1).setValues(caa);  // Reimbursement
  sheet.getRange(2,26,rc,1).setValues(bl.map(function(){return[""];}));

  for (var rowNum=2;rowNum<=rc+1;rowNum++) applySalaryFormulas(sheet,rowNum);
  // Restore month-wise Remarks from Drive JSON
  var monthRemarksBackup = readSalaryRemarksBackupFromDrive_(salaryMY.year, salaryMY.month);
  var remColNum = headers.indexOf("Remarks") + 1;
  for (var j = 0; j < uniq.length; j++) {
    var eid = uniq[j][0];
    var remarkValue = monthRemarksBackup[eid];
    if (remarkValue !== undefined && remarkValue !== null && remarkValue !== "") {
      sheet.getRange(j + 2, remColNum).setValue(remarkValue);
    }
  }
  sheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#d9ead3").setVerticalAlignment("middle");
  sheet.getRange(1,1,1,3).setHorizontalAlignment("left");
  if (headers.length>3) sheet.getRange(1,4,1,headers.length-3).setHorizontalAlignment("center");
  if (rc>0){
    sheet.getRange(2,1,rc,3).setHorizontalAlignment("left");
    if (headers.length>3) sheet.getRange(2,4,rc,headers.length-3).setHorizontalAlignment("center");
    sheet.getRange(2,1,rc,headers.length).setVerticalAlignment("middle");
  }
  sheet.setFrozenRows(1);sheet.autoResizeColumns(1,headers.length);
  sheet.getRange(1,1,Math.max(sheet.getLastRow(),1),headers.length).setBorder(true,true,true,true,true,true);
   // Highlight Expected Worked Day(s) column G if employee has OD/P or P/OD in Attendance
var attStatusMap = {};
// Build employee-wise attendance status list from Attendance record-Zoho
for (var ar = 1; ar < attVals.length; ar++) {
  var empIdAtt = String(attVals[ar][0] || "").trim();
  if (!empIdAtt) continue;
  var hasODP = false;
  for (var ac = 2; ac < attVals[ar].length; ac++) {
  var statusVal = String(attVals[ar][ac] || "").trim().toUpperCase();
  var allowedSpecialCases = [
    "W/P", "P/W",
    "H/P", "P/H",
    "P/OD", "OD/P"
  ];

  if (
    statusVal === "SL/0.5P" ||
    statusVal === "0.5P/SL" ||
    (
      /^[A-Z]+\/[A-Z]+$/.test(statusVal) &&
      allowedSpecialCases.indexOf(statusVal) === -1
    )
  ) {
    hasODP = true;
    break;
  }
}
  attStatusMap[empIdAtt] = hasODP;
}
// Apply background to Expected Worked Day(s), column G
var expectedWorkedBgs = [];
for (var er = 0; er < uniq.length; er++) {
  var empIdSalary = String(uniq[er][0] || "").trim();

  if (attStatusMap[empIdSalary]) {
    expectedWorkedBgs.push(["#f4cccc"]); 
  } else {
    expectedWorkedBgs.push(["#ffffff"]); 
  }
}

if (expectedWorkedBgs.length > 0) {
  sheet.getRange(2, 7, expectedWorkedBgs.length, 1).setBackgrounds(expectedWorkedBgs);
}
// Save the freshly restored remarks state for this month.
  setSalaryVisibleMonthYear_(sheet, salaryMY.year, salaryMY.month);
  return {sheetName: outName, rowsWritten: rc};
}

//════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════
function applySalaryFormulas(sheet, rowNum) {
  var props     = PropertiesService.getScriptProperties();
  var today     = new Date();
  var _month    = props.getProperty("SELECTED_MONTH");
  var _year     = props.getProperty("SELECTED_YEAR");
  var calcMonth = _month ? parseInt(_month) : today.getMonth();
  var calcYear  = _year  ? parseInt(_year)  : today.getFullYear();
  var totalDays = new Date(calcYear, calcMonth + 1, 0).getDate();
  var ATT       = "'" + CONFIG.ATTENDANCE_SHEET + "'";
  var R         = String(rowNum);

  sheet.getRange(rowNum,  4).setFormula("=IF(B"+R+"=\"\",\"\",DAY(EOMONTH("+ATT+"!$C$1,0))-U"+R+")");
  sheet.getRange(rowNum,  5).setFormula("=IF(B"+R+"=\"\",\"\",SUMPRODUCT(("+ATT+"!$C$1:$AG=\"W\")*("+ATT+"!$A$1:$A=B"+R+")))");
  sheet.getRange(rowNum,  6).setFormula("=IF(B"+R+"=\"\",\"\",SUMPRODUCT(("+ATT+"!$C$1:$AG=\"H\")*("+ATT+"!$A$1:$A=B"+R+")))");
  sheet.getRange(rowNum,  7).setFormula("=IF(B"+R+"=\"\",\"\",D"+R+"-E"+R+"-F"+R+")");
  sheet.getRange(rowNum,  8).setFormula("=IF(B"+R+"=\"\",\"\",I"+R+"+J"+R+")");
  sheet.getRange(rowNum,  9).setFormula(
    "=IF(B"+R+"=\"\",\"\",SUMPRODUCT((" +
    "IF(EXACT(\"P\","+ATT+"!$C$1:$AG),1,0)+" +
    "IF(ISNUMBER(SEARCH(\"OD/P\","+ATT+"!$C$1:$AG)),1,0)+" +
    "IF(ISNUMBER(SEARCH(\"H/P\","+ATT+"!$C$1:$AG)),1,0)+" +
    "IF(ISNUMBER(SEARCH(\"W/P\","+ATT+"!$C$1:$AG)),1,0)+" +
    "IF(ISNUMBER(SEARCH(\"0.5OWH/0.5P\","+ATT+"!$C$1:$AG)),0.5," +
    "IF(ISNUMBER(SEARCH(\"H/0.5P\","+ATT+"!$C$1:$AG)),0.5," +
    "IF(ISNUMBER(SEARCH(\"0.5P/\","+ATT+"!$C$1:$AG)),0.5," +
    "IF(ISNUMBER(SEARCH(\"/0.5P\","+ATT+"!$C$1:$AG)),0.5,0))))" +
    ")*--EXACT("+ATT+"!$A$1:$A,B"+R+")))"
  );
  sheet.getRange(rowNum, 10).setFormula(
  "=IF(B"+R+"=\"\",\"\",SUMPRODUCT((" +
  "IF(EXACT(\"OD\","+ATT+"!$C$1:$AG),1,0)+" +
  "IF(ISNUMBER(SEARCH(\"0.5OD\","+ATT+"!$C$1:$AG)),0.5,0)+" +
  "IF(ISNUMBER(SEARCH(\"OD/\","+ATT+"!$C$1:$AG))*" +
  "NOT(ISNUMBER(SEARCH(\"OD/A\","+ATT+"!$C$1:$AG)))*" +
  "NOT(ISNUMBER(SEARCH(\"OD/P\","+ATT+"!$C$1:$AG))),0.5,0)" +
  ")*EXACT("+ATT+"!$A$1:$A,B"+R+")))"
);
  sheet.getRange(rowNum, 11).setFormula("=IF(B"+R+"=\"\",\"\",L"+R+"+M"+R+"+N"+R+"+O"+R+"+P"+R+"+Q"+R+"+R"+R+")");
  sheet.getRange(rowNum, 12).setFormula(
    "=IF(B"+R+"=\"\",\"\",SUMPRODUCT((" +
    "IF(EXACT(\"EL\","+ATT+"!$C$1:$AG),1,0)+" +
    "IF(ISNUMBER(SEARCH(\"0.5EL/\","+ATT+"!$C$1:$AG)),0.5,0)+" +
    "IF(ISNUMBER(SEARCH(\"/0.5EL\","+ATT+"!$C$1:$AG)),0.5,0)" +
    ")*("+ATT+"!$A$1:$A=B"+R+")))"
  );
  sheet.getRange(rowNum, 13).setFormula(
    "=IF(B"+R+"=\"\",\"\",SUMPRODUCT((" +
    "IF(EXACT(\"SL\","+ATT+"!$C$1:$AG),1,0)+" +
    "IF(ISNUMBER(SEARCH(\"0.5SL/\","+ATT+"!$C$1:$AG)),0.5,0)+" +
    "IF(ISNUMBER(SEARCH(\"/0.5SL\","+ATT+"!$C$1:$AG)),0.5,0)" +
    ")*("+ATT+"!$A$1:$A=B"+R+")))"
  );
  sheet.getRange(rowNum, 14).setFormula(
    "=IF(B"+R+"=\"\",\"\",SUMPRODUCT((" +
    "IF(EXACT(\"CL\","+ATT+"!$C$1:$AG),1,0)+" +
    "IF(ISNUMBER(SEARCH(\"0.5CL/\","+ATT+"!$C$1:$AG)),0.5,0)+" +
    "IF(ISNUMBER(SEARCH(\"/0.5CL\","+ATT+"!$C$1:$AG)),0.5,0)" +
    ")*("+ATT+"!$A$1:$A=B"+R+")))"
  );
  sheet.getRange(rowNum, 15).setFormula("=IF(B"+R+"=\"\",\"\",SUMPRODUCT(("+ATT+"!$C$1:$AG=\"CO\")*("+ATT+"!$A$1:$A=B"+R+")))");
  sheet.getRange(rowNum, 16).setFormula(
    "=IF(B"+R+"=\"\",\"\",SUMPRODUCT((" +
    "IF(EXACT(\"OWH\","+ATT+"!$C$1:$AG),1,0)+" +
    "IF(ISNUMBER(SEARCH(\"OWH/\","+ATT+"!$C$1:$AG)),0.5,0)+" +
    "IF(ISNUMBER(SEARCH(\"0.5CL/0.5OWH\","+ATT+"!$C$1:$AG)),0.5,0)+" +
    "IF(ISNUMBER(SEARCH(\"/OWH\","+ATT+"!$C$1:$AG)),0.5,0)" +
    ")*("+ATT+"!$A$1:$A=B"+R+")))"
  );
  sheet.getRange(rowNum, 17).setFormula("=IF(B"+R+"=\"\",\"\",SUMPRODUCT(("+ATT+"!$C$1:$AG=\"ML\")*("+ATT+"!$A$1:$A=B"+R+")))");
  sheet.getRange(rowNum, 18).setFormula("=IF(B"+R+"=\"\",\"\",SUMPRODUCT(("+ATT+"!$C$1:$AG=\"PL\")*("+ATT+"!$A$1:$A=B"+R+")))");
  sheet.getRange(rowNum, 19).setFormula(
    "=IF(B"+R+"=\"\",\"\",SUMPRODUCT((" +
    "IF(ISNUMBER(SEARCH(\"/0.5LWP\","+ATT+"!$C$1:$AG)),0.5,0)+" +
    "IF(ISNUMBER(SEARCH(\"0.5LWP/\","+ATT+"!$C$1:$AG)),0.5,0)+" +
    "IF(EXACT(\"LWP\","+ATT+"!$C$1:$AG),1,0)+" +
    "IF(EXACT(\"A\","+ATT+"!$C$1:$AG),1,0)+" +
    "IF(ISNUMBER(SEARCH(\"0.5A/\","+ATT+"!$C$1:$AG)),0.5,0)+" +
    "IF(ISNUMBER(SEARCH(\"/0.5A\","+ATT+"!$C$1:$AG)),0.5,0)" +
    ")*("+ATT+"!$A$1:$A=B"+R+")))"
  );
  sheet.getRange(rowNum, 20).setFormula("=IF(B"+R+"=\"\",\"\",D"+R+"-S"+R+")");
  sheet.getRange(rowNum, 21).setFormula("=IF(B"+R+"=\"\",\"\",SUMPRODUCT(IF(EXACT(\"-\","+ATT+"!$C$1:$AG),1,0)*("+ATT+"!$A$1:$A=B"+R+")))");
 var COMP = "'"+CONFIG.COMPENSATION_SHEET+"'";
 var empIdForOT = String(sheet.getRange(rowNum, 2).getValue() || "").trim();
 var overtimeDays = getTotalOvertimeDaysForEmp(empIdForOT);

 sheet.getRange(rowNum, 23).setFormula(
  "=IFERROR(ROUND(V"+R+"/"+totalDays+"*"+overtimeDays+",2),0)"
);
  sheet.getRange(rowNum, 25).setFormula(
  "=IF(B"+R+"=\"\",\"\",ROUND(IFERROR(" +
  "INDEX('"+CONFIG.SALARY_PF_SHEET+"'!E:E," +
  "MATCH(B"+R+",'"+CONFIG.SALARY_PF_SHEET+"'!B:B,0))" +
  "/" + totalDays + "*T"+R+",0),2))"
);
}

//******************************************************* 
function getTotalOvertimeDaysForEmp(empId) {
  if (!empId) return 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.COMPENSATION_SHEET);

  if (!sheet || sheet.getLastRow() < 2) return 0;

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var empIdx = getColIndex(headers, ["Employee ID", "Employee Id", "Employee HRM", "Emp ID", "EmpID"]);
  var otIdx = getColIndex(headers, ["Overtime Days"]);
  var approvalIdx = getColIndex(headers, [
    "ApprovalStatus",
    "Approval Status",
    "approvalStatus",
    "Approval status"
  ]);

  if (empIdx === -1 || otIdx === -1) return 0;
  var total = 0;

  for (var i = 1; i < data.length; i++) {
    var rowEmpId = String(data[i][empIdx] || "").trim();
    if (rowEmpId !== empId) continue;
    // Only approved compensation should be counted in Salary sheet
    if (approvalIdx !== -1) {
      var approvalStatus = String(data[i][approvalIdx] || "")
        .trim()
        .toLowerCase();

      if (approvalStatus !== "approved") {
        continue;
      }
    }

    var val = data[i][otIdx];
    if (val === "" || val === null || val === undefined) continue;
    var num = typeof val === "number"
      ? val
      : parseFloat(String(val).replace(/,/g, "").trim());

    if (!isNaN(num)) {
      total += num;
    }
  }

  return total;
}

// ═══════════════════════════════════════════════════════════════════════════
function handleSalaryPFEdit(e) {
  if (!e || !e.range) return;
  var range = e.range;
  var startCol = range.getColumn();
  var endCol   = range.getLastColumn();
  var startRow = range.getRow();
  var endRow   = range.getLastRow();
  if (startRow === 1 && endRow === 1) return;
  if (startRow < 2) startRow = 2;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pfSheet = ss.getSheetByName(CONFIG.SALARY_PF_SHEET);
  var salSheet = ss.getSheetByName(CONFIG.SALARY_SHEET);
  if (!pfSheet) return;
  var pfLastCol = pfSheet.getLastColumn();
  var pfHeaders = pfSheet.getRange(1, 1, 1, pfLastCol).getValues()[0];

  var entityIdx  = getColIndex(pfHeaders, ["Entities", "Entity"]);
  var empIdIdx   = getColIndex(pfHeaders, ["Employee ID", "Employee Id", "Employee HRM"]);
  var monthlyIdx = getColIndex(pfHeaders, ["Monthly Fixed"]);
  var pfIdx      = getColIndex(pfHeaders, ["PF"]);

  if (entityIdx === -1 || empIdIdx === -1 || monthlyIdx === -1 || pfIdx === -1) {
    console.log("handleSalaryPFEdit skipped: required columns missing");
    return;
  }

  var watchedCols = [
    entityIdx + 1,
    monthlyIdx + 1,
    pfIdx + 1
  ];

  var hasWatched = false;

  for (var wc = 0; wc < watchedCols.length; wc++) {
    if (watchedCols[wc] >= startCol && watchedCols[wc] <= endCol) {
      hasWatched = true;
      break;
    }
  }

  if (!hasWatched) return;
  var numRows = endRow - startRow + 1;
  var pfData = pfSheet.getRange(startRow, 1, numRows, pfLastCol).getValues();

  // Update Salary sheet immediately if it exists.
  if (salSheet && salSheet.getLastRow() >= 2) {
    var salLastRow = salSheet.getLastRow();
    var salEmpIds = salSheet.getRange(2, 2, salLastRow - 1, 1).getValues();

    var salHdrs = salSheet.getRange(1, 1, 1, salSheet.getLastColumn()).getValues()[0];
    var salEntityIdx = getColIndex(salHdrs, ["Entities", "Entity"]);
    var salMonthlyIdx = getColIndex(salHdrs, ["Monthly Fixed"]);

    for (var i = 0; i < numRows; i++) {
      var pfRowData = pfData[i];

      var empIdVal = String(pfRowData[empIdIdx] || "").trim();
      if (!empIdVal) continue;
      var entityVal = pfRowData[entityIdx];
      var monthlyVal = pfRowData[monthlyIdx];

      for (var r = 0; r < salEmpIds.length; r++) {
        if (String(salEmpIds[r][0] || "").trim() !== empIdVal) continue;

        var salRow = r + 2;

        // If Entities was edited, update Salary sheet Entities.
        if (
          (entityIdx + 1) >= startCol &&
          (entityIdx + 1) <= endCol &&
          salEntityIdx !== -1
        ) {
          if (entityVal === "") {
            salSheet.getRange(salRow, salEntityIdx + 1).clearContent();
          } else {
            salSheet.getRange(salRow, salEntityIdx + 1).setValue(entityVal);
          }
        }

        // If Monthly Fixed was edited, update Salary sheet Monthly Fixed.
        if (
          (monthlyIdx + 1) >= startCol &&
          (monthlyIdx + 1) <= endCol &&
          salMonthlyIdx !== -1
        ) {
          salSheet.getRange(salRow, salMonthlyIdx + 1).setValue(monthlyVal || 0);
        }

        // If PF was edited, reapply Salary sheet formulas.
        if ((pfIdx + 1) >= startCol && (pfIdx + 1) <= endCol) {
          applySalaryFormulas(salSheet, salRow);
        }

        break;
      }
    }
  }

  // Save after one-by-one edit, paste, or drag-fill.
  savePFData();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✔ Salary & PF Details saved for this attendance month",
    "Auto-saved",
    2
  );
}

//*******************************************
function handleSalaryRemarksEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SALARY_SHEET) return;
  var startCol = e.range.getColumn();
  var endCol   = e.range.getLastColumn();
  var startRow = e.range.getRow();
  var endRow   = e.range.getLastRow();

  if (startRow === 1 && endRow === 1) return;
  if (startRow < 2) startRow = 2;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var remarksIdx = getColIndex(headers, ["Remarks"]);
  if (remarksIdx === -1) return;

  var remarksColNum = remarksIdx + 1;

  if (remarksColNum < startCol || remarksColNum > endCol) return;

  // Save after one-by-one edit, paste, or drag-fill.
  saveSalaryRemarksData();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✔ Salary Remarks saved for this attendance month",
    "Auto-saved",
    2
  );
}

//*************************************************
function getAttendanceMonthYearKey() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attSheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  var props = PropertiesService.getScriptProperties();

  var today = new Date();

  var selectedMonth = props.getProperty("SELECTED_MONTH");
  var selectedYear  = props.getProperty("SELECTED_YEAR");

  var calcMonth = selectedMonth !== null && selectedMonth !== ""
    ? parseInt(selectedMonth, 10)
    : today.getMonth();

  var calcYear = selectedYear !== null && selectedYear !== ""
    ? parseInt(selectedYear, 10)
    : today.getFullYear();

  // Detect month/year from Attendance sheet headers if possible.
  // Attendance date columns start from column C.
  if (attSheet && attSheet.getLastColumn() >= 3) {
    var headers = attSheet
      .getRange(1, 3, 1, attSheet.getLastColumn() - 2)
      .getValues()[0];

    for (var i = 0; i < headers.length; i++) {
      var h = headers[i];

      // If header is an actual Date object
      if (h instanceof Date && !isNaN(h.getTime())) {
        calcMonth = h.getMonth();
        calcYear = h.getFullYear();
        break;
      }

      // If header is like "1 - Apr"
      var s = String(h || "").trim();
      var m = s.match(/^\d+\s*-\s*([A-Za-z]{3})$/);

      if (m) {
        var monthNames = [
          "Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ];

        var monthIdx = monthNames.indexOf(m[1]);

        if (monthIdx !== -1) {
          calcMonth = monthIdx;

          // Year is not available in "1 - Apr", so use selected year.
          if (selectedYear !== null && selectedYear !== "") {
            calcYear = parseInt(selectedYear, 10);
          }

          break;
        }
      }
    }
  }

  return {
    month: calcMonth,
    year: calcYear,
    key: calcYear + "_" + calcMonth
  };
}