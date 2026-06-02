// ═══════════════════ATTENDANCE — fetch from Zoho,write to sheet,save snapshot══════════════════
function attendance(ACCESS_TOKEN, targetSheetName, selMonth, selYear) {
  var authHeader = { "Authorization": "Zoho-oauthtoken " + ACCESS_TOKEN };
  var today = new Date();
  var month = (selMonth !== null && selMonth !== undefined) ? parseInt(selMonth) : today.getMonth();
  var year  = (selYear  !== null && selYear  !== undefined) ? parseInt(selYear)  : today.getFullYear();
  var sdate = new Date(year, month, 1);
  var edate = new Date(year, month + 1, 0);
  function fmtApi(d)    { return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"); }
  function parseZohoExitDate_(raw) {
  if (!raw) return null;

  var s = String(raw).trim();
  // yyyy-MM-dd
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var p1 = s.split("-");
    return new Date(Number(p1[0]), Number(p1[1]) - 1, Number(p1[2]));
  }

  var m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
  if (m) {
    var months = {
      Jan: 0, Feb: 1, Mar: 2, Apr: 3,
      May: 4, Jun: 5, Jul: 6, Aug: 7,
      Sep: 8, Oct: 9, Nov: 10, Dec: 11
    };

    return new Date(Number(m[3]), months[m[2]], Number(m[1]));
  }

  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  return null;
}

function dateKeyToDate_(key) {
  var p = String(key).split("-");
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}
  function fmtHeader(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), "d - MMM"); }
  function fullName(emp){ return ((emp["first name"]||"")+" "+(emp["last name"]||"")).trim(); }

 function normalizeStatus(dayObj) {
    if (!dayObj) return "-";
    var s = (dayObj.Status || "").toString().trim();
    var parsedCombo = parseZohoCombinedStatus(s);
    if (parsedCombo) {
      return parsedCombo;
    }
    var hasPunch =
    dayObj.FirstIn && dayObj.FirstIn !== "-" &&
    dayObj.LastOut && dayObj.LastOut !== "-";
    var statusLower = s.toLowerCase();
    if (
    statusLower.indexOf("client visit") !== -1 ||
    statusLower.indexOf("clientvisit") !== -1 ||
    statusLower.indexOf("official event") !== -1 ||
    statusLower.indexOf("work-related absence") !== -1 ||
    statusLower.indexOf("work related absence") !== -1 ||
    statusLower.indexOf("work from home") !== -1 ||
    statusLower === "wfh" ||
    statusLower.indexOf("on duty") !== -1 ||
    statusLower.indexOf("onduty") !== -1
  ) {
    return "OD";
  }

    if ((statusLower === "weekend" || statusLower === "w") && hasPunch) {
      return "W/P";
    }
    if (
    (statusLower === "holiday" ||
    statusLower === "holidays" ||
    statusLower === "h" ||
    statusLower.indexOf("holiday") !== -1) &&
    hasPunch
  ) {
    return "H/P";
  }

  if (
    statusLower === "holiday" ||
    statusLower === "holidays" ||
    statusLower === "h" ||
    statusLower.indexOf("holiday") !== -1
  ) {
    return "H";
  }
    if (/^(0\.5)?[A-Za-z]+\/(0\.5)?[A-Za-z]+$/.test(s)) {
   return s.toUpperCase();
  }
    if(s===""&&(!dayObj.FirstIn||dayObj.FirstIn==="-")&&(!dayObj.LastOut||dayObj.LastOut==="-")) return "-";
    if (/^0\.5[A-Z]+\/0\.5[A-Z]+$/i.test(s)) return s.toUpperCase();
    var direct = ["P","H","W","A","ABSENT","CL","SL","EL","ML","PL","SBL","LWP","OWH","WFH","CO","OD"];
    if (s.toUpperCase() === "ABSENT") return "A";
    if (direct.indexOf(s.toUpperCase()) !== -1) return s.toUpperCase();

    // ── Handle half-day combos from Zoho ──────────────────────────────────
    // Format: "Casual Leave(First Half), 0.5 day Absent" OR "Casual Leave(First Half), 0.5 day Present"
    if(s.indexOf("Half")!==-1 && s.indexOf(",")!==-1){
      var parts   = s.split(",");
      var first   = parts[0].trim();  // e.g. "Casual Leave(First Half)"
      var second  = parts[1].trim();  // e.g. "0.5 day Absent" or "0.5 day Present"
      // Detect leave type from first part
      var leaveCode = "CL";
      if(first.indexOf("Casual Leave")!==-1)       leaveCode="CL";
      else if(first.indexOf("Sick Leave")!==-1)    leaveCode="SL";
      else if(first.indexOf("Earned Leave")!==-1)  leaveCode="EL";
      else if(first.indexOf("Maternity")!==-1)     leaveCode="ML";
      else if(first.indexOf("Paternity")!==-1)     leaveCode="PL";
      else if(first.indexOf("Optional WFH")!==-1)  leaveCode="OWH";
      else if(first.indexOf("Compensatory")!==-1)  leaveCode="CO";
      else if(first.indexOf("On Duty")!==-1)       leaveCode="OD";
      else if(first.indexOf("Leave Without Pay")!==-1) leaveCode="LWP";
      else if(first.indexOf("Sabbatical")!==-1)    leaveCode="SBL";

      // Detect other half status
      var otherCode = "A"; 
      if(second.indexOf("Present")!==-1)  otherCode="P";
      else if(second.indexOf("Absent")!==-1) otherCode="A";

      // First Half = leave in morning, other in afternoon
      if(first.indexOf("First Half")!==-1){
        return "0.5"+leaveCode+"/0.5"+otherCode;
      }
      // Second Half = other in morning, leave in afternoon
      if(first.indexOf("Second Half")!==-1){
        return "0.5"+otherCode+"/0.5"+leaveCode;
      }
    }

    var map = {
      "Present":"P","Holidays":"H","Holiday":"H","Weekend":"W","Absent":"A",
      "Casual Leave":"CL","CasualLeave":"CL","Sick Leave":"SL","SickLeave":"SL",
      "Earned Leave":"EL","EarnedLeave":"EL","Maternity Leave":"ML","MaternityLeave":"ML",
      "Paternity Leave":"PL","PaternityLeave":"PL","Sabbatical Leave":"SBL",
      "Leave Without Pay":"LWP","LeaveWithoutPay":"LWP",
      "Optional WFH":"OWH","Compensatory Off":"CO","CompOff":"CO",
      "On Duty":"OD","OnDuty":"OD","Paid Leave":"Paid Leave","Unpaid Leave":"Unpaid Leave"
    };
    if (map[s] !== undefined) return map[s];
    if (s===""&&dayObj.FirstIn&&dayObj.FirstIn!=="-"&&dayObj.LastOut&&dayObj.LastOut!=="-") return "P";
    return "-";
  }

  var allResults = [];
  var startIndex = 0;
  var batchSize  = 15;
  while(true){
    var params = "sdate="+encodeURIComponent(fmtApi(sdate))+
                 "&edate="+encodeURIComponent(fmtApi(edate))+
                 "&dateFormat=yyyy-MM-dd"+
                 "&startIndex="+startIndex+
                 "&limit="+batchSize;
    var url  = "https://people.zoho.in/people/api/attendance/getUserReport?"+params;
    var resp = UrlFetchApp.fetch(url, {
    method: "get",
    headers: authHeader,
    muteHttpExceptions: true
  });

    // console.log("ATTENDANCE API URL:");
    // console.log(url);

    // console.log("ATTENDANCE API STATUS:");
    // console.log(resp.getResponseCode());

    // console.log("ATTENDANCE API RAW RESPONSE:");
    // console.log(resp.getContentText());

    var json = JSON.parse(resp.getContentText());

    // console.log("ATTENDANCE API PARSED JSON:");
    // console.log(JSON.stringify(json, null, 2));
      if(!json.result || !Array.isArray(json.result) || json.result.length===0) break;
      allResults = allResults.concat(json.result);
      // If returned less than batchSize, we have fetched all
      if(json.result.length < batchSize) break;
      startIndex += batchSize;
      Utilities.sleep(1000); // 1 second pause to avoid bandwidth quota
    }
      // ── Deduplicate allResults by employee id ──────────────────────────
    var seen = {};
    var dedupedResults = [];
    allResults.forEach(function(emp) {
      var id = String((emp.employeeDetails || {}).id || "").trim();
      if (!id) return;
      if (seen[id]) {
        // Already have this employee — merge attendanceDetails
        // (keep whichever has more data)
        var existing = seen[id];
        var existingDays = Object.keys(existing.attendanceDetails || {}).length;
        var newDays = Object.keys(emp.attendanceDetails || {}).length;
        if (newDays > existingDays) {
          seen[id].attendanceDetails = emp.attendanceDetails;
        }
        return;
      }
      seen[id] = emp;
      dedupedResults.push(emp);
    });
    allResults = dedupedResults;
    // console.log("Total employees after dedup: " + allResults.length);
  var odMap = fetchODRequestMap(authHeader, sdate, edate);
  var leaveMap = fetchLeaveRequestMap(authHeader, sdate, edate);
  var joinDateMap = fetchEmployeeJoinDateMap(authHeader);
  // ── Append resigned employees attendance ──────────────────────────────
  var activeIds = {};
  allResults.forEach(function(emp) {
    var id = String((emp.employeeDetails || {}).id || "").trim();
    if (id) activeIds[id] = true;
  });

  // console.log("========== ACTIVE EMPLOYEES BEFORE RESIGNED APPEND ==========");
  // console.log("Active employee count: " + Object.keys(activeIds).length);
  // console.log("Active employee IDs: " + JSON.stringify(Object.keys(activeIds)));
  // console.log("=============================================================");

  var resignedResults = fetchResignedAttendance(authHeader, fmtApi(sdate), fmtApi(edate));

  // console.log("========== RESIGNED RESULTS RETURNED TO attendance() ==========");
  // console.log("Resigned results count: " + resignedResults.length);
  // console.log(JSON.stringify(resignedResults, null, 2));
  // console.log("==============================================================");

  resignedResults.forEach(function(emp) {
    var id = String((emp.employeeDetails || {}).id || "").trim();
    if (!id) return;

    if (!activeIds[id]) {
      allResults.push(emp);
      activeIds[id] = true;
      // console.log("APPENDED RESIGNED EMPLOYEE TO allResults: " + id);
    } else {
      // console.log("RESIGNED EMPLOYEE ALREADY EXISTS IN ACTIVE RESULT: " + id);

      for (var i = 0; i < allResults.length; i++) {
        var existingId = String((allResults[i].employeeDetails || {}).id || "").trim();
        if (existingId === id) {
          allResults[i].employeeDetails.dateOfExit = emp.employeeDetails.dateOfExit;
          // console.log("UPDATED DATE OF EXIT FOR EXISTING EMPLOYEE: " + id + " => " + emp.employeeDetails.dateOfExit);
          break;
        }
      }
    }
  });

  if(allResults.length === 0) {
  var ssNoData = SpreadsheetApp.getActiveSpreadsheet();
  var noDataSheetName = targetSheetName || CONFIG.ATTENDANCE_SHEET;
  var noDataSheet = ssNoData.getSheetByName(noDataSheetName);
  if (!noDataSheet) {
    noDataSheet = ssNoData.insertSheet(noDataSheetName);
  } else {
    noDataSheet.clearContents();
    noDataSheet.clearFormats();
  }

  noDataSheet.getRange(1, 1).setValue("No attendance data found for selected month/year")
    .setFontWeight("bold");

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "⚠️ No attendance data returned from Zoho for selected month. Attendance sheet cleared.",
    "Warning",
    6
  );
  return {status: 0, rowsWritten: 0, sheetName: noDataSheetName};
}
  var json = {result: allResults};
  var registry=loadRegistry(), fetchedIds={};
  json.result.forEach(function(emp) {
    var empId=String((emp.employeeDetails||{}).id||"").trim();
    if (!empId) return;
    fetchedIds[empId]=emp;
    if (!registry[empId]) {
      var ad=emp.attendanceDetails||{}, joinDate="", d=new Date(sdate);
      while (d<=edate) {
        var dk=fmtApi(d), day=ad[dk];
        if (day&&normalizeStatus(day)!==""&&normalizeStatus(day)!=="-"){joinDate=dk;break;}
        d.setDate(d.getDate()+1);
      }
      registry[empId]={name:fullName(emp.employeeDetails||{}),order:getNextOrder(),joinDate:joinDate};
    } else {
      registry[empId].name=fullName(emp.employeeDetails||{});
    }
  });

  var deletedIds=[];
  for (var rid in registry){if(!fetchedIds.hasOwnProperty(rid))deletedIds.push(rid);}
  // console.log("=== fetchedIds from Zoho: "+JSON.stringify(Object.keys(fetchedIds))+" ===");
  // console.log("=== deletedIds: "+JSON.stringify(deletedIds)+" ===");
  deletedIds.forEach(function(id){delete registry[id];});
  saveRegistry(registry);
  if (deletedIds.length>0) {
    var ss0=SpreadsheetApp.getActiveSpreadsheet();
    deletedIds.forEach(function(empId){removeEmployeeFromAllSheets(ss0,empId);});
    SpreadsheetApp.flush();
  }

  var sortedEmps=allResults.slice().sort(function(a,b){
    return parseInt(String((a.employeeDetails||{}).id||"").replace(/\D/g,""),10) -
           parseInt(String((b.employeeDetails||{}).id||"").replace(/\D/g,""),10);
  });

  var apiDates=[], headerDates=[], dd=new Date(sdate);
  while (dd<=edate){apiDates.push(fmtApi(dd));headerDates.push(fmtHeader(dd));dd.setDate(dd.getDate()+1);}
  var finalData=[["Employee Id","Employee Name"].concat(headerDates)];
  var writtenIds = {};   // ← ADD THIS
  sortedEmps.forEach(function(emp){
  var ed = emp.employeeDetails || {}, ad = emp.attendanceDetails || {};
  var empId = String(ed.id || "").trim();
  if (!empId) return;
  if (writtenIds[empId]) return;
  var actualJoinDate = joinDateMap[empId] || null;
  // Employee joined after selected month, so skip completely
  if (actualJoinDate && actualJoinDate > edate) {
    // console.log("SKIPPED NOT JOINED IN THIS MONTH: " + empId);
    return;
  }
  writtenIds[empId] = true;
  var row = [empId, fullName(ed)];
    var firstRealIdx=-1;
    apiDates.forEach(function(dk,idx){
      if(firstRealIdx!==-1) return;
      var day=ad[dk];
      if(!day) return;
      var s=(day.Status||"").toString().trim().toLowerCase();
      // Real working day = not weekend/holiday
      if(s!==""&&s!=="-"&&s!=="weekend"&&s!=="holiday"&&s!=="holidays"){
        firstRealIdx=idx;
      }
    });

    if(firstRealIdx===-1){
  // Skip employee completely — all days are '-', W, or H with no real activity
  return;
   }
    // Step 2: Go backward and include only W/H that have "OverTime" key
    var joinIdx=firstRealIdx;
    for(var b=firstRealIdx-1;b>=0;b--){
      var day=ad[apiDates[b]];
      if(!day) break;
      var s=(day.Status||"").toString().trim().toLowerCase();
      if((s==="weekend"||s==="holiday"||s==="holidays") && day.hasOwnProperty("OverTime")){
        joinIdx=b;
      } else {
        break;
      }
    }
    // Step 3: Build row
    var exitDate = parseZohoExitDate_(ed.dateOfExit);
    apiDates.forEach(function(dk, idx) {
    var currentDate = dateKeyToDate_(dk);
    // Before employee joining date, always show "-"
    if (actualJoinDate && currentDate < actualJoinDate) {
      row.push("-");
      return;
    }

    // Before first real attendance, also show "-"
    if (idx < joinIdx) {
      row.push("-");
      return;
    }
      // If employee resigned on 27-Apr, then 28/29/30-Apr must be "-"
      if (exitDate) {
        var currentDate = dateKeyToDate_(dk);
        if (currentDate > exitDate) {
          row.push("-");
          return;
        }
      }

     var baseStatus = normalizeStatus(ad[dk]);
      var odInfo = odMap[empId + "|" + dk];
      var leaveInfo = leaveMap[empId + "|" + dk];

      // If OD / Leave API has data, do not stop at "-"
      if (baseStatus === "-" && !odInfo && !leaveInfo) {
        row.push("-");
        return;
      }

     var finalStatus = mergeLeaveODWithAttendanceStatus(baseStatus, leaveInfo, odInfo);
     if (String(finalStatus).toUpperCase() === "ABSENT") {
       finalStatus = "A";
      }

     row.push(finalStatus);
    });
    finalData.push(row);
  });

  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheetName=targetSheetName||"Attendance";
  var sheet=ss.getSheetByName(sheetName);
  if (!sheet){sheet=ss.insertSheet(sheetName);}else{sheet.clearContents();sheet.clearFormats();}
  var CHUNK=200, tRows=finalData.length, tCols=finalData[0].length;
  for (var si=0;si<tRows;si+=CHUNK){
    var chunk=finalData.slice(si,Math.min(si+CHUNK,tRows));
    sheet.getRange(si+1,1,chunk.length,tCols).setValues(chunk);
    SpreadsheetApp.flush();
  }

  var lastRow=sheet.getLastRow(), lastCol=sheet.getLastColumn();
  sheet.getRange(1,1,1,lastCol).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground("#d9e2f3");
  sheet.getRange(1,1,1,2).setHorizontalAlignment("left");
  if (lastRow>1){
    sheet.getRange(2,1,lastRow-1,2).setHorizontalAlignment("left").setVerticalAlignment("middle");
    if (lastCol>2) sheet.getRange(2,3,lastRow-1,lastCol-2).setHorizontalAlignment("center").setVerticalAlignment("middle");
  }
  sheet.setFrozenRows(1);sheet.setFrozenColumns(2);sheet.setRowHeight(1,32);
  sheet.setColumnWidth(1,110);sheet.setColumnWidth(2,180);
  for (var ci=3;ci<=lastCol;ci++) sheet.setColumnWidth(ci,72);
  sheet.getRange(1,1,lastRow,lastCol).setBorder(true,true,true,true,true,true);

  if (lastRow>1&&lastCol>2){
    var drc=lastRow-1,dcc=lastCol-2;
    for (var rs=0;rs<drc;rs+=CHUNK){
      var re=Math.min(rs+CHUNK,drc),rc=re-rs;
      var grid=sheet.getRange(rs+2,3,rc,dcc),vals=grid.getValues(),bgs=[],fcs=[];
      for (var ii=0;ii<vals.length;ii++){
        var bgRow=[],fcRow=[];
        for (var jj=0;jj<vals[ii].length;jj++){
          var v=vals[ii][jj],bg="#ffffff";
          if(v==="P")bg="#e2f0d9";else if(v==="W")bg="#fff2cc";else if(v==="H")bg="#ddebf7";
          else if(v==="A")bg="#f4cccc";else if(v==="OD")bg="#e4d5ff";
          else if(["CL","SL","EL","PL","UL"].indexOf(v)!==-1)bg="#fce5cd";
          else if(v==="OWH")bg="#d9eaf7";
          else if(/^0\.5[A-Z]+\/0\.5[A-Z]+$/i.test(v))bg="#d9ead3";
          else if(v==="CO")bg="#fce5cd";
          else if(["OT","WFH","LOP"].indexOf(v)!==-1)bg="#ead1dc";
          bgRow.push(bg);fcRow.push("#000000");
        }
        bgs.push(bgRow);fcs.push(fcRow);
      }
      grid.setBackgrounds(bgs);grid.setFontColors(fcs);SpreadsheetApp.flush();
    }
  }

  // Save snapshot after writing attendance so onChange can diff against it
  var snapshot={};
  finalData.slice(1).forEach(function(row){
    var id=String(row[0]||"").trim();
    if (id) snapshot[id]=String(row[1]||"").trim();
  });
  PropertiesService.getScriptProperties().setProperty(CONFIG.PROP_ATT_SNAPSHOT, JSON.stringify(snapshot));
  return {status:resp.getResponseCode(), rowsWritten:tRows-1, sheetName:sheetName};
}

//***************************************
function mergeLeaveODWithAttendanceStatus(baseStatus, leaveInfo, odInfo) {
  baseStatus = String(baseStatus || "").trim().toUpperCase();

  // If Zoho attendance API already gives exact combo, keep it.
  if (/^(0\.5)?[A-Z]+\/(0\.5)?[A-Z]+$/.test(baseStatus)) {
    return baseStatus;
  }

  if (!baseStatus || baseStatus === "-") {
    baseStatus = "A";
  }

  if (leaveInfo && odInfo) {
    var leaveCode = String(leaveInfo.code || "").trim().toUpperCase();

    var isHalf =
      leaveInfo.half === "FIRST_HALF" ||
      leaveInfo.half === "SECOND_HALF" ||
      leaveInfo.half === "HALF_UNKNOWN" ||
      odInfo.half === "FIRST_HALF" ||
      odInfo.half === "SECOND_HALF" ||
      odInfo.half === "HALF_UNKNOWN";

    return combineAttendanceParts("OD", leaveCode, isHalf);
  }

  // Leave only
  if (leaveInfo) {
    var lc = String(leaveInfo.code || "").trim().toUpperCase();

    if (leaveInfo.half === "FIRST_HALF") {
      return combineAttendanceParts(lc, baseStatus, true);
    }

    if (leaveInfo.half === "SECOND_HALF") {
      return combineAttendanceParts(baseStatus, lc, true);
    }

    if (leaveInfo.half === "HALF_UNKNOWN") {
      return combineAttendanceParts(lc, baseStatus, true);
    }

    return lc;
  }

  // OD only
  if (odInfo) {
    return mergeODWithAttendanceStatus(baseStatus, odInfo);
  }

  return baseStatus;
}

//*************************************************
function mergeODWithAttendanceStatus(baseStatus, odInfo) {
  if (!odInfo) return baseStatus;
  baseStatus = String(baseStatus || "").trim().toUpperCase();
  // If baseStatus already has a combo from Zoho, keep it.
  // Example: OD/CL, W/P, 0.5CL/0.5A
  if (/^(0\.5)?[A-Z]+\/(0\.5)?[A-Z]+$/.test(baseStatus)) {
    return baseStatus;
  }

  var odTypeLower = String(odInfo.odType || "").trim().toLowerCase();

// Zoho Work From Home should be shown as OD, not "-"
  if (
    odTypeLower === "work from home" ||
    odTypeLower === "work from home " ||
    odTypeLower.indexOf("work from home") !== -1
  ) {
    return "OD";
  }

  if (baseStatus === "" || baseStatus === "-") {
    baseStatus = "A";
  }

  function half(code) {
    code = String(code || "").trim().toUpperCase();
    if (!code || code === "-") return "";
    if (code.indexOf("0.5") === 0) return code;
    return "0.5" + code;
  }

  function combine(left, right, isHalf) {
    left = String(left || "").trim().toUpperCase();
    right = String(right || "").trim().toUpperCase();

    if (!left || left === "-") return right || "-";
    if (!right || right === "-") return left;
    if (left === right) return left;

    if (isHalf) {
      return half(left) + "/" + half(right);
    }

    return left + "/" + right;
  }

  if (odInfo.half === "FIRST_HALF") {
    return combine("OD", baseStatus, true);
  }

  if (odInfo.half === "SECOND_HALF") {
    return combine(baseStatus, "OD", true);
  }

  if (odInfo.half === "HALF_UNKNOWN") {
    if (baseStatus === "A" && String(odInfo.dayTaken) === "0") {
      return "OD/P";
    }
    return combine("OD", baseStatus, true);
  }

  if (odInfo.half === "FULL_DAY") {
    if (baseStatus === "A" && String(odInfo.dayTaken) === "0") {
      return "OD/P";
    }

    return combine("OD", baseStatus, false);
  }

  return baseStatus;
}

function parseZohoCombinedStatus(rawStatus) {
  var s = String(rawStatus || "").trim();
  if (!s) return "";

  // Already short combination like W/P, OD/CL, 0.5CL/0.5A
  if (/^(0\.5)?[A-Za-z]+\/(0\.5)?[A-Za-z]+$/.test(s)) {
    return s.toUpperCase();
  }

  // Handle Zoho comma statuses:
  // "Weekend, Present" → W/P
  // "On Duty, Casual Leave" → OD/CL
  // "On Duty(Second Half), Casual Leave(First Half)" → 0.5OD/0.5CL
  // "Casual Leave(First Half), 0.5 day Absent" → 0.5CL/0.5A
  if (s.indexOf(",") !== -1) {
    var parts = s.split(",");
    var out = [];

    for (var i = 0; i < parts.length; i++) {
      var part = String(parts[i] || "").trim();
      var code = statusTextToCode(part);
      if (!code) continue;
      var lower = part.toLowerCase();
      var isHalf =
        lower.indexOf("half") !== -1 ||
        lower.indexOf("0.5") !== -1 ||
        lower.indexOf(".5") !== -1;

      if (isHalf && code.indexOf("0.5") !== 0) {
        code = "0.5" + code;
      }

      out.push(code);
    }

    if (out.length > 0) {
      return out.join("/");
    }
  }

  return "";
}

//**************************************
function statusTextToCode(text) {
  var s = String(text || "").trim().toLowerCase();

  if (!s) return "";

  if (s.indexOf("present") !== -1) return "P";
  if (s.indexOf("weekend") !== -1) return "W";
  if (s.indexOf("holiday") !== -1 || s.indexOf("holidays") !== -1) return "H";
  if (s.indexOf("absent") !== -1) return "A";

  if (s.indexOf("on duty") !== -1 || s.indexOf("onduty") !== -1) return "OD";

  if (s.indexOf("casual leave") !== -1 || s.indexOf("casualleave") !== -1) return "CL";
  if (s.indexOf("sick leave") !== -1 || s.indexOf("sickleave") !== -1) return "SL";
  if (s.indexOf("earned leave") !== -1 || s.indexOf("earnedleave") !== -1) return "EL";
  if (s.indexOf("maternity") !== -1) return "ML";
  if (s.indexOf("paternity") !== -1) return "PL";
  if (s.indexOf("sabbatical") !== -1) return "SBL";

  if (
    s.indexOf("leave without pay") !== -1 ||
    s.indexOf("leavewithoutpay") !== -1 ||
    s.indexOf("lwp") !== -1
  ) {
    return "LWP";
  }

  if (s.indexOf("optional wfh") !== -1 || s.indexOf("optionalwfh") !== -1) return "OWH";
  if (s.indexOf("compensatory") !== -1 || s.indexOf("comp off") !== -1 || s.indexOf("compoff") !== -1) return "CO";

  return "";
}

function combineAttendanceParts(left, right, isHalfDay) {
  left = String(left || "").trim().toUpperCase();
  right = String(right || "").trim().toUpperCase();

  if (!left || left === "-") return right || "-";
  if (!right || right === "-") return left;
  if (left === right) return left;

  function half(code) {
    code = String(code || "").trim().toUpperCase();
    if (code.indexOf("0.5") === 0) return code;
    return "0.5" + code;
  }

  if (isHalfDay) {
    return half(left) + "/" + half(right);
  }

  return left + "/" + right;
}
//*****************************************
function leaveTypeToCode(leaveType) {
  var s = String(leaveType || "").toLowerCase().replace(/\s+/g, "");
  if (s.indexOf("absent") !== -1) return "A";
  if (s.indexOf("casual") !== -1) return "CL";
  if (s.indexOf("sick") !== -1) return "SL";
  if (s.indexOf("earned") !== -1) return "EL";
  if (s.indexOf("maternity") !== -1) return "ML";
  if (s.indexOf("paternity") !== -1) return "PL";
  if (s.indexOf("sabbatical") !== -1) return "SBL";
  if (s.indexOf("withoutpay") !== -1 || s.indexOf("lwp") !== -1) return "LWP";
  if (s.indexOf("optionalwfh") !== -1) return "OWH";
  if (s.indexOf("compensatory") !== -1) return "CO";

  return String(leaveType || "").trim().toUpperCase();
}
//*****************************************
function toDateKeyFromZohoDate(raw) {
  if (!raw) return "";
  var s = String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  var m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
  if (!m) return "";

  var months = {
    Jan: "01", Feb: "02", Mar: "03", Apr: "04",
    May: "05", Jun: "06", Jul: "07", Aug: "08",
    Sep: "09", Oct: "10", Nov: "11", Dec: "12"
  };

  return m[3] + "-" + months[m[2]] + "-" + ("0" + m[1]).slice(-2);
}


















