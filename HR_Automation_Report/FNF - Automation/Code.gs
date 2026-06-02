// ════════════════════════════════════GET ACCESS TOKEN══════════════════════════════════════════════
function getAccessToken() {
  var props = PropertiesService.getScriptProperties();
  var response = UrlFetchApp.fetch("https://accounts.zoho.in/oauth/v2/token", {
    method: "post",
    payload: {
      grant_type: "refresh_token",
      client_id: props.getProperty("CLIENT_ID"),
      client_secret: props.getProperty("CLIENT_SECRET"),
      refresh_token: props.getProperty("REFRESH_TOKEN")
    },
    muteHttpExceptions: true
  });

  var text = response.getContentText();
  var data = JSON.parse(text);

  if (!data.access_token) {
    throw new Error("Token error: " + text);
  }

  return data.access_token;
}

function isDuplicateEmployeeIdTab_(ss, currentSheet, employeeId) {
  var sheets = ss.getSheets();
  var duplicateTabs = [];

  sheets.forEach(function(sh) {
    if (sh.getSheetId() === currentSheet.getSheetId()) return;

    var name = sh.getName();
    var match = name.match(/^(.*?)\s*[-–—]\s*(HRM\d+)$/i);

    if (!match) return;

    var otherEmployeeId = match[2].trim().toUpperCase();

    if (otherEmployeeId === employeeId) {
      duplicateTabs.push(name);
    }
  });

  return duplicateTabs;
}
// ════════════════════════════════════LEAVE ENCASHMENT══════════════════════════════════════════════
function fetchLeaveEncashmentForCurrentTab(targetSheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = targetSheet || ss.getActiveSheet();
  var tabName = sheet.getName();

  var match = tabName.match(/^(.*?)\s*[-–—]\s*(HRM\d+)$/i);

  if (!match) {
    console.log("Skipped non-employee tab: " + tabName);
    return;
  }

  var employeeNameFromTab = match[1].trim();
  var employeeIdFromTab = match[2].trim().toUpperCase();
  var duplicateTabs = isDuplicateEmployeeIdTab_(ss, sheet, employeeIdFromTab);
  if (duplicateTabs.length > 0) {
    sheet.getRange("A1")
      .setValue(
        "Duplicate Employee ID not allowed: " +
        employeeIdFromTab +
        ". Already used in tab: " +
        duplicateTabs.join(", ")
      )
      .setFontWeight("bold")
      .setBackground("#f4cccc");

    ss.toast(
      "Duplicate Employee ID found: " + employeeIdFromTab,
      "Duplicate Tab",
      6
    );

    return;
  }

  var oldManualSalaryBlock = sheet.getRange("B7:B10").getValues();

  /*******************************************************
   * Fetch Zoho data
   *******************************************************/
  var accessToken = getAccessToken();
  var url = "https://people.zoho.in/people/api/v2/leavetracker/reports/bookedAndBalance";

  function fetchBookedAndBalanceByPeriod(fromDate, toDate) {
    var payload = {
      from: fromDate,
      to: toDate,
      unit: "Day",
      employeeStatus: JSON.stringify(["EX_EMPLOYEES"]),
      limit: 200,
      startIndex: 0
    };

    var query = Object.keys(payload)
      .map(function(key) {
        return encodeURIComponent(key) + "=" + encodeURIComponent(payload[key]);
      })
      .join("&");

    var response = UrlFetchApp.fetch(url + "?" + query, {
      method: "get",
      headers: {
        Authorization: "Zoho-oauthtoken " + accessToken
      },
      muteHttpExceptions: true
    });

    var text = response.getContentText();
    var code = response.getResponseCode();

    if (code !== 200) {
      throw new Error("Zoho Leave API Error: " + code + " | " + text);
    }

    return JSON.parse(text);
  }

  // Previous year data = Previous Y Left
  var jsonPrevYear = fetchBookedAndBalanceByPeriod(
    "01-Jan-2025",
    "31-Dec-2025"
  );

  // Current year data = Actual Used Leaves - CY
  var jsonCurrentYear = fetchBookedAndBalanceByPeriod(
    "01-Jan-2026",
    "31-Dec-2026"
  );

  var reportPrevYear = jsonPrevYear.report || {};
  var reportCurrentYear = jsonCurrentYear.report || {};

  var leaveTypes =
    jsonCurrentYear.leavetypes ||
    jsonPrevYear.leavetypes ||
    {};


  /*******************************************************
   * Employee DOJ and Date of Exit
   *******************************************************/
  function parseZohoEmpDate(raw) {
    if (!raw) return null;

    var s = String(raw).trim();

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


  function fetchEmployeeDateMap() {
    var map = {};
    var startIndex = 0;
    var limit = 200;

    while (true) {
      var empUrl =
        "https://people.zoho.in/people/api/forms/employee/getRecords" +
        "?isActive=false" +
        "&startIndex=" + startIndex +
        "&limit=" + limit;

      var empResp = UrlFetchApp.fetch(empUrl, {
        method: "get",
        headers: {
          Authorization: "Zoho-oauthtoken " + accessToken
        },
        muteHttpExceptions: true
      });

      var text = empResp.getContentText();
      var code = empResp.getResponseCode();

      if (code !== 200) {
        throw new Error("Zoho Employee API Error: " + code + " | " + text);
      }

      var data = JSON.parse(text);
      var batch = [];

      if (Array.isArray(data)) {
        batch = data;
      } else if (data && data.response && data.response.result) {
        batch = data.response.result;
      }

      if (!batch || batch.length === 0) break;

      batch.forEach(function(empRecord) {
        var recordId = Object.keys(empRecord)[0];
        var empData = empRecord[recordId];

        if (Array.isArray(empData)) {
          empData = empData[0] || {};
        }

        var empId = String(
          empData["EmployeeID"] ||
          empData["EmployeeId"] ||
          empData["Employee ID"] ||
          ""
        ).trim().toUpperCase();

        if (!empId) return;

        var dojRaw =
          empData["Dateofjoining"] ||
          empData["DateOfJoining"] ||
          empData["Date of Joining"] ||
          empData["Joining Date"] ||
          "";

        var exitRaw =
          empData["Dateofexit"] ||
          empData["DateOfExit"] ||
          empData["Date of Exit"] ||
          empData["Exit Date"] ||
          "";

        map[empId] = {
          doj: parseZohoEmpDate(dojRaw),
          exitDate: parseZohoEmpDate(exitRaw)
        };
      });

      if (batch.length < limit) break;

      startIndex += limit;
      Utilities.sleep(300);
    }

    return map;
  }

  var employeeDateMap = fetchEmployeeDateMap();
  var currentYear = 2026;


  /*******************************************************
   * Helpers
   *******************************************************/
  function findEmployeeInReport(reportObj, employeeNameFromTab, employeeIdFromTab) {
    var matched = null;

    Object.keys(reportObj).forEach(function(recordId) {
      var empData = reportObj[recordId];

      if (!empData.employee) return;

      var apiEmpId = String(empData.employee.id || "").trim();
      var apiEmpName = String(empData.employee.name || "").trim();

      if (
        apiEmpId === employeeIdFromTab ||
        apiEmpName.toLowerCase() === employeeNameFromTab.toLowerCase()
      ) {
        matched = empData;
      }
    });

    return matched;
  }


  function findLeaveTypeIdByName(leaveTypesObj, targetName) {
    targetName = String(targetName || "").toLowerCase().trim();

    var foundId = "";

    Object.keys(leaveTypesObj).forEach(function(leaveTypeId) {
      var name = String(leaveTypesObj[leaveTypeId].name || "")
        .toLowerCase()
        .trim();

      if (name === targetName) {
        foundId = leaveTypeId;
      }
    });

    return foundId;
  }


  /*******************************************************
   * Match employee data
   *******************************************************/
  var matchedPrevYearEmployeeData = findEmployeeInReport(
    reportPrevYear,
    employeeNameFromTab,
    employeeIdFromTab
  );

  var matchedCurrentYearEmployeeData = findEmployeeInReport(
    reportCurrentYear,
    employeeNameFromTab,
    employeeIdFromTab
  );

  if (!matchedPrevYearEmployeeData && !matchedCurrentYearEmployeeData) {
    sheet.getRange("A1")
      .setValue("Employee not found in Zoho leave summary API: " + employeeNameFromTab + " - " + employeeIdFromTab)
      .setFontWeight("bold")
      .setBackground("#f4cccc");

    return;
  }


  /*******************************************************
   * DOJ / Exit Date Logic
   *******************************************************/
  var empDates = employeeDateMap[employeeIdFromTab] || {};

  var doj = empDates.doj;
  var exitDate = empDates.exitDate;

  var yearStartDate;

  if (doj && doj.getFullYear() === currentYear) {
    yearStartDate = doj;
  } else {
    yearStartDate = new Date(currentYear, 0, 1);
  }

  var dateOfExitValue = exitDate || new Date(currentYear, 11, 31);


  /*******************************************************
   * Leave configuration
   *******************************************************/
  var leaveConfig = [
    {
      shortName: "CL",
      zohoLeaveName: "Casual Leave",
      totalAllot: 7
    },
    {
      shortName: "SL",
      zohoLeaveName: "Sick Leave",
      totalAllot: 8
    },
    {
      shortName: "EL",
      zohoLeaveName: "Earned Leave",
      totalAllot: 15
    }
  ];


  /*******************************************************
   * Clear only output area, not whole sheet
   *******************************************************/
  sheet.getRange("A1:L10").clearContent();
  sheet.getRange("A1:L10").clearFormat();


  /*******************************************************
   * Header
   *******************************************************/
  sheet.getRange("A1").setValue(employeeNameFromTab);
  sheet.getRange("B1").setValue("Previous Y Left");
  sheet.getRange("C1").setValue("Total Leaves Allot");
  sheet.getRange("D1").setValue("Eligible For");
  sheet.getRange("E1").setValue("Actual Used Leaves - CY");
  sheet.getRange("F1").setValue("Balance no.");
  sheet.getRange("G1").setValue("Total leave for Encasment");
  sheet.getRange("H1").setValue("Year Start/DOJ");
  sheet.getRange("I1").setValue("Date of exit");
  sheet.getRange("J1").setValue("Current year Duration in Months");
  sheet.getRange("K1").setValue("Per Month");
  sheet.getRange("L1").setValue("Days");


  /*******************************************************
   * Leave rows
   *******************************************************/
  leaveConfig.forEach(function(cfg, index) {
    var row = 2 + index;

    var zohoLeaveTypeId = findLeaveTypeIdByName(
      leaveTypes,
      cfg.zohoLeaveName
    );

    var prevYearLeaveData =
      zohoLeaveTypeId && matchedPrevYearEmployeeData
        ? matchedPrevYearEmployeeData[zohoLeaveTypeId] || {}
        : {};

    var currentYearLeaveData =
      zohoLeaveTypeId && matchedCurrentYearEmployeeData
        ? matchedCurrentYearEmployeeData[zohoLeaveTypeId] || {}
        : {};

    var previousYLeft = 0;

    if (cfg.shortName === "EL") {
      previousYLeft =
        prevYearLeaveData.balance !== undefined
          ? Number(prevYearLeaveData.balance)
          : 0;
    }

    var actualUsedCY =
      currentYearLeaveData.booked !== undefined
        ? Number(currentYearLeaveData.booked)
        : 0;

    sheet.getRange(row, 1).setValue(cfg.shortName);
    sheet.getRange(row, 2).setValue(previousYLeft);
    sheet.getRange(row, 3).setValue(cfg.totalAllot);
    sheet.getRange(row, 4).setFormula("=L" + row);
    sheet.getRange(row, 5).setValue(actualUsedCY);
    sheet.getRange(row, 6).setFormula("=D" + row + "-E" + row);

    if (cfg.shortName === "EL") {
      sheet.getRange(row, 7).setFormula("=F4+B4+SUMIF(F2:F3,\"<0\")");
    } else {
      sheet.getRange(row, 7).setValue("");
    }

    sheet.getRange(row, 8).setValue(yearStartDate);
    sheet.getRange(row, 9).setValue(dateOfExitValue);

    sheet.getRange(row, 10).setFormula(
      "=(YEAR(I" + row + ")-YEAR(H" + row + "))*12" +
      "+(MONTH(I" + row + ")-MONTH(H" + row + "))" +
      "+(DAY(I" + row + ")+1-DAY(H" + row + "))/DAY(EOMONTH(I" + row + ",0))"
    );

    sheet.getRange(row, 11).setFormula("=C" + row + "/12");
    sheet.getRange(row, 12).setFormula("=K" + row + "*J" + row);
  });

//**************** Salary calculation block — MANUAL VALUES *******************
  sheet.getRange("A6").setValue("Salary Calculation");
  sheet.getRange("A7").setValue("Last Drawn Salary");
  sheet.getRange("A8").setValue("Total leave for Encashment");
  sheet.getRange("A9").setValue("Basic Pay (50% of Total Salary)");
  sheet.getRange("A10").setValue("Total Earned Leave Paid Amount");
  // Restore previous manual values
  sheet.getRange("B7:B10").setValues(oldManualSalaryBlock);
  sheet.getRange("B8").setFormula("=G4");

//*********************Formatting****************
  sheet.getRange("A1:L1")
    .setFontWeight("bold")
    .setBackground("#c9c9c9");

  sheet.getRange("A1")
    .setBackground("#ffff00")
    .setFontSize(14);

  sheet.getRange("A2:A4")
    .setBackground("#c9c9c9")
    .setFontWeight("bold");

  sheet.getRange("B2:B4").setBackground("#ffffff");
  sheet.getRange("E2:E4").setBackground("#ffffff");

  sheet.getRange("C2:D4").setBackground("#ffffff");
  sheet.getRange("F2:L4").setBackground("#ffffff");

  sheet.getRange("A6")
    .setBackground("#c9c9c9")
    .setFontWeight("bold")
    .setFontSize(12);

  sheet.getRange("A7:A10")
    .setBackground("#ffffff")
    .setFontWeight("normal");

  sheet.getRange("A6")
    .setFontWeight("bold")
    .setFontSize(12);

  sheet.getRange("B7:B10")
    .setBackground("#ffffff");

  sheet.getRange("A1:L4")
    .setBorder(true, true, true, true, true, true);

  sheet.getRange("A6:B10")
    .setBorder(true, true, true, true, true, true);

  sheet.getRange("B2:G4").setNumberFormat("0.00");
  sheet.getRange("H2:I4").setNumberFormat("d/m/yyyy");
  sheet.getRange("J2:L4").setNumberFormat("0.00");
  sheet.getRange("B7:B10").setNumberFormat("0.00");

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 170);
  sheet.setColumnWidth(8, 90);
  sheet.setColumnWidth(9, 90);
  sheet.setColumnWidth(10, 260);
  sheet.setColumnWidth(11, 90);
  sheet.setColumnWidth(12, 70);

  sheet.getRange("A1:L10").setHorizontalAlignment("center");
  sheet.getRange("A6:A10").setHorizontalAlignment("left");

  SpreadsheetApp.flush();

  ss.toast(
    "Leave encashment updated for " + employeeNameFromTab + " - " + employeeIdFromTab,
    "Done",
    5
  );
}


// ═══════════════════════════════INSTALL AUTO TRIGGER═══════════════════════════════
function installLeaveEncashmentAutoTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "onLeaveEncashmentSheetChange") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onLeaveEncashmentSheetChange")
    .forSpreadsheet(ss)
    .onChange()
    .create();

  ss.toast("Leave encashment auto trigger installed", "Done", 5);
}


// ═══════════════════════════════AUTO RUN ON NEW EMPLOYEE TAB═══════════════════════════════
function onLeaveEncashmentSheetChange(e) {
  if (!e) return;
  var changeType = e.changeType || "";
  if (changeType !== "INSERT_GRID" && changeType !== "OTHER") {
    return;
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var targetSheet = null;

    for (var i = sheets.length - 1; i >= 0; i--) {
      var sh = sheets[i];
      var tabName = sh.getName();
      var match = tabName.match(/^(.*?)\s*[-–—]\s*(HRM\d+)$/i);
      if (!match) {
        continue;
      }

      var employeeNameFromTab = match[1].trim();
      var a1Value = String(sh.getRange("A1").getValue() || "").trim();
      var b1Value = String(sh.getRange("B1").getValue() || "").trim();
      var alreadyBuilt =
        a1Value.toLowerCase() === employeeNameFromTab.toLowerCase() &&
        b1Value === "Previous Y Left";

      if (alreadyBuilt) {
        continue;
      }
      targetSheet = sh;
      break;
    }

    if (!targetSheet) {
      console.log("No new unprocessed employee tab found.");
      return;
    }

    fetchLeaveEncashmentForCurrentTab(targetSheet);

  } catch (err) {
    console.error("Leave encashment auto fetch error: " + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}
