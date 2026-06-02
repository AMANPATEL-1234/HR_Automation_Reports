
// ════════════════════════════════════GET ACCESS TOKEN══════════════════════════════════════════════
function getAccessToken() {
  var props = PropertiesService.getScriptProperties();
  var response = UrlFetchApp.fetch("https://accounts.zoho.in/oauth/v2/token", {
    method: "post",
    payload: {
      grant_type   : "refresh_token",
      client_id    : props.getProperty("CLIENT_ID"),
      client_secret: props.getProperty("CLIENT_SECRET"),
      refresh_token: props.getProperty("REFRESH_TOKEN")
    },
    muteHttpExceptions: true
  });
  var data = JSON.parse(response.getContentText());
  if (!data.access_token) throw new Error("Token error: " + JSON.stringify(data));
  return data.access_token;
}
//**************************************
function updateSalaryPFDOJAndExitDate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pfSheet = ss.getSheetByName(CONFIG.SALARY_PF_SHEET);
  if (!pfSheet) {
    throw new Error("Sheet not found: " + CONFIG.SALARY_PF_SHEET);
  }

  if (pfSheet.getLastRow() < 2) {
    ss.toast("No employee rows found in Salary and PF Details.", "DOJ / Exit Date", 5);
    return;
  }

  var accessToken = getAccessToken();
  var authHeader = {
    Authorization: "Zoho-oauthtoken " + accessToken
  };

  var empDateMap = {};
  function parseZohoDate_(raw) {
    if (!raw) return "";
    var s = String(raw).trim();
    // yyyy-MM-dd
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      var p1 = s.split("-");
      return new Date(Number(p1[0]), Number(p1[1]) - 1, Number(p1[2]));
    }

    // dd-MMM-yyyy, example: 27-Apr-2026
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

    return "";
  }

  function fetchEmployeesByStatus_(isActiveValue) {
    var startIndex = 0;
    var limit = 200;

    while (true) {
      var url =
        "https://people.zoho.in/people/api/forms/employee/getRecords" +
        "?isActive=" + isActiveValue +
        "&startIndex=" + startIndex +
        "&limit=" + limit;

      var resp = UrlFetchApp.fetch(url, {
        method: "get",
        headers: authHeader,
        muteHttpExceptions: true
      });

      var text = resp.getContentText();
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
          empData["id"] ||
          ""
        ).trim();

        if (!empId) return;

        var rawDOJ =
          empData["Dateofjoining"] ||
          empData["DateOfJoining"] ||
          empData["Date of Joining"] ||
          empData["Joining Date"] ||
          empData["DateofJoining"] ||
          empData["dateOfJoining"] ||
          "";

        var rawExit =
          empData["Dateofexit"] ||
          empData["DateOfExit"] ||
          empData["Date of Exit"] ||
          empData["Exit Date"] ||
          empData["DateofExit"] ||
          empData["dateOfExit"] ||
          "";

        empDateMap[empId] = {
          doj: parseZohoDate_(rawDOJ),
          exitDate: parseZohoDate_(rawExit)
        };
      });

      if (batch.length < limit) break;
      startIndex += limit;
      Utilities.sleep(500);
    }
  }

  fetchEmployeesByStatus_("true");
  fetchEmployeesByStatus_("false");

  var lastRow = pfSheet.getLastRow();
  var lastCol = pfSheet.getLastColumn();
  var headers = pfSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var empIdIdx = getColIndex(headers, [
    "Employee ID",
    "Employee Id",
    "Employee HRM",
    "Emp ID",
    "EmpID"
  ]);

  if (empIdIdx === -1) {
    throw new Error("Employee ID column not found in Salary and PF Details");
  }

  var dojIdx = getColIndex(headers, ["DOJ"]);

  if (dojIdx === -1) {
    pfSheet.insertColumnAfter(pfSheet.getLastColumn());
    dojIdx = pfSheet.getLastColumn() - 1;
    pfSheet.getRange(1, dojIdx + 1).setValue("DOJ");
  }

  headers = pfSheet.getRange(1, 1, 1, pfSheet.getLastColumn()).getValues()[0];

  var exitIdx = getColIndex(headers, ["Exit Date"]);

  if (exitIdx === -1) {
    pfSheet.insertColumnAfter(pfSheet.getLastColumn());
    exitIdx = pfSheet.getLastColumn() - 1;
    pfSheet.getRange(1, exitIdx + 1).setValue("Exit Date");
  }

  lastCol = pfSheet.getLastColumn();
  var data = pfSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var dojValues = [];
  var exitValues = [];

  for (var r = 0; r < data.length; r++) {
    var empId = String(data[r][empIdIdx] || "").trim();
    if (!empId || !empDateMap[empId]) {
      dojValues.push([""]);
      exitValues.push([""]);
      continue;
    }

    dojValues.push([empDateMap[empId].doj || ""]);
    exitValues.push([empDateMap[empId].exitDate || ""]);
  }

  pfSheet.getRange(2, dojIdx + 1, dojValues.length, 1).setValues(dojValues);
  pfSheet.getRange(2, exitIdx + 1, exitValues.length, 1).setValues(exitValues);

  pfSheet.getRange(2, dojIdx + 1, dojValues.length, 1).setNumberFormat("d mmm yyyy");
  pfSheet.getRange(2, exitIdx + 1, exitValues.length, 1).setNumberFormat("d mmm yyyy");

  pfSheet.getRange(1, 1, 1, pfSheet.getLastColumn())
    .setFontWeight("bold")
    .setBackground("#d9e2f3");

  pfSheet.autoResizeColumns(1, pfSheet.getLastColumn());
  ss.toast(
    "DOJ and Exit Date updated in Salary and PF Details.",
    "Done",
    5
  );
}
//************************************** 
function fetchEmployeeJoinDateMap(authHeader) {
  var joinMap = {};
  var startIndex = 0;
  var limit = 200;

  function parseJoinDate(raw) {
    if (!raw) return null;
    var s = String(raw).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      var p = s.split("-");
      return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
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

  while (true) {
    var url =
      "https://people.zoho.in/people/api/forms/employee/getRecords" +
      "?isActive=true" +
      "&startIndex=" + startIndex +
      "&limit=" + limit;

    var resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: authHeader,
      muteHttpExceptions: true
    });

    // console.log("EMPLOYEE MASTER API URL:");
    // console.log(url);
    // console.log("EMPLOYEE MASTER API STATUS:");
    // console.log(resp.getResponseCode());
    // console.log("EMPLOYEE MASTER API RAW RESPONSE:");
    // console.log(resp.getContentText());

    var data = JSON.parse(resp.getContentText());

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

      if (Array.isArray(empData)) empData = empData[0] || {};

      var empId = String(
        empData["EmployeeID"] ||
        empData["EmployeeId"] ||
        empData["Employee ID"] ||
        empData["id"] ||
        ""
      ).trim();

      var rawJoinDate =
        empData["Dateofjoining"] ||
        empData["DateOfJoining"] ||
        empData["Date of Joining"] ||
        empData["Joining Date"] ||
        empData["DateofJoining"] ||
        empData["dateOfJoining"] ||
        "";

      if (!empId) return;

      var joinDate = parseJoinDate(rawJoinDate);
      if (joinDate) {
        joinMap[empId] = joinDate;
      }
    });

    if (batch.length < limit) break;

    startIndex += limit;
    Utilities.sleep(500);
  }

  // console.log("JOIN DATE MAP:");
  // console.log(JSON.stringify(joinMap, null, 2));
  return joinMap;
}

//************resigned ******************
function fetchResignedAttendance(authHeader, sdateStr, edateStr) {
  var results = [];
  var startIndex = 0;
  var batchSize = 100;
  var resignedEmps = [];

  var monthStart = new Date(sdateStr);
  var monthEnd = new Date(edateStr);

  function parseExitDate(raw) {
    if (!raw) return null;

    var s = String(raw).trim();

    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      var p = s.split("-");
      return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
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

  while (true) {
    var url =
      "https://people.zoho.in/people/api/forms/employee/getRecords" +
      "?isActive=false" +
      "&startIndex=" + startIndex +
      "&limit=" + batchSize;

    var resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: authHeader,
      muteHttpExceptions: true
    });

    // console.log("RESIGNED EMPLOYEE API URL:");
    // console.log(url);

    // console.log("RESIGNED EMPLOYEE API STATUS:");
    // console.log(resp.getResponseCode());

    // console.log("RESIGNED EMPLOYEE API RAW RESPONSE:");
    // console.log(resp.getContentText());

    var data = JSON.parse(resp.getContentText());

    // console.log("RESIGNED EMPLOYEE API PARSED JSON:");
    // console.log(JSON.stringify(data, null, 2));

    var batch = [];

    if (Array.isArray(data)) {
      batch = data;
    } else if (data && data.response && data.response.result) {
      batch = data.response.result;
    }

    if (!batch || batch.length === 0) break;

    resignedEmps = resignedEmps.concat(batch);

    if (batch.length < batchSize) break;

    startIndex += batchSize;
    Utilities.sleep(300);
  }

  // console.log("Inactive employees found: " + resignedEmps.length);
  var eligibleCount = 0;
  resignedEmps.forEach(function(empRecord) {
    var recordId = Object.keys(empRecord)[0];
    var empData = empRecord[recordId];

    if (Array.isArray(empData)) empData = empData[0] || {};

    var hrmId = String(empData["EmployeeID"] || "").trim();
    var firstName = String(empData["FirstName"] || "").trim();
    var lastName = String(empData["LastName"] || "").trim();
    var empStatus = String(empData["Employeestatus"] || "").trim().toLowerCase();
    var dateOfExit = String(empData["Dateofexit"] || "").trim();

    if (!hrmId) return;
    if (empStatus !== "resigned") return;
    if (!dateOfExit) return;

    var exitDate = parseExitDate(dateOfExit);
    if (!exitDate) return;
    if (exitDate < monthStart) {
      return;
    }

    eligibleCount++;
    var attUrl =
      "https://people.zoho.in/people/api/attendance/getUserReport" +
      "?sdate=" + encodeURIComponent(sdateStr) +
      "&edate=" + encodeURIComponent(edateStr) +
      "&dateFormat=yyyy-MM-dd" +
      "&empId=" + encodeURIComponent(hrmId) +
      "&limit=1";

    try {
      var attResp = UrlFetchApp.fetch(attUrl, {
      method: "get",
      headers: authHeader,
      muteHttpExceptions: true
    });
      // console.log("RESIGNED ATTENDANCE API URL:");
      // console.log(attUrl);

      // console.log("RESIGNED ATTENDANCE API EMPLOYEE:");
      // console.log(hrmId);

      // console.log("RESIGNED ATTENDANCE API STATUS:");
      // console.log(attResp.getResponseCode());

      // console.log("RESIGNED ATTENDANCE API RAW RESPONSE:");
      // console.log(attResp.getContentText());

      var attData = JSON.parse(attResp.getContentText());
      // console.log("RESIGNED ATTENDANCE API PARSED JSON:");
      // console.log(JSON.stringify(attData, null, 2));

      if (attData && typeof attData === "object" && !attData.response && Object.keys(attData).length > 0) {
        results.push({
          employeeDetails: {
            id: hrmId,
            "first name": firstName,
            "last name": lastName,
            dateOfExit: dateOfExit
          },
          attendanceDetails: attData
        });

        // console.log("Resigned attendance fetched: " + hrmId);
      }

      Utilities.sleep(200);
    } catch (err) {
      console.error("Error fetching resigned attendance for " + hrmId + ": " + err.message);
    }
  });

  // console.log("Eligible resigned employees for selected month: " + eligibleCount);
  // console.log("Total resigned attendance fetched: " + results.length);
  return results;
}

//*************************************************
function fetchODRequestMap(authHeader, sdate, edate) {
  var odMap = {};
  var startIndex = 0;

  function fmtODDate(d) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  }

  function normalizeDate(raw) {
    if (!raw) return "";
    var s = String(raw).trim();
    // Already yyyy-MM-dd
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    // dd-MMM-yyyy, example: 05-Mar-2026
    var m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
    if (m) {
      var months = {
        Jan: "01", Feb: "02", Mar: "03", Apr: "04",
        May: "05", Jun: "06", Jul: "07", Aug: "08",
        Sep: "09", Oct: "10", Nov: "11", Dec: "12"
      };

      return m[3] + "-" + months[m[2]] + "-" + ("0" + m[1]).slice(-2);
    }

    return "";
  }

  function detectHalf(obj) {
    var text = JSON.stringify(obj || {}).toLowerCase();

    if (text.indexOf("first half") !== -1 || text.indexOf("firsthalf") !== -1) {
      return "FIRST_HALF";
    }

    if (text.indexOf("second half") !== -1 || text.indexOf("secondhalf") !== -1) {
      return "SECOND_HALF";
    }

    if (
      text.indexOf("0.5") !== -1 ||
      text.indexOf("half day") !== -1 ||
      text.indexOf("halfday") !== -1
    ) {
      return "HALF_UNKNOWN";
    }

    return "FULL_DAY";
  }

  while (true) {
    var params =
      "sDate=" + encodeURIComponent(fmtODDate(sdate)) +
      "&eDate=" + encodeURIComponent(fmtODDate(edate)) +
      "&startIndex=" + startIndex;

    var url = "https://people.zoho.in/people/api/attendance/getMyODRequest?" + params;

    var resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: authHeader,
      muteHttpExceptions: true
    });
    // console.log("OD API URL:");
    // console.log(url);

    // console.log("OD API STATUS:");
    // console.log(resp.getResponseCode());

    // console.log("OD API RAW RESPONSE:");
    // console.log(resp.getContentText());

    var text = resp.getContentText();
    var json;

    try {
      json = JSON.parse(text);
      // console.log("OD API PARSED JSON:");
      // console.log(JSON.stringify(json, null, 2));
    } catch (err) {
      console.error("OD API parse error: " + text);
      break;
    }

    var list = [];
    if (json.result && Array.isArray(json.result.list)) {
      list = json.result.list;
    } else if (Array.isArray(json.result)) {
      list = json.result;
    } else if (Array.isArray(json)) {
      list = json;
    }

    if (!list || list.length === 0) break;

    list.forEach(function(item) {
      var approval = String(
        item.approvalStatus ||
        item.ApprovalStatus ||
        item.status ||
        item.Status ||
        ""
      ).toLowerCase();

      if (
        approval.indexOf("reject") !== -1 ||
        approval.indexOf("cancel") !== -1
      ) {
        return;
      }

      var empId = String(
        item.employeeId ||
        item.employeeID ||
        item.EmployeeID ||
        item["Employee ID"] ||
        item.erecno ||
        ""
      ).trim();

      if (!empId) return;

      var details = item.odDetails || item.ODDetails || item.details || [];

      if (Array.isArray(details) && details.length > 0) {
        details.forEach(function(d) {
          var rawDate =
            d.originday ||
            d.originDay ||
            d.date ||
            d.Date ||
            d.odDate ||
            item.startDate ||
            item.sDate ||
            "";

          var dateKey = normalizeDate(rawDate);
          if (!dateKey) return;

         odMap[empId + "|" + dateKey] = {
         half: detectHalf(d),
         dayTaken: d.dayTaken,
         unit: item.unit,
         odType: item.type || item.Type || "",
         desc: item.desc || item.description || "",
         approvalStatus: d.indApprovalStatus || item.approvalStatus || ""
       };
        });
      } else {
        var startDate = normalizeDate(
          item.startDate ||
          item.sDate ||
          item.fromDate ||
          item.fDate ||
          ""
        );

        var endDate = normalizeDate(
          item.endDate ||
          item.eDate ||
          item.toDate ||
          item.tDate ||
          ""
        );

        if (!startDate) return;
        if (!endDate) endDate = startDate;

        var d1 = new Date(startDate);
        var d2 = new Date(endDate);

        while (d1 <= d2) {
          var dk = Utilities.formatDate(d1, Session.getScriptTimeZone(), "yyyy-MM-dd");

          odMap[empId + "|" + dk] = {
          half: detectHalf(item),
          dayTaken: item.dayTaken,
          unit: item.unit,
          odType: item.type || item.Type || "",
          desc: item.desc || item.description || "",
          approvalStatus: item.approvalStatus || ""
        };
          d1.setDate(d1.getDate() + 1);
        }
      }
    });

    startIndex += list.length;

    if (list.length < 200) break;

    Utilities.sleep(1000);
  }

  // console.log("OD Map Count: " + Object.keys(odMap).length);
  return odMap;
}

//*********************************************** */
function fetchLeaveRequestMap(authHeader, sdate, edate) {
  var leaveMap = {};
  var startIndex = 0;
  var limit = 200;

  function fmtLeaveDate(d) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  }

  while (true) {
    var params =
    "from=" + encodeURIComponent(fmtLeaveDate(sdate)) +
    "&to=" + encodeURIComponent(fmtLeaveDate(edate)) +
    "&dateFormat=" + encodeURIComponent("dd-MMM-yyyy") +
    "&startIndex=" + startIndex +
    "&limit=" + limit +
    "&dataSelect=" + encodeURIComponent("ALL");

    var url = "https://people.zoho.in/people/api/v2/leavetracker/leaves/records?" + params;

    var resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: authHeader,
      muteHttpExceptions: true
    });

    // console.log("LEAVE API URL:");
    // console.log(url);

    // console.log("LEAVE API STATUS:");
    // console.log(resp.getResponseCode());

    // console.log("LEAVE API RAW RESPONSE:");
    // console.log(resp.getContentText());

    var text = resp.getContentText();
    var json;

    try {
      json = JSON.parse(text);
      console.log("LEAVE API PARSED JSON:");
      console.log(JSON.stringify(json, null, 2));
    } catch (err) {
      console.error("Leave API parse error: " + text);
      break;
    }

    // console.log("Leave API URL: " + url);
    var recordsObj = json.records || {};
    var recordKeys = Object.keys(recordsObj);

    if (recordKeys.length === 0) break;

    recordKeys.forEach(function(recordId) {
      var rec = recordsObj[recordId] || {};

      var approval = String(rec.ApprovalStatus || rec.approvalStatus || "").toLowerCase();
      if (approval.indexOf("reject") !== -1 || approval.indexOf("cancel") !== -1) return;

      /*
        IMPORTANT:
        Zoho leave API may return internal Employee.ID, not HRM ID.
        If your response has Employee_ID or EmployeeID as HRM07, this works directly.
        Otherwise we need one extra employee-id map from Zoho record id to HRM ID.
      */
      var empId = String(
      rec["EmployeeId"] ||      // IMPORTANT: your API response has this: HRM07
      rec["Employee_ID"] ||
      rec["Employee ID"] ||
      rec["EmployeeID"] ||
      rec["Employee HRM"] ||
      rec["Employee.HRM"] ||
      ""
    ).trim();

      // If HRM ID is not available, keep internal id temporarily.
      if (!empId) {
        empId = String(rec["Employee.ID"] || rec["Employee.Id"] || "").trim();
      }

      if (!empId) return;

      var leaveCode = leaveTypeToCode(rec.Leavetype || rec.LeaveType || rec["Leave Type"]);
      var days = rec.Days || rec.days || {};

      Object.keys(days).forEach(function(dayKeyRaw) {
        var dayInfo = days[dayKeyRaw] || {};
        var dateKey = toDateKeyFromZohoDate(dayKeyRaw);
        if (!dateKey) return;
        var leaveCount = String(dayInfo.LeaveCount || dayInfo.leaveCount || "").trim();
        if (
      leaveCount === "0" ||
      leaveCount === "0.0" ||
      leaveCount === ""
    ) {
      return;
    }
        var session = String(dayInfo.Session || dayInfo.session || "").trim();

        var half = "FULL_DAY";

        if (leaveCount === "0.5" || leaveCount === ".5") {
          if (session === "1") {
            half = "FIRST_HALF";
          } else if (session === "2") {
            half = "SECOND_HALF";
          } else {
            half = "HALF_UNKNOWN";
          }
        }

        leaveMap[empId + "|" + dateKey] = {
          code: leaveCode,
          half: half,
          leaveCount: leaveCount,
          session: session
        };
      });
    });

    if (recordKeys.length < limit) break;

    startIndex += limit;
    Utilities.sleep(1000);
  }

  // console.log("Leave Map Count: " + Object.keys(leaveMap).length);
  return leaveMap;
}




















