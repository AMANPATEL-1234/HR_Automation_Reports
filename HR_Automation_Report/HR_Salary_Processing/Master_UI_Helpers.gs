
//******************************************
function combineAttendanceParts(left, right, isHalfDay) {
  left = String(left || "").trim().toUpperCase();
  right = String(right || "").trim().toUpperCase();
  if (!left || left === "-") return right || "-";
  if (!right || right === "-") return left;
  if (left === right) return left;
  // Avoid double 0.5 if already exists
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
// ═══════════════════════════════════════Final sheet creation══════════════════════════════════════
function buildFinalEmployeeSS(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetName = sheetName || CONFIG.FINAL_EMPLOYEE_SS_SHEET;
  var sheet = ss.getSheetByName(targetName);
  if (!sheet) { sheet = ss.insertSheet(targetName); } else { sheet.clearContents(); sheet.clearFormats(); }
  // Get selected month/year for dynamic header
  var props     = PropertiesService.getScriptProperties();
  var today     = new Date();
  var _month    = props.getProperty("SELECTED_MONTH");
  var _year     = props.getProperty("SELECTED_YEAR");
  var calcMonth = _month ? parseInt(_month) : today.getMonth();
  var calcYear  = _year  ? parseInt(_year)  : today.getFullYear();
  var totalDays = new Date(calcYear, calcMonth + 1, 0).getDate();
  var monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  var monthLabel = monthNames[calcMonth];
  // Headers
  var headers = [
  "Entities",
  "Employee ID",
  "Employee Name",
  "Total Day(s) - " + monthLabel,
  "Total Payable Days",
  "Monthly Fixed",
  "Basic Pay",
  "HRA",
  "Transport Allowance",
  "Special Allowance",
  monthLabel + " Payable",
  "Compensation Request",
  "One Time",
  "Arrears",
  "Reimbursement",
  "Referral",
  "Gross Payable Amount",  
  "ESI/PF",               
  "TDS",                 
  "Advance given salary/Other Deductions", 
  "Net Payable",          
  "Status",               
  "Remarks"              
];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  // Read source sheets
  var salSheet  = ss.getSheetByName(CONFIG.SALARY_SHEET);
  var attSheet  = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if (!salSheet) throw new Error("Salary sheet not found");
  if (!attSheet) throw new Error("Attendance sheet not found");
  var salVals  = salSheet.getDataRange().getValues();
  var salHdrs  = salVals[0];
  // Helper to get col index from salary sheet
  function nk(s){return String(s||"").toLowerCase().replace(/\s+/g,"").replace(/[^a-z0-9]/g,"");}
  function gci(names){
    for(var i=0;i<names.length;i++){
      for(var j=0;j<salHdrs.length;j++){
        if(nk(salHdrs[j])===nk(names[i])) return j;
      }
    }
    return -1;
  }

  // Get column indices from salary sheet
  var idxEntity   = gci(["Entities","Entity"]);
  var idxEmpId    = gci(["Employee ID","Employee Id"]);
  var idxEmpName  = gci(["Employee Name"]);
  var idxPayDays  = gci(["Total Payable Days"]);
  var idxMonFixed = gci(["Monthly Fixed"]);
  var idxComp     = gci(["Compensation request","Compensation Request"]);
  var idxReimb    = gci(["Reimbursement"]);
  var idxESIPF    = gci(["ESI/PF"]);
  // Build data rows
  var dataRows = [];
  for (var r = 1; r < salVals.length; r++) {
    var row = salVals[r];
    var empId = String(row[idxEmpId] || "").trim();
    if (!empId) continue;
    dataRows.push([
      idxEntity  >= 0 ? row[idxEntity]  : "",  // A
      empId,                                    // B
      idxEmpName >= 0 ? row[idxEmpName] : "",  // C
      totalDays,                                // D
      idxPayDays >= 0 ? row[idxPayDays] : "",  // E
      idxMonFixed>= 0 ? row[idxMonFixed]: "",  // F
      "",  // G - Basic Pay
      "",  // H - HRA
      "",  // I - Transport
      "",  // J - Special Allowance
      "",  // K - Month Payable
      idxComp  >= 0 ? row[idxComp]  : "",      // L
      "",  // M - One Time
      "",  // N - Arrears
      idxReimb >= 0 ? row[idxReimb] : "",      // O
      "",  // P - Referral
      "",  // Q - Gross Payable (formula)
      "",  // R - ESI/PF, formula applied below
      "",  // S - TDS (was R)
      "",  // T - Advance
      "",  // U - Net Payable
      "",  // V - Status
      ""   // W - Remarks
]);
  }

  if (dataRows.length === 0) {
    sheet.getRange(2, 1).setValue("No data found");
    return;
  }
  var finalBackup = readFinalSSBackupFromDrive_(calcYear, calcMonth);
  // Write data
  sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  // Restore saved permanent column values per employee
  var restoredCount = 0;
  for (var i = 0; i < dataRows.length; i++) {
    var rowEmpId = String(dataRows[i][1] || "").trim(); // col B = index 1
    var bk = finalBackup[rowEmpId];
    if (!bk) continue;
    var R2 = i + 2;
    if (bk.oneTime  !== "") sheet.getRange(R2, 13).setValue(bk.oneTime);   
    if (bk.arrears  !== "") sheet.getRange(R2, 14).setValue(bk.arrears);   
    if (bk.referral !== "") sheet.getRange(R2, 16).setValue(bk.referral);  
    if (bk.tds      !== "") sheet.getRange(R2, 19).setValue(bk.tds);       
    if (bk.advance  !== "") sheet.getRange(R2, 20).setValue(bk.advance);   
    if (bk.status   !== "") sheet.getRange(R2, 22).setValue(bk.status);   
    if (bk.remarks  !== "") sheet.getRange(R2, 23).setValue(bk.remarks);   
    restoredCount++;
  }

  // Apply formulas row by row
  for (var i = 0; i < dataRows.length; i++) {
    var R = String(i + 2);
    sheet.getRange(i+2, 7).setFormula("=ROUND((F"+R+"/2)/D"+R+"*E"+R+",2)");   // G BasicPay
    sheet.getRange(i+2, 8).setFormula("=ROUND(G"+R+"/2,2)");                    // H HRA
    sheet.getRange(i+2, 9).setFormula("=ROUND(1600/D"+R+"*E"+R+",2)");         // I Transport
    sheet.getRange(i+2,11).setFormula("=ROUND(F"+R+"/D"+R+"*E"+R+",2)");
    sheet.getRange(i+2,10).setFormula("=ROUND(K"+R+"-SUM(G"+R+":I"+R+"),2)");
    sheet.getRange(i+2,17).setFormula("=ROUND(K"+R+"+L"+R+"+M"+R+"+N"+R+"+O"+R+"+P"+R+",2)"); // Q GrossPayable (was col 19)
    sheet.getRange(i+2, 18).setFormula(
   "=IFERROR(INDEX('" + CONFIG.SALARY_SHEET + "'!Y:Y,MATCH(B" + R + ",'" + CONFIG.SALARY_SHEET + "'!B:B,0)),0)"
);
  sheet.getRange(i+2, 21).setFormula("=Q"+R+"-R"+R+"-S"+R+"-T"+R);
  }
  // Formatting
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9e2f3").setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet.getRange(1, 1, sheet.getLastRow(), headers.length).setBorder(true,true,true,true,true,true);
  SpreadsheetApp.flush();

  // ────────── Show restoration notification ────────────
  if (restoredCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "✔ Permanent data restored for " + restoredCount + " employee(s) — One Time, Arrears, Referral, TDS, Advance, Status, Remarks",
      "Final SS Restored",
      6
    );
  }
  setFinalSSVisibleMonthYear_(sheet, calcYear, calcMonth);
  return {sheetName: targetName, rowsWritten: dataRows.length};
}

//****************************************************
function clearEmployeeSheetsWhenNoAttendance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetsToClear = [
    CONFIG.REIMBURSEMENT_SHEET,
    CONFIG.COMPENSATION_SHEET,
    CONFIG.SALARY_SHEET,
    CONFIG.FINAL_EMPLOYEE_SS_SHEET
  ];

  sheetsToClear.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    sheet.clearContents();
    sheet.clearFormats();

    sheet.getRange(1, 1)
      .setValue("No employee attendance data found for selected month/year")
      .setFontWeight("bold")
      .setBackground("#f4cccc");

    sheet.autoResizeColumn(1);
  });

  SpreadsheetApp.flush();
}


// ═════════════════════════════════ REIMBURSEMENT════════════════════════════════════════════════════════
function reimbursement(ACCESS_TOKEN, targetSheetName, selMonth, selYear) {
  var hdrs = { Authorization: "Zoho-oauthtoken " + ACCESS_TOKEN };
  var allData = [];
  var startIndex = 0;
  var batchSize = 15;
  var filterMonth = parseInt(selMonth, 10);
  var filterYear = parseInt(selYear, 10);
  var monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  var seenRecordIds = {};
  var foundTargetMonth = false;
  while (true) {
    var url = "https://people.zoho.in/api/forms/reimbursement_view/records?startIndex=" + startIndex + "&limit=" + batchSize;
    var resp;
    try {
      resp = UrlFetchApp.fetch(url, {
        method: "get",
        headers: hdrs,
        muteHttpExceptions: true
      });
    } catch (err) {
      console.error("REIMBURSEMENT API FETCH ERROR at startIndex " + startIndex + ": " + err.message);
      break;
    }

    console.log("REIMBURSEMENT API URL:");
    console.log(url);
    console.log("REIMBURSEMENT API STATUS:");
    console.log(resp.getResponseCode());
    console.log("REIMBURSEMENT API RAW RESPONSE:");
    console.log(resp.getContentText());
    var batch = JSON.parse(resp.getContentText());
    console.log("REIMBURSEMENT API PARSED JSON:");
    console.log(JSON.stringify(batch, null, 2));
    if (!Array.isArray(batch) || batch.length === 0) break;

    var newRecordFound = false;
    var stopBecauseOlderMonth = false;
    batch.forEach(function(item) {
      var recordId = String(item["recordId"] || "").trim();
      if (recordId && seenRecordIds[recordId]) {
        return;
      }

      if (recordId) {
        seenRecordIds[recordId] = true;
        newRecordFound = true;
      }

      var rawDate =
        item["Date of Payment"] ||
        item["Date of payment"] ||
        "";

      if (!rawDate) return;

      var parts = String(rawDate).split("-");
      if (parts.length < 3) return;

      var itemMonth = monthNames.indexOf(parts[1]);
      var itemYear = parseInt(parts[2], 10);
      var itemDay = parseInt(parts[0], 10);

      // Keep only selected attendance month reimbursement
      if (itemMonth === filterMonth && itemYear === filterYear) {
        allData.push(item);
        foundTargetMonth = true;
        return;
      }

      // If selected month data was found and now older month started, stop fetching more
      var itemDate = new Date(itemYear, itemMonth, itemDay);
      var targetMonthStart = new Date(filterYear, filterMonth, 1);

      if (foundTargetMonth && itemDate < targetMonthStart) {
        stopBecauseOlderMonth = true;
      }
    });

    if (!newRecordFound) {
      console.log("Stopping reimbursement fetch because Zoho repeated same records.");
      break;
    }

    if (stopBecauseOlderMonth) {
      console.log("Stopping reimbursement fetch because older month started.");
      break;
    }

    if (batch.length < batchSize) break;

    startIndex += batchSize;
    Utilities.sleep(1000);
  }
  var data = allData;
  var sn = targetSheetName || "Reimbursement";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sn);
  if (!sheet) { sheet = ss.insertSheet(sn); } else { sheet.clearContents(); sheet.clearFormats(); }

  if (!Array.isArray(data) || data.length === 0) {
    sheet.clearContents();
    sheet.clearFormats();
    sheet.getRange(1, 1).setValue("No reimbursement data found for selected month/year").setFontWeight("bold");
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "⚠️ No reimbursement records found for selected month — sheet cleared.",
      "Reimbursement", 5
    );
    return [];
  }

 var hRow = [
  "ZOHO_LINK_ID",
  "Added Time",
  "Employee ID",
  "First Name",
  "Last Name",
  "Department",
  "Approval Status",
  "Date of Payment",
  "Total Amount",
  "Bill Receipt",
  "Payment Image",
  "Description",
  "Event Name",
  "Expense Type",
  "Payment Mode",
  "Reporting Manager",
  "Approver",
  "Approval Time"
];

function getReimbursementValue(item, colName) {
  if (colName === "ZOHO_LINK_ID") {
    return item["recordId"] || "";
  }

  if (colName === "Approval Status") {
    return item["ApprovalStatus"] || item["Approval Status"] || "";
  }

  if (colName === "Bill Receipt") {
    return (
      item[" Bill Receipt"] ||
      item["Bill Receipt"] ||
      ""
    );
  }

  if (colName === "Payment Image") {
    return (
      item["Payment Image"] ||
      ""
    );
  }

  if (colName === "Approver") {
    var approvers =
      item["approvalDetails"] &&
      item["approvalDetails"]["approvers"];

    if (Array.isArray(approvers) && approvers.length > 0) {
      return approvers.map(function(a) {
        return a["displayName"] || "";
      }).filter(String).join(", ");
    }

    return "";
  }

  if (colName === "Approval Time") {
    return (
      item["ApprovalTime"] ||
      (
        item["approvalDetails"] &&
        item["approvalDetails"]["approvalCompletionTime"]
      ) ||
      ""
    );
  }

  return decodeHtmlEntities_(item[colName] != null ? item[colName] : "");
}
  var seenReimb = {};
  var vals = [hRow];
  var duplicateRows = [];

  data.forEach(function (item) {
    var empId   = String(item["Employee ID"] || item["Employee HRM"] || "").trim();
    var payDate = String(
    item["Date of Payment"] ||
    item["Date of payment"] ||
    ""
  ).trim();
    var key = empId + "|" + payDate + "|" + String(item["Total Amount"] || "").trim();

    if (empId && payDate) {
      if (!seenReimb[key]) {
        seenReimb[key] = true;
      } else {
        // Second and further occurrences are duplicates
        duplicateRows.push(vals.length + 1); // +1 because header is row 1
      }
    }
    vals.push(hRow.map(function (h) {
    return getReimbursementValue(item, h);
    }));
  });

  // ── Write all data to sheet ───────────────────────────────────────────
  sheet.getRange(1, 1, vals.length, vals[0].length).setValues(vals);

  // ── Highlight duplicate rows in red ──────────────────────────────────
  if (duplicateRows.length > 0) {
    var totalCols = vals[0].length;
    duplicateRows.forEach(function (rowNum) {
      sheet.getRange(rowNum, 1, 1, totalCols)
        .setBackground("#f4cccc")
        .setFontColor("#cc0000");
    });
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "⚠️ " + duplicateRows.length + " duplicate row(s) highlighted in red (same Employee + Date).",
      "Duplicates Found", 6
    );
  }

  // ── Formatting ────────────────────────────────────────────────────────
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setHorizontalAlignment("left").setVerticalAlignment("middle");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setWrap(true);
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setBorder(true, true, true, true, true, true);
  var ACCESS_TOKEN_IMG = getAccessToken();

  var imgHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var paymentImgIdx = getColIndex(imgHeaders, ["Payment Image"]);
  var billImgIdx = getColIndex(imgHeaders, ["Bill Receipt"]);

  var selectedMonthName = monthNames[filterMonth];
  var parentFolderId = PropertiesService
  .getScriptProperties()
  .getProperty(CONFIG.PROP_BACKUP_PARENT_FOLDER_ID);

  if (!parentFolderId) {
    throw new Error(
      "Parent folder ID not found in Script Properties key: " +
      CONFIG.PROP_BACKUP_PARENT_FOLDER_ID
    );
  }

  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var hrDataFolders = parentFolder.getFoldersByName("Hr_Data");
  if (!hrDataFolders.hasNext()) {
    throw new Error("Subfolder not found inside parent folder: Hr_Data");
  }

  var hrDataFolder = hrDataFolders.next();
  var reimbursementFolders = hrDataFolder.getFoldersByName(CONFIG.REIMBURSEMENT_IMAGE_SUBFOLDER_NAME);
  if (!reimbursementFolders.hasNext()) {
    throw new Error("Subfolder not found inside Hr_Data: " + CONFIG.REIMBURSEMENT_IMAGE_SUBFOLDER_NAME);
  }

  var folder = reimbursementFolders.next();
  console.log("Using reimbursement image folder:");
  console.log(folder.getName());
  console.log(folder.getId());
  console.log("IMAGE UPLOAD SECTION STARTED");
  console.log("Drive folder found: " + folder.getName());
  console.log("paymentImgIdx: " + paymentImgIdx);
  console.log("billImgIdx: " + billImgIdx);

  function setDriveLinkInCell(sheetRow, colNum, linkText, driveUrl) {
    var richText = SpreadsheetApp.newRichTextValue()
      .setText(linkText)
      .setLinkUrl(driveUrl)
      .build();

    sheet.getRange(sheetRow, colNum).setRichTextValue(richText);
  }

  function getExtFromOriginalName(originalFileName) {
    var name = String(originalFileName || "").trim();
    var m = name.match(/\.([A-Za-z0-9]+)$/);

    if (m) {
      return m[1].toLowerCase();
    }

    return "";
  }

  function detectRealFileType(bytes) {
    if (!bytes || bytes.length < 4) {
      return "";
    }

    // PDF = %PDF
    if (
      bytes[0] === 37 &&
      bytes[1] === 80 &&
      bytes[2] === 68 &&
      bytes[3] === 70
    ) {
      return "pdf";
    }

    // JPG = FF D8 FF
    if (
      bytes[0] === -1 &&
      bytes[1] === -40 &&
      bytes[2] === -1
    ) {
      return "jpg";
    }

    // PNG = 89 50 4E 47
    if (
      bytes[0] === -119 &&
      bytes[1] === 80 &&
      bytes[2] === 78 &&
      bytes[3] === 71
    ) {
      return "png";
    }

    return "";
  }

  function getContentTypeFromExt(ext) {
    ext = String(ext || "").toLowerCase();

    if (ext === "pdf") return "application/pdf";
    if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
    if (ext === "png") return "image/png";
    if (ext === "webp") return "image/webp";

    return "";
  }

  function findExistingFileByBaseName(fileNameBase) {
    var files = folder.getFiles();

    while (files.hasNext()) {
      var f = files.next();
      if (f.getName().indexOf(fileNameBase) === 0) {
        return f;
      }
    }

    return null;
  }

  function uploadOneReimbursementFile(item, sheetRow, colIdx, imageType) {
    if (colIdx === -1) return;

    var colNum = colIdx + 1;

    var recordId = String(item["recordId"] || "").trim();
    if (!recordId) {
      recordId = "row_" + sheetRow;
    }

    var rawUrl = "";
    var originalFileName = "";
    var linkText = "";

    if (imageType === "payment") {
      rawUrl = String(item["Payment Image_downloadUrl"] || "").trim();
      originalFileName = String(item["Payment Image"] || "").trim();
      linkText = "Open Payment Image";
    } else {
      rawUrl = String(
        item[" Bill Receipt_downloadUrl"] ||
        item["Bill Receipt_downloadUrl"] ||
        ""
      ).trim();

      originalFileName = String(
        item[" Bill Receipt"] ||
        item["Bill Receipt"] ||
        ""
      ).trim();

      linkText = "Open Bill Receipt";
    }

    if (!rawUrl || rawUrl.indexOf("http") === -1) {
      console.log("No Zoho URL found | row " + sheetRow + " | type " + imageType);
      return;
    }

    var originalExt = getExtFromOriginalName(originalFileName);

    var fileNameBase =
      "reimb_" +
      filterYear + "_" +
      selectedMonthName + "_" +
      recordId + "_" +
      imageType;

    var existingFile = findExistingFileByBaseName(fileNameBase);

    if (existingFile) {
      var existingName = existingFile.getName();
      var existingExt = getExtFromOriginalName(existingName);

      // If old saved file extension does not match original Zoho file extension,
      // delete old file and upload again.
      if (originalExt && existingExt && originalExt !== existingExt) {
        console.log("Old wrong file extension found. Deleting: " + existingName);
        existingFile.setTrashed(true);
      } else {
        var existingDriveUrl = "https://drive.google.com/file/d/" + existingFile.getId() + "/preview";
        setDriveLinkInCell(sheetRow, colNum, linkText, existingDriveUrl);

        console.log("Existing Drive file used: " + existingName);
        return;
      }
    }

    try {
      console.log("Downloading Zoho file | row " + sheetRow + " | type " + imageType);
      console.log("Original file name: " + originalFileName);
      console.log("Zoho URL: " + rawUrl);

      var imgResp = UrlFetchApp.fetch(rawUrl, {
        method: "get",
        headers: {
          Authorization: "Zoho-oauthtoken " + ACCESS_TOKEN_IMG
        },
        muteHttpExceptions: true
      });

      console.log("Image response status row " + sheetRow + ": " + imgResp.getResponseCode());

      if (imgResp.getResponseCode() !== 200) {
        console.log("Image download failed row " + sheetRow);
        console.log(imgResp.getContentText());
        return;
      }

      var blob = imgResp.getBlob();
      var bytes = blob.getBytes();

      console.log("Blob size row " + sheetRow + ": " + bytes.length);
      console.log("Blob content type row " + sheetRow + ": " + blob.getContentType());
      console.log("First bytes row " + sheetRow + ": " + bytes.slice(0, 10).join(","));

      var realExt = detectRealFileType(bytes);

      if (!realExt) {
        console.log("File is not valid PDF/JPG/PNG from Zoho. Row " + sheetRow);
        sheet.getRange(sheetRow, colNum).setValue("File not previewable");
        return;
      }

      // Use real detected extension, not guessed extension.
      var ext = realExt;
      var contentType = getContentTypeFromExt(ext);

      if (contentType) {
        blob = blob.setContentType(contentType);
      }

      var fileName = fileNameBase + "." + ext;

      var file = folder.createFile(blob.setName(fileName));

      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {
        console.error("File uploaded but sharing failed row " + sheetRow + ": " + shareErr.message);
      }

      var driveUrl = "https://drive.google.com/file/d/" + file.getId() + "/preview";
      setDriveLinkInCell(sheetRow, colNum, linkText, driveUrl);

      console.log("Uploaded and linked Drive file: " + fileName);

      Utilities.sleep(300);

    } catch (err) {
      console.error("Image fetch/upload error row " + sheetRow + ": " + err.message);
    }
  }
  // Sheet row 2 = data[0], row 3 = data[1]
  data.forEach(function(item, index) {
    var sheetRow = index + 2;

    uploadOneReimbursementFile(item, sheetRow, billImgIdx, "bill");
    uploadOneReimbursementFile(item, sheetRow, paymentImgIdx, "payment");
  });
  return data;
}

// ════════════════════════════════COMPENSATION══════════════════════════════════════════════════════
function compensation(ACCESS_TOKEN, targetSheetName, selMonth, selYear) {
  var hdrs = { Authorization: "Zoho-oauthtoken " + ACCESS_TOKEN };
  var allData = [];
  var startIndex = 0;
  var batchSize = 15;
  while (true) {
    var url = "https://people.zoho.in/api/forms/compensation_for_additional_workdays_view/records?startIndex=" + startIndex + "&limit=" + batchSize;
    var resp = UrlFetchApp.fetch(url, {
    method: "get",
    headers: hdrs,
    muteHttpExceptions: true
  });

  console.log("COMPENSATION API URL:");
  console.log(url);

  console.log("COMPENSATION API STATUS:");
  console.log(resp.getResponseCode());

  console.log("COMPENSATION API RAW RESPONSE:");
  console.log(resp.getContentText());

  if (resp.getResponseCode() === 429 || resp.getContentText().indexOf("Bandwidth quota exceeded") !== -1) {
    Utilities.sleep(7000);

    resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: hdrs,
      muteHttpExceptions: true
    });

    console.log("COMPENSATION API RETRY STATUS:");
    console.log(resp.getResponseCode());

    console.log("COMPENSATION API RETRY RAW RESPONSE:");
    console.log(resp.getContentText());
  }

  var batch = JSON.parse(resp.getContentText());

  console.log("COMPENSATION API PARSED JSON:");
  console.log(JSON.stringify(batch, null, 2));  
    if (!Array.isArray(batch) || batch.length === 0) break;
    allData = allData.concat(batch);
    if (batch.length < batchSize) break;
    startIndex += batchSize;
    Utilities.sleep(4000);
  }
  var data = allData;

  var sn = targetSheetName || "Compensation";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sn);
  if (!sheet) { sheet = ss.insertSheet(sn); } else { sheet.clearContents(); sheet.clearFormats(); }

  if (selMonth !== null && selMonth !== undefined && selYear !== null && selYear !== undefined) {
    var filterMonth = parseInt(selMonth);
    var filterYear  = parseInt(selYear);
    var monthNames  = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

    data = data.filter(function (item) {
      var rawDate =
        item["Worked Date"] ||
        item["Date of working"] ||
        item["Date of Working"] ||
        item["Date of payment"] ||
        item["Date of Payment"] ||
        "";

      if (!rawDate) return false;

      var parts = String(rawDate).split("-");
      if (parts.length < 3) return false;

      return monthNames.indexOf(parts[1]) === filterMonth &&
            parseInt(parts[2], 10) === filterYear;
    });
  }

  if (!Array.isArray(data) || data.length === 0) {
    sheet.clearContents();
    sheet.clearFormats();
    sheet.getRange(1, 1).setValue("No compensation data found for selected month/year").setFontWeight("bold");
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "⚠️ No compensation records found for selected month — sheet cleared.",
      "Compensation", 5
    );
    return [];
  }

  var removeCompensationColumns = ["modifiedTime","ownerID","Modified Time","recordid","Modified By","ownerName","createdTime","Added By"];
  var hRow = Object.keys(data[0]).filter(function(h) {
    return removeCompensationColumns.indexOf(h) === -1;
  });

  hRow.push("Overtime Days");
  // ── Second pass: build rows and track which rows are duplicates ───────
  var seenOvertime = {};
  var vals = [hRow];
  var duplicateRows = [];

  data.forEach(function (item) {
    var empId    = String(item["Employee ID"] || item["Employee HRM"] || "").trim();
   var workDate = String(item["Worked Date"] ||item["Date of working"] ||item["Date of Working"] ||item["Date of payment"] ||item["Date of Payment"] ||
    ""
  ).trim();
    var key      = empId + "|" + workDate;
    var overtimeDay = "";

    if (empId && workDate) {
      if (!seenOvertime[key]) {
        seenOvertime[key] = true;
        overtimeDay = 1;
      } else {
        overtimeDay = 0;
        // Only flag the second (and further) duplicate rows
        duplicateRows.push(vals.length + 1);
      }
    }

    vals.push(hRow.map(function (h) {
      if (h === "Overtime Days") return overtimeDay;
      return decodeHtmlEntities_(item[h] != null ? item[h] : "");
    }));
  });

  // ── Write all data to sheet ───────────────────────────────────────────
  sheet.getRange(1, 1, vals.length, vals[0].length).setValues(vals);

  // ── Highlight ALL duplicate rows in red ──────────────────────────────
  if (duplicateRows.length > 0) {
    var totalCols = vals[0].length;
    duplicateRows.forEach(function (rowNum) {
      sheet.getRange(rowNum, 1, 1, totalCols)
        .setBackground("#f4cccc")
        .setFontColor("#cc0000");
    });
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "⚠️ " + duplicateRows.length + " duplicate row(s) highlighted in red (same Employee + Date).",
      "Duplicates Found", 6
    );
  }

  // ── Formatting ────────────────────────────────────────────────────────
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setHorizontalAlignment("left").setVerticalAlignment("middle");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setWrap(true);
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setBorder(true, true, true, true, true, true);
  return data;
}


//************************************************
function highlightAttendanceAndPFMismatches() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var attSheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  var pfSheet  = ss.getSheetByName(CONFIG.SALARY_PF_SHEET);

  if (!attSheet || !pfSheet) {
    ss.toast(
      "Attendance or Salary and PF Details sheet not found.",
      "Employee Status Check",
      5
    );
    return;
  }

  var attLastRow = attSheet.getLastRow();
  var pfLastRow  = pfSheet.getLastRow();
  var pfLastCol  = pfSheet.getLastColumn();

  if (attLastRow < 2 || pfLastRow < 2) return;

  var red = "#f4cccc";          // Light Red 3
  var lightGreen3 = "#d9ead3";  // Light Green 3
  var lightPurple3 = "#d9d2e9"; // Light Purple 3
  var white = "#ffffff";

  function normalizeEmpId(value) {
    return String(value || "")
      .replace(/\u00A0/g, " ")
      .replace(/[\u200B-\u200D\uFEFF]/g, "")
      .replace(/[\r\n\t]/g, "")
      .trim()
      .toUpperCase()
      .replace(/\s+/g, "");
  }

  function addToMap(map, value) {
    var id = normalizeEmpId(value);
    if (!id) return;
    map[id] = true;
  }

  function existsInMap(map, value) {
    var id = normalizeEmpId(value);
    if (!id) return false;
    return map[id] === true;
  }

  // Attendance record-Zoho: Employee Id = column A, Employee Name = column B
  var attData = attSheet.getRange(2, 1, attLastRow - 1, 2).getValues();

  // Salary and PF Details full data
  var pfData = pfSheet.getDataRange().getValues();
  var pfHeaders = pfData[0];

  var pfEmpIdIdx = getColIndex(pfHeaders, [
    "Employee ID",
    "Employee Id",
    "Employee HRM",
    "Emp ID",
    "EmpID"
  ]);

  var pfExitDateIdx = getColIndex(pfHeaders, [
    "Exit Date",
    "Date of Exit",
    "Dateofexit",
    "Date Of Exit"
  ]);

  if (pfEmpIdIdx === -1) {
    throw new Error("Employee ID column not found in Salary and PF Details");
  }

  var attIds = {};
  var pfIds = {};

  // Build Attendance employee ID map
  for (var a = 0; a < attData.length; a++) {
    addToMap(attIds, attData[a][0]);
  }

  // Build Salary & PF employee ID map
  for (var p = 1; p < pfData.length; p++) {
    addToMap(pfIds, pfData[p][pfEmpIdIdx]);
  }

  // Clear old colors
  attSheet.getRange(2, 1, attLastRow - 1, 2).setBackground(white);
  pfSheet.getRange(2, 1, pfLastRow - 1, pfLastCol).setBackground(white);

  var missingInAttendanceWithExitCount = 0;
  var missingInAttendanceNoExitCount = 0;
  var presentInAttendanceCount = 0;

  var pfBgs = [];

  for (var pr = 1; pr < pfData.length; pr++) {
    var pfEmpId = pfData[pr][pfEmpIdIdx];
    var empMissingInAttendance =
      normalizeEmpId(pfEmpId) &&
      !existsInMap(attIds, pfEmpId);

    var rowColor = white;

    if (empMissingInAttendance) {
      var exitDateVal = "";

      if (pfExitDateIdx !== -1) {
        exitDateVal = pfData[pr][pfExitDateIdx];
      }

      var hasExitDate =
        exitDateVal !== "" &&
        exitDateVal !== null &&
        exitDateVal !== undefined;

      if (hasExitDate) {
        rowColor = red;
        missingInAttendanceWithExitCount++;
      } else {
        rowColor = lightGreen3;
        missingInAttendanceNoExitCount++;
      }
    } else {
      presentInAttendanceCount++;
    }

    var rowBg = [];
    for (var c = 0; c < pfLastCol; c++) {
      rowBg.push(rowColor);
    }

    pfBgs.push(rowBg);
  }

  if (pfBgs.length > 0) {
    pfSheet.getRange(2, 1, pfBgs.length, pfLastCol).setBackgrounds(pfBgs);
  }

  var missingInPFCount = 0;
  var attBgs = [];

  for (var ar = 0; ar < attData.length; ar++) {
    var attEmpId = attData[ar][0];

    if (normalizeEmpId(attEmpId) && !existsInMap(pfIds, attEmpId)) {
      attBgs.push([lightPurple3, lightPurple3]);
      missingInPFCount++;
    } else {
      attBgs.push([white, white]);
    }
  }

  if (attBgs.length > 0) {
    attSheet.getRange(2, 1, attBgs.length, 2).setBackgrounds(attBgs);
  }

  SpreadsheetApp.flush();
  ss.toast("Check done. PF missing in Attendance + Exit Date: ",8 );
}

function showMonthYearPicker() {
  var today = new Date();
  var html = HtmlService.createHtmlOutput(
    '<select id="m">' +
    ['January','February','March','April','May','June',
     'July','August','September','October','November','December']
    .map(function(n,i){
      var selected = (i === today.getMonth()) ? ' selected' : '';
      return '<option value="'+i+'"'+selected+'>'+n+'</option>';
    }).join('') +
    '</select>' +
    '<select id="y">' +
    [2023,2024,2025,2026].map(function(y){
      var selected = (y === today.getFullYear()) ? ' selected' : '';
      return '<option value="'+y+'"'+selected+'>'+y+'</option>';
    }).join('') +
    '</select>' +
    '<br><br><button onclick="save()">Save</button>' +
'<script>' +
'function save(){' +
'  var m = document.getElementById("m").value;' +
'  var y = document.getElementById("y").value;' +
'  google.script.run' +
'    .withSuccessHandler(function(){' +
'      google.script.host.close();' +
'    })' +
'    .saveMonthYear(m, y);' +
'}' +
'<\/script>'
  ).setWidth(300).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, "Select Month & Year");
}
//****************************************
function saveMonthYear(month, year) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("SELECTED_MONTH", String(month));
  props.setProperty("SELECTED_YEAR",  String(year));
  console.log("=== saveMonthYear → month: " + month + " year: " + year + " ===");
}

//********************************************************* 
function showInternSheetPicker() {
  var props = PropertiesService.getScriptProperties();
  var savedUrl = props.getProperty("INTERN_SHEET_URL") || "";
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body { font-family: Arial, sans-serif; padding: 16px; }' +
    'label { font-weight: bold; display: block; margin-bottom: 6px; }' +
    'input { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; }' +
    'button { margin-top: 12px; padding: 8px 20px; background: #4a86e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; }' +
    'button:hover { background: #3a76d8; }' +
    '.msg { margin-top: 10px; font-size: 12px; color: green; display: none; }' +
    '</style>' +
    '<label>Paste Intern Attendance Google Sheet URL:</label>' +
    '<input type="text" id="url" value="' + savedUrl + '" placeholder="https://docs.google.com/spreadsheets/d/..." />' +
    '<br>' +
    '<button onclick="save()">Save & Use This Sheet</button>' +
    '<div class="msg" id="msg">✔ Saved successfully!</div>' +
    '<script>' +
    'function save() {' +
    '  var url = document.getElementById("url").value.trim();' +
    '  if (!url) { alert("Please paste a valid Google Sheet URL."); return; }' +
    '  google.script.run' +
    '    .withSuccessHandler(function() {' +
    '      document.getElementById("msg").style.display = "block";' +
    '      setTimeout(function(){ google.script.host.close(); }, 1000);' +
    '    })' +
    '    .withFailureHandler(function(err) {' +
    '      alert("Error: " + err.message);' +
    '    })' +
    '    .saveInternSheetUrl(url);' +
    '}' +
    '<\/script>'
  ).setWidth(480).setHeight(160);
  
  SpreadsheetApp.getUi().showModalDialog(html, "Intern Sheet Selection");
}


function saveInternSheetUrl(url) {
  if (!url) throw new Error("URL cannot be empty.");
  PropertiesService.getScriptProperties().setProperty("INTERN_SHEET_URL", url.trim());
  console.log("Intern sheet URL saved: " + url.trim());
}

 // ═══════════════════════════════MASTER════════════════════════════════════════
function master() {
  var html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      margin: 0;
      padding: 14px;
      font-family: Arial, sans-serif;
      background: radial-gradient(circle, #eef6fb 0%, #d8e2ea 65%, #c8d1d9 100%);
      overflow: hidden;
    }

    .box {
      width: 340px;
      height: 105px;
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
    }

    .progress-wrap {
      position: relative;
      width: 280px;
      height: 16px;
      background: #b8c6cf;
      border-radius: 20px;
      box-shadow:
        inset 0 2px 5px rgba(0,0,0,0.25),
        0 1px 1px rgba(255,255,255,0.7);
      overflow: visible;
    }

    .progress-fill {
      height: 100%;
      width: 0%;
      border-radius: 20px;
      background:
        repeating-linear-gradient(
          -45deg,
          rgba(255,255,255,0.25) 0px,
          rgba(255,255,255,0.25) 10px,
          rgba(0,185,230,0.25) 10px,
          rgba(0,185,230,0.25) 20px
        ),
        linear-gradient(#62d9f7, #08a9dc);
      box-shadow:
        inset 0 2px 3px rgba(255,255,255,0.7),
        inset 0 -2px 3px rgba(0,0,0,0.15);
      transition: width 0.5s ease;
    }

    .bubble {
      position: absolute;
      top: -36px;
      left: 0%;
      transform: translateX(-50%);
      background: #4a4a4a;
      color: #ffffff;
      font-size: 12px;
      font-weight: bold;
      padding: 5px 8px;
      border-radius: 4px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.35);
      transition: left 0.5s ease;
    }

    .bubble:after {
      content: "";
      position: absolute;
      bottom: -9px;
      left: 50%;
      transform: translateX(-50%);
      border-width: 9px 7px 0 7px;
      border-style: solid;
      border-color: #4a4a4a transparent transparent transparent;
    }

    .status {
      margin-top: 16px;
      font-size: 11px;
      color: #333;
      text-align: center;
      font-weight: bold;
    }

    .done {
      color: green;
    }

    .error {
      color: #b00020;
    }
  </style>
</head>

<body>
  <div class="box">
    <div class="progress-wrap">
      <div id="fill" class="progress-fill"></div>
      <div id="bubble" class="bubble">0%</div>
    </div>
    <div id="status" class="status">Starting Employee Master...</div>
  </div>

  <script>
    const steps = [
      { name: "Saving current manual values...", fn: "employeeProgressSaveManualData", percent: 10 },
      { name: "Fetching attendance from Zoho...", fn: "employeeProgressAttendance", percent: 25 },
      { name: "Checking attendance and PF mismatches...", fn: "employeeProgressMismatch", percent: 35 },
      { name: "Fetching reimbursement data...", fn: "employeeProgressReimbursement", percent: 50 },
      { name: "Fetching compensation data...", fn: "employeeProgressCompensation", percent: 65 },
      { name: "Building salary sheet...", fn: "employeeProgressBuildSalary", percent: 80 },
      { name: "Building final employee sheet...", fn: "employeeProgressBuildFinalSS", percent: 95 },
      { name: "Saving final backup...", fn: "employeeProgressFinalSave", percent: 100 }
    ];

    function setProgress(percent, text, done, error) {
      document.getElementById("fill").style.width = percent + "%";
      document.getElementById("bubble").style.left = percent + "%";
      document.getElementById("bubble").innerText = percent + "%";

      const status = document.getElementById("status");
      status.innerText = text;
      status.className = "status" + (done ? " done" : "") + (error ? " error" : "");
    }

    function runStep(index) {
      if (index >= steps.length) {
        setProgress(100, "Employee Master completed successfully ✔", true, false);
        setTimeout(function() {
          google.script.host.close();
        }, 1500);
        return;
      }

      const step = steps[index];
      setProgress(step.percent, step.name, false, false);

      google.script.run
        .withSuccessHandler(function() {
          runStep(index + 1);
        })
        .withFailureHandler(function(err) {
          setProgress(step.percent, "Error: " + err.message, false, true);
        })[step.fn]();
    }

    runStep(0);
  </script>
</body>
</html>
  `).setWidth(390).setHeight(150);

SpreadsheetApp.getUi().showModalDialog(html, "Employee Master Progress");
}

function getSelectedMonthYearForEmployeeRun_() {
  var props = PropertiesService.getScriptProperties();
  var today = new Date();

  return {
    selMonth: props.getProperty("SELECTED_MONTH") || String(today.getMonth()),
    selYear: props.getProperty("SELECTED_YEAR") || String(today.getFullYear())
  };
}

function employeeProgressSaveManualData() {
  try {
    saveSalaryRemarksData();
  } catch (err) {
    console.error("Pre-save Salary Remarks failed: " + err.message);
  }

  try {
    saveFinalSSData();
  } catch (err) {
    console.error("Pre-save Final Employee SS failed: " + err.message);
  }
}

function employeeProgressAttendance() {
  var selected = getSelectedMonthYearForEmployeeRun_();
  console.log(
    "=== employeeProgressAttendance → selMonth: " +
    selected.selMonth +
    " selYear: " +
    selected.selYear +
    " ==="
  );

  var ACCESS_TOKEN = getAccessToken();
  updateSalaryPFDOJAndExitDate();
  SpreadsheetApp.flush();

  var attResult = attendance(ACCESS_TOKEN,CONFIG.ATTENDANCE_SHEET,selected.selMonth,selected.selYear);
  if (!attResult || attResult.rowsWritten === 0) {
    clearEmployeeSheetsWhenNoAttendance();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "⚠️ No employee attendance data found.",
      "Employee Master",
      6
    );

    throw new Error("No employee attendance data found for selected month/year.");
  }
}

function employeeProgressMismatch() {
  highlightAttendanceAndPFMismatches();
}

function employeeProgressReimbursement() {
  var selected = getSelectedMonthYearForEmployeeRun_();
  var ACCESS_TOKEN = getAccessToken();

  reimbursement(ACCESS_TOKEN,CONFIG.REIMBURSEMENT_SHEET,selected.selMonth,selected.selYear );
}

function employeeProgressCompensation() {
  var selected = getSelectedMonthYearForEmployeeRun_();
  var ACCESS_TOKEN = getAccessToken();

  compensation(ACCESS_TOKEN,CONFIG.COMPENSATION_SHEET,selected.selMonth,selected.selYear);
}

function employeeProgressBuildSalary() {
  buildSalarySheet(CONFIG.ATTENDANCE_SHEET, CONFIG.REIMBURSEMENT_SHEET,CONFIG.COMPENSATION_SHEET);
}

function employeeProgressBuildFinalSS() {
  buildFinalEmployeeSS(CONFIG.FINAL_EMPLOYEE_SS_SHEET);
}

function employeeProgressFinalSave() {
  saveSalaryRemarksData();
  saveFinalSSData();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "All employee sheets refreshed successfully ✔",
    "Done",
    5
  );
}


// ════════════════════════════════════════════════SHARED HELPERS════════════════════════════════════════════════
function getColIndex(headersRow, names) {
  var n=headersRow.map(function(h){return String(h).toLowerCase().replace(/\s+/g,"").replace(/[^a-z0-9]/g,"");});
  for (var i=0;i<names.length;i++){var x=n.indexOf(names[i].toLowerCase().replace(/\s+/g,"").replace(/[^a-z0-9]/g,""));if(x!==-1)return x;}
  return -1;
}
function getSheetByGid_(ss,gid){
  var sheets=ss.getSheets();
  for (var i=0;i<sheets.length;i++){if(String(sheets[i].getSheetId())===String(gid))return sheets[i];}
  return null;
}
function copyBorders_(src,dst){
  var ss=src.getSheet(),ds=dst.getSheet(),sr=src.getRow(),sc=src.getColumn(),nr=src.getNumRows(),nc=src.getNumColumns();
  for (var r=0;r<nr;r++) for (var c=0;c<nc;c++){var b=ss.getRange(sr+r,sc+c).getBorder();if(b)ds.getRange(dst.getRow()+r,dst.getColumn()+c).setBorder(b.top,b.left,b.bottom,b.right,b.vertical,b.horizontal);}
}

function getHrDataFolder_() {
  var parentFolderId = PropertiesService
    .getScriptProperties()
    .getProperty(CONFIG.PROP_BACKUP_PARENT_FOLDER_ID);

  if (!parentFolderId) {
    throw new Error(
      "Parent folder ID not found in Script Properties key: " +
      CONFIG.PROP_BACKUP_PARENT_FOLDER_ID
    );
  }

  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var hrDataFolders = parentFolder.getFoldersByName("Hr_Data");
  if (hrDataFolders.hasNext()) {
    return hrDataFolders.next();
  }
  return parentFolder.createFolder("Hr_Data");
}

//----------------------------------------------------
function normalizeInternCode_(value) {
  var s = String(value || "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/[\r\n\t]/g, "")
    .trim();

  if (s.indexOf(" - ") !== -1) {
    s = s.split(" - ")[0].trim();
  }

  if (/^\d+(\.0+)?$/.test(s)) {
    s = String(Math.round(Number(s)));
  }

  while (s.length > 0 && s.length < 4 && /^\d+$/.test(s)) {
    s = "0" + s;
  }
  return s.toUpperCase().replace(/\s+/g, "");
}

function decodeHtmlEntities_(value) {
  if (value === null || value === undefined) return "";
  var s = String(value);
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'");
}


