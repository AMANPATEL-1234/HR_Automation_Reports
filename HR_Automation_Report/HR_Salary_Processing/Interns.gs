// ═══════════════════════════════════COPY INTERN ATTENDANCE════════════════════════════════════════
function copyInternAttendanceGoogleSheet() {
 var internUrl = PropertiesService.getScriptProperties().getProperty("INTERN_SHEET_URL");
 if (!internUrl || internUrl.trim() === "") {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "⚠️ No Intern Sheet URL found! Go to HR All Automation → Intern Sheet Selection and paste the link first.",
    "❌ Intern Sheet Missing",
    15
  );
  return;
}

var gidMatch = internUrl.match(/[#&]gid=(\d+)/);
var gid = gidMatch ? gidMatch[1] : "0";
var cleanUrl = internUrl.split("?")[0].split("#")[0];
var srcSheet = getSheetByGid_(SpreadsheetApp.openByUrl(cleanUrl), gid);
  if (!srcSheet) throw new Error("Source intern sheet not found.");
  var srcRange=srcSheet.getDataRange(),values=srcRange.getValues();
  if (!values||values.length===0) throw new Error("No data in intern source sheet.");
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var dest=ss.getSheetByName(CONFIG.INTERN_ATTENDANCE_SHEET)||ss.insertSheet(CONFIG.INTERN_ATTENDANCE_SHEET);
  dest.clearContents();dest.clearFormats();
  var nr=values.length,nc=values[0].length,dr=dest.getRange(1,1,nr,nc);
  dr.setValues(values);dr.setBackgrounds(srcRange.getBackgrounds());dr.setFontColors(srcRange.getFontColors());
  dr.setFontWeights(srcRange.getFontWeights());dr.setFontSizes(srcRange.getFontSizes());
  dr.setFontFamilies(srcRange.getFontFamilies());dr.setHorizontalAlignments(srcRange.getHorizontalAlignments());
  dr.setVerticalAlignments(srcRange.getVerticalAlignments());dr.setWrapStrategies(srcRange.getWrapStrategies());
  dr.setNumberFormats(srcRange.getNumberFormats());
  for (var c=1;c<=nc;c++) dest.setColumnWidth(c,srcSheet.getColumnWidth(c));
  for (var r=1;r<=nr;r++) dest.setRowHeight(r,srcSheet.getRowHeight(r));
  dest.setFrozenRows(srcSheet.getFrozenRows());dest.setFrozenColumns(srcSheet.getFrozenColumns());
}

//***************************************************
function refillInternAttendanceDOJFromInternsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dest = ss.getSheetByName(CONFIG.INTERN_ATTENDANCE_SHEET);
  var internsData = ss.getSheetByName(CONFIG.INTERN_DATA_SHEET);
  if (!dest || !internsData) return;
  if (dest.getLastRow() < 2 || internsData.getLastRow() < 2) return;
  var idVals = internsData.getDataRange().getValues();
  var idHdrs = idVals[0];

  var dojCol = -1;
  var codeCol = -1;

  for (var h = 0; h < idHdrs.length; h++) {
    var hn = String(idHdrs[h]).trim().toLowerCase();
    if (hn === "doj") dojCol = h;
    if (hn === "intern code") codeCol = h;
  }

  if (dojCol === -1 || codeCol === -1) return;

  var dojMap = {};

  for (var i = 1; i < idVals.length; i++) {
    var code = normalizeInternCode_(idVals[i][codeCol]);
    if (!code) continue;
    dojMap[code] = idVals[i][dojCol] || "";
  }

  var headers = dest.getRange(1, 1, 1, dest.getLastColumn()).getValues()[0];
  var dojColNum = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim().toLowerCase() === "doj") {
      dojColNum = c + 1;
      break;
    }
  }

  if (dojColNum === -1) return;
  dest.getRange(2, dojColNum, dest.getLastRow() - 1, 1).setNumberFormat("@");

  for (var r = 2; r <= dest.getLastRow(); r++) {
    var colAVal = String(dest.getRange(r, 1).getValue() || "").trim();
    var internCode = normalizeInternCode_(colAVal);
    var dojVal = dojMap[internCode] || "";
    var formattedDoj = "";

    if (dojVal && dojVal instanceof Date && !isNaN(dojVal.getTime())) {
      formattedDoj = Utilities.formatDate(
        dojVal,
        Session.getScriptTimeZone(),
        "d MMM yyyy"
      );
    } else if (dojVal) {
      var parsedDoj = new Date(dojVal);

      if (!isNaN(parsedDoj.getTime())) {
        formattedDoj = Utilities.formatDate(
          parsedDoj,
          Session.getScriptTimeZone(),
          "d MMM yyyy"
        );
      } else {
        formattedDoj = String(dojVal);
      }
    }

    dest.getRange(r, dojColNum).setValue(String(formattedDoj));
  }
  SpreadsheetApp.flush();
}

// ════════════════════════════════════COPY SELECTED COLUMNS FROM INTERN SHEET════════════════════════════════════
function copySelectedColumnsFromInternSheet() {
  var internDataUrl = PropertiesService
  .getScriptProperties()
  .getProperty("INTERN_DATA_SOURCE_URL");
  if (!internDataUrl) {
    throw new Error("Missing Script Property: INTERN_DATA_SOURCE_URL");
  }

  var gidMatch = internDataUrl.match(/[?#&]gid=(\d+)/);
  var gid = gidMatch ? gidMatch[1] : "0";
  var cleanUrl = internDataUrl.split("?")[0].split("#")[0];
  var srcSheet = getSheetByGid_(
    SpreadsheetApp.openByUrl(cleanUrl),
    gid
  );

  if (!srcSheet) {
    throw new Error("Intern data source tab not found for gid: " + gid);
  }
  if (!srcSheet) throw new Error("Intern data source sheet not found.");
  var rh = [
  "Entity",
  "Name",
  "Months since start",
  "DOJ",
  "Date of Ending",
  "Intern Code",
  "Stipend",
  "Designation",
  "Hometown- Location",
  "Status/Remarks"
];
  var srcRange=srcSheet.getDataRange(),data=srcRange.getValues(),bg=srcRange.getBackgrounds();
  if (!data||data.length===0) throw new Error("No data in intern source.");
  function normHeader_(v) {
  return String(v || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[^a-z0-9]/g, "");
}
  var hMap = {};
  data[0].forEach(function(h, i) {
    hMap[normHeader_(h)] = i;
  });

  function col_(names) {
    for (var i = 0; i < names.length; i++) {
      var key = normHeader_(names[i]);
      if (hMap[key] !== undefined) return hMap[key];
    }
    return -1;
  }
  // Debug: print raw Intern Code values from source
  for (var dr=1; dr<data.length; dr++) {
  var rawCode = data[dr][col_(["Intern Code"])];
  console.log("Row "+dr+" → Intern Code raw: ["+rawCode+"] type: "+typeof rawCode);
}
  var headerAliases = {
  "Entity": ["Entity"],
  "Name": ["Name"],
  "Months since start": [], // this will be calculated by formula, not copied
  "DOJ": ["DOJ"],
  "Date of Ending": ["Date of Ending", "Date of Ending "],
  "Intern Code": ["Intern Code"],
  "Stipend": ["Stipend"],
  "Designation": ["Designation"],
  "Hometown- Location": ["Hometown- Location", "Hometown - Location", "Hometown-Location"],
  "Status/Remarks": ["Status/Remarks", "Status Remarks"]
};

var miss = rh.filter(function(h) {
  if (h === "Months since start") return false; // formula column, not source column
  return col_(headerAliases[h] || [h]) === -1;
});
if (miss.length) throw new Error("Missing headers: " + miss.join(", "));
  //***********************************
  // Detect attendance month/year from intern attendance sheet headers
  var attSheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.INTERN_ATTENDANCE_SHEET);
  var today      = new Date();
  var calcMonth  = today.getMonth();
  var calcYear   = today.getFullYear();
  if (attSheet) {
    var attHdrs = attSheet.getRange(1, 1, 1, attSheet.getLastColumn()).getValues()[0];
    for (var h = 1; h < attHdrs.length; h++) {
      var hdr = attHdrs[h];
      if (hdr instanceof Date && !isNaN(hdr.getTime())) {
      calcMonth = hdr.getMonth();
      calcYear  = hdr.getFullYear();
      break;
    }
  }
}
// First day and last day of detected attendance month
var monthStart = new Date(calcYear, calcMonth, 1);
var monthEnd   = new Date(calcYear, calcMonth + 1, 0); // last day of month
console.log("Intern data filter → month: " + calcMonth + " year: " + calcYear);
var output = [rh];
for (var r = 1; r < data.length; r++) {
  var dojRaw = data[r][col_(["DOJ"])];
  var endRaw = data[r][col_(["Date of Ending", "Date of Ending "])];
  // Parse DOJ
  var doj = dojRaw ? new Date(dojRaw) : null;
  // Parse Date of Ending — if blank/empty treat as still active (use far future date)
  var doe = (endRaw && String(endRaw).trim() !== "") ? new Date(endRaw) : new Date(9999, 11, 31);
  // Skip if DOJ is invalid
  if (!doj || isNaN(doj.getTime())) continue;
  if (doj > monthEnd || doe < monthStart) continue;
  output.push(rh.map(function(h){
    if (h === "Months since start") {
      return ""; // formula will be added after writing
    }
    var idx = col_(headerAliases[h] || [h]);
    var v = idx === -1 ? "" : data[r][idx];
    if (h === "DOJ" && v) {
      v = new Date(v);
    }

    if (h === "Date of Ending" && v) {
      v = new Date(v);
    }

    if (h === "Intern Code") {
    v = normalizeInternCode_(v);
  }
    return v;
}));
}
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var dest=ss.getSheetByName(CONFIG.INTERN_DATA_SHEET)||ss.insertSheet(CONFIG.INTERN_DATA_SHEET);
  dest.clearContents();dest.clearFormats();

  // Sort Interns data by DOJ before writing
  var headerRow = output[0];
  var dataRows = output.slice(1);
  var dojIndex = headerRow.indexOf("DOJ");

  if (dojIndex !== -1) {
    dataRows.sort(function(a, b) {
      var da = a[dojIndex] ? new Date(a[dojIndex]) : new Date(9999, 11, 31);
      var db = b[dojIndex] ? new Date(b[dojIndex]) : new Date(9999, 11, 31);
      return da - db;
    });
  }

  output = [headerRow].concat(dataRows);
  var internCodeColNum = rh.indexOf("Intern Code") + 1;
  if (internCodeColNum > 0 && output.length > 1) {
    dest.getRange(2, internCodeColNum, output.length - 1, 1).setNumberFormat("@");
  }
  dest.getRange(1, 1, output.length, output[0].length).setValues(output);
  
   var dojColNum = rh.indexOf("DOJ") + 1;
   var endColNum = rh.indexOf("Date of Ending") + 1;

  if (output.length > 1) {
    if (dojColNum > 0) {
      dest.getRange(2, dojColNum, output.length - 1, 1).setNumberFormat("d mmm yyyy");
    }

  if (endColNum > 0) {
    dest.getRange(2, endColNum, output.length - 1, 1).setNumberFormat("d mmm yyyy");
  }
}
  // Add formula in Months since start column using DOJ column
  var monthsColNum = rh.indexOf("Months since start") + 1;
  var dojColNum = rh.indexOf("DOJ") + 1;
  function columnLetter_(colNum) {
    var letter = "";
    while (colNum > 0) {
      var rem = (colNum - 1) % 26;
      letter = String.fromCharCode(65 + rem) + letter;
      colNum = Math.floor((colNum - 1) / 26);
    }
    return letter;
  }
  if (monthsColNum > 0 && dojColNum > 0 && output.length > 1) {
    var dojLetter = columnLetter_(dojColNum);
    var formulas = [];
    for (var fr = 2; fr <= output.length; fr++) {
      formulas.push([
        '=IF(' + dojLetter + fr + '="","",ROUND((DATEDIF(' + dojLetter + fr + ',TODAY(),"M")+DATEDIF(' + dojLetter + fr + ',TODAY(),"MD")/31)*2,0)/2)'
      ]);
    }
    dest.getRange(2, monthsColNum, formulas.length, 1).setFormulas(formulas);
  }
  dest.getRange(1,1,1,output[0].length).setFontWeight("bold");
  dest.setFrozenRows(1);dest.autoResizeColumns(1,output[0].length);
  dest.getRange(1,1,output.length,output[0].length).setBorder(true,true,true,true,true,true);
  dest.getRange(1, 1, 1, output[0].length).setBackground("#d9e2f3");
  applyInternsDataHeaderHoverLinks();
}

//***********************************************
function applyInternsDataHeaderHoverLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.INTERN_DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 1) return;
  var internDetailUrl = PropertiesService
  .getScriptProperties()
  .getProperty("INTERN_DATA_SOURCE_URL");

  if (!internDetailUrl) {
    console.log("INTERN_DATA_SOURCE_URL not found, skipping header hover links.");
    return;
  }

  var targetHeaders = [
    "Entity",
    "Name",
    "Months since start",
    "DOJ",
    "Date of Ending",
    "Intern Code",
    "Stipend",
    "Designation",
    "Hometown- Location",
    "Status/Remarks"
  ];

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  function clean(v) {
    return String(v || "")
      .toLowerCase()
      .replace(/\s+/g, "")
      .replace(/[^a-z0-9]/g, "");
  }

  var blackHeaderStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor("#000000")
    .setBold(true)
    .setUnderline(false)
    .build();

  for (var c = 0; c < headers.length; c++) {
    var headerText = String(headers[c] || "");

    var shouldLink = targetHeaders.some(function(h) {
      return clean(h) === clean(headerText);
    });

    if (!shouldLink) continue;

    var richText = SpreadsheetApp.newRichTextValue()
      .setText(headerText)
      .setLinkUrl(internDetailUrl)
      .setTextStyle(blackHeaderStyle)
      .build();

    sheet.getRange(1, c + 1)
      .setRichTextValue(richText)
      .setFontColor("#000000")
      .setFontWeight("bold")
      .setFontLine("none");
  }
}
//***********************************************
function highlightInternCodeMismatches() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var attSheet   = ss.getSheetByName(CONFIG.INTERN_ATTENDANCE_SHEET);
  var dataSheet  = ss.getSheetByName(CONFIG.INTERN_DATA_SHEET);
  var finalSheet = ss.getSheetByName(CONFIG.FINAL_INTERNS_SHEET);

  if (!attSheet || !dataSheet || !finalSheet) {
    ss.toast("One or more intern sheets not found.", "Intern Code Check", 5);
    return;
  }

  var red = "#f4cccc";
  var white = "#ffffff";

  function normalizeCode(value) {
    var s = String(value || "")
      .replace(/\u00A0/g, " ")
      .replace(/[\u200B-\u200D\uFEFF]/g, "")
      .replace(/[\r\n\t]/g, "")
      .trim();

    // Attendance record-intern column A format: "2509 - Name"
    if (s.indexOf(" - ") !== -1) {
      s = s.split(" - ")[0].trim();
    }

    // Convert 2509.0 to 2509
    if (/^\d+(\.0+)?$/.test(s)) {
      s = String(Math.round(Number(s)));
    }

    // Pad numeric codes to 4 digits: 211 -> 0211
    while (s.length > 0 && s.length < 4 && /^\d+$/.test(s)) {
      s = "0" + s;
    }

    return s.toUpperCase().replace(/\s+/g, "");
  }

  function getColIndexLocal(headers, names) {
    function clean(v) {
      return String(v || "")
        .toLowerCase()
        .replace(/\s+/g, "")
        .replace(/[^a-z0-9]/g, "");
    }

    var cleanHeaders = headers.map(clean);

    for (var i = 0; i < names.length; i++) {
      var idx = cleanHeaders.indexOf(clean(names[i]));
      if (idx !== -1) return idx;
    }

    return -1;
  }

  function addToMap(map, value) {
    var code = normalizeCode(value);
    if (code) map[code] = true;
  }

  function exists(map, value) {
    var code = normalizeCode(value);
    return !!(code && map[code]);
  }

  var attLastRow = attSheet.getLastRow();
  var attLastCol = attSheet.getLastColumn();

  var dataLastRow = dataSheet.getLastRow();
  var dataLastCol = dataSheet.getLastColumn();

  var finalLastRow = finalSheet.getLastRow();
  var finalLastCol = finalSheet.getLastColumn();

  if (attLastRow < 2 || dataLastRow < 2 || finalLastRow < 2) {
    ss.toast("One or more intern sheets have no data rows.", "Intern Code Check", 5);
    return;
  }

  var attValues = attSheet
    .getRange(1, 1, attLastRow, attLastCol)
    .getDisplayValues();

  var dataValues = dataSheet
    .getRange(1, 1, dataLastRow, dataLastCol)
    .getDisplayValues();

  var finalValues = finalSheet
    .getRange(1, 1, finalLastRow, finalLastCol)
    .getDisplayValues();

  var dataCodeIdx  = getColIndexLocal(dataValues[0], ["Intern Code"]);
  var dataNameIdx  = getColIndexLocal(dataValues[0], ["Name"]);

  var finalCodeIdx = getColIndexLocal(finalValues[0], ["Intern Code"]);
  var finalNameIdx = getColIndexLocal(finalValues[0], ["Name"]);

  if (dataCodeIdx === -1) {
    throw new Error("Intern Code column not found in Interns data");
  }

  if (dataNameIdx === -1) {
    throw new Error("Name column not found in Interns data");
  }

  if (finalCodeIdx === -1) {
    throw new Error("Intern Code column not found in Final_Interns");
  }

  if (finalNameIdx === -1) {
    throw new Error("Name column not found in Final_Interns");
  }

  var attMap = {};
  var dataMap = {};
  var finalMap = {};

  // Attendance record-intern: code from column A
  for (var a = 1; a < attValues.length; a++) {
    addToMap(attMap, attValues[a][0]);
  }

  // Interns data: code from Intern Code column
  for (var d = 1; d < dataValues.length; d++) {
    addToMap(dataMap, dataValues[d][dataCodeIdx]);
  }

  // Final_Interns: code from Intern Code column
  for (var f = 1; f < finalValues.length; f++) {
    addToMap(finalMap, finalValues[f][finalCodeIdx]);
  }
  // Clear old mismatch red color from full rows first
  attSheet.getRange(2, 1, attLastRow - 1, 1).setBackground(white);
  dataSheet.getRange(2, dataNameIdx + 1, dataLastRow - 1, 1).setBackground(white);
  finalSheet.getRange(2, finalNameIdx + 1, finalLastRow - 1, 1).setBackground(white);
  var attMismatchCount = 0;
  var dataMismatchCount = 0;
  var finalMismatchCount = 0;
  
  var attNameBgs = [];
  for (var ar = 1; ar < attValues.length; ar++) {
    var attCode = attValues[ar][0];

    var attMissing =
      normalizeCode(attCode) &&
      (!exists(dataMap, attCode) || !exists(finalMap, attCode));

    if (attMissing) {
      attNameBgs.push([red]);
      attMismatchCount++;
    } else {
      attNameBgs.push([white]);
    }
  }

  if (attNameBgs.length > 0) {
    attSheet.getRange(2, 1, attNameBgs.length, 1).setBackgrounds(attNameBgs);
  }
  var dataNameBgs = [];
  for (var dr = 1; dr < dataValues.length; dr++) {
    var dataCode = dataValues[dr][dataCodeIdx];
    var dataMissing =
      normalizeCode(dataCode) &&
      (!exists(attMap, dataCode) || !exists(finalMap, dataCode));

    if (dataMissing) {
      dataNameBgs.push([red]);
      dataMismatchCount++;
    } else {
      dataNameBgs.push([white]);
    }
  }

  if (dataNameBgs.length > 0) {
    dataSheet
      .getRange(2, dataNameIdx + 1, dataNameBgs.length, 1)
      .setBackgrounds(dataNameBgs);
  }

  var finalNameBgs = [];
  for (var fr = 1; fr < finalValues.length; fr++) {
    var finalCode = finalValues[fr][finalCodeIdx];
    var finalMissing =
      normalizeCode(finalCode) &&
      (!exists(attMap, finalCode) || !exists(dataMap, finalCode));

    if (finalMissing) {
      finalNameBgs.push([red]);
      finalMismatchCount++;
    } else {
      finalNameBgs.push([white]);
    }
  }

  if (finalNameBgs.length > 0) {
    finalSheet
      .getRange(2, finalNameIdx + 1, finalNameBgs.length, 1)
      .setBackgrounds(finalNameBgs);
  }

  SpreadsheetApp.flush();
  ss.toast(
    "Intern mismatch done. Attendance: " + attMismatchCount +
    " | Interns data: " + dataMismatchCount +
    " | Final_Interns: " + finalMismatchCount,
    "Intern Code Check",
    8
  );

  console.log("Attendance record-intern codes: " + JSON.stringify(Object.keys(attMap)));
  console.log("Interns data codes: " + JSON.stringify(Object.keys(dataMap)));
  console.log("Final_Interns codes: " + JSON.stringify(Object.keys(finalMap)));
}

//******************************************* 
function buildFinalInterns() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var internData = ss.getSheetByName(CONFIG.INTERN_DATA_SHEET);
  var attSheet   = ss.getSheetByName(CONFIG.INTERN_ATTENDANCE_SHEET);
  if (!internData) throw new Error("Sheet 'Interns data' not found");
  if (!attSheet)   throw new Error("Sheet 'Attendance record-intern' not found");
  // ── Read Interns data sheet ──────────────────────────────────────────
  var idVals   = internData.getDataRange().getValues();
  var idHdrs   = idVals[0];
  function getCol(headers, names) {
    for (var i = 0; i < names.length; i++) {
      for (var j = 0; j < headers.length; j++) {
        if (String(headers[j]).trim().toLowerCase() === names[i].toLowerCase()) return j;
      }
    }
    return -1;
  }

  var colEntity   = getCol(idHdrs, ["Entity"]);
  var colName     = getCol(idHdrs, ["Name"]);
  var colCode     = getCol(idHdrs, ["Intern Code"]);
  var colStipend  = getCol(idHdrs, ["Stipend"]);   // used as Monthly Fixed

  // ── Build a lookup map: Name → {entity, code, stipend} ──────────────
  var internMap = {};
  for (var i = 1; i < idVals.length; i++) {
    var row  = idVals[i];
    var name = String(row[colName] || "").trim();
    if (!name) continue;
    var code = colCode >= 0 ? String(row[colCode] || "").trim() : "";
    var entry = {
      entity  : colEntity  >= 0 ? String(row[colEntity]  || "").trim() : "",
      code    : code,
      stipend : colStipend >= 0 ? row[colStipend] || ""                : ""
    };
    // Key by plain name
    internMap[name] = entry;
    // Also key by "CODE - Name" format so attendance sheet names match too
    if (code) internMap[code + " - " + name] = entry;
  }

  // ── Read Attendance record-intern sheet ─────────────────────────────
  var attVals = attSheet.getDataRange().getValues();
  var attHdrs = attVals[0];

  // Get selected month/year for Total Days calculation
  var props     = PropertiesService.getScriptProperties();
  var today     = new Date();
  var calcMonth = today.getMonth();
  var calcYear  = today.getFullYear();

// Headers in intern attendance sheet are actual Date objects e.g. 01/04/2026
// Scan from col 1 to find first Date object header
for (var h = 1; h < attHdrs.length; h++) {
  var hdr = attHdrs[h];
  if (hdr instanceof Date && !isNaN(hdr.getTime())) {
    calcMonth = hdr.getMonth();    // 0-based, April = 3
    calcYear  = hdr.getFullYear(); // 2026
    break;
  }
}

var totalDays = new Date(calcYear, calcMonth + 1, 0).getDate();
console.log("Intern sheet detected → month: " + calcMonth + " year: " + calcYear + " totalDays: " + totalDays);
  // ── Build Final_Interns sheet ────────────────────────
  var targetSheet = ss.getSheetByName(CONFIG.FINAL_INTERNS_SHEET);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(CONFIG.FINAL_INTERNS_SHEET);
  } else {
    targetSheet.clearContents();
    targetSheet.clearFormats();
  }

 var headers = [
  "Entity", "Name", "Designation", "DOJ", "Intern Code", "Monthly Fixed",
  "Total Days", "Total Payable Days","Paid Leave","WFH","Unpaid Leaves",
  "One Time", "Deduction", "Net Payable", "Remarks"
];
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ── Restore backup for manual columns ───────────────────────────────
  var internMY = getFinalInternsMonthYearKey();
  var backup = readFinalInternsBackupFromDrive_(internMY.year, internMY.month);
  var colTWD     = -1;
  var colPayable = -1;
  var colExtra1  = -1;
  var colExtra2  = -1;

  for (var h = 0; h < attHdrs.length; h++) {
    var hn = String(attHdrs[h]).trim();

    if (hn === "TWD") {
      colTWD = h;
    }

    if (hn === "Expected_Payable_Days") {
      colPayable = h;
    }

    if (hn === "Leave") {
      colExtra1 = h;
    }

    if (hn === "WFH") {
      colExtra2 = h;
    }
  }

  // ── Build lookup map keyed by INTERN CODE from Interns data ─────────
  // Rebuild internMap keyed by code only
  var codeMap = {};
  for (var i = 1; i < idVals.length; i++) {
    var idRow = idVals[i];
    var code = colCode >= 0 ? idRow[colCode] : "";
  if (code === "" || code === null || code === undefined) continue;
  if (typeof code === "number") {
  code = String(Math.round(code));
  while (code.length < 4) code = "0" + code;
} else {
  code = String(code).trim();
}

    // Log every entry being added to codeMap
    console.log("codeMap entry → code:[" + code + "] name:[" + (colName >= 0 ? idRow[colName] : "") + "] stipend:[" + (colStipend >= 0 ? idRow[colStipend] : "") + "]");

    var colDesignation = getCol(idHdrs, ["Designation"]);
    var colDOJ         = getCol(idHdrs, ["DOJ"]);
    codeMap[code] = {
    entity      : colEntity      >= 0 ? String(idRow[colEntity]      || "").trim() : "",
    name        : colName        >= 0 ? String(idRow[colName]        || "").trim() : "",
    stipend     : colStipend     >= 0 ? idRow[colStipend]            || ""         : "",
    designation : colDesignation >= 0 ? String(idRow[colDesignation] || "").trim() : "",
    doj         : colDOJ         >= 0 ? idRow[colDOJ]                || ""         : ""
};}
  // ── Build one row per intern from attendance sheet ───────────────────
var dataRows = [];
var dojFormatted = ""; 
var baseBackup = {};
for (var r = 1; r < attVals.length; r++) {
    var attRow = attVals[r];
    var colA   = String(attRow[0] || "").trim();  // e.g. "2509 - Aman Patel"
    if (!colA) continue;

    // Extract intern code and name from col A
    var internCode = "";
    var internName = colA;
    if (colA.indexOf(" - ") !== -1) {
  internCode = colA.split(" - ")[0].trim();
  internName = colA.split(" - ").slice(1).join(" - ").trim();
  // Pad intern code to 4 digits to match codeMap keys
  while (internCode.length < 4) internCode = "0" + internCode;
}

    // Lookup from codeMap using intern code
    var info = codeMap[internCode] || codeMap[String(parseInt(internCode, 10))] || {};
    // Use pre-calculated columns if available, else count manually
    var twd     = totalDays;
    var payable = colPayable >= 0 ? (attRow[colPayable] || 0) : 0;
    var extra1 = colExtra1 >= 0 ? attRow[colExtra1] : "";
    var extra2 = colExtra2 >= 0 ? attRow[colExtra2] : "";
    // Count unpaid leaves from date columns (col 1 up to TWD col)
    var unpaid     = 0;
    var dateEndCol = colTWD > 0 ? colTWD : attHdrs.length;
    for (var c = 1; c < dateEndCol; c++) {
      var cell = String(attRow[c] || "").trim().toUpperCase();
      if (cell === "-" || cell === "" || cell === "N/A") continue;
      if (cell === "LWP" || cell === "A") unpaid += 1;
      else if (cell.indexOf("0.5LWP") !== -1 || cell.indexOf("0.5A") !== -1) unpaid += 0.5;
    }
baseBackup[internCode] = {
  totalPayableDays: payable,
  paidLeave: extra1
};

// Use intern code as backup key for manual columns
var bk = backup[internCode] || {};
var expectedPayable = Number(payable || 0);
var expectedPaidLeave = Number(extra1 || 0);
var finalUnpaidLeaves = (bk.unpaidLeaves !== undefined && bk.unpaidLeaves !== null && bk.unpaidLeaves !== "")
  ? Number(bk.unpaidLeaves)
  : 0;
var finalPayable = expectedPayable - finalUnpaidLeaves;
if (finalPayable < 0) finalPayable = 0;
var finalPaidLeave = expectedPaidLeave - finalUnpaidLeaves;
if (finalPaidLeave < 0) finalPaidLeave = 0;
// Update backup also
bk.totalPayableSource = expectedPayable;
bk.paidLeaveSource = expectedPaidLeave;
bk.totalPayableDays = finalPayable;
bk.paidLeave = finalPaidLeave;
bk.unpaidLeaves = finalUnpaidLeaves;
  dojFormatted = "";
try {
  if (info.doj && String(info.doj).trim() !== "") {
    dojFormatted = Utilities.formatDate(new Date(info.doj), Session.getScriptTimeZone(), "d MMM yyyy");
  }
} catch(e) {
  dojFormatted = String(info.doj || "");
}
dataRows.push([
  info.entity       || "",         // Entity
  internName,                      // Name
  info.designation  || "",         // Designation
  dojFormatted,                    // DOJ
  internCode,                      // Intern Code
  info.stipend      || "",         // Monthly Fixed
  twd,                             // Total Days
  finalPayable,                    // Total Payable Days
  finalPaidLeave,                  // Paid Leave
  extra2,                          // WFH from Attendance record-intern
  finalUnpaidLeaves,               // Unpaid Leaves
  (bk.oneTime   !== undefined && bk.oneTime   !== null && bk.oneTime   !== "") ? bk.oneTime   : "",
  (bk.deduction !== undefined && bk.deduction !== null && bk.deduction !== "") ? bk.deduction : "",
   "",
  (bk.remarks   !== undefined && bk.remarks   !== null && bk.remarks   !== "") ? bk.remarks   : "",
]);
  }

  if (dataRows.length === 0) {
    targetSheet.getRange(2, 1).setValue("No intern attendance data found");
    return;
  }

    // Sort Final_Interns by DOJ before writing
    var dojIndex = headers.indexOf("DOJ"); // zero-based index
    dataRows.sort(function(a, b) {
      var da = a[dojIndex] ? new Date(a[dojIndex]) : new Date(9999, 11, 31);
      var db = b[dojIndex] ? new Date(b[dojIndex]) : new Date(9999, 11, 31);

      return da - db;
    });

    targetSheet.getRange(2, 5, dataRows.length, 1).setNumberFormat("@");
    targetSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    // Force DOJ column format in Final_Interns
    var finalDojCol = headers.indexOf("DOJ") + 1;

    if (finalDojCol > 0 && dataRows.length > 0) {
      targetSheet
        .getRange(2, finalDojCol, dataRows.length, 1)
        .setNumberFormat("d mmm yyyy");
    }

    var finalInternHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    var finalInternCodeCol = getColIndex(finalInternHeaders, ["Intern Code"]) + 1;
    var finalNetPayCol = getColIndex(finalInternHeaders, ["Net Payable"]) + 1;
    var finalInternHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    var finalNetPayCol = getColIndex(finalInternHeaders, ["Net Payable"]) + 1;
    if (finalNetPayCol > 0 && dataRows.length > 0) {
      var netPayFormulas = [];
      for (var nr = 0; nr < dataRows.length; nr++) {
        var rowNum = nr + 2;
        netPayFormulas.push([
          "=IFERROR(ROUND(F" + rowNum + "/G" + rowNum + "*H" + rowNum + "+L" + rowNum + "-M" + rowNum + ",2),0)"
        ]);
      }

  targetSheet
    .getRange(2, finalNetPayCol, dataRows.length, 1)
    .setFormulas(netPayFormulas)
    .setNumberFormat("0.00");
}
   
  // ────────────────── Formatting ─────────────────────
  targetSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d9e2f3")
    .setHorizontalAlignment("center");
  targetSheet.getRange(2, 1, dataRows.length, headers.length)
    .setVerticalAlignment("middle");
  targetSheet.getRange(2, 1, dataRows.length, 3)
    .setHorizontalAlignment("left");
  targetSheet.getRange(2, 4, dataRows.length, headers.length - 3)
    .setHorizontalAlignment("center");
  var netPayColForFormat = getColIndex(headers, ["Net Payable"]) + 1;
  if (netPayColForFormat > 0) {
    targetSheet
      .getRange(2, netPayColForFormat, dataRows.length, 1)
      .setNumberFormat("0.00");
  }
  targetSheet.setFrozenRows(1);
  targetSheet.autoResizeColumns(1, headers.length);
  targetSheet.getRange(1, 1, dataRows.length + 1, headers.length)
    .setBorder(true, true, true, true, true, true);

  highlightInvalidStipendPayableDays();
  SpreadsheetApp.getActiveSpreadsheet().toast(
  "Final_Interns built — " + dataRows.length + " interns ✔",
  "Done",
  4
  );
  writeFinalInternsBackupToDrive_(internMY.year, internMY.month, backup);
  setFinalInternsVisibleMonthYear_(targetSheet, internMY.year, internMY.month);
  return { rowsWritten: dataRows.length };
}
//*******************************************
function highlightInvalidStipendPayableDays() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.FINAL_INTERNS_SHEET); // Stipend Sheet
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var totalDaysCol = getColIndex(headers, ["Total Days"]) + 1;
  var payableCol   = getColIndex(headers, ["Total Payable Days"]) + 1;
  var unpaidCol    = getColIndex(headers, ["Unpaid Leaves"]) + 1;
  if (totalDaysCol <= 0 || payableCol <= 0 || unpaidCol <= 0) {
    console.log("Stipend validation skipped: required columns not found.");
    return;
  }

  var numRows = lastRow - 1;
  var totalDaysVals = sheet.getRange(2, totalDaysCol, numRows, 1).getValues();
  var payableVals   = sheet.getRange(2, payableCol, numRows, 1).getValues();
  var unpaidVals    = sheet.getRange(2, unpaidCol, numRows, 1).getValues();
  var bgs = [];
  for (var i = 0; i < numRows; i++) {
    var totalDays = Number(totalDaysVals[i][0] || 0);
    var payable   = Number(payableVals[i][0] || 0);
    var unpaid    = Number(unpaidVals[i][0] || 0);
    if ((payable + unpaid) > totalDays) {
      bgs.push(["#f4cccc"]); // red
    } else {
      bgs.push(["#ffffff"]); // normal
    }
  }

  sheet.getRange(2, payableCol, numRows, 1).setBackgrounds(bgs);
  SpreadsheetApp.flush();
}
//*******************************************
function handleFinalInternsEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.FINAL_INTERNS_SHEET) return;

  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();
  var startRow = e.range.getRow();
  var endRow = e.range.getLastRow();

  if (startRow === 1) return;
  if (startRow < 2) startRow = 2;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colInternCode = getColIndex(headers, ["Intern Code"]) + 1;
  var colPayable = getColIndex(headers, ["Total Payable Days"]) + 1;
  var colPaidLeave = getColIndex(headers, ["Paid Leave", "Leave"]) + 1;
  var colUnpaid = getColIndex(headers, ["Unpaid Leaves"]) + 1;
  var colOneTime = getColIndex(headers, ["One Time"]) + 1;
  var colDeduction = getColIndex(headers, ["Deduction"]) + 1;
  var colRemarks = getColIndex(headers, ["Remarks"]) + 1;
  var watchedCols = [
    colPayable,
    colPaidLeave,
    colUnpaid,
    colOneTime,
    colDeduction,
    colRemarks
  ].filter(function(c) {
    return c > 0;
  });

  var colsToProcess = [];
  for (var c = startCol; c <= endCol; c++) {
    if (watchedCols.indexOf(c) !== -1) {
      colsToProcess.push(c);
    }
  }

  if (colsToProcess.length === 0) return;

  var numRows = endRow - startRow + 1;
  var internCodes = sheet.getRange(startRow, colInternCode, numRows, 1).getValues();
  var colData = {};
  for (var ci = 0; ci < colsToProcess.length; ci++) {
    var col = colsToProcess[ci];

    colData[col] = {
      formulas: sheet.getRange(startRow, col, numRows, 1).getFormulas(),
      values: sheet.getRange(startRow, col, numRows, 1).getValues()
    };
  }

  var my = getFinalInternsVisibleMonthYear_(sheet);
  var year = my.year;
  var month = my.month;
  var backup = readFinalInternsBackupFromDrive_(year, month);

  for (var r = 0; r < numRows; r++) {
    var internCode = String(internCodes[r][0] || "").trim();
    if (!internCode) continue;

    if (!backup[internCode]) {
    backup[internCode] = {
    totalPayableDays: "",
    totalPayableSource: "",
    paidLeave: "",
    paidLeaveSource: "",
    unpaidLeaves: "",
    oneTime: "",
    deduction: "",
    netPay: "",
    netPayRow: "",
    remarks: ""
  };
    }

    for (var ci2 = 0; ci2 < colsToProcess.length; ci2++) {
      var colNum = colsToProcess[ci2];
      var formula = colData[colNum].formulas[r][0];
      var val = colData[colNum].values[r][0];
      var toSave = formula !== "" ? formula : valueOrBlankInterns_(val);

     if (colNum === colUnpaid) {
      var unpaidVal = Number(toSave || 0);
      var rowNum = startRow + r;
      var expectedPayable = Number(backup[internCode].totalPayableSource || 0);
      var expectedPaidLeave = Number(backup[internCode].paidLeaveSource || 0);
      if (!expectedPayable) {
        expectedPayable =
          Number(sheet.getRange(rowNum, colPayable).getValue() || 0) +
          Number(backup[internCode].unpaidLeaves || 0);
      }

      if (!expectedPaidLeave) {
        expectedPaidLeave =
          Number(sheet.getRange(rowNum, colPaidLeave).getValue() || 0) +
          Number(backup[internCode].unpaidLeaves || 0);
      }

      var newPayable = expectedPayable - unpaidVal;
      var newPaidLeave = expectedPaidLeave - unpaidVal;
      if (newPayable < 0) newPayable = 0;
      if (newPaidLeave < 0) newPaidLeave = 0;
      sheet.getRange(rowNum, colPayable).setValue(newPayable);
      sheet.getRange(rowNum, colPaidLeave).setValue(newPaidLeave);
      backup[internCode].totalPayableSource = expectedPayable;
      backup[internCode].paidLeaveSource = expectedPaidLeave;
      backup[internCode].totalPayableDays = newPayable;
      backup[internCode].paidLeave = newPaidLeave;
      backup[internCode].unpaidLeaves = unpaidVal;
    }
      else if (colNum === colPayable) {
        backup[internCode].totalPayableDays = toSave;
      }
      else if (colNum === colPaidLeave) {
        backup[internCode].paidLeave = toSave;
      }
      else if (colNum === colOneTime) {
        backup[internCode].oneTime = toSave;
      }
      else if (colNum === colDeduction) {
        backup[internCode].deduction = toSave;
      }
      else if (colNum === colRemarks) {
        backup[internCode].remarks = toSave;
      }
    }
  }

  writeFinalInternsBackupToDrive_(year, month, backup);
  highlightInvalidStipendPayableDays();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✔ Final_Interns saved to Drive JSON for " + year + "-" + pad2_(month + 1),
    "Auto-saved",
    2
  );
}

//*******************************************
function saveFinalInternsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.FINAL_INTERNS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;

  var data = sheet.getDataRange().getValues();
  var formulas = sheet.getDataRange().getFormulas();
  var headers = data[0];

  var codeIdx = getColIndex(headers, ["Intern Code"]);
  var totalPayableIdx = getColIndex(headers, ["Total Payable Days"]);
  var paidLeaveIdx = getColIndex(headers, ["Paid Leave", "Leave"]);
  var unpaidLeavesIdx = getColIndex(headers, ["Unpaid Leaves"]);
  var oneTimeIdx = getColIndex(headers, ["One Time"]);
  var deductionIdx = getColIndex(headers, ["Deduction"]);
  var remarksIdx = getColIndex(headers, ["Remarks"]);

  if (codeIdx === -1) return;

  var my = getFinalInternsVisibleMonthYear_(sheet);
  var year = my.year;
  var month = my.month;

  var backup = readFinalInternsBackupFromDrive_(year, month);

  for (var r = 1; r < data.length; r++) {
    var internCode = String(data[r][codeIdx] || "").trim();
    if (!internCode) continue;
    if (!backup[internCode]) {
    backup[internCode] = {
    totalPayableDays: "",
    totalPayableSource: "",
    paidLeave: "",
    paidLeaveSource: "",
    unpaidLeaves: "",
    oneTime: "",
    deduction: "",
    remarks: ""
  };
    }

    var totalPayableVal = totalPayableIdx >= 0 ? valueOrBlankInterns_(data[r][totalPayableIdx]) : "";
    var paidLeaveVal = paidLeaveIdx >= 0 ? valueOrBlankInterns_(data[r][paidLeaveIdx]) : "";
    var unpaidLeavesVal = unpaidLeavesIdx >= 0 ? valueOrBlankInterns_(data[r][unpaidLeavesIdx]) : "";
    var oneTimeVal = oneTimeIdx >= 0 ? valueOrBlankInterns_(data[r][oneTimeIdx]) : "";
    var deductionVal = deductionIdx >= 0 ? valueOrBlankInterns_(data[r][deductionIdx]) : "";
    var remarksVal = remarksIdx >= 0 ? valueOrBlankInterns_(data[r][remarksIdx]) : "";

    if (totalPayableVal !== "") {
    backup[internCode].totalPayableDays = totalPayableVal;
    if (
      backup[internCode].totalPayableSource === undefined ||
      backup[internCode].totalPayableSource === null ||
      backup[internCode].totalPayableSource === ""
    ) {
      backup[internCode].totalPayableSource = totalPayableVal;
    }
  }
    if (paidLeaveVal !== "") backup[internCode].paidLeave = paidLeaveVal;
    if (unpaidLeavesVal !== "") backup[internCode].unpaidLeaves = unpaidLeavesVal;
    if (oneTimeVal !== "") backup[internCode].oneTime = oneTimeVal;
    if (deductionVal !== "") backup[internCode].deduction = deductionVal;
    if (remarksVal !== "") backup[internCode].remarks = remarksVal;
      }

  var url = writeFinalInternsBackupToDrive_(year, month, backup);
  console.log("Final_Interns JSON backup saved for " +year + "-" + pad2_(month + 1) +" interns: " + Object.keys(backup).length + " file: "
  );
}

//******************************************* 
var FINAL_INTERNS_BACKUP_PARENT_FOLDER_ID = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_BACKUP_PARENT_FOLDER_ID);
var FINAL_INTERNS_BACKUP_SUBFOLDER_NAME = CONFIG.FINAL_INTERNS_BACKUP_SUBFOLDER_NAME;

function getFinalInternsBackupFolder_() {
  var hrDataFolder = getHrDataFolder_();
  var folders = hrDataFolder.getFoldersByName(FINAL_INTERNS_BACKUP_SUBFOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return hrDataFolder.createFolder(FINAL_INTERNS_BACKUP_SUBFOLDER_NAME);
}
// month is zero-based: Jan=0, Feb=1, Mar=2
function getFinalInternsBackupFileName_(year, month) {
  return "Final_Interns_" + year + "-" + pad2_(Number(month) + 1) + ".json";
}

//***************************************
function readFinalInternsBackupFromDrive_(year, month) {
  var folder = getFinalInternsBackupFolder_();
  var fileName = getFinalInternsBackupFileName_(year, month);
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

  if (parsed.interns) {
    return parsed.interns;
  }

  return parsed;
}

//*****************************************
function writeFinalInternsBackupToDrive_(year, month, internsBackup) {
  var folder = getFinalInternsBackupFolder_();
  var fileName = getFinalInternsBackupFileName_(year, month);
  var files = folder.getFilesByName(fileName);

  var monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  var payload = {
    year: Number(year),
    month: Number(month) + 1, // human month: April = 4
    monthIndex: Number(month), // script month: April = 3
    monthName: monthNames[Number(month)],
    updatedAt: new Date().toISOString(),
    interns: internsBackup || {}
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

function pad2_(n) {
  return String(n).padStart(2, "0");
}

function setFinalInternsVisibleMonthYear_(sheet, year, month) {
  if (!sheet) return;
  sheet.getRange(1, 1).setNote(JSON.stringify({
  year: Number(year),
  month: Number(month) + 1,      // human month: April = 4
  monthIndex: Number(month)      // script month: April = 3
}));
}

//********************************************
function getFinalInternsVisibleMonthYear_(sheet) {
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

  var my = getFinalInternsMonthYearKey();

  return {
    year: Number(my.year),
    month: Number(my.month)
  };
}

//*******************************************
function getFinalInternsMonthYearKey() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attSheet = ss.getSheetByName(CONFIG.INTERN_ATTENDANCE_SHEET);

  var today = new Date();
  var calcMonth = today.getMonth();
  var calcYear = today.getFullYear();

  if (attSheet && attSheet.getLastColumn() > 1) {
    var hdrs = attSheet.getRange(1, 1, 1, attSheet.getLastColumn()).getValues()[0];

    for (var h = 1; h < hdrs.length; h++) {
      var hdr = hdrs[h];

      if (hdr instanceof Date && !isNaN(hdr.getTime())) {
        calcMonth = hdr.getMonth();
        calcYear = hdr.getFullYear();
        break;
      }
    }
  }

  return {
    month: calcMonth,
    year: calcYear,
    key: calcYear + "_" + calcMonth
  };
}

function valueOrBlankInterns_(v) {
  if (v === null || v === undefined || v === "") {
    return "";
  }
  return v; // keeps 0 as 0
}

//*******************************************
function adjustFinalInternFormulaRow_(formula, oldRow, newRow) {
  if (!formula || String(formula).charAt(0) !== "=") return formula;
  formula = String(formula);
  newRow = Number(newRow);

  if (!newRow) return formula;
  return formula.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/gi, function(match, colPart, rowNum) {
    return colPart + newRow;
  });
}

//*********************************
function addDOJAndSortInternAttendance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var attSheet = ss.getSheetByName(CONFIG.INTERN_ATTENDANCE_SHEET);
  var internDataSheet = ss.getSheetByName(CONFIG.INTERN_DATA_SHEET);

  if (!attSheet) {
    throw new Error("Sheet 'Attendance record-intern' not found");
  }

  if (!internDataSheet) {
    throw new Error("Sheet 'Interns data' not found");
  }

  if (attSheet.getLastRow() < 2 || internDataSheet.getLastRow() < 2) {
    return;
  }

  var internData = internDataSheet.getDataRange().getValues();
  var internHeaders = internData[0];

  var dojIdx = getColIndex(internHeaders, ["DOJ"]);
  var codeIdx = getColIndex(internHeaders, ["Intern Code"]);

  if (dojIdx === -1) {
    throw new Error("DOJ column not found in Interns data");
  }

  if (codeIdx === -1) {
    throw new Error("Intern Code column not found in Interns data");
  }

  // Build map: Intern Code -> DOJ
  var dojMap = {};

  for (var i = 1; i < internData.length; i++) {
    var code = normalizeInternCode_(internData[i][codeIdx]);
    if (!code) continue;

    var doj = internData[i][dojIdx] || "";
    dojMap[code] = doj;
  }

  var lastRow = attSheet.getLastRow();
  var lastCol = attSheet.getLastColumn();

  // Check if DOJ column already exists
  var attHeaders = attSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var attDojCol = getColIndex(attHeaders, ["DOJ"]) + 1;

  // If DOJ does not exist, insert after column A
  if (attDojCol === 0) {
    attSheet.insertColumnAfter(1);
    attDojCol = 2;
    lastCol = attSheet.getLastColumn();

    attSheet
      .getRange(1, attDojCol)
      .setValue("DOJ")
      .setFontWeight("bold")
      .setBackground("#6dbb4b");
  }

  // Fill DOJ values
  var attValues = attSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var dojValues = [];

  for (var r = 0; r < attValues.length; r++) {
    var colAVal = attValues[r][0]; // Example: "2509 - Aman Patel"
    var internCode = normalizeInternCode_(colAVal);
    var dojVal = dojMap[internCode] || "";

    if (dojVal instanceof Date && !isNaN(dojVal.getTime())) {
      dojValues.push([dojVal]);
    } else if (dojVal) {
      var parsed = new Date(dojVal);
      if (!isNaN(parsed.getTime())) {
        dojValues.push([parsed]);
      } else {
        dojValues.push([String(dojVal)]);
      }
    } else {
      dojValues.push([""]);
    }
  }

  attSheet
    .getRange(2, attDojCol, dojValues.length, 1)
    .setValues(dojValues)
    .setNumberFormat("d mmm yyyy");

  // Sort full attendance rows by DOJ
  lastRow = attSheet.getLastRow();
  lastCol = attSheet.getLastColumn();

  var sortRange = attSheet.getRange(2, 1, lastRow - 1, lastCol);

  sortRange.sort([
    {
      column: attDojCol,
      ascending: true
    }
  ]);

  attSheet.autoResizeColumns(1, lastCol);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✔ DOJ added and Attendance record-intern sorted by DOJ",
    "Intern Attendance",
    4
  );
}

//*********************************
function internMaster() {
  copyInternAttendanceGoogleSheet();
  copySelectedColumnsFromInternSheet();
  addDOJAndSortInternAttendance();
  SpreadsheetApp.flush();
  buildFinalInterns();
  SpreadsheetApp.flush();
  highlightInternCodeMismatches();
  SpreadsheetApp.flush();
  saveFinalInternsData();
}














