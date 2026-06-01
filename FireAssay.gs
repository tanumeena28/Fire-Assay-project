/**
 * Gold & Silver Assay Laboratory Management System (LMS)
 * Client: NCH Services Pvt. Ltd., Surat, Gujarat
 * Phase 4: Fire Assay Lab Data Entry Module (FireAssay.gs)
 *
 * This file manages the Fire Assay data entry, batch validation,
 * fineness calculations with CG correction, and updating job card status.
 */

/**
 * Menu Callback: Opens the "Fire Assay Data Entry" HTML modal dialog.
 */
function menuFireAssay() {
  try {
    var html;
    try {
      html = HtmlService.createHtmlOutputFromFile('FireAssay');
    } catch (err) {
      html = HtmlService.createHtmlOutputFromFile('fireAssay');
    }
    html.setWidth(980).setHeight(720);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Fire Assay Data Entry');
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("UI Error", "Could not open Fire Assay entry form: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Backend API: Returns all pending Job Cards that are ready for Fire Assay testing.
 * (Job cards with status = "Pending").
 * 
 * @return {Array<Object>} List of pending job cards with metadata.
 */
function getPendingJobCards() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var jcSheet = ss.getSheetByName("Job_Cards");
    if (!jcSheet) return [];
    
    var data = jcSheet.getDataRange().getValues();
    var list = [];
    
    for (var i = 1; i < data.length; i++) {
      var status = data[i][10].toString().trim(); // Column K: Status
      if (status === "Pending") {
        list.push({
          subJobNo: data[i][0].toString().trim(),     // Column A: Job Card No
          receiptNo: data[i][1].toString().trim(),    // Column B: Receipt No
          custName: data[i][3].toString().trim(),     // Column D: Customer Name
          material: data[i][4].toString().trim(),     // Column E: Material Type
          weight: parseFloat(data[i][6]) || 0,       // Column G: Weight Received (g)
          xrfAu: parseFloat(data[i][7]) || 0          // Column H: XRF Au Reading (ppt)
        });
      }
    }
    return list;
  } catch (e) {
    Logger.log("Error in getPendingJobCards: " + e.toString());
    return [];
  }
}

/**
 * Calculates Fineness (ppt) using the Calibration Gold (CG) blank correction.
 * Formula: ((finalWt - cgFinal) / (initialWt - cgInitial)) * 1000
 * 
 * @param {number} initialWt Initial weight of sample in mg.
 * @param {number} finalWt Final weight of sample in mg (cornet weight).
 * @param {number} cgInitial Initial weight of CG check gold in mg.
 * @param {number} cgFinal Final weight of CG check gold in mg.
 * @return {number} Calculated fineness in ppt.
 */
function calculateFineness(initialWt, finalWt, cgInitial, cgFinal) {
  var diffInitial = initialWt - cgInitial;
  if (diffInitial === 0) return 0;
  return ((finalWt - cgFinal) / diffInitial) * 1000;
}

/**
 * Validates the Calibration Gold (CG) fineness reading.
 * CG average fineness must be 1000.0 ppt ± 0.5 (999.5 to 1000.5) to be valid.
 * 
 * @param {Object} batchData Form data containing cgInitialWt1, cgFinalWt1, etc.
 * @return {Object} Validation result with status and message.
 */
function validateBatch(batchData) {
  var cgInitial1 = parseFloat(batchData.cgInitialWt1) || 0;
  var cgFinal1 = parseFloat(batchData.cgFinalWt1) || 0;
  var cgInitial2 = parseFloat(batchData.cgInitialWt2) || 0;
  var cgFinal2 = parseFloat(batchData.cgFinalWt2) || 0;
  
  if (cgInitial1 <= 0 || cgInitial2 <= 0) {
    return { valid: false, message: "Calibration Gold initial weights must be greater than 0." };
  }
  
  var cgFineness1 = (cgFinal1 / cgInitial1) * 1000;
  var cgFineness2 = (cgFinal2 / cgInitial2) * 1000;
  var cgFinenessAvg = (cgFineness1 + cgFineness2) / 2;
  var valid = (cgFinenessAvg >= 999.5 && cgFinenessAvg <= 1000.5);
  
  return {
    valid: valid,
    fineness: cgFinenessAvg,
    message: valid ? "Valid" : "Calibration Gold average fineness is outside acceptable range: " + cgFinenessAvg.toFixed(2) + " ppt (Must be 999.5–1000.5 ppt)"
  };
}

/**
 * Backend API: Saves the submitted batch of Fire Assay samples, performs calculations,
 * and updates corresponding Job Cards to "Completed".
 * 
 * @param {Object} formData Submitted values from modal form.
 * @return {string} Confirmation summary message.
 */
function saveFireAssayBatch(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var faSheet = ss.getSheetByName("Fire_Assay_Sheet");
    var jcSheet = ss.getSheetByName("Job_Cards");
    
    if (!faSheet || !jcSheet) {
      throw new Error("Fire_Assay_Sheet or Job_Cards sheet not found. Please run setup.");
    }
    
    // 1. Validate CG Fineness first
    var validation = validateBatch(formData);
    if (!validation.valid) {
      throw new Error(validation.message);
    }
    
    var cgInitial1 = parseFloat(formData.cgInitialWt1);
    var cgFinal1 = parseFloat(formData.cgFinalWt1);
    var cgInitial2 = parseFloat(formData.cgInitialWt2);
    var cgFinal2 = parseFloat(formData.cgFinalWt2);
    
    var cgFineness1 = (cgFinal1 / cgInitial1) * 1000;
    var cgFineness2 = (cgFinal2 / cgInitial2) * 1000;
    var cgFinenessAvg = (cgFineness1 + cgFineness2) / 2;
    var cgCorrectionFactor = 999.9 / cgFinenessAvg;
    
    // Format Batch Sheet Name as DDMMYYYY
    var batchDateVal = new Date(formData.batchDate);
    var batchSheetName = Utilities.formatDate(batchDateVal, Session.getScriptTimeZone(), "ddMMyyyy");
    var batchDateStr = Utilities.formatDate(batchDateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    var rowsToAppend = [];
    
    // 2. Prepare 2 Rows for CG Calibration (CG Check Gold = TRUE)
    var cgSampleWt = (cgInitial1 + cgInitial2) / 1000; // Total weight in grams (e.g. 0.270 g)
    
    rowsToAppend.push([
      batchSheetName,       // A: Batch Sheet Name
      "Check Gold",         // B: Sub-Job Number
      "cg",                 // C: Additional Info
      cgSampleWt,           // D: Received Weight (g)
      999.9,                // E: XRF Reading (ppt)
      formData.cgSerialNo1, // F: Serial No (Crucible)
      cgInitial1,           // G: Initial Weight mg
      cgFinal1,             // H: Final Weight mg
      999.9,                // I: Fineness ppt
      999.90,               // J: Mean Value
      0.0,                  // K: Return Weight (g)
      batchDateStr,         // L: Batch Date
      formData.testedBy,    // M: Tested By
      true,                 // N: CG Check Gold flag
      formData.cupelNo,     // O: Cupel No
      formData.sheetNo      // P: Sheet No
    ]);
    
    rowsToAppend.push([
      batchSheetName,       // A
      "Check Gold",         // B
      "cg",                 // C
      cgSampleWt,           // D
      999.9,                // E
      formData.cgSerialNo2, // F
      cgInitial2,           // G
      cgFinal2,             // H
      999.9,                // I
      999.90,               // J
      0.0,                  // K
      batchDateStr,         // L
      formData.testedBy,    // M
      true,                 // N
      formData.cupelNo,     // O
      formData.sheetNo      // P
    ]);
    
    var completedSubJobs = [];
    var samples = formData.samples || [];
    
    // 3. Prepare 2 Rows for each sample in the batch (CG = FALSE)
    for (var i = 0; i < samples.length; i++) {
      var s = samples[i];
      var subJobNo = s.subJobNo;
      var serialNo1 = s.serialNo1;
      var serialNo2 = s.serialNo2;
      var returnWeight = parseFloat(s.returnWeight) || 0;
      
      var initialWt1 = parseFloat(s.initialWt1) || 0;
      var finalWt1 = parseFloat(s.finalWt1) || 0;
      var initialWt2 = parseFloat(s.initialWt2) || 0;
      var finalWt2 = parseFloat(s.finalWt2) || 0;
      
      // Calculate corrected finenesses
      var rawF1 = (finalWt1 / initialWt1) * 1000;
      var rawF2 = (finalWt2 / initialWt2) * 1000;
      
      var fineness1 = Math.round(rawF1 * cgCorrectionFactor * 10) / 10;
      var fineness2 = Math.round(rawF2 * cgCorrectionFactor * 10) / 10;
      var meanFineness = (fineness1 + fineness2) / 2;
      
      completedSubJobs.push(subJobNo);
      
      // Row for Reading 1
      rowsToAppend.push([
        batchSheetName,
        subJobNo,
        (i + 1).toString(), // Add Info (sequential index e.g. "1")
        parseFloat(s.weight) || 0,
        parseFloat(s.xrfAu) || 0,
        serialNo1,
        initialWt1,
        finalWt1,
        fineness1,
        meanFineness,
        returnWeight,
        batchDateStr,
        formData.testedBy,
        false,
        formData.cupelNo,
        formData.sheetNo
      ]);
      
      // Row for Reading 2
      rowsToAppend.push([
        batchSheetName,
        subJobNo,
        (i + 1).toString(), // Add Info
        parseFloat(s.weight) || 0,
        parseFloat(s.xrfAu) || 0,
        serialNo2,
        initialWt2,
        finalWt2,
        fineness2,
        meanFineness,
        returnWeight,
        batchDateStr,
        formData.testedBy,
        false,
        formData.cupelNo,
        formData.sheetNo
      ]);
    }
    
    // 4. Save all rows in bulk to Fire_Assay_Sheet
    var startRow = faSheet.getLastRow() + 1;
    var range = faSheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length);
    
    // Text format for identifiers
    var formats = [];
    for (var r = 0; r < rowsToAppend.length; r++) {
      var formatRow = [];
      for (var c = 0; c < rowsToAppend[0].length; c++) {
        if (c === 0 || c === 1 || c === 5 || c === 14 || c === 15) {
          formatRow.push("@"); 
        } else {
          formatRow.push("");
        }
      }
      formats.push(formatRow);
    }
    
    range.setNumberFormats(formats);
    range.setValues(rowsToAppend);
    
    // 5. Update Job_Cards Status to "Completed"
    var jcData = jcSheet.getDataRange().getValues();
    for (var j = 0; j < completedSubJobs.length; j++) {
      var targetSubJob = completedSubJobs[j];
      for (var i = 1; i < jcData.length; i++) {
        if (jcData[i][0].toString().trim() === targetSubJob) {
          jcSheet.getRange(i + 1, 11).setValue("Completed"); // Column K: Status
          break;
        }
      }
    }
    
    // 6. Generate Printable A4 Sheet
    generatePrintableFireAssaySheet(batchSheetName);
    
    return "Successfully saved Fire Assay batch: " + batchSheetName + ".\n" + 
           "Sheet: " + formData.sheetNo + ", Cupel: " + formData.cupelNo + ".\n" +
           "Crucibles: " + samples.length + " samples (" + (samples.length * 2) + " assays) saved.\n" +
           "Calibration Gold Avg Fineness: " + cgFinenessAvg.toFixed(2) + " ppt.\n" +
           "Job cards updated. Printable sheet has been generated.";
           
  } catch (e) {
    throw new Error("Failed to save Fire Assay batch: " + e.toString());
  }
}

/**
 * Formats columns and headers in the 'Fire_Assay_Sheet' sheet.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss The active spreadsheet.
 */
function setupFireAssaySheet(ss) {
  var faSheet = ss.getSheetByName("Fire_Assay_Sheet");
  if (!faSheet) return;
  
  var lastRow = faSheet.getLastRow();
  if (lastRow > 0) {
    // If already initialized, make sure headers O and P are present
    var headersRange = faSheet.getRange(1, 1, 1, 16);
    var headers = headersRange.getValues()[0];
    if (headers[14] !== "Cupel No" || headers[15] !== "Sheet No") {
      faSheet.getRange(1, 15).setValue("Cupel No");
      faSheet.getRange(1, 16).setValue("Sheet No");
    }
    return;
  }
  
  var headers = [
    "Batch Sheet Name",    // A (date: DDMMYYYY)
    "Sub-Job Number",      // B
    "Additional Info",     // C (Reading 1 / Reading 2)
    "Received Weight (g)", // D
    "XRF Reading (ppt)",    // E
    "Serial No",           // F (crucible number within batch)
    "Initial Weight mg",    // G (before assay)
    "Final Weight mg",     // H (after assay, cornet weight)
    "Fineness ppt",        // I
    "Mean Value",          // J (average of 2 readings for same sub-job)
    "Return Weight (g)",   // K
    "Batch Date",          // L
    "Tested By",           // M
    "CG Check Gold",        // N (TRUE/FALSE flag)
    "Cupel No",            // O
    "Sheet No"             // P
  ];
  
  faSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatSheetCommon(faSheet, headers.length, "#F97316"); // Tab color: orange, headers styled via formatSheetCommon
}

/**
 * Compiles a beautiful, landscape, print-ready "Print_FA_[batchSheetName]" tab
 * for the Fire Assay batch data.
 * 
 * @param {string} batchSheetName The batch name (date format DDMMYYYY) to compile.
 */
function generatePrintableFireAssaySheet(batchSheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var faSheet = ss.getSheetByName("Fire_Assay_Sheet");
    if (!faSheet) throw new Error("Fire_Assay_Sheet not found.");
    
    var data = faSheet.getDataRange().getValues();
    var batchRows = [];
    
    // Read all rows matching this batch sheet name
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === batchSheetName) {
        batchRows.push(data[i]);
      }
    }
    
    if (batchRows.length === 0) {
      throw new Error("No data found for batch: " + batchSheetName);
    }
    
    // Extract metadata from the first row
    var batchDateVal = batchRows[0][11]; // L: Batch Date
    var testedBy = batchRows[0][12];    // M: Tested By
    var cupelNo = batchRows[0][14] || "";  // O: Cupel No
    var sheetNo = batchRows[0][15] || "";  // P: Sheet No
    
    var formatDatePretty = function(dateVal) {
      if (!dateVal) return "";
      try {
        var d = dateVal;
        if (!(d instanceof Date)) {
          d = new Date(dateVal.toString());
        }
        if (!isNaN(d.getTime())) {
          return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy");
        }
      } catch (e) {
        // Fallback to raw string split
      }
      
      var dateStr = dateVal.toString();
      var parts = dateStr.split("-");
      if (parts.length === 3) {
        var year = parts[0];
        var monthIdx = parseInt(parts[1], 10) - 1;
        var day = parseInt(parts[2], 10);
        var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        var monthStr = months[monthIdx] || "";
        var dayStr = day < 10 ? "0" + day : day.toString();
        return dayStr + " " + monthStr + " " + year;
      }
      return dateStr;
    };
    
    var prettyDate = formatDatePretty(batchDateVal);
    
    // Group rows into pairs by Sub-Job Number
    var groupedSamples = [];
    // Since rows are written in pairs, we can iterate step 2
    for (var j = 0; j < batchRows.length; j += 2) {
      var row1 = batchRows[j];
      var row2 = batchRows[j + 1];
      if (!row2) {
        // Fallback in case of odd number of rows
        row2 = row1;
      }
      
      var subJobNo = row1[1]; // B: Sub-Job Number
      var addInfo = row1[2];  // C: Additional Info
      var sampleWt = parseFloat(row1[3]) || 0; // D: Received Weight (g)
      var xrfReading = parseFloat(row1[4]) || 0; // E: XRF Reading (ppt)
      
      var sNo1 = row1[5];     // F: Serial No 1
      var init1 = parseFloat(row1[6]) || 0; // G: Initial Weight mg
      var fin1 = parseFloat(row1[7]) || 0;  // H: Final Weight mg
      var fineness1 = parseFloat(row1[8]) || 0; // I: Fineness ppt
      
      var sNo2 = row2[5];     // F: Serial No 2
      var init2 = parseFloat(row2[6]) || 0; // G: Initial Weight mg
      var fin2 = parseFloat(row2[7]) || 0;  // H: Final Weight mg
      var fineness2 = parseFloat(row2[8]) || 0; // I: Fineness ppt
      
      var meanValue = parseFloat(row1[9]) || 0; // J: Mean Value
      var returnWt = parseFloat(row1[10]) || 0; // K: Return Weight (g)
      
      // Calculate Cornet WT (g) = (Final Weight 1 + Final Weight 2) / 1000
      var cornetWt = (fin1 + fin2) / 1000;
      
      // Calculate Variation = Mean Value - XRF Reading
      var variation = meanValue - xrfReading;
      if (subJobNo === "Check Gold") {
        variation = 0; // Calibration gold variation is 0
      }
      
      groupedSamples.push({
        subJobNo: subJobNo,
        addInfo: addInfo,
        sampleWt: sampleWt,
        xrfReading: xrfReading,
        sNo1: sNo1,
        init1: init1,
        fin1: fin1,
        fineness1: fineness1,
        sNo2: sNo2,
        init2: init2,
        fin2: fin2,
        fineness2: fineness2,
        meanValue: meanValue,
        returnWt: returnWt,
        cornetWt: cornetWt,
        variation: variation
      });
    }
    
    // Create or Clear tab named Print_FA_[batchSheetName]
    var printTabName = "Print_FA_" + batchSheetName;
    var printTab = ss.getSheetByName(printTabName);
    if (printTab) {
      ss.deleteSheet(printTab);
    }
    printTab = ss.insertSheet(printTabName);
    printTab.setHiddenGridlines(true);
    
    // Columns A-L widths (fit A4 Landscape layout)
    var widths = [130, 60, 100, 100, 50, 90, 90, 80, 80, 100, 100, 80];
    for (var col = 1; col <= widths.length; col++) {
      printTab.setColumnWidth(col, widths[col - 1]);
    }
    
    // Styles
    var primaryColor = "#1E293B"; // Slate-800
    var grayBorderColor = "#94A3B8"; // Slate-400 (medium gray for clear print borders)
    
    // 1. BRAND HEADER
    printTab.getRange("A2:L2").merge().setValue("FIRE ASSAY SHEET AS PER IS:1418")
      .setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center").setFontColor(primaryColor);
    
    // 2. METADATA BOXES (Row 4 & 5)
    // Cupel No
    printTab.getRange("A4").setValue("Cupel").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(10).setFontColor(primaryColor);
    printTab.getRange("A5").setValue(cupelNo).setFontWeight("bold").setHorizontalAlignment("center").setFontSize(11)
      .setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
    // Sheet Number
    printTab.getRange("D4").setValue("Sheet Number").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(10).setFontColor(primaryColor);
    printTab.getRange("D5").setValue(sheetNo).setFontWeight("bold").setHorizontalAlignment("center").setFontSize(11)
      .setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
    // Date
    printTab.getRange("K4:L4").merge().setValue("Date").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(10).setFontColor(primaryColor);
    printTab.getRange("K5:L5").merge().setValue(prettyDate).setFontWeight("bold").setHorizontalAlignment("center").setFontSize(11)
      .setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
    // 3. TABLE HEADERS (Row 7)
    var headers = [
      "Sub Job.No",
      "Add Info",
      "Sample WT\n(in grams)",
      "XRF Reading\n(in ppt)",
      "S.No",
      "Initial S.WT",
      "Final S.WT",
      "Fineness\n(in ppt)",
      "Mean Value\n(in ppt)",
      "Return WT\n(in grams)",
      "Cornet WT\n(in grams)",
      "Variation"
    ];
    
    var headerRange = printTab.getRange(7, 1, 1, headers.length);
    headerRange.setValues([headers]).setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBackground("#F8FAFC").setFontColor(primaryColor);
    
    printTab.setRowHeight(7, 32);
    
    // Write Data Rows
    var startRow = 8;
    for (var k = 0; k < groupedSamples.length; k++) {
      var s = groupedSamples[k];
      
      // Write Row 1
      printTab.getRange(startRow, 1).setValue(s.subJobNo);
      printTab.getRange(startRow, 2).setValue(s.addInfo);
      printTab.getRange(startRow, 3).setValue(s.sampleWt).setNumberFormat("0.000");
      printTab.getRange(startRow, 4).setValue(s.xrfReading).setNumberFormat("0.0");
      printTab.getRange(startRow, 5).setValue(s.sNo1).setNumberFormat("@");
      printTab.getRange(startRow, 6).setValue(s.init1).setNumberFormat("0.000");
      printTab.getRange(startRow, 7).setValue(s.fin1).setNumberFormat("0.000");
      printTab.getRange(startRow, 8).setValue(s.fineness1).setNumberFormat("0.0");
      printTab.getRange(startRow, 9).setValue(s.meanValue).setNumberFormat("0.00");
      printTab.getRange(startRow, 10).setValue(s.returnWt).setNumberFormat("0.000");
      printTab.getRange(startRow, 11).setValue(s.cornetWt).setNumberFormat("0.000");
      printTab.getRange(startRow, 12).setValue(s.variation).setNumberFormat("+0.000;-0.000;0.000");
      
      // Write Row 2
      printTab.getRange(startRow + 1, 5).setValue(s.sNo2).setNumberFormat("@");
      printTab.getRange(startRow + 1, 6).setValue(s.init2).setNumberFormat("0.000");
      printTab.getRange(startRow + 1, 7).setValue(s.fin2).setNumberFormat("0.000");
      printTab.getRange(startRow + 1, 8).setValue(s.fineness2).setNumberFormat("0.0");
      
      // Merge cells across row 1 & row 2
      printTab.getRange(startRow, 1, 2, 1).merge(); // Sub Job No
      printTab.getRange(startRow, 2, 2, 1).merge(); // Add Info
      printTab.getRange(startRow, 3, 2, 1).merge(); // Sample WT
      printTab.getRange(startRow, 4, 2, 1).merge(); // XRF Reading
      printTab.getRange(startRow, 9, 2, 1).merge(); // Mean Value
      printTab.getRange(startRow, 10, 2, 1).merge(); // Return WT
      printTab.getRange(startRow, 11, 2, 1).merge(); // Cornet WT
      printTab.getRange(startRow, 12, 2, 1).merge(); // Variation
      
      printTab.setRowHeight(startRow, 24);
      printTab.setRowHeight(startRow + 1, 24);
      
      startRow += 2;
    }
    
    // Style the Table Range
    var endRow = startRow - 1;
    var tableRange = printTab.getRange(7, 1, endRow - 6, 12);
    tableRange.setFontSize(9).setVerticalAlignment("middle");
    tableRange.setBorder(true, true, true, true, true, true, grayBorderColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // Set alignment overrides on tableRange
    // Left align Column A (Sub Job No)
    printTab.getRange(8, 1, endRow - 7, 1).setHorizontalAlignment("left");
    // Center align Column B & E (Add Info, S.No)
    printTab.getRange(8, 2, endRow - 7, 1).setHorizontalAlignment("center");
    printTab.getRange(8, 5, endRow - 7, 1).setHorizontalAlignment("center");
    // Right align numeric weights and fineness
    printTab.getRange(8, 3, endRow - 7, 2).setHorizontalAlignment("right"); // Sample WT, XRF Reading
    printTab.getRange(8, 6, endRow - 7, 3).setHorizontalAlignment("right"); // Initial S.WT, Final S.WT, Fineness
    printTab.getRange(8, 9, endRow - 7, 4).setHorizontalAlignment("right"); // Mean, Return, Cornet, Variation
    
    // Medium outer border
    printTab.getRange(7, 1, endRow - 6, 12).setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    // Signatures
    var footerRow = endRow + 3;
    printTab.setRowHeight(footerRow, 40); // Spacer for actual signature
    
    // Name Row
    printTab.getRange(footerRow + 1, 1, 1, 3).merge().setValue(testedBy)
      .setFontSize(9).setHorizontalAlignment("left").setFontColor(primaryColor);
    printTab.getRange(footerRow + 1, 10, 1, 3).merge().setValue("Ketan Varlekar")
      .setFontSize(9).setHorizontalAlignment("right").setFontColor(primaryColor);
      
    // Label Row
    printTab.getRange(footerRow + 2, 1, 1, 3).merge().setValue("Tested By")
      .setFontWeight("bold").setFontSize(9).setHorizontalAlignment("left").setFontColor(primaryColor);
    printTab.getRange(footerRow + 2, 10, 1, 3).merge().setValue("Checked and Authorized By")
      .setFontWeight("bold").setFontSize(9).setHorizontalAlignment("right").setFontColor(primaryColor);
    
    // Shift focus
    ss.setActiveSheet(printTab);
    
  } catch (e) {
    Logger.log("Error in generatePrintableFireAssaySheet: " + e.toString());
  }
}
