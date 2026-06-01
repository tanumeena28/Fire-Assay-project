/**
 * Gold & Silver Assay Laboratory Management System (LMS)
 * Client: NCH Services Pvt. Ltd., Surat, Gujarat
 * Phase 5: Test Report Generation and Printing (Reports.gs)
 *
 * This file handles database setups, report generation logic, 
 * and querying completed fire assay batches to identify pending reports.
 */

/**
 * Menu Callback: Opens the "Report Dialog" HTML modal dialog.
 */
function menuGenerateReport() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('ReportDialog')
        .setWidth(650)
        .setHeight(500);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Report Desk');
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("UI Error", "Could not open Report Desk dialog: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Menu Callback: Aligned to open the same central Report Desk.
 */
function menuPrintReport() {
  menuGenerateReport();
}

/**
 * Backend API: Returns a list of all sub-jobs that have completed Fire Assay testing
 * and are ready to be compiled into Reports.
 * 
 * @return {Array<Object>} List of completed sub-jobs with mean fineness values.
 */
function getCompletedJobCards() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var faSheet = ss.getSheetByName("Fire_Assay_Sheet");
    var reportsSheet = ss.getSheetByName("Reports");
    if (!faSheet) return [];
    
    var faData = faSheet.getDataRange().getValues();
    var repData = reportsSheet ? reportsSheet.getDataRange().getValues() : [];
    
    // Track already reported sub-jobs to flag them in the UI
    var reportedMap = {};
    for (var i = 1; i < repData.length; i++) {
      var rNo = repData[i][0] ? repData[i][0].toString().trim() : "";
      if (rNo) {
        reportedMap[rNo] = repData[i][21] ? repData[i][21].toString().trim() : "";
      }
    }
    
    // Group Fire Assay rows by sub-job
    var completedMap = {};
    for (var i = 1; i < faData.length; i++) {
      var subJob = faData[i][1] ? faData[i][1].toString().trim() : "";
      var isCg = faData[i][13];
      
      if (subJob && !isCg && subJob !== "CG_CALIBRATION") {
        if (!completedMap[subJob]) {
          completedMap[subJob] = {
            subJobNo: subJob,
            meanVal: parseFloat(faData[i][9]) || 0, // Column J: Mean Value
            date: faData[i][11]                      // Column L: Batch Date
          };
        }
      }
    }
    
    var list = [];
    var keys = Object.keys(completedMap);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      var status = reportedMap[k] || "Pending"; // "Draft", "Final", or "Pending"
      
      var dateVal = completedMap[k].date;
      var dateStr = "";
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        dateStr = dateVal ? dateVal.toString() : "";
      }
      
      list.push({
        subJobNo: completedMap[k].subJobNo,
        meanVal: completedMap[k].meanVal,
        date: dateStr,
        status: status
      });
    }
    
    return list;
  } catch (e) {
    Logger.log("Error in getCompletedJobCards: " + e.toString());
    return [];
  }
}

/**
 * Generates a unique ULR Number (Unique Lab Reference) for accreditation purposes.
 * Format: ULR/YEAR/CLEAN_SUBJOB
 * 
 * @param {string} subJobcardNo Sub-jobcard identifier.
 * @return {string} Generated ULR number.
 */
function generateULRNumber(subJobcardNo) {
  var year = new Date().getFullYear();
  var cleanSubJob = subJobcardNo.replace(/[^A-Za-z0-9]/g, ""); 
  return "ULR/" + year + "/" + cleanSubJob;
}

/**
 * Backend API: Compiles a report record for a specific sub-job.
 * Pulls data from Fire_Assay, Receiving, and Job_Cards, and writes it to the Reports sheet.
 * 
 * @param {string} subJobcardNo The Sub-Jobcard Number to compile.
 * @return {string} Generated Report Number.
 */
function generateReport(subJobcardNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var faSheet = ss.getSheetByName("Fire_Assay_Sheet");
    var recSheet = ss.getSheetByName("Receiving");
    var jcSheet = ss.getSheetByName("Job_Cards");
    var repSheet = ss.getSheetByName("Reports");
    
    if (!faSheet || !recSheet || !jcSheet || !repSheet) {
      throw new Error("Required sheets (Fire_Assay_Sheet, Receiving, Job_Cards, Reports) not found.");
    }
    
    // Helper to safely format values
    var safeStr = function(val) { return val ? val.toString().trim() : ""; };
    
    // 1. Fetch Fire Assay readings for this sub-job (We expect 2 rows: duplicate 1 & duplicate 2)
    var faData = faSheet.getDataRange().getValues();
    var fineness1 = 0;
    var fineness2 = 0;
    var meanVal = 0;
    var dateOfTest = "";
    var testedBy = "";
    var finalWt1 = 0;
    var finalWt2 = 0;
    var returnSampleWt = 0;
    var matchCount = 0;
    
    for (var i = 1; i < faData.length; i++) {
      var faSubJob = safeStr(faData[i][1]);
      if (faSubJob === subJobcardNo) {
        meanVal = parseFloat(faData[i][9]) || 0;    // Column J
        returnSampleWt = parseFloat(faData[i][10]) || 0; // Column K
        testedBy = safeStr(faData[i][12]);      // Column M
        
        var dateVal = faData[i][11]; // Column L
        if (dateVal instanceof Date) {
          dateOfTest = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          dateOfTest = dateVal ? dateVal.toString() : "";
        }
        
        matchCount++;
        if (matchCount === 1) {
          fineness1 = parseFloat(faData[i][8]) || 0; // Column I
          finalWt1 = parseFloat(faData[i][7]) || 0;    // Column H: Final Weight mg
        } else if (matchCount === 2) {
          fineness2 = parseFloat(faData[i][8]) || 0; // Column I
          finalWt2 = parseFloat(faData[i][7]) || 0;    // Column H: Final Weight mg
        }
      }
    }
    
    // Return Cornet Weight (g) = sum of cornet weights in mg / 1000
    var returnCornetWt = (finalWt1 + finalWt2) / 1000;
    
    // 2. Fetch Customer details from Receiving
    var recData = recSheet.getDataRange().getValues();
    var baseReceiptNo = "";
    var customerName = "";
    var customerAddress = "";
    var receiptDateStr = "";
    var recdWt = 0;
    var material = "Gold";
    var certificateName = "";
    
    for (var i = 1; i < recData.length; i++) {
      var recSubJob = safeStr(recData[i][0]);
      if (recSubJob === subJobcardNo) {
        customerName = safeStr(recData[i][3]);    // Column D
        customerAddress = safeStr(recData[i][6]); // Column G
        recdWt = parseFloat(recData[i][9]) || 0;           // Column J
        material = safeStr(recData[i][7]);        // Column H
        baseReceiptNo = recData[i][14] ? safeStr(recData[i][14]) : ""; // Column O
        certificateName = recData[i][21] ? safeStr(recData[i][21]) : customerName; // Column V
        
        var rDateVal = recData[i][1]; // Column B
        if (rDateVal instanceof Date) {
          receiptDateStr = Utilities.formatDate(rDateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          receiptDateStr = rDateVal ? rDateVal.toString() : "";
        }
        break;
      }
    }
    
    // Description of sample
    var description = material + " Ornament / Sample";
    
    // 3. Format Date variables
    var today = new Date();
    var reportDateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // Generate ULR Number
    var ulrNo = generateULRNumber(subJobcardNo);
    
    // 4. Check if report already exists in Reports sheet (if yes, we update it; else, append)
    var repData = repSheet.getDataRange().getValues();
    var existingRowIndex = -1;
    for (var i = 1; i < repData.length; i++) {
      if (safeStr(repData[i][0]) === subJobcardNo) {
        existingRowIndex = i + 1; // 1-indexed row number
        break;
      }
    }
    
    // Reports Schema:
    // A: Report No, B: Report Date, C: Receipt No, D: Receipt Date, E: Date of Test, F: Customer Name,
    // G: Customer Address, H: Description, I: Received Weight, J: Sampling Method, K: Date of Sampling,
    // L: Return Cornet Wt, M: Return Sample Wt, N: Gold Fineness 1, O: Gold Fineness 2, P: Mean Value,
    // Q: Additional Info, R: Deviations, S: Unusual Features, T: Tested By, U: Reviewed By, V: Status, W: ULR
    var rowValues = [
      subJobcardNo,       // A
      reportDateStr,      // B
      baseReceiptNo,      // C
      receiptDateStr,     // D
      dateOfTest,         // E
      certificateName,    // F
      customerAddress,    // G
      description,        // H
      recdWt,             // I
      "Scraping",         // J: Sampling Method (Default)
      receiptDateStr,     // K: Date of Sampling (Same as Receipt Date)
      returnCornetWt,     // L
      returnSampleWt,     // M
      fineness1,          // N
      fineness2,          // O
      meanVal,            // P
      "",                 // Q: Additional Info
      "No",               // R: Deviations (Default)
      "No",               // S: Unusual Features (Default)
      testedBy,           // T
      "Technical Manager",// U: Reviewed By (Default)
      "Draft",            // V: Status (Default Draft)
      ulrNo               // W: ULR Number
    ];
    
    if (existingRowIndex !== -1) {
      // Update existing row
      repSheet.getRange(existingRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      // Append new row
      var startRow = repSheet.getLastRow() + 1;
      var range = repSheet.getRange(startRow, 1, 1, rowValues.length);
      
      // Formatting
      var formats = [];
      var formatRow = [];
      for (var c = 0; c < rowValues.length; c++) {
        if (c === 0 || c === 2 || c === 22) {
          formatRow.push("@"); 
        } else {
          formatRow.push("");
        }
      }
      formats.push(formatRow);
      range.setNumberFormats(formats);
      range.setValues([rowValues]);
    }
    
    return subJobcardNo;
  } catch (e) {
    throw new Error("Failed to generate report for " + subJobcardNo + ": " + e.toString());
  }
}

/**
 * Backend API: Scans the Fire_Assay sheet for completed samples and automatically
 * generates Draft Reports for any sub-jobs that are not yet in the Reports sheet.
 * 
 * @return {string} Count and status message of reports processed.
 */
function generateAllPendingReports() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var faSheet = ss.getSheetByName("Fire_Assay_Sheet");
    var repSheet = ss.getSheetByName("Reports");
    if (!faSheet || !repSheet) {
      throw new Error("Fire_Assay_Sheet or Reports sheet not found.");
    }
    
    var faData = faSheet.getDataRange().getValues();
    var repData = repSheet.getDataRange().getValues();
    
    // Create map of existing reports
    var existingReports = {};
    for (var i = 1; i < repData.length; i++) {
      var rNo = repData[i][0].toString().trim();
      if (rNo) {
        existingReports[rNo] = true;
      }
    }
    
    // Find unique sub-jobs in Fire Assay
    var pendingSubJobs = [];
    var subJobCheck = {};
    
    for (var i = 1; i < faData.length; i++) {
      var subJob = faData[i][1].toString().trim();
      var isCg = faData[i][13];
      
      if (subJob && !isCg && subJob !== "CG_CALIBRATION") {
        if (!existingReports[subJob] && !subJobCheck[subJob]) {
          subJobCheck[subJob] = true;
          pendingSubJobs.push(subJob);
        }
      }
    }
    
    var count = 0;
    for (var i = 0; i < pendingSubJobs.length; i++) {
      generateReport(pendingSubJobs[i]);
      count++;
    }
    
    return count + " reports generated successfully.";
  } catch (e) {
    throw new Error("Failed to generate all pending reports: " + e.toString());
  }
}

/**
 * Formats columns and headers in the 'Reports' sheet.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss The active spreadsheet.
 */
function setupReportsSheet(ss) {
  var repSheet = ss.getSheetByName("Reports");
  if (!repSheet) return;
  
  var lastRow = repSheet.getLastRow();
  if (lastRow > 0) {
    // Already initialized. Skip.
    return;
  }
  
  var headers = [
    "Report No",             // A (Sub-Jobcard No)
    "Report Date",            // B
    "Receipt No",            // C
    "Receipt Date",           // D
    "Date of Test",           // E
    "Customer Name",         // F
    "Customer Address",      // G
    "Description of Sample", // H
    "Received Weight (g)",   // I
    "Sampling Method",       // J
    "Date of Sampling",      // K
    "Return Cornet Weight (g)", // L
    "Return Sample Weight (g)", // M
    "Gold Fineness 1 (ppt)", // N
    "Gold Fineness 2 (ppt)", // O
    "Mean Value (ppt)",      // P
    "Additional Information",// Q
    "Deviations from Standard", // R
    "Unusual Features",      // S
    "Tested By",             // T
    "Reviewed By",           // U
    "Report Status",         // V (Draft/Final)
    "ULR Number"             // W (Unique Lab Reference)
  ];
  
  repSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatSheetCommon(repSheet, headers.length, "#10B981"); // Tab color: green, headers styled via formatSheetCommon
}
