/**
 * Gold & Silver Assay Laboratory Management System (LMS)
 * Client: NCH Services Pvt. Ltd., Surat, Gujarat
 * Phase 3: Quality Manager Review Module (QCReview.gs)
 *
 * This file manages the QC Review logic. It syncs pending samples from Receiving
 * to the Review_of_Request sheet, serves pending receipt lists, and saves reviews.
 */

/**
 * Menu Callback: Opens the "QC Review Form" HTML sidebar.
 */
function menuReviewOfRequest() {
  try {
    var html;
    try {
      html = HtmlService.createHtmlOutputFromFile('QCReview');
    } catch (err) {
      html = HtmlService.createHtmlOutputFromFile('qCReview');
    }
    html.setWidth(960).setHeight(700);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Quality Manager Review');
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("UI Error", "Could not open QC Review form: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Backward compatibility wrapper for old menu item registration.
 */
function menuQCReview() {
  menuReviewOfRequest();
}

/**
 * Backend API: Reads all rows from the Receiving sheet and finds samples not yet reviewed.
 * Populates the Review_of_Request sheet with these pending samples and highlights them in yellow.
 */
function loadPendingQC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var receivingSheet = ss.getSheetByName("Receiving");
  var qcSheet = ss.getSheetByName("Review_of_Request");
  
  if (!receivingSheet || !qcSheet) {
    throw new Error("Receiving or Review_of_Request sheet not found. Please run project setup.");
  }
  
  var receivingData = receivingSheet.getDataRange().getValues();
  var qcData = qcSheet.getDataRange().getValues();
  
  // Create map of already loaded sub-jobcards in Review_of_Request to avoid duplicates
  var existingQCSubJobs = {};
  for (var i = 1; i < qcData.length; i++) {
    var subJob = qcData[i][1].toString().trim(); // Column B: Sub-Jobcard No
    if (subJob) {
      existingQCSubJobs[subJob] = true;
    }
  }
  
  var pendingRows = [];
  for (var i = 1; i < receivingData.length; i++) {
    var row = receivingData[i];
    var subJob = row[0].toString().trim(); // Column A: Receipt No (Sub-Jobcard No)
    if (subJob && !existingQCSubJobs[subJob]) {
      var baseReceiptNo = row[14] ? row[14].toString().trim() : "";
      var custName = row[3];     // Column D: Customer Name
      var material = row[7];     // Column H: Material Type
      var weight = row[9];       // Column J: Total Weight (g)
      var witnessed = row[16];   // Column Q: Customer Witnessed
      
      // Review_of_Request Columns:
      // A: Receipt No, B: Sub-Jobcard No, C: Customer Name, D: Material, E: Weight (g),
      // F: Weight OK, G: Customer Witnessed, H: Scope, I: Sample Accepted,
      // J: XRF Check By, K: XRF Fineness, L: Plate Fineness, M: QC Done By, N: QC Date, O: Remarks
      pendingRows.push([
        baseReceiptNo, // A
        subJob,        // B
        custName,      // C
        material,      // D
        weight,        // E
        "",            // F: Weight OK (Pending)
        witnessed,     // G: Customer Witnessed (From Receiving)
        "",            // H: Scope (Pending)
        "",            // I: Sample Accepted (Pending)
        "",            // J: XRF Check By (Pending)
        "",            // K: XRF Fineness (Pending)
        "",            // L: Plate Fineness (Pending)
        "",            // M: QC Done By (Pending)
        "",            // N: QC Date (Pending)
        ""             // O: Remarks
      ]);
    }
  }
  
  if (pendingRows.length > 0) {
    var startRow = qcSheet.getLastRow() + 1;
    var range = qcSheet.getRange(startRow, 1, pendingRows.length, pendingRows[0].length);
    
    // Set text format for receipt identifiers
    var formats = [];
    for (var r = 0; r < pendingRows.length; r++) {
      var formatRow = [];
      for (var c = 0; c < pendingRows[0].length; c++) {
        if (c === 0 || c === 1) {
          formatRow.push("@"); // Plain text formatting
        } else {
          formatRow.push("");
        }
      }
      formats.push(formatRow);
    }
    range.setNumberFormats(formats);
    range.setValues(pendingRows);
    
    // Highlight newly loaded rows in soft yellow (#FEF08A)
    var highlightRange = qcSheet.getRange(startRow, 1, pendingRows.length, pendingRows[0].length);
    highlightRange.setBackground("#FEF08A");
    
    return pendingRows.length + " pending samples loaded successfully.";
  } else {
    return "No new pending samples found in Receiving.";
  }
}

/**
 * Backend API: Returns unique Receipt Numbers that have pending reviews in Review_of_Request.
 * (A receipt is pending if at least one sub-jobcard does not have "Sample Accepted" filled).
 * 
 * @return {Array<string>} List of pending receipt numbers.
 */
function getPendingQCReceipts() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var qcSheet = ss.getSheetByName("Review_of_Request");
    if (!qcSheet) return [];
    
    var data = qcSheet.getDataRange().getValues();
    var pendingReceiptsMap = {};
    
    for (var i = 1; i < data.length; i++) {
      var receiptNo = data[i][0].toString().trim(); // Column A
      var accepted = data[i][8].toString().trim();  // Column I: Sample Accepted
      if (receiptNo && !accepted) {
        pendingReceiptsMap[receiptNo] = true;
      }
    }
    return Object.keys(pendingReceiptsMap);
  } catch (e) {
    Logger.log("Error in getPendingQCReceipts: " + e.toString());
    return [];
  }
}

/**
 * Backend API: Retrieves the detailed records of all sub-jobcards under a specific Receipt.
 * Used to populate the review UI fields.
 * 
 * @param {string} receiptNo The Receipt Number to filter by.
 * @return {Array<Object>} List of sample details under the receipt.
 */
function getQCDetailsForReceipt(receiptNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var qcSheet = ss.getSheetByName("Review_of_Request");
    if (!qcSheet) return [];
    
    var data = qcSheet.getDataRange().getValues();
    var samples = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rNo = row[0].toString().trim();
      if (rNo === receiptNo) {
        samples.push({
          rowNum: i + 1, // 1-indexed row number in the sheet
          subJobNo: row[1].toString(),
          customerName: row[2].toString(),
          material: row[3].toString(),
          weight: parseFloat(row[4]) || 0,
          witnessed: row[6].toString()
        });
      }
    }
    return samples;
  } catch (e) {
    Logger.log("Error in getQCDetailsForReceipt: " + e.toString());
    return [];
  }
}

/**
 * Backend API: Commits QC Review inputs back to the Review_of_Request sheet.
 * Clears the yellow background color once the review is completed.
 * 
 * @param {Object} formData Form variables containing XRF fineness, acceptance details, etc.
 * @return {string} Success confirmation message.
 */
function saveQCReview(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var qcSheet = ss.getSheetByName("Review_of_Request");
    if (!qcSheet) {
      throw new Error("Review_of_Request sheet not found.");
    }
    
    var samples = formData.samples || [];
    var now = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    if (samples.length === 0) return "No samples to update.";
    
    // Find min and max row to check if they are contiguous for batching
    var minRow = Infinity;
    var maxRow = -Infinity;
    var rowMap = {};
    
    for (var i = 0; i < samples.length; i++) {
      var sample = samples[i];
      var rowNum = parseInt(sample.rowNum, 10);
      if (isNaN(rowNum)) continue;
      if (rowNum < minRow) minRow = rowNum;
      if (rowNum > maxRow) maxRow = rowNum;
      rowMap[rowNum] = sample;
    }
    
    var isContiguous = (maxRow - minRow + 1 === samples.length);
    
    if (isContiguous && samples.length > 0) {
      // Contiguous batch update: write all rows in a single batch
      var updateValues = [];
      for (var r = minRow; r <= maxRow; r++) {
        var sample = rowMap[r];
        updateValues.push([
          sample.weightOk ? "Yes" : "No",         // F: Weight OK
          formData.witnessed,                     // G: Customer Witnessed
          formData.scope,                         // H: Scope
          formData.accepted,                      // I: Sample Accepted
          formData.xrfCheckBy,                    // J: XRF Check By
          parseFloat(sample.xrfFineness) || 0,    // K: XRF Fineness
          parseFloat(sample.plateFineness) || 0,  // L: Plate Fineness
          formData.qcDoneBy,                      // M: QC Done By
          dateStr,                                // N: QC Date
          formData.remarks                        // O: Remarks
        ]);
      }
      // Set values and backgrounds in single calls
      qcSheet.getRange(minRow, 6, updateValues.length, 10).setValues(updateValues);
      qcSheet.getRange(minRow, 1, updateValues.length, 15).setBackground(null);
    } else {
      // Non-contiguous fallback: write one row at a time (still optimized with setValues)
      for (var i = 0; i < samples.length; i++) {
        var sample = samples[i];
        var rowNum = parseInt(sample.rowNum, 10);
        if (isNaN(rowNum)) continue;
        
        var rowValues = [
          sample.weightOk ? "Yes" : "No",
          formData.witnessed,
          formData.scope,
          formData.accepted,
          formData.xrfCheckBy,
          parseFloat(sample.xrfFineness) || 0,
          parseFloat(sample.plateFineness) || 0,
          formData.qcDoneBy,
          dateStr,
          formData.remarks
        ];
        qcSheet.getRange(rowNum, 6, 1, 10).setValues([rowValues]);
        qcSheet.getRange(rowNum, 1, 1, 15).setBackground(null);
      }
    }
    
    return "QC Review successfully saved for Receipt: " + formData.receiptNo;
  } catch (e) {
    throw new Error("Failed to save QC Review: " + e.toString());
  }
}

/**
 * Formats columns and headers in the 'Review_of_Request' sheet.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss The active spreadsheet.
 */
function setupReviewOfRequestSheet(ss) {
  var qcSheet = ss.getSheetByName("Review_of_Request");
  if (!qcSheet) return;
  
  var lastRow = qcSheet.getLastRow();
  if (lastRow > 0) {
    // Already initialized. Skip.
    return;
  }
  
  var headers = [
    "Receipt No",             // A
    "Sub-Jobcard No",         // B
    "Customer Name",          // C
    "Material",               // D
    "Weight (g)",             // E
    "Weight OK (Yes/No)",     // F
    "Customer Witnessed (Yes/No)", // G
    "Scope (Yes/No)",         // H
    "Sample Accepted (Yes/No)", // I
    "XRF Check By",           // J
    "XRF Fineness Reading (ppt)", // K
    "Plate Fineness Check (ppt)", // L
    "QC Done By",             // M
    "QC Date",                // N
    "Remarks"                 // O
  ];
  
  qcSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatSheetCommon(qcSheet, headers.length, "#7C3AED"); // Tab color: purple
}


