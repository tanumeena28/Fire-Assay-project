/**
 * Gold & Silver Assay Laboratory Management System (LMS)
 * Client: NCH Services Pvt. Ltd., Surat, Gujarat
 * Phase 6: Dispatch Tracking + Invoice Generation (Dispatch.gs)
 *
 * This file handles database setup for the Dispatch sheet,
 * querying completed receipts, and logging dispatch actions.
 */

/**
 * Menu Callback: Opens the "Dispatch Entry" HTML dialog modal.
 */
function menuDispatchEntry() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('DispatchForm')
        .setWidth(650)
        .setHeight(600);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Dispatch Tracking');
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("UI Error", "Could not open Dispatch Tracking dialog: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Backend API: Returns a list of unique Receipt Numbers that have generated reports.
 * 
 * @return {Array<string>} Unique Receipt Numbers.
 */
function getCompletedReceiptsForDispatch() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var repSheet = ss.getSheetByName("Reports");
    if (!repSheet) return [];
    
    var data = repSheet.getDataRange().getValues();
    var receipts = [];
    var seen = {};
    
    for (var i = 1; i < data.length; i++) {
      var receiptNo = data[i][2] ? data[i][2].toString().trim() : ""; // Column C: Receipt No
      if (receiptNo && !seen[receiptNo]) {
        seen[receiptNo] = true;
        receipts.push(receiptNo);
      }
    }
    
    // Sort receipt numbers descending
    receipts.sort(function(a, b) {
      return b.localeCompare(a);
    });
    
    return receipts;
  } catch (e) {
    Logger.log("Error in getCompletedReceiptsForDispatch: " + e.toString());
    return [];
  }
}

/**
 * Backend API: Retrieves customer name, report count, and existing dispatch log (if any)
 * for a specific Receipt Number.
 * 
 * @param {string} receiptNo The Receipt Number to look up.
 * @return {Object} Details of the receipt.
 */
function getReceiptDetailsForDispatch(receiptNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var repSheet = ss.getSheetByName("Reports");
    var dispSheet = ss.getSheetByName("Dispatch");
    
    if (!repSheet || !dispSheet) {
      throw new Error("Reports or Dispatch sheet not found.");
    }
    
    var repData = repSheet.getDataRange().getValues();
    var customerName = "";
    var reportCount = 0;
    
    // 1. Gather data from Reports sheet
    for (var i = 1; i < repData.length; i++) {
      var rNo = repData[i][2] ? repData[i][2].toString().trim() : "";
      if (rNo === receiptNo) {
        if (!customerName) {
          customerName = repData[i][5] ? repData[i][5].toString().trim() : ""; // Column F: Customer Name
        }
        reportCount++;
      }
    }
    
    // 2. Check if a Dispatch log already exists
    var dispData = dispSheet.getDataRange().getValues();
    var existingLog = null;
    
    for (var i = 1; i < dispData.length; i++) {
      var dNo = dispData[i][0] ? dispData[i][0].toString().trim() : "";
      if (dNo === receiptNo) {
        var dDateVal = dispData[i][3]; // Column D
        var dDateStr = "";
        if (dDateVal instanceof Date) {
          dDateStr = Utilities.formatDate(dDateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          dDateStr = dDateVal ? dDateVal.toString() : "";
        }
        
        existingLog = {
          dispatchDate: dDateStr,
          dispatchedBy: dispData[i][4] ? dispData[i][4].toString().trim() : "",
          courierName: dispData[i][5] ? dispData[i][5].toString().trim() : "",
          fireReportSent: dispData[i][6] === "Yes" || dispData[i][6] === true,
          invoiceGenerated: dispData[i][7] === "Yes" || dispData[i][7] === true,
          stickerDone: dispData[i][8] === "Yes" || dispData[i][8] === true,
          qrCodeDone: dispData[i][9] === "Yes" || dispData[i][9] === true,
          reportVerification: dispData[i][10] === "Yes" || dispData[i][10] === true,
          remarks: dispData[i][11] ? dispData[i][11].toString().trim() : ""
        };
        break;
      }
    }
    
    return {
      receiptNo: receiptNo,
      customerName: customerName,
      reportCount: reportCount,
      existingLog: existingLog
    };
  } catch (e) {
    throw new Error("Failed to get receipt details: " + e.toString());
  }
}

/**
 * Backend API: Saves or updates a dispatch record in the "Dispatch" sheet.
 * 
 * @param {Object} data The dispatch entry fields.
 * @return {string} Status message.
 */
function saveDispatchEntry(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dispSheet = ss.getSheetByName("Dispatch");
    if (!dispSheet) throw new Error("Dispatch sheet not found.");
    
    var receiptNo = data.receiptNo;
    var customerName = data.customerName;
    var reportCount = parseInt(data.reportCount) || 0;
    var dispatchDate = data.dispatchDate; // yyyy-MM-dd
    var dispatchedBy = data.dispatchedBy;
    var courierName = data.courierName;
    var fireReportSent = data.fireReportSent ? "Yes" : "No";
    var invoiceGenerated = data.invoiceGenerated ? "Yes" : "No";
    var stickerDone = data.stickerDone ? "Yes" : "No";
    var qrCodeDone = data.qrCodeDone ? "Yes" : "No";
    var reportVerification = data.reportVerification ? "Yes" : "No";
    var remarks = data.remarks;
    
    var dispData = dispSheet.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var i = 1; i < dispData.length; i++) {
      var dNo = dispData[i][0] ? dispData[i][0].toString().trim() : "";
      if (dNo === receiptNo) {
        rowIndex = i + 1; // 1-indexed row number
        break;
      }
    }
    
    var rowValues = [
      receiptNo,           // A
      customerName,        // B
      reportCount,         // C
      dispatchDate,        // D
      dispatchedBy,        // E
      courierName,         // F
      fireReportSent,      // G
      invoiceGenerated,    // H
      stickerDone,         // I
      qrCodeDone,          // J
      reportVerification,  // K
      remarks              // L
    ];
    
    if (rowIndex !== -1) {
      // Update existing record
      dispSheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      // Append new record
      var startRow = dispSheet.getLastRow() + 1;
      var range = dispSheet.getRange(startRow, 1, 1, rowValues.length);
      
      // Formatting
      var formats = [];
      var formatRow = [];
      for (var c = 0; c < rowValues.length; c++) {
        if (c === 0) {
          formatRow.push("@"); 
        } else {
          formatRow.push("");
        }
      }
      formats.push(formatRow);
      range.setNumberFormats(formats);
      range.setValues([rowValues]);
    }
    
    // Also, if the report status needs to be updated to Dispatched, we can do it here.
    // For now we keep it as a ledger log.
    
    return "Dispatch details saved successfully for " + receiptNo;
  } catch (e) {
    throw new Error("Failed to save dispatch entry: " + e.toString());
  }
}

/**
 * Formats columns and headers in the 'Dispatch' sheet.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss The active spreadsheet.
 */
function setupDispatchSheet(ss) {
  var dispSheet = ss.getSheetByName("Dispatch");
  if (!dispSheet) return;
  
  var lastRow = dispSheet.getLastRow();
  if (lastRow > 0) {
    // Already initialized. Skip.
    return;
  }
  
  var headers = [
    "Receipt No",              // A
    "Customer Name",           // B
    "No of Reports",           // C
    "Dispatch Date",           // D
    "Dispatched By",           // E
    "Courier Name",            // F
    "Fire Report Sent (Yes/No)",// G
    "Invoice Generated (Yes/No)",// H
    "Sticker Done (Yes/No)",   // I
    "QR Code Done (Yes/No)",   // J
    "Report Verification (Yes/No)",// K
    "Remarks"                  // L
  ];
  
  dispSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatSheetCommon(dispSheet, headers.length, null); // Headers styled via formatSheetCommon, no custom tab color requested
}
