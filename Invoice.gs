/**
 * Gold & Silver Assay Laboratory Management System (LMS)
 * Client: NCH Services Pvt. Ltd., Surat, Gujarat
 * Phase 6: Dispatch Tracking + Invoice Generation (Invoice.gs)
 *
 * This file handles database setups for the Invoice sheet,
 * invoice calculation logic, and building printable Invoice slips.
 */

/**
 * Menu Callback: Opens the "New Invoice" HTML dialog modal.
 */
function menuNewInvoice() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('InvoiceDialog')
        .setWidth(650)
        .setHeight(600);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Invoice Management');
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("UI Error", "Could not open Invoice dialog: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Backend API: Returns a list of receipts with their invoice statuses.
 * Includes both pending and already invoiced receipts to allow re-printing.
 * 
 * @return {Array<Object>} List of receipts.
 */
function getInvoiceReceiptList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var recSheet = ss.getSheetByName("Receiving");
    var invSheet = ss.getSheetByName("Invoice");
    if (!recSheet) return [];
    
    var recData = recSheet.getDataRange().getValues();
    var invData = invSheet ? invSheet.getDataRange().getValues() : [];
    
    // 1. Gather all unique base receipts and customer details from Receiving
    var receiptMap = {};
    for (var i = 1; i < recData.length; i++) {
      var subJobNo = recData[i][0] ? recData[i][0].toString().trim() : "";
      if (!subJobNo) continue;
      
      var receiptNo = recData[i][14] ? recData[i][14].toString().trim() : "";
      if (!receiptNo) continue;
      
      if (!receiptMap[receiptNo]) {
        receiptMap[receiptNo] = {
          receiptNo: receiptNo,
          customerName: recData[i][3] ? recData[i][3].toString().trim() : "",
          sampleCount: 0,
          rate: parseFloat(recData[i][11]) || 200, // Column L: Rate
          invoiceNo: "",
          status: "Pending"
        };
      }
      receiptMap[receiptNo].sampleCount++;
    }
    
    // 2. Cross-reference with existing Invoice sheet records
    for (var i = 1; i < invData.length; i++) {
      var rNo = invData[i][2] ? invData[i][2].toString().trim() : ""; // Column C: Receipt No
      if (rNo && receiptMap[rNo]) {
        receiptMap[rNo].invoiceNo = invData[i][0] ? invData[i][0].toString().trim() : ""; // Column A: Invoice No
        receiptMap[rNo].status = "Invoiced";
      }
    }
    
    var list = [];
    var keys = Object.keys(receiptMap);
    for (var i = 0; i < keys.length; i++) {
      var r = receiptMap[keys[i]];
      list.push(r);
    }
    
    // Sort receipt numbers descending
    list.sort(function(a, b) {
      return b.receiptNo.localeCompare(a.receiptNo);
    });
    
    return list;
  } catch (e) {
    Logger.log("Error in getInvoiceReceiptList: " + e.toString());
    return [];
  }
}

/**
 * Computes the next unique Invoice Number in sequence (INV-YYYY-001, INV-YYYY-002...).
 * 
 * @return {string} Next formatted invoice number.
 */
function getNextInvoiceNumber() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  var year = new Date().getFullYear();
  if (!sheet) return "INV-" + year + "-001";
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "INV-" + year + "-001";
  
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxNum = 0;
  
  for (var i = 0; i < values.length; i++) {
    var val = values[i][0].toString();
    if (val.indexOf("INV-") === 0) {
      var parts = val.split('-');
      if (parts.length === 3) {
        var num = parseInt(parts[2], 10);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
  }
  
  var nextNum = maxNum + 1;
  var paddedNum = ("000" + nextNum).slice(-3);
  return "INV-" + year + "-" + paddedNum;
}

/**
 * Backend API: Generates or updates an Invoice row for a receipt.
 * 
 * @param {string} receiptNo The Receipt Number to invoice.
 * @param {number} ratePerSample Billing rate per sample.
 * @param {string} notes Additional billing notes.
 * @return {string} Created or updated Invoice Number.
 */
function generateInvoice(receiptNo, ratePerSample, notes) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var recSheet = ss.getSheetByName("Receiving");
    var invSheet = ss.getSheetByName("Invoice");
    
    if (!recSheet || !invSheet) {
      throw new Error("Receiving or Invoice sheet not found.");
    }
    
    var recData = recSheet.getDataRange().getValues();
    var customerName = "";
    var customerGST = "";
    var customerAddress = "";
    var sampleCount = 0;
    
    // Gather details from Receiving
    for (var i = 1; i < recData.length; i++) {
      var subJobNo = recData[i][0] ? recData[i][0].toString().trim() : "";
      var baseReceiptNo = recData[i][14] ? recData[i][14].toString().trim() : "";
      
      if (baseReceiptNo === receiptNo) {
        if (!customerName) {
          customerName = recData[i][3] ? recData[i][3].toString().trim() : "";
          customerGST = recData[i][5] ? recData[i][5].toString().trim() : "";
          customerAddress = recData[i][6] ? recData[i][6].toString().trim() : "";
        }
        sampleCount++;
      }
    }
    
    if (sampleCount === 0) {
      throw new Error("No samples found for receipt no: " + receiptNo);
    }
    
    // Check if an Invoice already exists
    var invData = invSheet.getDataRange().getValues();
    var existingRowIndex = -1;
    var invoiceNo = "";
    var existingStatus = "Pending";
    var existingPayDate = "";
    var existingPayMode = "";
    
    for (var i = 1; i < invData.length; i++) {
      var rNo = invData[i][2] ? invData[i][2].toString().trim() : "";
      if (rNo === receiptNo) {
        existingRowIndex = i + 1;
        invoiceNo = invData[i][0] ? invData[i][0].toString().trim() : "";
        existingStatus = invData[i][12] ? invData[i][12].toString().trim() : "Pending";
        
        var pDateVal = invData[i][13];
        if (pDateVal instanceof Date) {
          existingPayDate = Utilities.formatDate(pDateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          existingPayDate = pDateVal ? pDateVal.toString() : "";
        }
        existingPayMode = invData[i][14] ? invData[i][14].toString().trim() : "";
        break;
      }
    }
    
    if (!invoiceNo) {
      invoiceNo = getNextInvoiceNumber();
    }
    
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var rate = parseFloat(ratePerSample) || 200.0;
    var subtotal = sampleCount * rate;
    var gstPercent = 18.0;
    var gstAmount = subtotal * (gstPercent / 100);
    var totalAmount = subtotal + gstAmount;
    
    var rowValues = [
      invoiceNo,          // A: Invoice No
      todayStr,           // B: Invoice Date
      receiptNo,          // C: Receipt No
      customerName,       // D: Customer Name
      customerGST,        // E: Customer GST
      customerAddress,    // F: Customer Address
      sampleCount,        // G: No of Samples
      rate,               // H: Rate per Sample
      subtotal,           // I: Subtotal
      gstPercent,         // J: GST %
      gstAmount,          // K: GST Amount
      totalAmount,        // L: Total Amount
      existingStatus,     // M: Payment Status
      existingPayDate,    // N: Payment Date
      existingPayMode,    // O: Payment Mode
      notes || ""         // P: Notes
    ];
    
    if (existingRowIndex !== -1) {
      // Update existing
      invSheet.getRange(existingRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      // Append new
      var startRow = invSheet.getLastRow() + 1;
      var range = invSheet.getRange(startRow, 1, 1, rowValues.length);
      
      // Formatting
      var formats = [];
      var formatRow = [];
      for (var c = 0; c < rowValues.length; c++) {
        if (c === 0 || c === 2 || c === 4) {
          formatRow.push("@"); 
        } else if (c === 7 || c === 8 || c === 10 || c === 11) {
          formatRow.push("₹#,##0.00");
        } else {
          formatRow.push("");
        }
      }
      formats.push(formatRow);
      range.setNumberFormats(formats);
      range.setValues([rowValues]);
    }
    
    return invoiceNo;
  } catch (e) {
    throw new Error("Failed to generate invoice: " + e.toString());
  }
}

/**
 * Backend API: Creates a beautifully styled, print-ready "Invoice_Print" sheet
 * using the metadata of a specific invoice.
 * 
 * @param {string} invoiceNo The Invoice Number to compile and format.
 * @return {string} Success confirmation message.
 */
function printInvoice(invoiceNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = ss.getSheetByName("Invoice");
    if (!invSheet) throw new Error("Invoice sheet not found.");
    
    var invData = invSheet.getDataRange().getValues();
    var rowValues = null;
    
    for (var i = 1; i < invData.length; i++) {
      var currentInv = invData[i][0] ? invData[i][0].toString().trim() : "";
      if (currentInv === invoiceNo) {
        rowValues = invData[i];
        break;
      }
    }
    
    if (!rowValues) {
      throw new Error("Invoice not found: " + invoiceNo);
    }
    
    // Extract variables
    var invoiceDateVal = rowValues[1];
    var receiptNo = rowValues[2];
    var customerName = rowValues[3];
    var customerGST = rowValues[4];
    var customerAddress = rowValues[5];
    var sampleCount = parseInt(rowValues[6]) || 0;
    var rate = parseFloat(rowValues[7]) || 0;
    var subtotal = parseFloat(rowValues[8]) || 0;
    var gstPercent = parseFloat(rowValues[9]) || 0;
    var gstAmount = parseFloat(rowValues[10]) || 0;
    var totalAmount = parseFloat(rowValues[11]) || 0;
    var notes = rowValues[15] || "";
    
    var formatDate = function(val) {
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd-MM-yyyy");
      }
      return val ? val.toString() : "";
    };
    
    var invoiceDate = formatDate(invoiceDateVal);
    
    // Create or clear Invoice_Print sheet
    var printSheetName = "Invoice_Print";
    var tempSheet = ss.getSheetByName(printSheetName);
    if (tempSheet) {
      ss.deleteSheet(tempSheet);
    }
    tempSheet = ss.insertSheet(printSheetName);
    
    // Hide gridlines
    tempSheet.setHiddenGridlines(true);
    
    // Set column widths (A: S.No, B: Description, C: Qty, D: Amount)
    var widths = [50, 240, 110, 140];
    for (var col = 1; col <= widths.length; col++) {
      tempSheet.setColumnWidth(col, widths[col - 1]);
    }
    
    // Color Palette
    var primaryColor = "#1E293B"; // slate-800
    var dividerColor = "#E2E8F0"; // slate-200
    var lightBgColor = "#F8FAFC"; // slate-50
    var grayTextColor = "#64748B"; // slate-500
    
    // 1. BILLING SLIP LETTERHEAD
    tempSheet.getRange("A2:D2").merge().setValue("NCH SERVICES PVT. LTD.").setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center").setFontColor(primaryColor);
    tempSheet.getRange("A3:D3").merge().setValue("GSTIN: 24AADCN6275G1ZF").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setFontColor(grayTextColor);
    tempSheet.getRange("A4:D4").merge().setValue("L-22/23, Gujarat Hira Bourse, GJB Ichchhapore, Hazira, Surat - 394510").setFontSize(9).setHorizontalAlignment("center").setFontColor(grayTextColor);
    tempSheet.getRange("A5:D5").merge().setBorder(null, null, true, null, null, null, primaryColor, SpreadsheetApp.BorderStyle.DOUBLE);
    
    // 2. INVOICE TITLE
    tempSheet.getRange("A7:D7").merge().setValue("TAX INVOICE").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center").setBackground(lightBgColor).setFontColor(primaryColor);
    tempSheet.getRange("A7:D7").setBorder(true, true, true, true, null, null, dividerColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // 3. METADATA BLOCK
    // Left Grid: Invoice Metadata
    tempSheet.getRange("A9").setValue("Invoice No:").setFontWeight("bold").setFontColor(grayTextColor).setFontSize(9);
    tempSheet.getRange("B9").setValue(invoiceNo).setNumberFormat("@").setFontSize(9);
    tempSheet.getRange("A10").setValue("Invoice Date:").setFontWeight("bold").setFontColor(grayTextColor).setFontSize(9);
    tempSheet.getRange("B10").setValue(invoiceDate).setFontSize(9);
    tempSheet.getRange("A11").setValue("Receipt Ref:").setFontWeight("bold").setFontColor(grayTextColor).setFontSize(9);
    tempSheet.getRange("B11").setValue(receiptNo).setNumberFormat("@").setFontSize(9);
    
    // Right Grid: Billing Info
    tempSheet.getRange("C9").setValue("Bill To:").setFontWeight("bold").setFontColor(primaryColor).setFontSize(9);
    tempSheet.getRange("D9").setValue(customerName).setFontWeight("bold").setFontSize(9);
    tempSheet.getRange("C10").setValue("Address:").setFontWeight("bold").setFontColor(grayTextColor).setFontSize(9);
    tempSheet.getRange("D10").setValue(customerAddress).setFontSize(8).setWrap(true);
    tempSheet.getRange("C11").setValue("Customer GST:").setFontWeight("bold").setFontColor(grayTextColor).setFontSize(9);
    tempSheet.getRange("D11").setValue(customerGST || "N/A").setNumberFormat("@").setFontSize(9);
    
    // Border metadata
    tempSheet.getRange("A9:D12").setBorder(true, true, true, true, null, null, dividerColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // 4. LINE ITEMS TABLE HEADERS
    tempSheet.getRange("A14").setValue("S.No").setFontWeight("bold").setHorizontalAlignment("center").setBackground(lightBgColor).setFontColor(primaryColor);
    tempSheet.getRange("B14").setValue("Description").setFontWeight("bold").setHorizontalAlignment("left").setBackground(lightBgColor).setFontColor(primaryColor);
    tempSheet.getRange("C14").setValue("Qty (Samples)").setFontWeight("bold").setHorizontalAlignment("center").setBackground(lightBgColor).setFontColor(primaryColor);
    tempSheet.getRange("D14").setValue("Amount").setFontWeight("bold").setHorizontalAlignment("right").setBackground(lightBgColor).setFontColor(primaryColor);
    
    // 5. TABLE ROW
    tempSheet.getRange("A15").setValue("1").setHorizontalAlignment("center");
    tempSheet.getRange("B15").setValue("Fire Assay Lab Testing (IS 1418:2009)\nRate: ₹" + rate.toFixed(2) + " per sample").setWrap(true);
    tempSheet.getRange("C15").setValue(sampleCount).setHorizontalAlignment("center");
    tempSheet.getRange("D15").setValue(subtotal).setNumberFormat("₹#,##0.00").setHorizontalAlignment("right");
    
    // Empty padding rows in table
    tempSheet.getRange("A16").setValue("");
    tempSheet.getRange("B16").setValue("");
    tempSheet.getRange("C16").setValue("");
    tempSheet.getRange("D16").setValue("");
    
    // Table border
    tempSheet.getRange("A14:D16").setBorder(true, true, true, true, true, true, dividerColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // 6. TOTALS PANEL
    var startT = 17;
    // Subtotal
    tempSheet.getRange(startT, 2).setValue("Subtotal").setHorizontalAlignment("right").setFontWeight("bold").setFontColor(grayTextColor);
    tempSheet.getRange(startT, 4).setValue(subtotal).setNumberFormat("₹#,##0.00").setHorizontalAlignment("right");
    
    // GST
    tempSheet.getRange(startT + 1, 2).setValue("GST @ 18%").setHorizontalAlignment("right").setFontWeight("bold").setFontColor(grayTextColor);
    tempSheet.getRange(startT + 1, 4).setValue(gstAmount).setNumberFormat("₹#,##0.00").setHorizontalAlignment("right");
    
    // TOTAL
    tempSheet.getRange(startT + 2, 2).setValue("TOTAL (INR)").setHorizontalAlignment("right").setFontWeight("bold").setFontColor(primaryColor);
    tempSheet.getRange(startT + 2, 4).setValue(totalAmount).setNumberFormat("₹#,##0.00").setHorizontalAlignment("right").setFontWeight("bold").setFontColor("#10B981");
    
    tempSheet.getRange(startT, 2, 3, 3).setBorder(true, true, true, true, true, true, dividerColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // Notes block
    if (notes) {
      tempSheet.getRange("A21").setValue("Notes:").setFontWeight("bold").setFontColor(grayTextColor);
      tempSheet.getRange("B21:D21").merge().setValue(notes).setFontStyle("italic").setFontColor(grayTextColor);
    }
    
    // 7. AUTHORISED SIGNATORY
    tempSheet.getRange("A23:D23").merge().setValue("Authorised Signatory: NCH Services Pvt. Ltd.").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("right").setFontColor(primaryColor);
    tempSheet.getRange("A24:D24").merge().setValue("------------------------------------------------").setHorizontalAlignment("right").setFontColor(grayTextColor);
    
    // Outer border
    tempSheet.getRange("A1:D26").setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    // Row height setups
    tempSheet.setRowHeight(1, 15);
    tempSheet.setRowHeight(2, 24);
    tempSheet.setRowHeight(3, 18);
    tempSheet.setRowHeight(4, 18);
    tempSheet.setRowHeight(5, 10);
    tempSheet.setRowHeight(7, 24);
    tempSheet.setRowHeight(14, 24);
    tempSheet.setRowHeight(15, 30); // fit 2 lines description
    
    // Set view focus
    ss.setActiveSheet(tempSheet);
    
    // Alert user
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      "Invoice Slip Ready",
      "Invoice " + invoiceNo + " is compiled in tab '" + printSheetName + "'.\n\n" +
      "Printing Instructions:\n" +
      "1. Press Ctrl + P to open Print settings.\n" +
      "2. Select Portrait layout and Fit to Width.\n" +
      "3. Set margins to Normal/Custom to align borders correctly.",
      ui.ButtonSet.OK
    );
    
    return "Successfully loaded invoice print preview.";
  } catch (e) {
    throw new Error("Failed to compile invoice print: " + e.toString());
  }
}

/**
 * Formats columns and headers in the 'Invoice' sheet.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss The active spreadsheet.
 */
function setupInvoiceSheet(ss) {
  var invSheet = ss.getSheetByName("Invoice");
  if (!invSheet) return;
  
  var lastRow = invSheet.getLastRow();
  if (lastRow > 0) {
    // Already initialized. Skip.
    return;
  }
  
  var headers = [
    "Invoice No",        // A
    "Invoice Date",      // B
    "Receipt No",        // C
    "Customer Name",     // D
    "Customer GST",      // E
    "Customer Address",  // F
    "No of Samples",     // G
    "Rate per Sample (₹)", // H
    "Subtotal (₹)",      // I
    "GST % (18%)",       // J
    "GST Amount (₹)",    // K
    "Total Amount (₹)",  // L
    "Payment Status",    // M (Pending/Paid)
    "Payment Date",      // N
    "Payment Mode",      // O
    "Notes"              // P
  ];
  
  invSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatSheetCommon(invSheet, headers.length, null); // Headers styled via formatSheetCommon, no custom tab color requested
}
