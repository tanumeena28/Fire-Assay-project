/**
 * Gold & Silver Assay Laboratory Management System (LMS)
 * Client: NCH Services Pvt. Ltd., Surat, Gujarat
 * Phase 5: Test Report Generation and Printing (ReportPrint.gs)
 *
 * This file contains the formatting engine that constructs the official
 * printable NCH letterhead layout on the 'Print_Preview' tab.
 */

/**
 * Backend API: Creates a beautifully styled, print-ready "Print_Preview" tab
 * using the metadata of a specific report.
 * 
 * @param {string} reportNo The Report Number (Sub-Jobcard) to compile and format.
 * @return {string} Success confirmation message.
 */
function printTestReport(reportNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var repSheet = ss.getSheetByName("Reports");
    if (!repSheet) throw new Error("Reports sheet not found. Please run setup.");
    
    var data = repSheet.getDataRange().getValues();
    var rowValues = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === reportNo) {
        rowValues = data[i];
        break;
      }
    }
    
    if (!rowValues) {
      throw new Error("Report not found for Report No: " + reportNo);
    }
    
    // Extract variables
    var reportDateVal = rowValues[1];
    var receiptNo = rowValues[2];
    var receiptDateVal = rowValues[3];
    var dateOfTestVal = rowValues[4];
    var customerName = rowValues[5];
    var customerAddress = rowValues[6];
    var description = rowValues[7];
    var recdWt = parseFloat(rowValues[8]) || 0;
    var samplingMethod = rowValues[9];
    var dateSampVal = rowValues[10];
    var returnCornetWt = parseFloat(rowValues[11]) || 0;
    var returnSampleWt = parseFloat(rowValues[12]) || 0;
    var fineness1 = parseFloat(rowValues[13]) || 0;
    var fineness2 = parseFloat(rowValues[14]) || 0;
    var meanVal = parseFloat(rowValues[15]) || 0;
    var additionalInfo = rowValues[16];
    var deviations = rowValues[17];
    var unusualFeatures = rowValues[18];
    var testedBy = rowValues[19];
    var reviewedBy = rowValues[20];
    var status = rowValues[21];
    var ulrNo = rowValues[22];
    
    // Format dates
    var formatDate = function(val) {
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd-MM-yyyy");
      }
      return val ? val.toString() : "";
    };
    
    var reportDate = formatDate(reportDateVal);
    var receiptDate = formatDate(receiptDateVal);
    var dateOfTest = formatDate(dateOfTestVal);
    var dateSamp = formatDate(dateSampVal);
    
    // Create or clear Print_Preview tab (optimized to prevent slow sheet creation delay)
    var tempSheetName = "Print_Preview";
    var tempSheet = ss.getSheetByName(tempSheetName);
    if (tempSheet) {
      tempSheet.clear();
      tempSheet.clearFormats();
      tempSheet.clearNotes();
    } else {
      tempSheet = ss.insertSheet(tempSheetName);
    }
    
    // Setup page and hide gridlines for clean print document feel
    tempSheet.setHiddenGridlines(true);
    
    // Set column widths to A = 160, B = 190, C = 160, D = 190 (Total 700px)
    var columnWidths = [160, 190, 160, 190];
    for (var col = 1; col <= columnWidths.length; col++) {
      tempSheet.setColumnWidth(col, columnWidths[col - 1]);
    }
    
    var primaryColor = "#1E293B"; // Slate-800
    
    // 1. BRAND HEADER (A2:D4)
    tempSheet.getRange("B2:C2").merge().setValue("NCH SERVICES PVT LTD")
      .setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(primaryColor);
    tempSheet.getRange("B3:C3").merge().setValue("( Assay Laboratory )")
      .setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(primaryColor);
    tempSheet.getRange("A4:D4").merge().setValue("Ichchhapore, Surat, Gujarat.")
      .setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(primaryColor);
    
    // Top box borders
    tempSheet.getRange("A2:A4").setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
    tempSheet.getRange("D2:D4").setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
    tempSheet.getRange("B2:C4").setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
    tempSheet.getRange("A2:D4").setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID);

    // 2. REPORT TITLE (A6:D7)
    tempSheet.getRange("A6:D6").merge().setValue("TEST REPORT (GOLD FIRE ASSAY)")
      .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(primaryColor);
    tempSheet.getRange("A7:D7").merge().setValue("Fire Assay As Per IS 1418:2009")
      .setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(primaryColor);
    tempSheet.getRange("A6:D7").setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // 3. METADATA GRID (Rows 9–16)
    // Row 9
    tempSheet.getRange("A9").setValue("TEST REPORT NO");
    tempSheet.getRange("B9").setValue(reportNo).setNumberFormat("@");
    tempSheet.getRange("C9").setValue("REPORT DATE");
    tempSheet.getRange("D9").setValue(reportDate);
    
    // Row 10
    tempSheet.getRange("A10").setValue("RECEIPT NO.");
    tempSheet.getRange("B10").setValue(receiptNo).setNumberFormat("@");
    tempSheet.getRange("C10").setValue("RECEIPT DATE");
    tempSheet.getRange("D10").setValue(receiptDate);
    
    // Row 11
    tempSheet.getRange("A11").setValue("DATE OF TEST");
    tempSheet.getRange("B11:D11").merge().setValue(dateOfTest);
    
    // Row 12
    tempSheet.getRange("A12").setValue("CUSTOMER NAME");
    tempSheet.getRange("B12:D12").merge().setValue(customerName);
    
    // Row 13
    tempSheet.getRange("A13").setValue("CUSTOMER ADDRESS");
    tempSheet.getRange("B13:D13").merge().setValue(customerAddress);
    
    // Row 14
    tempSheet.getRange("A14").setValue("DESCRIPTION OF SAMPLE");
    tempSheet.getRange("B14").setValue(description);
    tempSheet.getRange("C14").setValue("RECEIVED SAMPLE WEIGHT (in grams)");
    tempSheet.getRange("D14").setValue(recdWt).setNumberFormat("0.000");
    
    // Row 15
    tempSheet.getRange("A15").setValue("Sampling Plan & Procedure");
    tempSheet.getRange("B15:D15").merge().setValue(samplingMethod || "Not Applicable - Sample tested as received.");
    
    // Row 16
    tempSheet.getRange("A16").setValue("RETURN CORNET WEIGHT (in grams)");
    tempSheet.getRange("B16").setValue(returnCornetWt).setNumberFormat("0.000");
    tempSheet.getRange("C16").setValue("RETURN SAMPLE WEIGHT (in grams)");
    tempSheet.getRange("D16").setValue(returnSampleWt).setNumberFormat("0.000");
    
    // 4. GOLD FINENESS SUB-TABLE (Rows 17-18)
    tempSheet.getRange("A17:A18").merge().setValue("GOLD FINENESS IN\nPPT")
      .setFontWeight("bold").setFontSize(9).setVerticalAlignment("middle");
    
    tempSheet.getRange("B17").setValue("FIRST VALUE").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setBackground("#F1F5F9");
    tempSheet.getRange("C17").setValue("SECOND VALUE").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setBackground("#F1F5F9");
    tempSheet.getRange("D17").setValue("MEAN VALUE").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setBackground("#F1F5F9");
    
    tempSheet.getRange("B18").setValue(fineness1).setNumberFormat("0.0").setHorizontalAlignment("center");
    tempSheet.getRange("C18").setValue(fineness2).setNumberFormat("0.0").setHorizontalAlignment("center");
    tempSheet.getRange("D18").setValue(meanVal).setNumberFormat("0.0").setFontWeight("bold").setHorizontalAlignment("center");
    
    // Row 19: Additional Info
    tempSheet.getRange("A19").setValue("ADDITIONAL\nINFORMATION");
    tempSheet.getRange("B19:D19").merge().setValue(additionalInfo || "N/A");
    
    // Row 20: Deviations
    tempSheet.getRange("A20:C20").merge().setValue("IF RELEVANT ANY DEVIATIONS FROM THE METHOD SPECIFIED\nIN THIS STANDARD.")
      .setFontSize(8).setVerticalAlignment("middle");
    tempSheet.getRange("D20").setValue(deviations || "No").setHorizontalAlignment("left");
    
    // Row 21: Unusual features
    tempSheet.getRange("A21:C21").merge().setValue("ANY UNUSUAL FEATURES OBSERVED DURING\nTHE DETERMINATION.")
      .setFontSize(8).setVerticalAlignment("middle");
    tempSheet.getRange("D21").setValue(unusualFeatures || "No").setHorizontalAlignment("left");
    
    // Set borders for metadata grid & fineness subtable (A9:D21)
    tempSheet.getRange("A9:D21").setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // Formatting properties of grid ranges
    tempSheet.getRange("A9:D21").setFontSize(9).setVerticalAlignment("middle");
    tempSheet.getRange("A9:A16").setFontWeight("bold");
    tempSheet.getRange("C9").setFontWeight("bold");
    tempSheet.getRange("C10").setFontWeight("bold");
    tempSheet.getRange("C14").setFontWeight("bold");
    tempSheet.getRange("C16").setFontWeight("bold");
    tempSheet.getRange("A19").setFontWeight("bold");
    tempSheet.getRange("A20").setFontWeight("bold");
    tempSheet.getRange("A21").setFontWeight("bold");
    
    // Alignments
    tempSheet.getRange("A9:A16").setHorizontalAlignment("left");
    tempSheet.getRange("C9").setHorizontalAlignment("left");
    tempSheet.getRange("C10").setHorizontalAlignment("left");
    tempSheet.getRange("C14").setHorizontalAlignment("left");
    tempSheet.getRange("C16").setHorizontalAlignment("left");
    
    tempSheet.getRange("B9").setHorizontalAlignment("left");
    tempSheet.getRange("B10").setHorizontalAlignment("left");
    tempSheet.getRange("B11").setHorizontalAlignment("left");
    tempSheet.getRange("B12").setHorizontalAlignment("left");
    tempSheet.getRange("B13").setHorizontalAlignment("left");
    tempSheet.getRange("B14").setHorizontalAlignment("left");
    tempSheet.getRange("B15").setHorizontalAlignment("left");
    tempSheet.getRange("B16").setHorizontalAlignment("left");
    tempSheet.getRange("D9").setHorizontalAlignment("left");
    tempSheet.getRange("D10").setHorizontalAlignment("left");
    tempSheet.getRange("D14").setHorizontalAlignment("left");
    tempSheet.getRange("D16").setHorizontalAlignment("left");
    tempSheet.getRange("B19").setHorizontalAlignment("left");
    
    // Row 22: END OF TEST REPORT
    tempSheet.getRange("A22:D22").merge().setValue("***END OF TEST REPORT***")
      .setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor("#64748B");
    tempSheet.getRange("A22:D22").setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // Row 24: Signatures spacer / placeholder
    tempSheet.setRowHeight(23, 20); // Empty row height
    tempSheet.setRowHeight(24, 35); // Signature area height
    
    // Tested By
    tempSheet.getRange("A25:B25").merge().setValue(testedBy || "Eknath Dalal").setHorizontalAlignment("center").setFontSize(9);
    tempSheet.getRange("A26:B26").merge().setValue("TESTED BY").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(9);
    
    // Reviewed / Authorized By
    tempSheet.getRange("C25:D25").merge().setValue(reviewedBy || "Ketan Varlekar").setHorizontalAlignment("center").setFontSize(9);
    tempSheet.getRange("C26:D26").merge().setValue("CHECKED AND AUTHORIZED BY\n( Technical Manager )")
      .setFontWeight("bold").setHorizontalAlignment("center").setFontSize(9).setVerticalAlignment("middle");
    
    // Borders for Signature Boxes
    tempSheet.getRange("A25:B26").setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
    tempSheet.getRange("C25:D26").setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // Row 27 Spacer
    tempSheet.setRowHeight(27, 10);
    
    // 5. NOTES AND DISCLAIMERS
    tempSheet.getRange("A28:D28").merge().setValue("NOTE :1) THIS TEST REPORT IS ONLY FOR THE SAMPLE TESTED")
      .setFontWeight("bold").setFontSize(8).setHorizontalAlignment("left").setVerticalAlignment("middle");
    tempSheet.getRange("A29:D29").merge().setValue("2) THIS REPORT IS FULL OR IN PART SHALL NOT TO BE REPRODUCED, PUBLISHED OR ADVERTISED OR ANY LEGAL ACTION UNLESS PRIOR PERMISSION HAS BEEN SECURED FROM THE DIRECTOR OF NCH SERVICES PVT LTD.")
      .setFontWeight("bold").setFontSize(8).setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true);
    tempSheet.setRowHeight(29, 32);
    
    // Divider
    tempSheet.getRange("A30:D30").merge().setBorder(null, null, true, null, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // Row 31: Document Control Line
    tempSheet.getRange("A31:D31").merge().setValue("Doc. No. F/7.8/01 / Issue No.: 01 / Issue Date: 01-08-2023 / Amend No. : 00 / Amend Date: -- / Page 1 of 1 / Controlled Document")
      .setFontSize(7).setHorizontalAlignment("left").setVerticalAlignment("middle").setFontColor("#64748B");
    
    // Set row height configurations for spacious aesthetic but compact enough for 1 page
    tempSheet.setRowHeight(2, 20);
    tempSheet.setRowHeight(3, 15);
    tempSheet.setRowHeight(4, 18);
    tempSheet.setRowHeight(6, 20);
    tempSheet.setRowHeight(7, 16);
    
    for (var r = 9; r <= 16; r++) {
      tempSheet.setRowHeight(r, 20);
    }
    tempSheet.setRowHeight(17, 18);
    tempSheet.setRowHeight(18, 20);
    tempSheet.setRowHeight(19, 20);
    tempSheet.setRowHeight(20, 24);
    tempSheet.setRowHeight(21, 24);
    tempSheet.setRowHeight(22, 18);
    tempSheet.setRowHeight(26, 22);
    
    // Wrap text globally for metadata table where needed (like Address, Sampling, Deviations)
    tempSheet.getRange("B13:D13").setWrap(true);
    tempSheet.getRange("B15:D15").setWrap(true);
    tempSheet.getRange("A20:C20").setWrap(true);
    tempSheet.getRange("A21:C21").setWrap(true);
    
    // Set outer medium border around the entire A1:D31 sheet layout
    tempSheet.getRange("A1:D31").setBorder(true, true, true, true, null, null, primaryColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    // Set view focus
    ss.setActiveSheet(tempSheet);
    
    // Prompt print instructions
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      "Report Ready", 
      "Test Report for " + reportNo + " has been successfully generated in tab 'Print_Preview' matching the original layout.\n\n" +
      "Instructions:\n" +
      "1. Press Ctrl + P to open Print settings.\n" +
      "2. Select Portrait layout and Fit to Page Width.\n" +
      "3. Set margins to Normal/Custom to align the A4 page layout correctly.", 
      ui.ButtonSet.OK
    );
    
    return "Successfully loaded print preview.";
  } catch (e) {
    throw new Error("Failed to compile print preview: " + e.toString());
  }
}
