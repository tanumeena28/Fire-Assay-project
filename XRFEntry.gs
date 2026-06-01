/**
 * Gold & Silver Assay Laboratory Management System (LMS)
 * Client: NCH Services Pvt. Ltd., Surat, Gujarat
 * Module: XRF Entry and Printable Ledger Sheet Generation (XRFEntry.gs)
 *
 * This file handles fetching job cards, saving raw element analysis,
 * updating Job Cards with Au/Ag ppt readings, and generating the printable XRF ledger.
 */

/**
 * Menu Callback: Opens the "XRF Sheet Entry" HTML dialog.
 */
function menuXRFEntry() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('XRFEntryForm')
        .setWidth(950)
        .setHeight(650);
    SpreadsheetApp.getUi().showModelessDialog(html, 'XRF Sheet Entry');
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("UI Error", "Could not open XRF sheet entry dialog: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Backend API: Returns all Job Cards matching a given Receipt Number to fill their XRF details.
 * 
 * @param {string} receiptNo The Receipt Number to look up.
 * @return {Array<Object>} List of samples in the job card.
 */
function getXrfPendingSamples(query) {
  try {
    if (!query) return [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var jcSheet = ss.getSheetByName("Job_Cards");
    if (!jcSheet) return [];
    
    var data = jcSheet.getDataRange().getValues();
    var samples = [];
    var addedJobNos = {}; // To prevent duplicates
    
    // Split query by commas and clean each term
    var parts = query.split(",").map(function(p) { return p.trim().toUpperCase(); }).filter(Boolean);
    
    for (var p = 0; p < parts.length; p++) {
      var part = parts[p];
      for (var i = 1; i < data.length; i++) {
        var jcNo = data[i][0].toString().trim().toUpperCase(); // Column A: Job Card No
        var rNo = data[i][1].toString().trim().toUpperCase();  // Column B: Receipt No
        
        // Match either Receipt No or Job Card No
        if ((rNo === part || jcNo === part) && !addedJobNos[jcNo]) {
          addedJobNos[jcNo] = true;
          samples.push({
            subJobNo: data[i][0].toString().trim(),     // Column A: Job Card No
            receiptNo: data[i][1].toString().trim(),    // Column B: Receipt No
            customerName: data[i][3].toString().trim(), // Column D: Customer Name
            material: data[i][4].toString().trim(),     // Column E: Material
            declaredFineness: data[i][5].toString().trim(), // Column F: Declared Fineness
            weight: parseFloat(data[i][6]) || 0         // Column G: Weight Received (g)
          });
        }
      }
    }
    return samples;
  } catch (e) {
    Logger.log("Error in getXrfPendingSamples: " + e.toString());
    return [];
  }
}

/**
 * Backend API: Saves raw XRF data to XRF_Gold/XRF_Silver sheets, updates Job_Cards, and renders printable ledger.
 * 
 * @param {Object} formData Submitted values containing receiptNo, samples element breakdowns, and metadata.
 * @return {string} Success confirmation message.
 */
function saveXrfData(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var queryStr = formData.receiptNo; // Can be comma-separated receipt/job card nos
    var samples = formData.samples || [];
    
    if (samples.length === 0) {
      throw new Error("No sample records provided to save.");
    }
    
    var jcSheet = ss.getSheetByName("Job_Cards");
    if (!jcSheet) {
      throw new Error("Job_Cards sheet not found.");
    }
    
    var jcData = jcSheet.getDataRange().getValues();
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    for (var i = 0; i < samples.length; i++) {
      var s = samples[i];
      var materialType = s.material;
      var dbSheetName = (materialType === "Silver") ? "XRF_Silver" : "XRF_Gold";
      var dbSheet = ss.getSheetByName(dbSheetName);
      
      if (!dbSheet) {
        throw new Error("Database sheet " + dbSheetName + " not found. Run setup.");
      }
      
      // XRF Database columns:
      // Sub-Jobcard No, Receipt No, Customer Name, Material, Weight (g), Declared Fineness,
      // Au, Ag, Cu, Ni, Pd, Zn, Cd, Ir, Ru, Os, In, Sn, Other,
      // Date, Temperature, Time, Check By, XRF Sheet No, Created Date
      var row = [
        s.subJobNo,
        s.receiptNo || queryStr,
        s.customerName,
        s.material,
        parseFloat(s.weight) || 0,
        s.declaredFineness,
        parseFloat(s.Au) || 0,
        parseFloat(s.Ag) || 0,
        parseFloat(s.Cu) || 0,
        parseFloat(s.Ni) || 0,
        parseFloat(s.Pd) || 0,
        parseFloat(s.Zn) || 0,
        parseFloat(s.Cd) || 0,
        parseFloat(s.Ir) || 0,
        parseFloat(s.Ru) || 0,
        parseFloat(s.Os) || 0,
        parseFloat(s.In) || 0,
        parseFloat(s.Sn) || 0,
        parseFloat(s.Other) || 0,
        formData.date,
        formData.temperature,
        formData.time,
        formData.checkBy,
        formData.xrfSheetNo,
        todayStr
      ];
      
      // Save record to respective database
      var startRow = dbSheet.getLastRow() + 1;
      var range = dbSheet.getRange(startRow, 1, 1, row.length);
      
      var formatRow = [];
      for (var c = 0; c < row.length; c++) {
        if (c === 0 || c === 1 || c === 23) {
          formatRow.push("@");
        } else {
          formatRow.push("");
        }
      }
      range.setNumberFormats([formatRow]);
      range.setValues([row]);
      
      // Update Job_Cards with Au & Ag ppt readings (percent * 10)
      var auPpt = (parseFloat(s.Au) || 0) * 10;
      var agPpt = (parseFloat(s.Ag) || 0) * 10;
      
      for (var j = 1; j < jcData.length; j++) {
        if (jcData[j][0].toString().trim() === s.subJobNo) {
          jcSheet.getRange(j + 1, 8).setValue(auPpt); // Column H: XRF Au Reading (ppt)
          jcSheet.getRange(j + 1, 9).setValue(agPpt); // Column I: Ag XRF Reading (ppt)
        }
      }
    }
    
    // Flush changes to ensure database is updated
    SpreadsheetApp.flush();
    
    // Generate the printable layout sheet
    generatePrintableXrfSheet(formData);
    
    // Sanitize receipt string for safe Google Sheets tab name
    var cleanName = queryStr.replace(/[:\\\/?*\[\]]/g, "-");
    if (cleanName.length > 20) {
      cleanName = cleanName.substring(0, 20);
    }
    var tempSheetName = "Print_XRF_" + cleanName;
    
    // Alert the user on the sheet side so they know it is done and ready to print
    var ui = SpreadsheetApp.getUi();
    ui.alert("XRF Sheet Generated", "Printable XRF Sheet for " + queryStr + " has been successfully generated in tab '" + tempSheetName + "'.\n\nPlease go to this tab and press Ctrl + P to print.", ui.ButtonSet.OK);
    
    return "XRF data successfully saved and printable ledger created.";
  } catch (e) {
    throw new Error("Failed to save XRF data: " + e.toString());
  }
}

/**
 * Creates a beautiful, print-ready XRF Sheet matching the user's ledger format.
 * 
 * @param {Object} formData The XRF Entry data structure.
 */
function generatePrintableXrfSheet(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var receiptNo = formData.receiptNo;
  
  var cleanName = receiptNo.replace(/[:\\\/?*\[\]]/g, "-");
  if (cleanName.length > 20) {
    cleanName = cleanName.substring(0, 20);
  }
  var tempSheetName = "Print_XRF_" + cleanName;
  var tempSheet = ss.getSheetByName(tempSheetName);
  if (tempSheet) {
    ss.deleteSheet(tempSheet);
  }
  tempSheet = ss.insertSheet(tempSheetName);
  
  tempSheet.setHiddenGridlines(true);
  
  // Set Column Widths A-P
  tempSheet.setColumnWidth(1, 110); // A: CODE NO
  for (var col = 2; col <= 14; col++) {
    tempSheet.setColumnWidth(col, 42); // B-N: Elements
  }
  tempSheet.setColumnWidth(15, 110); // O: Labels
  tempSheet.setColumnWidth(16, 120); // P: Values
  
  // Outer border around A2:P7
  var headerRange = tempSheet.getRange("A2:P7");
  headerRange.setBorder(true, true, true, true, false, false, "#1A3C5E", SpreadsheetApp.BorderStyle.MEDIUM);
  
  // Logo block (A2:B7)
  tempSheet.getRange("A2:B7").merge()
           .setValue("❖")
           .setFontColor("#2563EB")
           .setFontSize(28)
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle")
           .setBorder(null, null, null, true, null, null, "#1A3C5E", SpreadsheetApp.BorderStyle.SOLID);
           
  // Title block (C2:K7)
  tempSheet.getRange("C2:K3").merge()
           .setValue("NCH SERVICES PVT. LTD.")
           .setFontWeight("bold")
           .setFontSize(14)
           .setFontColor("#1A3C5E")
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle");
  tempSheet.getRange("C4:K4").merge()
           .setValue("(Assay Laboratory)")
           .setFontStyle("italic")
           .setFontSize(10)
           .setFontColor("#64748B")
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle");
  tempSheet.getRange("C5:K7").merge()
           .setValue("XRF SHEET")
           .setFontWeight("bold")
           .setFontSize(13)
           .setFontColor("#1A3C5E")
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle");
           
  tempSheet.getRange("K2:K7").setBorder(null, null, null, true, null, null, "#1A3C5E", SpreadsheetApp.BorderStyle.SOLID);
  
  // Doc Control L2:P7
  var setDocText = function(cell, val, bold, size, align) {
    var r = tempSheet.getRange(cell);
    r.setValue(val).setFontSize(size || 8).setFontColor("#1E293B").setVerticalAlignment("middle");
    if (bold) r.setFontWeight("bold");
    if (align) r.setHorizontalAlignment(align);
  };
  
  // Row 2: Doc No
  tempSheet.getRange("L2:N2").merge();
  setDocText("L2", "Doc. No. :", true, 8, "left");
  tempSheet.getRange("O2:P2").merge();
  setDocText("O2", "F/7.6/02", false, 8, "left");
  
  // Row 3: Issue No & Date
  tempSheet.getRange("L3:M3").merge();
  setDocText("L3", "Issue No. :", true, 8, "left");
  setDocText("N3", "01", false, 8, "center");
  setDocText("O3", "Issue Date:", true, 8, "left");
  setDocText("P3", "01-08-23", false, 8, "left");
  
  // Row 4: Amend No & Date
  tempSheet.getRange("L4:M4").merge();
  setDocText("L4", "Amend No.:", true, 8, "left");
  setDocText("N4", "00", false, 8, "center");
  setDocText("O4", "Amend Date:", true, 8, "left");
  setDocText("P4", "--", false, 8, "left");
  
  // Row 5: Prepared & Issued by
  tempSheet.getRange("L5:N5").merge();
  setDocText("L5", "Prepared & Issued by", true, 8, "left");
  tempSheet.getRange("O5:P5").merge();
  setDocText("O5", "Quality Manager", false, 8, "left");
  
  // Row 6: Reviewed & Approved by
  tempSheet.getRange("L6:N6").merge();
  setDocText("L6", "Reviewed & Approved by", true, 8, "left");
  tempSheet.getRange("O6:P6").merge();
  setDocText("O6", "Chairman", false, 8, "left");
  
  // Row 7: Page & MASTER Doc
  tempSheet.getRange("L7:N7").merge();
  setDocText("L7", "Page No.2", true, 8, "left");
  tempSheet.getRange("O7:P7").merge();
  setDocText("O7", "MASTER Document", true, 8, "left");
  
  tempSheet.getRange("L2:P7").setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  
  // Sheet No
  tempSheet.getRange("O1:P1").merge()
           .setValue("XRF SHEET NO:-" + (formData.xrfSheetNo || ""))
           .setFontWeight("bold")
           .setFontSize(10)
           .setFontColor("#1E293B")
           .setHorizontalAlignment("right")
           .setVerticalAlignment("bottom");
           
  // Table headers in Row 9
  var headers = ["CODE NO", "Au", "Ag", "Cu", "Ni", "Pd", "Zn", "Cd", "Ir", "Ru", "Os", "In", "Sn", "Other"];
  tempSheet.getRange(9, 1, 1, 14).setValues([headers]);
  tempSheet.getRange("O9").setValue("DATE");
  
  var dateStr = "";
  if (formData.date) {
    var dateObj = new Date(formData.date);
    if (!isNaN(dateObj.getTime())) {
      dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd MMM yyyy");
    } else {
      dateStr = formData.date.toString();
    }
  }
  tempSheet.getRange("P9").setValue(dateStr);
  
  // Construct 8 Rows data
  var rowData = [];
  var samples = formData.samples || [];
  for (var r = 0; r < 8; r++) {
    var row = [];
    if (r < samples.length) {
      var s = samples[r];
      row.push(s.subJobNo);
      row.push(s.Au !== "" ? parseFloat(s.Au) : "-");
      row.push(s.Ag !== "" ? parseFloat(s.Ag) : "-");
      row.push(s.Cu !== "" ? parseFloat(s.Cu) : "-");
      row.push(s.Ni !== "" ? parseFloat(s.Ni) : "-");
      row.push(s.Pd !== "" ? parseFloat(s.Pd) : "-");
      row.push(s.Zn !== "" ? parseFloat(s.Zn) : "-");
      row.push(s.Cd !== "" ? parseFloat(s.Cd) : "-");
      row.push(s.Ir !== "" ? parseFloat(s.Ir) : "-");
      row.push(s.Ru !== "" ? parseFloat(s.Ru) : "-");
      row.push(s.Os !== "" ? parseFloat(s.Os) : "-");
      row.push(s.In !== "" ? parseFloat(s.In) : "-");
      row.push(s.Sn !== "" ? parseFloat(s.Sn) : "-");
      row.push(s.Other !== "" ? parseFloat(s.Other) : "-");
    } else {
      row.push("");
      for (var c = 0; c < 13; c++) row.push("-");
    }
    rowData.push(row);
  }
  tempSheet.getRange(10, 1, 8, 14).setValues(rowData);
  
  // Sidebar Metadatas
  tempSheet.getRange("O10:O11").merge().setValue("TEMPERATURE");
  tempSheet.getRange("P10:P11").merge().setValue(formData.temperature || "24 C°");
  
  tempSheet.getRange("O12:O13").merge().setValue("TIME");
  tempSheet.getRange("P12:P13").merge().setValue(formData.time || "30 second");
  
  tempSheet.getRange("O14:O17").merge().setValue("CHECK BY");
  tempSheet.getRange("P14").setValue(formData.checkBy || "");
  tempSheet.getRange("P15:P17").merge().setValue(""); // Signature Space
  
  // Alignments and styles
  tempSheet.setRowHeight(9, 26);
  for (var r = 10; r <= 17; r++) {
    tempSheet.setRowHeight(r, 22);
  }
  
  var tableRange = tempSheet.getRange("A9:P17");
  tableRange.setFontSize(9)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setFontColor("#1E293B");
            
  tempSheet.getRange("A9:P9").setFontWeight("bold").setBackground("#F8FAFC");
  tempSheet.getRange("O10:O17").setFontWeight("bold").setBackground("#F8FAFC");
  tempSheet.getRange("A10:A17").setHorizontalAlignment("left").setFontWeight("bold");
  
  // Format numeric values
  tempSheet.getRange("B10:N17").setNumberFormat("0.00");
  
  // Borders
  tableRange.setBorder(true, true, true, true, null, null, "#1A3C5E", SpreadsheetApp.BorderStyle.MEDIUM);
  tempSheet.getRange("A9:N17").setBorder(null, null, null, null, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  tempSheet.getRange("N9:N17").setBorder(null, null, null, true, null, null, "#1A3C5E", SpreadsheetApp.BorderStyle.SOLID);
  tempSheet.getRange("O9:O17").setBorder(null, null, null, true, null, null, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  tempSheet.getRange("O9:P9").setBorder(null, null, true, null, null, null, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  tempSheet.getRange("O10:P11").setBorder(null, null, true, null, null, null, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  tempSheet.getRange("O12:P13").setBorder(null, null, true, null, null, null, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  tempSheet.getRange("P14").setBorder(null, null, true, null, null, null, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  
  ss.setActiveSheet(tempSheet);
  SpreadsheetApp.flush();
}
