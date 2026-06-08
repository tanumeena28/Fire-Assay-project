/**
 * Gold & Silver Assay Laboratory Management System (LMS)
 * Client: NCH Services Pvt. Ltd., Surat, Gujarat
 *
 * Phase 1: Complete Project Setup & Core Logic
 */

/**
 * Special Google Sheets trigger that runs automatically when the spreadsheet is opened.
 * Builds the custom menu "QAI Lab System" on the menu bar.
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('QAI Lab System')
      .addItem('📥 New Receipt', 'menuNewReceipt')
      .addItem('🔍 Review of Request', 'menuReviewOfRequest')
      .addItem('📋 Create Job Cards', 'menuNewJobCard')
      .addItem('📊 XRF Sheet Entry', 'menuXRFEntry')
      .addItem('🔥 Fire Assay Entry', 'menuFireAssay')
      .addSeparator()
      .addItem('📄 Generate Test Report', 'menuGenerateReport')
      .addItem('🖨️ Print Report', 'menuPrintReport')
      .addSeparator()
      .addItem('📦 Inventory Desk', 'menuInventory')
      .addSeparator()
      .addItem('🔄 Refresh Dashboard', 'menuRefreshDashboard')
      .addItem('⚙️ Setup All Sheets', 'setupProject')
      .addToUi();
  } catch (e) {
    Logger.log("Error in onOpen: " + e.toString());
  }
}

/**
 * Menu Callback: Opens the "New Customer Receipt" HTML form as a Modal Dialog.
 */
function menuNewReceipt() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('Receiving')
        .setWidth(960)
        .setHeight(700);
    SpreadsheetApp.getUi().showModelessDialog(html, 'New Customer Receipt');
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("UI Error", "Could not open receiving form: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Master Project Setup Function.
 * Initializes all sheets, configures columns and formats, seeds initial customer data,
 * and sets up sheet formatting (tab colors, frozen headers, etc.).
 */
function setupProject() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    ui = null;
  }
  
  Logger.log("=== STARTING SETUP ===");
  
  try {
    var targetSheets = [
      "Dashboard",
      "Master",
      "Receiving",
      "Review_of_Request",
      "Job_Cards",
      "XRF_Gold",
      "XRF_Silver",
      "Fire_Assay_Sheet",
      "Reports",
      "Dispatch",
      "Invoice",
      "Inventory",
      "Inventory_Transactions"
    ];
    
    // Ensure all target sheets exist
    Logger.log("Step 1: Checking/creating target sheets...");
    var sheets = ss.getSheets();
    var currentNames = sheets.map(function(s) { return s.getName(); });
    var newlyCreatedSheets = {};
    
    for (var i = 0; i < targetSheets.length; i++) {
      var sheetName = targetSheets[i];
      if (currentNames.indexOf(sheetName) === -1) {
        ss.insertSheet(sheetName, i);
        newlyCreatedSheets[sheetName] = true;
        Logger.log("   Created sheet: " + sheetName);
      }
    }
    
    // Skip deleting Sheet1 to avoid Sheet dependency update hangs
    Logger.log("Step 2: Skipped Sheet1 deletion to prevent API hang.");
    
    // Setup newly created sheets
    Logger.log("Step 3: Setting up individual sheets...");
    
    if (newlyCreatedSheets["Master"]) {
      Logger.log("   - Setting up Master sheet...");
      setupMasterSheet(ss);
    } else {
      Logger.log("   - Master sheet exists, skipping setup.");
    }
    
    if (newlyCreatedSheets["Receiving"]) {
      Logger.log("   - Setting up Receiving sheet...");
      setupReceivingSheet(ss);
    } else {
      Logger.log("   - Receiving sheet exists, skipping setup.");
    }
    
    if (newlyCreatedSheets["Review_of_Request"]) {
      Logger.log("   - Setting up Review of Request sheet...");
      setupReviewOfRequestSheet(ss);
    } else {
      Logger.log("   - Review of Request sheet exists, skipping setup.");
    }
    
    if (newlyCreatedSheets["Job_Cards"]) {
      Logger.log("   - Setting up Job Cards sheet...");
      setupJobCardsSheet(ss);
    } else {
      Logger.log("   - Job Cards sheet exists, skipping setup.");
    }
    
    if (newlyCreatedSheets["XRF_Gold"] || newlyCreatedSheets["XRF_Silver"]) {
      Logger.log("   - Setting up XRF database sheets...");
      setupXRFDatabaseSheets(ss);
    } else {
      Logger.log("   - XRF database sheets exist, skipping setup.");
    }
    
    if (newlyCreatedSheets["Fire_Assay_Sheet"]) {
      Logger.log("   - Setting up Fire Assay sheet...");
      setupFireAssaySheet(ss);
    } else {
      Logger.log("   - Fire Assay sheet exists, skipping setup.");
    }
    
    if (newlyCreatedSheets["Reports"]) {
      Logger.log("   - Setting up Reports sheet...");
      setupReportsSheet(ss);
    } else {
      Logger.log("   - Reports sheet exists, skipping setup.");
    }
    
    if (newlyCreatedSheets["Dispatch"]) {
      Logger.log("   - Setting up Dispatch sheet...");
      setupDispatchSheet(ss);
    } else {
      Logger.log("   - Dispatch sheet exists, skipping setup.");
    }
    
    if (newlyCreatedSheets["Invoice"]) {
      Logger.log("   - Setting up Invoice sheet...");
      setupInvoiceSheet(ss);
    } else {
      Logger.log("   - Invoice sheet exists, skipping setup.");
    }
    
    Logger.log("   - Setting up Inventory sheets...");
    setupInventorySheets(ss);
    
    // Reset selection to Dashboard (Skipped to prevent layout recalculation hang)
    Logger.log("Step 4: Skipped focusing on Dashboard to prevent API hang.");
    
    if (ui) {
      ss.toast("QAI Lab System Ready!", "System Setup", 5);
    } else {
      Logger.log("Setup Completed Successfully!");
    }
    Logger.log("=== SETUP COMPLETED SUCCESSFULLY ===");
    
  } catch (e) {
    if (ui) {
      ui.alert("Setup Failed", "Error: " + e.toString(), ui.ButtonSet.OK);
    } else {
      Logger.log("Setup Failed - Error: " + e.toString());
    }
  }
}

/**
 * Common formatting function to apply branding header, freeze row, tab colors, and auto-resize.
 */
function formatSheetCommon(sheet, numColumns, tabColorHex) {
  if (!sheet) return;
  var cols = numColumns || sheet.getLastColumn() || 1;
  
  // Format header row (Row 1)
  var headerRange = sheet.getRange(1, 1, 1, cols);
  headerRange.setFontWeight("bold")
             .setFontColor("#FFFFFF")
             .setBackground("#1a3c5e") // bold, background #1a3c5e, white text
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");
  
  sheet.setRowHeight(1, 30);
  sheet.setFrozenRows(1);
  
  // Set tab color
  if (tabColorHex) {
    sheet.setTabColor(tabColorHex);
  }
  
}

/**
 * Formats columns and populates customer seed records in the 'Master' sheet.
 */
function setupMasterSheet(ss) {
  var masterSheet = ss.getSheetByName("Master");
  if (!masterSheet) return;
  
  var lastRow = masterSheet.getLastRow();
  if (lastRow > 0) return; // Skip if already initialized
  
  var headers = [
    "Customer Name", "Contact", "GST No", "Address", "Certificate Name", "Email", "State", "City", "Pincode"
  ];
  
  var seedCustomers = [
    ["Hemratna Jewellers (Varachha)", "0261255666", "", "Shop No.3-4 Ground Floor Rajhans Point Varachha Gujarat", "Hemratna Jewellers", "", "Gujarat", "Varachha", ""],
    ["Mani Jewels", "8320225611", "24AAIFM4858N1ZP", "GHB Ichchhapore Surat", "Mani Jewels", "", "Gujarat", "Surat", "394510"],
    ["Svaraa Jewellers P Ltd", "7359002604", "24ABCCS2246L1ZB", "Ground floor N-02 Gujarat Hira Bourse Surat", "Svaraa Jewellers P Ltd", "", "Gujarat", "Surat", "394510"],
    ["NCH Services P Ltd", "", "24AADCN6275G1ZF", "L 22-23 GHB Ichchhapore Surat", "NCH Services P Ltd", "", "Gujarat", "Surat", "394510"],
    ["KTC Hallmarkings", "", "", "227 AVR Building Telugu Street Coimbatore TN 641001", "KTC Hallmarkings", "", "Tamil Nadu", "Coimbatore", "641001"]
  ];
  
  masterSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var dataRange = masterSheet.getRange(2, 1, seedCustomers.length, headers.length);
  dataRange.setNumberFormat("@");
  dataRange.setValues(seedCustomers);
  
  formatSheetCommon(masterSheet, headers.length, null);
}

/**
 * Formats columns and headers in the 'Receiving' sheet.
 */
function setupReceivingSheet(ss) {
  var receivingSheet = ss.getSheetByName("Receiving");
  if (!receivingSheet) return;
  
  var lastRow = receivingSheet.getLastRow();
  if (lastRow > 0) return; // Skip if already initialized
  
  var headers = [
    "Sub-Job No",          // A
    "Date",                // B
    "Time",                // C
    "Customer Name",       // D
    "Contact No",          // E
    "GST No",              // F
    "Address",             // G
    "Material Type",       // H
    "No. of Samples",      // I
    "Total Weight (g)",    // J
    "Process Name",        // K
    "Rate per Sample (₹)", // L
    "Payment Status",      // M
    "Sample Condition",    // N
    "Receipt No",          // O
    "Public Domain",       // P
    "Customer Witnessed",  // Q
    "Report Copy",         // R
    "Qty OK",              // S
    "Received By",         // T
    "Remarks",             // U
    "Certificate Name"     // V
  ];
  
  receivingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatSheetCommon(receivingSheet, headers.length, "#2563EB"); // Tab color: blue
}

/**
 * Formats columns and headers in the 'Review_of_Request' sheet.
 */
function setupReviewOfRequestSheet(ss) {
  var qcSheet = ss.getSheetByName("Review_of_Request");
  if (!qcSheet) return;
  
  var lastRow = qcSheet.getLastRow();
  if (lastRow > 0) return;
  
  var headers = [
    "Receipt No",
    "Sub-Jobcard No",
    "Customer Name",
    "Material",
    "Weight (g)",
    "Weight OK (Yes/No)",
    "Customer Witnessed (Yes/No)",
    "Scope (Yes/No)",
    "Sample Accepted (Yes/No)",
    "XRF Check By",
    "XRF Fineness Reading (ppt)",
    "Plate Fineness Check (ppt)",
    "QC Done By",
    "QC Date",
    "Remarks"
  ];
  
  qcSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatSheetCommon(qcSheet, headers.length, "#7C3AED"); // Tab color: purple
}

/**
 * Formats columns and headers in 'XRF_Gold' and 'XRF_Silver' database sheets.
 */
function setupXRFDatabaseSheets(ss) {
  var sheets = ["XRF_Gold", "XRF_Silver"];
  var headers = [
    "Sub-Jobcard No", "Receipt No", "Customer Name", "Material", "Weight (g)", "Declared Fineness",
    "Au", "Ag", "Cu", "Ni", "Pd", "Zn", "Cd", "Ir", "Ru", "Os", "In", "Sn", "Other",
    "Date", "Temperature", "Time", "Check By", "XRF Sheet No", "Created Date"
  ];
  
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i];
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow === 0 || lastRow === 1) { // setup headers if empty or only 1 row (header row check)
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        formatSheetCommon(sheet, headers.length, "#0D9488"); // Tab color: teal
      }
    }
  }
}

/**
 * Returns a list of customer objects from the Master sheet to populate dropdowns.
 */
function getCustomerList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName("Master");
    if (!masterSheet) return [];
    
    var lastRow = masterSheet.getLastRow();
    if (lastRow <= 1) return [];
    
    var data = masterSheet.getRange(2, 1, lastRow - 1, 5).getValues(); // Fetch 5 columns (A to E)
    var list = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) {
        list.push({
          name: data[i][0].toString(),
          contact: data[i][1] ? data[i][1].toString() : "",
          gst: data[i][2] ? data[i][2].toString() : "",
          address: data[i][3] ? data[i][3].toString() : "",
          certificateName: data[i][4] ? data[i][4].toString() : ""
        });
      }
    }
    return list;
  } catch (e) {
    Logger.log("Error in getCustomerList: " + e.toString());
    return [];
  }
}

/**
 * Helper to get financial year string like "24/25", "25/26" starting April 1st.
 */
function getFinancialYearString(date) {
  var year = date.getFullYear();
  var month = date.getMonth(); // 0-indexed: 0 = Jan, 3 = April
  var fyStart, fyEnd;
  if (month >= 3) { // April or later
    fyStart = year;
    fyEnd = year + 1;
  } else { // Jan, Feb, Mar
    fyStart = year - 1;
    fyEnd = year;
  }
  var startYY = fyStart.toString().slice(-2);
  var endYY = fyEnd.toString().slice(-2);
  return startYY + "/" + endYY;
}

/**
 * Computes the next unique Receipt Number in sequence (ICH-YY/YY-XXXX).
 */
function getNextReceiptNumber() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Receiving");
  
  var now = new Date();
  var fyStr = getFinancialYearString(now); // e.g. "24/25"
  var defaultReceipt = "ICH-" + fyStr + "-0001";
  
  if (!sheet) return defaultReceipt;
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return defaultReceipt;
  
  // Column O (15th column) contains the Receipt No
  var values = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
  var maxNum = 0;
  
  for (var i = 0; i < values.length; i++) {
    var val = values[i][0].toString();
    if (val.indexOf("ICH-") === 0) {
      var parts = val.split('-');
      if (parts.length === 3) {
        var valFy = parts[1]; // "24/25"
        if (valFy === fyStr) {
          var numPart = parts[2]; // "0001"
          var num = parseInt(numPart, 10);
          if (!isNaN(num) && num > maxNum) {
            maxNum = num;
          }
        }
      }
    }
  }
  
  var nextNum = maxNum + 1;
  var paddedNum = ("0000" + nextNum).slice(-4);
  return "ICH-" + fyStr + "-" + paddedNum;
}

/**
 * Calculates the next sequential batch number for the sub-job code (ICHNFAR7, ICHNFAR8...).
 */
function getNextBatchNumber(materialType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Receiving");
  
  var now = new Date();
  var fyStr = getFinancialYearString(now); // e.g. "24/25"
  var defaultBatch = 1; // Default start batch for a new FY is 1
  
  if (!sheet) return defaultBatch;
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return defaultBatch;
  
  // Read Column A (Sub-Job No) and Column O (Receipt No, which contains the FY)
  // Column A is index 0, Column O is index 14, so we read 15 columns
  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  var maxBatch = 0;
  
  var targetPrefix = (materialType === "Silver") ? "ICHNSAR" : "ICHNFAR";
  
  for (var i = 0; i < data.length; i++) {
    var subJob = data[i][0].toString().trim();
    var receipt = data[i][14].toString().trim(); // Column O
    
    // Check if the receipt belongs to the current financial year (e.g., contains ICH-24/25-)
    if (receipt.indexOf("ICH-" + fyStr + "-") === 0) {
      if (subJob.indexOf(targetPrefix) === 0) {
        var basePart = subJob.split('-')[0]; // "ICHNFAR1" or "ICHNSAR1"
        var numPart = basePart.substring(targetPrefix.length); // e.g. "1"
        var num = parseInt(numPart, 10);
        if (!isNaN(num) && num > maxBatch) {
          maxBatch = num;
        }
      }
    }
  }
  
  return maxBatch > 0 ? maxBatch + 1 : defaultBatch;
}

/**
 * Backend API: Saves the submitted form data as rows in the "Receiving" sheet.
 * Generates Receipt No (BH-YY/YY-XXXX) and Sub-Job No (BHNFAR7-XX) for each sample.
 */
function saveCustomerReceipt(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var receivingSheet = ss.getSheetByName("Receiving");
    if (!receivingSheet) {
      throw new Error("Receiving sheet not found.");
    }
    
    var saveSingle = function(materialType, noOfSamples, weights) {
      var now = new Date();
      var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
      var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");
      
      var receiptNo = getNextReceiptNumber();
      var batchNo = getNextBatchNumber(materialType);
      
      var numSamples = parseInt(noOfSamples, 10) || 0;
      var wList = weights || [];
      
      while (wList.length < numSamples) {
        wList.push(0);
      }
      
      var prefix = (materialType === "Silver") ? "ICHNSAR" : "ICHNFAR";
      var rowsToAppend = [];
      for (var i = 0; i < numSamples; i++) {
        var subJobNo = prefix + batchNo + "-" + ("0" + (i + 1)).slice(-2); // ICHNFAR1-01 or ICHNSAR1-01
        var weight = parseFloat(wList[i]) || 0;
        
        var row = [
          subJobNo,              // A
          dateStr,               // B
          timeStr,               // C
          formData.customerName, // D
          formData.contactNo,    // E
          formData.gstNo,        // F
          formData.address,      // G
          materialType,          // H
          numSamples,            // I
          weight,                // J
          formData.processName,  // K
          formData.rate,         // L
          formData.paymentStatus,// M
          formData.sampleCondition, // N
          receiptNo,             // O
          formData.publicDomain, // P
          formData.customerWitnessed, // Q
          formData.reportCopy,   // R
          formData.qtyOk,        // S
          formData.receivedBy,   // T
          formData.remarks,      // U
          formData.certificateName // V
        ];
        rowsToAppend.push(row);
      }
      
      var startRow = receivingSheet.getLastRow() + 1;
      var range = receivingSheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length);
      
      var formats = [];
      for (var r = 0; r < rowsToAppend.length; r++) {
        var formatRow = [];
        for (var c = 0; c < rowsToAppend[0].length; c++) {
          if (c === 0 || c === 4 || c === 5 || c === 14 || c === 21) {
            formatRow.push("@"); 
          } else {
            formatRow.push("");
          }
        }
        formats.push(formatRow);
      }
      
      range.setNumberFormats(formats);
      range.setValues(rowsToAppend);
      
      return receiptNo;
    };
    
    if (formData.materialType === "Both") {
      var goldReceiptNo = saveSingle("Gold", formData.goldNoOfSamples, formData.goldWeights);
      SpreadsheetApp.flush();
      var silverReceiptNo = saveSingle("Silver", formData.silverNoOfSamples, formData.silverWeights);
      SpreadsheetApp.flush();
      
      return {
        type: "Both",
        goldReceiptNo: goldReceiptNo,
        silverReceiptNo: silverReceiptNo
      };
    } else {
      var receiptNo = saveSingle(formData.materialType, formData.noOfSamples, formData.weights);
      SpreadsheetApp.flush();
      return {
        type: "Single",
        receiptNo: receiptNo
      };
    }
  } catch (e) {
    throw new Error("Failed to save receipt: " + e.toString());
  }
}

/**
 * Backend API: Creates a temporary receipt slip worksheet for printing.
 * Custom formatted matching premium letterhead layout.
 */
function printReceiptSlip(receiptNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var receivingSheet = ss.getSheetByName("Receiving");
    if (!receivingSheet) {
      throw new Error("Receiving sheet not found.");
    }
    
    var data = receivingSheet.getDataRange().getValues();
    var matchingRows = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rNo = row[14].toString().trim(); // Column O
      if (rNo === receiptNo) {
        matchingRows.push(row);
      }
    }
    
    if (matchingRows.length === 0) {
      SpreadsheetApp.getUi().alert("Print Error", "No samples found for Receipt: " + receiptNo, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Extract base metadata from first record
    var firstRow = matchingRows[0];
    var dateVal = firstRow[1];
    var custName = firstRow[3];
    var contact = firstRow[4];
    var gst = firstRow[5];
    var address = firstRow[6];
    var material = firstRow[7];
    var processName = firstRow[10];
    var rate = parseFloat(firstRow[11]) || 0;
    var paymentStatus = firstRow[12];
    var receivedBy = firstRow[19];
    var remarks = firstRow[20];
    var certificateName = firstRow[21] ? firstRow[21].toString().trim() : custName;
    
    var totalWeight = 0;
    for (var i = 0; i < matchingRows.length; i++) {
      totalWeight += parseFloat(matchingRows[i][9]) || 0;
    }
    
    var tempSheetName = "Receipt_Slip_" + receiptNo.replace(/\//g, "-"); // escape FY slashes for sheet names
    var tempSheet = ss.getSheetByName(tempSheetName);
    if (tempSheet) {
      ss.deleteSheet(tempSheet);
    }
    tempSheet = ss.insertSheet(tempSheetName);
    
    tempSheet.setHiddenGridlines(true);
    
    var columnWidths = [120, 150, 120, 110, 100];
    for (var col = 1; col <= columnWidths.length; col++) {
      tempSheet.setColumnWidth(col, columnWidths[col - 1]);
    }
    
    tempSheet.setRowHeight(1, 15);
    tempSheet.setRowHeight(2, 26);
    tempSheet.setRowHeight(3, 18);
    tempSheet.setRowHeight(4, 18);
    tempSheet.setRowHeight(5, 28);
    tempSheet.setRowHeight(6, 15);
    
    tempSheet.setRowHeight(7, 22);
    tempSheet.setRowHeight(8, 22);
    tempSheet.setRowHeight(9, 22);
    tempSheet.setRowHeight(10, 22);
    tempSheet.setRowHeight(11, 22); // Address Row
    tempSheet.setRowHeight(12, 15); // Spacer Row
    
    // Brand Header
    tempSheet.getRange("A2:E2").merge().setValue("NCH SERVICES PVT. LTD.").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#1E293B");
    tempSheet.getRange("A3:E3").merge().setValue("Gold & Silver Assay Laboratory").setFontSize(10).setFontStyle("italic").setHorizontalAlignment("center").setFontColor("#475569");
    tempSheet.getRange("A4:E4").merge().setValue("L-22/23, Gujarat Hira Bourse, Gems and Jewellery Park, Ichchhapore, Hazira, Surat - 394510").setFontSize(8).setHorizontalAlignment("center").setFontColor("#64748B");
    tempSheet.getRange("A5:E5").merge().setValue("CUSTOMER RECEIPT SLIP").setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#1E293B").setFontColor("#FFFFFF");
    
    // Receipt Details (Row 7)
    tempSheet.getRange("A7").setValue("Receipt No:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    tempSheet.getRange("B7").setValue(receiptNo).setNumberFormat("@").setFontSize(10).setFontWeight("bold");
    tempSheet.getRange("C7").setValue("Customer Name:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    tempSheet.getRange("D7:E7").merge().setValue(custName).setFontSize(10).setFontWeight("bold");
    
    // Row 8
    tempSheet.getRange("A8").setValue("Date:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    var dateStr = dateVal instanceof Date ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd-MM-yyyy") : dateVal;
    tempSheet.getRange("B8").setValue(dateStr).setNumberFormat("@").setFontSize(10);
    tempSheet.getRange("C8").setValue("Certificate Name:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    tempSheet.getRange("D8:E8").merge().setValue(certificateName).setFontSize(10).setFontWeight("bold");
    
    // Row 9
    tempSheet.getRange("A9").setValue("Material Type:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    tempSheet.getRange("B9").setValue(material).setFontSize(10);
    tempSheet.getRange("C9").setValue("Contact No:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    tempSheet.getRange("D9:E9").merge().setValue(contact).setNumberFormat("@").setFontSize(10);
    
    // Row 10
    tempSheet.getRange("A10").setValue("Process:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    tempSheet.getRange("B10").setValue(processName).setFontSize(10);
    tempSheet.getRange("C10").setValue("GST No:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    tempSheet.getRange("D10:E10").merge().setValue(gst).setNumberFormat("@").setFontSize(10);
    
    // Row 11
    tempSheet.getRange("A11").setValue("Address:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    tempSheet.getRange("B11:E11").merge().setValue(address).setFontSize(9);
    
    tempSheet.getRange("A11:E11").setBorder(null, null, true, null, null, null, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
    
    // Samples Table Header
    var startRow = 13;
    tempSheet.setRowHeight(startRow, 26);
    tempSheet.getRange(startRow, 1).setValue("Sub-Job No").setFontWeight("bold").setBackground("#334155").setFontColor("#FFFFFF").setHorizontalAlignment("center").setFontSize(10);
    tempSheet.getRange(startRow, 2, 1, 2).merge().setValue("Description").setFontWeight("bold").setBackground("#334155").setFontColor("#FFFFFF").setHorizontalAlignment("left").setFontSize(10);
    tempSheet.getRange(startRow, 4).setValue("Weight (g)").setFontWeight("bold").setBackground("#334155").setFontColor("#FFFFFF").setHorizontalAlignment("right").setFontSize(10);
    tempSheet.getRange(startRow, 5).setValue("Rate (₹)").setFontWeight("bold").setBackground("#334155").setFontColor("#FFFFFF").setHorizontalAlignment("right").setFontSize(10);
    
    var currentRow = startRow + 1;
    for (var i = 0; i < matchingRows.length; i++) {
      var rowData = matchingRows[i];
      tempSheet.setRowHeight(currentRow, 24);
      tempSheet.getRange(currentRow, 1).setValue(rowData[0]).setNumberFormat("@").setHorizontalAlignment("center").setFontSize(9);
      tempSheet.getRange(currentRow, 2, 1, 2).merge().setValue(material + " ornament / scrap").setFontSize(9);
      tempSheet.getRange(currentRow, 4).setValue(parseFloat(rowData[9]) || 0).setNumberFormat("0.000").setHorizontalAlignment("right").setFontSize(9);
      tempSheet.getRange(currentRow, 5).setValue(rate).setNumberFormat("₹#,##0.00").setHorizontalAlignment("right").setFontSize(9);
      currentRow++;
    }
    
    // Total Row
    tempSheet.setRowHeight(currentRow, 24);
    tempSheet.getRange(currentRow, 1).setValue("").setBackground("#F8FAFC");
    tempSheet.getRange(currentRow, 2, 1, 2).merge().setValue("TOTALS:").setFontWeight("bold").setBackground("#F8FAFC").setHorizontalAlignment("right").setFontSize(10);
    tempSheet.getRange(currentRow, 4).setValue(totalWeight).setFontWeight("bold").setBackground("#F8FAFC").setNumberFormat("0.000").setHorizontalAlignment("right").setFontSize(10);
    var totalAmount = matchingRows.length * rate;
    tempSheet.getRange(currentRow, 5).setValue(totalAmount).setFontWeight("bold").setBackground("#F8FAFC").setNumberFormat("₹#,##0.00").setHorizontalAlignment("right").setFontSize(10);
    
    var tableRange = tempSheet.getRange(startRow, 1, matchingRows.length + 2, 5);
    tableRange.setBorder(true, true, true, true, true, true, "#94A3B8", SpreadsheetApp.BorderStyle.SOLID);
    
    // Extra details
    currentRow += 2;
    tempSheet.setRowHeight(currentRow, 22);
    tempSheet.getRange(currentRow, 1).setValue("Payment Status:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    
    var payCell = tempSheet.getRange(currentRow, 2);
    payCell.setValue(paymentStatus).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
    if (paymentStatus === "Paid") {
      payCell.setBackground("#DCFCE7").setFontColor("#15803D");
    } else {
      payCell.setBackground("#FEE2E2").setFontColor("#B91C1C");
    }
    
    currentRow += 1;
    tempSheet.setRowHeight(currentRow, 22);
    tempSheet.getRange(currentRow, 1).setValue("Received By:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
    tempSheet.getRange(currentRow, 2).setValue(receivedBy).setFontSize(10);
    
    if (remarks) {
      currentRow += 1;
      tempSheet.setRowHeight(currentRow, 22);
      tempSheet.getRange(currentRow, 1).setValue("Remarks:").setFontWeight("bold").setFontColor("#475569").setFontSize(9);
      tempSheet.getRange(currentRow, 2, 1, 4).merge().setValue(remarks).setFontSize(9).setFontStyle("italic");
    }
    
    // Customer Request Table Section (matches second screenshot)
    currentRow += 2;
    tempSheet.setRowHeight(currentRow, 24);
    tempSheet.getRange(currentRow, 1, 1, 5).merge()
             .setValue("Customer Request")
             .setFontWeight("bold")
             .setFontSize(10)
             .setFontColor("#1E293B");
             
    var requestStartRow = currentRow + 1;
    
    // Item 1
    tempSheet.setRowHeight(requestStartRow, 22);
    tempSheet.getRange(requestStartRow, 1, 1, 4).merge()
             .setValue("Information provided by you are intended to place in the public domain by the testing laboratory, are you agree?")
             .setFontSize(8)
             .setFontColor("#475569")
             .setVerticalAlignment("middle");
    tempSheet.getRange(requestStartRow, 5)
             .setValue(firstRow[15] || "No") // Column P (Public Domain)
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle")
             .setFontSize(9);
             
    // Item 2
    tempSheet.setRowHeight(requestStartRow + 1, 22);
    tempSheet.getRange(requestStartRow + 1, 1, 1, 4).merge()
             .setValue("Test witness by the customer & its representative?")
             .setFontSize(8)
             .setFontColor("#475569")
             .setVerticalAlignment("middle");
    tempSheet.getRange(requestStartRow + 1, 5)
             .setValue(firstRow[16] || "No") // Column Q (Customer Witnessed)
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle")
             .setFontSize(9);
             
    // Item 3
    tempSheet.setRowHeight(requestStartRow + 2, 22);
    tempSheet.getRange(requestStartRow + 2, 1, 1, 4).merge()
             .setValue("Hard Copy or Soft Copy of the Test Report required?")
             .setFontSize(8)
             .setFontColor("#475569")
             .setVerticalAlignment("middle");
    tempSheet.getRange(requestStartRow + 2, 5)
             .setValue(firstRow[17] || "Soft") // Column R (Report Copy)
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle")
             .setFontSize(9);
             
    var reqTableRange = tempSheet.getRange(requestStartRow, 1, 3, 5);
    reqTableRange.setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
    
    currentRow = requestStartRow + 3;
    
    // Signatures
    currentRow += 4;
    var customerSigLine = tempSheet.getRange(currentRow, 1, 1, 2).merge().setValue("Customer Signature").setHorizontalAlignment("center").setFontWeight("bold").setFontSize(9).setFontColor("#475569");
    var authSigLine = tempSheet.getRange(currentRow, 4, 1, 2).merge().setValue("Authorized Signatory").setHorizontalAlignment("center").setFontWeight("bold").setFontSize(9).setFontColor("#475569");
    
    customerSigLine.setBorder(true, null, null, null, null, null, "#475569", SpreadsheetApp.BorderStyle.SOLID);
    authSigLine.setBorder(true, null, null, null, null, null, "#475569", SpreadsheetApp.BorderStyle.SOLID);
    
    var totalHeightRows = currentRow + 2;
    tempSheet.setRowHeight(totalHeightRows, 15);
    
    var outerBorderRange = tempSheet.getRange(1, 1, totalHeightRows, 5);
    outerBorderRange.setBorder(true, true, true, true, null, null, "#1E293B", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    ss.setActiveSheet(tempSheet);
    
    var ui = SpreadsheetApp.getUi();
    ui.alert("Receipt Generated", "Receipt Slip for " + receiptNo + " has been successfully generated in tab '" + tempSheetName + "'.\n\nPlease go to this tab and press Ctrl + P to print.", ui.ButtonSet.OK);
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Print Error", "Could not generate print slip: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Placeholder alert display helper for unimplemented custom menu items.
 */
function showPlaceholderAlert(actionName) {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Coming soon", "The feature '" + actionName + "' is coming soon in the next phase.", ui.ButtonSet.OK);
  } catch (e) {
    Logger.log("Error displaying alert: " + e.toString());
  }
}

function menuRefreshDashboard() {
  showPlaceholderAlert("Refresh Dashboard");
}
