/**
 * Gold & Silver Assay Laboratory Management System (LMS)
 * Client: NCH Services Pvt. Ltd., Surat, Gujarat
 * Phase 3: Quality Manager Review + Job Card Creation (JobCardCreation.gs)
 *
 * This file handles Job Card generation. It checks the QC Review status,
 * inserts Job Cards for accepted samples, formats columns, and applies 
 * dropdown validation rules.
 */

function menuNewJobCard() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('JobCardCreationForm')
        .setWidth(850)
        .setHeight(550);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Create Job Cards');
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("UI Error", "Could not open Job Card creator dialog: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Backend API: Returns all accepted samples from QC_Review for a receipt that do not have Job Cards yet.
 * 
 * @param {string} receiptNo The Receipt Number to scan.
 * @return {Array<Object>} List of accepted samples details.
 */
function getAcceptedSamplesForReceipt(receiptNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var receivingSheet = ss.getSheetByName("Receiving");
    var jobCardSheet = ss.getSheetByName("Job_Cards");
    if (!receivingSheet || !jobCardSheet) return [];
    
    // Get list of existing Job Cards to prevent showing duplicate options
    var jcData = jobCardSheet.getDataRange().getValues();
    var existingCards = {};
    for (var i = 1; i < jcData.length; i++) {
      var jcNo = jcData[i][0].toString().trim();
      if (jcNo) {
        existingCards[jcNo] = true;
      }
    }
    
    var data = receivingSheet.getDataRange().getValues();
    var samples = [];
    
    for (var i = 1; i < data.length; i++) {
      var subJobNo = data[i][0].toString().trim(); // Column A: Sub-Job No
      var custName = data[i][3].toString().trim(); // Column D: Customer Name
      var material = data[i][7].toString().trim(); // Column H: Material Type
      var weight = parseFloat(data[i][9]) || 0;    // Column J: Total Weight (g)
      var rNo = data[i][14].toString().trim();      // Column O: Receipt No
      
      if (rNo === receiptNo && !existingCards[subJobNo]) {
        samples.push({
          subJobNo: subJobNo,
          customerName: custName,
          material: material,
          weight: weight
        });
      }
    }
    return samples;
  } catch (e) {
    Logger.log("Error in getAcceptedSamplesForReceipt: " + e.toString());
    return [];
  }
}

/**
 * Backend API: Creates Job Card entries on the "Job_Cards" sheet for a given Receipt.
 * populating the custom Declared Fineness and technician.
 * 
 * @param {Object} formData Submitted values containing receiptNo and samples array.
 * @return {string} Status confirmation message detailing number of cards generated.
 */
function saveJobCards(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var jobCardSheet = ss.getSheetByName("Job_Cards");
    var receivingSheet = ss.getSheetByName("Receiving");
    
    if (!jobCardSheet || !receivingSheet) {
      throw new Error("Job_Cards or Receiving sheet not found.");
    }
    
    var receiptNo = formData.receiptNo;
    var samples = formData.samples || [];
    
    // Fetch Receiving sheet records to map receipt date
    var receivingData = receivingSheet.getDataRange().getValues();
    var receivingDatesMap = {};
    for (var i = 1; i < receivingData.length; i++) {
      var subJob = receivingData[i][0].toString().trim();
      var rDate = receivingData[i][1];
      if (subJob) {
        receivingDatesMap[subJob] = rDate;
      }
    }
    
    var rowsToAppend = [];
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    for (var i = 0; i < samples.length; i++) {
      var s = samples[i];
      var subJobNo = s.subJobNo;
      
      var receiptDate = receivingDatesMap[subJobNo] || new Date();
      if (receiptDate instanceof Date) {
        receiptDate = Utilities.formatDate(receiptDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      
      // Job_Cards Columns:
      // A: Job Card No, B: Receipt No, C: Receipt Date, D: Customer Name, E: Material Type,
      // F: Declared Fineness, G: Weight Received (g), H: XRF Au Reading (ppt), I: Ag XRF Reading (ppt),
      // J: Assigned To, K: Status, L: Created Date, M: Notes
      rowsToAppend.push([
        subJobNo,                     // A
        receiptNo,                    // B
        receiptDate,                  // C
        s.customerName,               // D
        s.material,                   // E
        s.declaredFineness,           // F
        parseFloat(s.weight) || 0,    // G
        0.0,                          // H (XRF Au Reading ppt) - populated in XRF step
        0.0,                          // I (Ag XRF Reading ppt) - populated in XRF step
        s.assignedTo,                 // J
        "Pending",                    // K
        todayStr,                     // L
        ""                            // M
      ]);
    }
    
    if (rowsToAppend.length > 0) {
      var startRow = jobCardSheet.getLastRow() + 1;
      var range = jobCardSheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length);
      
      var formats = [];
      for (var r = 0; r < rowsToAppend.length; r++) {
        var formatRow = [];
        for (var c = 0; c < rowsToAppend[0].length; c++) {
          if (c === 0 || c === 1) {
            formatRow.push("@"); 
          } else {
            formatRow.push("");
          }
        }
        formats.push(formatRow);
      }
      range.setNumberFormats(formats);
      range.setValues(rowsToAppend);
      
      // Apply data validations dynamically
      applyJobCardsValidation(startRow, rowsToAppend.length);
      return rowsToAppend.length + " Job Card(s) successfully created for receipt " + receiptNo + ".";
    } else {
      return "No new job cards to generate.";
    }
  } catch (e) {
    throw new Error("Failed to save Job Cards: " + e.toString());
  }
}

/**
 * Applies Google Sheets Data Validation rules to newly appended Job Cards.
 * 
 * @param {number} startRow First row index to apply rules.
 * @param {number} numRows Count of rows to apply rules to.
 */
function applyJobCardsValidation(startRow, numRows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Job_Cards");
  if (!sheet) return;
  
  // Column F is Declared Fineness (6th column)
  var finenessRange = sheet.getRange(startRow, 6, numRows, 1);
  var finenessRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["375", "585", "750", "916", "999", "999.9"], true)
    .setAllowInvalid(false)
    .setHelpText("Select a valid declared fineness.")
    .build();
  finenessRange.setDataValidation(finenessRule);
  
  // Column K is Status (11th column)
  var statusRange = sheet.getRange(startRow, 11, numRows, 1);
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending", "In Progress", "Completed"], true)
    .setAllowInvalid(false)
    .setHelpText("Select a valid status.")
    .build();
  statusRange.setDataValidation(statusRule);
}

/**
 * Formats columns and headers in the 'Job_Cards' sheet.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss The active spreadsheet.
 */
function setupJobCardsSheet(ss) {
  var jobCardSheet = ss.getSheetByName("Job_Cards");
  if (!jobCardSheet) return;
  
  var lastRow = jobCardSheet.getLastRow();
  if (lastRow > 0) {
    // Already initialized. Skip.
    return;
  }
  
  var headers = [
    "Job Card No",        // A (Sub-Jobcard No)
    "Receipt No",         // B
    "Receipt Date",       // C
    "Customer Name",      // D
    "Material Type",      // E
    "Declared Fineness",  // F
    "Weight Received (g)",// G
    "XRF Au Reading (ppt)", // H
    "Ag XRF Reading (ppt)", // I
    "Assigned To",        // J (Technician)
    "Status",             // K (Pending/In Progress/Completed)
    "Created Date",       // L
    "Notes"               // M
  ];
  
  jobCardSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatSheetCommon(jobCardSheet, headers.length, null); // Headers styled via formatSheetCommon, no custom tab color requested
}
