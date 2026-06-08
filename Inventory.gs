/**
 * Gold & Silver Assay Laboratory Management System (LMS)
 * Client: NCH Services Pvt. Ltd., Surat, Gujarat
 * Inventory Management Module (Inventory.gs)
 *
 * Handles backend transactions, stock checks, auto-seeding,
 * and dialog rendering for Gold, Silver, Lead, and Nickel consumables.
 */

/**
 * Menu Callback: Opens the "Inventory Desk" HTML modal dialog.
 */
function menuInventory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    setupInventorySheets(ss);
    
    var html = HtmlService.createHtmlOutputFromFile('InventoryDialog')
        .setWidth(900)
        .setHeight(700);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Inventory Desk');
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("UI Error", "Could not open Inventory Desk dialog: " + e.toString(), ui.ButtonSet.OK);
  }
}

function setupInventorySheets(ss) {
  var invSheet = ss.getSheetByName("Inventory");
  if (!invSheet) {
    invSheet = ss.insertSheet("Inventory");
  }
  
  var metals = ["Gold", "Silver", "Lead", "Copper", "Nickel"];
  var reportHeaders = ["Date", "Time", "Type", "Quantity", "Job Card / Receipt No", "Reference / Reason", "Staff Name"];
  
  // 1. Setup individual report sheets for each metal
  for (var i = 0; i < metals.length; i++) {
    var sheetName = metals[i] + "_Report";
    var sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
      sh.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]);
      formatSheetCommon(sh, reportHeaders.length, "#10B981"); // Green tab color
    }
  }
  
  // 2. Migrate transactions from old combined transaction sheet if it exists
  var oldTxSheet = ss.getSheetByName("Inventory_Transactions");
  if (oldTxSheet) {
    var oldTxLastRow = oldTxSheet.getLastRow();
    if (oldTxLastRow > 1) {
      var oldTxData = oldTxSheet.getRange(2, 1, oldTxLastRow - 1, 8).getValues();
      for (var k = 0; k < oldTxData.length; k++) {
        var row = oldTxData[k];
        var date = row[1];
        var time = row[2];
        var itemName = row[3] ? row[3].toString().trim() : "";
        var type = row[4];
        var qty = row[5];
        var reason = row[6];
        var staff = row[7];
        
        var matchItem = "";
        for (var m = 0; m < metals.length; m++) {
          if (metals[m].toLowerCase() === itemName.toLowerCase()) {
            matchItem = metals[m];
            break;
          }
        }
        
        if (matchItem) {
          var destSheet = ss.getSheetByName(matchItem + "_Report");
          if (destSheet) {
            var destData = destSheet.getDataRange().getValues();
            var alreadyMigrated = false;
            for (var d = 1; d < destData.length; d++) {
              if (destData[d][0].toString() === date.toString() &&
                  destData[d][1].toString() === time.toString() &&
                  destData[d][2] === type &&
                  parseFloat(destData[d][3]) === parseFloat(qty) &&
                  destData[d][5] === reason) {
                alreadyMigrated = true;
                break;
              }
            }
            if (!alreadyMigrated) {
              destSheet.appendRow([date, time, type, qty, "N/A", reason, staff]);
              destSheet.getRange(destSheet.getLastRow(), 4).setNumberFormat("0.00");
            }
          }
        }
      }
    }
    try {
      ss.deleteSheet(oldTxSheet);
    } catch (e) {
      Logger.log("Could not delete old transactions sheet: " + e.toString());
    }
  }
  
  // 3. Setup Inventory Summary Sheet Headers & dynamic SUMIF formulas
  var headers = ["Item ID", "Item Name", "Current Stock", "Unit", "Total Purchased", "Total Consumed", "Min Alert Level", "Last Updated"];
  
  // Preserve any existing user-configured Min Alert Level or Last Updated values
  var oldData = invSheet.getDataRange().getValues();
  var minAlertMap = {
    "Gold": 5.0,
    "Silver": 20.0,
    "Lead": 500.0,
    "Copper": 1.0,
    "Nickel": 1.0
  };
  var lastUpdatedMap = {
    "Gold": "",
    "Silver": "",
    "Lead": "",
    "Copper": "",
    "Nickel": ""
  };
  
  if (oldData.length > 1) {
    for (var r = 1; r < oldData.length; r++) {
      var name = oldData[r][1] ? oldData[r][1].toString().trim() : "";
      if (name) {
        if (oldData[r].length > 4 && oldData[r][4] !== undefined && oldData[r][4] !== "") {
          var val = parseFloat(oldData[r][4]);
          if (!isNaN(val)) minAlertMap[name] = val;
        }
        if (oldData[r].length > 6 && oldData[r][6] !== undefined && oldData[r][6] !== "") {
          var val = parseFloat(oldData[r][6]);
          if (!isNaN(val)) minAlertMap[name] = val;
        }
        if (oldData[r].length > 5 && oldData[r][5] !== undefined && oldData[r][5] !== "") {
          lastUpdatedMap[name] = oldData[r][5].toString();
        }
        if (oldData[r].length > 7 && oldData[r][7] !== undefined && oldData[r][7] !== "") {
          lastUpdatedMap[name] = oldData[r][7].toString();
        }
      }
    }
  }
  
  // Overwrite headers and rebuild formulas
  invSheet.clear();
  invSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatSheetCommon(invSheet, headers.length, "#8B5CF6"); // Purple tab color
  
  var seedRows = [
    ["INV-01", "Gold",   "=E2-F2", "g", "=SUMIF(Gold_Report!C:C, \"Stock In\", Gold_Report!D:D)",   "=SUMIF(Gold_Report!C:C, \"Stock Out\", Gold_Report!D:D)",   minAlertMap["Gold"],   lastUpdatedMap["Gold"]],
    ["INV-02", "Silver", "=E3-F3", "g", "=SUMIF(Silver_Report!C:C, \"Stock In\", Silver_Report!D:D)", "=SUMIF(Silver_Report!C:C, \"Stock Out\", Silver_Report!D:D)", minAlertMap["Silver"], lastUpdatedMap["Silver"]],
    ["INV-03", "Lead",   "=E4-F4", "g", "=SUMIF(Lead_Report!C:C, \"Stock In\", Lead_Report!D:D)",   "=SUMIF(Lead_Report!C:C, \"Stock Out\", Lead_Report!D:D)",   minAlertMap["Lead"],   lastUpdatedMap["Lead"]],
    ["INV-04", "Copper", "=E5-F5", "g", "=SUMIF(Copper_Report!C:C, \"Stock In\", Copper_Report!D:D)", "=SUMIF(Copper_Report!C:C, \"Stock Out\", Copper_Report!D:D)", minAlertMap["Copper"], lastUpdatedMap["Copper"]],
    ["INV-05", "Nickel", "=E6-F6", "g", "=SUMIF(Nickel_Report!C:C, \"Stock In\", Nickel_Report!D:D)", "=SUMIF(Nickel_Report!C:C, \"Stock Out\", Nickel_Report!D:D)", minAlertMap["Nickel"], lastUpdatedMap["Nickel"]]
  ];
  
  invSheet.getRange(2, 1, seedRows.length, headers.length).setValues(seedRows);
  invSheet.getRange(2, 3, seedRows.length, 1).setNumberFormat("0.00");
  invSheet.getRange(2, 5, seedRows.length, 2).setNumberFormat("0.00");
  invSheet.getRange(2, 7, seedRows.length, 1).setNumberFormat("0.00");
}

/**
 * Backend API: Returns current stock balance for Gold, Silver, Lead, and Nickel.
 * 
 * @return {Array<Object>} List of inventory items with alert triggers.
 */
function getInventoryData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Inventory");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var itemName = data[i][1] ? data[i][1].toString().trim() : "";
      if (itemName) {
        var currentStock = parseFloat(data[i][2]) || 0; // Column C
        var minAlert = parseFloat(data[i][6]) || 0; // Column G
        var unit = data[i][3] ? data[i][3].toString().trim() : ""; // Column D
        var totalPurchased = parseFloat(data[i][4]) || 0; // Column E
        var totalConsumed = parseFloat(data[i][5]) || 0; // Column F
        var lastUpdated = data[i][7] ? data[i][7].toString() : ""; // Column H
        
        list.push({
          itemId: data[i][0] ? data[i][0].toString() : "",
          itemName: itemName,
          currentStock: currentStock,
          unit: unit,
          minAlert: minAlert,
          lastUpdated: lastUpdated,
          isLow: currentStock < minAlert,
          totalPurchased: totalPurchased,
          totalConsumed: totalConsumed
        });
      }
    }
    return list;
  } catch (e) {
    Logger.log("Error in getInventoryData: " + e.toString());
    return [];
  }
}

/**
 * Backend API: Returns the last 15 transaction log rows.
 * 
 * @return {Array<Object>} Transaction history.
 */
function getTransactionHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var metals = ["Gold", "Silver", "Lead", "Copper", "Nickel"];
    var allTx = [];
    
    for (var m = 0; m < metals.length; m++) {
      var sheetName = metals[m] + "_Report";
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) continue;
      
      var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var dateVal = row[0];
        var dateStr = "";
        if (dateVal instanceof Date) {
          dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          dateStr = dateVal ? dateVal.toString() : "";
        }
        
        var timeVal = row[1];
        var timeStr = "";
        if (timeVal instanceof Date) {
          timeStr = Utilities.formatDate(timeVal, Session.getScriptTimeZone(), "HH:mm:ss");
        } else {
          timeStr = timeVal ? timeVal.toString() : "";
        }
        
        var qty = parseFloat(row[3]) || 0;
        var jobCardNo = row[4] ? row[4].toString() : "";
        var reason = row[5] ? row[5].toString() : "";
        var staff = row[6] ? row[6].toString() : "";
        
        var timestamp = 0;
        try {
          var dtStr = dateStr + "T" + (timeStr || "00:00:00");
          timestamp = Date.parse(dtStr) || 0;
        } catch (e) {}
        
        allTx.push({
          date: dateStr,
          time: timeStr,
          itemName: metals[m],
          type: row[2] ? row[2].toString() : "",
          qty: qty,
          jobCardNo: jobCardNo,
          reason: reason,
          staff: staff,
          timestamp: timestamp
        });
      }
    }
    
    allTx.sort(function(a, b) {
      if (b.timestamp !== a.timestamp) {
        return b.timestamp - a.timestamp;
      }
      return b.date.localeCompare(a.date) || b.time.localeCompare(a.time);
    });
    
    return allTx.slice(0, 15);
  } catch (e) {
    Logger.log("Error in getTransactionHistory: " + e.toString());
    return [];
  }
}

/**
 * Backend API: Adjusts inventory stock count and logs transaction audit trail.
 * 
 * @param {string} itemName Name of target item ("Gold", "Silver", "Lead", "Nickel").
 * @param {string} type Adjustment type ("Stock In" or "Stock Out").
 * @param {number} qty Amount to add/remove.
 * @param {string} reason Reference reason for adjustment.
 * @param {string} staffName Authorized staff member.
 * @return {string} Success statement message.
 */
function adjustInventory(itemName, type, qty, jobCardNo, reason, staffName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = ss.getSheetByName("Inventory");
    
    // Handle backwards compatibility signature variations
    if (staffName === undefined) {
      staffName = reason;
      reason = jobCardNo;
      jobCardNo = "N/A";
    }
    
    if (!invSheet) {
      throw new Error("Inventory sheets are not initialized. Please run project setup.");
    }
    
    var amount = parseFloat(qty);
    if (isNaN(amount) || amount <= 0) {
      throw new Error("Quantity must be a valid positive number.");
    }
    
    var targetSheetName = itemName.trim() + "_Report";
    var metalSheet = ss.getSheetByName(targetSheetName);
    if (!metalSheet) {
      throw new Error("Metal report sheet '" + targetSheetName + "' not found.");
    }
    
    // 1. Find the target item in the Inventory summary sheet to check stock level
    var invData = invSheet.getDataRange().getValues();
    var rowIndex = -1;
    var currentStock = 0;
    var unit = "g";
    
    for (var i = 1; i < invData.length; i++) {
      if (invData[i][1] && invData[i][1].toString().trim().toLowerCase() === itemName.trim().toLowerCase()) {
        rowIndex = i + 1; // 1-indexed row number
        currentStock = parseFloat(invData[i][2]) || 0;
        unit = invData[i][3] ? invData[i][3].toString() : "g";
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error("Item '" + itemName + "' not found in Inventory Summary.");
    }
    
    if (type === "Stock Out" && currentStock < amount) {
      throw new Error("Insufficient stock. Only " + currentStock.toFixed(2) + " " + unit + " available, but " + amount.toFixed(2) + " " + unit + " requested.");
    }
    
    var now = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");
    var timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    // 2. Append entry to specific metal's report sheet
    var txRow = [
      dateStr,
      timeStr,
      type,
      amount,
      jobCardNo || "N/A",
      reason || "N/A",
      staffName || "System"
    ];
    
    metalSheet.appendRow(txRow);
    metalSheet.getRange(metalSheet.getLastRow(), 4).setNumberFormat("0.00"); // Format Quantity
    
    // 3. Update Last Updated field in the Inventory Summary sheet
    invSheet.getRange(rowIndex, 8).setValue(timestampStr); // Column H
    
    SpreadsheetApp.flush();
    
    // Read the formula recalculated stock level
    var updatedStock = parseFloat(invSheet.getRange(rowIndex, 3).getValue()) || 0;
    
    return "Stock successfully updated for " + itemName + ". New Stock: " + updatedStock.toFixed(2) + " " + unit;
  } catch (e) {
    throw new Error(e.toString());
  }
}

/**
 * Backend API: Returns the list of last 10 unique batch sheet names from Fire_Assay_Sheet.
 * 
 * @return {Array<string>} Unique batch sheet names.
 */
function getRecentBatches() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var faSheet = ss.getSheetByName("Fire_Assay_Sheet");
    if (!faSheet) return [];
    
    var data = faSheet.getDataRange().getValues();
    var batches = [];
    var seen = {};
    
    // Read from bottom to top to get the most recent batches first
    for (var i = data.length - 1; i >= 1; i--) {
      var batchName = data[i][0] ? data[i][0].toString().trim() : "";
      if (batchName && !seen[batchName]) {
        seen[batchName] = true;
        batches.push(batchName);
        if (batches.length >= 10) break;
      }
    }
    return batches;
  } catch (e) {
    Logger.log("Error in getRecentBatches: " + e.toString());
    return [];
  }
}

/**
 * Backend API: Fetches the number of samples in a Fire Assay batch
 * and returns calculated consumable consumption values.
 * 
 * @param {string} batchName The Batch Sheet Name (e.g. "08062026").
 * @return {Object} Calculated consumption weights for Gold, Silver, Lead, Nickel.
 */
function getBatchConsumptionDetails(batchName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var faSheet = ss.getSheetByName("Fire_Assay_Sheet");
    if (!faSheet) {
      throw new Error("Fire_Assay_Sheet not found.");
    }
    
    var data = faSheet.getDataRange().getValues();
    var sampleCount = 0;
    var cgCount = 0;
    
    for (var i = 1; i < data.length; i++) {
      var rowBatch = data[i][0] ? data[i][0].toString().trim() : "";
      if (rowBatch === batchName) {
        var isCG = data[i][13] === true || data[i][13] === "true"; // CG Check Gold column N
        if (isCG) {
          cgCount++;
        } else {
          sampleCount++;
        }
      }
    }
    
    // Each sample row represents one cupellation (2 readings per sub-job)
    // Total cupellations in this batch = sampleCount + cgCount
    var totalCupellations = sampleCount + cgCount;
    var distinctSamples = sampleCount / 2; // Since there are 2 readings per sample
    
    // Default consumption formulas:
    // Lead: 4 grams of lead foil per cupellation
    var leadConsumed = totalCupellations * 4.0;
    
    // Silver: 0.6 grams per cupellation (inquation)
    var silverConsumed = totalCupellations * 0.6;
    
    // Gold: Check Gold standard uses 250mg (0.25g) per check gold cupellation
    var goldConsumed = cgCount * 0.25;
    
    // Nickel: Typically not used or a small default (e.g. 0)
    var nickelConsumed = 0.0;
    
    // Copper: Typically not used or a small default (e.g. 0)
    var copperConsumed = 0.0;
    
    return {
      batchName: batchName,
      totalCupellations: totalCupellations,
      distinctSamples: distinctSamples,
      cgCount: cgCount,
      gold: goldConsumed,
      silver: silverConsumed,
      lead: leadConsumed,
      nickel: nickelConsumed,
      copper: copperConsumed
    };
  } catch (e) {
    throw new Error("Failed to fetch batch details: " + e.toString());
  }
}
