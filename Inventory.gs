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
  
  var txSheet = ss.getSheetByName("Inventory_Transactions");
  if (!txSheet) {
    txSheet = ss.insertSheet("Inventory_Transactions");
  }
  
  // 1. Setup Inventory Sheet Headers
  var invLastRow = invSheet.getLastRow();
  if (invLastRow <= 1) {
    invSheet.clear();
    var headers = ["Item ID", "Item Name", "Current Stock", "Unit", "Min Alert Level", "Last Updated"];
    invSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatSheetCommon(invSheet, headers.length, "#8B5CF6"); // Purple tab color
    
    // Seed initial items (Gold, Silver, Lead, Copper, Nickel) with new alert thresholds
    var seedItems = [
      ["INV-01", "Gold", 0.0, "g", 5.0, ""],
      ["INV-02", "Silver", 0.0, "g", 20.0, ""],
      ["INV-03", "Lead", 0.0, "g", 500.0, ""],
      ["INV-04", "Copper", 0.0, "g", 1.0, ""],
      ["INV-05", "Nickel", 0.0, "g", 1.0, ""]
    ];
    invSheet.getRange(2, 1, seedItems.length, headers.length).setValues(seedItems);
    invSheet.getRange(2, 3, seedItems.length, 1).setNumberFormat("0.00");
    invSheet.getRange(2, 5, seedItems.length, 1).setNumberFormat("0.00");
  } else {
    // Check if Copper is present, append if missing
    var data = invSheet.getDataRange().getValues();
    var hasCopper = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim().toLowerCase() === "copper") {
        hasCopper = true;
        break;
      }
    }
    if (!hasCopper) {
      var nextId = "INV-0" + (data.length);
      var copperRow = [nextId, "Copper", 0.0, "g", 1.0, ""];
      var nextRow = invSheet.getLastRow() + 1;
      invSheet.getRange(nextRow, 1, 1, copperRow.length).setValues([copperRow]);
      invSheet.getRange(nextRow, 3).setNumberFormat("0.00");
      invSheet.getRange(nextRow, 5).setNumberFormat("0.00");
    }
    
    // Update alert levels dynamically for existing items
    var updatedLimits = {
      "gold": 5.0,
      "silver": 20.0,
      "lead": 500.0,
      "nickel": 1.0,
      "copper": 1.0
    };
    for (var i = 1; i < data.length; i++) {
      var name = data[i][1] ? data[i][1].toString().trim().toLowerCase() : "";
      if (updatedLimits[name] !== undefined) {
        invSheet.getRange(i + 1, 5).setValue(updatedLimits[name]);
      }
    }
  }
  
  // 2. Setup Inventory Transactions Sheet Headers
  var txLastRow = txSheet.getLastRow();
  if (txLastRow === 0) {
    var txHeaders = ["Transaction ID", "Date", "Time", "Item Name", "Type", "Quantity", "Reference / Reason", "Staff Name"];
    txSheet.getRange(1, 1, 1, txHeaders.length).setValues([txHeaders]);
    formatSheetCommon(txSheet, txHeaders.length, "#EC4899"); // Pink tab color
  }
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
    var txSheet = ss.getSheetByName("Inventory_Transactions");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    
    // Fetch transaction data to compute totals
    var txData = [];
    if (txSheet) {
      txData = txSheet.getDataRange().getValues();
    }
    
    // Compute total purchased (Stock In) and total consumed (Stock Out) per item
    var totals = {};
    for (var j = 1; j < txData.length; j++) {
      if (!txData[j]) continue;
      var item = txData[j][3] ? txData[j][3].toString().trim() : "";
      var type = txData[j][4] ? txData[j][4].toString().trim() : "";
      var qty = parseFloat(txData[j][5]) || 0;
      
      if (!totals[item]) {
        totals[item] = { purchased: 0, consumed: 0 };
      }
      
      if (type === "Stock In") {
        totals[item].purchased += qty;
      } else if (type === "Stock Out") {
        totals[item].consumed += qty;
      }
    }
    
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var itemName = data[i][1] ? data[i][1].toString().trim() : "";
      if (itemName) {
        var currentStock = parseFloat(data[i][2]) || 0;
        var minAlert = parseFloat(data[i][4]) || 0;
        var unit = data[i][3] ? data[i][3].toString().trim() : "";
        var lastUpdated = data[i][5] ? data[i][5].toString() : "";
        
        var itemTotals = totals[itemName] || { purchased: 0, consumed: 0 };
        
        list.push({
          itemId: data[i][0] ? data[i][0].toString() : "",
          itemName: itemName,
          currentStock: currentStock,
          unit: unit,
          minAlert: minAlert,
          lastUpdated: lastUpdated,
          isLow: currentStock < minAlert,
          totalPurchased: itemTotals.purchased,
          totalConsumed: itemTotals.consumed
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
    var sheet = ss.getSheetByName("Inventory_Transactions");
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    var startRow = Math.max(2, lastRow - 14); // Fetch up to 15 recent entries
    var numRows = lastRow - startRow + 1;
    
    var data = sheet.getRange(startRow, 1, numRows, 8).getValues();
    var list = [];
    
    // Iterate backwards to show newest first
    for (var i = data.length - 1; i >= 0; i--) {
      var dateVal = data[i][1];
      var dateStr = "";
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        dateStr = dateVal ? dateVal.toString() : "";
      }
      
      var timeVal = data[i][2];
      var timeStr = "";
      if (timeVal instanceof Date) {
        timeStr = Utilities.formatDate(timeVal, Session.getScriptTimeZone(), "HH:mm:ss");
      } else {
        timeStr = timeVal ? timeVal.toString() : "";
      }
      
      list.push({
        txId: data[i][0] ? data[i][0].toString() : "",
        date: dateStr,
        time: timeStr,
        itemName: data[i][3] ? data[i][3].toString() : "",
        type: data[i][4] ? data[i][4].toString() : "",
        qty: parseFloat(data[i][5]) || 0,
        reason: data[i][6] ? data[i][6].toString() : "",
        staff: data[i][7] ? data[i][7].toString() : ""
      });
    }
    return list;
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
function adjustInventory(itemName, type, qty, reason, staffName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = ss.getSheetByName("Inventory");
    var txSheet = ss.getSheetByName("Inventory_Transactions");
    
    if (!invSheet || !txSheet) {
      throw new Error("Inventory sheets are not initialized. Please run project setup.");
    }
    
    var amount = parseFloat(qty);
    if (isNaN(amount) || amount <= 0) {
      throw new Error("Quantity must be a valid positive number.");
    }
    
    // 1. Find the target item in the Inventory sheet
    var invData = invSheet.getDataRange().getValues();
    var rowIndex = -1;
    var currentStock = 0;
    var unit = "g";
    var minAlert = 0;
    
    for (var i = 1; i < invData.length; i++) {
      if (invData[i][1] && invData[i][1].toString().trim().toLowerCase() === itemName.trim().toLowerCase()) {
        rowIndex = i + 1; // 1-indexed row number
        currentStock = parseFloat(invData[i][2]) || 0;
        unit = invData[i][3] ? invData[i][3].toString() : "g";
        minAlert = parseFloat(invData[i][4]) || 0;
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error("Item '" + itemName + "' not found in Inventory Master.");
    }
    
    // 2. Calculate new stock level
    var newStock = currentStock;
    if (type === "Stock In") {
      newStock += amount;
    } else if (type === "Stock Out") {
      if (currentStock < amount) {
        throw new Error("Insufficient stock. Only " + currentStock + " " + unit + " available, but " + amount + " " + unit + " requested.");
      }
      newStock -= amount;
    } else {
      throw new Error("Invalid transaction type: " + type);
    }
    
    var now = new Date();
    var timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    // 3. Update stock balance in Inventory sheet
    invSheet.getRange(rowIndex, 3).setValue(newStock);
    invSheet.getRange(rowIndex, 6).setValue(timestampStr);
    
    // 4. Record entry in Inventory_Transactions sheet
    var txId = "TX-" + Utilities.formatDate(now, Session.getScriptTimeZone(), "yyMMddHHmmss");
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");
    
    var txRow = [
      txId,
      dateStr,
      timeStr,
      itemName,
      type,
      amount,
      reason || "N/A",
      staffName || "System"
    ];
    
    var nextTxRow = txSheet.getLastRow() + 1;
    txSheet.getRange(nextTxRow, 1, 1, txRow.length).setValues([txRow]);
    
    // Format transaction columns
    txSheet.getRange(nextTxRow, 1).setNumberFormat("@"); // Text ID
    txSheet.getRange(nextTxRow, 6).setNumberFormat("0.00"); // Qty decimal
    
    SpreadsheetApp.flush();
    return "Stock successfully updated for " + itemName + ". New Stock: " + newStock.toFixed(2) + " " + unit;
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
