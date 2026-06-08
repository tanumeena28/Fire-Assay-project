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
    var html = HtmlService.createHtmlOutputFromFile('InventoryDialog')
        .setWidth(900)
        .setHeight(700);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Inventory Desk');
  } catch (e) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("UI Error", "Could not open Inventory Desk dialog: " + e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Checks and creates the background inventory sheets, seeding initial rows if empty.
 * 
 * @param {SpreadsheetApp.Spreadsheet} ss The active spreadsheet.
 */
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
  if (invLastRow === 0) {
    var headers = ["Item ID", "Item Name", "Current Stock", "Unit", "Min Alert Level", "Last Updated"];
    invSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatSheetCommon(invSheet, headers.length, "#8B5CF6"); // Purple tab color
    
    // Seed initial items (Gold, Silver, Lead, Nickel)
    var seedItems = [
      ["INV-01", "Gold", 0.0, "g", 10.0, ""],
      ["INV-02", "Silver", 0.0, "g", 50.0, ""],
      ["INV-03", "Lead", 0.0, "g", 1000.0, ""],
      ["INV-04", "Nickel", 0.0, "g", 100.0, ""]
    ];
    invSheet.getRange(2, 1, seedItems.length, headers.length).setValues(seedItems);
    invSheet.getRange(2, 3, seedItems.length, 1).setNumberFormat("0.00");
    invSheet.getRange(2, 5, seedItems.length, 1).setNumberFormat("0.00");
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
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var list = [];
    
    for (var i = 1; i < data.length; i++) {
      var itemName = data[i][1] ? data[i][1].toString().trim() : "";
      if (itemName) {
        var currentStock = parseFloat(data[i][2]) || 0;
        var minAlert = parseFloat(data[i][4]) || 0;
        var unit = data[i][3] ? data[i][3].toString().trim() : "";
        var lastUpdated = data[i][5] ? data[i][5].toString() : "";
        
        list.push({
          itemId: data[i][0] ? data[i][0].toString() : "",
          itemName: itemName,
          currentStock: currentStock,
          unit: unit,
          minAlert: minAlert,
          lastUpdated: lastUpdated,
          isLow: currentStock < minAlert
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
