// ============================================================================
// STOCK TOOLS.GS - Stock Tools for Moving Skids Between Staging Areas
// WITH STOCK_TOTALS INTEGRATION
// ============================================================================

// Hardcoded column positions (0-indexed) - matches Bin_Stock/Floor_Stock_Levels/Inbound_Staging layout
// A=Bin_Code, B=Bin_Name, C=Push_Number, D=FBPN, E=Manufacturer, F=Project, 
// G=Initial_Qty, H=Current_Qty, I=Stock_Pct, J=AUDIT, K=Skid_ID, L=SKU
const STOCK_COLS = {
  BIN_CODE: 0,
  BIN_NAME: 1,
  PUSH_NUMBER: 2,
  FBPN: 3,
  MANUFACTURER: 4,
  PROJECT: 5,
  INITIAL_QTY: 6,
  CURRENT_QTY: 7,
  STOCK_PCT: 8,
  AUDIT: 9,
  SKID_ID: 10,
  SKU: 11
};

/**
 * Open the Stock Tools Modal
 */
function openStockToolsModal() {
  const html = HtmlService.createTemplateFromFile('StockToolsModal')
    .evaluate()
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Stock Tools - Move Skids');
}

/**
 * Get all bins with inventory from a specific tab
 * Returns bins that have at least one FBPN (not empty)
 */
function getBinsWithInventory(tabName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(tabName);

    if (!sheet) {
      throw new Error(`Sheet "${tabName}" not found`);
    }

    const data = sheet.getDataRange().getValues();
    const bins = [];

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fbpn = String(row[STOCK_COLS.FBPN] || '');
      const qty = Number(row[STOCK_COLS.CURRENT_QTY] || 0);

      // Only include rows with FBPN (has inventory)
      if (fbpn || qty > 0) {
        bins.push({
          binCode: row[STOCK_COLS.BIN_CODE],
          fbpn: fbpn,
          qty: qty,
          skidId: row[STOCK_COLS.SKID_ID],
          manufacturer: row[STOCK_COLS.MANUFACTURER],
          project: row[STOCK_COLS.PROJECT],
          pushNumber: row[STOCK_COLS.PUSH_NUMBER]
        });
      }
    }
    return bins;

  } catch (error) {
    Logger.log('Error in getBinsWithInventory: ' + error.toString());
    throw error;
  }
}

/**
 * Get destination bins (ONLY EMPTY ones)
 * Filter: Current_Quantity (Col H) is empty or 0 AND FBPN (Col D) is empty
 */
function getDestinationBins(tabName, emptyOnly) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(tabName);

    if (!sheet) {
      throw new Error(`Sheet "${tabName}" not found`);
    }

    const data = sheet.getDataRange().getValues();
    const bins = [];

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const binCode = String(row[STOCK_COLS.BIN_CODE]);
      if (!binCode) continue;

      const currentQty = Number(row[STOCK_COLS.CURRENT_QTY] || 0);
      const fbpn = String(row[STOCK_COLS.FBPN] || '');

      // Strict check for "Empty": Qty <= 0 AND no FBPN
      const isEmpty = (currentQty <= 0 && fbpn === '');

      if (emptyOnly) {
        if (isEmpty) bins.push(binCode);
      } else {
        bins.push(binCode);
      }
    }
    return bins.sort();

  } catch (error) {
    Logger.log('Error in getDestinationBins: ' + error.toString());
    throw error;
  }
}

/**
 * Move Skid Logic (Put-Away)
 * Handles multi-item skids by inserting rows at destination.
 * Writes ONLY to C, D, E, F, G, H, K. Copies A & B.
 */
function moveSkid(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(payload.sourceTab);
    const destSheet = ss.getSheetByName(payload.destTab); 
    
    if (!sourceSheet || !destSheet) throw new Error("Missing sheet(s)");
    
    // 1. Find Source Items (Rows with specific Skid_ID)
    const sourceData = sourceSheet.getDataRange().getValues();
    const skidId = payload.skidId;
    const sourceRowsIndices = [];
    const itemsToMove = [];
    
    // Scan source for skid items (Col K / Index 10)
    for(let i=1; i<sourceData.length; i++) {
      if(String(sourceData[i][STOCK_COLS.SKID_ID]) === String(skidId)) {
         sourceRowsIndices.push(i+1); // 1-based row index
         itemsToMove.push({
           pushNumber: sourceData[i][STOCK_COLS.PUSH_NUMBER],
           fbpn: sourceData[i][STOCK_COLS.FBPN],
           manufacturer: sourceData[i][STOCK_COLS.MANUFACTURER],
           project: sourceData[i][STOCK_COLS.PROJECT],
           initialQty: sourceData[i][STOCK_COLS.INITIAL_QTY], 
           currentQty: sourceData[i][STOCK_COLS.CURRENT_QTY],
           sku: sourceData[i][STOCK_COLS.SKU] 
         });
      }
    }
    
    if(itemsToMove.length === 0) throw new Error("No items found for skid " + skidId);
    
    // 2. Find Destination Bin Row
    const destData = destSheet.getDataRange().getValues();
    let destRowIndex = -1;
    let destBinCode = '';
    let destBinName = '';
    
    for(let i=1; i<destData.length; i++) {
      if(String(destData[i][STOCK_COLS.BIN_CODE]) === String(payload.destBinCode)) {
        destRowIndex = i+1;
        destBinCode = destData[i][STOCK_COLS.BIN_CODE];
        destBinName = destData[i][STOCK_COLS.BIN_NAME]; // Capture Bin Name for copying
        break;
      }
    }
    
    if(destRowIndex === -1) throw new Error("Destination bin not found: " + payload.destBinCode);
    
    // 3. Write First Item into the Existing Empty Slot
    const firstItem = itemsToMove[0];
    
    // Update specific columns only (C, D, E, F, G, H, K)
    // Note: getRange(row, col) is 1-indexed. Index 2 = Col C (3rd col)
    destSheet.getRange(destRowIndex, STOCK_COLS.PUSH_NUMBER + 1).setValue(firstItem.pushNumber);
    destSheet.getRange(destRowIndex, STOCK_COLS.FBPN + 1).setValue(firstItem.fbpn);
    destSheet.getRange(destRowIndex, STOCK_COLS.MANUFACTURER + 1).setValue(firstItem.manufacturer);
    destSheet.getRange(destRowIndex, STOCK_COLS.PROJECT + 1).setValue(firstItem.project);
    destSheet.getRange(destRowIndex, STOCK_COLS.INITIAL_QTY + 1).setValue(firstItem.initialQty);
    destSheet.getRange(destRowIndex, STOCK_COLS.CURRENT_QTY + 1).setValue(firstItem.currentQty);
    destSheet.getRange(destRowIndex, STOCK_COLS.SKID_ID + 1).setValue(skidId);
    
    // 4. Handle Additional Items (Insert Rows Below)
    if(itemsToMove.length > 1) {
      // We insert rows immediately after the first one we just filled.
      // We loop starting from the second item.
      
      for(let k=1; k<itemsToMove.length; k++) {
        const item = itemsToMove[k];
        const insertAt = destRowIndex + k; 
        
        // Insert 1 row after the previous one
        destSheet.insertRowAfter(insertAt - 1);
        
        // Construct row data. 
        // IMPORTANT: Copy Bin_Code (A) and Bin_Name (B) from parent
        // Initialize an empty array matching the sheet width
        const newRowData = new Array(destData[0].length).fill('');
        
        newRowData[STOCK_COLS.BIN_CODE] = destBinCode;     // A (Copy)
        newRowData[STOCK_COLS.BIN_NAME] = destBinName;     // B (Copy)
        newRowData[STOCK_COLS.PUSH_NUMBER] = item.pushNumber; // C
        newRowData[STOCK_COLS.FBPN] = item.fbpn;           // D
        newRowData[STOCK_COLS.MANUFACTURER] = item.manufacturer; // E
        newRowData[STOCK_COLS.PROJECT] = item.project;     // F
        newRowData[STOCK_COLS.INITIAL_QTY] = item.initialQty; // G
        newRowData[STOCK_COLS.CURRENT_QTY] = item.currentQty; // H
        // I (Stock Pct) & J (Audit) left empty
        newRowData[STOCK_COLS.SKID_ID] = skidId;           // K
        // L (SKU) left empty or passed if needed, request said only specific cols
        
        // Set values for the new row
        destSheet.getRange(insertAt, 1, 1, newRowData.length).setValues([newRowData]);
      }
    }
    
    // 5. Clear Source Rows
    // Only clear content cols C(3) through L(12) to keep the Staging Slot (Bin Code A) available
    // and ready for the next inbound shipment.
    sourceRowsIndices.forEach(r => {
      // Clear specific range: Cols C to L (Index 3 to 12)
      // getRange(row, column, numRows, numColumns)
      // C is col 3. L is col 12. Length = 10 columns.
      sourceSheet.getRange(r, 3, 1, 10).clearContent(); 
    });
    
    // 6. Log Transaction
    const logEntries = itemsToMove.map(item => ({
      timestamp: new Date(),
      action: 'PUT_AWAY',
      fbpn: item.fbpn,
      manufacturer: item.manufacturer,
      binCode: payload.destBinCode,
      qty: item.currentQty,
      project: item.project,
      sku: item.sku,
      user_email: Session.getActiveUser().getEmail(),
      description: `Moved Skid ${skidId} from ${payload.sourceTab} to ${payload.destTab}`
    }));
    
    writeStockToolsLocationLog(logEntries);
    
    return { success: true, message: `Successfully moved Skid ${skidId} (${itemsToMove.length} items) to ${payload.destBinCode}` };
    
  } catch (e) {
    Logger.log("moveSkid Error: " + e.toString());
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}


// ============================================================================
// MANUAL ADD + CONSOLIDATION
// ============================================================================

/**
 * Manual entry function for adding inventory directly to bins
 * (does NOT touch Stock_Totals)
 */
function manualAddInventory(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const destSheet = ss.getSheetByName(data.destTab);

    if (!destSheet) {
      return { success: false, message: 'Destination sheet not found' };
    }

    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date().toISOString();

    // Find an empty row in the destination bin
    const destData = destSheet.getDataRange().getValues();
    let emptyRowNum = -1;

    for (let i = 1; i < destData.length; i++) {
      const row = destData[i];
      if (row[STOCK_COLS.BIN_CODE] === data.binCode &&
          (!row[STOCK_COLS.FBPN] || row[STOCK_COLS.FBPN] === '')) {
        emptyRowNum = i + 1;
        break;
      }
    }

    if (emptyRowNum === -1) {
      return { success: false, message: 'No empty rows available in destination bin' };
    }

    // Write data
    destSheet.getRange(emptyRowNum, STOCK_COLS.PUSH_NUMBER + 1).setValue(data.pushNumber || '');
    destSheet.getRange(emptyRowNum, STOCK_COLS.FBPN + 1).setValue(data.fbpn);
    destSheet.getRange(emptyRowNum, STOCK_COLS.MANUFACTURER + 1).setValue(data.manufacturer || '');
    destSheet.getRange(emptyRowNum, STOCK_COLS.PROJECT + 1).setValue(data.project || '');
    destSheet.getRange(emptyRowNum, STOCK_COLS.INITIAL_QTY + 1).setValue(data.qty || 0);
    destSheet.getRange(emptyRowNum, STOCK_COLS.CURRENT_QTY + 1).setValue(data.qty || 0);
    destSheet.getRange(emptyRowNum, STOCK_COLS.STOCK_PCT + 1).setValue(100);
    destSheet.getRange(emptyRowNum, STOCK_COLS.SKID_ID + 1).setValue(data.skidId || '');
    destSheet.getRange(emptyRowNum, STOCK_COLS.SKU + 1).setValue(data.sku || '');

    // Log entry
    const logEntry = [{
      timestamp: timestamp,
      action: 'MANUAL_ADD',
      fbpn: data.fbpn,
      manufacturer: data.manufacturer || '',
      fromBin: 'N/A',
      fromTab: 'N/A',
      toBin: data.binCode,
      toTab: data.destTab,
      qty: data.qty,
      user: userEmail,
      skidId: data.skidId || '',
      sku: data.sku || ''
    }];

    writeStockToolsLocationLog(logEntry);

    return {
      success: true,
      message: `Successfully added ${data.qty} units of ${data.fbpn} to ${data.binCode}`
    };

  } catch (error) {
    Logger.log('Error in manualAddInventory: ' + error.toString());
    return {
      success: false,
      message: 'Error adding inventory: ' + error.toString()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Consolidate all stock from one bin into another
 * (no Stock_Totals change, this is intra-warehouse reshuffle)
 * payload: { sourceTab, destTab, sourceBinCode, destBinCode }
 */
function consolidateBinInventory(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(payload.sourceTab);
    const destSheet   = ss.getSheetByName(payload.destTab);

    if (!sourceSheet || !destSheet) {
      return { success: false, message: 'Source or destination sheet not found' };
    }

    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date().toISOString();
    const logEntries = [];

    const sourceData = sourceSheet.getDataRange().getValues();
    const destData   = destSheet.getDataRange().getValues();
    const headerLen  = destData[0].length;

    const items = [];
    const rowsToClear = [];

    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];
      if (row[STOCK_COLS.BIN_CODE] === payload.sourceBinCode &&
          row[STOCK_COLS.FBPN]) {
        items.push({
          fbpn: row[STOCK_COLS.FBPN],
          pushNumber: row[STOCK_COLS.PUSH_NUMBER] || '',
          manufacturer: row[STOCK_COLS.MANUFACTURER] || '',
          project: row[STOCK_COLS.PROJECT] || '',
          initialQty: row[STOCK_COLS.INITIAL_QTY] || row[STOCK_COLS.CURRENT_QTY] || 0,
          currentQty: row[STOCK_COLS.CURRENT_QTY] || 0,
          skidId: row[STOCK_COLS.SKID_ID] || '',
          sku: row[STOCK_COLS.SKU] || ''
        });
        rowsToClear.push(i + 1);
      }
    }

    if (!items.length) {
      return { success: false, message: 'No inventory found in source bin to consolidate.' };
    }

    // Destination bin display name (if any)
    let destBinName = '';
    for (let i = 1; i < destData.length; i++) {
      const row = destData[i];
      if (row[STOCK_COLS.BIN_CODE] === payload.destBinCode) {
        destBinName = row[STOCK_COLS.BIN_NAME] || payload.destBinCode;
        break;
      }
    }
    if (!destBinName) destBinName = payload.destBinCode;

    // Use generic placement for all
    items.forEach(item => {
      const qtyMoved = Number(item.currentQty) || 0;
      if (!qtyMoved) return;

      // Find an empty slot in the destination bin
      let emptyRow = -1;
      let lastBinRow = -1;

      for (let i = 1; i < destData.length; i++) {
        const row = destData[i];
        if (row[STOCK_COLS.BIN_CODE] === payload.destBinCode) {
          lastBinRow = i + 1;
          if (!row[STOCK_COLS.FBPN] && emptyRow === -1) emptyRow = i + 1;
        }
      }

      let newRowIdx;
      if (emptyRow !== -1) {
        newRowIdx = emptyRow;
      } else {
        const insertAfter = lastBinRow > 0 ? lastBinRow : destData.length;
        destSheet.insertRowAfter(insertAfter);
        newRowIdx = insertAfter + 1;
      }

      const vals = new Array(headerLen).fill('');
      vals[STOCK_COLS.BIN_CODE]      = payload.destBinCode;
      vals[STOCK_COLS.BIN_NAME]      = destBinName;
      vals[STOCK_COLS.PUSH_NUMBER] = item.pushNumber || '';
      vals[STOCK_COLS.FBPN]        = item.fbpn || '';
      vals[STOCK_COLS.MANUFACTURER]= item.manufacturer || '';
      vals[STOCK_COLS.PROJECT]     = item.project || '';
      vals[STOCK_COLS.INITIAL_QTY] = item.initialQty || item.currentQty || 0;
      vals[STOCK_COLS.CURRENT_QTY] = item.currentQty || 0;
      vals[STOCK_COLS.STOCK_PCT]   = 100;
      vals[STOCK_COLS.SKID_ID]     = item.skidId || '';
      vals[STOCK_COLS.SKU]         = item.sku || '';

      destSheet.getRange(newRowIdx, 1, 1, headerLen).setValues([vals]);

      logEntries.push({
        timestamp,
        action: 'CONSOLIDATE',
        fbpn: item.fbpn,
        manufacturer: item.manufacturer || '',
        fromBin: payload.sourceBinCode,
        fromTab: payload.sourceTab,
        toBin: payload.destBinCode,
        toTab: payload.destTab,
        qty: qtyMoved,
        user: userEmail,
        skidId: item.skidId || '',
        sku: item.sku || ''
      });
    });

    // Clear source bin rows
    rowsToClear.forEach(rowNum => {
      sourceSheet.getRange(rowNum, STOCK_COLS.PUSH_NUMBER + 1).setValue('');
      sourceSheet.getRange(rowNum, STOCK_COLS.FBPN + 1).setValue('');
      sourceSheet.getRange(rowNum, STOCK_COLS.MANUFACTURER + 1).setValue('');
      sourceSheet.getRange(rowNum, STOCK_COLS.PROJECT + 1).setValue('');
      sourceSheet.getRange(rowNum, STOCK_COLS.INITIAL_QTY + 1).setValue('');
      sourceSheet.getRange(rowNum, STOCK_COLS.CURRENT_QTY + 1).setValue('');
      sourceSheet.getRange(rowNum, STOCK_COLS.STOCK_PCT + 1).setValue('');
      sourceSheet.getRange(rowNum, STOCK_COLS.SKID_ID + 1).setValue('');
      sourceSheet.getRange(rowNum, STOCK_COLS.SKU + 1).setValue('');
    });

    writeStockToolsLocationLog(logEntries);

    return {
      success: true,
      message: `Consolidated ${items.length} line(s) from ${payload.sourceTab}/${payload.sourceBinCode} to ${payload.destTab}/${payload.destBinCode}`
    };

  } catch (error) {
    Logger.log('Error in consolidateBinInventory: ' + error.toString());
    return {
      success: false,
      message: 'Error consolidating inventory: ' + error.toString()
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// LOGGING HELPER
// ============================================================================

function writeStockToolsLocationLog(logEntries) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('LocationLog');
  if (!logSheet) return;
  
  const headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  
  logEntries.forEach(entry => {
    const row = [];
    headers.forEach(h => {
      switch(h) {
        case 'Timestamp': row.push(entry.timestamp); break;
        case 'Action': row.push(entry.action); break;
        case 'FBPN': row.push(entry.fbpn); break;
        case 'Manufacturer': row.push(entry.manufacturer); break;
        case 'Bin_Code': row.push(entry.binCode); break;
        case 'Qty_Changed': row.push(entry.qty); break;
        case 'Project': row.push(entry.project); break;
        case 'User_Email': row.push(entry.user_email); break;
        case 'Description': row.push(entry.description); break;
        default: row.push('');
      }
    });
    logSheet.appendRow(row);
  });
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet ${name} not found`);
  return sheet;
}