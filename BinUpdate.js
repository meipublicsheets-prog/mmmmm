// ============================================================================
// INVENTORY MANAGEMENT.GS - Add, Remove, Move Inventory Operations
// ============================================================================

function openAddInventoryModal() {
  const html = HtmlService.createTemplateFromFile('AddInventoryModal')
    .evaluate()
    .setWidth(900)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Inventory to Bins');
}

function openRemoveInventoryModal() {
  const html = HtmlService.createTemplateFromFile('RemoveInventoryModal')
    .evaluate()
    .setWidth(900)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Remove Inventory from Bins');
}

function openMoveInventoryModal() {
  const html = HtmlService.createTemplateFromFile('MoveInventoryModal')
    .evaluate()
    .setWidth(900)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Move Inventory Between Bins');
}

function openMoveFromStagingModal() {
  const html = HtmlService.createTemplateFromFile('MoveFromStagingModal')
    .evaluate()
    .setWidth(900)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Move Inventory from Inbound Staging to Bins');
}

function openTransferInventoryModal() {
  const html = HtmlService.createTemplateFromFile('TransferInventoryModal')
    .evaluate()
    .setWidth(900)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Transfer Inventory Between Projects');
}

/**
 * Adds inventory to one or more bins based on the provided data.
 * @param {Array} inventoryData - Array of objects { binCode, fbpn, manufacturer, project, qtyToAdd, notes, pushNumber, skidId }
 * @returns {Object} Result summary
 */
function addInventoryBatch(inventoryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const binStockSheet = ss.getSheetByName('Bin_Stock');
    const locationLogSheet = ss.getSheetByName('LocationLog');
    const floorStockSheet = ss.getSheetByName('Floor_Stock_Levels');

    if (!binStockSheet || !locationLogSheet || !floorStockSheet) {
      throw new Error('Required sheets not found. Please check that Bin_Stock, LocationLog, and Floor_Stock_Levels exist.');
    }

    const userEmail = Session.getActiveUser().getEmail();

    // Get existing data
    const binStockData = binStockSheet.getDataRange().getValues();
    const binHeaders = binStockData[0];

    const logData = locationLogSheet.getDataRange().getValues();
    const logHeaders = logData[0];

    const floorData = floorStockSheet.getDataRange().getValues();
    const floorHeaders = floorData[0];

    // Find column indexes in Bin_Stock
    const binCodeIndex = binHeaders.indexOf('Bin_Code');
    const fbpnIndex = binHeaders.indexOf('FBPN');
    const manufacturerIndex = binHeaders.indexOf('Manufacturer');
    const projectIndex = binHeaders.indexOf('Project');
    const currentQtyIndex = binHeaders.indexOf('Current_Quantity');
    const initialQtyIndex = binHeaders.indexOf('Initial_Quantity');
    const stockPercentageIndex = binHeaders.indexOf('Stock_Percentage');
    const pushNumberIndex = binHeaders.indexOf('Push_Number');
    const skidIdIndex = binHeaders.indexOf('Skid_ID');

    // LocationLog indexes
    const logTimestampIndex = logHeaders.indexOf('Timestamp');
    const logActionIndex = logHeaders.indexOf('Action');
    const logFbpnIndex = logHeaders.indexOf('FBPN');
    const logManufacturerIndex = logHeaders.indexOf('Manufacturer');
    const logBinCodeIndex = logHeaders.indexOf('Bin_Code');
    const logQtyChangedIndex = logHeaders.indexOf('Qty_Changed');
    const logResultingQtyIndex = logHeaders.indexOf('Resulting_Qty');
    const logDescriptionIndex = logHeaders.indexOf('Description');
    const logUserEmailIndex = logHeaders.indexOf('User_Email');
    const logProjectIndex = logHeaders.indexOf('Project');

    // Floor_Stock_Levels indexes
    const floorFbpnIndex = floorHeaders.indexOf('FBPN');
    const floorProjectIndex = floorHeaders.indexOf('Project');
    const floorQtyIndex = floorHeaders.indexOf('Qty_On_Floor');

    const binStockUpdates = [];
    const binStockNewRows = [];
    const logNewRows = [];
    const floorStockUpdates = [];

    let successCount = 0;
    const errors = [];

    // Map existing Bin_Stock records for faster lookup
    const binStockMap = {};
    for (let i = 1; i < binStockData.length; i++) {
      const row = binStockData[i];
      const binCodeVal = row[binCodeIndex];
      const fbpnVal = row[fbpnIndex];
      const manufacturerVal = row[manufacturerIndex];
      const projectVal = row[projectIndex];

      if (binCodeVal && fbpnVal && manufacturerVal && projectVal) {
        const key = `${binCodeVal}|${fbpnVal}|${manufacturerVal}|${projectVal}`;
        binStockMap[key] = { rowIndex: i + 1, rowData: row };
      }
    }

    // Map existing Floor_Stock_Levels records
    const floorStockMap = {};
    for (let i = 1; i < floorData.length; i++) {
      const row = floorData[i];
      const fbpnVal = row[floorFbpnIndex];
      const projectVal = row[floorProjectIndex];

      if (fbpnVal && projectVal) {
        const key = `${fbpnVal}|${projectVal}`;
        floorStockMap[key] = { rowIndex: i + 1, rowData: row };
      }
    }

    // Process each inventory record
    inventoryData.forEach((item, index) => {
      try {
        const binCode = item.binCode;
        const fbpn = item.fbpn;
        const manufacturer = item.manufacturer;
        const project = item.project;
        const qtyToAdd = parseInt(item.qtyToAdd) || 0;
        const notes = item.notes || '';
        const pushNumber = item.pushNumber || '';
        const skidId = item.skidId || '';

        if (!binCode || !fbpn || !manufacturer || !project) {
          throw new Error(`Missing required fields in item #${index + 1}`);
        }

        if (qtyToAdd <= 0) {
          throw new Error(`Quantity to add must be positive in item #${index + 1}`);
        }

        // Update or create Bin_Stock record
        const key = `${binCode}|${fbpn}|${manufacturer}|${project}`;
        if (binStockMap[key]) {
          const record = binStockMap[key];
          const rowIndex = record.rowIndex;
          const rowData = record.rowData;

          const currentQty = parseInt(rowData[currentQtyIndex]) || 0;
          const newQty = currentQty + qtyToAdd;

          const initialQty = rowData[initialQtyIndex] || newQty;
          const stockPercentage = ((newQty / initialQty) * 100).toFixed(2);

          const updateRange = binStockSheet.getRange(rowIndex, 1, 1, binHeaders.length);
          const updatedRow = rowData.slice();
          updatedRow[currentQtyIndex] = newQty;
          updatedRow[initialQtyIndex] = initialQty;
          updatedRow[stockPercentageIndex] = parseFloat(stockPercentage);
          updatedRow[pushNumberIndex] = pushNumber || rowData[pushNumberIndex];
          updatedRow[skidIdIndex] = skidId || rowData[skidIdIndex];
          binStockUpdates.push({ range: updateRange, values: [updatedRow] });

          // Update map
          record.rowData = updatedRow;
        } else {
          // Create new row
          const initialQty = qtyToAdd;
          const stockPercentage = 100.0;

          const newRow = new Array(binHeaders.length).fill('');
          newRow[binCodeIndex] = binCode;
          newRow[fbpnIndex] = fbpn;
          newRow[manufacturerIndex] = manufacturer;
          newRow[projectIndex] = project;
          newRow[currentQtyIndex] = qtyToAdd;
          newRow[initialQtyIndex] = initialQty;
          newRow[stockPercentageIndex] = stockPercentage;
          newRow[pushNumberIndex] = pushNumber;
          newRow[skidIdIndex] = skidId;

          binStockNewRows.push(newRow);
        }

        // Update Floor_Stock_Levels (if exists)
        const floorKey = `${fbpn}|${project}`;
        if (floorStockMap[floorKey]) {
          const record = floorStockMap[floorKey];
          const rowIndex = record.rowIndex;
          const rowData = record.rowData;

          const currentFloorQty = parseInt(rowData[floorQtyIndex]) || 0;
          const newFloorQty = currentFloorQty + qtyToAdd;

          const updateRange = floorStockSheet.getRange(rowIndex, 1, 1, floorHeaders.length);
          const updatedRow = rowData.slice();
          updatedRow[floorQtyIndex] = newFloorQty;
          floorStockUpdates.push({ range: updateRange, values: [updatedRow] });

          // Update map
          record.rowData = updatedRow;
        } else {
          // Create new floor stock record
          const newRow = new Array(floorHeaders.length).fill('');
          newRow[floorFbpnIndex] = fbpn;
          newRow[floorProjectIndex] = project;
          newRow[floorQtyIndex] = qtyToAdd;
          floorStockSheet.appendRow(newRow);
        }

        // Add LocationLog entry
        const logRow = new Array(logHeaders.length).fill('');
        logRow[logTimestampIndex] = new Date();
        logRow[logActionIndex] = 'ADD';
        logRow[logFbpnIndex] = fbpn;
        logRow[logManufacturerIndex] = manufacturer;
        logRow[logBinCodeIndex] = binCode;
        logRow[logQtyChangedIndex] = qtyToAdd;
        logRow[logResultingQtyIndex] = ''; // Will be updated after all operations
        logRow[logDescriptionIndex] = notes || `Added ${qtyToAdd} units to bin ${binCode}`;
        logRow[logUserEmailIndex] = userEmail;
        logRow[logProjectIndex] = project;

        logNewRows.push(logRow);

        successCount++;
      } catch (itemError) {
        errors.push(`Error processing item #${index + 1}: ${itemError.message}`);
      }
    });

    // Apply Bin_Stock updates
    binStockUpdates.forEach(update => {
      update.range.setValues(update.values);
    });

    // Append new Bin_Stock rows
    if (binStockNewRows.length > 0) {
      binStockSheet.getRange(binStockSheet.getLastRow() + 1, 1, binStockNewRows.length, binHeaders.length)
        .setValues(binStockNewRows);
    }

    // Apply Floor_Stock_Levels updates
    floorStockUpdates.forEach(update => {
      update.range.setValues(update.values);
    });

    // Append LocationLog rows
    if (logNewRows.length > 0) {
      locationLogSheet.getRange(locationLogSheet.getLastRow() + 1, 1, logNewRows.length, logHeaders.length)
        .setValues(logNewRows);
    }

    return {
      success: errors.length === 0,
      message: `Processed ${inventoryData.length} items. Successfully added inventory for ${successCount} items.`,
      errors: errors
    };
  } catch (error) {
    Logger.log('Error in addInventoryBatch: ' + error.toString());
    return {
      success: false,
      message: 'Error adding inventory: ' + error.message
    };
  }
}

/**
 * Removes inventory from one or more bins.
 * @param {Array} inventoryData - Array of objects { binCode, fbpn, manufacturer, project, qtyToRemove, notes }
 * @returns {Object} Result summary
 */
function removeInventoryBatch(inventoryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const binStockSheet = ss.getSheetByName('Bin_Stock');
    const locationLogSheet = ss.getSheetByName('LocationLog');
    const floorStockSheet = ss.getSheetByName('Floor_Stock_Levels');

    if (!binStockSheet || !locationLogSheet || !floorStockSheet) {
      throw new Error('Required sheets not found. Please check that Bin_Stock, LocationLog, and Floor_Stock_Levels exist.');
    }

    const userEmail = Session.getActiveUser().getEmail();

    const binStockData = binStockSheet.getDataRange().getValues();
    const binHeaders = binStockData[0];

    const logData = locationLogSheet.getDataRange().getValues();
    const logHeaders = logData[0];

    const floorData = floorStockSheet.getDataRange().getValues();
    const floorHeaders = floorData[0];

    const binCodeIndex = binHeaders.indexOf('Bin_Code');
    const fbpnIndex = binHeaders.indexOf('FBPN');
    const manufacturerIndex = binHeaders.indexOf('Manufacturer');
    const projectIndex = binHeaders.indexOf('Project');
    const currentQtyIndex = binHeaders.indexOf('Current_Quantity');
    const initialQtyIndex = binHeaders.indexOf('Initial_Quantity');
    const stockPercentageIndex = binHeaders.indexOf('Stock_Percentage');

    const logTimestampIndex = logHeaders.indexOf('Timestamp');
    const logActionIndex = logHeaders.indexOf('Action');
    const logFbpnIndex = logHeaders.indexOf('FBPN');
    const logManufacturerIndex = logHeaders.indexOf('Manufacturer');
    const logBinCodeIndex = logHeaders.indexOf('Bin_Code');
    const logQtyChangedIndex = logHeaders.indexOf('Qty_Changed');
    const logResultingQtyIndex = logHeaders.indexOf('Resulting_Qty');
    const logDescriptionIndex = logHeaders.indexOf('Description');
    const logUserEmailIndex = logHeaders.indexOf('User_Email');
    const logProjectIndex = logHeaders.indexOf('Project');

    const floorFbpnIndex = floorHeaders.indexOf('FBPN');
    const floorProjectIndex = floorHeaders.indexOf('Project');
    const floorQtyIndex = floorHeaders.indexOf('Qty_On_Floor');

    const binStockUpdates = [];
    const logNewRows = [];
    const floorStockUpdates = [];

    let successCount = 0;
    const errors = [];

    const binStockMap = {};
    for (let i = 1; i < binStockData.length; i++) {
      const row = binStockData[i];
      const binCodeVal = row[binCodeIndex];
      const fbpnVal = row[fbpnIndex];
      const manufacturerVal = row[manufacturerIndex];
      const projectVal = row[projectIndex];

      if (binCodeVal && fbpnVal && manufacturerVal && projectVal) {
        const key = `${binCodeVal}|${fbpnVal}|${manufacturerVal}|${projectVal}`;
        binStockMap[key] = { rowIndex: i + 1, rowData: row };
      }
    }

    const floorStockMap = {};
    for (let i = 1; i < floorData.length; i++) {
      const row = floorData[i];
      const fbpnVal = row[floorFbpnIndex];
      const projectVal = row[floorProjectIndex];

      if (fbpnVal && projectVal) {
        const key = `${fbpnVal}|${projectVal}`;
        floorStockMap[key] = { rowIndex: i + 1, rowData: row };
      }
    }

    inventoryData.forEach((item, index) => {
      try {
        const binCode = item.binCode;
        const fbpn = item.fbpn;
        const manufacturer = item.manufacturer;
        const project = item.project;
        const qtyToRemove = parseInt(item.qtyToRemove) || 0;
        const notes = item.notes || '';

        if (!binCode || !fbpn || !manufacturer || !project) {
          throw new Error(`Missing required fields in item #${index + 1}`);
        }

        if (qtyToRemove <= 0) {
          throw new Error(`Quantity to remove must be positive in item #${index + 1}`);
        }

        const key = `${binCode}|${fbpn}|${manufacturer}|${project}`;
        if (!binStockMap[key]) {
          throw new Error(`No existing record found for bin ${binCode}, FBPN ${fbpn}, manufacturer ${manufacturer}, project ${project} in item #${index + 1}`);
        }

        const record = binStockMap[key];
        const rowIndex = record.rowIndex;
        const rowData = record.rowData;

        const currentQty = parseInt(rowData[currentQtyIndex]) || 0;
        if (qtyToRemove > currentQty) {
          throw new Error(`Insufficient quantity in bin ${binCode} for item #${index + 1}. Current: ${currentQty}, requested to remove: ${qtyToRemove}`);
        }

        const newQty = currentQty - qtyToRemove;
        const initialQty = rowData[initialQtyIndex] || currentQty;
        const stockPercentage = initialQty > 0 ? ((newQty / initialQty) * 100).toFixed(2) : 0;

        const updateRange = binStockSheet.getRange(rowIndex, 1, 1, binHeaders.length);
        const updatedRow = rowData.slice();
        updatedRow[currentQtyIndex] = newQty;
        updatedRow[initialQtyIndex] = initialQty;
        updatedRow[stockPercentageIndex] = parseFloat(stockPercentage);
        binStockUpdates.push({ range: updateRange, values: [updatedRow] });

        record.rowData = updatedRow;

        const floorKey = `${fbpn}|${project}`;
        if (floorStockMap[floorKey]) {
          const floorRecord = floorStockMap[floorKey];
          const floorRowIndex = floorRecord.rowIndex;
          const floorRowData = floorRecord.rowData;

          const currentFloorQty = parseInt(floorRowData[floorQtyIndex]) || 0;
          const newFloorQty = Math.max(0, currentFloorQty - qtyToRemove);

          const floorUpdateRange = floorStockSheet.getRange(floorRowIndex, 1, 1, floorHeaders.length);
          const updatedFloorRow = floorRowData.slice();
          updatedFloorRow[floorQtyIndex] = newFloorQty;
          floorStockUpdates.push({ range: floorUpdateRange, values: [updatedFloorRow] });

          floorRecord.rowData = updatedFloorRow;
        }

        const logRow = new Array(logHeaders.length).fill('');
        logRow[logTimestampIndex] = new Date();
        logRow[logActionIndex] = 'REMOVE';
        logRow[logFbpnIndex] = fbpn;
        logRow[logManufacturerIndex] = manufacturer;
        logRow[logBinCodeIndex] = binCode;
        logRow[logQtyChangedIndex] = -qtyToRemove;
        logRow[logResultingQtyIndex] = newQty;
        logRow[logDescriptionIndex] = notes || `Removed ${qtyToRemove} units from bin ${binCode}`;
        logRow[logUserEmailIndex] = userEmail;
        logRow[logProjectIndex] = project;

        logNewRows.push(logRow);

        successCount++;
      } catch (itemError) {
        errors.push(`Error processing item #${index + 1}: ${itemError.message}`);
      }
    });

    binStockUpdates.forEach(update => {
      update.range.setValues(update.values);
    });

    floorStockUpdates.forEach(update => {
      update.range.setValues(update.values);
    });

    if (logNewRows.length > 0) {
      locationLogSheet.getRange(locationLogSheet.getLastRow() + 1, 1, logNewRows.length, logHeaders.length)
        .setValues(logNewRows);
    }

    return {
      success: errors.length === 0,
      message: `Processed ${inventoryData.length} items. Successfully removed inventory for ${successCount} items.`,
      errors: errors
    };
  } catch (error) {
    Logger.log('Error in removeInventoryBatch: ' + error.toString());
    return {
      success: false,
      message: 'Error removing inventory: ' + error.message
    };
  }
}

/**
 * Moves inventory between bins.
 * @param {Array} moveData - Array of objects { sourceBinCode, targetBinCode, fbpn, manufacturer, project, qtyToMove, notes }
 * @returns {Object} Result summary
 */
function moveInventoryBatch(moveData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const binStockSheet = ss.getSheetByName('Bin_Stock');
    const locationLogSheet = ss.getSheetByName('LocationLog');
    const floorStockSheet = ss.getSheetByName('Floor_Stock_Levels');

    if (!binStockSheet || !locationLogSheet || !floorStockSheet) {
      throw new Error('Required sheets not found. Please check that Bin_Stock, LocationLog, and Floor_Stock_Levels exist.');
    }

    const userEmail = Session.getActiveUser().getEmail();

    const binStockData = binStockSheet.getDataRange().getValues();
    const binHeaders = binStockData[0];

    const logData = locationLogSheet.getDataRange().getValues();
    const logHeaders = logData[0];

    const floorData = floorStockSheet.getDataRange().getValues();
    const floorHeaders = floorData[0];

    const binCodeIndex = binHeaders.indexOf('Bin_Code');
    const fbpnIndex = binHeaders.indexOf('FBPN');
    const manufacturerIndex = binHeaders.indexOf('Manufacturer');
    const projectIndex = binHeaders.indexOf('Project');
    const currentQtyIndex = binHeaders.indexOf('Current_Quantity');
    const initialQtyIndex = binHeaders.indexOf('Initial_Quantity');
    const stockPercentageIndex = binHeaders.indexOf('Stock_Percentage');

    const logTimestampIndex = logHeaders.indexOf('Timestamp');
    const logActionIndex = logHeaders.indexOf('Action');
    const logFbpnIndex = logHeaders.indexOf('FBPN');
    const logManufacturerIndex = logHeaders.indexOf('Manufacturer');
    const logBinCodeIndex = logHeaders.indexOf('Bin_Code');
    const logQtyChangedIndex = logHeaders.indexOf('Qty_Changed');
    const logResultingQtyIndex = logHeaders.indexOf('Resulting_Qty');
    const logDescriptionIndex = logHeaders.indexOf('Description');
    const logUserEmailIndex = logHeaders.indexOf('User_Email');
    const logProjectIndex = logHeaders.indexOf('Project');

    const floorFbpnIndex = floorHeaders.indexOf('FBPN');
    const floorProjectIndex = floorHeaders.indexOf('Project');
    const floorQtyIndex = floorHeaders.indexOf('Qty_On_Floor');

    const binStockUpdates = [];
    const logNewRows = [];

    let successCount = 0;
    const errors = [];

    const binStockMap = {};
    for (let i = 1; i < binStockData.length; i++) {
      const row = binStockData[i];
      const binCodeVal = row[binCodeIndex];
      const fbpnVal = row[fbpnIndex];
      const manufacturerVal = row[manufacturerIndex];
      const projectVal = row[projectIndex];

      if (binCodeVal && fbpnVal && manufacturerVal && projectVal) {
        const key = `${binCodeVal}|${fbpnVal}|${manufacturerVal}|${projectVal}`;
        binStockMap[key] = { rowIndex: i + 1, rowData: row };
      }
    }

    const floorStockMap = {};
    for (let i = 1; i < floorData.length; i++) {
      const row = floorData[i];
      const fbpnVal = row[floorFbpnIndex];
      const projectVal = row[floorProjectIndex];

      if (fbpnVal && projectVal) {
        const key = `${fbpnVal}|${projectVal}`;
        floorStockMap[key] = { rowIndex: i + 1, rowData: row };
      }
    }

    moveData.forEach((item, index) => {
      try {
        const sourceBinCode = item.sourceBinCode;
        const targetBinCode = item.targetBinCode;
        const fbpn = item.fbpn;
        const manufacturer = item.manufacturer;
        const project = item.project;
        const qtyToMove = parseInt(item.qtyToMove) || 0;
        const notes = item.notes || '';

        if (!sourceBinCode || !targetBinCode || !fbpn || !manufacturer || !project) {
          throw new Error(`Missing required fields in item #${index + 1}`);
        }

        if (qtyToMove <= 0) {
          throw new Error(`Quantity to move must be positive in item #${index + 1}`);
        }

        const sourceKey = `${sourceBinCode}|${fbpn}|${manufacturer}|${project}`;
        const targetKey = `${targetBinCode}|${fbpn}|${manufacturer}|${project}`;

        if (!binStockMap[sourceKey]) {
          throw new Error(`No existing record found for source bin ${sourceBinCode} in item #${index + 1}`);
        }

        const sourceRecord = binStockMap[sourceKey];
        const sourceRowIndex = sourceRecord.rowIndex;
        const sourceRowData = sourceRecord.rowData;

        const sourceCurrentQty = parseInt(sourceRowData[currentQtyIndex]) || 0;
        if (qtyToMove > sourceCurrentQty) {
          throw new Error(`Insufficient quantity in source bin ${sourceBinCode} for item #${index + 1}. Current: ${sourceCurrentQty}, requested to move: ${qtyToMove}`);
        }

        const newSourceQty = sourceCurrentQty - qtyToMove;
        const sourceInitialQty = sourceRowData[initialQtyIndex] || sourceCurrentQty;
        const sourceStockPercentage = sourceInitialQty > 0 ? ((newSourceQty / sourceInitialQty) * 100).toFixed(2) : 0;

        const sourceUpdateRange = binStockSheet.getRange(sourceRowIndex, 1, 1, binHeaders.length);
        const updatedSourceRow = sourceRowData.slice();
        updatedSourceRow[currentQtyIndex] = newSourceQty;
        updatedSourceRow[initialQtyIndex] = sourceInitialQty;
        updatedSourceRow[stockPercentageIndex] = parseFloat(sourceStockPercentage);
        binStockUpdates.push({ range: sourceUpdateRange, values: [updatedSourceRow] });

        sourceRecord.rowData = updatedSourceRow;

        if (binStockMap[targetKey]) {
          const targetRecord = binStockMap[targetKey];
          const targetRowIndex = targetRecord.rowIndex;
          const targetRowData = targetRecord.rowData;

          const targetCurrentQty = parseInt(targetRowData[currentQtyIndex]) || 0;
          const newTargetQty = targetCurrentQty + qtyToMove;

          const targetInitialQty = targetRowData[initialQtyIndex] || newTargetQty;
          const targetStockPercentage = ((newTargetQty / targetInitialQty) * 100).toFixed(2);

          const targetUpdateRange = binStockSheet.getRange(targetRowIndex, 1, 1, binHeaders.length);
          const updatedTargetRow = targetRowData.slice();
          updatedTargetRow[currentQtyIndex] = newTargetQty;
          updatedTargetRow[initialQtyIndex] = targetInitialQty;
          updatedTargetRow[stockPercentageIndex] = parseFloat(targetStockPercentage);
          binStockUpdates.push({ range: targetUpdateRange, values: [updatedTargetRow] });

          targetRecord.rowData = updatedTargetRow;
        } else {
          const initialQty = qtyToMove;
          const stockPercentage = 100.0;

          const newRow = new Array(binHeaders.length).fill('');
          newRow[binCodeIndex] = targetBinCode;
          newRow[fbpnIndex] = fbpn;
          newRow[manufacturerIndex] = manufacturer;
          newRow[projectIndex] = project;
          newRow[currentQtyIndex] = qtyToMove;
          newRow[initialQtyIndex] = initialQty;
          newRow[stockPercentageIndex] = stockPercentage;

          binStockSheet.appendRow(newRow);
        }

        const sourceLogRow = new Array(logHeaders.length).fill('');
        sourceLogRow[logTimestampIndex] = new Date();
        sourceLogRow[logActionIndex] = 'MOVE_OUT';
        sourceLogRow[logFbpnIndex] = fbpn;
        sourceLogRow[logManufacturerIndex] = manufacturer;
        sourceLogRow[logBinCodeIndex] = sourceBinCode;
        sourceLogRow[logQtyChangedIndex] = -qtyToMove;
        sourceLogRow[logResultingQtyIndex] = newSourceQty;
        sourceLogRow[logDescriptionIndex] = notes || `Moved ${qtyToMove} units from bin ${sourceBinCode} to bin ${targetBinCode}`;
        sourceLogRow[logUserEmailIndex] = userEmail;
        sourceLogRow[logProjectIndex] = project;

        logNewRows.push(sourceLogRow);

        const targetLogRow = new Array(logHeaders.length).fill('');
        targetLogRow[logTimestampIndex] = new Date();
        targetLogRow[logActionIndex] = 'MOVE_IN';
        targetLogRow[logFbpnIndex] = fbpn;
        targetLogRow[logManufacturerIndex] = manufacturer;
        targetLogRow[logBinCodeIndex] = targetBinCode;
        targetLogRow[logQtyChangedIndex] = qtyToMove;
        targetLogRow[logResultingQtyIndex] = ''; 
        targetLogRow[logDescriptionIndex] = notes || `Moved ${qtyToMove} units to bin ${targetBinCode} from bin ${sourceBinCode}`;
        targetLogRow[logUserEmailIndex] = userEmail;
        targetLogRow[logProjectIndex] = project;

        logNewRows.push(targetLogRow);

        successCount++;
      } catch (itemError) {
        errors.push(`Error processing move item #${index + 1}: ${itemError.message}`);
      }
    });

    binStockUpdates.forEach(update => {
      update.range.setValues(update.values);
    });

    if (logNewRows.length > 0) {
      locationLogSheet.getRange(locationLogSheet.getLastRow() + 1, 1, logNewRows.length, logHeaders.length)
        .setValues(logNewRows);
    }

    return {
      success: errors.length === 0,
      message: `Processed ${moveData.length} move items. Successfully moved inventory for ${successCount} items.`,
      errors: errors
    };
  } catch (error) {
    Logger.log('Error in moveInventoryBatch: ' + error.toString());
    return {
      success: false,
      message: 'Error moving inventory: ' + error.message
    };
  }
}

/**
 * Moves inventory from Inbound_Staging to Bin_Stock and updates Floor_Stock_Levels.
 * @param {Array} moveData - Array of objects { skidId, binCode, fbpn, manufacturer, project, qtyToMove, notes, inboundStagingQty, pushNumber }
 * @returns {Object} Result summary
 */
function moveFromStaging(moveData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inboundStagingSheet = ss.getSheetByName('Inbound_Staging');
    const binStockSheet = ss.getSheetByName('Bin_Stock');
    const locationLogSheet = ss.getSheetByName('LocationLog');
    const floorStockSheet = ss.getSheetByName('Floor_Stock_Levels');

    if (!inboundStagingSheet || !binStockSheet || !locationLogSheet || !floorStockSheet) {
      throw new Error('Required sheets not found. Please check that Inbound_Staging, Bin_Stock, LocationLog, and Floor_Stock_Levels exist.');
    }

    const userEmail = Session.getActiveUser().getEmail();

    const inboundData = inboundStagingSheet.getDataRange().getValues();
    const inboundHeaders = inboundData[0];

    const binStockData = binStockSheet.getDataRange().getValues();
    const binHeaders = binStockData[0];

    const logData = locationLogSheet.getDataRange().getValues();
    const logHeaders = logData[0];

    const floorData = floorStockSheet.getDataRange().getValues();
    const floorHeaders = floorData[0];

    const inboundSkidIdIndex = inboundHeaders.indexOf('Skid_ID');
    const inboundFbpnIndex = inboundHeaders.indexOf('FBPN');
    const inboundManufacturerIndex = inboundHeaders.indexOf('Manufacturer');
    const inboundProjectIndex = inboundHeaders.indexOf('Project');
    const inboundQtyIndex = inboundHeaders.indexOf('Qty_In_Staging');
    const inboundPushNumberIndex = inboundHeaders.indexOf('Push_Number');

    const binCodeIndex = binHeaders.indexOf('Bin_Code');
    const fbpnIndex = binHeaders.indexOf('FBPN');
    const manufacturerIndex = binHeaders.indexOf('Manufacturer');
    const projectIndex = binHeaders.indexOf('Project');
    const currentQtyIndex = binHeaders.indexOf('Current_Quantity');
    const initialQtyIndex = binHeaders.indexOf('Initial_Quantity');
    const stockPercentageIndex = binHeaders.indexOf('Stock_Percentage');
    const binPushNumberIndex = binHeaders.indexOf('Push_Number');
    const binSkidIdIndex = binHeaders.indexOf('Skid_ID');

    const logTimestampIndex = logHeaders.indexOf('Timestamp');
    const logActionIndex = logHeaders.indexOf('Action');
    const logFbpnIndex = logHeaders.indexOf('FBPN');
    const logManufacturerIndex = logHeaders.indexOf('Manufacturer');
    const logBinCodeIndex = logHeaders.indexOf('Bin_Code');
    const logQtyChangedIndex = logHeaders.indexOf('Qty_Changed');
    const logResultingQtyIndex = logHeaders.indexOf('Resulting_Qty');
    const logDescriptionIndex = logHeaders.indexOf('Description');
    const logUserEmailIndex = logHeaders.indexOf('User_Email');
    const logProjectIndex = logHeaders.indexOf('Project');

    const floorFbpnIndex = floorHeaders.indexOf('FBPN');
    const floorProjectIndex = floorHeaders.indexOf('Project');
    const floorQtyIndex = floorHeaders.indexOf('Qty_On_Floor');

    const inboundMap = {};
    for (let i = 1; i < inboundData.length; i++) {
      const row = inboundData[i];
      const skidIdVal = row[inboundSkidIdIndex];
      const fbpnVal = row[inboundFbpnIndex];
      const manufacturerVal = row[inboundManufacturerIndex];
      const projectVal = row[inboundProjectIndex];

      if (skidIdVal && fbpnVal && manufacturerVal && projectVal) {
        const key = `${skidIdVal}|${fbpnVal}|${manufacturerVal}|${projectVal}`;
        inboundMap[key] = { rowIndex: i + 1, rowData: row };
      }
    }

    const binStockMap = {};
    for (let i = 1; i < binStockData.length; i++) {
      const row = binStockData[i];
      const binCodeVal = row[binCodeIndex];
      const fbpnVal = row[fbpnIndex];
      const manufacturerVal = row[manufacturerIndex];
      const projectVal = row[projectIndex];

      if (binCodeVal && fbpnVal && manufacturerVal && projectVal) {
        const key = `${binCodeVal}|${fbpnVal}|${manufacturerVal}|${projectVal}`;
        binStockMap[key] = { rowIndex: i + 1, rowData: row };
      }
    }

    const floorStockMap = {};
    for (let i = 1; i < floorData.length; i++) {
      const row = floorData[i];
      const fbpnVal = row[floorFbpnIndex];
      const projectVal = row[floorProjectIndex];

      if (fbpnVal && projectVal) {
        const key = `${fbpnVal}|${projectVal}`;
        floorStockMap[key] = { rowIndex: i + 1, rowData: row };
      }
    }

    const inboundUpdates = [];
    const binStockUpdates = [];
    const binStockNewRows = [];
    const floorStockUpdates = [];
    const logNewRows = [];

    let successCount = 0;
    const errors = [];

    moveData.forEach((item, index) => {
      try {
        const skidId = item.skidId;
        const binCode = item.binCode;
        const fbpn = item.fbpn;
        const manufacturer = item.manufacturer;
        const project = item.project;
        const qtyToMove = parseInt(item.qtyToMove) || 0;
        const notes = item.notes || '';
        const pushNumber = item.pushNumber || '';

        if (!skidId || !binCode || !fbpn || !manufacturer || !project) {
          throw new Error(`Missing required fields in item #${index + 1}`);
        }

        if (qtyToMove <= 0) {
          throw new Error(`Quantity to move must be positive in item #${index + 1}`);
        }

        const inboundKey = `${skidId}|${fbpn}|${manufacturer}|${project}`;
        if (!inboundMap[inboundKey]) {
          throw new Error(`No matching record found in Inbound_Staging for Skid_ID ${skidId}, FBPN ${fbpn}, manufacturer ${manufacturer}, project ${project} in item #${index + 1}`);
        }

        const inboundRecord = inboundMap[inboundKey];
        const inboundRowIndex = inboundRecord.rowIndex;
        const inboundRowData = inboundRecord.rowData;

        const currentInboundQty = parseInt(inboundRowData[inboundQtyIndex]) || 0;
        if (qtyToMove > currentInboundQty) {
          throw new Error(`Insufficient quantity in Inbound_Staging for item #${index + 1}. Current: ${currentInboundQty}, requested to move: ${qtyToMove}`);
        }

        const newInboundQty = currentInboundQty - qtyToMove;
        const inboundUpdateRange = inboundStagingSheet.getRange(inboundRowIndex, 1, 1, inboundHeaders.length);
        const updatedInboundRow = inboundRowData.slice();
        updatedInboundRow[inboundQtyIndex] = newInboundQty;
        inboundUpdates.push({ range: inboundUpdateRange, values: [updatedInboundRow] });

        inboundRecord.rowData = updatedInboundRow;

        const binKey = `${binCode}|${fbpn}|${manufacturer}|${project}`;
        let resultingBinQty = 0;

        if (binStockMap[binKey]) {
          const binRecord = binStockMap[binKey];
          const binRowIndex = binRecord.rowIndex;
          const binRowData = binRecord.rowData;

          const currentBinQty = parseInt(binRowData[currentQtyIndex]) || 0;
          const newBinQty = currentBinQty + qtyToMove;

          const initialQty = binRowData[initialQtyIndex] || newBinQty;
          const stockPercentage = ((newBinQty / initialQty) * 100).toFixed(2);

          const binUpdateRange = binStockSheet.getRange(binRowIndex, 1, 1, binHeaders.length);
          const updatedBinRow = binRowData.slice();
          updatedBinRow[currentQtyIndex] = newBinQty;
          updatedBinRow[initialQtyIndex] = initialQty;
          updatedBinRow[stockPercentageIndex] = parseFloat(stockPercentage);
          updatedBinRow[binPushNumberIndex] = pushNumber || binRowData[binPushNumberIndex];
          updatedBinRow[binSkidIdIndex] = skidId || binRowData[binSkidIdIndex];
          binStockUpdates.push({ range: binUpdateRange, values: [updatedBinRow] });

          binRecord.rowData = updatedBinRow;
          resultingBinQty = newBinQty;
        } else {
          const initialQty = qtyToMove;
          const stockPercentage = 100.0;

          const newRow = new Array(binHeaders.length).fill('');
          newRow[binCodeIndex] = binCode;
          newRow[fbpnIndex] = fbpn;
          newRow[manufacturerIndex] = manufacturer;
          newRow[projectIndex] = project;
          newRow[currentQtyIndex] = qtyToMove;
          newRow[initialQtyIndex] = initialQty;
          newRow[stockPercentageIndex] = stockPercentage;
          newRow[binPushNumberIndex] = pushNumber;
          newRow[binSkidIdIndex] = skidId;

          binStockNewRows.push(newRow);
          resultingBinQty = qtyToMove;
        }

        const floorKey = `${fbpn}|${project}`;
        if (floorStockMap[floorKey]) {
          const floorRecord = floorStockMap[floorKey];
          const floorRowIndex = floorRecord.rowIndex;
          const floorRowData = floorRecord.rowData;

          const currentFloorQty = parseInt(floorRowData[floorQtyIndex]) || 0;
          const newFloorQty = currentFloorQty + qtyToMove;

          const floorUpdateRange = floorStockSheet.getRange(floorRowIndex, 1, 1, floorHeaders.length);
          const updatedFloorRow = floorRowData.slice();
          updatedFloorRow[floorQtyIndex] = newFloorQty;
          floorStockUpdates.push({ range: floorUpdateRange, values: [updatedFloorRow] });

          floorRecord.rowData = updatedFloorRow;
        } else {
          const newRow = new Array(floorHeaders.length).fill('');
          newRow[floorFbpnIndex] = fbpn;
          newRow[floorProjectIndex] = project;
          newRow[floorQtyIndex] = qtyToMove;
          floorStockSheet.appendRow(newRow);
        }

        const logRow = new Array(logHeaders.length).fill('');
        logRow[logTimestampIndex] = new Date();
        logRow[logActionIndex] = 'STAGING_TO_BIN';
        logRow[logFbpnIndex] = fbpn;
        logRow[logManufacturerIndex] = manufacturer;
        logRow[logBinCodeIndex] = binCode;
        logRow[logQtyChangedIndex] = qtyToMove;
        logRow[logResultingQtyIndex] = resultingBinQty;
        logRow[logDescriptionIndex] = notes || `Moved ${qtyToMove} units from Inbound_Staging skid ${skidId} to bin ${binCode}`;
        logRow[logUserEmailIndex] = userEmail;
        logRow[logProjectIndex] = project;

        logNewRows.push(logRow);

        successCount++;
      } catch (itemError) {
        errors.push(`Error processing staging move item #${index + 1}: ${itemError.message}`);
      }
    });

    inboundUpdates.forEach(update => {
      update.range.setValues(update.values);
    });

    binStockUpdates.forEach(update => {
      update.range.setValues(update.values);
    });

    if (binStockNewRows.length > 0) {
      binStockSheet.getRange(binStockSheet.getLastRow() + 1, 1, binStockNewRows.length, binHeaders.length)
        .setValues(binStockNewRows);
    }

    floorStockUpdates.forEach(update => {
      update.range.setValues(update.values);
    });

    if (logNewRows.length > 0) {
      locationLogSheet.getRange(locationLogSheet.getLastRow() + 1, 1, logNewRows.length, logHeaders.length)
        .setValues(logNewRows);
    }

    return {
      success: errors.length === 0,
      message: `Processed ${moveData.length} staging move items. Successfully moved inventory for ${successCount} items.`,
      errors: errors
    };
  } catch (error) {
    Logger.log('Error in moveFromStaging: ' + error.toString());
    return {
      success: false,
      message: 'Error moving from staging: ' + error.message
    };
  }
}
