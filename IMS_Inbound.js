/* ===== FILE CONTENT START: IMS_Inbound_FIXED5.gs ===== */

// ============================================================================
// HELPERS & LOOKUPS
// ============================================================================

function lookupProjectFromPO(customerPO) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = TABS.PO_MASTER || 'PO_Master';
    const poMasterSheet = ss.getSheetByName(sheetName);
    if (!poMasterSheet) return '';

    const data = poMasterSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === customerPO.trim()) {
        return data[i][1] ? data[i][1].toString() : '';
      }
    }
    return '';
  } catch (error) {
    Logger.log('Error in lookupProjectFromPO: ' + error.toString());
    return '';
  }
}

/**
 * Generates a random alphanumeric ID
 */
function generateRandomId(prefix, length) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return prefix + result;
}

function getManufacturers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = new Set();

  function addFromSheet_(sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;

    const data = sh.getDataRange().getValues();
    const headers = data[0].map(h => (h || '').toString().trim());
    const mCol = headers.indexOf('Manufacturer');
    if (mCol === -1) return;

    for (let i = 1; i < data.length; i++) {
      const v = (data[i][mCol] || '').toString().trim();
      if (v) out.add(v);
    }
  }

  addFromSheet_(TABS.ITEM_MASTER);
  if (out.size === 0) addFromSheet_(TABS.PROJECT_MASTER);

  return Array.from(out).sort((a, b) => a.localeCompare(b));
}

function getFBPNList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemMasterSheet = ss.getSheetByName(TABS.ITEM_MASTER);
  if (!itemMasterSheet) return [];
  const data = itemMasterSheet.getDataRange().getValues();
  const headers = data[0];
  const fbpnIdx = headers.indexOf('FBPN');
  const set = new Set();
  for (let i = 1; i < data.length; i++) {
    const v = data[i][fbpnIdx];
    if (v) set.add(v);
  }
  return Array.from(set).sort();
}

/**
 * Creates a lookup map from Project_Master for fast retrieval of MFPN/Description
 * Key: FBPN (uppercase) -> Value: {mfpn, description}
 */
function getProjectMasterMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.PROJECT_MASTER || 'Project_Master');
  const map = {};

  if (!sheet) return map;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;

  const headers = data[0].map(h => String(h).trim());
  const fbpnIdx = headers.indexOf('FBPN');
  const mfpnIdx = headers.indexOf('MFPN');
  const descIdx = headers.indexOf('Description');

  if (fbpnIdx === -1) return map;

  for (let i = 1; i < data.length; i++) {
    const fbpn = String(data[i][fbpnIdx] || '').toUpperCase().trim();
    if (fbpn) {
      map[fbpn] = {
        mfpn: mfpnIdx > -1 ? String(data[i][mfpnIdx] || '') : '',
        description: descIdx > -1 ? String(data[i][descIdx] || '') : ''
      };
    }
  }
  return map;
}

function generateSKU(fbpn, manufacturer) {
  if (!fbpn || !manufacturer) return '';
  const cleanFBPN = fbpn.toString().trim();
  const cleanMan = manufacturer.toString().trim();
  if (!cleanFBPN || !cleanMan) return '';
  const manPrefix = cleanMan.substring(0, 3).toUpperCase();
  return `${cleanFBPN}-${manPrefix}`;
}

function buildStockSkuFromFBPNAndManufacturer(fbpn, manufacturer) {
  return generateSKU(fbpn, manufacturer);
}

function getNextStagingSequence() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.INBOUND_STAGING);
  if (!sheet) return 1;

  if (sheet.getLastRow() <= 1) return 1;

  const data = sheet.getDataRange().getValues();
  const binCol = data[0].indexOf('Bin_Code');

  if (binCol === -1) return 1;

  let maxSeq = 0;
  for (let i = 1; i < data.length; i++) {
    const val = String(data[i][binCol]);
    const match = val.match(/IS\.(\d+)/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (!isNaN(num) && num > maxSeq) maxSeq = num;
    }
  }
  return maxSeq + 1;
}

function getNextSkidIdBase() {
  return 0; // Deprecated
}

function getNextInboundBinIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.INBOUND_STAGING);
  if (!sheet || sheet.getLastRow() <= 1) return 1;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const binCol = headers.indexOf('Bin_Code');

  if (binCol === -1) return 1;

  let maxIndex = 0;
  for (let i = 1; i < data.length; i++) {
    const bin = String(data[i][binCol] || '');
    const match = bin.match(/^IS\.(\d+)$/i);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxIndex) maxIndex = num;
    }
  }
  return maxIndex + 1;
}

function createInboundFolder_(dateObj, bolNumber) {
  const rootId = (typeof FOLDERS !== 'undefined' && FOLDERS.INBOUND_UPLOADS)
    ? FOLDERS.INBOUND_UPLOADS
    : (typeof FOLDERS !== 'undefined' && FOLDERS.IMS_ROOT ? FOLDERS.IMS_ROOT : '');

  if (!rootId) throw new Error("Inbound Uploads Folder ID not configured.");

  const rootFolder = DriveApp.getFolderById(rootId);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  const monthName = `${months[dateObj.getMonth()]} ${dateObj.getFullYear()}`;
  const monthFolder = getOrCreateSubfolder_(rootFolder, monthName);

  const dayName = String(dateObj.getDate()).padStart(2, '0');
  const dayFolder = getOrCreateSubfolder_(monthFolder, dayName);

  const safeBol = String(bolNumber || 'NO_BOL').trim().replace(/[\/\\?%*:|"<>\.]/g, '_');
  const bolFolder = getOrCreateSubfolder_(dayFolder, safeBol);

  return bolFolder;
}

/**
 * Parse inbound date coming from HTML modal.
 * Fixes the common “day behind” bug when input is a date-only string (YYYY-MM-DD),
 * which JS parses as UTC and shifts in local time.
 */
function parseInboundDate_(v, fallbackDate) {
  const fb = fallbackDate || new Date();
  if (!v) return fb;

  if (Object.prototype.toString.call(v) === '[object Date]') {
    const d = v;
    return isNaN(d.getTime()) ? fb : d;
  }

  const s = String(v).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const y = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    const da = parseInt(m[3], 10);
    const d = new Date(y, mo, da); // local midnight (no UTC shift)
    return isNaN(d.getTime()) ? fb : d;
  }

  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? fb : d2;
}

// ============================================================================
// MAIN PROCESSING
// ============================================================================

function processInboundSubmission(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inboundSkidsSheet = ss.getSheetByName(TABS.INBOUND_SKIDS);
  const masterLogSheet = ss.getSheetByName(TABS.MASTER_LOG);
  const inboundStagingSheet = ss.getSheetByName(TABS.INBOUND_STAGING);

  if (!inboundSkidsSheet || !masterLogSheet || !inboundStagingSheet) {
    throw new Error('Required sheets (Inbound_Skids, Master_Log, or Inbound_Staging) not found.');
  }

  const userEmail = Session.getActiveUser().getEmail();
  const now = new Date();

  const basic = payload.tab1 || {};
  const skidsData = payload.skids || [];
  const options = payload.options || {};
  options.generateLabels = (options.generateLabels === false) ? false : true;

  if (!skidsData.length) throw new Error('No skids provided in inbound submission.');

  const timestamp = new Date();

  // IMPORTANT: Column B / Date_Received must be the date selected in the HTML modal calendar.
  // Also fixes “day behind” from YYYY-MM-DD parsing.
  const dateReceivedRaw = basic.deliveryDate || basic.dateReceived || timestamp;
  const dateReceived = parseInboundDate_(dateReceivedRaw, timestamp);

  // Generate Random TXN ID (INB-XXXXXXXX)
  const txnId = generateRandomId('INB-', 8);

  // Create Inbound Folder
  let orderFolder = null;
  try {
    orderFolder = createInboundFolder_(dateReceived, basic.bolNumber);
  } catch (e) {
    Logger.log("Could not create inbound folder: " + e.toString());
  }

  // Handle File Uploads (Save to Order Folder)
  if (orderFolder && payload.files) {
    try {
      if (payload.files.bol) {
        const f = payload.files.bol;
        const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mimeType, f.name);
        orderFolder.createFile(blob);
      }
      if (payload.files.packingList) {
        const f = payload.files.packingList;
        const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mimeType, f.name);
        orderFolder.createFile(blob);
      }
    } catch (e) {
      Logger.log("Error saving uploaded files: " + e.toString());
    }
  }

  const mapHeaders = (headers) => {
    const map = {};
    headers.forEach((h, i) => map[h.toString().trim()] = i);
    return map;
  };

  const masterHeaders = masterLogSheet.getDataRange().getValues()[0];
  const skidHeaders = inboundSkidsSheet.getDataRange().getValues()[0];
  const stagingHeaders = inboundStagingSheet.getDataRange().getValues()[0];

  const mIdx = (h) => masterHeaders.indexOf(h);
  const sIdx = (h) => skidHeaders.indexOf(h);
  const stIdx = (h) => stagingHeaders.indexOf(h);

  // Pre-load data maps & validations
  const projectMasterMap = getProjectMasterMap();

  // -- VALIDATION CHECK START --
  const projectColIdx = sIdx('Project') + 1; // 1-based index
  const validProjects = getValidProjects_(inboundSkidsSheet, projectColIdx);
  // -- VALIDATION CHECK END --

  let binNumericBase = getNextInboundBinIndex();

  const masterLogRows = [];
  const inboundStagingRows = [];
  const inboundSkidsRows = [];
  const labelData = [];

  const totalSkidCount = skidsData.length;

  // Aggregation Map for Master Log
  const masterAggregation = {}; // { fbpn: { qty, ...info } }

  skidsData.forEach((skidEntry, index) => {
    // Generate Random Skid ID (SKD-XXXXXXXX)
    const skidId = generateRandomId('SKD-', 8);

    const skidItems = skidEntry.items || [{ fbpn: skidEntry.fbpn, qty: skidEntry.qty }];

    const currentBinIndex = binNumericBase + index;
    const binCode = `IS.${currentBinIndex}`;
    const binName = `Inbound Staging - Skid ${currentBinIndex}`;
    const skidSequenceNum = index + 1;

    skidItems.forEach(item => {
      const fbpn = String(item.fbpn || '').toUpperCase().trim();
      const qty = parseFloat(item.qty) || 0;
      const manufacturer = basic.manufacturer || '';
      const sku = generateSKU(fbpn, manufacturer);

      const meta = projectMasterMap[fbpn] || { mfpn: '', description: '' };
      const mfpn = meta.mfpn;
      const description = meta.description;

      // Validate Project Name
      let project = basic.project;
      if (validProjects.length > 0 && !validProjects.includes(project)) {
        const match = validProjects.find(p => p.toUpperCase() === String(project).toUpperCase());
        if (match) {
          project = match;
        } else {
          throw new Error(`Invalid Project '${project}'. Allowed: ${validProjects.slice(0, 5).join(', ')}...`);
        }
      }

      // --- Aggregate for Master Log ---
      if (!masterAggregation[fbpn]) {
        masterAggregation[fbpn] = {
          fbpn: fbpn,
          qty: 0,
          manufacturer: manufacturer,
          sku: sku,
          mfpn: mfpn,
          description: description,
          project: project
        };
      }
      masterAggregation[fbpn].qty += qty;

      // --- Inbound Skids Row (Detailed Breakdown) ---
      const sRow = new Array(skidHeaders.length).fill('');
      if (sIdx('Skid_ID') > -1) sRow[sIdx('Skid_ID')] = skidId;
      if (sIdx('TXN_ID') > -1) sRow[sIdx('TXN_ID')] = txnId;

      // IMPORTANT: Date column must use the selected Delivery Date (not "now")
      if (sIdx('Date') > -1) sRow[sIdx('Date')] = dateReceived;

      if (sIdx('FBPN') > -1) sRow[sIdx('FBPN')] = fbpn;
      if (sIdx('Qty_on_Skid') > -1) sRow[sIdx('Qty_on_Skid')] = qty;
      if (sIdx('SKU') > -1) sRow[sIdx('SKU')] = sku;
      if (sIdx('Project') > -1) sRow[sIdx('Project')] = project;
      if (sIdx('MFPN') > -1) sRow[sIdx('MFPN')] = mfpn;
      if (sIdx('Is_Mixed') > -1) sRow[sIdx('Is_Mixed')] = skidItems.length > 1 ? 'TRUE' : 'FALSE';
      if (sIdx('Skid_Sequence') > -1) sRow[sIdx('Skid_Sequence')] = skidSequenceNum;

      // Timestamp stays as actual submission time
      if (sIdx('Timestamp') > -1) sRow[sIdx('Timestamp')] = now;

      inboundSkidsRows.push(sRow);

      // --- Inbound Staging Row (Detailed Breakdown) ---
      const stRow = new Array(stagingHeaders.length).fill('');
      if (stIdx('Bin_Code') > -1) stRow[stIdx('Bin_Code')] = binCode;
      if (stIdx('Bin_Name') > -1) stRow[stIdx('Bin_Name')] = binName;
      if (stIdx('Push_Number') > -1) stRow[stIdx('Push_Number')] = basic.pushNumber;
      if (stIdx('FBPN') > -1) stRow[stIdx('FBPN')] = fbpn;
      if (stIdx('Manufacturer') > -1) stRow[stIdx('Manufacturer')] = manufacturer;
      if (stIdx('Project') > -1) stRow[stIdx('Project')] = project;
      if (stIdx('Initial_Quantity') > -1) stRow[stIdx('Initial_Quantity')] = qty;
      if (stIdx('Current_Quantity') > -1) stRow[stIdx('Current_Quantity')] = qty;
      if (stIdx('Stock_Percentage') > -1) stRow[stIdx('Stock_Percentage')] = 1;
      if (stIdx('AUDIT NEEDED') > -1) stRow[stIdx('AUDIT NEEDED')] = 'FALSE';
      if (stIdx('Skid_ID') > -1) stRow[stIdx('Skid_ID')] = skidId;
      if (stIdx('SKU') > -1) stRow[stIdx('SKU')] = sku;
      inboundStagingRows.push(stRow);

      labelData.push({
        skidId: skidId,
        fbpn: fbpn,
        quantity: qty,
        sku: sku,
        manufacturer: manufacturer,
        project: project,
        pushNumber: basic.pushNumber,
        dateReceived: formatDate(dateReceived),
        skidNumber: skidSequenceNum,
        totalSkids: totalSkidCount
      });
    });
  });

  // --- Build Master Log Rows from Aggregation ---
  for (const fbpn in masterAggregation) {
    const item = masterAggregation[fbpn];
    const mRow = new Array(masterHeaders.length).fill('');

    if (mIdx('Txn_ID') > -1) mRow[mIdx('Txn_ID')] = txnId;

    // IMPORTANT: Date_Received must use the selected Delivery Date (not "now")
    if (mIdx('Date_Received') > -1) mRow[mIdx('Date_Received')] = dateReceived;

    if (mIdx('Transaction_Type') > -1) mRow[mIdx('Transaction_Type')] = 'Inbound';
    if (mIdx('FBPN') > -1) mRow[mIdx('FBPN')] = item.fbpn;
    if (mIdx('Qty_Received') > -1) mRow[mIdx('Qty_Received')] = item.qty; // Total Qty
    if (mIdx('Total_Skid_Count') > -1) mRow[mIdx('Total_Skid_Count')] = totalSkidCount;
    if (mIdx('Manufacturer') > -1) mRow[mIdx('Manufacturer')] = item.manufacturer;
    if (mIdx('SKU') > -1) mRow[mIdx('SKU')] = item.sku;
    if (mIdx('Received_By') > -1) mRow[mIdx('Received_By')] = userEmail;
    if (mIdx('Warehouse') > -1) mRow[mIdx('Warehouse')] = basic.warehouse;
    if (mIdx('Push #') > -1) mRow[mIdx('Push #')] = basic.pushNumber;
    if (mIdx('Carrier') > -1) mRow[mIdx('Carrier')] = basic.carrier;
    if (mIdx('BOL_Number') > -1) mRow[mIdx('BOL_Number')] = basic.bolNumber;
    if (mIdx('Customer_PO_Number') > -1) mRow[mIdx('Customer_PO_Number')] = basic.customerPO;
    if (mIdx('MFPN') > -1) mRow[mIdx('MFPN')] = item.mfpn;
    if (mIdx('Description') > -1) mRow[mIdx('Description')] = item.description;
    if (mIdx('Project') > -1) mRow[mIdx('Project')] = item.project;

    masterLogRows.push(mRow);
  }

  // Bulk Write: insert new entries at TOP, starting row 3 (skip row 2)
  appendRowsSafe_(masterLogSheet, masterLogRows);
  appendRowsSafe_(inboundSkidsSheet, inboundSkidsRows);
  appendRowsSafe_(inboundStagingSheet, inboundStagingRows);

  // Link Folder to BOL in Master Log
  // NOTE: rows now insert at row 3; link range must target row 3..(3+len-1)
  if (orderFolder && masterLogRows.length) {
    try {
      const bolColIdx = mIdx('BOL_Number') + 1;
      if (bolColIdx > 0) {
        const startRow = 3;
        const folderUrl = orderFolder.getUrl();

        const bolRange = masterLogSheet.getRange(startRow, bolColIdx, masterLogRows.length, 1);
        const richTextValues = masterLogRows.map(row => {
          const bolVal = row[mIdx('BOL_Number')] || 'Link';
          return SpreadsheetApp.newRichTextValue()
            .setText(String(bolVal))
            .setLinkUrl(folderUrl)
            .build();
        });
        bolRange.setRichTextValues(richTextValues.map(r => [r]));
      }
    } catch (e) {
      Logger.log("Error linking folder: " + e);
    }
  }

  inboundSkidsRows.forEach(row => {
    const fbpn = row[sIdx('FBPN')];
    const qty = row[sIdx('Qty_on_Skid')];
    if (fbpn && qty) {
      fulfillBackorders(ss, fbpn, qty, txnId);
    }
  });

  const labelResult = { labelsGenerated: false, labelFileUrl: '', labelHtmlUrl: '' };
  if (options.generateLabels) {
    try {
      let orderFolder2 = null;
      try { orderFolder2 = createInboundFolder_(dateReceived, basic.bolNumber); } catch (e) {}

      const res = generateSkidLabels(labelData, {
        bolNumber: basic.bolNumber,
        targetFolder: orderFolder2
      });
      labelResult.labelsGenerated = !!(res && res.success);
      labelResult.labelFileUrl = (res && res.pdfUrl) ? res.pdfUrl : '';
      labelResult.labelHtmlUrl = (res && res.htmlUrl) ? res.htmlUrl : '';
    } catch (e) {
      Logger.log('Error generating labels: ' + e.toString());
    }
  }

  return {
    success: true,
    txnId,
    totalSkids: skidsData.length,
    labelResult,
    labelPdfUrl: labelResult.labelFileUrl,
    labelHtmlUrl: labelResult.labelHtmlUrl
  };
}

/**
 * Insert rows at the TOP, starting at row 3 (skipping row 2).
 * Assumes row 1 = headers, row 2 reserved (filters/notes/etc).
 */
function appendRowsSafe_(sheet, rows) {
  if (!rows || rows.length === 0) return;

  const numCols = rows[0].length;

  // Ensure sheet has at least 2 rows (header + reserved row 2)
  const maxRows = sheet.getMaxRows();
  if (maxRows < 2) {
    sheet.insertRowsAfter(maxRows, 2 - maxRows);
  }

  // Always insert after row 2, then write starting at row 3
  sheet.insertRowsAfter(2, rows.length);
  sheet.getRange(3, 1, rows.length, numCols).setValues(rows);
}

function getValidProjects_(sheet, colIndex) {
  try {
    const range = sheet.getRange(2, colIndex);
    const rule = range.getDataValidation();
    if (rule) {
      const criteria = rule.getCriteriaType();
      if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        return rule.getCriteriaValues()[0];
      }
    }
    return [];
  } catch (e) {
    Logger.log('Error fetching validation rules: ' + e);
    return [];
  }
}

function updateInboundStaging(stagingRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inboundStagingSheet = ss.getSheetByName(TABS.INBOUND_STAGING);
  if (!inboundStagingSheet) throw new Error('Inbound_Staging sheet not found');
  appendRowsSafe_(inboundStagingSheet, stagingRows);
}

// ============================================================================
// BACKORDER LOGIC
// ============================================================================

function fulfillBackorders(ss, fbpn, qtyReceived, txnId) {
  const backordersSheet = ss.getSheetByName(TABS.BACKORDERS);
  const fulfillmentLogSheet = ss.getSheetByName(TABS.BACKORDERFULFILLMENT_LOG);
  const requestedItemsSheet = ss.getSheetByName(TABS.REQUESTED_ITEMS);
  const customerOrdersSheet = ss.getSheetByName(TABS.CUSTOMER_ORDERS);

  if (!backordersSheet || !fulfillmentLogSheet || !requestedItemsSheet || !customerOrdersSheet) return;

  let remainingQty = qtyReceived;
  const backData = backordersSheet.getDataRange().getValues();
  const headers = backData[0];
  const mapHeaders = (h) => { const m = {}; h.forEach((v, i) => m[v.toString().trim()] = i); return m; };
  const hIdx = mapHeaders(headers);

  const logHeaders = fulfillmentLogSheet.getDataRange().getValues()[0] || [];
  const lIdx = mapHeaders(logHeaders);

  const logRows = [];
  const userEmail = Session.getActiveUser().getEmail();
  const updatedBackorders = [];

  for (let i = 1; i < backData.length && remainingQty > 0; i++) {
    const row = backData[i];
    const boStatus = (row[hIdx['Status']] || '').toString().toUpperCase();
    const boFBPN = (row[hIdx['FBPN']] || '').toString().trim().toUpperCase();

    if (boStatus === 'CLOSED') continue;
    if (boFBPN !== fbpn.toUpperCase()) continue;

    const qtyNeeded = parseFloat(row[hIdx['Qty_Requested']]) || 0;
    const qtyFulfilledSoFar = parseFloat(row[hIdx['Qty_Fulfilled']]) || 0;
    const remainingNeed = qtyNeeded - qtyFulfilledSoFar;
    if (remainingNeed <= 0) continue;

    const fulfillQty = Math.min(remainingNeed, remainingQty);
    if (fulfillQty <= 0) continue;

    const newFulfilled = qtyFulfilledSoFar + fulfillQty;
    remainingQty -= fulfillQty;

    row[hIdx['Qty_Fulfilled']] = newFulfilled;
    row[hIdx['Status']] = newFulfilled >= qtyNeeded ? 'Closed' : 'Partial';

    updatedBackorders.push({ rowIndex: i + 1, rowValues: row });

    const boId = row[hIdx['Backorder_ID']];
    const orderId = row[hIdx['Order_ID']];

    updateAllocationWithFulfillment(ss, boId, orderId, fbpn, fulfillQty);
    updateCustomerOrderStockStatus(ss, orderId);

    const logRow = new Array(logHeaders.length).fill('');
    if (lIdx['Timestamp'] !== undefined) logRow[lIdx['Timestamp']] = new Date();
    if (lIdx['Txn_ID'] !== undefined) logRow[lIdx['Txn_ID']] = txnId;
    if (lIdx['Backorder_ID'] !== undefined) logRow[lIdx['Backorder_ID']] = boId;
    if (lIdx['Order_ID'] !== undefined) logRow[lIdx['Order_ID']] = orderId;
    if (lIdx['FBPN'] !== undefined) logRow[lIdx['FBPN']] = fbpn;
    if (lIdx['Qty_Fulfilled'] !== undefined) logRow[lIdx['Qty_Fulfilled']] = fulfillQty;
    if (lIdx['Fulfilled_By'] !== undefined) logRow[lIdx['Fulfilled_By']] = userEmail;
    logRows.push(logRow);
  }

  updatedBackorders.forEach(u => {
    backordersSheet.getRange(u.rowIndex, 1, 1, headers.length).setValues([u.rowValues]);
  });

  if (logRows.length) {
    appendRowsSafe_(fulfillmentLogSheet, logRows);
  }
}

function updateAllocationWithFulfillment(ss, backorderId, orderId, fbpn, qtyFulfilled) {
  const requestedItemsSheet = ss.getSheetByName(TABS.REQUESTED_ITEMS);
  const allocationLogSheet = ss.getSheetByName(TABS.ALLOCATION_LOG);
  if (!requestedItemsSheet || !allocationLogSheet) return;

  const reqData = requestedItemsSheet.getDataRange().getValues();
  const reqHeaders = reqData[0];
  const idxOrderId = reqHeaders.indexOf('Order_ID');
  const idxFBPN = reqHeaders.indexOf('FBPN');
  const idxBackorderedQty = reqHeaders.indexOf('Qty_Backordered');
  const idxAllocatedQty = reqHeaders.indexOf('Qty_Allocated');
  const idxStockStatus = reqHeaders.indexOf('Stock_Status');

  for (let i = 1; i < reqData.length; i++) {
    const row = reqData[i];
    if (row[idxOrderId] === orderId && (row[idxFBPN] || '').toString().toUpperCase() === fbpn.toUpperCase()) {
      const backordered = parseFloat(row[idxBackorderedQty]) || 0;
      const allocated = parseFloat(row[idxAllocatedQty]) || 0;
      const newBackordered = Math.max(0, backordered - qtyFulfilled);
      const newAllocated = allocated + qtyFulfilled;
      row[idxBackorderedQty] = newBackordered;
      row[idxAllocatedQty] = newAllocated;
      row[idxStockStatus] = newBackordered === 0 ? 'In Stock' : 'Partial';
      requestedItemsSheet.getRange(i + 1, 1, 1, reqHeaders.length).setValues([row]);
      break;
    }
  }

  const allocData = allocationLogSheet.getDataRange().getValues();
  const allocHeaders = allocData[0];
  const idxBoId = allocHeaders.indexOf('Backorder_ID');
  const idxAllocQty = allocHeaders.indexOf('Qty_Allocated');
  const idxAllocStatus = allocHeaders.indexOf('Allocation_Status');

  for (let i = 1; i < allocData.length; i++) {
    const row = allocData[i];
    if (row[idxBoId] === backorderId) {
      const allocQty = parseFloat(row[idxAllocQty]) || 0;
      row[idxAllocQty] = allocQty + qtyFulfilled;
      row[idxAllocStatus] = 'Fulfilled';
      allocationLogSheet.getRange(i + 1, 1, 1, allocHeaders.length).setValues([row]);
      break;
    }
  }
}

function updateCustomerOrderStockStatus(ss, orderId) {
  const customerOrdersSheet = ss.getSheetByName(TABS.CUSTOMER_ORDERS);
  const requestedItemsSheet = ss.getSheetByName(TABS.REQUESTED_ITEMS);
  if (!customerOrdersSheet || !requestedItemsSheet) return;

  const reqData = requestedItemsSheet.getDataRange().getValues();
  const reqHeaders = reqData[0];
  const idxOrderId = reqHeaders.indexOf('Order_ID');
  const idxStockStatus = reqHeaders.indexOf('Stock_Status');

  let hasBackorder = false;
  let hasInStock = false;

  for (let i = 1; i < reqData.length; i++) {
    const row = reqData[i];
    if (row[idxOrderId] === orderId) {
      const status = (row[idxStockStatus] || '').toString();
      if (status === 'Backorder' || status === 'Partial') hasBackorder = true;
      if (status === 'In Stock') hasInStock = true;
    }
  }

  const ordersData = customerOrdersSheet.getDataRange().getValues();
  const ordersHeaders = ordersData[0];
  const idxOrderIdCol = ordersHeaders.indexOf('Order_ID');
  const idxStockStatusCol = ordersHeaders.indexOf('Stock_Status');

  for (let i = 1; i < ordersData.length; i++) {
    const row = ordersData[i];
    if (row[idxOrderIdCol] === orderId) {
      let newStatus = row[idxStockStatusCol] || 'Pending';
      if (hasBackorder && hasInStock) newStatus = 'Partial Allocation';
      else if (hasBackorder) newStatus = 'Awaiting Stock';
      else if (hasInStock) newStatus = 'Allocated';
      customerOrdersSheet.getRange(i + 1, idxStockStatusCol + 1).setValue(newStatus);
      break;
    }
  }
}

// ============================================================================
// LABELS
// ============================================================================

function generateSkidLabels(labelData, meta) {
  try {
    const html = generateSkidLabelsHtml_(labelData || []);
    let folder = meta.targetFolder;
    if (!folder) {
      const parentId = (typeof FOLDERS !== 'undefined' && FOLDERS.INBOUND_UPLOADS) ? FOLDERS.INBOUND_UPLOADS : FOLDERS.IMS_ROOT;
      let rootFolder;
      try { rootFolder = DriveApp.getFolderById(parentId); } catch (e) { throw new Error("Cannot access root folder: " + parentId); }
      folder = getOrCreateSubfolder_(rootFolder, 'Labels');
    }

    const bol = meta.bolNumber || 'NO_BOL';
    const res = saveLabelsToDrive(html, folder, bol);

    return { success: true, htmlUrl: res.htmlFile.getUrl(), pdfUrl: res.pdfFile.getUrl() };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function generateSkidLabelsHtml_(labelData) {
  const base = getSkidLookupBaseUrl_();
  let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    @page { size: 6in 4in; margin: 0; }
    * { box-sizing: border-box; }
    body { margin: 0; padding: 0; font-family: Arial, sans-serif; font-weight: bold; }
    .label { width: 6in; height: 4in; padding: 0.18in 0.25in 0.12in 0.25in; page-break-after: always; border: 2px solid #000; display: flex; flex-direction: column; }
    .label:last-child { page-break-after: auto; }
    .top-row { display: flex; justify-content: space-between; align-items: flex-start; }
    .top-left { flex: 1.4; }
    .manufacturer { font-size: 18pt; line-height: 1.05; text-transform: uppercase; }
    .push-line { margin-top: 0.03in; font-size: 20pt; }
    .totalSkids-line { margin-top: 0.03in; font-size: 16pt; }
    .top-right { flex: 1; text-align: center; }
    .top-barcode { width: 100%; height: 0.8in; margin-bottom: 0.02in; }
    .skid-id-text { font-size: 16pt; text-align: center; margin-top: 0.02in; }
    .top-barcode-img { width: 100%; height: 100%; object-fit: contain; display: block; }
    .bottom-barcode-img { width: 100%; height: 100%; object-fit: cover; display: block; }
    .middle-block { margin-top: 0.08in; }
    .line-fbpn { font-size: 40pt; line-height: 1.0; }
    .line-qty { font-size: 34pt; line-height: 1.0; margin-top: 0.03in; }
    .line-project { font-size: 24pt; line-height: 1.0; margin-top: 0.03in; }
    .bottom-block { margin-top: auto; text-align: center; padding-top: 0.04in; width: 100%; }
    .scan-to-pick-text { font-size: 12pt; margin-bottom: 0.05in; }
    /* Padding added to bottom barcode container as requested */
    .bottom-barcode { width: 100%; height: 0.7in; padding: 5px 25px; box-sizing: border-box; overflow: hidden; display: flex; justify-content: center; }
  </style></head><body>`;

  labelData.forEach(d => {
    // Top Right: SKU barcode
    const skuBarcodeUri = bwipPngDataUri_('code128', d.sku, { scale: 3, height: 15 });
    // Bottom: Skid ID barcode (with deep link QR if base URL is configured)
    const deepLink = base ? `${base}?skid=${encodeURIComponent(d.skidId)}` : d.skidId;
    const bottomCodeType = base ? 'qrcode' : 'code128';
    const skidBarcodeUri = bwipPngDataUri_(bottomCodeType, deepLink, { scale: 4, height: 14 });

    html += `
    <div class="label">
      <div class="top-row">
        <div class="top-left">
          <div class="manufacturer">${escapeHtml_(d.manufacturer)}</div>
          <div class="push-line">PUSH #: ${escapeHtml_(d.pushNumber)}</div>
          <div class="totalSkids-line">TOTAL SKIDS: ${escapeHtml_(d.totalSkids)}</div>
        </div>
        <div class="top-right">
          <div class="top-barcode"><img src="${skuBarcodeUri}" class="top-barcode-img"></div>
          <div class="skid-id-text">${escapeHtml_(d.sku)}</div>
        </div>
      </div>
      <div class="middle-block">
        <div class="line-fbpn">FBPN: ${escapeHtml_(d.fbpn)}</div>
        <div class="line-qty">Quantity: ${escapeHtml_(String(d.quantity))}</div>
        <div class="line-project">PROJECT: ${escapeHtml_(d.project)}</div>
      </div>
      <div class="bottom-block">
        <div class="scan-to-pick-text">Scan Skid ID</div>
        <div class="bottom-barcode"><img src="${skidBarcodeUri}" class="bottom-barcode-img" style="width:100%; height:100%;"></div>
        <div style="font-size: 10pt; margin-top: 2px;">${escapeHtml_(d.skidId)}</div>
      </div>
    </div>`;
  });
  return html + `</body></html>`;
}

function saveLabelsToDrive(htmlContent, folder, bolNumber) {
  const safeBol = String(bolNumber || 'NO_BOL').trim().replace(/[\/\\?%*:|"<>\.]/g, '_');

  // Check if file exists, append timestamp if so
  const iter = folder.getFilesByName(`Labels_${safeBol}.html`);
  let baseName = `Labels_${safeBol}`;
  if (iter.hasNext()) {
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    baseName = `Labels_${safeBol}_${timestamp}`;
  }

  const htmlBlob = Utilities.newBlob(htmlContent, 'text/html', `${baseName}.html`);
  const htmlFile = folder.createFile(htmlBlob);

  const pdfBlob = htmlBlob.getAs('application/pdf');
  pdfBlob.setName(`${baseName}.pdf`);
  const pdfFile = folder.createFile(pdfBlob);

  return { htmlFile, pdfFile };
}

function saveLabelsToDrive_(html, folder, bolNumber) {
  return saveLabelsToDrive(html, folder, bolNumber);
}

function bwipPngDataUri_(bcid, text, opts) {
  const o = opts || {};
  const params = [`bcid=${encodeURIComponent(bcid)}`, `text=${encodeURIComponent(text)}`, `scale=${o.scale || 3}`];
  if (o.height) params.push(`height=${o.height}`);
  const url = 'https://bwipjs-api.metafloor.com/?' + params.join('&');
  const resp = UrlFetchApp.fetch(url);
  return 'data:image/png;base64,' + Utilities.base64Encode(resp.getContent());
}

function escapeHtml_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function getSkidLookupBaseUrl_() {
  return PropertiesService.getScriptProperties().getProperty('SKID_LOOKUP_BASE_URL') || '';
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

function getOrCreateSubfolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

/**
 * ONE-TIME FUNCTION: Generate labels for ALL past inbounds found in Master_Log/Inbound_Skids.
 * Saves them to Inbound_Uploads/{MonthYear}/{Day}/{BOL}.
 * * Updated: Defaults to processing ALL unless range is specified.
 */
function generateLabelsForAllPastInbounds(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(TABS.MASTER_LOG);
  const skidsSheet = ss.getSheetByName(TABS.INBOUND_SKIDS);

  if (!masterSheet || !skidsSheet) throw new Error('Missing sheets.');

  const mData = masterSheet.getDataRange().getValues();
  const sData = skidsSheet.getDataRange().getValues();

  // Headers
  const mHead = mData[0];
  const sHead = sData[0];
  const mCol = (n) => mHead.indexOf(n);
  const sCol = (n) => sHead.indexOf(n);

  const mTxn = mCol('Txn_ID');
  const mDate = mCol('Date_Received');
  const mBol = mCol('BOL_Number');
  const mPush = mCol('Push #');
  const mMan = mCol('Manufacturer');

  const sTxn = sCol('TXN_ID');
  const sSkid = sCol('Skid_ID');
  const sFbpn = sCol('FBPN');
  const sQty = sCol('Qty_on_Skid');
  const sSku = sCol('SKU');
  const sProj = sCol('Project');
  const sSeq = sCol('Skid_Sequence');

  if (mTxn < 0 || mDate < 0) throw new Error('Master_Log missing Txn_ID or Date_Received');
  if (sTxn < 0) throw new Error('Inbound_Skids missing TXN_ID');

  // Filter dates if provided
  let start = null, end = null;
  if (startDate) start = normalizeDateOnly_(startDate);
  if (endDate) end = normalizeDateOnly_(endDate);

  // Group Master Info by Txn_ID
  const txns = {};
  for (let i = 1; i < mData.length; i++) {
    const txnId = (mData[i][mTxn] || '').toString().trim();
    if (!txnId) continue;

    if (start || end) {
      const d = normalizeDateOnly_(mData[i][mDate]);
      if (!d) continue;
      if (start && d < start) continue;
      if (end && d > end) continue;
    }

    if (!txns[txnId]) {
      txns[txnId] = {
        txnId: txnId,
        date: mData[i][mDate],
        bol: mBol >= 0 ? mData[i][mBol] : '',
        push: mPush >= 0 ? mData[i][mPush] : '',
        manufacturer: mMan >= 0 ? mData[i][mMan] : '',
        project: '',
        items: []
      };
    }
  }

  const txnIdsSet = Object.keys(txns);
  if (txnIdsSet.length === 0) {
    return `No transactions found to generate.`;
  }

  // Map Skids to Txns
  for (let i = 1; i < sData.length; i++) {
    const txnId = (sData[i][sTxn] || '').toString().trim();
    if (!txnId || !txns[txnId]) continue;

    txns[txnId].items.push({
      skidId: sSkid >= 0 ? sData[i][sSkid] : '',
      fbpn: sFbpn >= 0 ? sData[i][sFbpn] : '',
      qty: sQty >= 0 ? sData[i][sQty] : '',
      sku: sSku >= 0 ? sData[i][sSku] : '',
      project: sProj >= 0 ? sData[i][sProj] : '',
      skidSeq: sSeq >= 0 ? (sData[i][sSeq] || 1) : 1
    });
  }

  let count = 0;

  // Process each transaction
  for (const txnId in txns) {
    const txn = txns[txnId];
    if (txn.items.length === 0) continue;

    const totalSkids = txn.items.length;

    // Build Label Data
    const labelData = txn.items.map(item => ({
      skidId: item.skidId,
      fbpn: item.fbpn,
      quantity: item.qty,
      sku: item.sku,
      manufacturer: txn.manufacturer,
      project: item.project,
      pushNumber: txn.push,
      dateReceived: formatDate(parseInboundDate_(txn.date, new Date(txn.date))),
      skidNumber: item.skidSeq,
      totalSkids: totalSkids
    }));

    // Create folder structure if date is valid
    let targetFolder = null;
    try {
      const d = parseInboundDate_(txn.date, new Date());
      if (!isNaN(d.getTime())) {
        targetFolder = createInboundFolder_(d, txn.bol);
      }
    } catch (e) {
      Logger.log('Folder creation failed for ' + txnId + ': ' + e);
    }

    // Generate
    try {
      generateSkidLabels(labelData, { bolNumber: txn.bol, targetFolder: targetFolder });
      count++;
    } catch (e) {
      Logger.log('Failed to generate labels for ' + txnId + ': ' + e);
    }
  }

  return `Generated labels for ${count} transactions.`;
}

function normalizeDateOnly_(v) {
  if (!v) return null;
  const d = (Object.prototype.toString.call(v) === '[object Date]') ? v : new Date(v);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/* ===== FILE CONTENT END: IMS_Inbound_FIXED5.gs ===== */
