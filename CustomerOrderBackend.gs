
// ============================================================================
// INTERNAL: TOP-INSERT WRITER (Row 3 insert)
// ============================================================================
function _insertRowsAtTop_(sheet, rows) {
  if (!sheet) return;
  if (!rows || !rows.length) return;

  // Headers row 1, spacer row 2, data starts row 3
  sheet.insertRowsBefore(3, rows.length);
  sheet.getRange(3, 1, rows.length, rows[0].length).setValues(rows);
}

// ============================================================================
// MODAL FUNCTIONS
// ============================================================================

function debugCheckCustomerOrdersRoot() {
  const id = '1G97y64fxlq6rBd8RItHREmNxRMRrK-VV';
  const folder = DriveApp.getFolderById(id);
  Logger.log('Found folder: ' + folder.getName() + ' | ' + folder.getUrl());
}

function showCustomerOrderModal() {
  const html = HtmlService.createTemplateFromFile('CustomerOrderModal')
    .evaluate()
    .setWidth(900)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Customer Order');
}

function getCompanies() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const supportSheet = ss.getSheetByName('Support_Sheet');
  if (!supportSheet) return [];

  const data = supportSheet.getDataRange().getValues();
  const companies = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) companies.add(data[i][0]);
  }
  return Array.from(companies).sort();
}

function getProjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const supportSheet = ss.getSheetByName('Support_Sheet');
  if (!supportSheet) return [];

  const data = supportSheet.getDataRange().getValues();
  const projects = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) projects.add(data[i][1]);
  }
  return Array.from(projects).sort();
}

// ============================================================================
// ORDER PROCESSING
// ============================================================================

function processCustomerOrder(orderData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userEmail = Session.getActiveUser().getEmail();

    // Generate Order ID
    const orderId = generateOrderId();

    // Look up contact information (fallback if not in orderData)
    const contactInfo = lookupContactInfo(orderData.company, orderData.project);

    // Use provided data, fallback to lookup
    const finalName = orderData.name || contactInfo.name;
    const finalPhone = orderData.phoneNumber || contactInfo.phoneNumber;

    // Create Drive folder structure
    const folderInfo = createOrderFolder(orderData.taskNumber, orderData.company);

    // Upload file if provided (Modal upload)
    let fileUrl = '';
    if (orderData.fileData) {
      fileUrl = uploadOrderFile(folderInfo.folderId, orderData.fileData);
    }
    // If processed from existing Drive file (Automation), move it to order folder
    else if (orderData.sourceFileId) {
      try {
        const file = DriveApp.getFileById(orderData.sourceFileId);
        const targetFolder = DriveApp.getFolderById(folderInfo.folderId);
        const movedFile = file.moveTo(targetFolder);
        fileUrl = movedFile.getUrl();
      } catch (e) {
        Logger.log("Error moving source file: " + e);
        fileUrl = `File ID: ${orderData.sourceFileId} (Move Failed)`;
      }
    }

    // Check stock availability for all items (Global Stock_Totals check)
    const stockCheckResults = checkStockAvailability(orderData.items);

    // Determine overall stock status
    const overallStockStatus = determineOverallStockStatus(stockCheckResults);

    // Write to Customer_Orders
    writeCustomerOrder(ss, {
      orderId: orderId,
      taskNumber: orderData.taskNumber,
      project: orderData.project,
      nbd: orderData.nbdDate,
      company: orderData.company,
      orderTitle: orderData.orderTitle,
      deliverTo: orderData.deliveryLocation,
      name: finalName,
      phoneNumber: finalPhone,
      originalOrderUrl: fileUrl,
      orderFolderUrl: folderInfo.folderUrl,
      stockStatus: overallStockStatus,
      createdBy: userEmail
    });

    // Process each requested item
    const backorderedItems = []; // [{ fbpn, sku, backorderId }]
    const allocatedItems = [];   // [{ fbpn, sku, allocationId }]

    const pickLogItems = [];     // per bin+sku lines (for Pick_Log + OutboundLog)

    orderData.items.forEach(item => {
      const stockCheck = stockCheckResults.find(sc => sc.fbpn === item.fbpn);

      // Base SKU Logic (fallback if bins don't provide SKU)
      let baseSku = '';
      if (item.manufacturer) {
        const mfrCode = item.manufacturer.substring(0, 3).toUpperCase();
        baseSku = item.fbpn + '-' + mfrCode;
      } else {
        const existingSku = getSkuForFBPN(ss, item.fbpn);
        baseSku = existingSku || (item.fbpn + '-UNK');
      }

      // Description (prefer SKU-based description if exists)
      const baseDescription = lookupDescriptionBySku(ss, baseSku) || item.description || '';

      // 1) BIN-SELECTION NOW RETURNS PER (SKU + BIN) LINES
      //    This is the core fix: bins can carry different SKUs for same FBPN.
      const qtyAvail = Number(stockCheck.qtyAvailable) || 0;
      const binSkuAllocLines = (qtyAvail > 0)
        ? allocateInventoryToBinsBySku(ss, item.fbpn, qtyAvail, {
            fallbackSku: baseSku,
            fallbackManufacturer: item.manufacturer || ''
          })
        : [];

      // 2) WRITE REQUESTED_ITEMS:
      //    - One line per unique SKU that was selected from bins (sum across bins)
      //    - PLUS a backorder line (if any) using baseSku
      const perSkuTotals = aggregateAllocationsBySku_(binSkuAllocLines); // Map sku -> { qtyAllocated, manufacturer }
      perSkuTotals.forEach((v, sku) => {
        const desc = lookupDescriptionBySku(ss, sku) || baseDescription || '';
        writeRequestedItem(ss, {
          orderId: orderId,
          fbpn: item.fbpn,
          description: desc,
          qtyRequested: v.qtyAllocated,           // per-SKU allocated qty
          stockStatus: (v.qtyAllocated > 0) ? 'Allocated' : (stockCheck.status || ''),
          qtyBackordered: 0,
          qtyReserved: v.qtyAllocated,
          sku: sku
        });
      });

      // If no allocations but item exists, still write a single Requested_Items line (keeps behavior sane)
      if (qtyAvail > 0 && perSkuTotals.size === 0) {
        writeRequestedItem(ss, {
          orderId: orderId,
          fbpn: item.fbpn,
          description: baseDescription,
          qtyRequested: qtyAvail,
          stockStatus: 'Allocated',
          qtyBackordered: 0,
          qtyReserved: qtyAvail,
          sku: baseSku
        });
      }

      // Backorder line (keeps backorder quantity visible in Requested_Items)
      if (Number(stockCheck.qtyBackordered) > 0) {
        writeRequestedItem(ss, {
          orderId: orderId,
          fbpn: item.fbpn,
          description: baseDescription,
          qtyRequested: Number(stockCheck.qtyBackordered) || 0,
          stockStatus: stockCheck.status,
          qtyBackordered: Number(stockCheck.qtyBackordered) || 0,
          qtyReserved: 0,
          sku: baseSku
        });
      }

      // 3) CREATE BACKORDER(S) — per SKU (we only have one SKU for backorder: baseSku)
      if (Number(stockCheck.qtyBackordered) > 0) {
        const backorderId = createBackorder(ss, {
          orderId: orderId,
          taskNumber: orderData.taskNumber,
          nbd: orderData.nbdDate,
          fbpn: item.fbpn,
          qtyRequested: Number(stockCheck.qtyBackordered) || 0,
          qtyBackordered: Number(stockCheck.qtyBackordered) || 0,
          stockStatus: stockCheck.status,
          sku: baseSku
        });
        backorderedItems.push({ fbpn: item.fbpn, sku: baseSku, backorderId: backorderId });
      }

      // 4) CREATE ALLOCATION(S) — one per unique SKU that was selected from bins
      perSkuTotals.forEach((v, sku) => {
        if ((Number(v.qtyAllocated) || 0) <= 0) return;

        const allocationId = createAllocation(ss, {
          orderId: orderId,
          fbpn: item.fbpn,
          qtyRequested: v.qtyAllocated,
          qtyAllocated: v.qtyAllocated,
          qtyBackordered: 0,
          allocatedBy: userEmail,
          backorderId: backorderedItems.find(bi => bi.fbpn === item.fbpn && bi.sku === baseSku)?.backorderId || '',
          sku: sku
        });
        allocatedItems.push({ fbpn: item.fbpn, sku: sku, allocationId: allocationId });
      });

      // If allocated but no perSkuTotals (fallback)
      if (qtyAvail > 0 && perSkuTotals.size === 0) {
        const allocationId = createAllocation(ss, {
          orderId: orderId,
          fbpn: item.fbpn,
          qtyRequested: qtyAvail,
          qtyAllocated: qtyAvail,
          qtyBackordered: 0,
          allocatedBy: userEmail,
          backorderId: backorderedItems.find(bi => bi.fbpn === item.fbpn && bi.sku === baseSku)?.backorderId || '',
          sku: baseSku
        });
        allocatedItems.push({ fbpn: item.fbpn, sku: baseSku, allocationId: allocationId });
      }

      // 5) PICK_LOG LINES — per bin + sku (already desired)
      //    Also store manufacturer for OutboundLog.
      binSkuAllocLines.forEach(a => {
        pickLogItems.push({
          fbpn: item.fbpn,
          description: lookupDescriptionBySku(ss, a.sku) || baseDescription || '',
          qtyTotalRequested: item.qty,
          qtyToPick: a.qtyToPick,
          binCode: a.binCode,
          sku: a.sku,
          manufacturer: a.manufacturer || item.manufacturer || ''
        });
      });

      // 6) Stock_Totals Integration (keep existing behavior: baseSku)
      //    If you want per-SKU stock totals updates, we can expand this later.
      if (baseSku) {
        try {
          updateStockTotals_OrderAllocation(
            baseSku,
            stockCheck.qtyAvailable,    // Qty allocated (reserved)
            stockCheck.qtyBackordered   // Qty backordered
          );
        } catch (e) {
          Logger.log(`Warning: Could not update Stock_Totals for ${baseSku}: ${e.toString()}`);
        }
      }
    });

    // Update Requested_Items IDs (now matches by Order_ID + FBPN + SKU)
    updateRequestedItemsWithIds(ss, orderId, backorderedItems, allocatedItems);

    // Pick_Log (detail rows)
    if (pickLogItems.length > 0) {
      addItemsToPickLog(ss, {
        nbd: orderData.nbdDate,
        orderId: orderId,
        taskNumber: orderData.taskNumber,
        company: orderData.company,
        project: orderData.project,
        items: pickLogItems
      });
      // Note: OutboundLog is written during outbound processing (after TOC/Packing slips)
      // in processPackingTOCAndShipment, not during order creation
    }

    const inStockCount = stockCheckResults.filter(sc => sc.status === 'In Stock').length;
    const backorderedCount = stockCheckResults.filter(sc => (Number(sc.qtyBackordered) || 0) > 0).length;

    return {
      success: true,
      orderId: orderId,
      totalItems: orderData.items.length,
      inStockCount: inStockCount,
      backorderedCount: backorderedCount,
      folderUrl: folderInfo.folderUrl
    };

  } catch (error) {
    Logger.log('Error in processCustomerOrder: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

// ============================================================================
// HELPER: Aggregate allocations by SKU (sum qty across bins)
// ============================================================================
function aggregateAllocationsBySku_(allocLines) {
  const out = new Map(); // sku -> { qtyAllocated, manufacturer }
  (allocLines || []).forEach(a => {
    const sku = String(a.sku || '').trim();
    if (!sku) return;
    const qty = Number(a.qtyToPick) || 0;
    if (qty <= 0) return;
    const prev = out.get(sku) || { qtyAllocated: 0, manufacturer: '' };
    prev.qtyAllocated += qty;
    if (!prev.manufacturer && a.manufacturer) prev.manufacturer = a.manufacturer;
    out.set(sku, prev);
  });
  return out;
}

// ============================================================================
// HELPER: Add Items to Pick_Log (Detail Rows)
// ============================================================================

function addItemsToPickLog(ss, data) {
  const sheet = ss.getSheetByName('Pick_Log');
  if (!sheet) return;

  const timestamp = new Date();
  const rows = [];

  data.items.forEach(item => {
    const pikId = 'PIK-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const binCode = item.binCode || 'N/A';
    const qtyForLog = item.qtyToPick || item.qtyTotalRequested;

    rows.push([
      pikId,                 // PIK_ID
      data.nbd,              // NBD
      data.orderId,          // Order_Number
      data.taskNumber,       // Task_Number
      data.company,          // Company
      data.project,          // Project
      item.fbpn,             // FBPN
      item.description,      // Description
      item.qtyTotalRequested || qtyForLog, // Total Request
      qtyForLog,             // Qty To Pick (bin-specific)
      binCode,               // Bin_Code
      0,                     // Qty_Picked
      'PENDING',             // Status
      '',                    // Picked_By
      '',                    // Shipped_Date
      timestamp,             // Timestamp
      item.sku               // SKU
    ]);
  });

  if (rows.length > 0) {
    _insertRowsAtTop_(sheet, rows);
  }
}

// ============================================================================
// HELPER: OUTBOUNDLOG WRITER — PER BIN + SKU (fixed schema provided)
// Headers: Date Order_Number Task_Number Transaction Type Warehouse Company Project
//          FBPN Manufacturer Qty Bin_Code Skid_ID SKU
// ============================================================================
function addItemsToOutboundLog(ss, data) {
  const sh = ss.getSheetByName('OutboundLog');
  if (!sh) return;

  const rows = [];
  const dt = data.date || new Date();
  const wh = data.warehouse || '';

  (data.items || []).forEach(item => {
    const qty = Number(item.qtyToPick) || 0;
    if (qty <= 0) return;

    rows.push([
      dt,                     // Date
      data.orderId,           // Order_Number
      data.taskNumber,        // Task_Number
      'Outbound',             // Transaction Type
      wh,                     // Warehouse
      data.company,           // Company
      data.project,           // Project
      item.fbpn,              // FBPN
      item.manufacturer || '',// Manufacturer
      qty,                    // Qty
      item.binCode || '',     // Bin_Code
      '',                     // Skid_ID
      item.sku || ''          // SKU
    ]);
  });

  if (rows.length > 0) {
    _insertRowsAtTop_(sh, rows);
  }
}

// ============================================================================
// HELPER: Lookup Description by SKU
// ============================================================================
function lookupDescriptionBySku(ss, sku) {
  const pmSheet = ss.getSheetByName('Project_Master');
  if (!pmSheet) return '';

  const data = pmSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][6]).trim() === String(sku).trim()) { // Column G = SKU
      return data[i][3]; // Column D = Description
    }
  }
  return '';
}

// ============================================================================
// HELPER: Get SKU for FBPN (Legacy/Fallback)
// ============================================================================
function getSkuForFBPN(ss, fbpn) {
  const stockTotalsSheet = ss.getSheetByName('Stock_Totals');
  if (!stockTotalsSheet) return '';

  const data = stockTotalsSheet.getDataRange().getValues();
  const headers = data[0];

  const skuCol = headers.indexOf('SKU');
  const fbpnCol = headers.indexOf('FBPN');

  if (skuCol === -1 || fbpnCol === -1) return '';

  for (let i = 1; i < data.length; i++) {
    if (data[i][fbpnCol] === fbpn) {
      return data[i][skuCol] || '';
    }
  }
  return '';
}

// ============================================================================
// GOOGLE DRIVE FOLDER MANAGEMENT
// ============================================================================
function createOrderFolder(taskNumber, company) {
  const rootFolderId = '1G97y64fxlq6rBd8RItHREmNxRMRrK-VV';
  const rootFolder = DriveApp.getFolderById(rootFolderId);

  let companyFolder;
  const companyFolders = rootFolder.getFoldersByName(company);
  if (companyFolders.hasNext()) companyFolder = companyFolders.next();
  else companyFolder = rootFolder.createFolder(company);

  const taskFolder = companyFolder.createFolder(taskNumber);

  return {
    folderId: taskFolder.getId(),
    folderUrl: taskFolder.getUrl()
  };
}

function uploadOrderFile(folderId, fileData) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.data),
      fileData.mimeType,
      fileData.name
    );
    const file = folder.createFile(blob);
    return file.getUrl();
  } catch (error) {
    Logger.log('Error uploading file: ' + error.toString());
    return '';
  }
}

// ============================================================================
// STOCK CHECKING
// ============================================================================
function checkStockAvailability(items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Build FBPN -> total available from BIN SOURCES ONLY
  const totalsByFbpn = new Map();

  const addFromSheet = (tabName, fallback) => {
    const sh = ss.getSheetByName(tabName);
    if (!sh) return;
    const values = sh.getDataRange().getValues();
    if (!values || values.length < 2) return;

    const headers = values[0].map(h => String(h || ''));
    const colFbpn = safeColIndex_(headers, ['fbpn', 'part_number', 'part number'], fallback.fbpn);
    const colQty  = safeColIndex_(headers, ['current_quantity','current qty','current quantity','qty','quantity','qty_on_hand','qty on hand'], fallback.qty);

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const fbpn = String(row[colFbpn] || '').trim();
      if (!fbpn) continue;
      const qty = Number(row[colQty]) || 0;
      if (qty <= 0) continue;

      const key = fbpn.toUpperCase();
      totalsByFbpn.set(key, (totalsByFbpn.get(key) || 0) + qty);
    }
  };

  // NOTE: fallback indices match your previous allocation helper defaults
  addFromSheet('Bin_Stock',          { fbpn: 3, qty: 7 });
  addFromSheet('Floor_Stock_Levels', { fbpn: 3, qty: 7 });
  addFromSheet('Inbound_Staging',    { fbpn: 3, qty: 7 }); // ✅ added

  return items.map(item => {
    const fbpnKey = String(item.fbpn || '').trim().toUpperCase();
    const availableQty = totalsByFbpn.get(fbpnKey) || 0;
    const qtyRequested = Number(item.qty) || 0;

    let status, qtyAvailable, qtyBackordered;
    if (availableQty >= qtyRequested) {
      status = 'In Stock';
      qtyAvailable = qtyRequested;
      qtyBackordered = 0;
    } else if (availableQty > 0) {
      status = 'Partial';
      qtyAvailable = availableQty;
      qtyBackordered = qtyRequested - availableQty;
    } else {
      status = 'Backordered';
      qtyAvailable = 0;
      qtyBackordered = qtyRequested;
    }

    return {
      fbpn: String(item.fbpn || '').trim(),
      qtyRequested: qtyRequested,
      status: status,
      qtyAvailable: qtyAvailable,
      qtyBackordered: qtyBackordered
    };
  });
}

function determineOverallStockStatus(stockCheckResults) {
  const hasBackorders = stockCheckResults.some(r => r.status === 'Backordered');
  const hasPartials = stockCheckResults.some(r => r.status === 'Partial');
  if (hasBackorders) return 'Contains Backorders';
  if (hasPartials) return 'Contains Partials';
  return 'In Stock';
}

// ============================================================================
// DATABASE WRITES
// ============================================================================
function writeCustomerOrder(ss, orderData) {
  const sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet) throw new Error('Customer_Orders sheet not found');

  const timestamp = new Date();

  const row = [[
    orderData.orderId,
    orderData.taskNumber,
    orderData.project,
    orderData.nbd,
    'Accepted',
    orderData.stockStatus,
    orderData.company,
    orderData.orderTitle,
    orderData.deliverTo,
    orderData.name,
    orderData.phoneNumber,
    '', // Original_Order
    '', // Order_Folder
    '', // Pick_Ticket_PDF
    '', // TOC_PDF
    '', // Packing_Lists
    timestamp,
    orderData.createdBy
  ]];

  _insertRowsAtTop_(sheet, row);

  // New row is always row 3 after insert
  const insertedRow = 3;

  if (orderData.originalOrderUrl) {
    sheet.getRange(insertedRow, 12).setFormula(`=HYPERLINK("${orderData.originalOrderUrl}", "Original Order")`);
  }
  if (orderData.orderFolderUrl) {
    sheet.getRange(insertedRow, 13).setFormula(`=HYPERLINK("${orderData.orderFolderUrl}", "Order Folder")`);
  }
}

function writeRequestedItem(ss, itemData) {
  const sheet = ss.getSheetByName('Requested_Items');
  if (!sheet) throw new Error('Requested_Items sheet not found');

  const row = [[
    itemData.orderId,
    itemData.fbpn,
    itemData.description || '',
    itemData.qtyRequested,
    itemData.stockStatus,
    itemData.qtyBackordered,
    itemData.qtyReserved,
    0,   // Qty_Shipped
    '',  // Backorder_ID
    '',  // Allocation_ID
    itemData.sku || ''
  ]];

  _insertRowsAtTop_(sheet, row);
}

function createBackorder(ss, backorderData) {
  const sheet = ss.getSheetByName('Backorders');
  if (!sheet) throw new Error('Backorders sheet not found');

  const backorderId = 'BO-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const timestamp = new Date();

  const row = [[
    backorderData.orderId,
    backorderData.nbd,
    'Open',
    backorderData.taskNumber,
    backorderData.stockStatus,
    backorderData.fbpn,
    backorderData.qtyRequested,
    backorderData.qtyBackordered,
    0,
    timestamp,
    '',
    '',
    backorderId,
    backorderData.sku || ''
  ]];

  _insertRowsAtTop_(sheet, row);

  return backorderId;
}

function createAllocation(ss, allocationData) {
  const sheet = ss.getSheetByName('Allocation_Log');
  if (!sheet) throw new Error('Allocation_Log sheet not found');

  const allocationId = 'ALLOC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const timestamp = new Date();
  const allocationStatus = allocationData.qtyAllocated > 0 ? 'Ready to Pick' : 'Pending';

  const row = [[
    allocationData.orderId,
    timestamp,
    allocationStatus,
    allocationData.fbpn,
    allocationData.qtyRequested,
    allocationData.qtyAllocated,
    allocationData.qtyBackordered,
    allocationData.allocatedBy,
    allocationData.backorderId || '',
    allocationId,
    allocationData.sku || ''
  ]];

  _insertRowsAtTop_(sheet, row);

  return allocationId;
}

// UPDATED: match by Order_ID + FBPN + SKU, and set Backorder_ID when Qty_Backordered > 0
function updateRequestedItemsWithIds(ss, orderId, backorderedItems, allocatedItems) {
  const sheet = ss.getSheetByName('Requested_Items');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const orderIdCol = headers.indexOf('Order_ID');
  const fbpnCol = headers.indexOf('FBPN');
  const skuCol = headers.indexOf('SKU');
  const qtyBackCol = headers.indexOf('Qty_Backordered');
  const backorderIdCol = headers.indexOf('Backorder_ID');
  const allocationIdCol = headers.indexOf('Allocation_ID');

  for (let i = 1; i < data.length; i++) {
    if (data[i][orderIdCol] !== orderId) continue;

    const fbpn = String(data[i][fbpnCol] || '').trim();
    const sku = (skuCol > -1) ? String(data[i][skuCol] || '').trim() : '';
    const qtyBack = (qtyBackCol > -1) ? (Number(data[i][qtyBackCol]) || 0) : 0;

    if (qtyBack > 0) {
      const bo = backorderedItems.find(bi => bi.fbpn === fbpn && bi.sku === sku);
      if (bo && backorderIdCol !== -1) {
        sheet.getRange(i + 1, backorderIdCol + 1).setValue(bo.backorderId);
      }
    } else {
      const al = allocatedItems.find(ai => ai.fbpn === fbpn && ai.sku === sku);
      if (al && allocationIdCol !== -1) {
        sheet.getRange(i + 1, allocationIdCol + 1).setValue(al.allocationId);
      }
    }
  }
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================
function generateOrderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet) return 'Order-1001';

  const data = sheet.getDataRange().getValues();
  let maxId = 1000;

  for (let i = 1; i < data.length; i++) {
    const orderId = data[i][0];
    if (orderId && typeof orderId === 'string' && orderId.startsWith('Order-')) {
      const num = parseInt(orderId.replace('Order-', ''));
      if (!isNaN(num) && num > maxId) maxId = num;
    }
  }
  return 'Order-' + (maxId + 1);
}

function lookupContactInfo(company, project) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const supportSheet = ss.getSheetByName('Support_Sheet');
  if (!supportSheet) return { name: '', phoneNumber: '' };

  const data = supportSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === company && data[i][1] === project) {
      return { name: data[i][2] || '', phoneNumber: data[i][3] || '' };
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === company) {
      return { name: data[i][2] || '', phoneNumber: data[i][3] || '' };
    }
  }
  return { name: '', phoneNumber: '' };
}

// ============================================================================
// 5. BIN ALLOCATION HELPERS
// A) Legacy (kept): allocateInventoryToBins()
// B) NEW: allocateInventoryToBinsBySku() — returns per (SKU + BIN) lines
// ============================================================================

function allocateInventoryToBins(ss, fbpn, qtyToAllocate) {
  const qtyNeededOriginal = Number(qtyToAllocate) || 0;
  const fbpnNorm = String(fbpn || '').trim().toUpperCase();
  if (!fbpnNorm || qtyNeededOriginal <= 0) return [];

  const sources = [
    { tab: 'Bin_Stock',          fallback: { bin: 0, fbpn: 3, qty: 7 } },
    { tab: 'Floor_Stock_Levels', fallback: { bin: 0, fbpn: 3, qty: 7 } },
    { tab: 'Inbound_Staging',    fallback: { bin: 0, fbpn: 3, qty: 7 } } // ✅ added
  ];

  // Aggregate availability: BIN_CODE -> qtyAvail
  const byBin = new Map();

  const addBinQty = (binCode, qtyAvail) => {
    const key = String(binCode || '').trim();
    if (!key) return;
    const qty = Number(qtyAvail) || 0;
    if (qty <= 0) return;
    byBin.set(key, (byBin.get(key) || 0) + qty);
  };

  sources.forEach(src => {
    const sh = ss.getSheetByName(src.tab);
    if (!sh) return;

    const values = sh.getDataRange().getValues();
    if (!values || values.length < 2) return;

    const headers = values[0].map(h => String(h || ''));
    const colBin  = safeColIndex_(headers, ['bin_code','bin code','bin','binlocation','bin_location'], src.fallback.bin);
    const colFbpn = safeColIndex_(headers, ['fbpn','part_number','part number'], src.fallback.fbpn);
    const colQty  = safeColIndex_(headers, ['current_quantity','current qty','current quantity','qty','quantity','qty_on_hand','qty on hand'], src.fallback.qty);

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const rowFbpn = String(row[colFbpn] || '').trim().toUpperCase();
      if (!rowFbpn) continue;
      if (rowFbpn !== fbpnNorm) continue;

      const binCode = String(row[colBin] || '').trim();
      if (!binCode) continue;

      const qtyAvail = Number(row[colQty]) || 0;
      if (qtyAvail <= 0) continue;

      addBinQty(binCode, qtyAvail);
    }
  });

  if (byBin.size === 0) {
    return [{ binCode: 'UNASSIGNED', qtyToPick: qtyNeededOriginal }];
  }

  const merged = Array.from(byBin.entries()).map(([binCode, qtyAvail]) => ({ binCode, qtyAvail }));

  // Prefer bins with more stock first (stable tie-breaker)
  merged.sort((a, b) => {
    if (b.qtyAvail !== a.qtyAvail) return b.qtyAvail - a.qtyAvail;
    return a.binCode.localeCompare(b.binCode);
  });

  const allocations = [];
  let remaining = qtyNeededOriginal;

  for (let i = 0; i < merged.length && remaining > 0; i++) {
    const take = Math.min(remaining, merged[i].qtyAvail);
    if (take > 0) {
      allocations.push({ binCode: merged[i].binCode, qtyToPick: take });
      remaining -= take;
    }
  }

  if (remaining > 0) allocations.push({ binCode: 'UNASSIGNED', qtyToPick: remaining });
  return allocations;
}

/**
 * NEW: allocateInventoryToBinsBySku
 * Reads bin sources and returns allocations that preserve SKU differences per bin.
 * Output: [{ binCode, qtyToPick, sku, manufacturer }]
 *
 * Assumes bin sources may have a SKU column and optionally Manufacturer.
 * If missing, falls back to provided fallbackSku / fallbackManufacturer.
 */
function allocateInventoryToBinsBySku(ss, fbpn, qtyToAllocate, opts) {
  const qtyNeededOriginal = Number(qtyToAllocate) || 0;
  const fbpnNorm = String(fbpn || '').trim().toUpperCase();
  if (!fbpnNorm || qtyNeededOriginal <= 0) return [];

  const fallbackSku = (opts && opts.fallbackSku) ? String(opts.fallbackSku) : '';
  const fallbackMfr = (opts && opts.fallbackManufacturer) ? String(opts.fallbackManufacturer) : '';

  const sources = [
    { tab: 'Bin_Stock',          fallback: { bin: 0, fbpn: 3, qty: 7 } },
    { tab: 'Floor_Stock_Levels', fallback: { bin: 0, fbpn: 3, qty: 7 } },
    { tab: 'Inbound_Staging',    fallback: { bin: 0, fbpn: 3, qty: 7 } }
  ];

  // Aggregate availability by (SKU + BIN)
  // key = sku||bin
  const bySkuBin = new Map();

  const addSkuBinQty = (sku, binCode, qtyAvail, manufacturer) => {
    const s = String(sku || '').trim();
    const b = String(binCode || '').trim();
    if (!b) return;
    const q = Number(qtyAvail) || 0;
    if (q <= 0) return;

    const finalSku = s || fallbackSku || (fbpnNorm + '-UNK');
    const mfr = String(manufacturer || '').trim() || fallbackMfr || '';

    const key = finalSku + '||' + b;
    const prev = bySkuBin.get(key) || { sku: finalSku, binCode: b, qtyAvail: 0, manufacturer: mfr };
    prev.qtyAvail += q;
    if (!prev.manufacturer && mfr) prev.manufacturer = mfr;
    bySkuBin.set(key, prev);
  };

  sources.forEach(src => {
    const sh = ss.getSheetByName(src.tab);
    if (!sh) return;

    const values = sh.getDataRange().getValues();
    if (!values || values.length < 2) return;

    const headers = values[0].map(h => String(h || ''));

    const colBin  = safeColIndex_(headers, ['bin_code','bin code','bin','binlocation','bin_location'], src.fallback.bin);
    const colFbpn = safeColIndex_(headers, ['fbpn','part_number','part number'], src.fallback.fbpn);
    const colQty  = safeColIndex_(headers, ['current_quantity','current qty','current quantity','qty','quantity','qty_on_hand','qty on hand'], src.fallback.qty);

    // Optional SKU/Manufacturer columns if they exist in bin sources
    const colSku = safeColIndexOptional_(headers, ['sku','item_sku','item sku']);
    const colMfr = safeColIndexOptional_(headers, ['manufacturer','mfr','vendor','brand']);

    for (let r = 1; r < values.length; r++) {
      const row = values[r];

      const rowFbpn = String(row[colFbpn] || '').trim().toUpperCase();
      if (!rowFbpn || rowFbpn !== fbpnNorm) continue;

      const binCode = String(row[colBin] || '').trim();
      if (!binCode) continue;

      const qtyAvail = Number(row[colQty]) || 0;
      if (qtyAvail <= 0) continue;

      const rowSku = (colSku > -1) ? String(row[colSku] || '').trim() : '';
      const rowMfr = (colMfr > -1) ? String(row[colMfr] || '').trim() : '';

      addSkuBinQty(rowSku, binCode, qtyAvail, rowMfr);
    }
  });

  if (bySkuBin.size === 0) {
    return [{
      binCode: 'UNASSIGNED',
      qtyToPick: qtyNeededOriginal,
      sku: fallbackSku || (fbpnNorm + '-UNK'),
      manufacturer: fallbackMfr || ''
    }];
  }

  const merged = Array.from(bySkuBin.values());

  // Prefer larger availability (stable)
  merged.sort((a, b) => {
    if (b.qtyAvail !== a.qtyAvail) return b.qtyAvail - a.qtyAvail;
    const s = a.sku.localeCompare(b.sku);
    if (s !== 0) return s;
    return a.binCode.localeCompare(b.binCode);
  });

  const allocations = [];
  let remaining = qtyNeededOriginal;

  for (let i = 0; i < merged.length && remaining > 0; i++) {
    const take = Math.min(remaining, merged[i].qtyAvail);
    if (take > 0) {
      allocations.push({
        binCode: merged[i].binCode,
        qtyToPick: take,
        sku: merged[i].sku,
        manufacturer: merged[i].manufacturer || ''
      });
      remaining -= take;
    }
  }

  if (remaining > 0) {
    allocations.push({
      binCode: 'UNASSIGNED',
      qtyToPick: remaining,
      sku: fallbackSku || (fbpnNorm + '-UNK'),
      manufacturer: fallbackMfr || ''
    });
  }

  return allocations;
}

function safeColIndex_(headers, possibleNames, fallbackIndex) {
  let idx = -1;

  try {
    if (typeof findColumnIndex === 'function') {
      idx = findColumnIndex(headers, possibleNames);
    } else {
      const normalizedHeaders = headers.map(h => String(h || '').toLowerCase().trim().replace(/[_\s]+/g, '_'));
      for (const name of possibleNames) {
        const normalizedName = String(name || '').toLowerCase().trim().replace(/[_\s]+/g, '_');
        const found = normalizedHeaders.indexOf(normalizedName);
        if (found >= 0) { idx = found; break; }
      }
    }
  } catch (e) {
    idx = -1;
  }

  if (idx == null || idx < 0) return fallbackIndex;
  return idx;
}

function safeColIndexOptional_(headers, possibleNames) {
  try {
    if (typeof findColumnIndex === 'function') {
      return findColumnIndex(headers, possibleNames);
    }
    const normalizedHeaders = headers.map(h => String(h || '').toLowerCase().trim().replace(/[_\s]+/g, '_'));
    for (const name of possibleNames) {
      const normalizedName = String(name || '').toLowerCase().trim().replace(/[_\s]+/g, '_');
      const found = normalizedHeaders.indexOf(normalizedName);
      if (found >= 0) return found;
    }
    return -1;
  } catch (e) {
    return -1;
  }
}

// ============================================================================
// CANCELLATION FUNCTIONS (with Stock_Totals integration)
// UPDATED: uses SKU column from Allocation_Log/Backorders if present
// ============================================================================

function cancelOrder(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Starting cancellation for Order ID: ' + orderId);

  try {
    // 1. Process Allocation_Log
    const allocSheet = ss.getSheetByName('Allocation_Log');
    if (allocSheet) {
      const data = allocSheet.getDataRange().getValues();
      const headers = data[0];
      const orderIdCol = headers.indexOf('Order_ID');
      const fbpnCol = headers.indexOf('FBPN');
      const qtyAllocCol = headers.indexOf('Qty_Allocated');
      const skuCol = headers.indexOf('SKU');

      const rowsToDelete = [];
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][orderIdCol]) === String(orderId)) {
          const fbpn = data[i][fbpnCol];
          const qtyAllocated = Number(data[i][qtyAllocCol]) || 0;
          const rowSku = (skuCol > -1) ? String(data[i][skuCol] || '').trim() : '';

          if (qtyAllocated > 0) {
            const sku = rowSku || getSkuForFBPN(ss, fbpn);
            if (sku) {
              try {
                if (typeof updateStockTotals_CancelAllocation === 'function') {
                  updateStockTotals_CancelAllocation(sku, qtyAllocated);
                  Logger.log(`Released allocation: ${qtyAllocated} of ${sku}`);
                }
              } catch (e) { Logger.log(`Warning: Could not release Stock_Totals allocation: ${e.toString()}`); }
            }
          }
          rowsToDelete.push(i + 1);
        }
      }
      rowsToDelete.sort((a, b) => b - a).forEach(r => allocSheet.deleteRow(r));
    }

    // 2. Process Backorders
    const backSheet = ss.getSheetByName('Backorders');
    if (backSheet) {
      const data = backSheet.getDataRange().getValues();
      const headers = data[0];
      const orderIdCol = headers.indexOf('Order_ID');
      const fbpnCol = headers.indexOf('FBPN');
      const qtyBackCol = headers.indexOf('Qty_Backordered');
      const qtyFulCol = headers.indexOf('Qty_Fulfilled');
      const skuCol = headers.indexOf('SKU');

      const rowsToDelete = [];
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][orderIdCol]) === String(orderId)) {
          const fbpn = data[i][fbpnCol];
          const remaining = (Number(data[i][qtyBackCol]) || 0) - (Number(data[i][qtyFulCol]) || 0);
          const rowSku = (skuCol > -1) ? String(data[i][skuCol] || '').trim() : '';

          if (remaining > 0) {
            const sku = rowSku || getSkuForFBPN(ss, fbpn);
            if (sku) {
              try {
                if (typeof updateStockTotals_CancelBackorder === 'function') {
                  updateStockTotals_CancelBackorder(sku, remaining);
                }
              } catch (e) { Logger.log(`Warning: Could not cancel backorder in Stock_Totals: ${e.toString()}`); }
            }
          }
          rowsToDelete.push(i + 1);
        }
      }
      rowsToDelete.sort((a, b) => b - a).forEach(r => backSheet.deleteRow(r));
    }

    // 3. Process Requested_Items
    const reqSheet = ss.getSheetByName('Requested_Items');
    if (reqSheet) {
      const data = reqSheet.getDataRange().getValues();
      const orderIdCol = data[0].indexOf('Order_ID');
      const rowsToDelete = [];
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][orderIdCol]) === String(orderId)) rowsToDelete.push(i + 1);
      }
      rowsToDelete.sort((a, b) => b - a).forEach(r => reqSheet.deleteRow(r));
    }

    // 4. Process Pick_Log
    const pickSheet = ss.getSheetByName('Pick_Log');
    if (pickSheet) {
      const data = pickSheet.getDataRange().getValues();
      let orderIdCol = data[0].indexOf('Order_Number');
      if (orderIdCol === -1) orderIdCol = data[0].indexOf('Order_ID');

      if (orderIdCol > -1) {
        const rowsToDelete = [];
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][orderIdCol]) === String(orderId)) rowsToDelete.push(i + 1);
        }
        rowsToDelete.sort((a, b) => b - a).forEach(r => pickSheet.deleteRow(r));
      }
    }

    // 4b. Process OutboundLog (matches your schema: Order_Number header)
    const outSheet = ss.getSheetByName('OutboundLog');
    if (outSheet) {
      const data = outSheet.getDataRange().getValues();
      const headers = data[0] || [];
      const orderCol = headers.indexOf('Order_Number');
      if (orderCol > -1) {
        const rowsToDelete = [];
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][orderCol]) === String(orderId)) rowsToDelete.push(i + 1);
        }
        rowsToDelete.sort((a, b) => b - a).forEach(r => outSheet.deleteRow(r));
      }
    }

    // 5. Update Customer_Orders Status
    const ordersSheet = ss.getSheetByName('Customer_Orders');
    if (ordersSheet) {
      const data = ordersSheet.getDataRange().getValues();
      const headers = data[0];
      const orderIdCol = headers.indexOf('Order_ID');
      const statusCol = headers.indexOf('Request_Status');

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][orderIdCol]) === String(orderId)) {
          ordersSheet.getRange(i + 1, statusCol + 1).setValue('Cancelled');
          break;
        }
      }
    }

    return { success: true, message: `Order ${orderId} cancelled.` };

  } catch (error) {
    Logger.log('Error cancelling order: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ============================================================================
// AUTOMATION: PROCESS UPLOADED ORDER FILES
// ============================================================================

function processPendingOrderUploads() {
  const CONFIG = (typeof getIMSConfig === 'function') ? getIMSConfig() : {};
  const sourceId = (CONFIG.SOURCE_NEW_ORDERS_FOLDER_ID) ? CONFIG.SOURCE_NEW_ORDERS_FOLDER_ID : AUTOMATION_FOLDER_ID;

  if (!sourceId) return;

  let sourceFolder;
  try {
    sourceFolder = DriveApp.getFolderById(sourceId);
  } catch (e) { return; }

  let processedFolder;
  const pIter = sourceFolder.getFoldersByName('Processed_Uploads');
  if (pIter.hasNext()) processedFolder = pIter.next();
  else processedFolder = sourceFolder.createFolder('Processed_Uploads');

  const fileTypes = [MimeType.MICROSOFT_EXCEL, MimeType.MICROSOFT_EXCEL_LEGACY];

  fileTypes.forEach(function (type) {
    const files = sourceFolder.getFilesByType(type);
    while (files.hasNext()) {
      const file = files.next();
      try {
        processSingleOrderFile_(file, processedFolder);
      } catch (e) {
        Logger.log('Failed to process file ' + file.getName() + ': ' + e.toString());
      }
    }
  });
}

function processSingleOrderFile_(file, processedFolder) {
  Logger.log('Processing file: ' + file.getName());

  const blob = file.getBlob();
  let tempFile;
  try {
    if (typeof Drive === 'undefined') throw new Error('Advanced Drive Service not enabled.');

    if (typeof Drive.Files.insert === 'function') {
      const resource = {
        title: '[TEMP] ' + file.getName(),
        parents: [{ id: processedFolder.getId() }],
        mimeType: MimeType.GOOGLE_SHEETS
      };
      tempFile = Drive.Files.insert(resource, blob);
    }
    else if (typeof Drive.Files.create === 'function') {
      const resource = {
        name: '[TEMP] ' + file.getName(),
        parents: [processedFolder.getId()],
        mimeType: MimeType.GOOGLE_SHEETS
      };
      tempFile = Drive.Files.create(resource, blob);
    } else {
      throw new Error('Drive API methods not found.');
    }
  } catch (e) {
    Logger.log('Conversion failed: ' + e.toString());
    return;
  }

  const ss = SpreadsheetApp.openById(tempFile.id);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  const orderData = {
    items: [],
    company: '',
    project: '',
    taskNumber: '',
    nbdDate: '',
    orderTitle: 'Upload: ' + file.getName(),
    sourceFileId: file.getId(),
    addToPickLog: true // ✅ REQUIRED: triggers Pick_Log writes
  };

  let itemHeaderRowIndex = -1;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    if (String(row[1]).trim().toLowerCase() === 'fbpn' || String(row[0]).trim().toLowerCase() === 'line number') {
      itemHeaderRowIndex = i;
      break;
    }

    [0, 2].forEach(keyIdx => {
      if (row.length <= keyIdx + 1) return;
      const key = String(row[keyIdx]).trim().replace(':', '').toLowerCase();
      const val = String(row[keyIdx + 1]).trim();
      if (!key || !val) return;

      if (key.includes('task')) orderData.taskNumber = val;
      else if (key.includes('company')) orderData.company = val;
      else if (key.includes('project')) orderData.project = val;
      else if (key.includes('title')) orderData.orderTitle = val;
      else if (key.includes('deliver')) orderData.deliveryLocation = val;
      else if (key.includes('phone') || key.includes('poc')) orderData.phoneNumber = val;
      else if (key.includes('nbd') || key.includes('date')) {
        var dateVal = row[keyIdx + 1];
        if (dateVal instanceof Date) {
          orderData.nbdDate = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          var parsed = new Date(dateVal);
          if (!isNaN(parsed.valueOf())) {
            orderData.nbdDate = Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          } else {
            orderData.nbdDate = val;
          }
        }
      }
      else if (key === 'poc') orderData.name = val;
    });
  }

  if (itemHeaderRowIndex > -1) {
    const headerRow = data[itemHeaderRowIndex].map(s => String(s).toLowerCase());
    const cFbpn = headerRow.indexOf('fbpn');
    const cDesc = headerRow.indexOf('description');
    const cQty = headerRow.findIndex(h => h.includes('qty') || h.includes('quantity'));
    const cMfr = headerRow.indexOf('manufacturer');

    if (cFbpn > -1 && cQty > -1) {
      for (let i = itemHeaderRowIndex + 1; i < data.length; i++) {
        const fbpn = String(data[i][cFbpn]).trim();
        const qty = parseInt(data[i][cQty], 10);
        const desc = cDesc > -1 ? String(data[i][cDesc]) : '';
        const manufacturer = cMfr > -1 ? String(data[i][cMfr]).trim() : '';

        if (fbpn && !isNaN(qty) && qty > 0) {
          orderData.items.push({
            fbpn: fbpn,
            qty: qty,
            description: desc,
            manufacturer: manufacturer
          });
        }
      }
    }
  }

  if (orderData.taskNumber && orderData.company && orderData.items.length > 0) {
    Logger.log('Submitting order for ' + orderData.company);
    const result = processCustomerOrder(orderData);
    if (result.success) Logger.log('Order Created: ' + result.orderId);
    else Logger.log('Order Creation Failed: ' + result.message);
  } else {
    Logger.log('Invalid file format or missing data (Task, Company, or Items).');
  }

  try {
    DriveApp.getFileById(tempFile.id).setTrashed(true);
  } catch (e) { Logger.log('Error cleaning temp file: ' + e); }
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Customer_Orders') return;

  const STATUS_COL_INDEX = 5;
  const ORDER_ID_COL_INDEX = 1;

  if (e.range.getColumn() === STATUS_COL_INDEX && e.value === 'Cancelled') {
    const row = e.range.getRow();
    const orderId = sheet.getRange(row, ORDER_ID_COL_INDEX).getValue();

    if (orderId) {
      Logger.log(`Auto-cancelling Order ${orderId} due to status change.`);
      cancelOrder(orderId);
      e.source.toast(`Order ${orderId} processing cancellation...`);
    }
  }
}

// ============================================================================
// 1. GET CUSTOMER ORDERS (Role-Based Filtering)
// ============================================================================

function getCustomerOrders(context) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Customer_Orders');

    if (!sheet) {
      return { success: false, message: 'Customer_Orders sheet not found', orders: [] };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: true, orders: [] };
    }

    const headers = data[0];

    // Map header indices
    const colMap = {
      orderId: findColumnIndex(headers, ['order_id', 'order id', 'orderid']),
      taskNumber: findColumnIndex(headers, ['task_number', 'task number', 'task#', 'task']),
      project: findColumnIndex(headers, ['project']),
      nbd: findColumnIndex(headers, ['nbd', 'need_by_date', 'need by date']),
      company: findColumnIndex(headers, ['company', 'company_name']),
      orderTitle: findColumnIndex(headers, ['order_title', 'title', 'order title']),
      deliverTo: findColumnIndex(headers, ['deliver_to', 'delivery_address', 'deliver to']),
      status: findColumnIndex(headers, ['request_status', 'status']),
      stockStatus: findColumnIndex(headers, ['stock_status', 'stock status']),
      trackingNumber: findColumnIndex(headers, ['tracking_number', 'tracking', 'tracking number']),
      shipDate: findColumnIndex(headers, ['ship_date', 'shipped_date', 'ship date']),
      carrier: findColumnIndex(headers, ['carrier']),
      pickTicketUrl: findColumnIndex(headers, ['pick_ticket_url', 'pick_ticket', 'pick ticket']),
      packingListUrl: findColumnIndex(headers, ['packing_list_url', 'packing_list', 'packing list']),
      tocUrl: findColumnIndex(headers, ['toc_url', 'toc']),
      createdAt: findColumnIndex(headers, ['created_at', 'date_created', 'timestamp']),
      createdBy: findColumnIndex(headers, ['created_by', 'user_email'])
    };

    const orders = [];

    // Determine user's company for filtering (Standard users only)
    let userCompany = '';
    if (context && context.accessLevel === 'STANDARD' && context.company) {
      userCompany = context.company.toLowerCase().trim();
    }

    // Process rows (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows
      const orderId = colMap.orderId >= 0 ? row[colMap.orderId] : '';
      if (!orderId) continue;

      // Get company for this row
      const rowCompany = colMap.company >= 0 ? String(row[colMap.company] || '').toLowerCase().trim() : '';

      // Apply role-based filtering
      if (userCompany && rowCompany !== userCompany) {
        continue; // Skip orders not belonging to user's company
      }

      // Build order object
      const order = {
        orderId: String(orderId),
        taskNumber: colMap.taskNumber >= 0 ? String(row[colMap.taskNumber] || '') : '',
        project: colMap.project >= 0 ? String(row[colMap.project] || '') : '',
        nbd: colMap.nbd >= 0 ? formatDateValue(row[colMap.nbd]) : '',
        company: colMap.company >= 0 ? String(row[colMap.company] || '') : '',
        orderTitle: colMap.orderTitle >= 0 ? String(row[colMap.orderTitle] || '') : '',
        deliverTo: colMap.deliverTo >= 0 ? String(row[colMap.deliverTo] || '') : '',
        status: colMap.status >= 0 ? String(row[colMap.status] || 'Pending') : 'Pending',
        stockStatus: colMap.stockStatus >= 0 ? String(row[colMap.stockStatus] || '') : '',
        trackingNumber: colMap.trackingNumber >= 0 ? String(row[colMap.trackingNumber] || '') : '',
        shipDate: colMap.shipDate >= 0 ? formatDateValue(row[colMap.shipDate]) : '',
        carrier: colMap.carrier >= 0 ? String(row[colMap.carrier] || '') : '',
        pickTicketUrl: colMap.pickTicketUrl >= 0 ? extractUrl(row[colMap.pickTicketUrl]) : '',
        packingListUrl: colMap.packingListUrl >= 0 ? extractUrl(row[colMap.packingListUrl]) : '',
        tocUrl: colMap.tocUrl >= 0 ? extractUrl(row[colMap.tocUrl]) : '',
        createdAt: colMap.createdAt >= 0 ? formatDateValue(row[colMap.createdAt]) : '',
        createdBy: colMap.createdBy >= 0 ? String(row[colMap.createdBy] || '') : '',
        rowIndex: i + 1 // 1-indexed row number for updates
      };

      orders.push(order);
    }

    // Sort by created date (newest first)
    orders.sort((a, b) => {
      const dateA = a.createdAt ? new Date(a.createdAt) : new Date(0);
      const dateB = b.createdAt ? new Date(b.createdAt) : new Date(0);
      return dateB - dateA;
    });

    return {
      success: true,
      orders: orders,
      totalCount: orders.length
    };

  } catch (error) {
    Logger.log('Error in getCustomerOrders: ' + error.toString());
    return {
      success: false,
      message: error.toString(),
      orders: []
    };
  }
}

/**
 * Helper: Find column index from multiple possible header names
 */
function findColumnIndex(headers, possibleNames) {
  const normalizedHeaders = headers.map(h => String(h || '').toLowerCase().trim().replace(/[_\s]+/g, '_'));

  for (const name of possibleNames) {
    const normalizedName = name.toLowerCase().trim().replace(/[_\s]+/g, '_');
    const idx = normalizedHeaders.indexOf(normalizedName);
    if (idx >= 0) return idx;
  }
  return -1;
}

/**
 * Helper: Format date values
 */
function formatDateValue(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}

/**
 * Helper: Extract URL from HYPERLINK formula or plain text
 */
function extractUrl(value) {
  if (!value) return '';
  const str = String(value);

  // Check for HYPERLINK formula
  const match = str.match(/=HYPERLINK\s*\(\s*"([^"]+)"/i);
  if (match) return match[1];

  // Check if it's already a URL
  if (str.startsWith('http://') || str.startsWith('https://')) {
    return str;
  }

  return '';
}

// ============================================================================
// 2. UPLOAD TO AUTOMATION FOLDER
// ============================================================================

function uploadToAutomationFolder(fileData) {
  try {
    if (!fileData || !fileData.content || !fileData.fileName) {
      return { success: false, message: 'Missing file data or filename' };
    }

    const folder = DriveApp.getFolderById(AUTOMATION_FOLDER_ID);

    // Decode base64 content
    const decodedContent = Utilities.base64Decode(fileData.content);
    const blob = Utilities.newBlob(decodedContent, fileData.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileData.fileName);

    // Create file in automation folder
    const file = folder.createFile(blob);

    return {
      success: true,
      message: 'File uploaded successfully. It will be processed by the automation system.',
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      fileName: file.getName()
    };

  } catch (error) {
    Logger.log('Error in uploadToAutomationFolder: ' + error.toString());
    return {
      success: false,
      message: 'Error uploading file: ' + error.toString()
    };
  }
}

// ============================================================================
// 3. REGENERATE ORDER DOCUMENT
// ============================================================================

function regenerateOrderDoc(orderId, docType) {
  try {
    if (!orderId) {
      return { success: false, message: 'Order ID is required' };
    }

    if (!['PICK', 'PACKING', 'TOC'].includes(docType)) {
      return { success: false, message: 'Invalid document type. Must be PICK, PACKING, or TOC' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Fetch order data
    const orderData = getOrderDataForDoc(ss, orderId);
    if (!orderData) {
      return { success: false, message: 'Order not found: ' + orderId };
    }

    // Fetch order items
    const items = getOrderItemsForDoc(ss, orderId);
    if (!items || items.length === 0) {
      return { success: false, message: 'No items found for order: ' + orderId };
    }

    // Build form data for document generation
    const formData = {
      orderNumber: orderId,
      orderId: orderId,
      taskNumber: orderData.taskNumber,
      company: orderData.company,
      project: orderData.project,
      orderTitle: orderData.orderTitle,
      deliverTo: orderData.deliverTo,
      name: orderData.name,
      phoneNumber: orderData.phoneNumber,
      date: new Date().toLocaleDateString(),
      shipDate: orderData.shipDate || new Date().toLocaleDateString(),
      items: items
    };

    let result;

    switch (docType) {
      case 'PICK':
        // Generate Pick Ticket
        if (typeof generatePickTicket !== 'function') {
          return { success: false, message: 'generatePickTicket function not available' };
        }
        result = generatePickTicket(formData);
        break;

      case 'PACKING':
        // Generate Packing Lists - need to structure as skids
        formData.skids = buildSkidsFromItems(items, orderId);
        formData.totalSkids = formData.skids.length;

        if (typeof generatePackingLists !== 'function') {
          return { success: false, message: 'generatePackingLists function not available' };
        }
        result = generatePackingLists(formData);
        break;

      case 'TOC':
        // Generate TOC - need to structure as skids
        formData.skids = buildSkidsFromItems(items, orderId);
        formData.totalSkids = formData.skids.length;

        if (typeof generateTOC !== 'function') {
          return { success: false, message: 'generateTOC function not available' };
        }
        result = generateTOC(formData);
        break;
    }

    if (result && result.success) {
      // Update the order record with the new document URL
      updateOrderDocUrl(ss, orderId, docType, result.pdfUrl || result.url);

      return {
        success: true,
        url: result.pdfUrl || result.url,
        docType: docType,
        message: docType + ' document generated successfully'
      };
    } else {
      return {
        success: false,
        message: result ? result.message : 'Document generation failed'
      };
    }

  } catch (error) {
    Logger.log('Error in regenerateOrderDoc: ' + error.toString());
    return {
      success: false,
      message: 'Error generating document: ' + error.toString()
    };
  }
}

function getOrderDataForDoc(ss, orderId) {
  const sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find column indices
  const colOrderId = findColumnIndex(headers, ['order_id', 'order id']);
  const colTask = findColumnIndex(headers, ['task_number', 'task number']);
  const colProject = findColumnIndex(headers, ['project']);
  const colCompany = findColumnIndex(headers, ['company']);
  const colTitle = findColumnIndex(headers, ['order_title', 'title']);
  const colDeliverTo = findColumnIndex(headers, ['deliver_to', 'delivery_address']);
  const colName = findColumnIndex(headers, ['name', 'contact_name']);
  const colPhone = findColumnIndex(headers, ['phone_number', 'phone']);
  const colShipDate = findColumnIndex(headers, ['ship_date', 'shipped_date']);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowOrderId = colOrderId >= 0 ? String(row[colOrderId]) : '';

    if (rowOrderId === String(orderId)) {
      return {
        orderId: rowOrderId,
        taskNumber: colTask >= 0 ? String(row[colTask] || '') : '',
        project: colProject >= 0 ? String(row[colProject] || '') : '',
        company: colCompany >= 0 ? String(row[colCompany] || '') : '',
        orderTitle: colTitle >= 0 ? String(row[colTitle] || '') : '',
        deliverTo: colDeliverTo >= 0 ? String(row[colDeliverTo] || '') : '',
        name: colName >= 0 ? String(row[colName] || '') : '',
        phoneNumber: colPhone >= 0 ? String(row[colPhone] || '') : '',
        shipDate: colShipDate >= 0 ? formatDateValue(row[colShipDate]) : ''
      };
    }
  }

  return null;
}

function getOrderItemsForDoc(ss, orderId) {
  const sheet = ss.getSheetByName('Requested_Items');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find column indices
  const colOrderId = findColumnIndex(headers, ['order_id', 'order_number']);
  const colFbpn = findColumnIndex(headers, ['fbpn']);
  const colDesc = findColumnIndex(headers, ['description', 'desc']);
  const colQtyReq = findColumnIndex(headers, ['qty_requested', 'qty', 'quantity']);
  const colSku = findColumnIndex(headers, ['sku']);

  const items = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowOrderId = colOrderId >= 0 ? String(row[colOrderId]) : '';

    // Match order ID (handle numeric vs string comparison)
    const match = rowOrderId === String(orderId) ||
      rowOrderId === String(Math.trunc(Number(orderId)));

    if (match) {
      items.push({
        fbpn: colFbpn >= 0 ? String(row[colFbpn] || '') : '',
        description: colDesc >= 0 ? String(row[colDesc] || '') : '',
        qtyRequested: colQtyReq >= 0 ? Number(row[colQtyReq] || 0) : 0,
        qty: colQtyReq >= 0 ? Number(row[colQtyReq] || 0) : 0,
        sku: colSku >= 0 ? String(row[colSku] || '') : ''
      });
    }
  }

  return items;
}

function buildSkidsFromItems(items, orderId) {
  if (!items || items.length === 0) return [];

  return [{
    skidNumber: 1,
    items: items.map(item => ({
      fbpn: item.fbpn,
      description: item.description,
      qtyRequested: item.qtyRequested,
      qty: item.qtyRequested,
      qtyOnSkid: item.qtyRequested
    }))
  }];
}

function updateOrderDocUrl(ss, orderId, docType, url) {
  const sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet || !url) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colOrderId = findColumnIndex(headers, ['order_id', 'order id']);

  // Determine which column to update
  let targetColName;
  switch (docType) {
    case 'PICK':
      targetColName = ['pick_ticket_url', 'pick_ticket'];
      break;
    case 'PACKING':
      targetColName = ['packing_list_url', 'packing_list'];
      break;
    case 'TOC':
      targetColName = ['toc_url', 'toc'];
      break;
    default:
      return;
  }

  const targetCol = findColumnIndex(headers, targetColName);
  if (targetCol < 0) return;

  for (let i = 1; i < data.length; i++) {
    const rowOrderId = colOrderId >= 0 ? String(data[i][colOrderId]) : '';

    if (rowOrderId === String(orderId)) {
      // Set as hyperlink formula
      const linkText = docType === 'PICK' ? 'Pick Ticket' :
        docType === 'PACKING' ? 'Packing List' : 'TOC';
      sheet.getRange(i + 1, targetCol + 1).setFormula(`=HYPERLINK("${url}", "${linkText}")`);
      break;
    }
  }
}

// ============================================================================
// 4. GET FORM DROPDOWN DATA
// ============================================================================

function getCompaniesFiltered(context) {
  try {
    const companies = getCompanies(); // Use existing function

    // If Standard user, only return their company
    if (context && context.accessLevel === 'STANDARD' && context.company) {
      return [context.company];
    }

    return companies;
  } catch (error) {
    Logger.log('Error in getCompaniesFiltered: ' + error.toString());
    return [];
  }
}

function getProjectsFiltered(company) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const supportSheet = ss.getSheetByName('Support_Sheet');

    if (!supportSheet) return getProjects(); // Fallback to all projects

    const data = supportSheet.getDataRange().getValues();
    const projects = new Set();

    for (let i = 1; i < data.length; i++) {
      const rowCompany = String(data[i][0] || '').trim();
      const rowProject = String(data[i][1] || '').trim();

      if (rowProject) {
        // If company filter provided, only include matching projects
        if (company) {
          if (rowCompany.toLowerCase() === company.toLowerCase()) {
            projects.add(rowProject);
          }
        } else {
          projects.add(rowProject);
        }
      }
    }

    return Array.from(projects).sort();
  } catch (error) {
    Logger.log('Error in getProjectsFiltered: ' + error.toString());
    return [];
  }
}

function getNextTaskNumber() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Customer_Orders');

    if (!sheet) return '1001';

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const taskCol = findColumnIndex(headers, ['task_number', 'task number', 'task']);

    if (taskCol < 0) return '1001';

    let maxNum = 1000;

    for (let i = 1; i < data.length; i++) {
      const taskStr = String(data[i][taskCol] || '');
      const num = parseInt(taskStr.replace(/\D/g, ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }

    return String(maxNum + 1);
  } catch (error) {
    return '1001';
  }
}

function validateFBPN(fbpn) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pmSheet = ss.getSheetByName('Project_Master');

    if (!pmSheet) return { valid: false, message: 'Project_Master not found' };

    const data = pmSheet.getDataRange().getValues();
    const headers = data[0];
    const fbpnCol = findColumnIndex(headers, ['fbpn', 'part_number']);
    const descCol = findColumnIndex(headers, ['description', 'desc']);
    const skuCol = findColumnIndex(headers, ['sku']);

    if (fbpnCol < 0) return { valid: false, message: 'FBPN column not found' };

    for (let i = 1; i < data.length; i++) {
      const rowFbpn = String(data[i][fbpnCol] || '').trim();

      if (rowFbpn.toLowerCase() === fbpn.toLowerCase()) {
        return {
          valid: true,
          fbpn: rowFbpn,
          description: descCol >= 0 ? String(data[i][descCol] || '') : '',
          sku: skuCol >= 0 ? String(data[i][skuCol] || '') : ''
        };
      }
    }

    return { valid: false, message: 'FBPN not found in system' };
  } catch (error) {
    return { valid: false, message: error.toString() };
  }
}
