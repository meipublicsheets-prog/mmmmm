// ============================================================================
// STOCK_TOTALS.GS - Runtime Stock Totals Engine (SKU primary, FBPN + MFR fallback)
// ============================================================================

const STOCK_TOTALS_COLS = {
  SKU: 0,
  FBPN: 1,
  MFPN: 2,
  MANUFACTURER: 3,
  QTY_AVAILABLE: 4,
  QTY_IN_RACKING: 5,
  QTY_ON_FLOOR: 6,
  QTY_INBOUND_STAGING: 7,
  QTY_ALLOCATED: 8,
  QTY_BACKORDERED: 9,
  QTY_SHIPPED: 10,
  LAST_UPDATED: 11
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getStockTotalsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TABS.STOCK_TOTALS);
  if (!sheet) throw new Error('Stock_Totals sheet not found.');
  return sheet;
}

function buildStockKey_(fbpn, manufacturer) {
  return `${(fbpn || '').toString().trim().toUpperCase()}|${(manufacturer || '').toString().trim().toUpperCase()}`;
}

function getIdentifierString_(itemInfo) {
  const fb = (itemInfo.fbpn || '').toString().trim();
  const mfr = (itemInfo.manufacturer || '').toString().trim();
  const sku = (itemInfo.sku || '').toString().trim();
  if (sku) return `SKU:${sku} (FBPN:${fb}, MFR:${mfr})`;
  return `FBPN:${fb} + MFR:${mfr}`;
}

// ---------------------------------------------------------------------------
// Row lookup / create (SKU primary, FBPN + Manufacturer keyed fallback)
// ---------------------------------------------------------------------------

function findStockTotalsRowByKey_(sheet, key) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowFb = (data[i][STOCK_TOTALS_COLS.FBPN] || '').toString().trim().toUpperCase();
    const rowMfr = (data[i][STOCK_TOTALS_COLS.MANUFACTURER] || '').toString().trim().toUpperCase();
    if (buildStockKey_(rowFb, rowMfr) === key) return i + 1; // 1-based
  }
  return -1;
}

function findStockTotalsRow_(sheet, itemInfo) {
  const data = sheet.getDataRange().getValues();
  const sku = (itemInfo.sku || '').toString().trim().toUpperCase();
  if (sku) {
    for (let i = 1; i < data.length; i++) {
      const rowSku = (data[i][STOCK_TOTALS_COLS.SKU] || '').toString().trim().toUpperCase();
      if (rowSku === sku) return i + 1; // 1-based
    }
  }

  const fb = (itemInfo.fbpn || '').toString().trim().toUpperCase();
  const mfr = (itemInfo.manufacturer || '').toString().trim().toUpperCase();
  if (!fb || !mfr) return -1;

  const key = buildStockKey_(fb, mfr);
  return findStockTotalsRowByKey_(sheet, key);
}

function createStockTotalsRow_(sheet, itemInfo) {
  const fb = (itemInfo.fbpn || '').toString().trim().toUpperCase();
  const mfr = (itemInfo.manufacturer || '').toString().trim().toUpperCase();
  const sku = (itemInfo.sku || '').toString().trim();
  const mfpn = (itemInfo.mfpn || '').toString().trim();

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];

  const newRow = new Array(headers.length).fill('');

  if (STOCK_TOTALS_COLS.SKU < headers.length) newRow[STOCK_TOTALS_COLS.SKU] = sku;
  if (STOCK_TOTALS_COLS.FBPN < headers.length) newRow[STOCK_TOTALS_COLS.FBPN] = fb;
  if (STOCK_TOTALS_COLS.MFPN < headers.length) newRow[STOCK_TOTALS_COLS.MFPN] = mfpn;
  if (STOCK_TOTALS_COLS.MANUFACTURER < headers.length) newRow[STOCK_TOTALS_COLS.MANUFACTURER] = mfr;
  if (STOCK_TOTALS_COLS.QTY_AVAILABLE < headers.length) newRow[STOCK_TOTALS_COLS.QTY_AVAILABLE] = 0;
  if (STOCK_TOTALS_COLS.QTY_IN_RACKING < headers.length) newRow[STOCK_TOTALS_COLS.QTY_IN_RACKING] = 0;
  if (STOCK_TOTALS_COLS.QTY_ON_FLOOR < headers.length) newRow[STOCK_TOTALS_COLS.QTY_ON_FLOOR] = 0;
  if (STOCK_TOTALS_COLS.QTY_INBOUND_STAGING < headers.length) newRow[STOCK_TOTALS_COLS.QTY_INBOUND_STAGING] = 0;
  if (STOCK_TOTALS_COLS.QTY_ALLOCATED < headers.length) newRow[STOCK_TOTALS_COLS.QTY_ALLOCATED] = 0;
  if (STOCK_TOTALS_COLS.QTY_BACKORDERED < headers.length) newRow[STOCK_TOTALS_COLS.QTY_BACKORDERED] = 0;
  if (STOCK_TOTALS_COLS.QTY_SHIPPED < headers.length) newRow[STOCK_TOTALS_COLS.QTY_SHIPPED] = 0;
  if (STOCK_TOTALS_COLS.LAST_UPDATED < headers.length) newRow[STOCK_TOTALS_COLS.LAST_UPDATED] = new Date();

  sheet.appendRow(newRow);
  return sheet.getLastRow();
}

// ---------------------------------------------------------------------------
// Qty Available recalculation helpers (kept for reference; formulas now own totals)
// ---------------------------------------------------------------------------

function recalcQtyAvailable_(sheet, rowIndex) {
  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const inRacking = Number(row[STOCK_TOTALS_COLS.QTY_IN_RACKING] || 0);
  const onFloor = Number(row[STOCK_TOTALS_COLS.QTY_ON_FLOOR] || 0);
  const inboundStaging = Number(row[STOCK_TOTALS_COLS.QTY_INBOUND_STAGING] || 0);
  const allocated = Number(row[STOCK_TOTALS_COLS.QTY_ALLOCATED] || 0);
  const backordered = Number(row[STOCK_TOTALS_COLS.QTY_BACKORDERED] || 0);
  const shipped = Number(row[STOCK_TOTALS_COLS.QTY_SHIPPED] || 0);

  const totalPhysical = inRacking + onFloor + inboundStaging;
  const logicalAvailable = totalPhysical - allocated - backordered - shipped;

  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_AVAILABLE + 1).setValue(logicalAvailable);
  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.LAST_UPDATED + 1).setValue(new Date());
}

// ---------------------------------------------------------------------------
// Inbound / Staging / Outbound stock updates (no longer used; preserved for reference)
// ---------------------------------------------------------------------------

function updateStockTotals_Inbound(itemInfo, qtyReceived) {
  const sheet = getStockTotalsSheet_();
  let rowIndex = findStockTotalsRow_(sheet, itemInfo);
  if (rowIndex <= 0) {
    rowIndex = createStockTotalsRow_(sheet, itemInfo);
  }

  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const inRacking = Number(row[STOCK_TOTALS_COLS.QTY_IN_RACKING] || 0);
  const inboundStaging = Number(row[STOCK_TOTALS_COLS.QTY_INBOUND_STAGING] || 0);

  sheet
    .getRange(rowIndex, STOCK_TOTALS_COLS.QTY_IN_RACKING + 1)
    .setValue(inRacking + qtyReceived);
  sheet
    .getRange(rowIndex, STOCK_TOTALS_COLS.QTY_INBOUND_STAGING + 1)
    .setValue(inboundStaging + 0);

  recalcQtyAvailable_(sheet, rowIndex);

  Logger.log(`Stock_Totals INBOUND: ${getIdentifierString_(itemInfo)} +${qtyReceived}`);
  return { success: true };
}

function updateStockTotals_MoveFromStaging(itemInfo, qty, destinationType) {
  const sheet = getStockTotalsSheet_();
  let rowIndex = findStockTotalsRow_(sheet, itemInfo);
  if (rowIndex <= 0) {
    rowIndex = createStockTotalsRow_(sheet, itemInfo);
  }

  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  let inboundStaging = Number(row[STOCK_TOTALS_COLS.QTY_INBOUND_STAGING] || 0);
  let inRacking = Number(row[STOCK_TOTALS_COLS.QTY_IN_RACKING] || 0);
  let onFloor = Number(row[STOCK_TOTALS_COLS.QTY_ON_FLOOR] || 0);

  inboundStaging = Math.max(0, inboundStaging - qty);
  if (destinationType === 'RACK') {
    inRacking += qty;
  } else if (destinationType === 'FLOOR') {
    onFloor += qty;
  }

  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_INBOUND_STAGING + 1).setValue(inboundStaging);
  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_IN_RACKING + 1).setValue(inRacking);
  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_ON_FLOOR + 1).setValue(onFloor);

  recalcQtyAvailable_(sheet, rowIndex);

  Logger.log(`Stock_Totals MOVE_FROM_STAGING: ${getIdentifierString_(itemInfo)} qty=${qty}, dest=${destinationType}`);
  return { success: true };
}

function updateStockTotals_Outbound(itemInfo, qtyShipped, sourceType) {
  const sheet = getStockTotalsSheet_();
  let rowIndex = findStockTotalsRow_(sheet, itemInfo);
  if (rowIndex <= 0) {
    rowIndex = createStockTotalsRow_(sheet, itemInfo);
  }

  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  let inRacking = Number(row[STOCK_TOTALS_COLS.QTY_IN_RACKING] || 0);
  let onFloor = Number(row[STOCK_TOTALS_COLS.QTY_ON_FLOOR] || 0);
  let shipped = Number(row[STOCK_TOTALS_COLS.QTY_SHIPPED] || 0);

  if (sourceType === 'Bin_Stock') {
    inRacking = Math.max(0, inRacking - qtyShipped);
  } else if (sourceType === 'Floor_Stock_Levels') {
    onFloor = Math.max(0, onFloor - qtyShipped);
  }
  shipped += qtyShipped;

  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_IN_RACKING + 1).setValue(inRacking);
  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_ON_FLOOR + 1).setValue(onFloor);
  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_SHIPPED + 1).setValue(shipped);

  recalcQtyAvailable_(sheet, rowIndex);

  Logger.log(`Stock_Totals OUTBOUND: ${getIdentifierString_(itemInfo)} -${qtyShipped} from ${sourceType}`);
  return { success: true };
}

// ---------------------------------------------------------------------------
// Allocation / Backorder only – these are the ONLY ones you should use now
// ---------------------------------------------------------------------------

function updateStockTotals_OrderAllocation(itemInfo, qtyAllocated, qtyBackordered) {
  const sheet = getStockTotalsSheet_();
  const rowIndex = findStockTotalsRow_(sheet, itemInfo);
  if (rowIndex <= 0) return { success: false, message: 'Item not found' };

  if (qtyAllocated > 0) {
    const curAlloc = Number(sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_ALLOCATED + 1).getValue() || 0);
    sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_ALLOCATED + 1).setValue(curAlloc + qtyAllocated);
  }

  if (qtyBackordered > 0) {
    const curBo = Number(sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_BACKORDERED + 1).getValue() || 0);
    sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_BACKORDERED + 1).setValue(curBo + qtyBackordered);
  }

  Logger.log(`Stock_Totals ALLOC: ${getIdentifierString_(itemInfo)} alloc=${qtyAllocated}, bo=${qtyBackordered}`);
  return { success: true };
}

function updateStockTotals_BackorderFulfilled(itemInfo, qtyFulfilled) {
  const sheet = getStockTotalsSheet_();
  const rowIndex = findStockTotalsRow_(sheet, itemInfo);
  if (rowIndex <= 0) return { success: false, message: 'Item not found' };

  const curBo = Number(sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_BACKORDERED + 1).getValue() || 0);
  const curAlloc = Number(sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_ALLOCATED + 1).getValue() || 0);

  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_BACKORDERED + 1).setValue(Math.max(0, curBo - qtyFulfilled));
  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_ALLOCATED + 1).setValue(curAlloc + qtyFulfilled);

  Logger.log(`Stock_Totals BO_FULFILLED: ${getIdentifierString_(itemInfo)} ${qtyFulfilled}`);
  return { success: true };
}

function updateStockTotals_CancelAllocation(itemInfo, qtyToRelease) {
  const sheet = getStockTotalsSheet_();
  const rowIndex = findStockTotalsRow_(sheet, itemInfo);
  if (rowIndex <= 0) return { success: false, message: 'Item not found' };

  const curAlloc = Number(sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_ALLOCATED + 1).getValue() || 0);
  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_ALLOCATED + 1).setValue(Math.max(0, curAlloc - qtyToRelease));

  Logger.log(`Stock_Totals CANCEL_ALLOC: ${getIdentifierString_(itemInfo)} -${qtyToRelease}`);
  return { success: true };
}

function updateStockTotals_CancelBackorder(itemInfo, qtyToCancel) {
  const sheet = getStockTotalsSheet_();
  const rowIndex = findStockTotalsRow_(sheet, itemInfo);
  if (rowIndex <= 0) return { success: false, message: 'Item not found' };

  const curBo = Number(sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_BACKORDERED + 1).getValue() || 0);
  sheet.getRange(rowIndex, STOCK_TOTALS_COLS.QTY_BACKORDERED + 1).setValue(Math.max(0, curBo - qtyToCancel));

  Logger.log(`Stock_Totals CANCEL_BO: ${getIdentifierString_(itemInfo)} -${qtyToCancel}`);
  return { success: true };
}

// ---------------------------------------------------------------------------
// Legacy wrappers (SKU-only) – still safe; only allocations/backorders in use
// ---------------------------------------------------------------------------

function updateStockTotals_Inbound_Legacy(sku, fbpn, mfpn, manufacturer, qtyReceived) {
  return updateStockTotals_Inbound({ sku, fbpn, mfpn, manufacturer }, qtyReceived);
}

function updateStockTotals_MoveFromStaging_Legacy(sku, qty, destinationType) {
  return updateStockTotals_MoveFromStaging({ sku: sku, fbpn: '', manufacturer: '' }, qty, destinationType);
}

function updateStockTotals_OrderAllocation_Legacy(sku, qtyAllocated, qtyBackordered) {
  return updateStockTotals_OrderAllocation({ sku: sku, fbpn: '', manufacturer: '' }, qtyAllocated, qtyBackordered);
}

function updateStockTotals_Outbound_Legacy(sku, qtyShipped, sourceType) {
  return updateStockTotals_Outbound({ sku: sku, fbpn: '', manufacturer: '' }, qtyShipped, sourceType);
}
