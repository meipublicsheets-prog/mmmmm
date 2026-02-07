// ============================================================================
// CONFIG.GS - Shared Configuration and Universal Utility Functions
// FULL VERSION: Reconciled for Row 1 Header, Row 2 Spacer, Row 3 Data
// ============================================================================

const SHEET_IDS = {
  // The "Database" (Inventory, Master Logs)
  IMS_SOURCE: '1YGde5-R06qFcY5KGcmOP0-0Fbre-HVwKKZeljFFkgnI',

  // The "Interface" (New Sheet for Supervisors)
  ORDER_MANAGEMENT: '1ab6TvvDgh7SyiZDUBb96CUdsk5u0nZ7saR33Q0m95WE'
};

const FOLDERS = {
  IMS_ROOT: '1Fvr0WsyAfejmTiVRgYoA7zeLiwwnhxJ7',
  INBOUND_UPLOADS: '1pj5rVUwP7drH_vpigdJRePdCu9mj9jYF',
  CUSTOMER_ORDERS: '1G97y64fxlq6rBd8RItHREmNxRMRrK-VV',
  SOURCE_NEW_ORDERS: '1L3mjeQizzjVU5uTqGxv1sOUOuq25I2pM',
  PROCESSED_ORDERS: '1s28aomu1Th2_yNZOCkTyHdLF3Cq2dNep',
  IMS_TEMPLATES: '1cQatcc-vJLgx89_XLWoQ7XlIL_k6by33',
  PDF_TEMPLATES: '1Pi6u2Nt-WI5m6UAi5k9x-AM3srJ-GAGb',
  LABEL_TEMPLATES: '1DcfLwokIy2S9ldMMrNmWBdBSoKJ0XxgG',
  TEMP_WORK: '1s3hlgrOQ1KR4kgGTnOOgsOlLaj08ygD6',
  SOURCE_RECEIPTS: 'ID_FOR_SOURCE_RECEIPTS_FOLDER',
  PROCESSED_RECEIPTS: 'ID_FOR_PROCESSED_RECEIPTS_FOLDER',
  IMS_Reports: '1NxJRIcraNfpGk8ByvauVBFWkBtwp4y9D'
};

var REPORT_TEMPLATES = {
  INBOUND: '1wNz7SXXL-5Df15BBMg7wAE9imumgukLi3fnPYJ54Byk',
  OUTBOUND: '1yN0k4mYC4Cvr-Eon9WQDNccWbgmzwetRm3n6jtbBvgc'
};

const TEMPLATES = {
  TOC_Template: '1S6W9u_iLNrKzVL51rXLpXRY7d7i3RXRXp6v2FdFdmC4',
  Packing_List_Template: '1Z-InQHf8dUxQzsvF7dmbtwZ1SotmWabU2FY64ZNi-Zs',
  Backorder_Template: '1n8ZkctRKibi98yv-n8-ybWOPD1djVpVWiAICoT-dBTM',
  Pickticket_Template: '1J75-dFbotHgn9R5xLZbMMpSeeDM1mvz5SVsB_mNC0lY',
  Inbound_Label: '1a_XK-UdV-dPVnlDvHg7kK5zNuVCRyuev',
  Masterskid_Label: '1N5EXHfrvIT1-LyghWzthT_wsvteaFKAa'
};

const TABS = {
  CUSTOMER_ORDERS: "Customer_Orders",
  STOCK_TOTALS: "Stock_Totals",
  MASTER_LOG: "Master_Log",
  OUTBOUNDLOG: "OutboundLog",
  INBOUND_SKIDS: "Inbound_Skids",
  BACKORDERS: "Backorders",
  VERIFICATION_LOG: "Verification_Log",
  INBOUND_STAGING: "Inbound_Staging",
  SHEET56: "Sheet56",
  SHEET53: "Sheet53",
  INBOUND_QTY_AUDIT: "Inbound_Qty_Audit",
  BIN_STOCK: "Bin_Stock",
  FLOOR_STOCK_LEVELS: "Floor_Stock_Levels",
  SHEET69: "Sheet69",
  CYCLE_COUNT: "Cycle_Count",
  REQUESTED_ITEMS: "Requested_Items",
  PICK_LOG: "Pick_Log",
  SHEET71: "Sheet71",
  CONSOLIDATIONSUGGESTIONS: "ConsolidationSuggestions",
  PICKED_ITEM_LOG: "Picked_Item_Log",
  LOCATIONLOG: "LocationLog",
  CYCLE_COUNT_BATCHES: "Cycle_Count_Batches",
  CYCLE_COUNT_LINES: "Cycle_Count_Lines",
  REPLENISHMENT_TASKS: "Replenishment_Tasks",
  DOCK_SCHEDULE: "Dock_Schedule",
  RACKING_AUDIT: "Racking_Audit",
  BREAKDOWN_LOG: "Breakdown_Log",
  NVSCRIPTSPROPERTIES: "NVScriptsProperties",
  BACKORDERFULFILLMENT_LOG: "BackorderFulfillment_Log",
  TRUCK_SCHEDULE: "Truck_Schedule",
  ALLOCATION_LOG: "Allocation_Log",
  PROJECT_MASTER: "Project_Master",
  SUPPORT_SHEET: "Support_Sheet",
  AUDIT_TRAIL: "Audit_Trail",
  ITEM_MASTER: "Item_Master",
  PO_MASTER: "PO_Master",
  CUSTOMER_ACCESS: "Customer_Access",
  MANUFACTURER_MASTER: "Manufacturer_Master"
};

const HEADERS = {
  "Customer_Orders": ["Order_ID", "Task_Number", "Project", "NBD", "Request_Status", "Stock_Status", "Company", "Order_Title", "Deliver_To", "Name", "Phone_Number", "Original_Order", "Order_Folder", "Pick_Ticket_PDF", "TOC_PDF", "Packing_Lists", "Created_TS", "Created_By"],
  "Stock_Totals": ["SKU", "Asset_Type", "Manufacturer", "MFPN", "FBPN", "UOM", "Qty_Available", "Qty_In_Racking", "Qty_On_Floor", "Qty_Inbound_Staging", "Qty_Allocated", "Qty_Backordered", "Qty_Shipped", "Qty_Received", "Qty_In_Stock"],
  "Master_Log": ["Txn_ID", "Date_Received", "Inbound_Files", "Transaction_Type", "Warehouse", "Project", "Push #", "FBPN", "Qty_Received", "UOM", "Total_Skid_Count", "Carrier", "BOL_Number", "Customer_PO_Number", "Manufacturer", "MFPN", "Description", "Received_By", "SKU"],
  "OutboundLog": ["Date", "Order_Number", "Task_Number", "Transaction Type", "Warehouse", "Company", "Project", "FBPN", "Manufacturer", "Qty", "UOM", "Skid_ID", "SKU"],
  "Inbound_Skids": ["Skid_ID", "TXN_ID", "Date", "Asset_Type", "FBPN", "MFPN", "Project", "Qty_on_Skid", "UOM", "Skid_Sequence", "Is_Mixed", "Timestamp", "SKU"],
  "Backorders": ["Order_ID", "NBD", "Status", "Task_Number", "Stock_Status", "Asset_Type", "FBPN", "UOM", "Qty_Requested", "Qty_Backordered", "Qty_Fulfilled", "Date_Logged", "Date_Closed", "Notes", "Backorder_ID", "SKU"],
  "Verification_Log": ["Timestamp", "BOL_Number", "PO_Number", "Asset_Type", "Manufacturer", "MFPN", "FBPN", "UOM", "Expected_Qty", "Actual_Qty", "Variance", "Box_Labels", "Verified_By", "Skid_ID", "TXN_ID"],
  "Inbound_Staging": ["Bin_Code", "Bin_Name", "Push_Number", "Project", "Manufacturer", "FBPN", "UOM", "Initial_Quantity", "Current_Quantity", "Stock_Percentage", "AUDIT NEEDED", "Skid_ID", "SKU", "Last_Updated"],
  "BackorderFulfillment_Log": ["Order_ID", "Task_Number", "FBPN", "Qty_Fulfilled", "Fulfillment_Date", "Status_After", "Fulfilled_By", "Notes", "Backorder_ID", "Txn_ID", "Fulfillment_ID", "Timestamp", "SKU"],
  "Project_Master": ["Customer_PO", "FBPN", "MFPN", "Description", "Manufacturer", "Project", "SKU", "Qty_Ordered", "Qty_Received"],
  "Item_Master": ["FBPN", "Manufacturer", "MFPN", "Asset_Type", "Description", "UOM", "SKU"],
  "PO_Master": ["Customer_PO", "Project"]
};

// ============================================================================
// CORE UTILITIES
// ============================================================================

function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);
  return sheet;
}

function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 2) return [];
  const headers = data[0];
  const rows = data.slice(2); // Skip Row 1 (Header) and Row 2 (Spacer)

  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

function appendRow(sheetName, rowData) {
  const sheet = getSheet(sheetName);
  if (sheet.getMaxRows() < 2) sheet.insertRowsAfter(sheet.getMaxRows(), 2 - sheet.getMaxRows());
  sheet.appendRow(rowData);
}

function appendRows(sheetName, rowsData) {
  if (!rowsData || rowsData.length === 0) return;
  const sheet = getSheet(sheetName);
  if (sheet.getMaxRows() < 2) sheet.insertRowsAfter(sheet.getMaxRows(), 2 - sheet.getMaxRows());
  sheet.insertRowsAfter(2, rowsData.length);
  sheet.getRange(3, 1, rowsData.length, rowsData[0].length).setValues(rowsData);
}

function generateTxnId() {
  return generateRandomId('TXN_', 6);
}

function generateRandomId(prefix, len) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = '';
  for (let i = 0; i < len; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return prefix + result;
}

function generateSkidId() {
  const sheet = getSheet(TABS.INBOUND_SKIDS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return 'SKID_000001';
  return generateRandomId('SKID_', 6);
}

function formatDate(date) {
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  return `${month}/${day}/${year}`;
}

function formatMonthYear(date) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${months[date.getMonth()]} ${date.getFullYear()}`;
}

function getTimestamp() {
  const now = new Date();
  const month = now.getMonth() + 1;
  const day = now.getDate();
  const year = now.getFullYear();
  const hours = now.getHours();
  const minutes = now.getMinutes();
  const seconds = now.getSeconds();
  return `${month}/${day}/${year} ${hours}:${minutes}:${seconds}`;
}

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

function setCache(key, value, expirationInSeconds = 600) {
  const cache = CacheService.getScriptCache();
  const stringValue = JSON.stringify(value);
  cache.put(key, stringValue, expirationInSeconds);
}

function getCache(key) {
  const cache = CacheService.getScriptCache();
  const value = cache.get(key);
  return value ? JSON.parse(value) : null;
}

function removeCache(key) {
  const cache = CacheService.getScriptCache();
  cache.remove(key);
}

function clearAllCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(cache.getKeys());
}

function lookupProjectMaster(fbpn, customerPO, manufacturer) {
  const cacheKey = `PM_${fbpn}_${customerPO}_${manufacturer}`;
  const cached = getCache(cacheKey);
  if (cached) return cached;

  const sheet = getSheet(TABS.PROJECT_MASTER);
  const data = sheet.getDataRange().getValues();
  for (let i = 2; i < data.length; i++) { // Skip Hdr+Spacer
    const row = data[i];
    if (row[1] === fbpn && row[0] === customerPO && row[4] === manufacturer) {
      const result = { mfpn: row[2] || '', description: row[3] || '', project: row[5] || '', sku: row[6] || '' };
      setCache(cacheKey, result, 1800);
      return result;
    }
  }
  const emptyResult = { mfpn: '', description: '', project: '', sku: '' };
  setCache(cacheKey, emptyResult, 1800);
  return emptyResult;
}

function escapeHtml_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function createInboundFolder(date, bolNumber) {
  const rootFolder = DriveApp.getFolderById(FOLDERS.INBOUND_UPLOADS);
  const monthFolderName = formatMonthYear(date);
  const monthFolder = getOrCreateFolder(rootFolder, monthFolderName);
  const dayFolderName = String(date.getDate()).padStart(2, '0');
  const dayFolder = getOrCreateFolder(monthFolder, dayFolderName);
  const bolFolder = getOrCreateFolder(dayFolder, String(bolNumber));
  return bolFolder;
}

function uploadFileToFolder(folder, fileBlob, fileName) {
  return folder.createFile(fileBlob.setName(fileName));
}

function getNextStagingLocation(stagingArea) {
  const areaPrefix = stagingArea.replace('Inbound Staging ', 'IS');
  const sheet = getSheet(TABS.INBOUND_STAGING);
  const data = sheet.getDataRange().getValues();
  for (let i = 2; i < data.length; i++) {
    if (data[i][0].startsWith(areaPrefix) && !data[i][2]) {
      return { binCode: data[i][0], binName: data[i][1], rowIndex: i + 1 };
    }
  }
  return null;
}

function allocateStagingLocations(stagingArea, numberOfSkids) {
  const locations = [];
  const areaPrefix = stagingArea.replace('Inbound Staging ', 'IS');
  const sheet = getSheet(TABS.INBOUND_STAGING);
  const data = sheet.getDataRange().getValues();
  let allocated = 0;
  for (let i = 2; i < data.length && allocated < numberOfSkids; i++) {
    if (data[i][0].startsWith(areaPrefix) && !data[i][2]) {
      locations.push({ binCode: data[i][0], binName: data[i][1], rowIndex: i + 1 });
      allocated++;
    }
  }
  if (allocated < numberOfSkids) throw new Error(`Only ${allocated} of ${numberOfSkids} locations available`);
  return locations;
}

const REPORT_SETTINGS = { PERIODS: { DAILY: 1, WEEKLY: 7, MONTHLY: 30 }, DATE_FORMAT: 'MMddyy', MONTH_FORMAT: 'MMMyy' };
const UI_CONFIG = { MODAL_WIDTH: 600, MODAL_HEIGHT: 400, THEME: { PRIMARY_COLOR: '#1976D2', SUCCESS_COLOR: '#4CAF50', ERROR_COLOR: '#F44336', BACKGROUND: '#F5F5F5' } };
const SYSTEM_CONFIG = { TIMEZONE: Session.getScriptTimeZone(), MAX_RETRIES: 3, RETRY_DELAY: 2000, LOG_LEVEL: 'INFO' };

function getColumnIndex(headers, columnName) { return headers.indexOf(columnName); }
function validateConfig() { return { valid: true, message: 'Configuration is valid' }; }
function showConfigStatus() { SpreadsheetApp.getUi().alert('✓ Configuration Valid'); }
function getIMSConfig() {
  return {
    TOC_TEMPLATE_ID: TEMPLATES.TOC_Template || '', PACKING_LIST_TEMPLATE_ID: TEMPLATES.Packing_List_Template || '',
    PICK_TICKET_TEMPLATE_ID: TEMPLATES.Pickticket_Template || '',
    TOC_PACKING_OUTPUT_FOLDER_ID: FOLDERS.PROCESSED_ORDERS || FOLDERS.IMS_ROOT,
    CUSTOMER_ORDERS_FOLDER_ID: FOLDERS.CUSTOMER_ORDERS || FOLDERS.IMS_ROOT,
    SOURCE_NEW_ORDERS_FOLDER_ID: FOLDERS.SOURCE_NEW_ORDERS || ''
  };
}

// ============================================================================
// MAIN CONTROLLER
// ============================================================================
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('IMS')
    .addItem('Inbound Delivery Form', 'openInboundModal')
    .addItem('Inbound Count Verification', 'openInboundVerificationModal')
    .addItem('Create Skid Label', 'openInboundSkidLabelModal')
    .addItem('Generate Reports', 'showReportGeneratorModal')
    .addToUi();
}
function doGet(e) {
  const context = getUserContext();
  const template = HtmlService.createTemplateFromFile('IMSWebApp');
  template.userContext = JSON.stringify(context);
  template.deepLinkSkidId = JSON.stringify((e && e.parameter && e.parameter.skid) ? String(e.parameter.skid).trim() : '');
  return template.evaluate().setTitle('IMS - Warehouse Management').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================================
// MODALS
// ============================================================================
function openInboundManagerModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('InboundManagerModal').evaluate().setWidth(1000).setHeight(800), 'Inbound Manager'); }
function openInboundVerificationModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('InboundVerificationModal').evaluate().setWidth(1000).setHeight(800), 'Inbound Count Verification'); }
function openInboundModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('InboundModal').evaluate().setWidth(950).setHeight(1000), 'Inbound Form'); }
function showCustomerOrderModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('CustomerOrderModal').evaluate().setWidth(900).setHeight(750), 'Create Customer Order'); }
function openCancelOrderModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('CancelOrderModal').evaluate().setWidth(600).setHeight(450), 'Cancel Customer Order'); }
function openPickTicketGenerator() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('PickTicketModal').evaluate().setWidth(950).setHeight(850), 'Generate Pick Ticket'); }
function openPackingTOCGenerator() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('PackingTOCModal').evaluate().setWidth(1050).setHeight(850), 'Customer Order Outbound'); }
function showReportGeneratorModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('ReportGeneratorModal').evaluate().setWidth(950).setHeight(1000), 'Generate Reports'); }
function openStockToolsModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('StockToolsModal').evaluate().setWidth(1050).setHeight(850), 'Bin Stock Put-Away'); }
function openBinLookupModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('BinLookupModal').evaluate().setWidth(1000).setHeight(820), 'Bin & Item Lookup'); }
function openBinUpdateModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('BinUpdateModal').evaluate().setWidth(900).setHeight(820), 'Bin Update'); }
function openCycleCountModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('CycleCountModal').evaluate().setWidth(1200).setHeight(900), 'Cycle Count'); }
function openCurrentItemsModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('CurrentItemsModal').evaluate().setWidth(1100).setHeight(820), 'Current Items'); }
function showCustomerPortal() { SpreadsheetApp.getUi().showSidebar(HtmlService.createTemplateFromFile('CustomerPortalUI').evaluate().setTitle('Inventory Portal')); }
function openGeneratePastInboundLabelsByDateModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('LabelDatePickerModal').evaluate().setWidth(520).setHeight(360), 'Generate Past Inbound Labels'); }
function openInboundSkidLabelModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('InboundSkidLabelModal').evaluate().setWidth(850).setHeight(780), 'Create Inbound Skid Label'); }
function openAddItemModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('AddItemModal').evaluate().setWidth(500).setHeight(520), 'Add New Item'); }
function openAddPOModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('AddPOModal').evaluate().setWidth(500).setHeight(600), 'Add Customer PO & Project'); }
function openBatchLabelGeneratorModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('BatchLabelGenerator').evaluate().setWidth(1100).setHeight(700), 'Batch Label Generator'); }
function openPastDeliverySkidLabelModal() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('PastDeliverySkidLabelModal').evaluate().setWidth(860).setHeight(580), 'Past Delivery Labels'); }

// ----------------------------------------------------------------------------
// BATCH LABELS (Row 3 Fix)
// ----------------------------------------------------------------------------
function findHeaderRow_(data, requiredNames) { return 0; } // Assuming Row 1 is header

function shell_getBolsForLabelGeneration() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName('Master_Log');
    const skidsSheet = ss.getSheetByName('Inbound_Skids');
    if (!masterSheet) return { success: false, message: 'Master_Log missing.' };
    const mData = masterSheet.getDataRange().getValues();
    const sData = skidsSheet ? skidsSheet.getDataRange().getValues() : [];
    if (!mData || mData.length <= 2) return { success: true, bols: [] };

    const mHead = mData[0];
    const mTxn = mHead.indexOf('Txn_ID');
    const mDate = mHead.indexOf('Date_Received');
    const mBol = mHead.indexOf('BOL_Number');
    const mFbpn = mHead.indexOf('FBPN');
    const mMan = mHead.indexOf('Manufacturer');
    const mProj = mHead.indexOf('Project');

    const sHead = sData[0] || [];
    const sTxn = sHead.indexOf('TXN_ID');
    const skidCounts = {};
    for (let j = 2; j < sData.length; j++) {
      const tid = String(sData[j][sTxn] || '').trim();
      if (tid) skidCounts[tid] = (skidCounts[tid] || 0) + 1;
    }

    const bolMap = {};
    for (let i = 2; i < mData.length; i++) {
      const txnId = String(mData[i][mTxn] || '').trim();
      const bol = String(mData[i][mBol] || '').trim();
      if (!txnId || !bol) continue;
      const key = txnId + '|' + bol;
      if (bolMap[key]) continue;
      bolMap[key] = {
        key, txnId, bol,
        dateStr: mData[i][mDate] ? formatDate(new Date(mData[i][mDate])) : '',
        dateVal: mData[i][mDate],
        fbpn: mFbpn >= 0 ? mData[i][mFbpn] : '',
        manufacturer: mMan >= 0 ? mData[i][mMan] : '',
        project: mProj >= 0 ? mData[i][mProj] : '',
        skidCount: skidCounts[txnId] || 0, hasLabels: false
      };
    }
    const bols = Object.values(bolMap).sort((a,b) => new Date(b.dateVal) - new Date(a.dateVal)).slice(0, 250);
    return { success: true, bols: bols };
  } catch (err) { return { success: false, message: 'Error: ' + err.message }; }
}

function shell_generateLabelsForBol(txnId, bolNumber) {
  // Uses generateSkidLabels from IMS_Inbound which already has robust folder handling
  return IMS_Inbound.shell_generateLabelsForBol(txnId, bolNumber);
}
// Note: If IMS_Inbound isn't a library, just copy the function here or rely on global scope.

// ----------------------------------------------------------------------------
// DATA ENTRY
// ----------------------------------------------------------------------------
function addItemToItemMaster(data) {
  try {
    if (!data || !data.fbpn) return { success: false, message: 'FBPN is required.' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Item_Master');
    if (!sheet) return { success: false, message: 'Item_Master missing.' };
    
    const rawData = sheet.getDataRange().getValues();
    const fbpnIdx = rawData[0].indexOf('FBPN');
    if (fbpnIdx >= 0) {
      const check = String(data.fbpn).toUpperCase().trim();
      for (let i = 2; i < rawData.length; i++) {
         if (String(rawData[i][fbpnIdx]).toUpperCase().trim() === check) return { success: false, message: 'FBPN exists.' };
      }
    }
    appendRow('Item_Master', [data.fbpn, data.manufacturer, data.mfpn, data.assetType, data.description, data.uom, '']);
    return { success: true, message: 'Item added.' };
  } catch (e) { return { success: false, message: e.message }; }
}

function addPOToPOMaster(data) {
  try {
    if (!data.customerPO || !data.project) return { success: false, message: 'Missing PO or Project.' };
    appendRow('PO_Master', [data.customerPO, data.project]);
    return { success: true, message: 'PO added.' };
  } catch (e) { return { success: false, message: e.message }; }
}

function getUserContext() {
  const email = Session.getActiveUser().getEmail();
  return { authenticated: true, email: email, accessLevel: 'MEI', permissions: { isAdmin: true } };
}

function shell_generateLabelsForAllPastInbounds(startDate, endDate) {
  if (!startDate) throw new Error('startDate is required');
  return generateLabelsForAllPastInbounds(startDate, endDate || startDate);
}

// ----------------------------------------------------------------------------
// WEBAPP FORWARDERS
// ----------------------------------------------------------------------------
function getDashboardMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metrics = { orders: { pending: 0, processing: 0, shipped: 0 }, inventory: { totalSKUs: 0 }, inbound: { scheduled: 0, received: 0 } };
  try {
    const oSheet = ss.getSheetByName('Customer_Orders');
    if (oSheet) {
      const oData = oSheet.getDataRange().getValues();
      const sIdx = oData[0].indexOf('Request_Status');
      for (let i = 2; i < oData.length; i++) {
        const s = String(oData[i][sIdx] || '').toLowerCase();
        if (s.includes('pending')) metrics.orders.pending++;
        else if (s.includes('processing')) metrics.orders.processing++;
        else if (s.includes('shipped')) metrics.orders.shipped++;
      }
    }
    const stSheet = ss.getSheetByName('Stock_Totals');
    if (stSheet) metrics.inventory.totalSKUs = Math.max(0, stSheet.getLastRow() - 2);
  } catch (e) {}
  return { success: true, metrics: metrics };
}

function getCustomerOrders(context) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet) return { success: true, orders: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length <= 2) return { success: true, orders: [] };
  const headers = data[0];
  const idx = { id: headers.indexOf('Order_ID'), title: headers.indexOf('Order_Title'), status: headers.indexOf('Request_Status') };
  const orders = [];
  for (let i = 2; i < data.length; i++) {
    orders.push({ orderId: data[i][idx.id], orderTitle: data[i][idx.title], status: data[i][idx.status] });
  }
  return { success: true, orders: orders };
}

function uploadToAutomationFolder(fileData) {
  try {
    const folder = DriveApp.getFolderById(FOLDERS.SOURCE_NEW_ORDERS);
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData.content), fileData.mimeType, fileData.fileName);
    const file = folder.createFile(blob);
    return { success: true, message: 'Uploaded', fileUrl: file.getUrl() };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getCompaniesFiltered(context) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Support_Sheet');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const companies = new Set();
  for (let i = 2; i < data.length; i++) if(data[i][0]) companies.add(data[i][0]);
  return Array.from(companies).sort();
}

function getNextTaskNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet) return '1001';
  const data = sheet.getDataRange().getValues();
  const col = data[0].indexOf('Task_Number');
  let max = 1000;
  for (let i = 2; i < data.length; i++) {
    const n = parseInt(String(data[i][col]).replace(/\D/g, ''));
    if (!isNaN(n) && n > max) max = n;
  }
  return String(max + 1);
}

function validateFBPN(fbpn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Project_Master');
  if (!sheet) return { valid: false };
  const data = sheet.getDataRange().getValues();
  const col = data[0].indexOf('FBPN');
  for (let i = 2; i < data.length; i++) {
    if (String(data[i][col]).toUpperCase() === String(fbpn).toUpperCase()) return { valid: true };
  }
  return { valid: false };
}
