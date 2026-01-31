// ============================================================================
// CONFIG.GS - Shared Configuration and Universal Utility Functions
// ============================================================================
const SHEET_IDS = {
  // The "Database" (Inventory, Master Logs)
  IMS_SOURCE: '10DqIhFdwuZQOniGGwiZoqBL0batKOYnPghbHPzDVq0Q', 
  
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
  Pickticket_Template: '1J75-dFbotHgn9R5xLZbMMpSeeDM1mvz5SVsB_mNC0lY', // UPDATED ID
  Inbound_Label: '1a_XK-UdV-dPVnlDvHg7kK5zNuVCRyuev',
  Masterskid_Label: '1N5EXHfrvIT1-LyghWzthT_wsvteaFKAa'
};
const TABS = {
  MASTER_LOG: "Master_Log",
  INBOUND_SKIDS: "Inbound_Skids",
  INBOUND_STAGING: "Inbound_Staging",
  OUTBOUNDLOG: "OutboundLog",
  BIN_STOCK: "Bin_Stock",
  FLOOR_STOCK_LEVELS: "Floor_Stock_Levels",
  LOCATIONLOG: "LocationLog",
  CYCLE_COUNT: "Cycle_Count",
  STOCK_TOTALS: "Stock_Totals",
  TRUCK_SCHEDULE: "Truck_Schedule",
  BACKORDERS: "Backorders",
  CUSTOMER_ORDERS: "Customer_Orders",
  BREAKDOWN_LOG: "Breakdown_Log",
  NVSCRIPTSPROPERTIES: "NVScriptsProperties",
  REQUESTED_ITEMS: "Requested_Items",
  BACKORDERFULFILLMENT_LOG: "BackorderFulfillment_Log",
  ALLOCATION_LOG: "Allocation_Log",
  SUPPORT_SHEET: "Support_Sheet",
  ITEM_MASTER: "Item_Master",
  PROJECT_MASTER: "Project_Master",
  PO_MASTER: "PO_Master",
  CUSTOMER_ACCESS: "Customer_Access",
  PICK_LOG: "Pick_Log"
};
const HEADERS = {
  "Master_Log": ["Txn_ID", "Date_Received", "Transaction_Type", "Warehouse", "Project", "Push #", "FBPN", "Qty_Received", "Total_Skid_Count", "Carrier", "BOL_Number", "Customer_PO_Number", "Manufacturer", "MFPN", "Description", "Received_By", "SKU"],
  "Inbound_Skids": ["Skid_ID", "TXN_ID", "Date", "FBPN", "MFPN", "Project", "Qty_on_Skid", "Skid_Sequence", "Is_Mixed", "Timestamp", "SKU"],
  "Inbound_Staging": ["Bin_Code", "Bin_Name", "Push_Number", "FBPN", "Manufacturer", "Project", "Initial_Quantity", "Current_Quantity", "Stock_Percentage", "AUDIT NEEDED", "Skid_ID", "SKU"],
  "OutboundLog": ["Date", "Order_Number", "Task_Number", "Transaction Type", "Warehouse", "Company", "Project", "FBPN", "Manufacturer", "Qty", "Bin_Code", "Skid_ID", "SKU"],
  "Bin_Stock": ["Bin_Code", "Bin_Name", "Push_Number", "FBPN", "Manufacturer", "Project", "Initial_Quantity", "Current_Quantity", "Stock_Percentage", "AUDIT NEEDED", "Skid_ID", "SKU"],
  "Floor_Stock_Levels": ["Bin_Code", "Bin_Name", "Push_Number", "FBPN", "Manufacturer", "Project", "Initial_Quantity", "Current_Quantity", "Stock_Percentage", "AUDIT NEEDED", "Skid_ID", "SKU"],
  "LocationLog": ["Timestamp", 
"Action", "FBPN", "Manufacturer", "Bin_Code", "Qty_Changed", "Resulting_Qty", "Description", "User_Email", "SKU"],
  "Cycle_Count": ["Date", "Bin_Code", "Bin_Name", "Push_Number", "Project", "Manufacturer", "FBPN", "Counted_Qty", "System_Qty", "Variance", "Corrected_FBPN", "Corrected_Manufacturer", "Corrected_Project", "Timestamp", "User", "SKU"],
  "Stock_Totals": ["SKU", "FBPN", "MFPN", "Manufacturer", "Qty_Available", "Qty_In_Racking", "Qty_On_Floor", "Qty_Inbound_Staging", "Qty_Allocated", "Qty_Backordered", "Qty_Shipped", "Qty_Received", "Qty_In_Stock"],
  "Truck_Schedule": ["Schedule_ID", "Delivery_Date", "Time", "Carrier", "BOL_Number", "PO", "Project", "Total_Skids", "FBPN", "Total_Qty", "Notes", "BOL_Packing_List"],
  "Backorders": ["Order_ID", "NBD", "Status", "Task_Number", "Stock_Status", "FBPN", "Qty_Requested", "Qty_Backordered", "Qty_Fulfilled", "Date_Logged", "Date_Closed", "Notes", "Backorder_ID", "SKU"],
  "Customer_Orders": ["Order_ID", "Task_Number", "Project", "NBD", "Request_Status", "Stock_Status", "Company", "Order_Title", "Deliver_To", "Name", "Phone_Number", "Original_Order", "Order_Folder", "Pick_Ticket_PDF", "TOC_PDF", "Packing_Lists", "Created_TS", "Created_By"],
  "Breakdown_Log": ["Breakdown_ID", "Skid_ID", "SKU", "FBPN", "Qty_Before", "Qty_Removed", 
"Qty_Remaining", "MasterSkid_ID", "Qty_On_ChildSkid", "Bin_ID_Child", "Date_Broken_Down", "Broken_Down_By", "Notes"],
  "NVScriptsProperties": ["autocratn", "autocratp"],
  "Requested_Items": ["Order_ID", "FBPN", "Description", "Qty_Requested", "Stock_Status", "Qty_Backordered", "Qty_Allocated", "Backorder_ID", "Allocation_ID", "SKU"],
  "BackorderFulfillment_Log": ["Order_ID", "Task_Number", "FBPN", "Qty_Fulfilled", "Fulfillment_Date", "Status_After", "Fulfilled_By", "Notes", "Backorder_ID", "Txn_ID", "Fulfillment_ID", "Timestamp", "SKU"],
  "Allocation_Log": ["Order_ID", "Timestamp", "Allocation_Status", "FBPN", "Qty_Requested", "Qty_Allocated", "Qty_Backordered", "Allocated_By", "Backorder_ID", "Allocation_ID", "SKU"],
  "Support_Sheet": ["Company", "Project", "Name", "Phone_Number"],
  "Item_Master": ["SKU", "FBPN", "Description", "Asset Type", "Manufacturer", "Model", "QR Code", "Bin_Code", "Stock_On_Hand"],
  "Project_Master": ["Customer_PO", "FBPN", "MFPN", "Description", "Manufacturer", "Project", "SKU"],
  "PO_Master": ["Customer_PO", "Project"],
  "Customer_Access": ["Email", "Name", "Company Name", "Email Domain", "Access_Level", "Project_Access", "Active"],
  "Pick_Log": ["PIK_ID", "NBD", "Order_Number", "Task_Number", 
"Company", "Project", "FBPN", "Description", "Qty_Requested", "Qty_To_Pick", "Bin_Code", "Qty_Picked", "Status", "Picked_By", "Shipped_Date", "Timestamp", "SKU"]
};




function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  return sheet;
}

function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
const headers = data[0];
  const rows = data.slice(1);
  
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
  sheet.appendRow(rowData);
}

function appendRows(sheetName, rowsData) {
  if (!rowsData || rowsData.length === 0) return;
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
sheet.getRange(lastRow + 1, 1, rowsData.length, rowsData[0].length).setValues(rowsData);
}

function generateTxnId() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let id = 'TXN_';
for (let i = 0; i < 6; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
}
  return id;
}

function generateSkidId() {
  const sheet = getSheet(TABS.INBOUND_SKIDS);
  const lastRow = sheet.getLastRow();
if (lastRow <= 1) {
    return 'SKID_000001';
  }
  
  const lastSkidId = sheet.getRange(lastRow, 1).getValue();
const numPart = parseInt(lastSkidId.split('_')[1]) || 0;
  const newNum = numPart + 1;
  
  return 'SKID_' + String(newNum).padStart(6, '0');
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
// Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
const rowPO = row[0];
    const rowFBPN = row[1];
    const rowManufacturer = row[4];
if (rowFBPN === fbpn && rowPO === customerPO && rowManufacturer === manufacturer) {
      const result = {
        mfpn: row[2] ||
'',
        description: row[3] ||
'',
        project: row[5] ||
'',
        sku: row[6] || ''
      };
setCache(cacheKey, result, 1800); // Cache for 30 minutes
      return result;
}
  }
  
  // Not found
  const emptyResult = { mfpn: '', description: '', project: '', sku: '' };
setCache(cacheKey, emptyResult, 1800);
  return emptyResult;
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}


function createInboundFolder(date, bolNumber) {
  const rootFolder = DriveApp.getFolderById(FOLDERS.INBOUND_UPLOADS);
  
  // Create or get month folder (e.g., "Jan 2024")
  const monthFolderName = formatMonthYear(date);
const monthFolder = getOrCreateFolder(rootFolder, monthFolderName);
  
  // Create or get day folder (e.g., "15")
  const dayFolderName = String(date.getDate()).padStart(2, '0');
const dayFolder = getOrCreateFolder(monthFolder, dayFolderName);
  
  // Create or get BOL folder
  const bolFolder = getOrCreateFolder(dayFolder, bolNumber);
  
  return bolFolder;
}


function uploadFileToFolder(folder, fileBlob, fileName) {
  return folder.createFile(fileBlob.setName(fileName));
}

function getNextStagingLocation(stagingArea) {
  const areaPrefix = stagingArea.replace('Inbound Staging ', 'IS');
const sheet = getSheet(TABS.INBOUND_STAGING);
  const data = sheet.getDataRange().getValues();
  
  // Find all bins in the specified area that are empty
  for (let i = 1; i < data.length; i++) {
    const binCode = data[i][0];
const binName = data[i][1];
    const fbpn = data[i][2];
    
    // Check if this bin is in the correct area and is empty
    if (binCode.startsWith(areaPrefix) && !fbpn) {
      return { binCode: binCode, binName: binName, rowIndex: i + 1 };
}
  }
  
  return null; // No available locations
}


function allocateStagingLocations(stagingArea, numberOfSkids) {
  const locations = [];
const areaPrefix = stagingArea.replace('Inbound Staging ', 'IS');
  const sheet = getSheet(TABS.INBOUND_STAGING);
  const data = sheet.getDataRange().getValues();
  
  let allocated = 0;
for (let i = 1; i < data.length && allocated < numberOfSkids; i++) {
    const binCode = data[i][0];
const binName = data[i][1];
    const fbpn = data[i][2];
    
    if (binCode.startsWith(areaPrefix) && !fbpn) {
      locations.push({ binCode: binCode, binName: binName, rowIndex: i + 1 });
allocated++;
    }
  }
  
  if (allocated < numberOfSkids) {
    throw new Error(`Only ${allocated} of ${numberOfSkids} locations available in ${stagingArea}`);
}
  
  return locations;
}
const REPORT_SETTINGS = {
  // Date ranges for each period type
  PERIODS: {
    DAILY: 1,    // Last 1 day
    WEEKLY: 7,   // Last 7 days
    MONTHLY: 30  // Last 30 days
  },
  
  // Filename format
  DATE_FORMAT: 'MMddyy',           // e.g., 111325
  MONTH_FORMAT: 'MMMyy',           // e.g., Nov25
  
  // File naming pattern: {Period}_{Type}_{Date}.pdf
  // Examples: Daily_Inbound_111325.pdf, Weekly_Outbound_111325.pdf
};

// ═══════════════════════════════════════════════════════════════════════════
// UI SETTINGS
// ═══════════════════════════════════════════════════════════════════════════

const UI_CONFIG = {
  MODAL_WIDTH: 600,
  MODAL_HEIGHT: 400,
  THEME: {
    PRIMARY_COLOR: '#1976D2',
    SUCCESS_COLOR: '#4CAF50',
    ERROR_COLOR: '#F44336',
    BACKGROUND: '#F5F5F5'
  }
};
// ═══════════════════════════════════════════════════════════════════════════
// SYSTEM SETTINGS
// ═══════════════════════════════════════════════════════════════════════════

const SYSTEM_CONFIG = {
  TIMEZONE: Session.getScriptTimeZone(),
  MAX_RETRIES: 3,           // Number of retry attempts for Drive operations
  RETRY_DELAY: 2000,        // Milliseconds between retries
  LOG_LEVEL: 'INFO'         // INFO, WARNING, ERROR
};
// ═══════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Get column index by name
 */
function getColumnIndex(headers, columnName) {
  const index = headers.indexOf(columnName);
if (index === -1) {
    Logger.log(`Warning: Column "${columnName}" not found`);
  }
  return index;
}
var SHEET_NAMES = {
  MASTER_LOG: 'Master_Log',           // Inbound transactions
  OUTBOUND_LOG: 'OutboundLog',        // Outbound transactions
  BACKORDERS: 'Backorders',           // Backorder tracking
  CUSTOMER_ORDERS: 'Customer_Orders', // Order data
  BIN_STOCK: 'Bin_Stock',            // Bin-level inventory
  STOCK_TOTALS: 'Stock_Totals'       // Aggregated stock levels
};
// ═══════════════════════════════════════════════════════════════════════════
// COLUMN MAPPINGS
// ═══════════════════════════════════════════════════════════════════════════

// Master_Log columns (Inbound)
var MASTER_LOG_COLUMNS = {
  TXN_ID: 'Txn_ID',
  DATE_RECEIVED: 'Date_Received',
  TRANSACTION_TYPE: 'Transaction_Type',
  WAREHOUSE: 'Warehouse',
  PUSH_NUM: 'Push #',
  FBPN: 'FBPN',
  QTY_RECEIVED: 'Qty_Received',
  TOTAL_SKID_COUNT: 'Total_Skid_Count',
  CARRIER: 'Carrier',
  BOL_NUMBER: 'BOL_Number',
  CUSTOMER_PO_NUMBER: 'Customer_PO_Number',
  MANUFACTURER: 'Manufacturer',
  MFPN: 'MFPN',
  DESCRIPTION: 'Description',
  RECEIVED_BY: 'Received_By',
  SKU: 'SKU'
};
// OutboundLog columns
var OUTBOUND_LOG_COLUMNS = {
  DATE: 'Date',
  ORDER_NUMBER: 'Order_Number',
  TASK_NUMBER: 'Task_Number',
  TRANSACTION_TYPE: 'Transaction Type',
  WAREHOUSE: 'Warehouse',
  COMPANY: 'Company',
  PROJECT: 'Project',
  FBPN: 'FBPN',
  MANUFACTURER: 'Manufacturer',
  QTY: 'Qty',
  TOC: 'TOC',
  PO_NUMBER: 'PO_Number',
  SKID_ID: 'Skid_ID',
  SKU: 'SKU'
};
/**
 * Validate configuration
 */
function validateConfig() {
  const errors = [];
// Check template IDs
  if (REPORT_TEMPLATES.INBOUND === '1ABC...') {
    errors.push('Inbound template ID not set');
}
  if (REPORT_TEMPLATES.OUTBOUND === '1DEF...') {
    errors.push('Outbound template ID not set');
}
  
  // Check sheet names exist
  const ss = SpreadsheetApp.getActiveSpreadsheet();
Object.values(SHEET_NAMES).forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      errors.push(`Sheet "${sheetName}" not found`);
    }
  });
if (errors.length > 0) {
    return {
      valid: false,
      errors: errors
    };
}
  
  return {
    valid: true,
    message: 'Configuration is valid'
  };
}
const DRIVE_CONFIG = {
  REPORTS_FOLDER: 'IMS Reports',      // Main folder for all reports
  CREATE_SUBFOLDERS: true,            // Create Inbound/Outbound subfolders
  SUBFOLDER_NAMES: {
    INBOUND: 'Inbound Reports',
    OUTBOUND: 'Outbound Reports'
  }
};
/**
 * Show configuration status
 */
function showConfigStatus() {
  const validation = validateConfig();
if (validation.valid) {
    SpreadsheetApp.getUi().alert('✓ Configuration Valid', validation.message, SpreadsheetApp.getUi().ButtonSet.OK);
} else {
    const errorMessage = 'Configuration Errors:\n\n' + validation.errors.join('\n');
    SpreadsheetApp.getUi().alert('✗ Configuration Errors', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
}
}


// ─────────────────────────────────────────────────────────────────────────────
// SHARED CONFIG ACCESSOR (used by Shipping_Docs / Outbound / other modules)
// ─────────────────────────────────────────────────────────────────────────────
/**
 * Canonical config accessor for the IMS system.
* Shipping_Docs expects these keys:
 * - TOC_TEMPLATE_ID
 * - PACKING_LIST_TEMPLATE_ID
 * - TOC_PACKING_OUTPUT_FOLDER_ID
 * - CUSTOMER_ORDERS_FOLDER_ID
 * - SOURCE_NEW_ORDERS_FOLDER_ID (Added for Automation)
 */
function getIMSConfig() {
  return {
    // Templates
    TOC_TEMPLATE_ID: (typeof TEMPLATES !== 'undefined' && TEMPLATES.TOC_Template) ?
TEMPLATES.TOC_Template : '',
    PACKING_LIST_TEMPLATE_ID: (typeof TEMPLATES !== 'undefined' && TEMPLATES.Packing_List_Template) ?
TEMPLATES.Packing_List_Template : '',
    // ADDED PICK TICKET ID EXPLICITLY HERE FOR SAFETY
    PICK_TICKET_TEMPLATE_ID: (typeof TEMPLATES !== 'undefined' && TEMPLATES.Pickticket_Template) ?
      TEMPLATES.Pickticket_Template : '',

    // Folders (default: save PDFs under PROCESSED_ORDERS)
    TOC_PACKING_OUTPUT_FOLDER_ID: (typeof FOLDERS !== 'undefined' && FOLDERS.PROCESSED_ORDERS) ?
FOLDERS.PROCESSED_ORDERS : ((typeof FOLDERS !== 'undefined' && FOLDERS.IMS_ROOT) ? FOLDERS.IMS_ROOT : ''),
      
    CUSTOMER_ORDERS_FOLDER_ID: (typeof FOLDERS !== 'undefined' && FOLDERS.CUSTOMER_ORDERS) ?
FOLDERS.CUSTOMER_ORDERS : ((typeof FOLDERS !== 'undefined' && FOLDERS.IMS_ROOT) ? FOLDERS.IMS_ROOT : ''),
      
    // Added Source Folder for Auto-Processing
    SOURCE_NEW_ORDERS_FOLDER_ID: (typeof FOLDERS !== 'undefined' && FOLDERS.SOURCE_NEW_ORDERS) ?
      FOLDERS.SOURCE_NEW_ORDERS : ''
  };
}

// ============================================================================
// MAIN CONTROLLER / SHELL FUNCTIONS
// ============================================================================

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  const customerOrdersMenu = ui.createMenu('Customer Orders')
    .addItem('Create Customer Order', 'showCustomerOrderModal')
    .addItem('Cancel Customer Order', 'openCancelOrderModal')
    .addItem('Generate Pick Ticket', 'openPickTicketGenerator')
    .addItem('Process Outbound / Packing & TOC', 'openPackingTOCGenerator')
    .addSeparator()
    .addItem('Sync Delivered Orders (One-Time)', 'runOneTimeDeliveredOrderSync')
    .addItem('Force Update All Delivered (Repair)', 'forceUpdateDeliveredOrders');

  const inventoryToolsMenu = ui.createMenu('Inventory Tools')
    .addItem('Bin Stock Put-Away & Consolidation', 'openStockToolsModal')
    .addItem('Bin Lookup', 'openBinLookupModal')

    .addSeparator()
    .addItem('Cycle Count', 'openCycleCountModal');

  ui.createMenu('IMS')
    .addItem('Open Customer Portal', 'showCustomerPortal')
    .addSeparator()
    .addItem('Inbound Delivery Form', 'openInboundModal')
    .addItem('Inbound Manager (Undo/Labels)', 'openInboundManagerModal')
    .addItem('Create Skid Label', 'openInboundSkidLabelModal')
    .addSeparator()
    .addSubMenu(customerOrdersMenu)
    .addSeparator()
    .addSubMenu(inventoryToolsMenu)
    .addSeparator()
    .addItem('Generate Reports', 'showReportGeneratorModal')
    .addToUi();
}

function onEdit(e) {
  // onEdit trigger - add custom logic here if needed
}

function doGet(e) {
  const context = getUserContext();

  const template = HtmlService.createTemplateFromFile('IMSWebApp');
  template.userContext = JSON.stringify(context);

  // Deep link param: ?skid=SKID_000123
  const skidParam = (e && e.parameter && e.parameter.skid) ? String(e.parameter.skid).trim() : '';
  template.deepLinkSkidId = JSON.stringify(skidParam); // keep JSON-safe

  const html = template.evaluate()
    .setTitle('IMS - Warehouse Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  return html;
}

// ----------------------------------------------------------------------------
// MODAL OPENERS
// ----------------------------------------------------------------------------
function openInboundManagerModal() {
  const html = HtmlService.createTemplateFromFile('InboundManagerModal')
    .evaluate()
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Inbound Manager');
}

function openInboundModal() {
  const html = HtmlService.createTemplateFromFile('InboundModal')
    .evaluate().setWidth(950).setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Inbound Delivery Receiving');
}

function showCustomerOrderModal() {
  const html = HtmlService.createTemplateFromFile('CustomerOrderModal')
    .evaluate().setWidth(900).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Customer Order');
}

function openCancelOrderModal() {
  const html = HtmlService.createTemplateFromFile('CancelOrderModal')
    .evaluate().setWidth(600).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cancel Customer Order');
}

function openPickTicketGenerator() {
  const html = HtmlService.createTemplateFromFile('PickTicketModal')
    .evaluate().setWidth(950).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Pick Ticket');
}

function openPackingTOCGenerator() {
  const html = HtmlService.createTemplateFromFile('PackingTOCModal')
    .evaluate().setWidth(1050).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Packing List + TOC Generator');
}

function showReportGeneratorModal() {
  const html = HtmlService.createTemplateFromFile('ReportGeneratorModal')
    .evaluate().setWidth(950).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Report Generator');
}

function openStockToolsModal() {
  const html = HtmlService.createTemplateFromFile('StockToolsModal')
    .evaluate().setWidth(1050).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bin Stock Put-Away & Consolidation');
}

function openBinLookupModal() {
  const html = HtmlService.createTemplateFromFile('BinLookupModal')
    .evaluate().setWidth(1000).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bin Lookup');
}

function openBinUpdateModal() {
  const html = HtmlService.createTemplateFromFile('BinUpdateModal')
    .evaluate().setWidth(900).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bin Update');
}

function openCycleCountModal() {
  const html = HtmlService.createTemplateFromFile('CycleCountModal')
    .evaluate().setWidth(1100).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cycle Count');
}

function openCurrentItemsModal() {
  const html = HtmlService.createTemplateFromFile('CurrentItemsModal')
    .evaluate().setWidth(1100).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Current Items');
}

function showCustomerPortal() {
  const html = HtmlService.createTemplateFromFile('CustomerPortalUI')
    .evaluate().setTitle('Inventory Portal');
  SpreadsheetApp.getUi().showSidebar(html);
}

function openGeneratePastInboundLabelsByDateModal() {
  const html = HtmlService.createTemplateFromFile('LabelDatePickerModal')
    .evaluate()
    .setWidth(520)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Past Inbound Labels');
}

// Manual Skid Label Modal Functions
function openInboundSkidLabelModal() {
  const html = HtmlService.createTemplateFromFile('InboundSkidLabelModal')
    .evaluate()
    .setWidth(850)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Inbound Skid Label');
}

// ----------------------------------------------------------------------------
// SHELL WRAPPER FUNCTIONS
// Note: Most functions are implemented in their respective modules:
// - IMS_Inbound.js: Inbound processing, labels, manufacturers, FBPNs
// - CustomerOrderBackend.js: Orders, companies, projects
// - SHIPPING_DOCS.js: Pick tickets, TOC, packing lists
// - BinLookup.js: Bin search, details, history
// - BinUpdate.js: Inventory batch operations
// - CycleCount.js: Cycle count functions
// - IMS_Inbound_Manager.js: Inbound undo, search, label regeneration
// - ReportGenerator.js: Report generation
// ----------------------------------------------------------------------------

function shell_generateLabelsForAllPastInbounds(startDate, endDate) {
  if (!startDate) throw new Error('startDate is required');
  return generateLabelsForAllPastInbounds(startDate, endDate || startDate);
}

// ----------------------------------------------------------------------------
// BOL LOOKUP FOR MANUAL LABEL MODAL
// ----------------------------------------------------------------------------
/**
 * Looks up BOL data from Master_Log to pre-populate the manual label form.
 * Returns the first matching record's data (Manufacturer, Project, Push #, Total Skids).
 */
function shell_lookupBOLData(bolNumber) {
  try {
    if (!bolNumber || !String(bolNumber).trim()) {
      return { success: false, message: 'BOL number is required.' };
    }

    const bol = String(bolNumber).trim().toUpperCase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master_Log');

    if (!sheet) {
      return { success: false, message: 'Master_Log sheet not found.' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: false, message: 'No data in Master_Log.' };
    }

    const headers = data[0];
    const colIdx = (name) => headers.indexOf(name);

    const bolCol = colIdx('BOL_Number');
    const mfgCol = colIdx('Manufacturer');
    const projCol = colIdx('Project');
    const pushCol = colIdx('Push #');
    const totalSkidsCol = colIdx('Total_Skid_Count');
    const fbpnCol = colIdx('FBPN');
    const warehouseCol = colIdx('Warehouse');

    if (bolCol < 0) {
      return { success: false, message: 'BOL_Number column not found in Master_Log.' };
    }

    // Find first matching BOL record
    for (let i = 1; i < data.length; i++) {
      const rowBol = String(data[i][bolCol] || '').trim().toUpperCase();
      if (rowBol === bol) {
        return {
          success: true,
          data: {
            manufacturer: mfgCol >= 0 ? String(data[i][mfgCol] || '') : '',
            project: projCol >= 0 ? String(data[i][projCol] || '') : '',
            pushNumber: pushCol >= 0 ? String(data[i][pushCol] || '') : '',
            totalSkids: totalSkidsCol >= 0 ? (parseInt(data[i][totalSkidsCol]) || 1) : 1,
            fbpn: fbpnCol >= 0 ? String(data[i][fbpnCol] || '') : '',
            warehouse: warehouseCol >= 0 ? String(data[i][warehouseCol] || '') : ''
          }
        };
      }
    }

    return { success: false, message: 'BOL not found in Master_Log.' };

  } catch (err) {
    Logger.log('shell_lookupBOLData error: ' + err.toString());
    return { success: false, message: 'Error: ' + err.message };
  }
}

// ----------------------------------------------------------------------------
// CUSTOMER PORTAL FUNCTIONS
// ----------------------------------------------------------------------------
function authenticateUser(email) {
  // Stub - implement customer authentication if needed
  return { success: false, message: 'authenticateUser not implemented' };
}

function searchInventoryForCustomer(email, criteria) {
  // Use getStockTotalsForWebApp for inventory search
  return getStockTotalsForWebApp({ email: email }, criteria);
}

function getAvailableFBPNsForOrder(email) {
  // Return FBPN list for ordering
  return getFBPNList();
}

function submitCustomerOrderFromPortal(email, data) {
  // Use processCustomerOrder for order submission
  return processCustomerOrder(data);
}

function getUserProjectAccess(email) {
  // Get user project access from Customer_Access sheet
  const context = getUserContextDirect_();
  return context.projectAccess || [];
}

function getProjectsForPortal() {
  return getProjects();
}

// ----------------------------------------------------------------------------
// MANUAL SKID LABEL GENERATION
// ----------------------------------------------------------------------------
function shell_generateManualSkidLabel(data) {
  return generateManualSkidLabelFromModal(data);
}

function generateManualSkidLabelFromModal(data) {
  try {
    if (!data.fbpn) throw new Error('FBPN is required.');
    if (!data.qty || data.qty <= 0) throw new Error('Quantity must be greater than 0.');
    if (!data.manufacturer) throw new Error('Manufacturer is required.');
    if (!data.project) throw new Error('Project is required.');

    const copies = Math.min(50, Math.max(1, parseInt(data.copies) || 1));
    const skidNumber = parseInt(data.skidNumber) || 1;
    const totalSkids = parseInt(data.totalSkids) || 1;
    const now = new Date();
    const dateStr = formatDateISO_(now);

    const labelData = [];
    for (let i = 0; i < copies; i++) {
      const skidId = generateRandomId_('SKD-', 8);
      const sku = generateSKU(data.fbpn, data.manufacturer);

      labelData.push({
        skidId: skidId,
        fbpn: String(data.fbpn).toUpperCase().trim(),
        quantity: data.qty,
        sku: sku,
        manufacturer: String(data.manufacturer).trim(),
        project: String(data.project).trim(),
        pushNumber: String(data.push || '').trim(),
        dateReceived: dateStr,
        skidNumber: skidNumber,
        totalSkids: totalSkids,
        notes: String(data.notes || '').trim()
      });
    }

    const bolNumber = String(data.bol || 'MANUAL').trim();
    const result = generateSkidLabels(labelData, { bolNumber: bolNumber });

    if (!result || !result.success) {
      return { success: false, message: (result && result.message) ? result.message : 'Label generation failed.' };
    }

    return {
      success: true,
      pdfUrl: result.pdfUrl || '',
      htmlUrl: result.htmlUrl || '',
      labelCount: labelData.length,
      message: 'Successfully generated ' + labelData.length + ' label(s).'
    };
  } catch (err) {
    Logger.log('generateManualSkidLabelFromModal error: ' + err.toString());
    return { success: false, message: 'Error: ' + err.message };
  }
}

function generateRandomId_(prefix, length) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return prefix + result;
}

// ----------------------------------------------------------------------------
// WEBAPP FUNCTIONS
// ----------------------------------------------------------------------------
function getUserContext() {
  return getUserContextDirect_();
}

function getUserContextDirect_() {
  try {
    const email = Session.getActiveUser().getEmail();

    if (!email) {
      return {
        authenticated: false,
        error: 'Unable to retrieve user email. Please ensure you are signed in.'
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Customer_Access');

    if (!sheet) {
      Logger.log('Customer_Access sheet not found');
      return {
        authenticated: false,
        error: 'System configuration error: Customer_Access sheet not found.'
      };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const emailCol = headers.indexOf('Email');
    const nameCol = headers.indexOf('Name');
    const companyCol = headers.indexOf('Company Name');
    const accessCol = headers.indexOf('Access_Level');
    const projectCol = headers.indexOf('Project_Access');
    const activeCol = headers.indexOf('Active');

    const normalizedEmail = email.toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][emailCol] || '').toLowerCase().trim();

      if (rowEmail === normalizedEmail) {
        const projectAccessRaw = String(data[i][projectCol] || '').trim();
        const projectAccess = projectAccessRaw.toUpperCase() === 'ALL'
          ? ['ALL']
          : projectAccessRaw.split(',').map(p => p.trim()).filter(p => p);

        const activeRaw = data[i][activeCol];
        const isActive = activeRaw === true ||
                         String(activeRaw).toUpperCase() === 'TRUE' ||
                         String(activeRaw).toUpperCase() === 'YES' ||
                         String(activeRaw).toUpperCase() === 'Y' ||
                         String(activeRaw).toUpperCase() === 'ACTIVE' ||
                         activeRaw === 1 ||
                         String(activeRaw) === '1' ||
                         (activeCol < 0);

        if (!isActive) {
          return {
            authenticated: false,
            error: 'Your account has been deactivated.'
          };
        }

        const accessLevel = data[i][accessCol] || 'Standard';

        return {
          authenticated: true,
          email: email,
          name: data[i][nameCol] || email.split('@')[0],
          company: data[i][companyCol] || '',
          accessLevel: accessLevel,
          projectAccess: projectAccess,
          isActive: true,
          permissions: buildPermissionsFromLevel_(accessLevel),
          timestamp: new Date().toISOString()
        };
      }
    }

    Logger.log('Access denied for unregistered user: ' + email);
    return {
      authenticated: false,
      error: 'Access denied. Your email is not registered in the system.'
    };

  } catch (error) {
    Logger.log('getUserContext error: ' + error.toString());
    return {
      authenticated: false,
      error: 'Authentication error: ' + error.message
    };
  }
}

function buildPermissionsFromLevel_(accessLevel) {
  const level = String(accessLevel || '').toUpperCase();
  const isMEI = level === 'MEI';
  const isTurner = level === 'TURNER';

  return {
    isAdmin: isMEI,
    canViewAllOrders: isMEI || isTurner,
    canCreateOrders: true,
    canAccessInventoryOps: isMEI,
    canAccessReports: isMEI || isTurner,
    canGenerateDocs: isMEI,
    canAccessInbound: isMEI,
    canAccessCycleCount: isMEI
  };
}

// ============================================================================
// WEBAPP FORWARDERS - DASHBOARD
// ============================================================================

function getDashboardMetrics() {
  return getDashboardMetricsDirect_();
}

function getDashboardMetricsDirect_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metrics = {
    orders: { pending: 0, processing: 0, shipped: 0 },
    inventory: { totalSKUs: 0, lowStock: 0, outOfStock: 0 },
    inbound: { scheduled: 0, received: 0 }
  };

  try {
    const ordersSheet = ss.getSheetByName('Customer_Orders');
    if (ordersSheet) {
      const ordersData = ordersSheet.getDataRange().getValues();
      const statusCol = ordersData[0].indexOf('Request_Status');
      if (statusCol >= 0) {
        for (let i = 1; i < ordersData.length; i++) {
          const status = String(ordersData[i][statusCol] || '').toLowerCase();
          if (status.includes('pending')) metrics.orders.pending++;
          else if (status.includes('processing') || status.includes('picking')) metrics.orders.processing++;
          else if (status.includes('shipped') || status.includes('delivered')) metrics.orders.shipped++;
        }
      }
    }

    const stockSheet = ss.getSheetByName('Stock_Totals');
    if (stockSheet) {
      metrics.inventory.totalSKUs = Math.max(0, stockSheet.getLastRow() - 1);
    }
  } catch (e) {
    Logger.log('getDashboardMetrics error: ' + e.toString());
  }

  return { success: true, metrics: metrics };
}

// ============================================================================
// WEBAPP FORWARDERS - CUSTOMER ORDERS
// ============================================================================

function getCustomerOrders(context) {
  return getCustomerOrdersDirect_(context);
}

function getCustomerOrdersDirect_(context) {
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
    const colMap = {};
    ['Order_ID', 'Task_Number', 'Project', 'NBD', 'Company', 'Order_Title',
     'Deliver_To', 'Request_Status', 'Stock_Status', 'Created_TS'].forEach(h => {
      colMap[h] = headers.indexOf(h);
    });

    const orders = [];
    const userCompany = (context && context.accessLevel === 'Standard' && context.company)
      ? context.company.toLowerCase() : '';

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const orderId = colMap['Order_ID'] >= 0 ? row[colMap['Order_ID']] : '';
      if (!orderId) continue;

      const rowCompany = colMap['Company'] >= 0 ? String(row[colMap['Company']] || '').toLowerCase() : '';
      if (userCompany && rowCompany !== userCompany) continue;

      orders.push({
        orderId: String(orderId),
        taskNumber: colMap['Task_Number'] >= 0 ? String(row[colMap['Task_Number']] || '') : '',
        project: colMap['Project'] >= 0 ? String(row[colMap['Project']] || '') : '',
        nbd: colMap['NBD'] >= 0 ? formatDateISO_(row[colMap['NBD']]) : '',
        company: colMap['Company'] >= 0 ? String(row[colMap['Company']] || '') : '',
        orderTitle: colMap['Order_Title'] >= 0 ? String(row[colMap['Order_Title']] || '') : '',
        status: colMap['Request_Status'] >= 0 ? String(row[colMap['Request_Status']] || 'Pending') : 'Pending',
        stockStatus: colMap['Stock_Status'] >= 0 ? String(row[colMap['Stock_Status']] || '') : '',
        createdAt: colMap['Created_TS'] >= 0 ? formatDateISO_(row[colMap['Created_TS']]) : ''
      });
    }

    orders.sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));

    return { success: true, orders: orders, totalCount: orders.length };
  } catch (error) {
    return { success: false, message: error.toString(), orders: [] };
  }
}

/**
 * Formats date as ISO (yyyy-MM-dd) for web app use
 */
function formatDateISO_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}

// ============================================================================
// WEBAPP FORWARDERS - DOCUMENT GENERATION
// ============================================================================

function regenerateOrderDoc(orderId, docType) {
  try {
    if (!orderId) {
      return { success: false, message: 'Order ID is required' };
    }

    const orderData = getFullOrderData_(orderId);

    if (!orderData) {
      return { success: false, message: 'Order not found: ' + orderId };
    }

    if (!orderData.items || orderData.items.length === 0) {
      return { success: false, message: 'No items found for order: ' + orderId };
    }

    let result;

    switch (docType) {
      case 'PICK':
        const pickData = {
          orderNumber: orderData.orderNumber,
          orderId: orderData.orderId,
          taskNumber: orderData.taskNumber,
          company: orderData.company,
          project: orderData.project,
          orderTitle: orderData.orderTitle,
          date: new Date().toLocaleDateString(),
          items: orderData.items.map(item => ({
            fbpn: item.fbpn,
            description: item.description,
            qtyRequested: item.qtyRequested,
            qtyToPick: item.qtyRequested,
            qty: item.qtyRequested
          }))
        };

        result = generatePickTicket(pickData);
        break;

      case 'PACKING':
      case 'TOC':
        const skids = [{
          skidNumber: 1,
          items: orderData.items.map(item => ({
            fbpn: item.fbpn,
            description: item.description,
            qtyRequested: item.qtyRequested,
            qtyOnSkid: item.qtyRequested,
            qty: item.qtyRequested
          }))
        }];

        const docData = {
          orderNumber: orderData.orderNumber,
          orderId: orderData.orderId,
          taskNumber: orderData.taskNumber,
          company: orderData.company,
          project: orderData.project,
          orderTitle: orderData.orderTitle,
          deliverTo: orderData.deliverTo || '',
          name: orderData.name || '',
          phoneNumber: orderData.phoneNumber || '',
          shipDate: orderData.shipDate || new Date().toLocaleDateString(),
          date: new Date().toLocaleDateString(),
          totalSkids: '1',
          skids: skids,
          items: orderData.items
        };

        if (docType === 'TOC') {
          result = generateTOC(docData);
        } else {
          result = generatePackingLists(docData);
        }
        break;

      default:
        return { success: false, message: 'Invalid document type: ' + docType };
    }

    if (result && (result.success || result.pdfUrl || result.url)) {
      return {
        success: true,
        url: result.pdfUrl || result.url,
        docType: docType,
        message: docType + ' generated successfully'
      };
    } else {
      return { success: false, message: result ? result.message : 'Document generation failed' };
    }

  } catch (e) {
    Logger.log('regenerateOrderDoc error: ' + e.toString());
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

function getFullOrderData_(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('Customer_Orders');
  if (!orderSheet) return null;

  const orderData = orderSheet.getDataRange().getValues();
  const orderHeaders = orderData[0];

  const findCol = (names) => {
    const normalized = orderHeaders.map(h => String(h || '').toLowerCase().trim().replace(/[_\s]+/g, '_'));
    for (const name of names) {
      const idx = normalized.indexOf(name.toLowerCase().replace(/[_\s]+/g, '_'));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const col = {
    orderId: findCol(['order_id', 'order_number', 'orderid']),
    taskNumber: findCol(['task_number', 'task_number', 'task']),
    project: findCol(['project']),
    nbd: findCol(['nbd', 'need_by_date']),
    company: findCol(['company']),
    orderTitle: findCol(['order_title', 'title']),
    deliverTo: findCol(['deliver_to', 'delivery_address']),
    name: findCol(['name', 'contact_name']),
    phoneNumber: findCol(['phone_number', 'phone']),
    status: findCol(['request_status', 'status']),
    shipDate: findCol(['ship_date', 'shipped_date'])
  };

  let orderRow = null;
  for (let i = 1; i < orderData.length; i++) {
    const rowOrderId = col.orderId >= 0 ? String(orderData[i][col.orderId]) : '';
    if (rowOrderId === String(orderId)) {
      orderRow = orderData[i];
      break;
    }
  }

  if (!orderRow) return null;

  const order = {
    orderNumber: String(orderId),
    orderId: String(orderId),
    taskNumber: col.taskNumber >= 0 ? String(orderRow[col.taskNumber] || '') : '',
    project: col.project >= 0 ? String(orderRow[col.project] || '') : '',
    nbd: col.nbd >= 0 ? formatDateISO_(orderRow[col.nbd]) : '',
    company: col.company >= 0 ? String(orderRow[col.company] || '') : '',
    orderTitle: col.orderTitle >= 0 ? String(orderRow[col.orderTitle] || '') : '',
    deliverTo: col.deliverTo >= 0 ? String(orderRow[col.deliverTo] || '') : '',
    name: col.name >= 0 ? String(orderRow[col.name] || '') : '',
    phoneNumber: col.phoneNumber >= 0 ? String(orderRow[col.phoneNumber] || '') : '',
    status: col.status >= 0 ? String(orderRow[col.status] || '') : '',
    shipDate: col.shipDate >= 0 ? formatDateISO_(orderRow[col.shipDate]) : '',
    date: new Date().toLocaleDateString()
  };

  const itemsSheet = ss.getSheetByName('Requested_Items');
  order.items = [];

  if (itemsSheet) {
    const itemsData = itemsSheet.getDataRange().getValues();
    const itemHeaders = itemsData[0];

    const findItemCol = (names) => {
      const normalized = itemHeaders.map(h => String(h || '').toLowerCase().trim().replace(/[_\s]+/g, '_'));
      for (const name of names) {
        const idx = normalized.indexOf(name.toLowerCase().replace(/[_\s]+/g, '_'));
        if (idx >= 0) return idx;
      }
      return -1;
    };

    const itemCol = {
      orderId: findItemCol(['order_id', 'order_number']),
      fbpn: findItemCol(['fbpn']),
      description: findItemCol(['description', 'desc']),
      qtyRequested: findItemCol(['qty_requested', 'qty', 'quantity']),
      sku: findItemCol(['sku'])
    };

    for (let i = 1; i < itemsData.length; i++) {
      const rowOrderId = itemCol.orderId >= 0 ? String(itemsData[i][itemCol.orderId]) : '';

      if (rowOrderId === String(orderId) || rowOrderId === String(Math.trunc(Number(orderId)))) {
        const fbpn = itemCol.fbpn >= 0 ? String(itemsData[i][itemCol.fbpn] || '').trim() : '';
        if (!fbpn) continue;

        order.items.push({
          fbpn: fbpn,
          description: itemCol.description >= 0 ? String(itemsData[i][itemCol.description] || '') : '',
          qtyRequested: itemCol.qtyRequested >= 0 ? Number(itemsData[i][itemCol.qtyRequested] || 0) : 0,
          qty: itemCol.qtyRequested >= 0 ? Number(itemsData[i][itemCol.qtyRequested] || 0) : 0,
          sku: itemCol.sku >= 0 ? String(itemsData[i][itemCol.sku] || '') : ''
        });
      }
    }
  }

  return order;
}

// ============================================================================
// WEBAPP FORWARDERS - ORDER DATA FOR SHIPPING MODAL
// ============================================================================

function getOrderDataForShipping(orderId) {
  return getFullOrderData_(orderId);
}

function getOrderByTaskNumber(taskNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('Customer_Orders');
  const itemsSheet = ss.getSheetByName('Requested_Items');

  if (!orderSheet || !itemsSheet) return null;

  const orderData = orderSheet.getDataRange().getValues();
  const orderHeaders = orderData[0];

  const findCol = (headers, names) => {
    const normalized = headers.map(h => String(h || '').toLowerCase().trim().replace(/[_\s]+/g, '_'));
    for (const name of names) {
      const idx = normalized.indexOf(name.toLowerCase().replace(/[_\s]+/g, '_'));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const cTask = findCol(orderHeaders, ['task_number', 'task']);
  const cOrder = findCol(orderHeaders, ['order_id', 'order_number']);
  const cProj = findCol(orderHeaders, ['project']);
  const cComp = findCol(orderHeaders, ['company']);
  const cTitle = findCol(orderHeaders, ['order_title', 'title']);
  const cDeliver = findCol(orderHeaders, ['deliver_to', 'delivery_address']);
  const cName = findCol(orderHeaders, ['name', 'contact_name']);
  const cPhone = findCol(orderHeaders, ['phone_number', 'phone']);

  const key = String(taskNumber).trim();
  let orderRow = null;

  for (let r = 1; r < orderData.length; r++) {
    const row = orderData[r];
    const vTask = cTask >= 0 ? String(row[cTask] || '').trim() : '';
    const vOrder = cOrder >= 0 ? String(row[cOrder] || '').trim() : '';

    if (vTask === key || vOrder === key ||
        String(Math.trunc(Number(vTask))) === key ||
        String(Math.trunc(Number(vOrder))) === key) {
      orderRow = row;
      break;
    }
  }

  if (!orderRow) return null;

  const orderId = (cOrder >= 0 ? String(orderRow[cOrder] || '') : '') || key;

  const itemsData = itemsSheet.getDataRange().getValues();
  const itemHeaders = itemsData[0];

  const iOrder = findCol(itemHeaders, ['order_id', 'order_number', 'task_number']);
  const iFbpn = findCol(itemHeaders, ['fbpn']);
  const iDesc = findCol(itemHeaders, ['description', 'desc']);
  const iQty = findCol(itemHeaders, ['qty_requested', 'qty']);

  const items = [];
  const matchKeys = [key, orderId, String(Math.trunc(Number(key))), String(Math.trunc(Number(orderId)))];

  for (let j = 1; j < itemsData.length; j++) {
    const row = itemsData[j];
    const ok = iOrder >= 0 ? String(row[iOrder] || '').trim() : '';

    if (!matchKeys.includes(ok)) continue;

    const fbpn = iFbpn >= 0 ? String(row[iFbpn] || '').trim() : '';
    if (!fbpn) continue;

    items.push({
      fbpn: fbpn,
      description: iDesc >= 0 ? String(row[iDesc] || '') : '',
      qtyRequested: iQty >= 0 ? Number(row[iQty] || 0) : 0
    });
  }

  const combined = {};
  items.forEach(item => {
    if (!combined[item.fbpn]) {
      combined[item.fbpn] = item;
    } else {
      combined[item.fbpn].qtyRequested += item.qtyRequested;
    }
  });

  return {
    orderId: orderId,
    orderNumber: orderId,
    orderTitle: cTitle >= 0 ? String(orderRow[cTitle] || '') : '',
    company: cComp >= 0 ? String(orderRow[cComp] || '') : '',
    project: cProj >= 0 ? String(orderRow[cProj] || '') : '',
    deliverTo: cDeliver >= 0 ? String(orderRow[cDeliver] || '') : '',
    name: cName >= 0 ? String(orderRow[cName] || '') : '',
    phoneNumber: cPhone >= 0 ? String(orderRow[cPhone] || '') : '',
    items: Object.values(combined)
  };
}

// ============================================================================
// WEBAPP FORWARDERS - FILE UPLOAD
// ============================================================================

const AUTOMATION_FOLDER_ID = '1L3mjeQizzjVU5uTqGxv1sOUOuq25I2pM';

function uploadToAutomationFolder(fileData) {
  try {
    if (!fileData || !fileData.content || !fileData.fileName) {
      return { success: false, message: 'Missing file data' };
    }

    const folder = DriveApp.getFolderById(AUTOMATION_FOLDER_ID);
    const decoded = Utilities.base64Decode(fileData.content);
    const blob = Utilities.newBlob(decoded,
      fileData.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      fileData.fileName);

    const file = folder.createFile(blob);

    return {
      success: true,
      message: 'File uploaded. It will be processed automatically.',
      fileUrl: file.getUrl()
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// WEBAPP FORWARDERS - FORM HELPERS
// ============================================================================

function getCompaniesFiltered(context) {
  if (context && context.accessLevel === 'Standard' && context.company) {
    return [context.company];
  }
  return getCompaniesDirect_();
}

function getCompaniesDirect_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Support_Sheet');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const companies = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) companies.add(data[i][0]);
  }
  return Array.from(companies).sort();
}

function getProjectsFiltered(company) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Support_Sheet');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const projects = new Set();

  for (let i = 1; i < data.length; i++) {
    const rowCompany = String(data[i][0] || '').trim();
    const rowProject = String(data[i][1] || '').trim();

    if (rowProject) {
      if (company && rowCompany.toLowerCase() === company.toLowerCase()) {
        projects.add(rowProject);
      } else if (!company) {
        projects.add(rowProject);
      }
    }
  }
  return Array.from(projects).sort();
}

function getNextTaskNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet) return '1001';

  const data = sheet.getDataRange().getValues();
  const taskCol = data[0].indexOf('Task_Number');
  if (taskCol < 0) return '1001';

  let maxNum = 1000;
  for (let i = 1; i < data.length; i++) {
    const num = parseInt(String(data[i][taskCol]).replace(/\D/g, ''), 10);
    if (!isNaN(num) && num > maxNum) maxNum = num;
  }
  return String(maxNum + 1);
}

function validateFBPN(fbpn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Project_Master');
  if (!sheet) return { valid: false, message: 'Project_Master not found' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const fbpnCol = headers.indexOf('FBPN');
  const descCol = headers.indexOf('Description');

  if (fbpnCol < 0) return { valid: false, message: 'FBPN column not found' };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][fbpnCol]).toLowerCase() === fbpn.toLowerCase()) {
      return {
        valid: true,
        fbpn: data[i][fbpnCol],
        description: descCol >= 0 ? String(data[i][descCol] || '') : ''
      };
    }
  }
  return { valid: false, message: 'FBPN not found' };
}

function api_getSkidById(skidId) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Inbound_Skids');
  if (!sh) return { success: false, message: 'Inbound_Skids not found' };

  const id = String(skidId || '').trim().toUpperCase();
  const values = sh.getDataRange().getValues();
  const headers = values[0];

  const c = name => headers.indexOf(name);
  const idxSkid = c('Skid_ID');
  if (idxSkid === -1) return { success:false, message:'Skid_ID column missing' };

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idxSkid]).toUpperCase() === id) {
      return {
        success: true,
        skid: {
          Skid_ID: id,
          FBPN: values[r][c('FBPN')],
          MFPN: values[r][c('MFPN')],
          Project: values[r][c('Project')],
          Qty: values[r][c('Qty_on_Skid')],
          SKU: values[r][c('SKU')],
          TXN_ID: values[r][c('TXN_ID')],
          Timestamp: values[r][c('Timestamp')]
        }
      };
    }
  }
  return { success:false, message:`Skid not found: ${id}` };
}