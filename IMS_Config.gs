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
  "Master_Log": ["Txn_ID", "Date_Received", "Transaction_Type", "Warehouse", "Push #", "FBPN", "Qty_Received", "Total_Skid_Count", "Carrier", "BOL_Number", "Customer_PO_Number", "Manufacturer", "MFPN", "Description", "Received_By", "SKU"],
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