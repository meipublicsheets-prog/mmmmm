// ============================================================================
// IMS_Shell.gs - MAIN CONTROLLER (Library Version)
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
  if (typeof IMS_LIB !== 'undefined' && typeof IMS_LIB.onEdit === 'function') {
    return IMS_LIB.onEdit(e);
  }
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
// ----------------------------------------------------------------------------
// LIBRARY STUBS
// ----------------------------------------------------------------------------
function generateLabelsForAllPastInbounds() {
  return IMS_LIB.generateLabelsForAllPastInbounds(); 
}
function shell_generateLabelsForAllPastInbounds(startDate, endDate) {
  if (!startDate) throw new Error('startDate is required');
  // If your library expects (start,end) you’re good; if it expects single day, pass same twice.
  return IMS_LIB.generateLabelsForAllPastInbounds(startDate, endDate || startDate);
}

// Inbound Functions (Now working in library)
function lookupProjectFromPO(customerPO) { return IMS_LIB.lookupProjectFromPO(customerPO); }

function processInboundSubmission(payload) {
  if (!payload) payload = {};
  if (!payload.options) payload.options = {};

  // Force labels ON (required)
  payload.options.generateLabels = true;

  const res = IMS_LIB.processInboundSubmission(payload);

  // Normalize for InboundModal.html (expects labelPdfUrl)
  if (res && res.success) {
    const pdfUrl =
      (res.labelPdfUrl) ||
      (res.labelResult && (res.labelResult.labelFileUrl || res.labelResult.pdfUrl)) ||
      '';

    const htmlUrl =
      (res.labelHtmlUrl) ||
      (res.labelResult && (res.labelResult.labelHtmlUrl || res.labelResult.htmlUrl)) ||
      '';

    res.labelPdfUrl = String(pdfUrl || '').trim();
    res.labelHtmlUrl = String(htmlUrl || '').trim();
  }

  return res;
}
function generateSkidLabelsHtml(labelData) {
  return IMS_LIB.generateSkidLabelsHtml_(labelData);
}
function getManufacturers() { return IMS_LIB.getManufacturers(); }
function saveLabelsToDrive(html) { return IMS_LIB.saveLabelsToDrive(html); }
function processPendingOrderUploads() { return IMS_LIB.processPendingOrderUploads(); }


// Other Library Functions
function getIMSConfig() { return IMS_LIB.getIMSConfig(); }
function getCompanies() { return IMS_LIB.getCompanies(); }
function getProjects() { return IMS_LIB.getProjects(); }
function processCustomerOrder(data) { return IMS_LIB.processCustomerOrder(data); }
function cancelOrder(id) { return IMS_LIB.cancelOrder(id); }
function getOutOfStockItems(email) { return IMS_LIB.getOutOfStockItems(email); }

function getTaskNumbers() { return IMS_LIB.getTaskNumbers(); }
function getOrderByTaskNumber(task) { return IMS_LIB.getOrderByTaskNumber(task); }
function generatePickTicket(data) { return IMS_LIB.generatePickTicket(data); }
function processPackingTOCAndShipment(data) { return IMS_LIB.processPackingTOCAndShipment(data); }
function processPackingTOC_DocsOnly(data) { return IMS_LIB.processPackingTOC_DocsOnly(data); }
function generateTOC(data) { return IMS_LIB.generateTOC(data); }
function generatePackingLists(data) { return IMS_LIB.generatePackingLists(data); }

function runOneTimeDeliveredOrderSync() {
  if (typeof IMS_LIB.runOneTimeDeliveredOrderSync === 'function') return IMS_LIB.runOneTimeDeliveredOrderSync();
  if (typeof IMS_LIB.forceUpdateDeliveredOrders === 'function') return IMS_LIB.forceUpdateDeliveredOrders();
}

function forceUpdateDeliveredOrders() { return IMS_LIB.forceUpdateDeliveredOrders(); }

function getFBPNList() { return IMS_LIB.getFBPNList(); }
function getNextSkidIdBase() { return IMS_LIB.getNextSkidIdBase(); }
function getNextStagingSequence() { return IMS_LIB.getNextStagingSequence(); }
function updateInboundStaging(stagingRows) { return IMS_LIB.updateInboundStaging(stagingRows); }
function generateSKU(fbpn, manufacturer) { return IMS_LIB.generateSKU(fbpn, manufacturer); }
function buildStockSkuFromFBPNAndManufacturer(fbpn, manufacturer) {
  return IMS_LIB.buildStockSkuFromFBPNAndManufacturer(fbpn, manufacturer);
}

function fulfillBackorders(ss, fbpn, qtyReceived, txnId) {
  return IMS_LIB.fulfillBackorders(ss, fbpn, qtyReceived, txnId);
}
function updateAllocationWithFulfillment(ss, backorderId, orderId, fbpn, qtyFulfilled) {
  return IMS_LIB.updateAllocationWithFulfillment(ss, backorderId, orderId, fbpn, qtyFulfilled);
}
function updateCustomerOrderStockStatus(ss, orderId) {
  return IMS_LIB.updateCustomerOrderStockStatus(ss, orderId);
}
function addInventoryBatch(data) { return IMS_LIB.addInventoryBatch(data); }
function removeInventoryBatch(data) { return IMS_LIB.removeInventoryBatch(data); }
function moveInventoryBatch(data) { return IMS_LIB.moveInventoryBatch(data); }
function moveFromStaging(data) { return IMS_LIB.moveFromStaging(data); }
function transferInventoryBatch(data) {
  if (typeof IMS_LIB.transferInventoryBatch === 'function') return IMS_LIB.transferInventoryBatch(data);
  return IMS_LIB.moveInventoryBatch(data);
}
function searchBins(params) { return IMS_LIB.searchBins(params); }
function getBinDetails(bin) { return IMS_LIB.getBinDetails(bin); }
function getBinHistory(bin) { return IMS_LIB.getBinHistory(bin); }
function quickBarcodeScan(val) { return IMS_LIB.quickBarcodeScan(val); }

function imsGetCycleCountBins(filter) { return IMS_LIB.imsGetCycleCountBins(filter); }
function imsCreateCycleCountBatch(data) { return IMS_LIB.imsCreateCycleCountBatch(data); }
function imsSubmitCycleCountLine(data) { return IMS_LIB.imsSubmitCycleCountLine(data); }
function imsGetCycleBatch(id) { return IMS_LIB.imsGetCycleBatch(id); }

function generateInboundReport(params) { return IMS_LIB.generateInboundReport(params); }
function generateOutboundReport(params) { return IMS_LIB.generateOutboundReport(params); }
function getUserContext() {
  // Library needs spreadsheet ID since getActiveSpreadsheet() won't work in web app context
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  if (typeof IMS_LIB !== 'undefined' && IMS_LIB.getUserContextForWebApp) {
    return IMS_LIB.getUserContextForWebApp(ssId);
  }
  
  // Fallback: Direct implementation if library function doesn't exist
  return getUserContextDirect_();
}

/**
 * Direct implementation fallback for getUserContext
 * Used if library doesn't have getUserContextForWebApp
 */
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
        // Handle various formats: TRUE, true, Yes, Y, 1, or checkbox true
        const isActive = activeRaw === true || 
                         String(activeRaw).toUpperCase() === 'TRUE' || 
                         String(activeRaw).toUpperCase() === 'YES' || 
                         String(activeRaw).toUpperCase() === 'Y' ||
                         String(activeRaw).toUpperCase() === 'ACTIVE' ||
                         activeRaw === 1 ||
                         String(activeRaw) === '1' ||
                         (activeCol < 0); // If no Active column exists, assume active
        
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
    
    // User not found - deny access
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

/**
 * Build permissions object from access level
 */
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
  if (typeof IMS_LIB !== 'undefined' && IMS_LIB.getDashboardMetrics) {
    return IMS_LIB.getDashboardMetrics();
  }
  return getDashboardMetricsDirect_();
}

function getDashboardMetricsDirect_() {
  // Basic implementation - returns sample metrics
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metrics = {
    orders: { pending: 0, processing: 0, shipped: 0 },
    inventory: { totalSKUs: 0, lowStock: 0, outOfStock: 0 },
    inbound: { scheduled: 0, received: 0 }
  };
  
  try {
    // Count orders by status
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
    
    // Count SKUs
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
  if (typeof IMS_LIB !== 'undefined' && IMS_LIB.getCustomerOrdersForWebApp) {
    return IMS_LIB.getCustomerOrdersForWebApp(context);
  }
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
        nbd: colMap['NBD'] >= 0 ? formatDate_(row[colMap['NBD']]) : '',
        company: colMap['Company'] >= 0 ? String(row[colMap['Company']] || '') : '',
        orderTitle: colMap['Order_Title'] >= 0 ? String(row[colMap['Order_Title']] || '') : '',
        status: colMap['Request_Status'] >= 0 ? String(row[colMap['Request_Status']] || 'Pending') : 'Pending',
        stockStatus: colMap['Stock_Status'] >= 0 ? String(row[colMap['Stock_Status']] || '') : '',
        createdAt: colMap['Created_TS'] >= 0 ? formatDate_(row[colMap['Created_TS']]) : ''
      });
    }
    
    orders.sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
    
    return { success: true, orders: orders, totalCount: orders.length };
  } catch (error) {
    return { success: false, message: error.toString(), orders: [] };
  }
}

function formatDate_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}

// ============================================================================
// WEBAPP FORWARDERS - DOCUMENT GENERATION
// ============================================================================

/**
 * Regenerate a document for an existing order
 * Matches the data format expected by SHIPPING_DOCS.js functions
 */
function regenerateOrderDoc(orderId, docType) {
  try {
    if (!orderId) {
      return { success: false, message: 'Order ID is required' };
    }
    
    // Get full order data with items
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
        // Pick Ticket uses orderNumber to look up from Pick_Log, or falls back to items
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
        
        if (typeof IMS_LIB !== 'undefined' && typeof IMS_LIB.generatePickTicket === 'function') {
          result = IMS_LIB.generatePickTicket(pickData);
        } else {
          return { success: false, message: 'generatePickTicket function not available in library' };
        }
        break;
        
      case 'PACKING':
      case 'TOC':
        // TOC and Packing Lists require skids structure
        // Build skids from items - put all items on one skid as default
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
          if (typeof IMS_LIB !== 'undefined' && typeof IMS_LIB.generateTOC === 'function') {
            result = IMS_LIB.generateTOC(docData);
          } else {
            return { success: false, message: 'generateTOC function not available in library' };
          }
        } else {
          if (typeof IMS_LIB !== 'undefined' && typeof IMS_LIB.generatePackingLists === 'function') {
            result = IMS_LIB.generatePackingLists(docData);
          } else {
            return { success: false, message: 'generatePackingLists function not available in library' };
          }
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

/**
 * Get full order data including items for document generation
 */
function getFullOrderData_(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('Customer_Orders');
  if (!orderSheet) return null;
  
  const orderData = orderSheet.getDataRange().getValues();
  const orderHeaders = orderData[0];
  
  // Build column map with flexible header matching
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
  
  // Find the order
  let orderRow = null;
  for (let i = 1; i < orderData.length; i++) {
    const rowOrderId = col.orderId >= 0 ? String(orderData[i][col.orderId]) : '';
    if (rowOrderId === String(orderId)) {
      orderRow = orderData[i];
      break;
    }
  }
  
  if (!orderRow) return null;
  
  // Build order object with all required fields
  const order = {
    orderNumber: String(orderId),
    orderId: String(orderId),
    taskNumber: col.taskNumber >= 0 ? String(orderRow[col.taskNumber] || '') : '',
    project: col.project >= 0 ? String(orderRow[col.project] || '') : '',
    nbd: col.nbd >= 0 ? formatDate_(orderRow[col.nbd]) : '',
    company: col.company >= 0 ? String(orderRow[col.company] || '') : '',
    orderTitle: col.orderTitle >= 0 ? String(orderRow[col.orderTitle] || '') : '',
    deliverTo: col.deliverTo >= 0 ? String(orderRow[col.deliverTo] || '') : '',
    name: col.name >= 0 ? String(orderRow[col.name] || '') : '',
    phoneNumber: col.phoneNumber >= 0 ? String(orderRow[col.phoneNumber] || '') : '',
    status: col.status >= 0 ? String(orderRow[col.status] || '') : '',
    shipDate: col.shipDate >= 0 ? formatDate_(orderRow[col.shipDate]) : '',
    date: new Date().toLocaleDateString()
  };
  
  // Get items from Requested_Items
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
      
      // Match order ID (handle numeric vs string)
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

/**
 * Get order data formatted for the shipping docs modal
 * This is called when opening the shipping modal from Customer Orders
 */
function getOrderDataForShipping(orderId) {
  return getFullOrderData_(orderId);
}

/**
 * Get order by task number - wrapper for library function
 */
function getOrderByTaskNumber(taskNumber) {
  if (typeof IMS_LIB !== 'undefined' && typeof IMS_LIB.getOrderByTaskNumber === 'function') {
    return IMS_LIB.getOrderByTaskNumber(taskNumber);
  }
  
  // Fallback implementation
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('Customer_Orders');
  const itemsSheet = ss.getSheetByName('Requested_Items');
  
  if (!orderSheet || !itemsSheet) return null;
  
  const orderData = orderSheet.getDataRange().getValues();
  const orderHeaders = orderData[0];
  
  // Find column indices with flexible matching
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
  
  // Find order row
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
  
  // Get items
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
  
  // Combine duplicate FBPNs
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
  if (typeof IMS_LIB !== 'undefined' && IMS_LIB.uploadToAutomationFolder) {
    return IMS_LIB.uploadToAutomationFolder(fileData);
  }
  
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
  return typeof IMS_LIB !== 'undefined' && IMS_LIB.getCompanies ? 
    IMS_LIB.getCompanies() : getCompaniesDirect_();
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
function searchInboundByBOL(bol) { return IMS_LIB.searchInboundByBOL(bol); }
function executeUndoByTxnId(txnId) { return IMS_LIB.executeUndoByTxnId(txnId); }
function getRecentInboundTransactions(limit) { return IMS_LIB.getRecentInboundTransactions(limit); }
function regenerateLabelsForTxn(txnId) { return IMS_LIB.regenerateLabelsForTxn(txnId); }
function generateManualLabel(data) { return IMS_LIB.generateManualLabel(data); }
function generateLabelsByBOL(bol) { return IMS_LIB.generateLabelsByBOL(bol); }

// Manual Skid Label Modal Functions
function openInboundSkidLabelModal() {
  const html = HtmlService.createTemplateFromFile('InboundSkidLabelModal')
    .evaluate()
    .setWidth(850)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Inbound Skid Label');
}

function shell_generateManualSkidLabel(data) {
  // Generate label using the library or local function
  if (typeof IMS_LIB !== 'undefined' && typeof IMS_LIB.generateManualSkidLabelFromModal === 'function') {
    return IMS_LIB.generateManualSkidLabelFromModal(data);
  }
  // Fallback to local implementation if library doesn't have it
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
    const dateStr = formatDate_(now);

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
    const result = IMS_LIB.generateSkidLabels(labelData, { bolNumber: bolNumber });

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

function authenticateUser(email) { return IMS_LIB.authenticateUser(email); }
function searchInventoryForCustomer(email, criteria) { return IMS_LIB.searchInventoryForCustomer(email, criteria); }
function getCustomerOrders(email) { return IMS_LIB.getCustomerOrders(email); }
function getAvailableFBPNsForOrder(email) { return IMS_LIB.getAvailableFBPNsForOrder(email); }
function submitCustomerOrderFromPortal(email, data) { return IMS_LIB.submitCustomerOrderFromPortal(email, data); }
function getUserProjectAccess(email) { return IMS_LIB.getUserProjectAccess(email); }
function getProjectsForPortal() { return IMS_LIB.getProjectsForPortal(); }