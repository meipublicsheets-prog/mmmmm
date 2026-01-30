// ============================================================================
// IMS_WEBAPP.GS (LIBRARY) - WebApp API + Auth Context + Data Endpoints
// Intended to be added to IMS_LIB and called by IMS_Shell webapp forwarders.
// ============================================================================

/**
 * Library-safe spreadsheet opener (WebApp context cannot rely on ActiveSpreadsheet).
 */
function imsw_openSpreadsheet_(ssId) {
  if (!ssId) throw new Error('Spreadsheet ID required.');
  return SpreadsheetApp.openById(ssId);
}

function imsw_tabs_() {
  // Prefer your canonical TABS map if present.
  if (typeof TABS !== 'undefined' && TABS) return TABS;
  return {
    CUSTOMER_ACCESS: 'Customer_Access',
    CUSTOMER_ORDERS: 'Customer_Orders',
    REQUESTED_ITEMS: 'Requested_Items',
    STOCK_TOTALS: 'Stock_Totals',
    MASTER_LOG: 'Master_Log',
    INBOUND_SKIDS: 'Inbound_Skids',
    TRUCK_SCHEDULE: 'Truck_Schedule',
    BACKORDERS: 'Backorders',
    ALLOCATION_LOG: 'Allocation_Log',
    OUTBOUNDLOG: 'OutboundLog'
  };
}

function imsw_normHeader_(h) {
  return String(h || '')
    .toLowerCase()
    .trim()
    .replace(/[\s]+/g, '_')
    .replace(/[^a-z0-9_]/g, '');
}

function imsw_col_(headers, aliases) {
  const norm = (headers || []).map(imsw_normHeader_);
  for (const a of aliases) {
    const want = imsw_normHeader_(a);
    const idx = norm.indexOf(want);
    if (idx >= 0) return idx;
  }
  return -1;
}

function imsw_safeNum_(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function imsw_upper_(v) {
  return String(v || '').trim().toUpperCase();
}

function imsw_buildPermissionsFromLevel_(accessLevel) {
  const lvl = String(accessLevel || 'Standard').trim().toLowerCase();

  const isAdmin = lvl.includes('admin');
  const isEmployee = lvl.includes('employee') || lvl.includes('warehouse') || lvl.includes('staff');
  const isStakeholder = lvl.includes('stake') || lvl.includes('manager') || lvl.includes('pm') || lvl.includes('lead');
  const isCustomer = lvl.includes('standard') || lvl.includes('customer') || lvl === '';

  // Conservative defaults: customer can view inventory + their orders.
  const perms = {
    isAdmin,
    isEmployee,
    isStakeholder,
    isCustomer,

    // Customer portal
    canViewInventory: true,
    canViewOrders: true,
    canCreateOrders: true,

    // Employee/stakeholder tools
    canInbound: isAdmin || isEmployee || isStakeholder,
    canOutbound: isAdmin || isEmployee || isStakeholder,
    canStockTools: isAdmin || isEmployee || isStakeholder,
    canReports: isAdmin || isEmployee || isStakeholder,
    canReprintLabels: isAdmin || isEmployee || isStakeholder,
    canGenerateDocs: isAdmin || isEmployee || isStakeholder,

    // Admin
    canAdmin: isAdmin
  };

  // Lock down customer portal for true “Standard”
  if (isCustomer && !(isAdmin || isEmployee || isStakeholder)) {
    perms.canInbound = false;
    perms.canOutbound = false;
    perms.canStockTools = false;
    perms.canReports = false;
    perms.canReprintLabels = false;
    perms.canGenerateDocs = false;
    perms.canAdmin = false;
  }

  return perms;
}

// ============================================================================
// 1) AUTH CONTEXT (called by IMS_Shell.getUserContext())
// ============================================================================

/**
 * Must match IMS_Shell expectation:
 *   IMS_LIB.getUserContextForWebApp(ssId)
 */
function getUserContextForWebApp(ssId) {
  try {
    const email = Session.getActiveUser().getEmail();

    if (!email) {
      return { authenticated: false, error: 'Unable to retrieve user email. Please ensure you are signed in.' };
    }

    const ss = imsw_openSpreadsheet_(ssId);
    const tabs = imsw_tabs_();
    const sheet = ss.getSheetByName(tabs.CUSTOMER_ACCESS) || ss.getSheetByName('Customer_Access');

    if (!sheet) {
      return { authenticated: false, error: 'System configuration error: Customer_Access sheet not found.' };
    }

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      return { authenticated: false, error: 'Customer_Access is empty.' };
    }

    const headers = data[0];

    // Prefer your exact header casing but tolerate variants
    const cEmail = imsw_col_(headers, ['Email', 'email']);
    const cName = imsw_col_(headers, ['Name', 'name']);
    const cCompany = imsw_col_(headers, ['Company Name', 'Company', 'company_name', 'company']);
    const cAccess = imsw_col_(headers, ['Access_Level', 'access_level', 'role', 'access']);
    const cProjects = imsw_col_(headers, ['Project_Access', 'project_access', 'projects']);
    const cActive = imsw_col_(headers, ['Active', 'active', 'enabled', 'status']);

    const want = String(email).toLowerCase().trim();

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const rowEmail = String(cEmail >= 0 ? row[cEmail] : '').toLowerCase().trim();
      if (!rowEmail || rowEmail !== want) continue;

      const activeRaw = cActive >= 0 ? row[cActive] : true;
      const isActive =
        activeRaw === true ||
        String(activeRaw).toUpperCase() === 'TRUE' ||
        String(activeRaw).toUpperCase() === 'YES' ||
        String(activeRaw).toUpperCase() === 'Y' ||
        String(activeRaw).toUpperCase() === 'ACTIVE' ||
        activeRaw === 1 ||
        String(activeRaw) === '1' ||
        cActive < 0;

      if (!isActive) return { authenticated: false, error: 'Your account has been deactivated.' };

      const projectAccessRaw = String(cProjects >= 0 ? row[cProjects] : '').trim();
      const projectAccess =
        imsw_upper_(projectAccessRaw) === 'ALL'
          ? ['ALL']
          : projectAccessRaw.split(',').map(s => s.trim()).filter(Boolean);

      const accessLevel = String(cAccess >= 0 ? row[cAccess] : 'Standard') || 'Standard';
      const company = String(cCompany >= 0 ? row[cCompany] : '') || '';

      return {
        authenticated: true,
        email,
        name: String(cName >= 0 ? row[cName] : '') || email.split('@')[0],
        company,
        accessLevel,
        projectAccess,
        isActive: true,
        permissions: imsw_buildPermissionsFromLevel_(accessLevel),
        timestamp: new Date().toISOString()
      };
    }

    return { authenticated: false, error: 'Access denied: user not found in Customer_Access.' };
  } catch (err) {
    Logger.log('getUserContextForWebApp error: ' + err);
    return { authenticated: false, error: 'Error determining access: ' + err.toString() };
  }
}

// ============================================================================
// 2) DASHBOARD METRICS (called by IMS_Shell.getDashboardMetrics())
// ============================================================================

function getDashboardMetrics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); // library might be invoked from container context too
    const tabs = imsw_tabs_();

    const metrics = {
      orders: { pending: 0, processing: 0, shipped: 0 },
      inventory: { totalSKUs: 0, lowStock: 0, outOfStock: 0 },
      inbound: { scheduled: 0, receivedToday: 0 }
    };

    // Orders by status
    const ordersSh = ss.getSheetByName(tabs.CUSTOMER_ORDERS) || ss.getSheetByName('Customer_Orders');
    if (ordersSh) {
      const od = ordersSh.getDataRange().getValues();
      if (od.length >= 2) {
        const h = od[0];
        const cStatus = imsw_col_(h, ['Request_Status', 'request_status', 'status']);
        for (let i = 1; i < od.length; i++) {
          const st = String(cStatus >= 0 ? od[i][cStatus] : '').toLowerCase();
          if (!st) continue;
          if (st.includes('pending') || st.includes('accepted') || st.includes('new')) metrics.orders.pending++;
          else if (st.includes('processing') || st.includes('picking') || st.includes('allocated')) metrics.orders.processing++;
          else if (st.includes('shipped') || st.includes('delivered') || st.includes('complete')) metrics.orders.shipped++;
        }
      }
    }

    // Inventory counts
    const stockSh = ss.getSheetByName(tabs.STOCK_TOTALS) || ss.getSheetByName('Stock_Totals');
    if (stockSh) {
      const sd = stockSh.getDataRange().getValues();
      metrics.inventory.totalSKUs = Math.max(0, sd.length - 1);

      if (sd.length >= 2) {
        const h = sd[0];
        const cAvail = imsw_col_(h, ['Qty_Available', 'qty_available', 'available', 'qty_available_total']);
        for (let i = 1; i < sd.length; i++) {
          const avail = cAvail >= 0 ? imsw_safeNum_(sd[i][cAvail]) : 0;
          if (avail <= 0) metrics.inventory.outOfStock++;
          else if (avail <= 5) metrics.inventory.lowStock++;
        }
      }
    }

    // Inbound received today (Master_Log)
    const masterSh = ss.getSheetByName(tabs.MASTER_LOG) || ss.getSheetByName('Master_Log');
    if (masterSh) {
      const md = masterSh.getDataRange().getValues();
      if (md.length >= 2) {
        const h = md[0];
        const cDate = imsw_col_(h, ['Date', 'date', 'date_received', 'timestamp']);
        const today = new Date(); today.setHours(0, 0, 0, 0);

        for (let i = 1; i < md.length; i++) {
          const v = cDate >= 0 ? md[i][cDate] : '';
          const d = (v instanceof Date) ? new Date(v) : new Date(v);
          if (isNaN(d.getTime())) continue;
          d.setHours(0, 0, 0, 0);
          if (d.getTime() === today.getTime()) metrics.inbound.receivedToday++;
        }
      }
    }

    // Inbound scheduled (Truck_Schedule) - count future/active rows if present
    const schedSh = ss.getSheetByName(tabs.TRUCK_SCHEDULE) || ss.getSheetByName('Truck_Schedule');
    if (schedSh) {
      const td = schedSh.getDataRange().getValues();
      if (td.length >= 2) {
        const h = td[0];
        const cDate = imsw_col_(h, ['Date', 'date', 'scheduled_date', 'eta']);
        const cStatus = imsw_col_(h, ['Status', 'status']);
        const today = new Date(); today.setHours(0, 0, 0, 0);

        for (let i = 1; i < td.length; i++) {
          const st = String(cStatus >= 0 ? td[i][cStatus] : '').toLowerCase();
          if (st && (st.includes('cancel') || st.includes('closed') || st.includes('complete'))) continue;

          const v = cDate >= 0 ? td[i][cDate] : '';
          const d = (v instanceof Date) ? new Date(v) : new Date(v);
          if (isNaN(d.getTime())) continue;
          d.setHours(0, 0, 0, 0);
          if (d.getTime() >= today.getTime()) metrics.inbound.scheduled++;
        }
      }
    }

    return { success: true, metrics };
  } catch (e) {
    Logger.log('getDashboardMetrics error: ' + e);
    return { success: false, message: e.toString(), metrics: null };
  }
}

// ============================================================================
// 3) CUSTOMER ORDERS LIST (called by IMS_Shell.getCustomerOrders(context))
// ============================================================================

function getCustomerOrdersForWebApp(context) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tabs = imsw_tabs_();
    const ordersSh = ss.getSheetByName(tabs.CUSTOMER_ORDERS) || ss.getSheetByName('Customer_Orders');
    const itemsSh = ss.getSheetByName(tabs.REQUESTED_ITEMS) || ss.getSheetByName('Requested_Items');

    if (!ordersSh) return { success: false, message: 'Customer_Orders sheet not found', orders: [] };

    const od = ordersSh.getDataRange().getValues();
    if (!od || od.length < 2) return { success: true, orders: [] };

    const oh = od[0];

    const cOrderId = imsw_col_(oh, ['Order_ID', 'order_id', 'order_number']);
    const cTask = imsw_col_(oh, ['Task_Number', 'task_number', 'task']);
    const cProject = imsw_col_(oh, ['Project', 'project']);
    const cNbd = imsw_col_(oh, ['NBD', 'nbd', 'need_by_date', 'need by date']);
    const cCompany = imsw_col_(oh, ['Company', 'company']);
    const cTitle = imsw_col_(oh, ['Order_Title', 'order_title', 'title']);
    const cDeliver = imsw_col_(oh, ['Deliver_To', 'deliver_to', 'delivery_address', 'deliver to']);
    const cStatus = imsw_col_(oh, ['Request_Status', 'request_status', 'status']);
    const cStock = imsw_col_(oh, ['Stock_Status', 'stock_status']);
    const cPick = imsw_col_(oh, ['Pick_Ticket_PDF', 'pick_ticket_pdf', 'pick_ticket']);
    const cToc = imsw_col_(oh, ['TOC_PDF', 'toc_pdf', 'toc']);
    const cPack = imsw_col_(oh, ['Packing_Lists', 'packing_lists', 'packing']);
    const cCreated = imsw_col_(oh, ['Created_TS', 'created_ts', 'timestamp', 'created']);

    // Aggregate items by order for quick stats
    const statsByOrder = {};
    if (itemsSh) {
      const id = itemsSh.getDataRange().getValues();
      if (id && id.length >= 2) {
        const ih = id[0];
        const iOrder = imsw_col_(ih, ['Order_ID', 'order_id', 'order_number', 'task_number']);
        const iQtyReq = imsw_col_(ih, ['Qty_Requested', 'qty_requested', 'qty']);
        const iQtyBack = imsw_col_(ih, ['Qty_Backordered', 'qty_backordered']);
        for (let i = 1; i < id.length; i++) {
          const ok = String(iOrder >= 0 ? id[i][iOrder] : '').trim();
          if (!ok) continue;
          if (!statsByOrder[ok]) statsByOrder[ok] = { itemLines: 0, qtyRequested: 0, qtyBackordered: 0 };
          statsByOrder[ok].itemLines++;
          statsByOrder[ok].qtyRequested += iQtyReq >= 0 ? imsw_safeNum_(id[i][iQtyReq]) : 0;
          statsByOrder[ok].qtyBackordered += iQtyBack >= 0 ? imsw_safeNum_(id[i][iQtyBack]) : 0;
        }
      }
    }

    const userLevel = String(context && context.accessLevel ? context.accessLevel : 'Standard').toLowerCase();
    const userCompany = String(context && context.company ? context.company : '').toLowerCase();

    const orders = [];
    for (let r = 1; r < od.length; r++) {
      const row = od[r];

      const orderId = String(cOrderId >= 0 ? row[cOrderId] : '').trim();
      if (!orderId) continue;

      const comp = String(cCompany >= 0 ? row[cCompany] : '').trim();
      if (userLevel.includes('standard') && userCompany) {
        if (String(comp).toLowerCase() !== userCompany) continue;
      }

      const stat = statsByOrder[orderId] || { itemLines: 0, qtyRequested: 0, qtyBackordered: 0 };

      orders.push({
        orderId,
        taskNumber: String(cTask >= 0 ? row[cTask] : '').trim(),
        project: String(cProject >= 0 ? row[cProject] : '').trim(),
        nbd: (cNbd >= 0 ? row[cNbd] : '') || '',
        company: comp,
        orderTitle: String(cTitle >= 0 ? row[cTitle] : '').trim(),
        deliverTo: String(cDeliver >= 0 ? row[cDeliver] : '').trim(),
        requestStatus: String(cStatus >= 0 ? row[cStatus] : '').trim(),
        stockStatus: String(cStock >= 0 ? row[cStock] : '').trim(),
        pickTicketUrl: cPick >= 0 ? row[cPick] : '',
        tocUrl: cToc >= 0 ? row[cToc] : '',
        packingListsUrl: cPack >= 0 ? row[cPack] : '',
        createdTs: cCreated >= 0 ? row[cCreated] : '',
        itemLines: stat.itemLines,
        qtyRequested: stat.qtyRequested,
        qtyBackordered: stat.qtyBackordered
      });
    }

    // Newest first if we can parse createdTs
    orders.sort((a, b) => {
      const da = new Date(a.createdTs || 0).getTime();
      const db = new Date(b.createdTs || 0).getTime();
      return (isNaN(db) ? 0 : db) - (isNaN(da) ? 0 : da);
    });

    return { success: true, orders };
  } catch (e) {
    Logger.log('getCustomerOrdersForWebApp error: ' + e);
    return { success: false, message: e.toString(), orders: [] };
  }
}

// ============================================================================
// 4) ORDER DETAILS (WebApp: open order)
// ============================================================================

function getOrderDetailsForWebApp(orderId, context) {
  try {
    if (!orderId) return { success: false, message: 'orderId required' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tabs = imsw_tabs_();
    const ordersSh = ss.getSheetByName(tabs.CUSTOMER_ORDERS) || ss.getSheetByName('Customer_Orders');
    const itemsSh = ss.getSheetByName(tabs.REQUESTED_ITEMS) || ss.getSheetByName('Requested_Items');
    if (!ordersSh) return { success: false, message: 'Customer_Orders sheet not found' };
    if (!itemsSh) return { success: false, message: 'Requested_Items sheet not found' };

    const od = ordersSh.getDataRange().getValues();
    const oh = od[0];

    const cOrderId = imsw_col_(oh, ['Order_ID', 'order_id', 'order_number']);
    const cCompany = imsw_col_(oh, ['Company', 'company']);

    const want = String(orderId).trim();
    let orderRow = null;

    for (let r = 1; r < od.length; r++) {
      const v = String(cOrderId >= 0 ? od[r][cOrderId] : '').trim();
      if (v === want) { orderRow = od[r]; break; }
    }
    if (!orderRow) return { success: false, message: 'Order not found: ' + want };

    // Enforce customer company isolation
    const userLevel = String(context && context.accessLevel ? context.accessLevel : 'Standard').toLowerCase();
    const userCompany = String(context && context.company ? context.company : '').toLowerCase();
    const orderCompany = String(cCompany >= 0 ? orderRow[cCompany] : '').toLowerCase().trim();

    if (userLevel.includes('standard') && userCompany && orderCompany && userCompany !== orderCompany) {
      return { success: false, message: 'Access denied for this order.' };
    }

    // Build order object from headers
    const order = {};
    for (let c = 0; c < oh.length; c++) order[String(oh[c] || ('COL_' + (c + 1)))] = orderRow[c];

    // Items
    const id = itemsSh.getDataRange().getValues();
    const ih = id[0];
    const iOrder = imsw_col_(ih, ['Order_ID', 'order_id', 'order_number', 'task_number']);
    const items = [];
    for (let i = 1; i < id.length; i++) {
      const ok = String(iOrder >= 0 ? id[i][iOrder] : '').trim();
      if (ok !== want) continue;

      const item = {};
      for (let c = 0; c < ih.length; c++) item[String(ih[c] || ('COL_' + (c + 1)))] = id[i][c];
      items.push(item);
    }

    return { success: true, orderId: want, order, items };
  } catch (e) {
    Logger.log('getOrderDetailsForWebApp error: ' + e);
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// 5) STOCK TOTALS LIST (WebApp: inventory table/search)
// ============================================================================

function getStockTotalsForWebApp(context, params) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tabs = imsw_tabs_();
    const sh = ss.getSheetByName(tabs.STOCK_TOTALS) || ss.getSheetByName('Stock_Totals');
    if (!sh) return { success: false, message: 'Stock_Totals sheet not found', rows: [], total: 0 };

    const data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, rows: [], total: 0, headers: data[0] || [] };

    const headers = data[0];

    const q = (params && params.q) ? String(params.q).toLowerCase().trim() : '';
    const offset = Math.max(0, Number(params && params.offset ? params.offset : 0) || 0);
    const limit = Math.min(500, Math.max(25, Number(params && params.limit ? params.limit : 200) || 200));

    const cSku = imsw_col_(headers, ['SKU', 'sku']);
    const cFbpn = imsw_col_(headers, ['FBPN', 'fbpn']);
    const cDesc = imsw_col_(headers, ['Description', 'description', 'desc']);

    // Optional: if you ever add a Company column here, we’ll honor it.
    const cCompany = imsw_col_(headers, ['Company', 'company']);

    const userLevel = String(context && context.accessLevel ? context.accessLevel : 'Standard').toLowerCase();
    const userCompany = String(context && context.company ? context.company : '').toLowerCase();

    const matches = [];
    for (let r = 1; r < data.length; r++) {
      const row = data[r];

      if (userLevel.includes('standard') && userCompany && cCompany >= 0) {
        const rc = String(row[cCompany] || '').toLowerCase().trim();
        if (rc && rc !== userCompany) continue;
      }

      if (q) {
        const sku = cSku >= 0 ? String(row[cSku] || '').toLowerCase() : '';
        const fbpn = cFbpn >= 0 ? String(row[cFbpn] || '').toLowerCase() : '';
        const desc = cDesc >= 0 ? String(row[cDesc] || '').toLowerCase() : '';
        if (!(sku.includes(q) || fbpn.includes(q) || desc.includes(q))) continue;
      }

      matches.push(row);
    }

    const total = matches.length;
    const page = matches.slice(offset, offset + limit);

    const rows = page.map(row => {
      const o = {};
      for (let c = 0; c < headers.length; c++) o[String(headers[c] || ('COL_' + (c + 1)))] = row[c];
      return o;
    });

    return { success: true, headers, rows, total, offset, limit };
  } catch (e) {
    Logger.log('getStockTotalsForWebApp error: ' + e);
    return { success: false, message: e.toString(), rows: [], total: 0 };
  }
}

// ============================================================================
// 6) CREATE ORDER (WebApp: customer creates order)
// ============================================================================

function createCustomerOrderForWebApp(orderData, context) {
  try {
    // Minimal permission gate
    const lvl = String(context && context.accessLevel ? context.accessLevel : 'Standard').toLowerCase();
    if (!lvl.includes('standard') && !lvl.includes('admin') && !lvl.includes('employee') && !lvl.includes('stake')) {
      return { success: false, message: 'Access denied.' };
    }

    // Force company for Standard customers (prevents spoofing)
    if (lvl.includes('standard') && context && context.company) {
      orderData = orderData || {};
      orderData.company = context.company;
    }

    if (typeof processCustomerOrder !== 'function') {
      return { success: false, message: 'processCustomerOrder() not available in this project.' };
    }

    // processCustomerOrder expects:
    // { taskNumber, company, project, orderTitle, deliveryLocation, nbdDate, name, phoneNumber, items:[{fbpn, qty, description, manufacturer}], fileData? }
    return processCustomerOrder(orderData);
  } catch (e) {
    Logger.log('createCustomerOrderForWebApp error: ' + e);
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// 7) DOC REGEN WRAPPER (WebApp: regenerate PICK/PACKING/TOC)
// ============================================================================

function regenerateOrderDocForWebApp(orderId, docType, context) {
  try {
    const perms = (context && context.permissions) ? context.permissions : imsw_buildPermissionsFromLevel_(context && context.accessLevel);
    if (!perms.canGenerateDocs) return { success: false, message: 'Access denied.' };

    if (typeof regenerateOrderDoc !== 'function') {
      return { success: false, message: 'regenerateOrderDoc() not available in this project.' };
    }
    return regenerateOrderDoc(orderId, docType);
  } catch (e) {
    Logger.log('regenerateOrderDocForWebApp error: ' + e);
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// 8) BOOTSTRAP (single call for UI init)
// ============================================================================

function getWebAppBootstrap(ssId) {
  try {
    const ctx = getUserContextForWebApp(ssId);
    if (!ctx || !ctx.authenticated) return { success: false, context: ctx };

    const dash = getDashboardMetrics();
    return {
      success: true,
      context: ctx,
      dashboard: dash && dash.success ? dash.metrics : null
    };
  } catch (e) {
    Logger.log('getWebAppBootstrap error: ' + e);
    return { success: false, message: e.toString() };
  }
}
