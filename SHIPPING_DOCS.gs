// ============================================================================
// SHIPPING_DOCS.GS
// Handles: Pick Tickets, Packing Lists, TOC, and Outbound Shipment Processing
// ============================================================================

// ─────────────────────────────────────────────────────────────────────────────
// 1. SPREADSHEET & BASIC HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function getSpreadsheetForShipping_() {
  if (typeof getSpreadsheet_ === 'function') {
    try { return getSpreadsheet_(); } catch (e) { Logger.log('getSpreadsheet_ failed: ' + e); }
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e2) { Logger.log('getActiveSpreadsheet failed: ' + e2); }

  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  throw new Error('getSpreadsheetForShipping_: No spreadsheet context available.');
}

function _col_(headers, names) {
  var h = (headers || []).map(x => String(x || '').trim().toLowerCase());
  for (var i = 0; i < names.length; i++) {
    var n = String(names[i]).trim().toLowerCase();
    var idx = h.indexOf(n);
    if (idx >= 0) return idx;
  }
  return -1;
}

function _norm_(v) {
  var s = String(v == null ? '' : v).trim();
  if (/^\d+\.0$/.test(s)) s = s.replace(/\.0$/, '');
  return s;
}

function _eq_(a, b) {
  return _norm_(a).toLowerCase() === _norm_(b).toLowerCase();
}

function _orderKey_(v) {
  return _norm_(v).toLowerCase().replace(/^order[-_]?/i, '').trim();
}

function _extractDriveIdFromUrl_(s) {
  const str = String(s || '').trim();
  if (!str) return '';
  if (/^[a-zA-Z0-9_-]{20,}$/.test(str)) return str;
  let m = str.match(/\/folders\/([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];
  m = str.match(/[-\w]{25,}/);
  if (m && m[0]) return m[0];
  return '';
}

function _safeNum_(v) {
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function _upper_(v) {
  return String(v || '').trim().toUpperCase();
}

// ─────────────────────────────────────────────────────────────────────────────
// 2. PLACEHOLDER REPLACEMENT (SUPPORTS {{Token}} AND {{(Token)}})
// ─────────────────────────────────────────────────────────────────────────────

function replacePlaceholders(docOrElement, placeholders) {
  let targets = [];
  if (docOrElement && typeof docOrElement.getBody === 'function') {
    targets.push(docOrElement.getBody());
    try { if (docOrElement.getHeader && docOrElement.getHeader()) targets.push(docOrElement.getHeader()); } catch (e) {}
    try { if (docOrElement.getFooter && docOrElement.getFooter()) targets.push(docOrElement.getFooter()); } catch (e2) {}
  } else {
    targets.push(docOrElement);
  }

  const esc = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  for (let i = 0; i < targets.length; i++) {
    const target = targets[i];
    if (!target) continue;

    for (let key in placeholders) {
      const val = String(placeholders[key] ?? '');
      const inner = String(key).replace(/^\{\{|\}\}$/g, '').trim();
      const pattern = "\\{\\{\\s*\\(?\\s*" + esc(inner) + "\\s*\\)?\\s*\\}\\}";
      try {
        target.replaceText(pattern, val);
      } catch (e) {
        Logger.log("replacePlaceholders failed for " + key + ": " + e);
      }
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. TASK/ORDER LOOKUPS
// ─────────────────────────────────────────────────────────────────────────────

function getTaskNumbers() {
  var ss = getSpreadsheetForShipping_();
  var sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet) return [];

  var v = sheet.getDataRange().getValues();
  if (v.length < 2) return [];

  var headers = v[0];
  var cTask = _col_(headers, ['task_number', 'task number', 'task#', 'task #', 'task']);
  var cOrder = _col_(headers, ['order_id', 'order id', 'order_number', 'order number', 'order']);
  var cStatus = _col_(headers, ['request_status', 'status', 'request status']);

  var set = {};
  for (var r = 1; r < v.length; r++) {
    if (cStatus >= 0) {
      var status = _upper_(v[r][cStatus]);
      if (status === 'DELIVERED' || status === 'CANCELLED') continue;
    }
    var key = _norm_(cTask >= 0 ? v[r][cTask] : '') || _norm_(cOrder >= 0 ? v[r][cOrder] : '');
    if (key && key.toLowerCase() !== 'undefined') set[key] = true;
  }

  return Object.keys(set).sort((a, b) => {
    var na = parseInt(a, 10), nb = parseInt(b, 10);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.localeCompare(b);
  });
}

function getOrderByTaskNumber(taskNumber) {
  var ss = getSpreadsheetForShipping_();
  var ordersSh = ss.getSheetByName('Customer_Orders');
  var itemsSh = ss.getSheetByName('Requested_Items');

  var key = _norm_(taskNumber);
  if (!key || !ordersSh || !itemsSh) return null;

  var od = ordersSh.getDataRange().getValues();
  var oh = od[0];
  var cTask = _col_(oh, ['task_number', 'task']);
  var cOrder = _col_(oh, ['order_id', 'order_number']);
  var cProj = _col_(oh, ['project']);
  var cComp = _col_(oh, ['company']);
  var cTitle = _col_(oh, ['order_title', 'title']);
  var cDeliver = _col_(oh, ['deliver_to', 'delivery_address', 'address', 'deliver to']);
  var cName = _col_(oh, ['name', 'contact_name', 'site_contact_name', 'contact name']);
  var cPhone = _col_(oh, ['phone_number', 'phone', 'contact_phone', 'phone number']);

  var orderRow = null;
  for (var r = 1; r < od.length; r++) {
    var row = od[r];
    var vTask = (cTask >= 0) ? row[cTask] : '';
    var vOrder = (cOrder >= 0) ? row[cOrder] : '';
    if (_eq_(vTask, key) || _eq_(vOrder, key)) { orderRow = row; break; }
  }
  if (!orderRow) return null;

  var orderId = _norm_(cOrder >= 0 ? orderRow[cOrder] : '') || _norm_(cTask >= 0 ? orderRow[cTask] : '') || key;

  var id = itemsSh.getDataRange().getValues();
  var ih = id[0];
  var iOrder = _col_(ih, ['order_id', 'order_number', 'task_number']);
  var iFbpn = _col_(ih, ['fbpn']);
  var iDesc = _col_(ih, ['description', 'desc']);
  var iQty = _col_(ih, ['qty_requested', 'qty']);

  var items = [];
  var matchKeys = [key, orderId, String(Math.trunc(Number(key))), String(Math.trunc(Number(orderId)))];

  for (var j = 1; j < id.length; j++) {
    var row2 = id[j];
    var ok = _norm_(iOrder >= 0 ? row2[iOrder] : '');
    if (!matchKeys.includes(ok)) continue;

    var fbpn = String(iFbpn >= 0 ? row2[iFbpn] : '').trim();
    if (!fbpn) continue;

    items.push({
      fbpn: fbpn,
      description: String(iDesc >= 0 ? row2[iDesc] : '').trim(),
      qtyRequested: Number(iQty >= 0 ? row2[iQty] : 0) || 0
    });
  }

  var combined = {};
  items.forEach(item => {
    var k = item.fbpn;
    if (!combined[k]) combined[k] = item;
    else combined[k].qtyRequested += item.qtyRequested;
  });

  return {
    orderId: orderId,
    orderTitle: (cTitle >= 0 ? orderRow[cTitle] : '') || '',
    company: (cComp >= 0 ? orderRow[cComp] : '') || '',
    project: (cProj >= 0 ? orderRow[cProj] : '') || '',
    deliverTo: (cDeliver >= 0 ? orderRow[cDeliver] : '') || '',
    name: (cName >= 0 ? orderRow[cName] : '') || '',
    phoneNumber: (cPhone >= 0 ? orderRow[cPhone] : '') || '',
    items: Object.values(combined)
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. ORIGINAL ORDER FOLDER RESOLUTION
// ─────────────────────────────────────────────────────────────────────────────

function _getOrderFolderFromCustomerOrders_(ss, orderNumber) {
  const sh = ss.getSheetByName('Customer_Orders');
  if (!sh) return null;
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const h = values[0];
  const cOrder = _col_(h, ['order_id', 'order_number', 'order id', 'order number']);
  if (cOrder < 0) return null;

  const cFolderA = _col_(h, ['order_folder_id', 'order folder id', 'folder_id', 'folder id']);
  const cFolderB = _col_(h, ['order_folder', 'order folder', 'folder', 'folder link', 'order_folder_url', 'order folder url', 'folder_url', 'folder url']);

  const want = _orderKey_(orderNumber);

  for (let r = 1; r < values.length; r++) {
    if (_orderKey_(values[r][cOrder]) !== want) continue;

    const rawA = (cFolderA >= 0) ? values[r][cFolderA] : '';
    const rawB = (cFolderB >= 0) ? values[r][cFolderB] : '';
    const id = _extractDriveIdFromUrl_(rawA) || _extractDriveIdFromUrl_(rawB);
    if (!id) return null;

    try { return DriveApp.getFolderById(id); }
    catch (e) {
      Logger.log('_getOrderFolderFromCustomerOrders_: bad folder for ' + orderNumber + ': ' + e);
      return null;
    }
  }
  return null;
}

function getOrderFolderForShipment(parentFolder, orderNumber, ss) {
  try {
    if (ss) {
      const f = _getOrderFolderFromCustomerOrders_(ss, orderNumber);
      if (f) return f;
    }
  } catch (e) {
    Logger.log('getOrderFolderForShipment Customer_Orders lookup failed: ' + e);
  }

  const folders = parentFolder.getFoldersByName(orderNumber);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(orderNumber);
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. CORE SHIPMENT ORCHESTRATION
// ─────────────────────────────────────────────────────────────────────────────

function processPackingTOCAndShipment(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (typeof getIMSConfig !== 'function') throw new Error('IMS Config missing.');
    const ss = getSpreadsheetForShipping_();

    const results = {
      success: true,
      tocUrl: null,
      packingListsUrl: null,
      pickLogUpdates: 0,
      inventoryUpdates: 0,
      outboundLogEntries: 0,
      backordersCreated: 0
    };

    data.skids = processSkidsWithCombinedFBPNs(data.skids || []);

    const tocResult = generateTOC(data);
    if (!tocResult.success) throw new Error('TOC Generation Failed: ' + (tocResult.message || 'Unknown Error'));
    results.tocUrl = tocResult.pdfUrl;

    const plResult = generatePackingLists(data);
    if (!plResult.success) throw new Error('Packing List Generation Failed: ' + (plResult.message || 'Unknown Error'));
    results.packingListsUrl = plResult.pdfUrl;

    const timestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail();

    // Summarize shipped qty by fbpn from skid payload
    const shippedQtyByFbpn = {};
    (data.skids || []).forEach(skid => {
      (skid.items || []).forEach(item => {
        const qty = Number(item.qtyOnSkid || item.qty || 0);
        if (qty > 0) shippedQtyByFbpn[item.fbpn] = (shippedQtyByFbpn[item.fbpn] || 0) + qty;
      });
    });

    // Update Pick_Log + reduce inventory + write OutboundLog
    const pickLogSheet = ss.getSheetByName('Pick_Log');
    if (pickLogSheet) {
      const pickData = pickLogSheet.getDataRange().getValues();
      const h = pickData[0];

      const cOrder = _col_(h, ['order_number', 'order_id', 'order #', 'order id']);
      const cFbpn = _col_(h, ['fbpn', 'item', 'part number']);
      const cQtyPick = _col_(h, ['qty_to_pick', 'quantity to pick', 'qty requested']);
      const cQtyPicked = _col_(h, ['qty_picked', 'quantity picked', 'picked qty']);
      const cBin = _col_(h, ['bin_code', 'bin', 'location']);
      const cStatus = _col_(h, ['status']);
      const cDate = _col_(h, ['shipped_date', 'date shipped']);
      const cUser = _col_(h, ['picked_by', 'user']);
      const cSku = _col_(h, ['sku']);

      let remainingToAssign = Object.assign({}, shippedQtyByFbpn);

      if (cOrder >= 0 && cFbpn >= 0) {
        for (let i = 1; i < pickData.length; i++) {
          const rowOrder = String(pickData[i][cOrder] || '');
          if (_orderKey_(rowOrder) !== _orderKey_(data.orderNumber || '')) continue;

          const status = cStatus >= 0 ? _upper_(pickData[i][cStatus]) : '';
          if (['SHIPPED', 'CANCELLED', 'COMPLETE', 'DELIVERED'].includes(status)) continue;

          const fbpn = String(pickData[i][cFbpn] || '').trim();
          const qtyToPick = cQtyPick >= 0 ? _safeNum_(pickData[i][cQtyPick]) : 0;

          const available = remainingToAssign[fbpn] || 0;
          if (available <= 0) continue;

          const qtyShipped = Math.min(qtyToPick, available);
          remainingToAssign[fbpn] -= qtyShipped;

          const rowIdx = i + 1;
          if (cQtyPicked >= 0) pickLogSheet.getRange(rowIdx, cQtyPicked + 1).setValue(qtyShipped);
          if (cStatus >= 0) pickLogSheet.getRange(rowIdx, cStatus + 1).setValue('SHIPPED');
          if (cDate >= 0) pickLogSheet.getRange(rowIdx, cDate + 1).setValue(timestamp);
          if (cUser >= 0) pickLogSheet.getRange(rowIdx, cUser + 1).setValue(userEmail);
          results.pickLogUpdates++;

          const binCode = cBin >= 0 ? String(pickData[i][cBin] || '').trim() : '';
          const rowSku = cSku >= 0 ? String(pickData[i][cSku] || '').trim() : '';

          if (qtyShipped > 0 && binCode) {
            const sourceSheetName = getSourceFromBinCode(ss, binCode, fbpn);
            reduceInventoryForShipment(ss, sourceSheetName, binCode, fbpn, qtyShipped);
            results.inventoryUpdates++;

            updateAllocationStatus(ss, data.orderNumber, fbpn, 'SHIPPED');

            const currentSku = rowSku || getCurrentBinSku(ss, sourceSheetName, binCode, fbpn);

            writeOutboundLogEntry(ss, {
              date: data.shipDate,
              orderNumber: data.orderNumber,
              taskNumber: data.taskNumber,
              company: data.company,
              project: data.project,
              fbpn: fbpn,
              qty: qtyShipped,
              binCode: binCode,
              sku: currentSku
            });
            results.outboundLogEntries++;
          }
        }
      }
    }

    updateCustomerOrderShipment(ss, data.orderNumber, {
      status: 'Delivered',
      shipDate: data.shipDate,
      carrier: data.carrier,
      trackingNumber: data.trackingNumber,
      bolNumber: data.bolNumber
    });

    return results;

  } catch (err) {
    Logger.log('Error in processPackingTOCAndShipment: ' + err);
    return { success: false, message: String(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function processPackingTOC_DocsOnly(formData) {
  if (!formData) throw new Error('Missing formData');

  formData.skids = processSkidsWithCombinedFBPNs(formData.skids || []);

  const tocResult = generateTOC(formData);
  if (!tocResult.success) throw new Error(tocResult.message || 'TOC Generation Failed');

  const plResult = generatePackingLists(formData);
  if (!plResult.success) throw new Error(plResult.message || 'Packing List Generation Failed');

  return {
    success: true,
    message: 'Documents rebuilt (docs-only).',
    tocUrl: tocResult.pdfUrl,
    packingListsUrl: plResult.pdfUrl
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// 6. DOC GENERATION: TOC + PACKING LISTS
// ─────────────────────────────────────────────────────────────────────────────

function generateTOC(params) {
  try {
    const CONFIG = getIMSConfig();
    const templateId = CONFIG.TOC_TEMPLATE_ID;
    if (!templateId) throw new Error("TOC Template ID is missing in IMS_Config.");

    const ss = getSpreadsheetForShipping_();
    const rootId = CONFIG.TOC_PACKING_OUTPUT_FOLDER_ID;
    if (!rootId) throw new Error("Output Folder ID is missing in IMS_Config.");

    const rootFolder = DriveApp.getFolderById(rootId);
    const folder = getOrderFolderForShipment(rootFolder, params.orderNumber, ss);

    const templateFile = DriveApp.getFileById(templateId);
    const newFile = templateFile.makeCopy('TOC_' + params.orderNumber + '_' + new Date().getTime(), folder);
    const doc = DocumentApp.openById(newFile.getId());

    replacePlaceholders(doc, {
      'Order_ID': params.orderNumber || '',
      'Order_Number': params.orderNumber || '',
      'Order_Title': params.orderTitle || '',
      'Task_Number': params.taskNumber || '',
      'Total_Skids': params.totalSkids || '',
      'Company': params.company || '',
      'Project': params.project || '',
      'Delivery_Address': params.deliverTo || '',
      'Deliver_To': params.deliverTo || '',
      'Site_Contact_Name': params.name || '',
      'Name': params.name || '',
      'Site_Contact_Phone': params.phoneNumber || '',
      'Phone_Number': params.phoneNumber || '',
      'Carrier': params.carrier || '',
      'Date': params.shipDate || new Date().toLocaleDateString()
    });

    // Fill item table (find header row with FBPN/Description, then use next row as template)
    const body = doc.getBody();
    const tables = body.getTables();
    let targetTable = null;
    let templateRowIndex = -1;

    for (let i = 0; i < tables.length; i++) {
      const t = tables[i];
      if (t.getNumRows() < 2) continue;
      for (let r = 0; r < Math.min(3, t.getNumRows()); r++) {
        const text = String(t.getRow(r).getText() || '').toLowerCase();
        if (text.includes('fbpn') && (text.includes('description') || text.includes('desc'))) {
          targetTable = t;
          templateRowIndex = Math.min(r + 1, t.getNumRows() - 1);
          break;
        }
      }
      if (targetTable) break;
    }

    if (targetTable && templateRowIndex > -1) {
      const templateRow = targetTable.getRow(templateRowIndex).copy();
      while (targetTable.getNumRows() > templateRowIndex) targetTable.removeRow(templateRowIndex);

      const allItems = [];
      (params.skids || []).forEach(skid => (skid.items || []).forEach(item => allItems.push(item)));
      const combined = combinePackingListItems(allItems);

      combined.forEach(item => {
        const row = templateRow.copy();
        if (row.getNumCells() > 0) row.getCell(0).setText(item.fbpn || '');
        if (row.getNumCells() > 1) row.getCell(1).setText(item.description || '');
        if (row.getNumCells() > 2) row.getCell(2).setText(String(item.qtyRequested || 0));
        if (row.getNumCells() > 3) row.getCell(3).setText(String(item.qty || item.qtyOnSkid || 0));
        targetTable.appendTableRow(row);
      });
    }

    doc.saveAndClose();

    const pdfBlob = doc.getAs(MimeType.PDF);
    const pdfFile = folder.createFile(pdfBlob).setName('TOC_' + params.orderNumber + '.pdf');
    newFile.setTrashed(true);

    linkPDFToCustomerOrder(ss, params.orderNumber, pdfFile.getUrl(), 'TOC_PDF', 'TOC');
    return { success: true, pdfUrl: pdfFile.getUrl() };

  } catch (e) {
    Logger.log("generateTOC Error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function generatePackingLists(params) {
  try {
    const CONFIG = getIMSConfig();
    const templateId = CONFIG.PACKING_LIST_TEMPLATE_ID;
    if (!templateId) throw new Error("Packing List Template ID is missing in IMS_Config.");

    const ss = getSpreadsheetForShipping_();
    const rootId = CONFIG.TOC_PACKING_OUTPUT_FOLDER_ID;
    if (!rootId) throw new Error("Output Folder ID is missing in IMS_Config.");

    const rootFolder = DriveApp.getFolderById(rootId);
    const folder = getOrderFolderForShipment(rootFolder, params.orderNumber, ss);


    // Only generate packing lists for skids that contain multiple FBPNs
    const skidsToPrint = (params.skids || []).filter(skid => {
      const items = (skid && skid.items) ? skid.items : [];
      const uniq = new Set(items.map(it => String(it.fbpn || '').trim()).filter(Boolean));
      return uniq.size > 1;
    });

    if (skidsToPrint.length === 0) {
      return {
        success: true,
        pdfUrl: '',
        message: 'No packing lists required (all skids contain a single FBPN).'
      };
    }

    const plFolder = folder.createFolder('PackingLists_' + params.orderNumber);

    (params.skids || []).forEach(skid => {
      const templateFile = DriveApp.getFileById(templateId);
      const newFile = templateFile.makeCopy('PL_Skid' + skid.skidNumber + '_' + params.orderNumber, plFolder);
      const doc = DocumentApp.openById(newFile.getId());

      replacePlaceholders(doc, {
        'Order_ID': params.orderNumber,
        'Order_Number': params.orderNumber,
        'Order_Title': params.orderTitle || '',
        'Task_Number': params.taskNumber || '',
        'Skid_Number': String(skid.skidNumber),
        'Skid_Sequence': String(skid.skidNumber),
        'Total_Skids': params.totalSkids || '',
        'Project': params.project || '',
        'Company': params.company || '',
        'Delivery_Address': params.deliverTo || '',
        'Deliver_To': params.deliverTo || '',
        'Name': params.name || '',
        'Phone_Number': params.phoneNumber || '',
        'Date': params.shipDate || new Date().toLocaleDateString()
      });

      const body = doc.getBody();
      const tables = body.getTables();

      let targetTable = null;
      let templateRowIndex = -1;

      for (let i = 0; i < tables.length; i++) {
        const t = tables[i];
        if (t.getNumRows() < 2) continue;
        for (let r = 0; r < Math.min(3, t.getNumRows()); r++) {
          const text = String(t.getRow(r).getText() || '').toLowerCase();
          if (text.includes('fbpn') && (text.includes('description') || text.includes('desc'))) {
            targetTable = t;
            templateRowIndex = Math.min(r + 1, t.getNumRows() - 1);
            break;
          }
        }
        if (targetTable) break;
      }

      if (targetTable && templateRowIndex > -1) {
        const templateRow = targetTable.getRow(templateRowIndex).copy();
        while (targetTable.getNumRows() > templateRowIndex) targetTable.removeRow(templateRowIndex);

        (skid.items || []).forEach(item => {
          const row = templateRow.copy();
          if (row.getNumCells() > 0) row.getCell(0).setText(item.fbpn || '');
          if (row.getNumCells() > 1) row.getCell(1).setText(item.description || '');
          if (row.getNumCells() > 2) row.getCell(2).setText(String(item.qtyRequested || 0));
          if (row.getNumCells() > 3) row.getCell(3).setText(String(item.qty || item.qtyOnSkid || 0));
          targetTable.appendTableRow(row);
        });
      }

      doc.saveAndClose();

      const pdfBlob = doc.getAs(MimeType.PDF);
      plFolder.createFile(pdfBlob).setName(`Skid${skid.skidNumber}_PackingList.pdf`);
      newFile.setTrashed(true);
    });

    return { success: true, pdfUrl: plFolder.getUrl() };

  } catch (e) {
    Logger.log("generatePackingLists Error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// 7. PICK TICKET (FILL TABLE ROWS)
// ─────────────────────────────────────────────────────────────────────────────

function generatePickTicket(formData) {
  const CONFIG = getIMSConfig();
  const TEMPLATE_ID = CONFIG.PICK_TICKET_TEMPLATE_ID || (typeof TEMPLATES !== 'undefined' ? TEMPLATES.Pickticket_Template : null);
  const ROOT_ID = CONFIG.TOC_PACKING_OUTPUT_FOLDER_ID;

  if (!TEMPLATE_ID) throw new Error("Pick Ticket Template ID missing in Config.");
  if (!ROOT_ID) throw new Error("TOC/Packing Output Folder ID missing in Config.");
  if (!formData || !formData.orderNumber) throw new Error("Missing orderNumber.");

  const ss = getSpreadsheetForShipping_();
  const rootFolder = DriveApp.getFolderById(ROOT_ID);
  const folder = getOrderFolderForShipment(rootFolder, formData.orderNumber, ss);

  // Prefer Pick_Log (so bins/push/skid come from real pick lines)
  const pickSheet = ss.getSheetByName('Pick_Log');
  if (!pickSheet) throw new Error("Pick_Log sheet not found.");

  const pickData = pickSheet.getDataRange().getValues();
  if (pickData.length < 2) throw new Error("Pick_Log empty.");

  const h = pickData[0];
  const cOrder = _col_(h, ['order_number', 'order_id', 'order #', 'order id']);
  const cFbpn  = _col_(h, ['fbpn']);
  const cDesc  = _col_(h, ['description']);
  const cReq   = _col_(h, ['qty_requested']);
  const cPick  = _col_(h, ['qty_to_pick', 'qty pick', 'qty to pick']);
  const cBin   = _col_(h, ['bin_code', 'bin', 'location']);
  const cPush  = _col_(h, ['push', 'push_number', 'push #', 'push number']);
  const cSkid  = _col_(h, ['skid', 'skid_number', 'skid number']);
  const cStatus = _col_(h, ['status']);

  if (cOrder < 0 || cFbpn < 0 || cPick < 0 || cBin < 0) {
    throw new Error("Pick_Log missing required headers (need Order + FBPN + Qty_To_Pick + Bin_Code).");
  }

  const wantKey = _orderKey_(formData.orderNumber);
  const rows = [];

  for (let r = 1; r < pickData.length; r++) {
    if (_orderKey_(pickData[r][cOrder]) !== wantKey) continue;

    const st = (cStatus >= 0) ? _upper_(pickData[r][cStatus]) : '';
    if (st === 'CANCELLED') continue;

    const qtyToPick = _safeNum_(pickData[r][cPick]);
    if (qtyToPick <= 0) continue;

    rows.push({
      fbpn: String(pickData[r][cFbpn] || '').trim(),
      description: cDesc >= 0 ? String(pickData[r][cDesc] || '') : '',
      qtyRequested: cReq >= 0 ? (Number(pickData[r][cReq]) || '') : '',
      qtyToPick: qtyToPick,
      pushNumber: cPush >= 0 ? String(pickData[r][cPush] || '') : '',
      binCode: String(pickData[r][cBin] || '').trim(),
      skidNumber: cSkid >= 0 ? String(pickData[r][cSkid] || '') : ''
    });
  }

  // Fallback: if your Pick_Log isn’t populated yet, use formData.items (no bins)
  if (rows.length === 0 && Array.isArray(formData.items) && formData.items.length) {
    formData.items.forEach(it => {
      const qty = _safeNum_(it.qtyToPick || it.qtyRequested || it.qty || 0);
      if (qty <= 0) return;
      rows.push({
        fbpn: String(it.fbpn || '').trim(),
        description: String(it.description || ''),
        qtyRequested: _safeNum_(it.qtyRequested || ''),
        qtyToPick: qty,
        pushNumber: '',
        binCode: '',
        skidNumber: ''
      });
    });
  }

  if (rows.length === 0) return { success: false, message: "No pickable lines found for this order." };

  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const newFile = templateFile.makeCopy('Pick_' + formData.orderNumber + '_' + new Date().getTime(), folder);
  const doc = DocumentApp.openById(newFile.getId());

  replacePlaceholders(doc, {
    'Order_ID': formData.orderNumber,
    'Order_Number': formData.orderNumber,
    'Order_Title': formData.orderTitle || '',
    'Task_Number': formData.taskNumber || '',
    'Project': formData.project || '',
    'Company': formData.company || '',
    'Date': formData.date || new Date().toLocaleDateString()
  });

  const body = doc.getBody();
  const tables = body.getTables();
  if (!tables || tables.length === 0) throw new Error("No table found in Pick Ticket template.");

  // Find the row that contains FBPN placeholder in ANY table
  let table = null;
  let templateRowIndex = -1;

  for (let t = 0; t < tables.length; t++) {
    const candidate = tables[t];
    for (let rr = 0; rr < candidate.getNumRows(); rr++) {
      const txt = String(candidate.getRow(rr).getText() || '');
      if (txt.includes('{{FBPN') || txt.includes('{{(FBPN')) {
        table = candidate;
        templateRowIndex = rr;
        break;
      }
    }
    if (table) break;
  }

  if (!table) {
    // fallback: use table that has "FBPN" header and take next row as template
    for (let t = 0; t < tables.length; t++) {
      const candidate = tables[t];
      for (let rr = 0; rr < Math.min(3, candidate.getNumRows()); rr++) {
        const txt = String(candidate.getRow(rr).getText() || '').toLowerCase();
        if (txt.includes('fbpn') && (txt.includes('qty') || txt.includes('description') || txt.includes('desc'))) {
          table = candidate;
          templateRowIndex = Math.min(rr + 1, candidate.getNumRows() - 1);
          break;
        }
      }
      if (table) break;
    }
  }

  if (!table) {
    table = tables[0];
    templateRowIndex = Math.min(1, table.getNumRows() - 1);
  }



  // Sort by Bin_Code for an optimized pick path (then FBPN)
  // Desired: sort by leading letter(s), then by the number AFTER the dot (e.g., A1.1 -> A then 1),
  // ignoring the first number before the dot for primary ordering.
  rows.sort((a, b) => {
    const aBin = String(a.binCode || '').trim().toUpperCase();
    const bBin = String(b.binCode || '').trim().toUpperCase();
    if (!aBin && bBin) return 1;
    if (aBin && !bBin) return -1;

    const parseBin = (code) => {
      const s = String(code || '').trim().toUpperCase();
      // Examples: A1.1, A12.3, OF1.2, REEL.A.12 (fallback handles weird formats)
      // Primary: leading letters (A, OF, REEL), secondary: number after last dot, tertiary: first number in string.
      const lettersMatch = s.match(/^([A-Z]+)/);
      const letters = lettersMatch ? lettersMatch[1] : '';
      const afterDotMatch = s.match(/\.(\d+)(?!.*\.(\d+))/); // last .number
      const afterDot = afterDotMatch ? parseInt(afterDotMatch[1], 10) : Number.POSITIVE_INFINITY;
      const firstNumMatch = s.match(/(\d+)/);
      const firstNum = firstNumMatch ? parseInt(firstNumMatch[1], 10) : Number.POSITIVE_INFINITY;
      return { letters, afterDot, firstNum, raw: s };
    };

    const ka = parseBin(aBin);
    const kb = parseBin(bBin);

    const cLetters = ka.letters.localeCompare(kb.letters, undefined, { sensitivity: 'base' });
    if (cLetters !== 0) return cLetters;

    if (ka.afterDot !== kb.afterDot) return ka.afterDot - kb.afterDot;

    // Tie-break: then sort by the first number (the one we "ignore" for primary ordering)
    if (ka.firstNum !== kb.firstNum) return ka.firstNum - kb.firstNum;

    // Final tie-break: full bin string
    const cRaw = ka.raw.localeCompare(kb.raw, undefined, { numeric: true, sensitivity: 'base' });
    if (cRaw !== 0) return cRaw;

    // Then FBPN
    const aFbpn = String(a.fbpn || '').trim().toUpperCase();
    const bFbpn = String(b.fbpn || '').trim().toUpperCase();
    return aFbpn.localeCompare(bFbpn, undefined, { numeric: true, sensitivity: 'base' });
  });
  const proto = table.getRow(templateRowIndex).copy();
  while (table.getNumRows() > templateRowIndex) table.removeRow(templateRowIndex);

  rows.forEach(item => {
    const row = proto.copy();

    const setCell = (i, v) => {
      if (row.getNumCells() > i) row.getCell(i).setText(String(v == null ? '' : v));
    };

    // Common 7-col layout: FBPN | Desc | Qty Req | Qty Pick | Push | Bin | Skid
    setCell(0, item.fbpn);
    setCell(1, item.description);
    setCell(2, item.qtyRequested);
    setCell(3, item.qtyToPick);
    setCell(4, item.pushNumber);
    setCell(5, item.binCode);
    setCell(6, item.skidNumber);

    table.appendTableRow(row);
  });

  doc.saveAndClose();

  const pdf = folder.createFile(doc.getAs(MimeType.PDF)).setName('PickTicket_' + formData.orderNumber + '.pdf');
  newFile.setTrashed(true);

  return { success: true, pdfUrl: pdf.getUrl(), url: pdf.getUrl(), name: pdf.getName() };
}

// ─────────────────────────────────────────────────────────────────────────────
// 8. ITEM COMBINERS (TOC / PACKING LISTS)
// ─────────────────────────────────────────────────────────────────────────────

function combinePackingListItems(items) {
  const map = new Map();
  (items || []).forEach(item => {
    const fbpn = String(item.fbpn || '').trim();
    if (!fbpn) return;
    const desc = String(item.description || '');
    const key = fbpn + '|' + desc;
    const qty = parseInt(item.qtyOnSkid || item.qty || item.quantity || 0, 10) || 0;
    const req = parseInt(item.qtyRequested || 0, 10) || 0;

    if (map.has(key)) {
      const ex = map.get(key);
      ex.qty += qty;
      ex.qtyOnSkid = (ex.qtyOnSkid || 0) + qty;
      ex.qtyRequested = Math.max(ex.qtyRequested || 0, req);
    } else {
      map.set(key, { fbpn, description: desc, qty: qty, qtyOnSkid: qty, qtyRequested: req });
    }
  });
  return Array.from(map.values());
}

function processSkidsWithCombinedFBPNs(skids) {
  return (skids || []).map((skid, idx) => {
    const combined = combinePackingListItems(skid.items || []);
    return { skidNumber: skid.skidNumber || (idx + 1), items: combined };
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// 9. INVENTORY + LOGGING HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function linkPDFToCustomerOrder(ss, orderId, url, colName, text) {
  const sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const h = data[0];

  const cLink = h.indexOf(colName);
  const cOrder = _col_(h, ['order_id', 'order_number']);
  if (cLink < 0 || cOrder < 0) return;

  for (let i = 1; i < data.length; i++) {
    if (_orderKey_(data[i][cOrder]) === _orderKey_(orderId)) {
      sheet.getRange(i + 1, cLink + 1).setFormula(`=HYPERLINK("${url}", "${text}")`);
      break;
    }
  }
}

function getSourceFromBinCode(ss, binCode, fbpn) {
  const tabs = ['Bin_Stock', 'Floor_Stock_Levels', 'Inbound_Staging'];
  for (const tabName of tabs) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) continue;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;
    const h = data[0];

    const cBin = h.indexOf('Bin_Code');
    const cFbpn = h.indexOf('FBPN');
    if (cBin < 0 || cFbpn < 0) continue;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][cBin] || '').trim() === String(binCode || '').trim() &&
          String(data[i][cFbpn] || '').trim() === String(fbpn || '').trim()) {
        return tabName;
      }
    }
  }
  return 'Bin_Stock';
}

function reduceInventoryForShipment(ss, tabName, binCode, fbpn, qty) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const h = data[0];
  const cBin = h.indexOf('Bin_Code');
  const cFbpn = h.indexOf('FBPN');
  const cQty = h.indexOf('Current_Quantity');

  if (cBin < 0 || cFbpn < 0 || cQty < 0) return;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cBin] || '').trim() === String(binCode || '').trim() &&
        String(data[i][cFbpn] || '').trim() === String(fbpn || '').trim()) {
      const current = _safeNum_(data[i][cQty]);
      const newVal = Math.max(0, current - _safeNum_(qty));
      sheet.getRange(i + 1, cQty + 1).setValue(newVal);
      break;
    }
  }
}

function getCurrentBinSku(ss, source, binCode, fbpn) {
  const sheet = ss.getSheetByName(source);
  if (!sheet) return '';
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return '';

  const h = data[0];
  const binCodeCol = h.indexOf('Bin_Code');
  const fbpnCol = h.indexOf('FBPN');
  const skuCol = h.indexOf('SKU');
  if (binCodeCol === -1 || fbpnCol === -1 || skuCol === -1) return '';

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][binCodeCol] || '').trim() === String(binCode || '').trim() &&
        String(data[i][fbpnCol] || '').trim() === String(fbpn || '').trim()) {
      return data[i][skuCol] || '';
    }
  }
  return '';
}

function writeOutboundLogEntry(ss, data) {
  const sheet = ss.getSheetByName('OutboundLog');
  if (!sheet) return;

  // Try to infer Manufacturer from Stock_Totals if not provided
  let mfr = data.manufacturer || '';
  if (!mfr) {
    const stSheet = ss.getSheetByName('Stock_Totals');
    if (stSheet) {
      const vals = stSheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][1] || '').trim() === String(data.fbpn || '').trim()) {
          mfr = vals[i][3];
          break;
        }
      }
    }
  }

  sheet.appendRow([
    data.date || new Date(),
    data.orderNumber,
    data.taskNumber,
    'OUTBOUND',
    'Warehouse',
    data.company,
    data.project,
    data.fbpn,
    mfr,
    data.qty,
    data.binCode,
    data.skidId || '',
    data.sku || ''
  ]);
}

function updateAllocationStatus(ss, orderNumber, fbpn, newStatus) {
  const sheet = ss.getSheetByName('Allocation_Log');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const h = data[0];
  const cOrder = _col_(h, ['order_id', 'order number', 'order_number']);
  const cFbpn = _col_(h, ['fbpn']);
  const cStatus = _col_(h, ['allocation_status', 'status']);

  if (cOrder < 0 || cFbpn < 0 || cStatus < 0) return;

  for (let i = 1; i < data.length; i++) {
    const rowOrder = data[i][cOrder];
    const matchOrder = (_orderKey_(rowOrder) === _orderKey_(orderNumber)) ||
      (_eq_(rowOrder, orderNumber)) ||
      (String(rowOrder || '').includes(String(orderNumber || '')) || String(orderNumber || '').includes(String(rowOrder || '')));

    if (!matchOrder) continue;

    const rowFbpn = data[i][cFbpn];
    if (fbpn && !_eq_(rowFbpn, fbpn)) continue;

    const currentStatus = _upper_(data[i][cStatus]);
    if (currentStatus === 'SHIPPED' || currentStatus === 'CANCELLED') continue;

    sheet.getRange(i + 1, cStatus + 1).setValue(newStatus);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// 10. ORDER STATUS + FINALIZATION (Requested_Items / Allocation_Log / Backorders)
// ─────────────────────────────────────────────────────────────────────────────

function updateCustomerOrderShipment(ss, orderId, info) {
  const sheet = ss.getSheetByName('Customer_Orders');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const h = data[0];
  const cOrder = _col_(h, ['order_id', 'order_number']);
  const cStatus = _col_(h, ['request_status', 'status']);
  const cDate = _col_(h, ['ship_date', 'date_shipped', 'shipped_date']);
  const cCarr = _col_(h, ['carrier']);
  const cTrack = _col_(h, ['tracking_number']);
  const cBol = _col_(h, ['bol_number']);

  for (let i = 1; i < data.length; i++) {
    if (_orderKey_(data[i][cOrder]) === _orderKey_(orderId)) {
      if (cStatus >= 0) sheet.getRange(i + 1, cStatus + 1).setValue(info.status);
      if (cDate >= 0) sheet.getRange(i + 1, cDate + 1).setValue(info.shipDate);
      if (cCarr >= 0 && info.carrier) sheet.getRange(i + 1, cCarr + 1).setValue(info.carrier);
      if (cTrack >= 0 && info.trackingNumber) sheet.getRange(i + 1, cTrack + 1).setValue(info.trackingNumber);
      if (cBol >= 0 && info.bolNumber) sheet.getRange(i + 1, cBol + 1).setValue(info.bolNumber);
      break;
    }
  }

  // Finalize related tabs
  finalizeRequestedItemsForShipment(ss, orderId);
  finalizeAllocationLogForShipment(ss, orderId);
  finalizeBackordersSheet(ss, orderId);
}

function _getShippedQtyByOrderFbpn_(ss, orderId) {
  const out = {};
  const sh = ss.getSheetByName('OutboundLog');
  if (!sh) return out;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return out;

  const h = data[0];
  const cOrder = _col_(h, ['order_number', 'order_id', 'order']);
  const cFbpn = _col_(h, ['fbpn']);
  const cQty = _col_(h, ['qty', 'quantity', 'qty_shipped', 'qty_dispatched']);

  if (cOrder < 0 || cFbpn < 0 || cQty < 0) return out;

  for (let i = 1; i < data.length; i++) {
    if (_orderKey_(data[i][cOrder]) !== _orderKey_(orderId)) continue;
    const fbpn = String(data[i][cFbpn] || '').trim();
    if (!fbpn) continue;
    out[fbpn] = (out[fbpn] || 0) + _safeNum_(data[i][cQty]);
  }
  return out;
}

function finalizeRequestedItemsForShipment(ss, orderId) {
  const sheet = ss.getSheetByName('Requested_Items');
  if (!sheet) return;

  const shipped = _getShippedQtyByOrderFbpn_(ss, orderId);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const h = data[0];

  // Your headers (from your message): Order_ID, FBPN, Description, Qty_Requested, Stock_Status,
  // Qty_Backordered, Qty_Allocated, Qty_Shipped, Backorder_ID, Allocation_ID, SKU
  const cOrder = _col_(h, ['order_id', 'order_number', 'order']);
  const cFbpn = _col_(h, ['fbpn']);
  const cQtyReq = _col_(h, ['qty_requested']);
  const cQtyBack = _col_(h, ['qty_backordered']);
  const cQtyAlloc = _col_(h, ['qty_allocated']);
  const cQtyShip = _col_(h, ['qty_shipped']);
  const cStockStatus = _col_(h, ['stock_status']);

  if (cOrder < 0 || cFbpn < 0) return;

  for (let i = 1; i < data.length; i++) {
    if (_orderKey_(data[i][cOrder]) !== _orderKey_(orderId)) continue;

    const fbpn = String(data[i][cFbpn] || '').trim();
    if (!fbpn) continue;

    const req = cQtyReq >= 0 ? _safeNum_(data[i][cQtyReq]) : 0;
    const shipQty = shipped[fbpn] || 0;
    const back = Math.max(0, req - shipQty);

    if (cQtyShip >= 0) sheet.getRange(i + 1, cQtyShip + 1).setValue(shipQty);
    if (cQtyAlloc >= 0) sheet.getRange(i + 1, cQtyAlloc + 1).setValue(0);
    if (cQtyBack >= 0) sheet.getRange(i + 1, cQtyBack + 1).setValue(back);

    if (cStockStatus >= 0) {
      const status = (back <= 0) ? 'Complete' : (shipQty > 0 ? 'Partial' : 'Out of Stock');
      sheet.getRange(i + 1, cStockStatus + 1).setValue(status);
    }
  }
}

function finalizeAllocationLogForShipment(ss, orderId) {
  const sheet = ss.getSheetByName('Allocation_Log');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const h = data[0];
  const cOrder = _col_(h, ['order_id', 'order_number', 'order']);
  const cQty = _col_(h, ['qty_allocated', 'allocated_qty', 'allocated']);
  const cStatus = _col_(h, ['allocation_status', 'status']);

  if (cOrder < 0) return;

  for (let i = 1; i < data.length; i++) {
    if (_orderKey_(data[i][cOrder]) !== _orderKey_(orderId)) continue;

    if (cQty >= 0) sheet.getRange(i + 1, cQty + 1).setValue(0);
    if (cStatus >= 0) sheet.getRange(i + 1, cStatus + 1).setValue('SHIPPED');
  }
}

function _ensureBackorderFromRequestedItems_(ss, orderId) {
  const req = ss.getSheetByName('Requested_Items');
  const bo = ss.getSheetByName('Backorders');
  if (!req || !bo) return 0;

  const reqData = req.getDataRange().getValues();
  if (reqData.length < 2) return 0;

  const rh = reqData[0];
  const rOrder = _col_(rh, ['order_id', 'order_number']);
  const rFbpn = _col_(rh, ['fbpn']);
  const rQtyReq = _col_(rh, ['qty_requested']);
  const rQtyBack = _col_(rh, ['qty_backordered']);
  const rStock = _col_(rh, ['stock_status']);
  const rSku = _col_(rh, ['sku']);

  if (rOrder < 0 || rFbpn < 0 || rQtyBack < 0) return 0;

  // load existing backorders for order+fbpn to avoid duplicates
  const boData = bo.getDataRange().getValues();
  const bh = boData[0] || [];
  const bOrder = _col_(bh, ['order_id', 'order_number']);
  const bFbpn = _col_(bh, ['fbpn']);
  const bQtyBack = _col_(bh, ['qty_backordered', 'qty_needed', 'qty_short']);
  const bStatus = _col_(bh, ['status']);
  const bSku = _col_(bh, ['sku']);
  const bBackId = _col_(bh, ['backorder_id']);

  const existing = {};
  if (boData.length >= 2 && bOrder >= 0 && bFbpn >= 0) {
    for (let i = 1; i < boData.length; i++) {
      const ok = _orderKey_(boData[i][bOrder]);
      const fb = String(boData[i][bFbpn] || '').trim();
      if (!ok || !fb) continue;
      existing[ok + '|' + fb] = true;
    }
  }

  let created = 0;
  for (let i = 1; i < reqData.length; i++) {
    if (_orderKey_(reqData[i][rOrder]) !== _orderKey_(orderId)) continue;

    const fbpn = String(reqData[i][rFbpn] || '').trim();
    const qtyBack = _safeNum_(reqData[i][rQtyBack]);
    if (!fbpn || qtyBack <= 0) continue;

    const key = _orderKey_(orderId) + '|' + fbpn;
    if (existing[key]) continue;

    const stockStatus = rStock >= 0 ? String(reqData[i][rStock] || '') : '';
    const qtyReq = rQtyReq >= 0 ? _safeNum_(reqData[i][rQtyReq]) : '';
    const sku = rSku >= 0 ? String(reqData[i][rSku] || '') : '';

    const backorderId = 'BO-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    // Try to match your prior header shape; if different, still appends usable row
    // Expected per your old comments:
    // Order_ID, NBD, Status, Task_Number, Stock_Status, FBPN, Qty_Requested, Qty_Backordered, Qty_Fulfilled,
    // Date_Logged, Date_Closed, Notes, Backorder_ID, SKU
    const row = [];
    row[bOrder >= 0 ? bOrder : 0] = orderId;
    if (bh.length < 14) {
      // append in the common order if headers unknown
      bo.appendRow([
        orderId, '', 'Open', '', stockStatus, fbpn, qtyReq, qtyBack, 0,
        new Date(), '', 'Auto-generated from shipment', backorderId, sku
      ]);
    } else {
      // if headers exist, set by col index where possible
      const mk = (col, val) => { if (col >= 0) row[col] = val; };
      mk(bOrder, orderId);
      mk(_col_(bh, ['nbd', 'need_by_date', 'need by date']), '');
      mk(bStatus, 'Open');
      mk(_col_(bh, ['task_number', 'task number']), '');
      mk(_col_(bh, ['stock_status']), stockStatus);
      mk(bFbpn, fbpn);
      mk(_col_(bh, ['qty_requested']), qtyReq);
      mk(bQtyBack, qtyBack);
      mk(_col_(bh, ['qty_fulfilled']), 0);
      mk(_col_(bh, ['date_logged']), new Date());
      mk(_col_(bh, ['date_closed']), '');
      mk(_col_(bh, ['notes']), 'Auto-generated from shipment');
      mk(bBackId, backorderId);
      mk(bSku, sku);
      // normalize length
      for (let k = 0; k < bh.length; k++) if (row[k] === undefined) row[k] = '';
      bo.appendRow(row);
    }

    created++;
    existing[key] = true;
  }

  return created;
}

function finalizeBackordersSheet(ss, orderId) {
  // Ensure backorders exist for any remaining shortages
  const created = _ensureBackorderFromRequestedItems_(ss, orderId);

  // Close backorders if fully fulfilled (Qty_Backordered <= Qty_Fulfilled)
  const bo = ss.getSheetByName('Backorders');
  if (!bo) return;

  const data = bo.getDataRange().getValues();
  if (data.length < 2) return;

  const h = data[0];
  const cOrder = _col_(h, ['order_id', 'order_number']);
  const cQtyBack = _col_(h, ['qty_backordered', 'qty_needed', 'qty_short']);
  const cQtyFul = _col_(h, ['qty_fulfilled', 'qty_fulfilled_total']);
  const cStatus = _col_(h, ['status']);
  const cClosed = _col_(h, ['date_closed']);

  if (cOrder < 0 || cQtyBack < 0 || cStatus < 0) return;

  for (let i = 1; i < data.length; i++) {
    if (_orderKey_(data[i][cOrder]) !== _orderKey_(orderId)) continue;

    const back = _safeNum_(data[i][cQtyBack]);
    const ful = cQtyFul >= 0 ? _safeNum_(data[i][cQtyFul]) : 0;

    if (back > 0 && ful >= back) {
      const st = _upper_(data[i][cStatus]);
      if (st !== 'CLOSED') {
        bo.getRange(i + 1, cStatus + 1).setValue('Closed');
        if (cClosed >= 0) bo.getRange(i + 1, cClosed + 1).setValue(new Date());
      }
    }
  }

  return created;
}

// ─────────────────────────────────────────────────────────────────────────────
// 11. DELIVERED ORDER SYNC UTILITIES
// ─────────────────────────────────────────────────────────────────────────────

function runOneTimeDeliveredOrderSync() {
  // Simple: same as forceUpdateDeliveredOrders; kept for menu compatibility
  return forceUpdateDeliveredOrders();
}

function forceUpdateDeliveredOrders() {
  var ss = getSpreadsheetForShipping_();
  var coSheet = ss.getSheetByName('Customer_Orders');
  if (!coSheet) return { success: false, message: 'Customer_Orders missing' };

  var data = coSheet.getDataRange().getValues();
  var h = data[0];

  var cOrder = _col_(h, ['order_id', 'order_number', 'order']);
  var cStatus = _col_(h, ['request_status', 'status']);

  if (cOrder < 0 || cStatus < 0) return { success: false, message: 'Missing columns in Customer_Orders' };

  var processed = 0;
  for (var i = 1; i < data.length; i++) {
    var orderId = String(data[i][cOrder] || '').trim();
    if (!orderId) continue;

    var status = _upper_(data[i][cStatus]);
    if (status === 'DELIVERED' || status === 'SHIPPED' || status === 'COMPLETE') {
      finalizeRequestedItemsForShipment(ss, orderId);
      finalizeAllocationLogForShipment(ss, orderId);
      finalizeBackordersSheet(ss, orderId);
      processed++;
    }
  }

  SpreadsheetApp.flush();
  return { success: true, processed: processed };
}
