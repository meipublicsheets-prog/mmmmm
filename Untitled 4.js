function debugPickTicketLookup(taskNumber) {
  var ss = getSpreadsheetForShipping_();
  var ordersSh = ss.getSheetByName('Customer_Orders');
  var pickSh = ss.getSheetByName('Pick_Log');

  var key = _norm_(taskNumber);
  var out = {
    ok: true,
    input: taskNumber,
    normInput: key,
    hasCustomerOrders: !!ordersSh,
    hasPickLog: !!pickSh,
    customerOrdersHeaders: [],
    pickLogHeaders: [],
    resolved: {
      orderId: '',
      taskNumber: ''
    },
    pickLog: {
      matchedRows: 0,
      sampleMatches: []
    }
  };

  if (!key) return { ok: false, error: 'Missing taskNumber' };
  if (!ordersSh) return { ok: false, error: 'Customer_Orders not found' };
  if (!pickSh) return { ok: false, error: 'Pick_Log not found' };

  var od = ordersSh.getDataRange().getValues();
  if (od.length < 2) return { ok: false, error: 'Customer_Orders empty' };

  var oh = od[0];
  out.customerOrdersHeaders = oh;

  var cTask = _col_(oh, ['task_number','task number','task#','task #','task']);
  var cOrder = _col_(oh, ['order_id','order id','order_number','order number','order']);
  var cStatus = _col_(oh, ['request_status','status','request status']);

  var orderRow = null;
  for (var r = 1; r < od.length; r++) {
    var row = od[r];
    var st = (cStatus >= 0) ? _upper_(row[cStatus]) : '';
    // don’t block on status for debug, just record
    var vTask = (cTask >= 0) ? row[cTask] : '';
    var vOrder = (cOrder >= 0) ? row[cOrder] : '';
    if (_eq_(vTask, key) || _eq_(vOrder, key)) { orderRow = row; break; }
  }
  if (!orderRow) return { ok: false, error: 'No matching row in Customer_Orders for task/order ' + key, debug: out };

  var orderId =
    _norm_(cOrder >= 0 ? orderRow[cOrder] : '') ||
    _norm_(cTask >= 0 ? orderRow[cTask] : '') ||
    key;

  out.resolved.orderId = orderId;
  out.resolved.taskNumber = _norm_(cTask >= 0 ? orderRow[cTask] : key);

  var pd = pickSh.getDataRange().getValues();
  if (pd.length < 2) return { ok: true, message: 'Pick_Log empty', debug: out };

  var ph = pd[0];
  out.pickLogHeaders = ph;

  // Detect *any* plausible order/task column
  var pOrder = _col_(ph, ['order_number','order number','order_id','order id','order #','order']);
  var pTask  = _col_(ph, ['task_number','task number','task#','task #','task']);
  var pFbpn  = _col_(ph, ['fbpn','item','part number']);
  var pQty   = _col_(ph, ['qty_to_pick','quantity to pick','qty pick','qty to pick','qty_requested','qty requested']);

  // Create match set across order + task + normalized keys
  var rawCandidates = [
    key,
    orderId,
    out.resolved.taskNumber
  ].filter(Boolean);

  var matchSet = {};
  rawCandidates.forEach(function(v) {
    matchSet[_norm_(v)] = true;
    matchSet[_orderKey_(v)] = true;
    // numeric trunc variants
    var n = Number(v);
    if (!isNaN(n)) {
      matchSet[String(Math.trunc(n))] = true;
      matchSet[_orderKey_(String(Math.trunc(n)))] = true;
    }
  });

  for (var i = 1; i < pd.length; i++) {
    var row2 = pd[i];

    var hit = false;
    if (pOrder >= 0) {
      var vOrder = row2[pOrder];
      if (matchSet[_norm_(vOrder)] || matchSet[_orderKey_(vOrder)]) hit = true;
    }
    if (!hit && pTask >= 0) {
      var vTask = row2[pTask];
      if (matchSet[_norm_(vTask)] || matchSet[_orderKey_(vTask)]) hit = true;
    }
    if (!hit) continue;

    // must have FBPN + qty>0 to be considered usable
    var fbpn = (pFbpn >= 0) ? String(row2[pFbpn] || '').trim() : '';
    var qty = (pQty >= 0) ? _safeNum_(row2[pQty]) : 0;
    if (!fbpn || qty <= 0) continue;

    out.pickLog.matchedRows++;
    if (out.pickLog.sampleMatches.length < 5) {
      out.pickLog.sampleMatches.push({
        row: i + 1,
        fbpn: fbpn,
        qty: qty,
        orderCell: (pOrder >= 0 ? row2[pOrder] : ''),
        taskCell: (pTask >= 0 ? row2[pTask] : '')
      });
    }
  }

  return out;
}
/**
 * Creates a new tab and compares Master_Log.Txn_ID to Inbound_Skids.TXN_ID.
 * Outputs:
 *  1) Txn_IDs present in Master_Log but missing in Inbound_Skids
 *  2) TXN_IDs present in Inbound_Skids but missing in Master_Log
 *
 * Master_Log headers:
 * Txn_ID Date_Received Transaction_Type Warehouse Push # FBPN Qty_Received Total_Skid_Count
 * Carrier BOL_Number Customer_PO_Number Manufacturer MFPN Description Received_By SKU
 *
 * Inbound_Skids headers:
 * Skid_ID TXN_ID Date FBPN MFPN Project Qty_on_Skid Skid_Sequence Is_Mixed Timestamp SKU
 */
function reconTxnIds_MasterLog_vs_InboundSkids() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const MASTER_SHEET = 'Master_Log';
  const SKIDS_SHEET = 'Inbound_Skids';

  const masterSh = ss.getSheetByName(MASTER_SHEET);
  const skidsSh = ss.getSheetByName(SKIDS_SHEET);

  if (!masterSh) throw new Error('Sheet not found: ' + MASTER_SHEET);
  if (!skidsSh) throw new Error('Sheet not found: ' + SKIDS_SHEET);

  const masterData = masterSh.getDataRange().getValues();
  const skidsData = skidsSh.getDataRange().getValues();

  if (masterData.length < 2) throw new Error(MASTER_SHEET + ' has no data rows.');
  if (skidsData.length < 2) throw new Error(SKIDS_SHEET + ' has no data rows.');

  const masterHdr = masterData[0].map(String);
  const skidsHdr = skidsData[0].map(String);

  const masterTxnIdx = masterHdr.indexOf('Txn_ID');
  const skidsTxnIdx = skidsHdr.indexOf('TXN_ID');

  if (masterTxnIdx === -1) throw new Error('Header "Txn_ID" not found in ' + MASTER_SHEET);
  if (skidsTxnIdx === -1) throw new Error('Header "TXN_ID" not found in ' + SKIDS_SHEET);

  // Build sets + lightweight detail maps for reporting
  const masterSet = new Set();
  const masterDetail = new Map(); // txn -> {date, bol, push, warehouse, receivedBy, type}
  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    const txn = (row[masterTxnIdx] || '').toString().trim();
    if (!txn) continue;
    masterSet.add(txn);

    if (!masterDetail.has(txn)) {
      masterDetail.set(txn, {
        Date_Received: getCellByHeader_(row, masterHdr, 'Date_Received'),
        Transaction_Type: getCellByHeader_(row, masterHdr, 'Transaction_Type'),
        Warehouse: getCellByHeader_(row, masterHdr, 'Warehouse'),
        'Push #': getCellByHeader_(row, masterHdr, 'Push #'),
        BOL_Number: getCellByHeader_(row, masterHdr, 'BOL_Number'),
        Received_By: getCellByHeader_(row, masterHdr, 'Received_By'),
      });
    }
  }

  const skidsSet = new Set();
  const skidsDetail = new Map(); // txn -> {firstSkidId, date, project, skuCount}
  const skidsSkidIdIdx = skidsHdr.indexOf('Skid_ID');
  const skidsDateIdx = skidsHdr.indexOf('Date');
  const skidsProjectIdx = skidsHdr.indexOf('Project');

  for (let i = 1; i < skidsData.length; i++) {
    const row = skidsData[i];
    const txn = (row[skidsTxnIdx] || '').toString().trim();
    if (!txn) continue;
    skidsSet.add(txn);

    if (!skidsDetail.has(txn)) {
      skidsDetail.set(txn, {
        Skid_ID: skidsSkidIdIdx >= 0 ? row[skidsSkidIdIdx] : '',
        Date: skidsDateIdx >= 0 ? row[skidsDateIdx] : '',
        Project: skidsProjectIdx >= 0 ? row[skidsProjectIdx] : '',
        Skid_Row_Count: 1,
      });
    } else {
      const d = skidsDetail.get(txn);
      d.Skid_Row_Count = (d.Skid_Row_Count || 0) + 1;
    }
  }

  // Compute diffs
  const missingInSkids = [];
  masterSet.forEach(txn => {
    if (!skidsSet.has(txn)) missingInSkids.push(txn);
  });

  const missingInMaster = [];
  skidsSet.forEach(txn => {
    if (!masterSet.has(txn)) missingInMaster.push(txn);
  });

  missingInSkids.sort();
  missingInMaster.sort();

  // Create report tab
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmmss');
  const reportName = `TXN_Recon_${stamp}`;
  const reportSh = ss.insertSheet(reportName);

  // Title + summary
  reportSh.getRange(1, 1).setValue('TXN_ID Reconciliation: Master_Log vs Inbound_Skids');
  reportSh.getRange(2, 1).setValue('Generated');
  reportSh.getRange(2, 2).setValue(new Date());
  reportSh.getRange(3, 1).setValue('Master_Log unique Txn_ID count');
  reportSh.getRange(3, 2).setValue(masterSet.size);
  reportSh.getRange(4, 1).setValue('Inbound_Skids unique TXN_ID count');
  reportSh.getRange(4, 2).setValue(skidsSet.size);
  reportSh.getRange(5, 1).setValue('Missing in Inbound_Skids');
  reportSh.getRange(5, 2).setValue(missingInSkids.length);
  reportSh.getRange(6, 1).setValue('Missing in Master_Log');
  reportSh.getRange(6, 2).setValue(missingInMaster.length);

  // Section 1: Master_Log Txn_ID missing in Inbound_Skids
  let r = 8;
  reportSh.getRange(r, 1).setValue('A) Present in Master_Log, missing in Inbound_Skids');
  r++;

  const aHeaders = ['Txn_ID', 'Date_Received', 'Transaction_Type', 'Warehouse', 'Push #', 'BOL_Number', 'Received_By'];
  reportSh.getRange(r, 1, 1, aHeaders.length).setValues([aHeaders]);
  r++;

  const aRows = missingInSkids.map(txn => {
    const d = masterDetail.get(txn) || {};
    return [
      txn,
      d.Date_Received || '',
      d.Transaction_Type || '',
      d.Warehouse || '',
      d['Push #'] || '',
      d.BOL_Number || '',
      d.Received_By || '',
    ];
  });

  if (aRows.length) {
    reportSh.getRange(r, 1, aRows.length, aHeaders.length).setValues(aRows);
    r += aRows.length + 2;
  } else {
    reportSh.getRange(r, 1).setValue('(none)');
    r += 3;
  }

  // Section 2: Inbound_Skids TXN_ID missing in Master_Log
  reportSh.getRange(r, 1).setValue('B) Present in Inbound_Skids, missing in Master_Log');
  r++;

  const bHeaders = ['TXN_ID', 'Skid_ID (example)', 'Date (example)', 'Project (example)', 'Skid Row Count'];
  reportSh.getRange(r, 1, 1, bHeaders.length).setValues([bHeaders]);
  r++;

  const bRows = missingInMaster.map(txn => {
    const d = skidsDetail.get(txn) || {};
    return [
      txn,
      d.Skid_ID || '',
      d.Date || '',
      d.Project || '',
      d.Skid_Row_Count || '',
    ];
  });

  if (bRows.length) {
    reportSh.getRange(r, 1, bRows.length, bHeaders.length).setValues(bRows);
  } else {
    reportSh.getRange(r, 1).setValue('(none)');
  }

  // Basic formatting
  reportSh.setFrozenRows(7);
  reportSh.autoResizeColumns(1, 12);

  // Return for logging / testing
  return {
    sheet: reportName,
    masterUnique: masterSet.size,
    skidsUnique: skidsSet.size,
    missingInSkids: missingInSkids.length,
    missingInMaster: missingInMaster.length,
  };
}

/**
 * Helper: get cell value from a row by header name (exact match).
 */
function getCellByHeader_(row, headers, headerName) {
  const idx = headers.indexOf(headerName);
  return idx >= 0 ? row[idx] : '';
}
