/**
 * Cycle Count Feature
 * - Create cycle count batches
 * - Track progress
 * - Update inventory based on count results
 */

// ============================================================================
// MODAL
// ============================================================================
function showCycleCountModal() {
  const template = HtmlService.createTemplateFromFile('CycleCountModal');
  const html = template.evaluate()
    .setWidth(1200)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Cycle Count');
}

// ============================================================================
// CORE CONFIG
// ============================================================================
function getCycleConfig_() {
  return {
    BIN_STOCK_SHEET: 'Bin_Stock',
    CYCLE_COUNT_SHEET: 'Cycle_Count',
    LOCATIONLOG_SHEET: 'LocationLog',

    HEADERS: {
      CYCLE_COUNT: [
        'Batch_ID',        // A
        'Status',          // B  (Open, In Progress, Completed, Canceled)
        'Created_At',      // C
        'Created_By',      // D
        'Bin_Code',        // E
        'FBPN',            // F
        'Manufacturer',    // G
        'Project',         // H
        'Current_Qty',     // I (snapshot at batch creation)
        'Counted_Qty',     // J
        'Variance',        // K (Counted - Current)
        'Notes',           // L
        'Counted_At',      // M
        'Counted_By'       // N
      ]
    }
  };
}

// ============================================================================
// PUBLIC API FOR HTML
// ============================================================================

/**
 * Fetch bins eligible for cycle counting, with filters.
 * filters: {
 *   project?: string,
 *   manufacturer?: string,
 *   fbpn?: string,
 *   binCode?: string,
 *   status?: string ("Open","In Progress","Completed","All")
 * }
 */
function imsGetCycleCountBins(filters) {
  const cfg = getCycleConfig_();
  const ss = SpreadsheetApp.getActive();
  const binSheet = ss.getSheetByName(cfg.BIN_STOCK_SHEET);
  const cycleSheet = ss.getSheetByName(cfg.CYCLE_COUNT_SHEET);

  if (!binSheet || !cycleSheet) {
    throw new Error('Bin_Stock or Cycle_Count sheet not found.');
  }

  const binData = binSheet.getDataRange().getValues();
  const binHeaders = binData[0];

  const idxBinCode = binHeaders.indexOf('Bin_Code');
  const idxFBPN = binHeaders.indexOf('FBPN');
  const idxMan = binHeaders.indexOf('Manufacturer');
  const idxProj = binHeaders.indexOf('Project');
  const idxQty = binHeaders.indexOf('Current_Quantity');

  if ([idxBinCode, idxFBPN, idxMan, idxProj, idxQty].some(i => i === -1)) {
    throw new Error('Bin_Stock headers missing required columns.');
  }

  const cycleData = cycleSheet.getDataRange().getValues();
  const cycleHeaders = cycleData[0];

  const idxBatch = cycleHeaders.indexOf('Batch_ID');
  const idxStatus = cycleHeaders.indexOf('Status');
  const idxBinCodeC = cycleHeaders.indexOf('Bin_Code');
  const idxFBPNC = cycleHeaders.indexOf('FBPN');
  const idxManC = cycleHeaders.indexOf('Manufacturer');
  const idxProjC = cycleHeaders.indexOf('Project');
  const idxCurrentC = cycleHeaders.indexOf('Current_Qty');
  const idxCounted = cycleHeaders.indexOf('Counted_Qty');
  const idxVar = cycleHeaders.indexOf('Variance');

  const filterStatus = (filters && filters.status) || 'Open';
  const filterBin = (filters && filters.binCode || '').toString().trim().toUpperCase();
  const filterFBPN = (filters && filters.fbpn || '').toString().trim().toUpperCase();
  const filterMan = (filters && filters.manufacturer || '').toString().trim().toUpperCase();
  const filterProj = (filters && filters.project || '').toString().trim().toUpperCase();

  const existingMap = {};
  for (let i = 1; i < cycleData.length; i++) {
    const r = cycleData[i];
    const bin = (r[idxBinCodeC] || '').toString().toUpperCase();
    const fbpn = (r[idxFBPNC] || '').toString().toUpperCase();
    const man = (r[idxManC] || '').toString().toUpperCase();
    const proj = (r[idxProjC] || '').toString().toUpperCase();

    const key = [bin, fbpn, man, proj].join('||');
    existingMap[key] = existingMap[key] || [];
    existingMap[key].push({
      batchId: r[idxBatch],
      status: r[idxStatus],
      currentQtySnapshot: r[idxCurrentC],
      countedQty: r[idxCounted],
      variance: r[idxVar]
    });
  }

  const results = [];

  for (let i = 1; i < binData.length; i++) {
    const row = binData[i];
    const binCode = (row[idxBinCode] || '').toString().toUpperCase();
    const fbpn = (row[idxFBPN] || '').toString().toUpperCase();
    const man = (row[idxMan] || '').toString().toUpperCase();
    const proj = (row[idxProj] || '').toString().toUpperCase();
    const qty = row[idxQty] || 0;

    if (filterBin && !binCode.includes(filterBin)) continue;
    if (filterFBPN && !fbpn.includes(filterFBPN)) continue;
    if (filterMan && !man.includes(filterMan)) continue;
    if (filterProj && !proj.includes(filterProj)) continue;

    const key = [binCode, fbpn, man, proj].join('||');
    const hist = existingMap[key] || [];

    let matchesStatus = false;
    if (filterStatus === 'All') {
      matchesStatus = true;
    } else if (!hist.length && filterStatus === 'Open') {
      matchesStatus = true;
    } else {
      matchesStatus = hist.some(h => h.status === filterStatus);
    }

    if (!matchesStatus) continue;

    results.push({
      binCode,
      fbpn,
      manufacturer: man,
      project: proj,
      currentQty: qty,
      history: hist
    });
  }

  return results;
}

/**
 * Create a new cycle count batch for selected lines from Bin_Stock.
 * payload: {
 *   lines: [{ binCode, fbpn, manufacturer, project }],
 *   batchId?: string,
 *   createdBy?: string
 * }
 */
function imsCreateCycleCountBatch(payload) {
  const cfg = getCycleConfig_();
  const ss = SpreadsheetApp.getActive();
  const binSheet = ss.getSheetByName(cfg.BIN_STOCK_SHEET);
  const cycleSheet = ss.getSheetByName(cfg.CYCLE_COUNT_SHEET);

  if (!binSheet || !cycleSheet) {
    throw new Error('Bin_Stock or Cycle_Count sheet not found.');
  }

  const lines = payload && payload.lines || [];
  if (!lines.length) throw new Error('No lines provided.');

  const createdBy = (payload && payload.createdBy) || Session.getActiveUser().getEmail();
  const now = new Date();
  const batchId = payload.batchId || generateCycleBatchId_(cycleSheet);

  const binData = binSheet.getDataRange().getValues();
  const binHeaders = binData[0];

  const idxBinCode = binHeaders.indexOf('Bin_Code');
  const idxFBPN = binHeaders.indexOf('FBPN');
  const idxMan = binHeaders.indexOf('Manufacturer');
  const idxProj = binHeaders.indexOf('Project');
  const idxQty = binHeaders.indexOf('Current_Quantity');

  if ([idxBinCode, idxFBPN, idxMan, idxProj, idxQty].some(i => i === -1)) {
    throw new Error('Bin_Stock headers missing required columns.');
  }

  const snapshotMap = {};
  for (let i = 1; i < binData.length; i++) {
    const r = binData[i];
    const bin = (r[idxBinCode] || '').toString().toUpperCase();
    const fbpn = (r[idxFBPN] || '').toString().toUpperCase();
    const man = (r[idxMan] || '').toString().toUpperCase();
    const proj = (r[idxProj] || '').toString().toUpperCase();
    const qty = r[idxQty] || 0;

    const key = [bin, fbpn, man, proj].join('||');
    snapshotMap[key] = qty;
  }

  const cycleData = cycleSheet.getDataRange().getValues();
  const cycleHeaders = cycleData[0];

  const idxBatch = cycleHeaders.indexOf('Batch_ID');
  const idxStatus = cycleHeaders.indexOf('Status');
  const idxCreatedAt = cycleHeaders.indexOf('Created_At');
  const idxCreatedBy = cycleHeaders.indexOf('Created_By');
  const idxBinCodeC = cycleHeaders.indexOf('Bin_Code');
  const idxFBPNC = cycleHeaders.indexOf('FBPN');
  const idxManC = cycleHeaders.indexOf('Manufacturer');
  const idxProjC = cycleHeaders.indexOf('Project');
  const idxCurrentC = cycleHeaders.indexOf('Current_Qty');
  const idxCounted = cycleHeaders.indexOf('Counted_Qty');
  const idxVar = cycleHeaders.indexOf('Variance');
  const idxNotes = cycleHeaders.indexOf('Notes');
  const idxCountedAt = cycleHeaders.indexOf('Counted_At');
  const idxCountedBy = cycleHeaders.indexOf('Counted_By');

  const newRows = [];

  lines.forEach(line => {
    const bin = (line.binCode || '').toString().toUpperCase();
    const fbpn = (line.fbpn || '').toString().toUpperCase();
    const man = (line.manufacturer || '').toString().toUpperCase();
    const proj = (line.project || '').toString().toUpperCase();

    if (!bin || !fbpn || !man || !proj) return;

    const key = [bin, fbpn, man, proj].join('||');
    const currentQtySnapshot = snapshotMap[key];

    if (typeof currentQtySnapshot === 'undefined') return;

    const row = new Array(cycleHeaders.length).fill('');
    row[idxBatch] = batchId;
    row[idxStatus] = 'Open';
    row[idxCreatedAt] = now;
    row[idxCreatedBy] = createdBy;
    row[idxBinCodeC] = bin;
    row[idxFBPNC] = fbpn;
    row[idxManC] = man;
    row[idxProjC] = proj;
    row[idxCurrentC] = currentQtySnapshot;
    row[idxCounted] = '';
    row[idxVar] = '';
    row[idxNotes] = '';
    row[idxCountedAt] = '';
    row[idxCountedBy] = '';

    newRows.push(row);
  });

  if (!newRows.length) {
    throw new Error('No valid lines to add to Cycle_Count.');
  }

  if (cycleSheet.getLastRow() === 0) {
    cycleSheet.appendRow(getCycleConfig_().HEADERS.CYCLE_COUNT);
  }

  const startRow = cycleSheet.getLastRow() + 1;
  cycleSheet.getRange(startRow, 1, newRows.length, cycleHeaders.length).setValues(newRows);

  return {
    success: true,
    batchId,
    addedLines: newRows.length
  };
}

/**
 * Update count results for a batch line.
 * payload: {
 *   batchId,
 *   binCode,
 *   fbpn,
 *   manufacturer,
 *   project,
 *   countedQty,
 *   notes
 * }
 */
function imsSubmitCycleCountLine(payload) {
  const cfg = getCycleConfig_();
  const ss = SpreadsheetApp.getActive();
  const cycleSheet = ss.getSheetByName(cfg.CYCLE_COUNT_SHEET);
  const binSheet = ss.getSheetByName(cfg.BIN_STOCK_SHEET);
  const logSheet = ss.getSheetByName(cfg.LOCATIONLOG_SHEET);

  if (!cycleSheet || !binSheet || !logSheet) {
    throw new Error('Required sheets not found.');
  }

  const batchId = (payload && payload.batchId) || '';
  const bin = (payload && payload.binCode || '').toString().toUpperCase();
  const fbpn = (payload && payload.fbpn || '').toString().toUpperCase();
  const man = (payload && payload.manufacturer || '').toString().toUpperCase();
  const proj = (payload && payload.project || '').toString().toUpperCase();
  const countedQty = parseFloat(payload && payload.countedQty);
  const notes = (payload && payload.notes) || '';
  const countedBy = Session.getActiveUser().getEmail();
  const now = new Date();

  if (!batchId || !bin || !fbpn || !man || !proj || isNaN(countedQty)) {
    throw new Error('Missing required fields for cycle count submission.');
  }

  const cycleData = cycleSheet.getDataRange().getValues();
  const headers = cycleData[0];

  const idxBatch = headers.indexOf('Batch_ID');
  const idxStatus = headers.indexOf('Status');
  const idxBinCode = headers.indexOf('Bin_Code');
  const idxFBPN = headers.indexOf('FBPN');
  const idxMan = headers.indexOf('Manufacturer');
  const idxProj = headers.indexOf('Project');
  const idxCurrent = headers.indexOf('Current_Qty');
  const idxCounted = headers.indexOf('Counted_Qty');
  const idxVar = headers.indexOf('Variance');
  const idxNotes = headers.indexOf('Notes');
  const idxCountedAt = headers.indexOf('Counted_At');
  const idxCountedBy = headers.indexOf('Counted_By');

  let rowIndex = -1;
  let currentQtySnapshot = 0;

  for (let i = 1; i < cycleData.length; i++) {
    const r = cycleData[i];
    if (
      r[idxBatch] === batchId &&
      (r[idxBinCode] || '').toString().toUpperCase() === bin &&
      (r[idxFBPN] || '').toString().toUpperCase() === fbpn &&
      (r[idxMan] || '').toString().toUpperCase() === man &&
      (r[idxProj] || '').toString().toUpperCase() === proj
    ) {
      rowIndex = i + 1;
      currentQtySnapshot = r[idxCurrent] || 0;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('Matching cycle count line not found.');
  }

  const variance = countedQty - currentQtySnapshot;

  const updateRow = cycleSheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  updateRow[idxStatus] = 'Completed';
  updateRow[idxCounted] = countedQty;
  updateRow[idxVar] = variance;
  updateRow[idxNotes] = notes;
  updateRow[idxCountedAt] = now;
  updateRow[idxCountedBy] = countedBy;

  cycleSheet.getRange(rowIndex, 1, 1, headers.length).setValues([updateRow]);

  applyCycleCountToInventory_(binSheet, logSheet, {
    bin,
    fbpn,
    manufacturer: man,
    project: proj,
    countedQty,
    variance,
    notes,
    countedBy,
    countedAt: now,
    batchId
  });

  updateBatchStatus_(cycleSheet, batchId);

  return {
    success: true,
    variance
  };
}

/**
 * Get batch summary + lines.
 * batchId: string
 */
function imsGetCycleBatch(batchId) {
  const cfg = getCycleConfig_();
  const ss = SpreadsheetApp.getActive();
  const cycleSheet = ss.getSheetByName(cfg.CYCLE_COUNT_SHEET);

  if (!cycleSheet) {
    throw new Error('Cycle_Count sheet not found.');
  }

  const data = cycleSheet.getDataRange().getValues();
  const headers = data[0];

  const idxBatch = headers.indexOf('Batch_ID');
  const idxStatus = headers.indexOf('Status');
  const idxCreatedAt = headers.indexOf('Created_At');
  const idxCreatedBy = headers.indexOf('Created_By');
  const idxBin = headers.indexOf('Bin_Code');
  const idxFBPN = headers.indexOf('FBPN');
  const idxMan = headers.indexOf('Manufacturer');
  const idxProj = headers.indexOf('Project');
  const idxCurrent = headers.indexOf('Current_Qty');
  const idxCounted = headers.indexOf('Counted_Qty');
  const idxVar = headers.indexOf('Variance');
  const idxNotes = headers.indexOf('Notes');
  const idxCountedAt = headers.indexOf('Counted_At');
  const idxCountedBy = headers.indexOf('Counted_By');

  const lines = [];
  let overallStatus = 'Open';
  let createdAt = null;
  let createdBy = null;

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r[idxBatch] !== batchId) continue;

    const lineStatus = r[idxStatus] || 'Open';
    if (!createdAt) createdAt = r[idxCreatedAt];
    if (!createdBy) createdBy = r[idxCreatedBy];

    lines.push({
      binCode: r[idxBin],
      fbpn: r[idxFBPN],
      manufacturer: r[idxMan],
      project: r[idxProj],
      currentQtySnapshot: r[idxCurrent],
      countedQty: r[idxCounted],
      variance: r[idxVar],
      notes: r[idxNotes],
      status: lineStatus,
      countedAt: r[idxCountedAt],
      countedBy: r[idxCountedBy]
    });
  }

  if (!lines.length) {
    throw new Error('Batch not found or empty.');
  }

  const statuses = lines.map(l => l.status || 'Open');
  if (statuses.every(s => s === 'Completed')) {
    overallStatus = 'Completed';
  } else if (statuses.some(s => s === 'Completed')) {
    overallStatus = 'In Progress';
  }

  return {
    batchId,
    createdAt,
    createdBy,
    status: overallStatus,
    lines
  };
}

// ============================================================================
// INTERNAL HELPERS
// ============================================================================

function generateCycleBatchId_(cycleSheet) {
  const data = cycleSheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxBatch = headers.indexOf('Batch_ID');

  if (idxBatch === -1 || data.length <= 1) {
    return 'CC-0001';
  }

  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][idxBatch] || '').toString();
    const match = id.match(/^CC-(\d+)$/);
    if (match) {
      const n = parseInt(match[1], 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
  }

  const next = maxNum + 1;
  return 'CC-' + String(next).padStart(4, '0');
}

function applyCycleCountToInventory_(binSheet, logSheet, payload) {
  const binData = binSheet.getDataRange().getValues();
  const binHeaders = binData[0];

  const idxBin = binHeaders.indexOf('Bin_Code');
  const idxFBPN = binHeaders.indexOf('FBPN');
  const idxMan = binHeaders.indexOf('Manufacturer');
  const idxProj = binHeaders.indexOf('Project');
  const idxCurrent = binHeaders.indexOf('Current_Quantity');
  const idxInitial = binHeaders.indexOf('Initial_Quantity');
  const idxPct = binHeaders.indexOf('Stock_Percentage');

  if ([idxBin, idxFBPN, idxMan, idxProj, idxCurrent, idxInitial, idxPct].some(i => i === -1)) {
    throw new Error('Bin_Stock headers missing required columns for cycle application.');
  }

  const logData = logSheet.getDataRange().getValues();
  const logHeaders = logData[0];

  const idxTime = logHeaders.indexOf('Timestamp');
  const idxAction = logHeaders.indexOf('Action');
  const idxLFbpn = logHeaders.indexOf('FBPN');
  const idxLMan = logHeaders.indexOf('Manufacturer');
  const idxLBin = logHeaders.indexOf('Bin_Code');
  const idxLQtyChanged = logHeaders.indexOf('Qty_Changed');
  const idxLResulting = logHeaders.indexOf('Resulting_Qty');
  const idxLDesc = logHeaders.indexOf('Description');
  const idxLUser = logHeaders.indexOf('User_Email');
  const idxLProj = logHeaders.indexOf('Project');

  const binKey = [
    payload.bin || payload.binCode,
    payload.fbpn,
    payload.manufacturer,
    payload.project
  ].map(v => (v || '').toString().toUpperCase());

  let binRowIndex = -1;
  let currentQty = 0;

  for (let i = 1; i < binData.length; i++) {
    const r = binData[i];
    const rowKey = [
      (r[idxBin] || '').toString().toUpperCase(),
      (r[idxFBPN] || '').toString().toUpperCase(),
      (r[idxMan] || '').toString().toUpperCase(),
      (r[idxProj] || '').toString().toUpperCase()
    ];
    if (rowKey.join('||') === binKey.join('||')) {
      binRowIndex = i + 1;
      currentQty = parseFloat(r[idxCurrent]) || 0;
      break;
    }
  }

  if (binRowIndex === -1) {
    throw new Error('Matching bin row not found for cycle count application.');
  }

  const countedQty = payload.countedQty;
  const variance = payload.variance;
  const newQty = countedQty;
  const rowValues = binSheet.getRange(binRowIndex, 1, 1, binHeaders.length).getValues()[0];

  rowValues[idxCurrent] = newQty;

  const initialQty = rowValues[idxInitial] || newQty || 0;
  rowValues[idxInitial] = initialQty;
  rowValues[idxPct] = initialQty ? (newQty / initialQty) * 100 : 0;

  binSheet.getRange(binRowIndex, 1, 1, binHeaders.length).setValues([rowValues]);

  const logRow = new Array(logHeaders.length).fill('');
  logRow[idxTime] = payload.countedAt || new Date();
  logRow[idxAction] = 'CYCLE_ADJUST';
  logRow[idxLFbpn] = payload.fbpn;
  logRow[idxLMan] = payload.manufacturer;
  logRow[idxLBin] = payload.bin || payload.binCode;
  logRow[idxLQtyChanged] = variance;
  logRow[idxLResulting] = newQty;
  logRow[idxLDesc] = (payload.notes || '') +
    ` [Cycle Count ${payload.batchId || ''}, from ${currentQty} to ${newQty}]`;
  logRow[idxLUser] = payload.countedBy || '';
  logRow[idxLProj] = payload.project;

  logSheet.appendRow(logRow);
}

function updateBatchStatus_(cycleSheet, batchId) {
  const data = cycleSheet.getDataRange().getValues();
  const headers = data[0];

  const idxBatch = headers.indexOf('Batch_ID');
  const idxStatus = headers.indexOf('Status');

  if (idxBatch === -1 || idxStatus === -1) return;

  let anyCompleted = false;
  let anyOpen = false;

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r[idxBatch] !== batchId) continue;
    const s = r[idxStatus] || 'Open';
    if (s === 'Completed') anyCompleted = true;
    if (s === 'Open' || s === 'In Progress') anyOpen = true;
  }

  const newStatus = anyOpen
    ? (anyCompleted ? 'In Progress' : 'Open')
    : 'Completed';

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r[idxBatch] !== batchId) continue;

    if (r[idxStatus] !== newStatus && r[idxStatus] !== 'Completed') {
      cycleSheet.getRange(i + 1, idxStatus + 1).setValue(newStatus);
    }
  }
}
