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

// ============================================================================
// WEEKLY REPORT FUNCTIONS
// ============================================================================

/**
 * Gets cycle count report data for a date range.
 * @param {Object} params - { startDate, endDate }
 * @returns {Object} Report data with summary and details
 */
function imsGetCycleCountReport(params) {
  try {
    const cfg = getCycleConfig_();
    const ss = SpreadsheetApp.getActive();
    const cycleSheet = ss.getSheetByName(cfg.CYCLE_COUNT_SHEET);

    if (!cycleSheet) {
      return { success: false, message: 'Cycle_Count sheet not found.' };
    }

    const data = cycleSheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: true, summary: getEmptySummary_(), details: [], dateRange: params };
    }

    const headers = data[0];
    const idx = {
      batchId: headers.indexOf('Batch_ID'),
      status: headers.indexOf('Status'),
      createdAt: headers.indexOf('Created_At'),
      createdBy: headers.indexOf('Created_By'),
      binCode: headers.indexOf('Bin_Code'),
      fbpn: headers.indexOf('FBPN'),
      manufacturer: headers.indexOf('Manufacturer'),
      project: headers.indexOf('Project'),
      currentQty: headers.indexOf('Current_Qty'),
      countedQty: headers.indexOf('Counted_Qty'),
      variance: headers.indexOf('Variance'),
      notes: headers.indexOf('Notes'),
      countedAt: headers.indexOf('Counted_At'),
      countedBy: headers.indexOf('Counted_By')
    };

    // Parse date range
    let startDate = null, endDate = null;
    if (params && params.startDate) {
      startDate = new Date(params.startDate);
      startDate.setHours(0, 0, 0, 0);
    }
    if (params && params.endDate) {
      endDate = new Date(params.endDate);
      endDate.setHours(23, 59, 59, 999);
    }

    // If no dates provided, default to last 7 days
    if (!startDate && !endDate) {
      endDate = new Date();
      endDate.setHours(23, 59, 59, 999);
      startDate = new Date();
      startDate.setDate(startDate.getDate() - 7);
      startDate.setHours(0, 0, 0, 0);
    }

    const details = [];
    const batchSummary = {};
    const binsCounted = new Set();
    let totalVariance = 0;
    let positiveVariance = 0;
    let negativeVariance = 0;
    let zeroVariance = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[idx.status] || '';

      // Only include completed counts
      if (status !== 'Completed') continue;

      // Check date range
      const countedAt = row[idx.countedAt];
      if (!countedAt) continue;

      const countDate = new Date(countedAt);
      if (isNaN(countDate.getTime())) continue;

      if (startDate && countDate < startDate) continue;
      if (endDate && countDate > endDate) continue;

      const binCode = row[idx.binCode] || '';
      const fbpn = row[idx.fbpn] || '';
      const manufacturer = row[idx.manufacturer] || '';
      const project = row[idx.project] || '';
      const currentQty = parseFloat(row[idx.currentQty]) || 0;
      const countedQty = parseFloat(row[idx.countedQty]) || 0;
      const variance = parseFloat(row[idx.variance]) || 0;
      const batchId = row[idx.batchId] || '';
      const notes = row[idx.notes] || '';
      const countedBy = row[idx.countedBy] || '';

      // Track variance statistics
      totalVariance += variance;
      if (variance > 0) positiveVariance += variance;
      else if (variance < 0) negativeVariance += Math.abs(variance);
      else zeroVariance++;

      // Track unique bins
      binsCounted.add(`${binCode}|${fbpn}|${manufacturer}|${project}`);

      // Track batch summary
      if (!batchSummary[batchId]) {
        batchSummary[batchId] = {
          batchId: batchId,
          totalLines: 0,
          completedLines: 0,
          discrepancies: 0,
          totalVariance: 0
        };
      }
      batchSummary[batchId].totalLines++;
      batchSummary[batchId].completedLines++;
      if (variance !== 0) batchSummary[batchId].discrepancies++;
      batchSummary[batchId].totalVariance += variance;

      details.push({
        batchId: batchId,
        binCode: binCode,
        fbpn: fbpn,
        manufacturer: manufacturer,
        project: project,
        systemQty: currentQty,
        countedQty: countedQty,
        variance: variance,
        hasDiscrepancy: variance !== 0,
        notes: notes,
        countedAt: Utilities.formatDate(countDate, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm'),
        countedBy: countedBy
      });
    }

    // Sort details by date descending
    details.sort((a, b) => {
      const da = new Date(a.countedAt);
      const db = new Date(b.countedAt);
      return db - da;
    });

    const discrepancies = details.filter(d => d.hasDiscrepancy);

    const summary = {
      totalBinsAudited: binsCounted.size,
      totalLinesCompleted: details.length,
      totalDiscrepancies: discrepancies.length,
      discrepancyRate: details.length > 0 ? ((discrepancies.length / details.length) * 100).toFixed(1) : 0,
      totalVariance: totalVariance,
      positiveVariance: positiveVariance,
      negativeVariance: negativeVariance,
      accurateCount: zeroVariance,
      batchCount: Object.keys(batchSummary).length,
      batches: Object.values(batchSummary)
    };

    return {
      success: true,
      summary: summary,
      details: details,
      discrepancies: discrepancies,
      dateRange: {
        start: startDate ? Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'MM/dd/yyyy') : '',
        end: endDate ? Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'MM/dd/yyyy') : ''
      }
    };

  } catch (err) {
    Logger.log('imsGetCycleCountReport error: ' + err.toString());
    return { success: false, message: 'Error generating report: ' + err.message };
  }
}

/**
 * Returns an empty summary structure.
 */
function getEmptySummary_() {
  return {
    totalBinsAudited: 0,
    totalLinesCompleted: 0,
    totalDiscrepancies: 0,
    discrepancyRate: 0,
    totalVariance: 0,
    positiveVariance: 0,
    negativeVariance: 0,
    accurateCount: 0,
    batchCount: 0,
    batches: []
  };
}

/**
 * Generates a PDF report for cycle count data.
 * @param {Object} params - { startDate, endDate }
 * @returns {Object} { success, pdfUrl, message }
 */
function imsGenerateCycleCountReportPdf(params) {
  try {
    const reportData = imsGetCycleCountReport(params);
    if (!reportData.success) {
      return reportData;
    }

    const html = buildCycleCountReportHtml_(reportData);

    // Create temp HTML file and convert to PDF
    const tempFolder = DriveApp.getRootFolder();
    const fileName = `Cycle_Count_Report_${reportData.dateRange.start}_to_${reportData.dateRange.end}`.replace(/\//g, '-');

    const htmlBlob = Utilities.newBlob(html, 'text/html', fileName + '.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName(fileName + '.pdf');

    const pdfFile = tempFolder.createFile(pdfBlob);

    return {
      success: true,
      pdfUrl: pdfFile.getUrl(),
      message: 'Report generated successfully.'
    };

  } catch (err) {
    Logger.log('imsGenerateCycleCountReportPdf error: ' + err.toString());
    return { success: false, message: 'Error generating PDF: ' + err.message };
  }
}

/**
 * Builds HTML for the cycle count report.
 */
function buildCycleCountReportHtml_(reportData) {
  const s = reportData.summary;
  const dateRange = reportData.dateRange;

  let discrepancyRows = '';
  if (reportData.discrepancies && reportData.discrepancies.length > 0) {
    reportData.discrepancies.forEach(d => {
      const varClass = d.variance > 0 ? 'positive' : 'negative';
      const varSign = d.variance > 0 ? '+' : '';
      discrepancyRows += `
        <tr>
          <td>${d.binCode}</td>
          <td>${d.fbpn}</td>
          <td>${d.manufacturer}</td>
          <td>${d.project}</td>
          <td style="text-align:right">${d.systemQty}</td>
          <td style="text-align:right">${d.countedQty}</td>
          <td style="text-align:right" class="${varClass}">${varSign}${d.variance}</td>
          <td>${d.notes || '-'}</td>
          <td>${d.countedAt}</td>
          <td>${d.countedBy}</td>
        </tr>`;
    });
  } else {
    discrepancyRows = '<tr><td colspan="10" style="text-align:center; color:#888;">No discrepancies found in this period.</td></tr>';
  }

  let allAuditsRows = '';
  if (reportData.details && reportData.details.length > 0) {
    reportData.details.forEach(d => {
      const varClass = d.variance > 0 ? 'positive' : (d.variance < 0 ? 'negative' : 'zero');
      const varSign = d.variance > 0 ? '+' : '';
      allAuditsRows += `
        <tr>
          <td>${d.binCode}</td>
          <td>${d.fbpn}</td>
          <td>${d.manufacturer}</td>
          <td>${d.project}</td>
          <td style="text-align:right">${d.systemQty}</td>
          <td style="text-align:right">${d.countedQty}</td>
          <td style="text-align:right" class="${varClass}">${varSign}${d.variance}</td>
          <td>${d.countedAt}</td>
        </tr>`;
    });
  } else {
    allAuditsRows = '<tr><td colspan="8" style="text-align:center; color:#888;">No audits completed in this period.</td></tr>';
  }

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Cycle Count Report</title>
  <style>
    @page { size: landscape; margin: 0.5in; }
    body { font-family: Arial, sans-serif; font-size: 10pt; color: #333; margin: 0; padding: 20px; }
    h1 { font-size: 18pt; color: #1a1a1a; margin-bottom: 5px; }
    h2 { font-size: 14pt; color: #333; margin-top: 25px; margin-bottom: 10px; border-bottom: 2px solid #2563eb; padding-bottom: 5px; }
    .date-range { color: #666; font-size: 11pt; margin-bottom: 20px; }
    .summary-grid { display: flex; flex-wrap: wrap; gap: 15px; margin-bottom: 25px; }
    .summary-card { background: #f8f9fa; border: 1px solid #e0e0e0; border-radius: 6px; padding: 12px 16px; min-width: 140px; }
    .summary-card .label { font-size: 9pt; color: #666; text-transform: uppercase; margin-bottom: 4px; }
    .summary-card .value { font-size: 18pt; font-weight: 700; color: #1a1a1a; }
    .summary-card .value.positive { color: #16a34a; }
    .summary-card .value.negative { color: #dc2626; }
    .summary-card .value.warning { color: #d97706; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
    th { background: #2563eb; color: white; padding: 8px 10px; text-align: left; font-size: 9pt; text-transform: uppercase; }
    td { padding: 6px 10px; border-bottom: 1px solid #e0e0e0; font-size: 9pt; }
    tr:nth-child(even) { background: #f8f9fa; }
    .positive { color: #16a34a; font-weight: 600; }
    .negative { color: #dc2626; font-weight: 600; }
    .zero { color: #666; }
    .page-break { page-break-before: always; }
    .footer { margin-top: 30px; font-size: 8pt; color: #999; text-align: center; }
  </style>
</head>
<body>
  <h1>Cycle Count Weekly Report</h1>
  <div class="date-range">Report Period: ${dateRange.start} to ${dateRange.end}</div>

  <div class="summary-grid">
    <div class="summary-card">
      <div class="label">Bins Audited</div>
      <div class="value">${s.totalBinsAudited}</div>
    </div>
    <div class="summary-card">
      <div class="label">Lines Completed</div>
      <div class="value">${s.totalLinesCompleted}</div>
    </div>
    <div class="summary-card">
      <div class="label">Discrepancies</div>
      <div class="value ${s.totalDiscrepancies > 0 ? 'warning' : ''}">${s.totalDiscrepancies}</div>
    </div>
    <div class="summary-card">
      <div class="label">Accuracy Rate</div>
      <div class="value ${parseFloat(s.discrepancyRate) > 10 ? 'negative' : 'positive'}">${(100 - parseFloat(s.discrepancyRate)).toFixed(1)}%</div>
    </div>
    <div class="summary-card">
      <div class="label">Net Variance</div>
      <div class="value ${s.totalVariance > 0 ? 'positive' : (s.totalVariance < 0 ? 'negative' : '')}">${s.totalVariance > 0 ? '+' : ''}${s.totalVariance}</div>
    </div>
    <div class="summary-card">
      <div class="label">Over (+)</div>
      <div class="value positive">+${s.positiveVariance}</div>
    </div>
    <div class="summary-card">
      <div class="label">Short (-)</div>
      <div class="value negative">-${s.negativeVariance}</div>
    </div>
    <div class="summary-card">
      <div class="label">Batches</div>
      <div class="value">${s.batchCount}</div>
    </div>
  </div>

  <h2>Discrepancies</h2>
  <table>
    <thead>
      <tr>
        <th>Bin</th>
        <th>FBPN</th>
        <th>Manufacturer</th>
        <th>Project</th>
        <th>System Qty</th>
        <th>Counted Qty</th>
        <th>Variance</th>
        <th>Notes</th>
        <th>Counted At</th>
        <th>Counted By</th>
      </tr>
    </thead>
    <tbody>
      ${discrepancyRows}
    </tbody>
  </table>

  <div class="page-break"></div>

  <h2>All Audits Completed</h2>
  <table>
    <thead>
      <tr>
        <th>Bin</th>
        <th>FBPN</th>
        <th>Manufacturer</th>
        <th>Project</th>
        <th>System Qty</th>
        <th>Counted Qty</th>
        <th>Variance</th>
        <th>Counted At</th>
      </tr>
    </thead>
    <tbody>
      ${allAuditsRows}
    </tbody>
  </table>

  <div class="footer">
    Generated on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss')}
  </div>
</body>
</html>`;
}

