function fillBinStockProjectFromPush(force) {
  force = !!force;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet name fallbacks (use your TABS constants if present in your project)
  const shBin = ss.getSheetByName((typeof TABS !== 'undefined' && TABS.BIN_STOCK) ? TABS.BIN_STOCK : 'Bin_Stock');
  const shMaster = ss.getSheetByName((typeof TABS !== 'undefined' && TABS.MASTER_LOG) ? TABS.MASTER_LOG : 'Master_Log');
  const shPO = ss.getSheetByName((typeof TABS !== 'undefined' && TABS.PO_MASTER) ? TABS.PO_MASTER : 'PO_Master');

  if (!shBin) throw new Error('Bin_Stock sheet not found.');
  if (!shMaster) throw new Error('Master_Log sheet not found.');
  if (!shPO) throw new Error('PO_Master sheet not found.');

  const norm = (v) => (v === null || v === undefined) ? '' : String(v).trim();

  // --- Build Push -> PO map from Master_Log (E=5, K=11) ---
  const mVals = shMaster.getDataRange().getValues();
  const pushToPO = new Map();

  for (let r = 1; r < mVals.length; r++) {
    const push = norm(mVals[r][4]);   // col E
    const po = norm(mVals[r][10]);    // col K
    if (!push || !po) continue;
    // First match wins (keeps stable). Change to always-set if you prefer last match.
    if (!pushToPO.has(push)) pushToPO.set(push, po);
  }

  // --- Build PO -> Project map from PO_Master (A=1, B=2) ---
  const pVals = shPO.getDataRange().getValues();
  const poToProject = new Map();

  for (let r = 1; r < pVals.length; r++) {
    const po = norm(pVals[r][0]);       // col A
    const project = norm(pVals[r][1]);  // col B
    if (!po || !project) continue;
    if (!poToProject.has(po)) poToProject.set(po, project);
  }

  // --- Read Bin_Stock and compute writes (C=3, F=6) ---
  const binRange = shBin.getDataRange();
  const bVals = binRange.getValues();

  let updates = 0;
  for (let r = 1; r < bVals.length; r++) {
    const push = norm(bVals[r][2]);    // col C
    const existingProject = norm(bVals[r][5]); // col F

    if (!push) continue;
    if (!force && existingProject) continue;

    const po = pushToPO.get(push);
    if (!po) continue;

    const project = poToProject.get(po);
    if (!project) continue;

    bVals[r][5] = project; // write to col F
    updates++;
  }

  if (updates) {
    binRange.setValues(bVals);
  }

  Logger.log(`fillBinStockProjectFromPush: updated ${updates} row(s).`);
}

/** Convenience menu-safe wrapper (does not overwrite existing Project cells). */
function fillBinStockProjectFromPush_noOverwrite() {
  fillBinStockProjectFromPush(false);
}

/** Convenience wrapper to force overwrite Bin_Stock col F. */
function fillBinStockProjectFromPush_forceOverwrite() {
  fillBinStockProjectFromPush(true);
}
function generateBinConsolidationSuggestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh =
    ss.getSheetByName((typeof TABS !== 'undefined' && TABS.BIN_STOCK) ? TABS.BIN_STOCK : 'Bin_Stock');
  if (!sh) throw new Error('Bin_Stock sheet not found.');

  const outName = 'ConsolidationSuggestions';
  const out = ss.getSheetByName(outName) || ss.insertSheet(outName);
  out.clearContents();

  const norm = (v) => (v === null || v === undefined) ? '' : String(v).trim();
  const toNum = (v) => {
    if (v === null || v === undefined || v === '') return NaN;
    if (typeof v === 'number') return v;
    const s = String(v).trim().replace('%', '');
    const n = Number(s);
    return Number.isFinite(n) ? n : NaN;
  };

  // ✅ Block ANY bin where Project contains "NAB" anywhere (case-insensitive)
  const isNABProject = (projectVal) => norm(projectVal).toUpperCase().includes('NAB');

  function binSortKey(binCode) {
    const b = norm(binCode);
    const letter = (b.match(/[A-Za-z]/) || ['~'])[0].toUpperCase();
    const nums = b.match(/\d+/g);
    const lastNum = nums && nums.length ? Number(nums[nums.length - 1]) : Number.POSITIVE_INFINITY;
    return { letter, lastNum, raw: b };
  }
  function cmpBin(a, b) {
    const ka = binSortKey(a);
    const kb = binSortKey(b);
    if (ka.letter !== kb.letter) return ka.letter < kb.letter ? -1 : 1;
    if (ka.lastNum !== kb.lastNum) return ka.lastNum - kb.lastNum;
    if (ka.raw !== kb.raw) return ka.raw < kb.raw ? -1 : 1;
    return 0;
  }

  const values = sh.getDataRange().getValues();
  const header = [[
    'SKU',
    'FBPN',
    'Bin_To_Empty',
    'Qty_In_Donor_Bin',
    'Bin_To_Fill',
    'Original_Qty',
    'Qty_To_Move',
    'Final_Qty_In_Bin'
  ]];

  if (values.length < 2) {
    out.getRange(1, 1, 1, 8).setValues(header);
    return;
  }

  // Bin_Stock columns (0-based)
  const COL_BIN = 0;      // A Bin_Code
  const COL_FBPN = 3;     // D FBPN
  const COL_PROJECT = 5;  // F Project
  const COL_QTY = 7;      // H Current_Quantity
  const COL_PCT = 8;      // I Stock_Percentage
  const COL_SKU = 11;     // L SKU

  // BIN total % = sum of % across ALL rows in that bin
  const binPctSum = new Map();
  for (let r = 1; r < values.length; r++) {
    const bin = norm(values[r][COL_BIN]);
    if (!bin) continue;
    const pct = toNum(values[r][COL_PCT]);
    if (!Number.isFinite(pct) || pct <= 0) continue;
    binPctSum.set(bin, (binPctSum.get(bin) || 0) + pct);
  }

  // SKU -> rows
  const rowsBySku = new Map();
  for (let r = 1; r < values.length; r++) {
    const sku = norm(values[r][COL_SKU]);
    if (!sku) continue;

    const qtyN = toNum(values[r][COL_QTY]);
    const pctN = toNum(values[r][COL_PCT]);

    const rowObj = {
      row: r + 1,
      bin: norm(values[r][COL_BIN]),
      project: norm(values[r][COL_PROJECT]),
      sku,
      fbpn: norm(values[r][COL_FBPN]) || '',
      qty: Number.isFinite(qtyN) ? qtyN : 0,
      pct: Number.isFinite(pctN) ? pctN : NaN
    };

    if (!rowsBySku.has(sku)) rowsBySku.set(sku, []);
    rowsBySku.get(sku).push(rowObj);
  }

  function estimateMaxCap(qty, pct) {
    if (!Number.isFinite(qty) || qty <= 0) return NaN;
    if (!Number.isFinite(pct) || pct <= 0) return NaN;
    return qty * (100 / pct);
  }

  // Track bin-level % added by our proposed moves (across all SKUs)
  const binPctAdded = new Map(); // bin -> pctAdded
  // Track qty moved into (SKU||Bin_To_Fill) so Final_Qty is consistent across multiple donor rows
  const movedQtyBySkuBin = new Map(); // key -> movedQty

  const rowsOut = [];

  for (const [sku, skuRows] of rowsBySku.entries()) {
    // ✅ Do not move INTO any bin whose Project contains NAB
    const targets = skuRows.filter(x =>
      Number.isFinite(x.pct) &&
      x.pct < 100 &&
      x.qty > 0 &&
      x.bin &&
      !isNABProject(x.project)
    );

    // ✅ Do not PULL FROM/empty any bin whose Project contains NAB
    const donorsAll = skuRows.filter(x =>
      x.qty > 0 &&
      x.bin &&
      !isNABProject(x.project)
    );

    if (!targets.length || !donorsAll.length) continue;

    targets.sort((a, b) => cmpBin(a.bin, b.bin));

    for (const t of targets) {
      const originalQty = t.qty;
      const maxCap = estimateMaxCap(originalQty, t.pct);
      if (!Number.isFinite(maxCap) || maxCap <= 0) continue;

      const remainingRowUnits0 = (maxCap - originalQty);
      if (!Number.isFinite(remainingRowUnits0) || remainingRowUnits0 <= 0) continue;

      const skuBinKey = `${sku}||${t.bin}`;
      let remainingRowUnits = remainingRowUnits0 - (movedQtyBySkuBin.get(skuBinKey) || 0);
      if (remainingRowUnits <= 0) continue;

      const donorList = donorsAll
        .filter(d => d.row !== t.row && d.qty > 0)
        .sort((a, b) => a.qty - b.qty);

      for (const d of donorList) {
        if (remainingRowUnits <= 0) break;

        // FULL donor move only + must fit row-level remaining
        if (d.qty > remainingRowUnits) continue;

        // Convert move to % increase for this target row (used as bin-level increment proxy)
        const pctInc = (d.qty / maxCap) * 100;

        // Bin-level constraint: sum(all row %) + planned% + this% <= 100
        const binCurrentPct = binPctSum.get(t.bin) || 0;
        const plannedPct = binPctAdded.get(t.bin) || 0;
        if ((binCurrentPct + plannedPct + pctInc) > 100.000001) continue;

        // Accept move
        const nextMovedQty = (movedQtyBySkuBin.get(skuBinKey) || 0) + d.qty;
        movedQtyBySkuBin.set(skuBinKey, nextMovedQty);
        binPctAdded.set(t.bin, plannedPct + pctInc);
        remainingRowUnits -= d.qty;

        const finalQty = originalQty + nextMovedQty;

        // Use FBPN from donor row (what you're physically moving); fallback to target
        const fbpn = d.fbpn || t.fbpn || '';

        rowsOut.push([
          sku,
          fbpn,
          d.bin,      // Bin_To_Empty
          d.qty,      // Qty_In_Donor_Bin
          t.bin,      // Bin_To_Fill
          originalQty,
          d.qty,      // Qty_To_Move
          finalQty
        ]);
      }
    }
  }

  // Sort by Bin_To_Empty
  rowsOut.sort((r1, r2) => {
    // [0 SKU,1 FBPN,2 Empty,3 DonorQty,4 Fill,5 Orig,6 Move,7 Final]
    let c = cmpBin(r1[2], r2[2]);
    if (c) return c;
    c = cmpBin(r1[4], r2[4]);
    if (c) return c;
    if (r1[0] !== r2[0]) return r1[0] < r2[0] ? -1 : 1;
    return 0;
  });

  const output = header.concat(rowsOut);
  out.getRange(1, 1, output.length, 8).setValues(output);
  out.autoResizeColumns(1, 8);
}


function generateConsolidationPDF_FromTemplate() {
  const TEMPLATE_ID = '1iG9Mny3sHge8Fa14BmOgw8mgo09S9M4Fid2U4c8z9sA';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('ConsolidationSuggestions');
  if (!sh) throw new Error('ConsolidationSuggestions sheet not found.');

  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error('No consolidation rows found.');

  // ConsolidationSuggestions (0-based):
  // A SKU, B FBPN, C Bin_To_Empty, D Qty_In_Donor_Bin, E Bin_To_Fill, F Original_Qty, G Qty_To_Move, H Final_Qty_In_Bin
  const COL_FBPN = 1;        // B
  const COL_BIN_EMPTY = 2;   // C
  const COL_BIN_FILL = 4;    // E
  const COL_QTY_MOVE = 6;    // G

  const norm = (v) => (v === null || v === undefined) ? '' : String(v).trim();

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const fbpn = norm(r[COL_FBPN]);
    const emptyBin = norm(r[COL_BIN_EMPTY]);
    const fillBin = norm(r[COL_BIN_FILL]);
    const qtyMove = norm(r[COL_QTY_MOVE]);
    if (!fbpn || !emptyBin || !fillBin || !qtyMove) continue;

    rows.push({
      Bin_To_Empty: emptyBin,
      FBPN: fbpn,
      Bin_To_Fill: fillBin,
      Qty_To_Move: qtyMove
    });
  }
  if (!rows.length) throw new Error('No valid rows to print.');

  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');

  // Copy template
  const copyName = `Bin Consolidation Log - ${Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HHmm')}`;
  const docId = DriveApp.getFileById(TEMPLATE_ID).makeCopy(copyName).getId();
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  // Replace Date placeholder globally
  body.replaceText(escapeForDocRegex_('{{Date}}'), dateStr);

  // Find placeholder row
  const info = findPlaceholderRow_(body, [
    '{{Bin_To_Empty}}',
    '{{FBPN}}',
    '{{Bin_To_Fill}}',
    '{{Qty_To_Move}}'
  ]);
  if (!info) {
    doc.saveAndClose();
    throw new Error('Template row with placeholders not found.');
  }

  const table = info.table;
  const templateRowIndex = info.rowIndex;
  const templateRow = table.getRow(templateRowIndex);

  // ✅ Pristine prototype row (keeps row height + full styling)
  const protoRow = templateRow.copy();

  // Fill first row (the real template row)
  fillRowPlaceholders_(templateRow, rows[0]);

  // Insert remaining rows, each as a true copy of the prototype row
  let insertAt = templateRowIndex;
  for (let i = 1; i < rows.length; i++) {
    insertAt++;

    const newRow = table.insertTableRow(insertAt, protoRow.copy()); // ✅ keeps height/formatting
    fillRowPlaceholders_(newRow, rows[i]);
  }

  doc.saveAndClose();

  // Export to PDF
  const pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF).setName(copyName + '.pdf');
  const pdfFile = DriveApp.createFile(pdfBlob);

  Logger.log('PDF created: ' + pdfFile.getUrl());
  return pdfFile.getUrl();
}

// ---------- helpers ----------

function findPlaceholderRow_(body, requiredTokens) {
  const n = body.getNumChildren();
  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;

    const table = el.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      const txt = table.getRow(r).getText();
      let ok = true;
      for (const t of requiredTokens) {
        if (txt.indexOf(t) === -1) { ok = false; break; }
      }
      if (ok) return { table, rowIndex: r };
    }
  }
  return null;
}

function fillRowPlaceholders_(row, map) {
  const repl = [
    ['{{Bin_To_Empty}}', map.Bin_To_Empty],
    ['{{FBPN}}', map.FBPN],
    ['{{Bin_To_Fill}}', map.Bin_To_Fill],
    ['{{Qty_To_Move}}', map.Qty_To_Move]
  ];
  for (let c = 0; c < row.getNumCells(); c++) {
    const cell = row.getCell(c);
    for (const [token, val] of repl) {
      cell.replaceText(escapeForDocRegex_(token), String(val));
    }
  }
}

function escapeForDocRegex_(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
/**
 * Generates a PDF of unique Bin_Codes (one row per bin) where Project contains "NAB" (case-insensitive).
 * Even if a bin has multiple FBPN rows, it appears ONCE (whole skid pulled).
 *
 * Sort order: by first letter in Bin_Code, then by the last number in the Bin_Code.
 *
 * Template ID:
 *  1vj-1ViazIpr92WTlAK8dYM3WgYzzqVnfeQkz4QYMF0Q
 *
 * Placeholders:
 *  {{Date}}, {{Bin_To_Pull}}, {{FBPN}}, {{Project}}, {{Current_Qty}}
 */
function generateNABBinsPDF_FromTemplate() {
  const TEMPLATE_ID = '1vj-1ViazIpr92WTlAK8dYM3WgYzzqVnfeQkz4QYMF0Q';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh =
    ss.getSheetByName((typeof TABS !== 'undefined' && TABS.BIN_STOCK) ? TABS.BIN_STOCK : 'Bin_Stock');
  if (!sh) throw new Error('Bin_Stock sheet not found.');

  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error('Bin_Stock has no rows.');

  const norm = (v) => (v === null || v === undefined) ? '' : String(v).trim();
  const hasNAB = (projectVal) => norm(projectVal).toUpperCase().includes('NAB');

  // Bin_Stock columns (0-based)
  const COL_BIN = 0;     // A Bin_Code
  const COL_FBPN = 3;    // D FBPN
  const COL_PROJECT = 5; // F Project
  const COL_CUR_QTY = 7; // H Current_Quantity

  // Bin -> aggregate (one row per bin)
  const byBin = new Map();

  for (let i = 1; i < data.length; i++) {
    const r = data[i];

    const bin = norm(r[COL_BIN]);
    if (!bin) continue;

    const project = norm(r[COL_PROJECT]);
    if (!hasNAB(project)) continue;

    const fbpn = norm(r[COL_FBPN]);
    const qtyNum = Number(norm(r[COL_CUR_QTY]).replace(/,/g, '')) || 0;

    if (!byBin.has(bin)) {
      byBin.set(bin, {
        Bin_To_Pull: bin,
        FBPN: fbpn || '',
        Project: project || '',
        Current_Qty: qtyNum
      });
    } else {
      // If multiple rows per bin, aggregate qty; keep first non-empty FBPN/project
      const agg = byBin.get(bin);
      agg.Current_Qty += qtyNum;
      if (!agg.FBPN && fbpn) agg.FBPN = fbpn;
      if (!agg.Project && project) agg.Project = project;
      byBin.set(bin, agg);
    }
  }

  const rows = Array.from(byBin.values());
  if (!rows.length) throw new Error('No Bin_Stock bins found with Project containing "NAB".');

  // Sort: by first letter, then last number, then raw string
  function binSortKey(binCode) {
    const b = norm(binCode);
    const letter = (b.match(/[A-Za-z]/) || ['~'])[0].toUpperCase();
    const nums = b.match(/\d+/g);
    const lastNum = nums && nums.length ? Number(nums[nums.length - 1]) : Number.POSITIVE_INFINITY;
    return { letter, lastNum, raw: b };
  }
  rows.sort((a, b) => {
    const ka = binSortKey(a.Bin_To_Pull);
    const kb = binSortKey(b.Bin_To_Pull);
    if (ka.letter !== kb.letter) return ka.letter < kb.letter ? -1 : 1;
    if (ka.lastNum !== kb.lastNum) return ka.lastNum - kb.lastNum;
    if (ka.raw !== kb.raw) return ka.raw < kb.raw ? -1 : 1;
    return 0;
  });

  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');

  // Copy template
  const copyName = `NAB_Bins_${Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm')}`;
  const docId = DriveApp.getFileById(TEMPLATE_ID).makeCopy(copyName).getId();
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  // Replace Date globally
  body.replaceText(escapeForDocRegex_('{{Date}}'), dateStr);

  // Find placeholder row
  const info = findPlaceholderRow_(body, [
    '{{Bin_To_Pull}}',
    '{{FBPN}}',
    '{{Project}}',
    '{{Current_Qty}}'
  ]);
  if (!info) {
    doc.saveAndClose();
    throw new Error('Template row with placeholders not found ({{Bin_To_Pull}}, {{FBPN}}, {{Project}}, {{Current_Qty}}).');
  }

  const table = info.table;
  const templateRowIndex = info.rowIndex;
  const templateRow = table.getRow(templateRowIndex);

  // Pristine prototype row (keeps row height + styling)
  const protoRow = templateRow.copy();

  // Fill first row
  fillNABRowPlaceholders_(templateRow, rows[0]);

  // Insert remaining rows with same formatting/height
  let insertAt = templateRowIndex;
  for (let i = 1; i < rows.length; i++) {
    insertAt++;
    const newRow = table.insertTableRow(insertAt, protoRow.copy());
    fillNABRowPlaceholders_(newRow, rows[i]);
  }

  doc.saveAndClose();

  // Export to PDF
  const pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF).setName(copyName + '.pdf');
  const pdfFile = DriveApp.createFile(pdfBlob);

  Logger.log('PDF created: ' + pdfFile.getUrl());
  return pdfFile.getUrl();
}

// ---------------- helpers ----------------

function findPlaceholderRow_(body, requiredTokens) {
  const n = body.getNumChildren();
  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;

    const table = el.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      const txt = table.getRow(r).getText();
      let ok = true;
      for (const t of requiredTokens) {
        if (txt.indexOf(t) === -1) { ok = false; break; }
      }
      if (ok) return { table, rowIndex: r };
    }
  }
  return null;
}

function fillNABRowPlaceholders_(row, map) {
  const repl = [
    ['{{Bin_To_Pull}}', map.Bin_To_Pull],
    ['{{FBPN}}', map.FBPN],
    ['{{Project}}', map.Project],
    ['{{Current_Qty}}', String(map.Current_Qty)]
  ];

  for (let c = 0; c < row.getNumCells(); c++) {
    const cell = row.getCell(c);
    for (const [token, val] of repl) {
      cell.replaceText(escapeForDocRegex_(token), String(val));
    }
  }
}

function escapeForDocRegex_(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
