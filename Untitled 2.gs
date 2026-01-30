// Debug function removed - was used for library inspection
// All functions are now in the same project
/**
 * Writes skids that exist in BOTH Bin_Stock and Inbound_Staging
 * to a new tab: "Skids_In_Both"
 *
 * Match key: SKU + Initial_Quantity + Push_Number
 *
 * Output columns include both Inbound row + Bin_Stock row details.
 */
function writeSkidsInBothBinStockAndInboundStagingToNewTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const INBOUND_SHEET = 'Inbound_Staging';
  const STOCK_SHEET = 'Bin_Stock';
  const OUT_SHEET = 'Skids_In_Both';

  const inboundSh = ss.getSheetByName(INBOUND_SHEET);
  const stockSh = ss.getSheetByName(STOCK_SHEET);
  if (!inboundSh) throw new Error('Missing sheet: ' + INBOUND_SHEET);
  if (!stockSh) throw new Error('Missing sheet: ' + STOCK_SHEET);

  const inbound = _readSheetAsObjectsStrict_(inboundSh);
  const stock = _readSheetAsObjectsStrict_(stockSh);

  // Build key -> stock rows map
  const stockMap = new Map();
  for (const r of stock.rows) {
    const key = _buildSkuInitQtyPushKey_(r);
    if (!key) continue;
    if (!stockMap.has(key)) stockMap.set(key, []);
    stockMap.get(key).push(r);
  }

  // Prepare output rows (one per matched pair; if multiple stock rows share key, we output one row per stock row)
  const out = [];
  for (const ir of inbound.rows) {
    const key = _buildSkuInitQtyPushKey_(ir);
    if (!key) continue;

    const matches = stockMap.get(key);
    if (!matches || !matches.length) continue;

    for (const sr of matches) {
      out.push(_buildOutputRow_(key, ir, sr));
    }
  }

  // Create/clear output sheet
  let outSh = ss.getSheetByName(OUT_SHEET);
  if (!outSh) outSh = ss.insertSheet(OUT_SHEET);
  outSh.clearContents();

  const headers = _outputHeaders_();
  outSh.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (out.length) {
    outSh.getRange(2, 1, out.length, headers.length).setValues(out);
  }

  outSh.autoResizeColumns(1, headers.length);
  outSh.setFrozenRows(1);

  return { ok: true, sheet: OUT_SHEET, matches: out.length };
}

// ==============================
// Internals
// ==============================

function _outputHeaders_() {
  return [
    'Match_Key',

    // Inbound_Staging
    'Inbound_Row',
    'Inbound_Bin_Code',
    'Inbound_Bin_Name',
    'Inbound_Push_Number',
    'Inbound_FBPN',
    'Inbound_Manufacturer',
    'Inbound_Project',
    'Inbound_Initial_Quantity',
    'Inbound_Current_Quantity',
    'Inbound_Stock_Percentage',
    'Inbound_AUDIT_NEEDED',
    'Inbound_Skid_ID',
    'Inbound_SKU',

    // Bin_Stock
    'BinStock_Row',
    'BinStock_Bin_Code',
    'BinStock_Bin_Name',
    'BinStock_Push_Number',
    'BinStock_FBPN',
    'BinStock_Manufacturer',
    'BinStock_Project',
    'BinStock_Initial_Quantity',
    'BinStock_Current_Quantity',
    'BinStock_Stock_Percentage',
    'BinStock_AUDIT_NEEDED',
    'BinStock_Skid_ID',
    'BinStock_SKU'
  ];
}

function _buildOutputRow_(key, inbound, stock) {
  return [
    key,

    // Inbound_Staging
    inbound.__rowNumber || '',
    inbound.Bin_Code ?? '',
    inbound.Bin_Name ?? '',
    inbound.Push_Number ?? '',
    inbound.FBPN ?? '',
    inbound.Manufacturer ?? '',
    inbound.Project ?? '',
    inbound.Initial_Quantity ?? '',
    inbound.Current_Quantity ?? '',
    inbound.Stock_Percentage ?? '',
    inbound['AUDIT NEEDED'] ?? '',
    inbound.Skid_ID ?? '',
    inbound.SKU ?? '',

    // Bin_Stock
    stock.__rowNumber || '',
    stock.Bin_Code ?? '',
    stock.Bin_Name ?? '',
    stock.Push_Number ?? '',
    stock.FBPN ?? '',
    stock.Manufacturer ?? '',
    stock.Project ?? '',
    stock.Initial_Quantity ?? '',
    stock.Current_Quantity ?? '',
    stock.Stock_Percentage ?? '',
    stock['AUDIT NEEDED'] ?? '',
    stock.Skid_ID ?? '',
    stock.SKU ?? ''
  ];
}

function _readSheetAsObjectsStrict_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { headers: [], rows: [] };

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => String(h || '').trim());

  const headerIndex = {};
  headers.forEach((h, i) => { if (h) headerIndex[h] = i; });

  // Require match fields
  ['SKU', 'Initial_Quantity', 'Push_Number'].forEach(req => {
    if (!(req in headerIndex)) throw new Error(`Sheet "${sheet.getName()}" missing required header: ${req}`);
  });

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(v => v === '' || v === null)) continue;

    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c];
      if (!h) continue;
      obj[h] = row[c];
    }
    obj.__rowNumber = r + 1;
    rows.push(obj);
  }

  return { headers, rows };
}

function _buildSkuInitQtyPushKey_(rowObj) {
  if (!rowObj) return '';
  const sku = String(rowObj.SKU ?? '').trim().toUpperCase();
  const push = String(rowObj.Push_Number ?? '').trim();
  const initQty = _normalizeNumber_(rowObj.Initial_Quantity);

  if (!sku || !push || initQty === '') return '';
  return `${sku}||${push}||${initQty}`;
}

function _normalizeNumber_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return String(v);
  const s = String(v).trim();
  if (!s) return '';
  const n = Number(s.replace(/,/g, ''));
  if (Number.isNaN(n)) return s;
  if (Math.floor(n) === n) return String(n);
  return String(n);
}
