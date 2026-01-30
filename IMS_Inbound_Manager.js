

/**
 * Searches for an Inbound Transaction by BOL Number
 */
function searchInboundByBOL(bolNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(TABS.MASTER_LOG);
  if (!masterSheet) throw new Error('Master_Log not found');

  const data = masterSheet.getDataRange().getValues();
  const headers = data[0];
  const idxBol = headers.indexOf('BOL_Number');
  const idxTxn = headers.indexOf('Txn_ID');
  const idxDate = headers.indexOf('Date_Received');
  const idxVendor = headers.indexOf('Manufacturer');
  const idxSkidCount = headers.indexOf('Total_Skid_Count');

  if (idxBol === -1) throw new Error('BOL_Number column not found');

  const bolSearch = String(bolNumber || '').trim().toUpperCase();
  // Find all transactions matching BOL (could be split, but usually one)
  // We return the unique TxnIDs found
  const transactions = {};

  for (let i = 1; i < data.length; i++) {
    const rowBol = String(data[i][idxBol] || '').trim().toUpperCase();
    if (rowBol === bolSearch) {
      const txnId = data[i][idxTxn];
      if (!transactions[txnId]) {
        transactions[txnId] = {
          txnId: txnId,
          date: formatDate(data[i][idxDate]),
          manufacturer: data[i][idxVendor],
          totalSkids: data[i][idxSkidCount]
        };
      }
    }
  }

  return Object.values(transactions);
}

/**
 * Wrapper for the undo function ensuring it connects to the BOL UI
 */
function executeUndoByTxnId(txnId) {
  // Call the existing undo function from IMS_Inbound_FIXED5.js
  if (typeof undoInboundSubmission === 'function') {
    return undoInboundSubmission(txnId);
  } else {
    throw new Error('undoInboundSubmission function not found. Please ensure IMS_Inbound is loaded.');
  }
}

/**
 * Fetches the last N inbound transactions for the Reprint UI
 */
function getRecentInboundTransactions(limit = 20) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(TABS.MASTER_LOG);
  if (!masterSheet) return [];

  const data = masterSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Indices
  const idxTxn = headers.indexOf('Txn_ID');
  const idxDate = headers.indexOf('Date_Received');
  const idxBol = headers.indexOf('BOL_Number');
  const idxMan = headers.indexOf('Manufacturer');
  const idxPush = headers.indexOf('Push #');

  const seenTxns = new Set();
  const recent = [];

  // Loop backwards from end
  for (let i = data.length - 1; i >= 1; i--) {
    const txnId = data[i][idxTxn];
    if (!txnId || seenTxns.has(txnId)) continue;

    seenTxns.add(txnId);
    recent.push({
      txnId: txnId,
      date: formatDate(data[i][idxDate]),
      bol: data[i][idxBol],
      manufacturer: data[i][idxMan],
      push: data[i][idxPush]
    });

    if (recent.length >= limit) break;
  }

  return recent;
}

/**
 * Re-generates labels for a past transaction by reading Inbound_Skids
 */
function regenerateLabelsForTxn(txnId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const skidsSheet = ss.getSheetByName(TABS.INBOUND_SKIDS);
  const masterSheet = ss.getSheetByName(TABS.MASTER_LOG); // To get Push#, BOL, Mfg if needed
  
  if (!skidsSheet || !masterSheet) throw new Error('Missing data sheets');

  // 1. Get Transaction Details from Master Log (for Label Header info)
  const mData = masterSheet.getDataRange().getValues();
  const mHeaders = mData[0];
  const mIdxTxn = mHeaders.indexOf('Txn_ID');
  
  let basicInfo = { manufacturer: '', pushNumber: '', project: '', bolNumber: '' };
  
  // Find first row of this txn to get general info
  for (let i = 1; i < mData.length; i++) {
    if (String(mData[i][mIdxTxn]) === String(txnId)) {
      basicInfo.manufacturer = mData[i][mHeaders.indexOf('Manufacturer')] || '';
      basicInfo.pushNumber = mData[i][mHeaders.indexOf('Push #')] || '';
      basicInfo.bolNumber = mData[i][mHeaders.indexOf('BOL_Number')] || '';
      break;
    }
  }

  // 2. Get Skid Details from Inbound_Skids
  const sData = skidsSheet.getDataRange().getValues();
  const sHeaders = sData[0];
  const sMap = {}; sHeaders.forEach((h, i) => sMap[h] = i);

  const labelData = [];
  const txnRows = sData.filter(r => String(r[sMap['TXN_ID']]) === String(txnId));

  txnRows.forEach((row, i) => {
    labelData.push({
      skidId: row[sMap['Skid_ID']],
      fbpn: row[sMap['FBPN']],
      quantity: row[sMap['Qty_on_Skid']],
      sku: row[sMap['SKU']],
      manufacturer: basicInfo.manufacturer,
      project: row[sMap['Project']],
      pushNumber: basicInfo.pushNumber,
      dateReceived: formatDate(row[sMap['Date']]),
      skidNumber: row[sMap['Skid_Sequence']] || (i + 1),
      totalSkids: txnRows.length
    });
  });

  if (labelData.length === 0) throw new Error('No skids found for this Transaction ID');

  // 3. Generate
  const res = generateSkidLabels(labelData, { bolNumber: basicInfo.bolNumber });
  return res;
}

/**
 * Generates labels for all skids associated with a specific BOL Number.
 * Handles cases where a BOL might span multiple transactions.
 */
function generateLabelsByBOL(bolNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(TABS.MASTER_LOG);
  const skidsSheet = ss.getSheetByName(TABS.INBOUND_SKIDS);
  
  if (!masterSheet || !skidsSheet) throw new Error('Missing required data sheets (Master_Log or Inbound_Skids).');

  const bolSearch = String(bolNumber || '').trim().toUpperCase();
  if (!bolSearch) throw new Error('Invalid BOL Number provided.');

  // 1. Find Transaction IDs linked to this BOL in Master_Log
  const mData = masterSheet.getDataRange().getValues();
  const mHeaders = mData[0];
  const idxBol = mHeaders.indexOf('BOL_Number');
  const idxTxn = mHeaders.indexOf('Txn_ID');
  const idxMan = mHeaders.indexOf('Manufacturer');
  const idxPush = mHeaders.indexOf('Push #');
  
  if (idxBol === -1 || idxTxn === -1) throw new Error('Required columns missing in Master_Log.');

  const txnMap = {}; // TxnID -> { manufacturer, push }
  
  for (let i = 1; i < mData.length; i++) {
    const rowBol = String(mData[i][idxBol] || '').trim().toUpperCase();
    if (rowBol === bolSearch) {
      const txnId = mData[i][idxTxn];
      if (txnId && !txnMap[txnId]) {
        txnMap[txnId] = {
          manufacturer: mData[i][idxMan] || '',
          pushNumber: mData[i][idxPush] || ''
        };
      }
    }
  }

  const txnIds = Object.keys(txnMap);
  if (txnIds.length === 0) throw new Error(`No transactions found for BOL: ${bolNumber}`);

  // 2. Find Skids matching these Txn IDs
  const sData = skidsSheet.getDataRange().getValues();
  const sHeaders = sData[0];
  const sCol = (name) => sHeaders.indexOf(name);
  
  if (sCol('TXN_ID') === -1 || sCol('Skid_ID') === -1) throw new Error('Required columns missing in Inbound_Skids.');

  const labelData = [];
  
  // Filter and map rows
  let skidSeqCounter = 1; 

  for (let i = 1; i < sData.length; i++) {
    const rowTxn = String(sData[i][sCol('TXN_ID')]);
    if (txnMap[rowTxn]) {
      const info = txnMap[rowTxn];
      
      labelData.push({
        skidId: sData[i][sCol('Skid_ID')],
        fbpn: sData[i][sCol('FBPN')],
        quantity: sData[i][sCol('Qty_on_Skid')],
        sku: sData[i][sCol('SKU')],
        project: sData[i][sCol('Project')],
        dateReceived: formatDate(sData[i][sCol('Date')]),
        manufacturer: info.manufacturer,
        pushNumber: info.pushNumber,
        skidNumber: sData[i][sCol('Skid_Sequence')] || skidSeqCounter++,
        totalSkids: 0 // Will update after collecting all
      });
    }
  }

  if (labelData.length === 0) throw new Error(`No skid details found for BOL: ${bolNumber} (Transactions exist but skids missing).`);

  const total = labelData.length;
  labelData.forEach(d => d.totalSkids = total);

  // 3. Generate Labels
  const res = generateSkidLabels(labelData, { bolNumber: bolNumber });
  
  return res;
}

/**
 * Generates a manual label from user input
 */
function generateManualLabel(data) {
  // data: { fbpn, qty, manufacturer, project, push, bol, copies }
  const copies = parseInt(data.copies) || 1;
  const labelData = [];
  const now = new Date();
  const dateStr = formatDate(now);
  
  // We need a temporary Skid ID for manual labels or generate a real one?
  // Usually manual labels are for existing skids or ad-hoc. 
  // Let's generate a placeholder SKID ID to ensure barcode renders.
  const tempSkidBase = getNextSkidIdBase();

  for (let i = 0; i < copies; i++) {
    const sku = generateSKU(data.fbpn, data.manufacturer);
    const skidId = `MANUAL_${String(tempSkidBase + i + 1).padStart(6, '0')}`;
    
    labelData.push({
      skidId: skidId,
      fbpn: data.fbpn.toUpperCase(),
      quantity: data.qty,
      sku: sku,
      manufacturer: data.manufacturer,
      project: data.project,
      pushNumber: data.push,
      dateReceived: dateStr,
      skidNumber: 1,
      totalSkids: 1
    });
  }

  const res = generateSkidLabels(labelData, { bolNumber: data.bol || 'MANUAL' });
  return res;
}