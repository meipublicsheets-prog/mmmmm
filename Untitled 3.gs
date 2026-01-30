function buildOneRowPerSkuPerProject_FromProjectMaster() {
  const ss = SpreadsheetApp.getActive();

  const SRC_NAME = 'Project_Master';
  const OUT_NAME = 'ProjectMaster_SkuByProject_Totals';

  const sh = ss.getSheetByName(SRC_NAME);
  if (!sh) throw new Error(`Missing sheet: ${SRC_NAME}`);

  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error(`${SRC_NAME} has no data rows.`);

  // Project_Master columns:
  // F Project, G SKU, H Qty_Ordered, I Qty_Received
  const iProject = 5; // F
  const iSku     = 6; // G
  const iOrd     = 7; // H
  const iRec     = 8; // I

  const trim = (v) => String(v ?? '').trim();
  const keyPart = (v) => trim(v).toUpperCase();
  const toNum = (v) => {
    if (v === '' || v == null) return 0;
    const n = Number(v);
    return isNaN(n) ? 0 : n;
  };
  const SEP = '\u0001';

  // Key = Project + SKU ensures one unique SKU per project
  const map = new Map(); // key -> {project, sku, qtyOrderedTotal, qtyReceivedTotal}

  for (let r = 1; r < data.length; r++) {
    const row = data[r];

    const project = trim(row[iProject]);
    const sku = trim(row[iSku]);
    if (!project || !sku) continue;

    const key = `${keyPart(project)}${SEP}${keyPart(sku)}`;

    const qtyOrdered = toNum(row[iOrd]);
    const qtyReceived = toNum(row[iRec]);

    const cur = map.get(key);
    if (cur) {
      cur.qtyOrderedTotal += qtyOrdered;
      cur.qtyReceivedTotal += qtyReceived;
    } else {
      map.set(key, {
        project,
        sku,
        qtyOrderedTotal: qtyOrdered,
        qtyReceivedTotal: qtyReceived
      });
    }
  }

  // Recreate output sheet
  const existing = ss.getSheetByName(OUT_NAME);
  if (existing) ss.deleteSheet(existing);
  const out = ss.insertSheet(OUT_NAME);

  const rows = [['Project', 'SKU', 'Total_Qty_Ordered', 'Total_Qty_Received']];

  // Optional: sort for readability
  const list = Array.from(map.values()).sort((a, b) => {
    const ap = a.project.toUpperCase(), bp = b.project.toUpperCase();
    if (ap !== bp) return ap < bp ? -1 : 1;
    const as = a.sku.toUpperCase(), bs = b.sku.toUpperCase();
    return as < bs ? -1 : as > bs ? 1 : 0;
  });

  for (const g of list) {
    rows.push([g.project, g.sku, g.qtyOrderedTotal, g.qtyReceivedTotal]);
  }

  out.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  out.setFrozenRows(1);
  out.autoResizeColumns(1, rows[0].length);
  if (rows.length > 1) out.getRange(2, 3, rows.length - 1, 2).setNumberFormat('#,##0.00');
}

function createZeroStockPutAwayLog() {
  // --- CONFIGURATION ---
  const TEMPLATE_ID = '1jguHxekNv22rWCVOxrk2TA8aJnUxpnQMnyi9DfofHl0';
  const SHEET_NAME = 'Bin_Stock';
  // ---------------------

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Error: Sheet '${SHEET_NAME}' not found.`);
    return;
  }

  // 1. Get Data
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const getIdx = (name) => headers.indexOf(name);
  
  const colIdx = {
    bin: getIdx('Bin_Code'),
    fbpn: getIdx('FBPN'),
    push: getIdx('Push_Number'),
    mfg: getIdx('Manufacturer'),
    qty: getIdx('Current_Quantity')
  };

  if (colIdx.bin === -1 || colIdx.qty === -1) {
    SpreadsheetApp.getUi().alert("Error: Missing 'Bin_Code' or 'Current_Quantity' column.");
    return;
  }

  // 2. Filter for Zero/Blank
  const itemsToProcess = data.slice(1).filter(row => {
    const q = row[colIdx.qty];
    return (q === 0 || q === "" || q === null);
  });

  if (itemsToProcess.length === 0) {
    SpreadsheetApp.getUi().alert("No zero stock items found.");
    return;
  }

  // 3. Prepare Doc
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  const tempFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(`PutAway_Log_${dateStr}_TEMP`);
  const doc = DocumentApp.openById(tempFile.getId());
  const body = doc.getBody();
  const docHeader = doc.getHeader(); // Access the header section if it exists

  // Replace {{Date}} in Body AND Header (just in case you moved it)
  body.replaceText("{{Date}}", dateStr);
  if (docHeader) docHeader.replaceText("{{Date}}", dateStr);

  // 4. Find Table
  const tables = body.getTables();
  let targetTable = null;
  let templateRowIndex = -1;

  for (let t of tables) {
    for (let r = 0; r < t.getNumRows(); r++) {
      if (t.getRow(r).getText().includes("{{Bin_Code}}")) {
        targetTable = t;
        templateRowIndex = r;
        break;
      }
    }
    if (targetTable) break;
  }

  if (!targetTable) {
    doc.saveAndClose();
    tempFile.setTrashed(true);
    SpreadsheetApp.getUi().alert("Error: Placeholder '{{Bin_Code}}' not found in table.");
    return;
  }

  // 5. Fill Table
  const templateRow = targetTable.getRow(templateRowIndex);

  itemsToProcess.forEach(row => {
    // FIX: Use appendTableRow instead of insertRow
    const newRow = targetTable.appendTableRow(templateRow.copy());
    
    newRow.getCell(0).setText(row[colIdx.bin] || ""); 
    newRow.getCell(1).setText(row[colIdx.fbpn] || "");
    newRow.getCell(2).setText(row[colIdx.push] || "");
    newRow.getCell(3).setText(row[colIdx.mfg] || ""); 
    newRow.getCell(4).setText(""); // Blank for manual count
  });

  // Remove the placeholder row (leaving the header row intact)
  targetTable.removeRow(templateRowIndex);

  // 6. Save & PDF
  doc.saveAndClose();
  const pdfBlob = tempFile.getAs(MimeType.PDF);
  const pdfFile = DriveApp.createFile(pdfBlob).setName(`PutAway_Log_${dateStr}.pdf`);
  tempFile.setTrashed(true);

  // 7. Show Link
  const html = `<p>PDF Created: <a href="${pdfFile.getUrl()}" target="_blank">Open PDF</a></p>
                <button onclick="google.script.host.close()">Close</button>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setHeight(100), "Success");
}


