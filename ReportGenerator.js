function showReportGeneratorModal() {
  const html = HtmlService.createTemplateFromFile('ReportGeneratorModal')
    .evaluate()
    .setWidth(950)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Reports');
}

function generateInboundReport(params) {
  try {
    const templateId = REPORT_TEMPLATES.INBOUND;
    // Get data from Master_Log
    const reportData = getInboundReportData(params);
    if (reportData.rows.length === 0) {
      return {
        success: false,
        message: 'No inbound data found for the selected date range'
      };
    }
    
    // Create document from template
    const templateDoc = DriveApp.getFileById(templateId);
    const tempDoc = templateDoc.makeCopy('TEMP_Inbound_Report_' + new Date().getTime());
    const docId = tempDoc.getId();
    const doc = DocumentApp.openById(docId);
    
    const body = doc.getBody();
    const header = doc.getHeader();
    const footer = doc.getFooter();
    
    // Replace header placeholders
    const headerPlaceholders = {
      '{{Frequency}}': params.frequency,
      '{{DateRange}}': reportData.dateRange
    };
    replacePlaceholders(body, headerPlaceholders);
    if (header) replacePlaceholders(header, headerPlaceholders);
    if (footer) replacePlaceholders(footer, headerPlaceholders);
    
    // Process table with alternating BOL backgrounds
    processInboundReportTable(body, reportData.rows);
    
    doc.saveAndClose();
    
    // Export to PDF
    const pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
    const filename = `${params.frequency}_Inbound_${formatDateForFilename(new Date())}.pdf`;
    pdfBlob.setName(filename);
    
    // Save to Reports folder
    const pdfFile = saveReportToFolder(pdfBlob, 'Inbound', params.frequency);
    
    // Delete temp document
    DriveApp.getFileById(docId).setTrashed(true);
    
    return {
      success: true,
      url: pdfFile.getUrl(),
      name: pdfFile.getName(),
      rowCount: reportData.rows.length
    };
  } catch (error) {
    Logger.log('Error generating inbound report: ' + error.toString());
    return {
      success: false,
      message: 'Error generating report: ' + error.toString()
    };
  }
}

/**
 * Generate Outbound Report
 * @param {Object} params - Report parameters {frequency, startDate, endDate}
 * @returns {Object} Result with PDF URL
 */
function generateOutboundReport(params) {
  try {
    const templateId = REPORT_TEMPLATES.OUTBOUND;
    // Get data from OutboundLog
    const reportData = getOutboundReportData(params);
    if (reportData.rows.length === 0) {
      return {
        success: false,
        message: 'No outbound data found for the selected date range'
      };
    }
    
    // Create document from template
    const templateDoc = DriveApp.getFileById(templateId);
    const tempDoc = templateDoc.makeCopy('TEMP_Outbound_Report_' + new Date().getTime());
    const docId = tempDoc.getId();
    const doc = DocumentApp.openById(docId);
    
    const body = doc.getBody();
    const header = doc.getHeader();
    const footer = doc.getFooter();
    
    // Replace header placeholders
    const headerPlaceholders = {
      '{{Frequency}}': params.frequency,
      '{{DateRange}}': reportData.dateRange
    };
    replacePlaceholders(body, headerPlaceholders);
    if (header) replacePlaceholders(header, headerPlaceholders);
    if (footer) replacePlaceholders(footer, headerPlaceholders);
    
    // Process table with alternating Order Number backgrounds
    processOutboundReportTable(body, reportData.rows);
    
    doc.saveAndClose();
    
    // Export to PDF
    const pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
    const filename = `${params.frequency}_Outbound_${formatDateForFilename(new Date())}.pdf`;
    pdfBlob.setName(filename);
    
    // Save to Reports folder
    const pdfFile = saveReportToFolder(pdfBlob, 'Outbound', params.frequency);
    
    // Delete temp document
    DriveApp.getFileById(docId).setTrashed(true);
    
    return {
      success: true,
      url: pdfFile.getUrl(),
      name: pdfFile.getName(),
      rowCount: reportData.rows.length
    };
  } catch (error) {
    Logger.log('Error generating outbound report: ' + error.toString());
    return {
      success: false,
      message: 'Error generating report: ' + error.toString()
    };
  }
}

/**
 * Get inbound report data from Master_Log
 * @param {Object} params - Report parameters
 * @returns {Object} Report data with rows and date range
 */
function getInboundReportData(params) {
  const sheet = getSheet(TABS.MASTER_LOG);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Get column indices
  const dateIndex = headers.indexOf('Date_Received');
  const warehouseIndex = headers.indexOf('Warehouse');
  const fbpnIndex = headers.indexOf('FBPN');
  const qtyIndex = headers.indexOf('Qty_Received');
  const manufacturerIndex = headers.indexOf('Manufacturer');
  const carrierIndex = headers.indexOf('Carrier');
  const poIndex = headers.indexOf('Customer_PO_Number');
  const bolIndex = headers.indexOf('BOL_Number');
  const pushIndex = headers.indexOf('Push #');
  
  const startDate = new Date(params.startDate);
  const endDate = new Date(params.endDate);
  endDate.setHours(23, 59, 59, 999); // Include entire end date
  
  const rows = [];
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = new Date(row[dateIndex]);
    
    if (rowDate >= startDate && rowDate <= endDate) {
      rows.push({
        date: rowDate, // Keep original date object for sorting
        dateReceived: formatDate(rowDate),
        warehouse: row[warehouseIndex] || '',
        fbpn: row[fbpnIndex] || '',
        qty: row[qtyIndex] || 0,
        manufacturer: row[manufacturerIndex] || '',
        carrier: row[carrierIndex] || '',
        poNumber: row[poIndex] || '',
        bol: row[bolIndex] || '',
        push: row[pushIndex] || ''
      });
    }
  }
  
  // Sort by date first, then by BOL number to group together
  rows.sort((a, b) => {
    // First sort by date
    if (a.date < b.date) return -1;
    if (a.date > b.date) return 1;
    // Then by BOL number
    if (a.bol < b.bol) return -1;
    if (a.bol > b.bol) return 1;
    return 0;
  });
  
  const dateRange = formatDate(startDate) + ' - ' + formatDate(endDate);
  
  return {
    rows: rows,
    dateRange: dateRange
  };
}

/**
 * Get outbound report data from OutboundLog
 * @param {Object} params - Report parameters
 * @returns {Object} Report data with rows and date range
 */
function getOutboundReportData(params) {
  const sheet = getSheet(TABS.OUTBOUNDLOG);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Get column indices
  const dateIndex = headers.indexOf('Date');
  const companyIndex = headers.indexOf('Company');
  const projectIndex = headers.indexOf('Project');
  const fbpnIndex = headers.indexOf('FBPN');
  const qtyIndex = headers.indexOf('Qty');
  const manufacturerIndex = headers.indexOf('Manufacturer');
  const taskNumberIndex = headers.indexOf('Task_Number');
  const orderNumberIndex = headers.indexOf('Order_Number');
  
  const startDate = new Date(params.startDate);
  const endDate = new Date(params.endDate);
  endDate.setHours(23, 59, 59, 999); // Include entire end date
  
  const rows = [];
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = new Date(row[dateIndex]);
    
    if (rowDate >= startDate && rowDate <= endDate) {
      rows.push({
        dateObj: rowDate, // Keep original date object for sorting
        date: formatDate(rowDate),
        company: row[companyIndex] || '',
        project: row[projectIndex] || '',
        fbpn: row[fbpnIndex] || '',
        qty: row[qtyIndex] || 0,
        manufacturer: row[manufacturerIndex] || '',
        taskNumber: row[taskNumberIndex] || '',
        orderNumber: row[orderNumberIndex] || ''
      });
    }
  }
  
  // Sort by date first, then by Order Number to group together
  rows.sort((a, b) => {
    // First sort by date
    if (a.dateObj < b.dateObj) return -1;
    if (a.dateObj > b.dateObj) return 1;
    // Then by Order Number
    if (a.orderNumber < b.orderNumber) return -1;
    if (a.orderNumber > b.orderNumber) return 1;
    return 0;
  });
  
  const dateRange = formatDate(startDate) + ' - ' + formatDate(endDate);
  
  return {
    rows: rows,
    dateRange: dateRange
  };
}

/**
 * Process inbound report table with alternating BOL backgrounds
 * @param {Body} body - Document body
 * @param {Array} rows - Data rows to insert
 */
function processInboundReportTable(body, rows) {
  // Find the table containing placeholders
  let table = null;
  let templateRowIndex = -1;
  const numChildren = body.getNumChildren();
  
  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);
    if (element.getType() === DocumentApp.ElementType.TABLE) {
      const tbl = element.asTable();
      const numRows = tbl.getNumRows();
      
      // Search through all rows for placeholder
      for (let r = 0; r < numRows; r++) {
        const row = tbl.getRow(r);
        const cellText = row.getText();
        
        // Check for any inbound placeholder
        if (cellText.indexOf('{{DateReceived}}') !== -1 || 
            cellText.indexOf('{{Warehouse}}') !== -1 ||
            cellText.indexOf('{{FBPN}}') !== -1 ||
            cellText.indexOf('{{Manufacturer}}') !== -1) {
          table = tbl;
          templateRowIndex = r;
          break;
        }
      }
      if (table) break;
    }
  }
  
  if (!table || templateRowIndex === -1) {
    Logger.log('Warning: Could not find table template in document');
    Logger.log('Number of children in body: ' + numChildren);
    
    // Log all elements in body for debugging
    for (let i = 0; i < numChildren; i++) {
      const element = body.getChild(i);
      Logger.log('Element ' + i + ': ' + element.getType());
    }
    return;
  }
  
  const templateRow = table.getRow(templateRowIndex);
  const numCells = templateRow.getNumCells();
  Logger.log('Found template table with ' + numCells + ' cells at row ' + templateRowIndex);
  
  // Track BOL numbers for alternating colors
  let currentBOL = '';
  let useAlternateColor = false;
  const colorA = '#FFFFFF'; // White
  const colorB = '#F0F9FF'; // Light blue
  const colorC = '#FEF3C7'; // Light yellow
  
  // Process each row
  rows.forEach((item, index) => {
    // Check if BOL changed
    if (item.bol !== currentBOL) {
      currentBOL = item.bol;
      useAlternateColor = !useAlternateColor;
    }
    
    const newRow = table.insertTableRow(templateRowIndex + index);
    
    // Copy cells from template and populate
    for (let c = 0; c < numCells; c++) {
      const templateCell = templateRow.getCell(c);
      const newCell = newRow.appendTableCell();
      
      // Get template text
      let cellText = templateCell.getText();
      
      // Replace placeholders
      cellText = cellText.replace(/{{DateReceived}}/g, item.dateReceived);
      cellText = cellText.replace(/{{Warehouse}}/g, item.warehouse);
      cellText = cellText.replace(/{{FBPN}}/g, item.fbpn);
      cellText = cellText.replace(/{{Qty}}/g, item.qty.toString());
      cellText = cellText.replace(/{{Manufacturer}}/g, item.manufacturer);
      cellText = cellText.replace(/{{Carrier}}/g, item.carrier);
      cellText = cellText.replace(/{{PO_Number}}/g, item.poNumber);
      cellText = cellText.replace(/{{BOL}}/g, item.bol);
      cellText = cellText.replace(/{{Push}}/g, item.push);
      
      // Set text
      newCell.clear();
      newCell.setText(cellText);
      
      // Copy formatting from template
      const templateAttrs = templateCell.getAttributes();
      newCell.setAttributes(templateAttrs);
      
      // Apply alternating background color based on BOL
      newCell.setBackgroundColor(useAlternateColor ? colorB : colorC);
    }
  });
  
  // Remove the template row
  try {
    table.removeRow(templateRowIndex + rows.length);
  } catch (e) {
    Logger.log('Warning: Could not remove template row: ' + e.toString());
  }
}

/**
 * Process outbound report table with alternating Order Number backgrounds
 * @param {Body} body - Document body
 * @param {Array} rows - Data rows to insert
 */
function processOutboundReportTable(body, rows) {
  // Find the table containing placeholders
  let table = null;
  let templateRowIndex = -1;
  const numChildren = body.getNumChildren();
  
  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);
    if (element.getType() === DocumentApp.ElementType.TABLE) {
      const tbl = element.asTable();
      const numRows = tbl.getNumRows();
      // Search through all rows for placeholder
      for (let r = 0; r < numRows; r++) {
        const row = tbl.getRow(r);
        const cellText = row.getText();
        
        // Check for any outbound placeholder
        if (cellText.indexOf('{{Date}}') !== -1 || 
            cellText.indexOf('{{Company}}') !== -1 ||
            cellText.indexOf('{{Order_Number}}') !== -1 ||
            cellText.indexOf('{{FBPN}}') !== -1 ||
            cellText.indexOf('{{Manufacturer}}') !== -1) {
          table = tbl;
          templateRowIndex = r;
          break;
        }
      }
      if (table) break;
    }
  }
  
  if (!table || templateRowIndex === -1) {
    Logger.log('Warning: Could not find table template in document');
    Logger.log('Number of children in body: ' + numChildren);
    
    // Log all elements in body for debugging
    for (let i = 0; i < numChildren; i++) {
      const element = body.getChild(i);
      Logger.log('Element ' + i + ': ' + element.getType());
    }
    return;
  }
  
  const templateRow = table.getRow(templateRowIndex);
  const numCells = templateRow.getNumCells();
  Logger.log('Found template table with ' + numCells + ' cells at row ' + templateRowIndex);
  
  // Track Order Numbers for alternating colors
  let currentOrderNumber = '';
  let useAlternateColor = false;
  const colorA = '#FFFFFF'; // White
  const colorB = '#F3F4F6'; // Light gray
  const colorC = '#FEF3C7'; // Light yellow
  
  // Process each row
  rows.forEach((item, index) => {
    // Check if Order Number changed
    if (item.orderNumber !== currentOrderNumber) {
      currentOrderNumber = item.orderNumber;
      useAlternateColor = !useAlternateColor;
    }
    
    const newRow = table.insertTableRow(templateRowIndex + index);
    
    // Copy cells from template and populate
    for (let c = 0; c < numCells; c++) {
      const templateCell = templateRow.getCell(c);
      const newCell = newRow.appendTableCell();
      
      // Get template text
      let cellText = templateCell.getText();
      
      // Replace placeholders
      cellText = cellText.replace(/{{Date}}/g, item.date);
      cellText = cellText.replace(/{{Company}}/g, item.company);
      cellText = cellText.replace(/{{Project}}/g, item.project);
      cellText = cellText.replace(/{{FBPN}}/g, item.fbpn);
      cellText = cellText.replace(/{{Qty}}/g, item.qty.toString());
      cellText = cellText.replace(/{{Manufacturer}}/g, item.manufacturer);
      cellText = cellText.replace(/{{Task_Number}}/g, item.taskNumber);
      cellText = cellText.replace(/{{Order_Number}}/g, item.orderNumber);
      
      // Set text
      newCell.clear();
      newCell.setText(cellText);
      
      // Copy formatting from template
      const templateAttrs = templateCell.getAttributes();
      newCell.setAttributes(templateAttrs);
      
      // Apply alternating background color based on Order Number
      newCell.setBackgroundColor(useAlternateColor ? colorB : colorC);
    }
  });
  
  // Remove the template row
  try {
    table.removeRow(templateRowIndex + rows.length);
  } catch (e) {
    Logger.log('Warning: Could not remove template row: ' + e.toString());
  }
}

/**
 * Save report PDF to organized folder structure
 * @param {Blob} pdfBlob - PDF blob to save
 * @param {string} reportType - 'Inbound' or 'Outbound'
 * @param {string} frequency - 'Daily', 'Weekly', or 'Monthly'
 * @returns {File} Saved PDF file
 */
function saveReportToFolder(pdfBlob, reportType, frequency) {
  const rootFolder = DriveApp.getFolderById(FOLDERS.IMS_Reports);
  
  // Create or get report type folder (Inbound Reports / Outbound Reports)
  const typeFolderName = reportType + ' Reports';
  const typeFolder = getOrCreateFolder(rootFolder, typeFolderName);
  
  // Create or get frequency folder (Daily / Weekly / Monthly)
  const frequencyFolder = getOrCreateFolder(typeFolder, frequency);
  
  // Create or get year folder
  const year = new Date().getFullYear().toString();
  const yearFolder = getOrCreateFolder(frequencyFolder, year);
  
  // Save PDF
  const pdfFile = yearFolder.createFile(pdfBlob);
  
  return pdfFile;
}

/**
 * Format date for filename (YYYYMMDD)
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDateForFilename(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

/**
 * Get date range based on preset selection
 * @param {string} preset - Preset name
 * @returns {Object} Object with startDate and endDate
 */
function getDateRangePreset(preset) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let startDate, endDate;
  
  switch(preset) {
    case 'today':
      startDate = new Date(today);
      endDate = new Date(today);
      break;
      
    case 'yesterday':
      startDate = new Date(today);
      startDate.setDate(startDate.getDate() - 1);
      endDate = new Date(startDate);
      break;
      
    case 'thisWeek':
      startDate = new Date(today);
      const dayOfWeek = startDate.getDay();
      startDate.setDate(startDate.getDate() - dayOfWeek); // Start of week (Sunday)
      endDate = new Date(today);
      break;
      
    case 'lastWeek':
      startDate = new Date(today);
      const lastWeekDay = startDate.getDay();
      startDate.setDate(startDate.getDate() - lastWeekDay - 7); // Start of last week
      endDate = new Date(startDate);
      endDate.setDate(endDate.getDate() + 6); // End of last week
      break;
      
    case 'thisMonth':
      startDate = new Date(today.getFullYear(), today.getMonth(), 1);
      endDate = new Date(today);
      break;
      
    case 'lastMonth':
      startDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      endDate = new Date(today.getFullYear(), today.getMonth(), 0); // Last day of last month
      break;
      
    default:
      startDate = new Date(today);
      endDate = new Date(today);
  }
  
  return {
    startDate: formatDate(startDate),
    endDate: formatDate(endDate)
  };
}
