function openBinLookupModal() {
  const html = HtmlService.createTemplateFromFile('BinLookupModal')
    .evaluate()
    .setWidth(950)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bin Lookup & Search');
}


function searchBins(searchParams) {
  try {
    const sheet = getSheet(TABS.BIN_STOCK);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const binCodeIndex = headers.indexOf('Bin_Code');
    const binNameIndex = headers.indexOf('Bin_Name');
    const pushNumberIndex = headers.indexOf('Push_Number');
    const fbpnIndex = headers.indexOf('FBPN');
    const manufacturerIndex = headers.indexOf('Manufacturer');
    const projectIndex = headers.indexOf('Project');
    const initialQtyIndex = headers.indexOf('Initial_Quantity');
    const currentQtyIndex = headers.indexOf('Current_Quantity');
    const stockPercentageIndex = headers.indexOf('Stock_Percentage');
    const skidIdIndex = headers.indexOf('Skid_ID');
    
    const results = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let matches = true;
      
      if (searchParams.binCode && searchParams.binCode.trim() !== '') {
        const searchTerm = searchParams.binCode.trim().toLowerCase();
        const binCode = (row[binCodeIndex] || '').toString().toLowerCase();
        if (!binCode.includes(searchTerm)) matches = false;
      }
      
      if (searchParams.fbpn && searchParams.fbpn.trim() !== '') {
        const searchTerm = searchParams.fbpn.trim().toLowerCase();
        const fbpn = (row[fbpnIndex] || '').toString().toLowerCase();
        if (!fbpn.includes(searchTerm)) matches = false;
      }
      
      if (searchParams.manufacturer && searchParams.manufacturer !== '') {
        if (row[manufacturerIndex] !== searchParams.manufacturer) matches = false;
      }
      
      if (searchParams.project && searchParams.project !== '') {
        if (row[projectIndex] !== searchParams.project) matches = false;
      }
      
      if (searchParams.stockStatus && searchParams.stockStatus !== '') {
        const currentQty = parseInt(row[currentQtyIndex]) || 0;
        
        if (searchParams.stockStatus === 'empty' && currentQty > 0) {
          matches = false;
        } else if (searchParams.stockStatus === 'occupied' && currentQty === 0) {
          matches = false;
        } else if (searchParams.stockStatus === 'low' && (currentQty === 0 || currentQty > row[initialQtyIndex] * 0.25)) {
          matches = false;
        }
      }
      
      if (matches) {
        results.push({
          binCode: row[binCodeIndex] || '',
          binName: row[binNameIndex] || '',
          pushNumber: row[pushNumberIndex] || '',
          fbpn: row[fbpnIndex] || '',
          manufacturer: row[manufacturerIndex] || '',
          project: row[projectIndex] || '',
          initialQty: row[initialQtyIndex] || 0,
          currentQty: row[currentQtyIndex] || 0,
          stockPercentage: row[stockPercentageIndex] || 0,
          skidId: row[skidIdIndex] || '',
          rowIndex: i + 1
        });
      }
    }
    
    return results;
    
  } catch (error) {
    Logger.log('Error searching bins: ' + error.toString());
    throw new Error('Error searching bins: ' + error.toString());
  }
}

function getBinDetails(binCode) {
  try {
    const sheet = getSheet(TABS.BIN_STOCK);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const binCodeIndex = headers.indexOf('Bin_Code');
    const binNameIndex = headers.indexOf('Bin_Name');
    const pushNumberIndex = headers.indexOf('Push_Number');
    const fbpnIndex = headers.indexOf('FBPN');
    const manufacturerIndex = headers.indexOf('Manufacturer');
    const projectIndex = headers.indexOf('Project');
    const initialQtyIndex = headers.indexOf('Initial_Quantity');
    const currentQtyIndex = headers.indexOf('Current_Quantity');
    const stockPercentageIndex = headers.indexOf('Stock_Percentage');
    const skidIdIndex = headers.indexOf('Skid_ID');
    
    const binItems = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[binCodeIndex] === binCode) {
        binItems.push({
          binCode: row[binCodeIndex] || '',
          binName: row[binNameIndex] || '',
          pushNumber: row[pushNumberIndex] || '',
          fbpn: row[fbpnIndex] || '',
          manufacturer: row[manufacturerIndex] || '',
          project: row[projectIndex] || '',
          initialQty: row[initialQtyIndex] || 0,
          currentQty: row[currentQtyIndex] || 0,
          stockPercentage: row[stockPercentageIndex] || 0,
          skidId: row[skidIdIndex] || '',
          rowIndex: i + 1
        });
      }
    }
    
    return binItems;
    
  } catch (error) {
    Logger.log('Error getting bin details: ' + error.toString());
    throw new Error('Error getting bin details: ' + error.toString());
  }
}

function getBinHistory(binCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('LocationLog');
    if (!logSheet) return [];
    
    const data = logSheet.getDataRange().getValues();
    const headers = data[0];
    
    const timestampIndex = headers.indexOf('Timestamp');
    const actionIndex = headers.indexOf('Action');
    const fbpnIndex = headers.indexOf('FBPN');
    const manufacturerIndex = headers.indexOf('Manufacturer');
    const binCodeIndex = headers.indexOf('Bin_Code');
    const qtyChangedIndex = headers.indexOf('Qty_Changed');
    const resultingQtyIndex = headers.indexOf('Resulting_Qty');
    const descriptionIndex = headers.indexOf('Description');
    const userEmailIndex = headers.indexOf('User_Email');
    
    const history = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[binCodeIndex] === binCode) {
        history.push({
          timestamp: row[timestampIndex] || '',
          action: row[actionIndex] || '',
          fbpn: row[fbpnIndex] || '',
          manufacturer: row[manufacturerIndex] || '',
          qtyChanged: row[qtyChangedIndex] || 0,
          resultingQty: row[resultingQtyIndex] || 0,
          description: row[descriptionIndex] || '',
          userEmail: row[userEmailIndex] || ''
        });
      }
    }
    
    history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    return history;
    
  } catch (error) {
    Logger.log('Error getting bin history: ' + error.toString());
    return [];
  }
}




/**
 * Quick barcode scanner function
 * Can detect if the scanned value is a bin code or FBPN
 */
function quickBarcodeScan(scanValue) {
  try {
    const sheet = getSheet(TABS.BIN_STOCK);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const binCodeIndex = headers.indexOf('Bin_Code');
    const fbpnIndex = headers.indexOf('FBPN');

    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const binCode = (row[binCodeIndex] || '').toString();
      const fbpn = (row[fbpnIndex] || '').toString();

      if (binCode === scanValue || fbpn === scanValue) {
        results.push({
          binCode: row[binCodeIndex] || '',
          binName: row[headers.indexOf('Bin_Name')] || '',
          fbpn: row[fbpnIndex] || '',
          manufacturer: row[headers.indexOf('Manufacturer')] || '',
          project: row[headers.indexOf('Project')] || '',
          initialQty: row[headers.indexOf('Initial_Quantity')] || 0,
          currentQty: row[headers.indexOf('Current_Quantity')] || 0,
          stockPercentage: row[headers.indexOf('Stock_Percentage')] || 0,
          skidId: row[headers.indexOf('Skid_ID')] || ''
        });
      }
    }

    if (results.length === 0) {
      return { success: false, message: 'No results found for: ' + scanValue };
    }

    return {
      success: true,
      scanType: results[0].binCode === scanValue ? 'BIN_CODE' : 'FBPN',
      results: results
    };

  } catch (error) {
    Logger.log('Error in barcode scan: ' + error.toString());
    return { success: false, message: 'Error scanning: ' + error.toString() };
  }
}
