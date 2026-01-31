// ============================================================================
// IMS_Inbound_SkidLabelModal.js - Manual Skid Label Generation
// ============================================================================

/**
 * Generates manual inbound skid labels from user input.
 * Called from InboundSkidLabelModal.html
 *
 * @param {Object} data - Label data from the modal form
 * @param {string} data.fbpn - The FBPN for the label
 * @param {number} data.qty - Quantity on the skid
 * @param {string} data.manufacturer - Manufacturer name
 * @param {string} data.project - Project name
 * @param {string} data.push - Push number (optional)
 * @param {string} data.bol - BOL number (optional, used for folder organization)
 * @param {number} data.skidNumber - Current skid number in sequence
 * @param {number} data.totalSkids - Total skids in delivery
 * @param {number} data.copies - Number of label copies to generate
 * @param {string} data.notes - Optional notes
 * @returns {Object} Result with success status and PDF/HTML URLs
 */
function generateManualSkidLabelFromModal(data) {
  try {
    // Validate required fields
    if (!data.fbpn) throw new Error('FBPN is required.');
    if (!data.qty || data.qty <= 0) throw new Error('Quantity must be greater than 0.');
    if (!data.manufacturer) throw new Error('Manufacturer is required.');
    if (!data.project) throw new Error('Project is required.');

    const copies = Math.min(50, Math.max(1, parseInt(data.copies) || 1));
    const skidNumber = parseInt(data.skidNumber) || 1;
    const totalSkids = parseInt(data.totalSkids) || 1;
    const now = new Date();
    const dateStr = formatDate(now);

    // Build label data array
    const labelData = [];

    for (let i = 0; i < copies; i++) {
      // Generate a unique Skid ID for each label
      const skidId = generateRandomId('SKD-', 8);
      const sku = generateSKU(data.fbpn, data.manufacturer);

      labelData.push({
        skidId: skidId,
        fbpn: String(data.fbpn).toUpperCase().trim(),
        quantity: data.qty,
        sku: sku,
        manufacturer: String(data.manufacturer).trim(),
        project: String(data.project).trim(),
        pushNumber: String(data.push || '').trim(),
        dateReceived: dateStr,
        skidNumber: skidNumber,
        totalSkids: totalSkids,
        notes: String(data.notes || '').trim()
      });
    }

    // Determine target folder
    let targetFolder = null;
    const bolNumber = String(data.bol || 'MANUAL').trim();

    try {
      // Try to create/use proper folder structure
      targetFolder = createManualLabelFolder_(now, bolNumber);
    } catch (e) {
      Logger.log('Could not create target folder: ' + e.toString());
      // Will fall back to default Labels folder in generateSkidLabels
    }

    // Generate labels using existing function
    const result = generateSkidLabels(labelData, {
      bolNumber: bolNumber,
      targetFolder: targetFolder
    });

    if (!result || !result.success) {
      return {
        success: false,
        message: (result && result.message) ? result.message : 'Label generation failed.'
      };
    }

    return {
      success: true,
      pdfUrl: result.pdfUrl || '',
      htmlUrl: result.htmlUrl || '',
      labelCount: labelData.length,
      message: `Successfully generated ${labelData.length} label(s).`
    };

  } catch (err) {
    Logger.log('generateManualSkidLabelFromModal error: ' + err.toString());
    return {
      success: false,
      message: 'Error: ' + err.message
    };
  }
}

/**
 * Creates folder structure for manual labels.
 * Structure: Inbound_Uploads/ManualLabels/{MonthYear}/{BOL or MANUAL}
 *
 * @param {Date} dateObj - Date for folder organization
 * @param {string} bolNumber - BOL number or 'MANUAL' for folder name
 * @returns {Folder} Google Drive folder
 */
function createManualLabelFolder_(dateObj, bolNumber) {
  const rootId = (typeof FOLDERS !== 'undefined' && FOLDERS.INBOUND_UPLOADS)
    ? FOLDERS.INBOUND_UPLOADS
    : (typeof FOLDERS !== 'undefined' && FOLDERS.IMS_ROOT ? FOLDERS.IMS_ROOT : '');

  if (!rootId) throw new Error("Inbound Uploads Folder ID not configured.");

  const rootFolder = DriveApp.getFolderById(rootId);

  // Create ManualLabels subfolder
  const manualLabelsFolder = getOrCreateSubfolder_(rootFolder, 'ManualLabels');

  // Create month folder
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const monthName = `${months[dateObj.getMonth()]} ${dateObj.getFullYear()}`;
  const monthFolder = getOrCreateSubfolder_(manualLabelsFolder, monthName);

  // Create BOL/day folder
  const safeBol = String(bolNumber || 'MANUAL').trim().replace(/[\/\\?%*:|"<>\.]/g, '_');
  const dayStr = String(dateObj.getDate()).padStart(2, '0');
  const folderName = safeBol === 'MANUAL' ? `Manual_${dayStr}` : safeBol;
  const targetFolder = getOrCreateSubfolder_(monthFolder, folderName);

  return targetFolder;
}

/**
 * Shell wrapper function called from HTML modal.
 * This is the entry point from the frontend.
 *
 * @param {Object} data - Label data from modal
 * @returns {Object} Result object
 */
function shell_generateManualSkidLabel(data) {
  return generateManualSkidLabelFromModal(data);
}

/**
 * Opens the manual skid label creation modal.
 */
function openInboundSkidLabelModal() {
  const html = HtmlService.createTemplateFromFile('InboundSkidLabelModal')
    .evaluate()
    .setWidth(850)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Inbound Skid Label');
}
