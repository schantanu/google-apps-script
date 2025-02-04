function test1() {
  // getAttributeData('ITEM');
  Logger.log(getCurrentValues('ITEM','96-002315-A'));
}

/**
 * Helper to determine field type based on header name
 * @param {string} header - Field header name
 * @returns {string} Field type (date, dropdown, or text)
 */
function getFieldType(header) {
  if (ATTRIBUTE_CONFIG.DATES.includes(header)) return 'date';

  const dropdownFields = [
    'VENDOR', 'MATERIAL GROUP', 'GROUP FUNCTION', 'FAMILY', 'SUBFAMILY',
    'DP HIERARCHY', 'TRACKED SET', 'FREQ', 'POWER', 'INTEGRATED', 'TECH'
  ];

  return dropdownFields.includes(header) ? 'dropdown' : 'text';
}

/**
 * Get attribute data configuration based on type (common or item level)
 * @param {string} type - The type of attributes to get ('common' or 'item')
 * @returns {Object} Object containing ids, attributes configuration, and dropdown values
 */
function getAttributeData(type) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(SHEET_CONFIG.DATA.NAME);
    const dropdownsSheet = ss.getSheetByName(SHEET_CONFIG.DROPDOWNS.NAME);

    // Get data
    const data = dataSheet.getRange(
      SHEET_CONFIG.DATA.POSITIONS.DATA_START_ROW,
      SHEET_CONFIG.DATA.POSITIONS.START_COL,
      dataSheet.getLastRow() - 1,
      dataSheet.getLastColumn()
    ).getValues();

    // Get selected fields based on type
    const columnConfig = type === 'COMMON' ?
      ATTRIBUTE_CONFIG.COMMON_LEVEL_FIELDS :
      ATTRIBUTE_CONFIG.ITEM_LEVEL_FIELDS;

    // Get column indices
    const headers = SHEET_CONFIG.DATA.COLUMNS.HEADERS;

    // Get ID column index from config
    const idField = type === 'COMMON' ? 'COMMON ID' : 'ITEM ID';
    const idColIndex = SHEET_CONFIG.DATA.COLUMNS.INDEX[idField] - 1;

    // Get unique ids
    const ids = [...new Set(data.map(row => row[idColIndex]).filter(Boolean))];

    // Map attributes
    const attributes = columnConfig.map(header => ({
      name: header,
      columnIndex: SHEET_CONFIG.DATA.COLUMNS.HEADERS.indexOf(header),
      type: getFieldType(header),
      maxLength: ATTRIBUTE_CONFIG.CHAR_LIMITS[header.replace(/\s+/g, '_')] || 100
    }));

    // Get dropdown values from Dropdowns sheet if needed
    const dropdowns = {};
    if (attributes.some(attr => attr.type === 'dropdown')) {
      const dropdownFields = type === 'COMMON' ?
        Object.keys(SHEET_CONFIG.DROPDOWNS.COLUMNS.INDEX) :
        ATTRIBUTE_CONFIG.ITEM_LEVEL_FIELDS.filter(field =>
          getFieldType(field) === 'dropdown'
        );

      dropdownFields.forEach(field => {
        const colIndex = SHEET_CONFIG.DROPDOWNS.COLUMNS.INDEX[field];
        const values = dropdownsSheet.getRange(
          SHEET_CONFIG.DROPDOWNS.POSITIONS.DATA_START_ROW,
          colIndex,
          dropdownsSheet.getLastRow()
        ).getValues()
          .map(row => row[0])
          .filter(Boolean);
        dropdowns[field.replace(/_/g, ' ')] = values;
      });
    }

    return {
      ids,
      attributes,
      dropdowns,
      config: ATTRIBUTE_CONFIG.ARCHIVED_FIELDS
    };

  } catch (error) {
    Logger.log(error.stack);
    throw new Error('Failed to get attribute data');
  }
}

/**
 * Get current values for selected ID
 * @param {string} type - common or item level attributes
 * @param {string} id - Selected ID value
 * @return {Object} Current values for the ID
 */
function getCurrentValues(type, id) {
  try {
    if (!type || !id) throw new Error('Invalid parameters');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(SHEET_CONFIG.DATA.NAME);

    // Get column config based on type
    const columnConfig = type === 'COMMON' ?
      ATTRIBUTE_CONFIG.COMMON_LEVEL_FIELDS :
      ATTRIBUTE_CONFIG.ITEM_LEVEL_FIELDS;

    // Get ID column index from config
    const idField = type === 'COMMON' ? 'COMMON ID' : 'ITEM ID';
    const idColIndex = SHEET_CONFIG.DATA.COLUMNS.INDEX[idField] - 1;

    // Get data range
    const headers = SHEET_CONFIG.DATA.COLUMNS.HEADERS;
    const lastRow = dataSheet.getLastRow();
    const searchRange = dataSheet.getRange(
      SHEET_CONFIG.DATA.POSITIONS.DATA_START_ROW,
      SHEET_CONFIG.DATA.POSITIONS.START_COL,
      lastRow - 1,
      headers.length
    );

    const values = searchRange.getDisplayValues();
    const rowIndex = values.findIndex(row => row[idColIndex] === id);

    if (rowIndex === -1) return null;

    // Create result object
    const result = {};
    columnConfig.forEach(field => {
      const colIndex = SHEET_CONFIG.DATA.COLUMNS.INDEX[field] - 1;
      if (colIndex !== undefined) {
        const key = field.toLowerCase().replace(/\s+/g, '-');
        result[key] = values[rowIndex][colIndex];
      }
    });

    return result;

  } catch (error) {
    Logger.log(error.stack);
    throw new Error(`Failed to get current values: ${error.message}`);
  }
}

/**
* Submit attribute update request and update sheet with pending changes
* @param {Object} formData Form data containing type, id and attributes
* @returns {boolean} True if submission successful
*/
function submitAttributeRequest(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(SHEET_CONFIG.DATA.NAME);
    const inputSheet = ss.getSheetByName(SHEET_CONFIG.INPUT.NAME);
    const requestsSheet = ss.getSheetByName(SHEET_CONFIG.REQUESTS.NAME);
    const headers = SHEET_CONFIG.DATA.COLUMNS.HEADERS;

    const timestamp = new Date().toISOString();
    const userEmail = Session.getActiveUser().getEmail();
    const idColName = formData.type === 'COMMON' ? 'COMMON ID' : 'ITEM ID';
    const idColIndex = headers.indexOf(idColName);

    Object.entries(formData.attributes).forEach(([key, newValue]) => {
      const headerName = key.replace(/-/g, ' ').toUpperCase();
      const colIndex = headers.indexOf(headerName);

      // Get first matching row for requests tab
      const data = dataSheet.getDataRange().getValues();
      const firstMatchingRow = data.find(row => row[idColIndex] === formData.id);

      if (firstMatchingRow && firstMatchingRow[colIndex] !== newValue) {
        // Single request entry
        requestsSheet.appendRow([
          timestamp,
          userEmail,
          'Pending',
          formData.id,
          formData.type,
          headerName,
          firstMatchingRow[colIndex] || '',
          newValue || '',
          formData.notes || ''
        ]);

        // Apply changes to all matching rows in both sheets
        [dataSheet, inputSheet].forEach(sheet => {
          const sheetData = sheet.getDataRange().getValues();
          sheetData.forEach((row, rowIndex) => {
            if (row[idColIndex] === formData.id) {
              // Highlight ID column
              sheet.getRange(rowIndex + 1, idColIndex + 1).setBackground('#ffe066');

              // Note and highlight for changed column
              const range = sheet.getRange(rowIndex + 1, colIndex + 1);
              const note = `User: ${userEmail}\nOld Value: ${row[colIndex]}\nNew Value:¬ ${newValue}${
                formData.notes ? '\n\nNotes: ' + formData.notes : ''
              }`;
              range.setNote(note).setBackground('#ffe066');
            }
          });
        });
      }
    });

    SpreadsheetApp.flush();
    return true;

  } catch (error) {
    Logger.log('Error in submitAttributeRequest: ' + error.stack);
    throw new Error('Failed to submit request: ' + error.message);
  }
}

/**
 * Add a change request to the tracking sheet with validation and access control
 * @param {Object} formData - Form data containing change request details
 * @param {string} formData.attributeLevel - Level of attribute ('Item' or 'Common')
 * @param {string} formData.id - Item ID or Common ID based on attribute level
 * @param {string} formData.attribute - Name of the attribute being changed
 * @param {string} formData.currentValue - Current value of the attribute
 * @param {string} formData.requestedValue - Requested new value
 * @returns {Object} Response object with status and message
 */
function addChangeRequest(formData) {
  try {
    // Validate user access
    if (!checkAccess('User')) {
      throw new Error('You do not have permission to submit change requests');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = ss.getSheetByName('Change Requests') || createRequestsSheet(ss);

    // Get current user and timestamp
    const user = Session.getActiveUser().getEmail();
    const timestamp = new Date();

    // Validate required fields
    if (!formData.attributeLevel || !formData.id || !formData.attribute ||
        !formData.currentValue || !formData.requestedValue) {
      throw new Error('All fields are required');
    }

    // Add new request row
    requestsSheet.appendRow([
      timestamp,              // Timestamp
      user,                   // User email
      'Pending',             // Initial status
      formData.attributeLevel,// Attribute level (Item/Common)
      formData.id,           // ID
      formData.attribute,    // Attribute name
      formData.currentValue, // Current value
      formData.requestedValue// Requested value
    ]);

    // Apply formatting to the new row
    const lastRow = requestsSheet.getLastRow();
    formatNewRequestRow(requestsSheet, lastRow);

    return {
      status: 'success',
      message: 'Change request submitted successfully'
    };

  } catch (error) {
    Logger.log(error.stack);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * Applies formatting to a newly added request row
 * @param {SpreadsheetApp.Sheet} sheet - The requests sheet
 * @param {number} rowNum - Row number to format
 */
function formatNewRequestRow(sheet, rowNum) {
  const range = sheet.getRange(rowNum, 1, 1, 8);

  // Set date format for timestamp
  sheet.getRange(rowNum, 1).setNumberFormat('MM/dd/yyyy HH:mm:ss');

  // Set borders and alignment
  range
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Set status cell background
  sheet.getRange(rowNum, 3)
    .setBackground('#ffd666')
    .setFontWeight('bold');
}

/**
 * Get pending change requests
 * @return {Array} Array of pending requests
 */
function getPendingRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestsSheet = ss.getSheetByName('Requests');

  if (!requestsSheet) return [];

  const data = requestsSheet.getDataRange().getValues();
  const headers = data[0];

  return data.slice(1)
    .filter(row => row[headers.indexOf('STATUS')] === 'Pending')
    .map(row => ({
      timestamp: row[headers.indexOf('TIMESTAMP')],
      type: row[headers.indexOf('ATTRIBUTE LEVEL')],
      id: row[headers.indexOf('ID')],
      attribute: row[headers.indexOf('ATTRIBUTE')],
      currentValue: row[headers.indexOf('CURRENT VALUE')],
      requestedValue: row[headers.indexOf('REQUESTED VALUE')]
    }));
}

/**
 * Handle change request approval/rejection
 * @param {number} index - Request index
 * @param {string} action - 'approve' or 'reject'
 */
function handleChangeRequest(index, action) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestsSheet = ss.getSheetByName('Change Requests');
  const dataSheet = ss.getSheetByName('Data');

  // Get request data
  const data = requestsSheet.getDataRange().getValues();
  const headers = data[0];
  const requests = data.slice(1).filter(row => row[headers.indexOf('Status')] === 'Pending');
  const request = requests[index];

  // Update request status
  const requestRow = data.indexOf(request) + 1;
  requestsSheet.getRange(requestRow, headers.indexOf('Status') + 1)
    .setValue(action === 'approve' ? 'Approved' : 'Rejected');

  if (action === 'approve') {
    // Update data
    const dataHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    const idCol = dataHeaders.indexOf(request[headers.indexOf('Type')] === 'COMMON' ? 'Common ID' : 'Item ID');
    const attrCol = dataHeaders.indexOf(request[headers.indexOf('Attribute')]);

    const dataRows = dataSheet.getDataRange().getValues();
    dataRows.slice(1).forEach((row, index) => {
      if (row[idCol] === request[headers.indexOf('ID')]) {
        const range = dataSheet.getRange(index + 2, attrCol + 1);
        range.setValue(request[headers.indexOf('Requested Value')]);
        range.setBackground(null);
        range.clearNote();
      }
    });
  } else {
    // Add rejection note
    const dataHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    const idCol = dataHeaders.indexOf(request[headers.indexOf('Type')] === 'COMMON' ? 'Common ID' : 'Item ID');
    const attrCol = dataHeaders.indexOf(request[headers.indexOf('Attribute')]);

    const dataRows = dataSheet.getDataRange().getValues();
    dataRows.slice(1).forEach((row, index) => {
      if (row[idCol] === request[headers.indexOf('ID')]) {
        const range = dataSheet.getRange(index + 2, attrCol + 1);
        range.setBackground(null);
        range.setNote('Change request rejected by admin. Please submit a new request if needed.');
      }
    });

    // Send email notification
    const userEmail = Session.getActiveUser().getEmail();
    const subject = 'Change Request Rejected';
    const message = `Your request to change ${request[headers.indexOf('Attribute')]} for ${request[headers.indexOf('ID')]} has been rejected. Please submit a new request if needed.`;

    MailApp.sendEmail(userEmail, subject, message);
  }
}