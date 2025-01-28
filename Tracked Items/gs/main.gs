/**
 * Get attribute data configuration based on type (common or item level)
 * @param {string} type - The type of attributes to get ('common' or 'item')
 * @returns {Object} Object containing ids, attributes configuration, and dropdown values
 */
function getAttributeData(type) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName('Data');
    const adminSheet = ss.getSheetByName('Admin');
    const data = dataSheet.getDataRange().getValues();
    const headers = data[0];

    // Define column configurations
    const columnConfig = {
      common: [
        'SHORT NAME',
        'COMMON DESCRIPTION',
        'VENDOR',
        'GROUP FUNCTION',
        'FAMILY',
        'SUBFAMILY',
        'PLANNED START DATE',
        'PLANNED END DATE',
        'BUDGET START DATE',
        'BUDGET END DATE',
        'POWER'
      ],
      item: [
        'ITEM DESCRIPTION'
      ]
    };

    // Get column indices
    const itemIdCol = headers.indexOf('ITEM ID');
    const commonIdCol = headers.indexOf('COMMON ID');

    // Get columns for selected type
    const selectedColumns = columnConfig[type];

    // Get unique IDs based on type
    const ids = [...new Set(data.slice(1).map(row => row[type === 'common' ? commonIdCol : itemIdCol]))];

    // Configure attributes based on column names
    const attributes = selectedColumns.map(header => {
      const colIndex = headers.indexOf(header);
      return {
        name: header,
        columnIndex: colIndex,
        type: header.toLowerCase().includes('date') ? 'date' :
              ['VENDOR', 'GROUP FUNCTION', 'FAMILY', 'SUBFAMILY'].includes(header) ? 'dropdown' : 'text',
        maxLength: 100
      };
    });

    // Get dropdown values from Admin sheet if needed
    const dropdowns = {};
    if (type === 'common') {
      // Get dropdown values for each category from Admin sheet
      const dropdownColumns = {
        'VENDOR': 3,        // Column C
        'GROUP FUNCTION': 4, // Column D
        'FAMILY': 5,        // Column E
        'SUBFAMILY': 6      // Column F
      };

      Object.entries(dropdownColumns).forEach(([field, col]) => {
        const values = adminSheet.getRange(3, col, adminSheet.getLastRow()).getValues()
          .map(row => row[0])
          .filter(Boolean);
        dropdowns[field] = values;
      });
    }

    return {
      ids: ids,
      attributes: attributes,
      dropdowns: dropdowns
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
    if (!type || !id) {
      throw new Error('Invalid parameters');
    }

    // Define column configurations
    const columnConfig = {
      common: [
        'SHORT NAME',
        'COMMON DESCRIPTION',
        'VENDOR',
        'GROUP FUNCTION',
        'FAMILY',
        'SUBFAMILY',
        'PLANNED START DATE',
        'PLANNED END DATE',
        'BUDGET START DATE',
        'BUDGET END DATE',
        'POWER'
      ],
      item: [
        'ITEM DESCRIPTION'
      ]
    };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(dataSheetName);
    const data = dataSheet.getDataRange().getDisplayValues();
    const headers = data[0];

    // Get column indices
    const idCol = headers.indexOf(type === 'common' ? 'COMMON ID' : 'ITEM ID');
    const selectedColumns = columnConfig[type];

    // Find row with matching ID
    const row = data.find(row => row[idCol] === id);
    if (!row) return null;

    // Create result object with current values
    const result = {};
    selectedColumns.forEach(header => {
      const colIndex = headers.indexOf(header);
      if (colIndex !== -1) {
        // Convert header to lowercase with hyphens for object key
        const key = header.toLowerCase().replace(/\s+/g, '-');
        result[key] = row[colIndex];
      }
    });

    Logger.log('Result for ID ' + id + ': ' + JSON.stringify(result));
    return result;

  } catch (error) {
    Logger.log('Error in getCurrentValues: ' + error.stack);
    throw new Error('Failed to get current values: ' + error.message);
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
    const dataSheet = ss.getSheetByName('Data');
    const requestsSheet = ss.getSheetByName('Requests');
    const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];

    const timestamp = new Date().toISOString();
    const userEmail = Session.getActiveUser().getEmail();

    const dateColumns = [
      'PLANNED START DATE',
      'PLANNED END DATE',
      'BUDGET START DATE',
      'BUDGET END DATE'
    ];

    const data = dataSheet.getDataRange().getValues();
    const idCol = headers.indexOf(formData.type === 'common' ? 'COMMON ID' : 'ITEM ID');
    const rowIndex = data.findIndex(row => row[idCol] === formData.id);
    if (rowIndex === -1) throw new Error('ID not found');

    Object.entries(formData.attributes).forEach(([key, newValue]) => {
      const attrName = key.replace(/-/g, ' ').toUpperCase();
      const colIndex = headers.indexOf(attrName);
      const currentValue = data[rowIndex][colIndex];

      // Only process if value has changed
      if (currentValue !== newValue) {
        // Handle null dates
        if (dateColumns.includes(attrName) && (!newValue || newValue.toLowerCase() === 'null')) {
          newValue = '';
        }

        // Add note and highlight
        const range = dataSheet.getRange(rowIndex + 1, colIndex + 1);
        range.setNote(`User: ${userEmail}\nChange: ${currentValue} → ${newValue}`)
              .setBackground('#ffe066');

        // Add to requests sheet
        requestsSheet.appendRow([
          timestamp,
          userEmail,
          'Pending',
          formData.id,
          formData.type,
          attrName,
          currentValue || '',
          newValue || ''
        ]);
      }
    });

    SpreadsheetApp.flush();
    return true;

  } catch (error) {
    Logger.log('Error in submitAttributeRequest: ' + error.stack);
    throw new Error('Failed to submit request: ' + error.message);
  }
}

// /**
//  * Submit attribute update request
//  * @param {Object} formData Form data containing type, id and attributes
//  */
// function submitAttributeRequest(formData) {
//   try {
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const changelogSheet = ss.getSheetByName('Changelog');

//     // Get current timestamp and user email
//     const timestamp = new Date().toISOString();
//     const userEmail = Session.getActiveUser().getEmail();

//     // Format changelog entry
//     const changelogRow = [
//       timestamp,
//       userEmail,
//       formData.type,
//       formData.id,
//       JSON.stringify(formData.attributes),
//       'Pending'
//     ];

//     // Append to changelog
//     changelogSheet.appendRow(changelogRow);

//     return true;
//   } catch (error) {
//     Logger.log('Error in submitAttributeRequest: ' + error.stack);
//     throw new Error('Failed to submit request: ' + error.message);
//   }
// }

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
  const requestsSheet = ss.getSheetByName('Change Requests');

  if (!requestsSheet) return [];

  const data = requestsSheet.getDataRange().getValues();
  const headers = data[0];

  return data.slice(1)
    .filter(row => row[headers.indexOf('Status')] === 'Pending')
    .map(row => ({
      timestamp: row[headers.indexOf('Timestamp')],
      type: row[headers.indexOf('Type')],
      id: row[headers.indexOf('ID')],
      attribute: row[headers.indexOf('Attribute')],
      currentValue: row[headers.indexOf('Current Value')],
      requestedValue: row[headers.indexOf('Requested Value')]
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
    const idCol = dataHeaders.indexOf(request[headers.indexOf('Type')] === 'common' ? 'Common ID' : 'Item ID');
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
    const idCol = dataHeaders.indexOf(request[headers.indexOf('Type')] === 'common' ? 'Common ID' : 'Item ID');
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