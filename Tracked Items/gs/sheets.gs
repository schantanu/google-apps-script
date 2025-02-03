// Sheet Functions
// ----------------------------------------------------------------------------------------------------------
//
// updateFont()                                   - Set default font for the whole Spreadsheet.
// setFontSize(sheetName)                         - Set font size for a sheet.
// resetSheet(sheetName, clearType = 'formats')   - Insert sheet, if it does not exist else clear formatting and data.
// clearSheetContents(sheetName)                  - Clear only the data contents of a sheet, not the formatting.
// hideSheet(sheetName)                           - Hide sheet
// activateCell(sheetName, cell)                  - Activate a given cell in a sheet.
// getColumnLastRow(sheetName, column, rowStart)  - Get the last empty row in a column of a given sheet.
// showSidebar(sidebarHtml, sidebarTitle)         - Show sidebar form.
//
// ----------------------------------------------------------------------------------------------------------

/**
 * Update the Font for the whole Spreadsheet.
 * Note: Need to execute only once.
 */
function updateFont() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const newFont = 'IBM Plex Sans';

    // Set Font
    ss.getSpreadsheetTheme().setFontFamily(newFont);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Set the Font Size for a given sheet.
 * @param {string} sheetName - The sheet to change the font size of.
 */
function setFontSize(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    const fontSize = 9;

    // Set Font Size for whole sheet
    const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    range.setFontSize(fontSize);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Insert sheet, if it does not exist else clear formatting and data.
 * @param {string} sheetName - The sheet name.
 */
function resetSheet(sheetName, clearType = 'formats') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    // Insert new sheet if not exists
    if (!sheet) {
      ss.insertSheet(sheetName);
      sheet = ss.getSheetByName(sheetName);

      // Update protection
      // updateSheetsProtection(); <------------------------------------------------------
    } else {
      // Clear based on clearType parameter
      switch(clearType) {
        case 'all':
          // Clear formatting and data
          sheet.clear();
          break;
        case 'formats':
        default:
          // Get sheet config for data start row
          const sheetConfig = SHEET_CONFIG[sheetName.toUpperCase()];
          if (!sheetConfig) throw new Error(`Sheet config not found for ${sheetName}`);
          const defaultDataStartRow = 2;
          const dataStartRow = sheetConfig.POSITIONS.DATA_START_ROW || defaultDataStartRow;

          // Clear formatting only for header section
          sheet.getRange(1, 1, dataStartRow - 1, sheet.getMaxColumns()).clearFormat();
      }
    }

    // If filter exists then remove
    if(sheet.getFilter()) {
      sheet.getFilter().remove();
    }

    // Set Font Size
    setFontSize(sheetName);

    return sheet;
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Clear only the data contents of a sheet, not the formatting.
 * @param {string} sheetName - The sheet name.
 */
function clearSheetContents(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    // Throw error if sheet not found
    if (!sheet) throw new Error('Sheet with name ' + sheetName + ' not found.');

    // Get sheet data range
    const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

    // Clear range contents only
    range.clearContent();
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Hide sheet.
 * @param {string} sheetName - The sheet name.
 */
function hideSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    // Hide sheet
    sheet.hideSheet();
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Activate a given cell in a sheet.
 * @param {string} sheetName - The sheet name.
 * @param {string} cell - The range in A1 notation, default = 'A1'
 */
function activateCell(sheetName, cell='A1') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    // Activate cell
    sheet.getRange(cell).activate();
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Get the last empty row in a column of a given sheet.
 * @param {string} role - The sheet name.
 * @param {integer} column - The column.
 * @param {integer} rowStart - The row from where the data starts.
 */
function getColumnLastRow(sheetName, column, rowStart) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    // Get range values
    const lastRow = sheet.getLastRow();

    // Calculate the correct range height
    const rangeHeight = lastRow - rowStart + 1;

    // Get range values with correct height
    const range = sheet.getRange(rowStart, column, rangeHeight);
    const values = range.getValues();

    // Reverse the array and find the first non-empty cell
    const reversedValues = values.reverse();
    const offset = reversedValues.findIndex(c => c[0] !== '');

    if (offset === -1) {
      return rowStart;
    }

    // Get last filled row
    const lastFilledRow = lastRow - offset;

    // Return the next empty row
    return lastFilledRow + 1;
  } catch (error) {
    Logger.log(error.stack);
    return rowStart;
  }
}

/**
 * Toggles filter view on/off for a specified sheet
 * @param {string} sheetName - Name of the sheet to toggle filter on
 * @param {boolean} filterState - true to apply filter, false to remove filter
 * @returns {void}
 * @throws {Error} If sheet operations fail
 */
function toggleSheetFilter(sheetName, headerRow, filterState) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    // Remove existing filter if any
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }

    // Add new filter if requested and sheet has headers
    if (filterState) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const dataRange = sheet.getRange(headerRow, 1, lastRow-(headerRow-1), lastCol);
      if (!dataRange.isBlank()) {
        const lastCol = dataRange.getLastColumn();
        dataRange.createFilter();
      }
    }
  } catch (error) {
    Logger.log(`Error setting filter:\n${error.stack}`);
    throw error;
  }
}