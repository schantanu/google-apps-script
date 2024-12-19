// Sheet Functions
// ----------------------------------------------------------------------------------------------------------
//
// updateFont()                                   - Set default font for the whole Spreadsheet
// setFontSize(sheetName)                         - Set font size for a sheet
// clearSheetContents(sheetName)                  - Clear only the data contents of a sheet, not the formatting.
// hideSheet(sheetName)                           - Hide sheet
// activateCell(sheetName, cell)                  - Activate a given cell in a sheet.
// getColumnLastRow(sheetName, column, rowStart)  - Get the last empty row in a column of a given sheet.
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
    const sheet = ss.getSheetByName(sheetName);

    // Insert new sheet if not exists
    if (!sheet) {
      ss.insertSheet(sheetName);

      // Update protection
      updateSheetsProtection();
    } else {
      // // Clear formatting and data
      // sheet.clear();

      // Clear based on clearType parameter
      switch(clearType) {
        case 'all':
          // Clear formatting and data
          sheet.clear();
          break;
        case 'formats':
        default:
          // Clear formatting only
          sheet.clearFormats();
      }
    }

    // If filter exists then remove
    if(sheet.getFilter()) {
      sheet.getFilter().remove();
    }

    // Set Font Size
    setFontSize(sheetName);
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
    const range = sheet.getRange(rowStart, column, lastRow);
    const values = range.getValues();

    // Reverse the array
    const reversedValues = values.reverse();
    const offset = reversedValues.findIndex(c => c[0] !== '');

    if (offset === -1) {
      return rowStart - 1;
    }

    // Get column last row value
    const columnLastRow = ((lastRow + rowStart) - offset) ;

    return columnLastRow;
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Show sidebar form.
 * @param {string} sidebarHtml - The html element to show as sidebar.
 * @param {string} sidebarTitle - The title of the sidebar.
 */
function showSidebar(sidebarHtml, sidebarTitle) {
  try {
    // Show Sidebar
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile(sidebarHtml).setTitle(sidebarTitle);
    ui.showSidebar(html);
  } catch (error) {
    Logger.log(error.stack);
  }
}