/**
 * Creates and initializes the Admin sheet with formatting
 */
function setupSheetAdmin() {
  try {
    // Reset sheet
    const sheet = resetSheet(SHEET_CONFIG.ADMIN.NAME);

    // Set Headers
    sheet.getRange(SHEET_CONFIG.ADMIN.RANGES.USERS_HEADER)
      .setValue(SHEET_CONFIG.ADMIN.HEADERS.USERS_HEADER)
      .setBackground(SHEET_CONFIG.ADMIN.COLORS.USERS_HEADER)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .mergeAcross();

    // Set Subheaders
    sheet.getRange(SHEET_CONFIG.ADMIN.RANGES.USERS_SUBHEADER)
      .setValues(SHEET_CONFIG.ADMIN.HEADERS.USERS_SUBHEADER)
      .setBackground(SHEET_CONFIG.ADMIN.COLORS.USERS_SUBHEADER)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // Set Column widths
    sheet.setColumnWidths(
      SHEET_CONFIG.ADMIN.COLUMNS.ADMINS,
      SHEET_CONFIG.ADMIN.COLUMNS.USERS,
      SHEET_CONFIG.ADMIN.COLUMN_WIDTH
    );

    // Freeze header row
    sheet.setFrozenRows(SHEET_CONFIG.ADMIN.POSITIONS.SUBHEADER_ROW);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Creates and initializes the Dropdowns sheet with formatting
 */
function setupSheetDropdowns() {
  try {
    // Reset Sheet
    const sheet = resetSheet(SHEET_CONFIG.DROPDOWNS.NAME);

    // Get header and data ranges
    const headers = SHEET_CONFIG.DROPDOWNS.HEADERS;
    const headerRange = sheet.getRange(
      SHEET_CONFIG.DROPDOWNS.POSITIONS.START_ROW,
      SHEET_CONFIG.DROPDOWNS.POSITIONS.START_COL,
      SHEET_CONFIG.DROPDOWNS.POSITIONS.HEADER_ROW,
      headers.length
    );
    const dataRange = sheet.getRange(
      SHEET_CONFIG.DROPDOWNS.POSITIONS.START_ROW,
      SHEET_CONFIG.DROPDOWNS.POSITIONS.START_COL,
      sheet.getMaxRows(),
      sheet.getMaxColumns()
    );

    // Set header and data range formatting
    headerRange
      .setValues([headers])
      .setBackground(SHEET_CONFIG.DROPDOWNS.COLORS.HEADER)
      .setFontWeight('bold')
      .setHorizontalAlignment('left');
    dataRange
      .setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // Set Column widths
    const columnWidths = SHEET_CONFIG.DROPDOWNS.COLUMN_WIDTHS;
    for (let i = 0; i < columnWidths.length; i++) {
      sheet.setColumnWidth(i + 1, columnWidths[i]);
    }

    // Freeze header row
    sheet.setFrozenRows(SHEET_CONFIG.DROPDOWNS.POSITIONS.HEADER_ROW);

    // Add filter
    toggleSheetFilter(SHEET_CONFIG.DROPDOWNS.NAME, SHEET_CONFIG.DROPDOWNS.POSITIONS.HEADER_ROW, true);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Creates and initializes the Change Requests sheet with formatting
 */
function setupSheetRequests() {
  try {
    // Reset Sheet
    const sheet = resetSheet(SHEET_CONFIG.REQUESTS.NAME);

    // Set up header row
    const headers = SHEET_CONFIG.REQUESTS.HEADERS;
    const headerRange = sheet.getRange(
      SHEET_CONFIG.REQUESTS.POSITIONS.START_ROW,
      SHEET_CONFIG.REQUESTS.POSITIONS.START_COL,
      SHEET_CONFIG.REQUESTS.POSITIONS.HEADER_ROW,
      headers.length
    );
    headerRange.setValues([headers]);

    // Set header formatting
    headerRange
      .setBackground(SHEET_CONFIG.REQUESTS.COLORS.HEADER)
      .setFontWeight('bold')
      .setHorizontalAlignment('left');

    // Get data range
    const dataRange = sheet.getRange(
      SHEET_CONFIG.REQUESTS.POSITIONS.DATA_START_ROW,
      SHEET_CONFIG.REQUESTS.POSITIONS.START_COL,
      sheet.getMaxRows(),
      headers.length
    );

    // Apply data formatting
    dataRange.setHorizontalAlignment('left');

    // Set Column widths
    const columnWidths = SHEET_CONFIG.REQUESTS.COLUMN_WIDTHS;
    for (let i = 0; i < columnWidths.length; i++) {
      sheet.setColumnWidth(i + 1, columnWidths[i]);
    }

    // Freeze header row
    sheet.setFrozenRows(SHEET_CONFIG.REQUESTS.POSITIONS.HEADER_ROW);

    // Add filter
    toggleSheetFilter(SHEET_CONFIG.REQUESTS.NAME, SHEET_CONFIG.REQUESTS.POSITIONS.HEADER_ROW, true);
  } catch (error) {
    Logger.log(error.stack);
  }
}

function getColumnWidths() {
  return {
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.ITEM_ID]             : 150,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.COMMON_ID]           : 150,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.FORECASTED_S4MM_ID]  : 150,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.SHORT_NAME]          : 200,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.COMMON_DESC]         : 200,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.ITEM_DESC]           : 200,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.VENDOR]              : 150,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.MATERIAL_GROUP]      : 90,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.GROUP_FUNCTION]      : 100,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.FAMILY]              : 100,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.SUBFAMILY]           : 100,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.PLANNED_START_DATE]  : 82,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.PLANNED_END_DATE]    : 82,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.BUDGET_LINE_ITEM]    : 200,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.BUDGET_LINE_ITEM2]   : 150,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.BUDGET_START_DATE]   : 82,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.BUDGET_END_DATE]     : 82,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.PARENT_COMMON_ID]    : 150,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.DP_HIERARCHY]        : 100,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.TRACKED_SET]         : 100,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.FREQ]                : 100,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.POWER]               : 100,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.INTEGRATED]          : 110,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.TECH]                : 100,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.PLANNER_NAME]        : 150,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.LAST_UPDATED_BY]     : 200,
    [SHEET_CONFIG.DATA.COLUMNS.INDEX.LAST_UPDATED]        : 150,
  };
}

/**
 * Common function to setup and format sheets
 * @param {string} sheetName - Name of the sheet to setup
 * @param {Object} config - Configuration object with sheet-specific settings
 */
function setupSheet(sheetName, config) {
  try {
    // Reset sheet
    const sheet = resetSheet(sheetName);

    // Add empty rows
    // const currentRows = sheet.getMaxRows();
    // const rowsToAdd = 15000 - currentRows;
    // if (rowsToAdd > 0) {
    //   sheet.insertRowsAfter(currentRows, rowsToAdd);
    // }

    // Set header formatting
    const headerRange = sheet.getRange(config.startRow, config.startCol, config.headerRow, config.dataColHeaders.length);
    headerRange
      .setBackground(SHEET_CONFIG.DATA.COLORS.HEADER)
      .setFontWeight('bold')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    // Set header row height
    sheet.setRowHeight(config.headerRow, 50);

    // Format data range
    const maxRows = sheet.getMaxRows();
    sheet.getRange(config.startRow + 1, config.startCol, maxRows, config.dataColHeaders.length)
      .setNumberFormat('@')
      .setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // Format date columns
    const dateColumns = [
      SHEET_CONFIG.DATA.COLUMNS.INDEX.PLANNED_START_DATE,
      SHEET_CONFIG.DATA.COLUMNS.INDEX.PLANNED_END_DATE,
      SHEET_CONFIG.DATA.COLUMNS.INDEX.BUDGET_START_DATE,
      SHEET_CONFIG.DATA.COLUMNS.INDEX.BUDGET_END_DATE,
      SHEET_CONFIG.DATA.COLUMNS.INDEX.LAST_UPDATED
    ];
    dateColumns.forEach(col => {
      sheet.getRange(config.startRow, col, maxRows).setNumberFormat('dd-mmm-yy');
    });

    // Sheet-specific operations
    if (config.additionalSetup) {
      config.additionalSetup(sheet, config);
    }

    // Apply column widths
    const columnWidths = getColumnWidths();
    Object.entries(columnWidths).forEach(([col, width]) => {
      sheet.setColumnWidth(col, width);
    });

    // Freeze header row
    sheet.setFrozenRows(config.headerRow);

    // Add filter
    toggleSheetFilter(sheetName, config.headerRow, true);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Creates and initializes the Data sheet with formatting
 */
function setupSheetData() {
  setupSheet(SHEET_CONFIG.DATA.NAME, {
    startRow: SHEET_CONFIG.DATA.POSITIONS.START_ROW,
    startCol: SHEET_CONFIG.DATA.POSITIONS.START_COL,
    headerRow: SHEET_CONFIG.DATA.POSITIONS.HEADER_ROW,
    dataColHeaders: SHEET_CONFIG.DATA.HEADERS,
    additionalSetup: (sheet, config) => {
      sheet.getRange(config.startRow, config.startCol, config.headerRow, config.dataColHeaders.length)
        .setValues([config.dataColHeaders]);
    }
  });
}

/**
 * Creates and initializes the Input sheet with formatting
 */
function setupSheetInput() {
  setupSheet(SHEET_CONFIG.INPUT.NAME, {
    startRow: SHEET_CONFIG.INPUT.POSITIONS.START_ROW,
    startCol: SHEET_CONFIG.INPUT.POSITIONS.START_COL,
    headerRow: SHEET_CONFIG.INPUT.POSITIONS.HEADER_ROW,
    dataColHeaders: SHEET_CONFIG.DATA.HEADERS,
    additionalSetup: (sheet, config) => {
      sheet.getRange(SHEET_CONFIG.INPUT.IMPORT.CELL)
        .setFormula(SHEET_CONFIG.INPUT.IMPORT.FORMULA);
    }
  });
}