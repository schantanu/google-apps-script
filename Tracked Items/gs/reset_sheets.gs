/**
 * Setup the 'Admin' sheet
 */
function setupSheetAdmin() {
  try {
    // Reset sheet
    resetSheet(adminSheetName);

    // Get sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(adminSheetName);

    // Add Role Management columns
    sheet.getRange(1, 1, 1, 2).mergeAcross();
    sheet.getRange('A1').setValue('Users List');
    sheet.getRange('A2:B2').setValues([['Admins','Users']]);

    // Format Role Management columns
    sheet.getRange('A1').setBackground('#FFFF00');
    sheet.getRange('A2:B2').setBackground('#C9DAF8');
    sheet.getRange('A1:B2').setHorizontalAlignment('center');
    sheet.getRange('A1:B2').setFontWeight('Bold');
    sheet.setColumnWidths(1, 2, 200);

    // Add Attribute dropdown columns
    sheet.getRange(1, 3, 1, adminColDropdownHeaders.length).mergeAcross();
    sheet.getRange(1, 3, 1, 1).setValue('Dropdown Attributes');
    sheet.getRange(2, 3, 1, adminColDropdownHeaders.length).setValues([adminColDropdownHeaders]);

    // Format Attribute dropdown columns
    sheet.getRange('C1').setBackground('#B6D7A8');
    sheet.getRange(2, 3, 1, adminColDropdownHeaders.length).setBackground('#C9DAF8');
    sheet.getRange(1, 3, 2, adminColDropdownHeaders.length).setHorizontalAlignment('center');
    sheet.getRange(1, 3, 2, adminColDropdownHeaders.length).setFontWeight('Bold');
    sheet.getRange(3, 3, sheet.getLastRow(), adminColDropdownHeaders.length).setHorizontalAlignment('left');
    sheet.setColumnWidths(3, adminColDropdownHeaders.length, 120);
    sheet.setFrozenRows(2);

    // Hide Sheet
    // hideSheet(adminSheetName);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Common function to setup and format sheets
 * @param {string} sheetName - Name of the sheet to setup
 * @param {Object} config - Configuration object with sheet-specific settings
 */
function setupSheet(sheetName, config) {
  try {
    // Reset sheet
    resetSheet(sheetName);

    // Get sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    // // Add empty rows
    // const currentRows = sheet.getMaxRows();
    // const rowsToAdd = 15000 - currentRows;
    // if (rowsToAdd > 0) {
    //   sheet.insertRowsAfter(currentRows, rowsToAdd);
    // }

    // Format header row
    const headerRange = sheet.getRange(config.startRow, config.startCol, config.headerRow, dataColHeaders.length);
    headerRange.setFontWeight('Bold')
               .setBackground('#C9DAF8')
               .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    // Format user editable columns
    const userEditableRange = sheet.getRange(config.startRow, dataColShortName, config.headerRow, (dataColLastUpdated - dataColShortName));
    userEditableRange.setBackground('#FFE599');

    // Set header row height
    sheet.setRowHeight(config.headerRow, 50);

    // Apply column widths
    const columnWidths = getColumnWidths();
    Object.entries(columnWidths).forEach(([col, width]) => {
      sheet.setColumnWidth(col, width);
    });

    // Format columns
    const maxRows = sheet.getMaxRows();
    sheet.getRange(config.startRow, config.startCol, maxRows, dataColLastUpdated).setNumberFormat('@');
    // sheet.getRange(r,c,nr,nc);

    // Format date columns
    const dateColumns = [dataColPlannedStartDate, dataColPlannedEndDate, dataColBudgetStartDate, dataColBudgetEndDate];
    dateColumns.forEach(col => {
      sheet.getRange(config.startRow, col, maxRows).setNumberFormat('dd-mmm-yy');
    });

    // Format data rows
    sheet.getRange(config.startRow + 1, config.startCol, maxRows, dataColLastUpdated)
         .setHorizontalAlignment('left')
         .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // Freeze header row
    sheet.setFrozenRows(config.headerRow);

    // Sheet-specific operations
    if (config.additionalSetup) {
      config.additionalSetup(sheet, config);
    }

    // Add filter
    toggleSheetFilter(sheetName, config.headerRow, true);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Get column width configurations
 * @return {Object} Column width mappings
 */
function getColumnWidths() {
  return {
    [dataColItemId]           : 150,
    [dataColCommonId]         : 150,
    [dataColForecastedS4MMId] : 105,
    [dataColShortName]        : 200,
    [dataColCommonDesc]       : 200,
    [dataColItemDesc]         : 200,
    [dataColVendor]           : 150,
    [dataColMaterialGroup]    : 90,
    [dataColGroupFunction]    : 100,
    [dataColFamily]           : 100,
    [dataColSubFamily]        : 100,
    [dataColPlannedStartDate] : 82,
    [dataColPlannedEndDate]   : 82,
    [dataColBudgetLineItem]   : 200,
    [dataColBudgetLineItem2]  : 150,
    [dataColBudgetStartDate]  : 82,
    [dataColBudgetEndDate]    : 82,
    [dataColParentCommonId]   : 150,
    [dataColDPHierarchy]      : 100,
    [dataColTrackedSet]       : 100,
    [dataColFreq]             : 100,
    [dataColPower]            : 100,
    [dataColIntegrated]       : 110,
    [dataColTech]             : 100,
    [dataColPlanner]          : 100,
    [dataColLastUpdatedBy]    : 200,
    [dataColLastUpdated]      : 150,
  };
}

/**
 * Setup the 'Data' sheet
 */
function setupSheetData() {
  setupSheet('Data', {
    startRow: dataStartRow,
    startCol: dataStartCol,
    headerRow: dataHeaderRow,
    additionalSetup: (sheet, config) => {
      sheet.getRange(config.startRow, config.startCol, config.headerRow, dataColHeaders.length)
           .setValues([dataColHeaders]);
    }
  });
}

/**
 * Setup the 'Input' sheet
 */
function setupSheetInput() {
  setupSheet('Input', {
    startRow: inputStartRow,
    startCol: inputStartCol,
    headerRow: inputHeaderRow,
    additionalSetup: (sheet, config) => {
      const importCell = sheet.getRange('A' + config.headerRow);
      importCell.setFormula(inputArrayFormula);
    }
  });
}

/**
 * Creates and initializes the Change Requests sheet with proper formatting
 * @param {SpreadsheetApp.Spreadsheet} ss - Active spreadsheet
 * @returns {SpreadsheetApp.Sheet} The created requests sheet
 */
function setupSheetRequests() {
  // Reset Sheet
  resetSheet(requestsSheetName);

  // Get sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(requestsSheetName);

  // Set up header row
  const headers = ['Timestamp','User','Status','ID','Attribute Level','Attribute','Current Value','Requested Value'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // Apply header formatting
  headerRange
    .setBackground('#C9DAF8')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Auto-resize columns
  // sheet.autoResizeColumns(1, headers.length);

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 70);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 200);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Add filter
  toggleSheetFilter(requestsSheetName, 1, true);
}

/**
 * Setup the 'Data' sheet
 */
// function setupSheetData() {
//   try {
//     // Reset sheet
//     resetSheet(dataSheetName);

//     // Get sheet
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName(dataSheetName);

//     // Add empty rows
//     const currentRows = sheet.getMaxRows();
//     const rowsToAdd = 15000 - currentRows;
//     if (rowsToAdd > 0) {
//       sheet.insertRowsAfter(currentRows, rowsToAdd);
//     }

//     // Set Headers
//     sheet.getRange(dataStartRow, dataStartCol, dataHeaderRow, dataColHeaders.length).setValues([dataColHeaders]);

//     // Freeze top row
//     sheet.setFrozenRows(dataHeaderRow);

//     // Format Header row
//     headerRange = sheet.getRange(dataStartRow, dataStartCol, dataHeaderRow, dataColLastUpdated);
//     headerRange.setFontWeight('Bold');
//     headerRange.setBackground('#C9DAF8');
//     headerRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

//     // Show User editable columns
//     userEditableRange = sheet.getRange(dataStartRow, dataColShortName, dataHeaderRow, (dataColLastUpdated - dataColShortName));
//     userEditableRange.setBackground('#FFE599');

//     // Set Header row height
//     sheet.setRowHeight(dataHeaderRow, 50);

//     // Set column widths
//     sheet.setColumnWidth(dataColItemId            , 150);
//     sheet.setColumnWidth(dataColCommonId          , 150);
//     sheet.setColumnWidth(dataColForecastedS4MMId  , 105);
//     sheet.setColumnWidth(dataColShortName         , 200);
//     sheet.setColumnWidth(dataColCommonDesc        , 200);
//     sheet.setColumnWidth(dataColItemDesc          , 200);
//     sheet.setColumnWidth(dataColVendor            , 150);
//     sheet.setColumnWidth(dataColMaterialGroup     , 90);
//     sheet.setColumnWidth(dataColGroupFunction     , 100);
//     sheet.setColumnWidth(dataColFamily            , 100);
//     sheet.setColumnWidth(dataColSubFamily         , 100);
//     sheet.setColumnWidth(dataColPlannedStartDate  , 82);
//     sheet.setColumnWidth(dataColPlannedEndDate    , 82);
//     sheet.setColumnWidth(dataColBudgetLineItem    , 200);
//     sheet.setColumnWidth(dataColBudgetLineItem2   , 150);
//     sheet.setColumnWidth(dataColBudgetStartDate   , 82);
//     sheet.setColumnWidth(dataColBudgetEndDate     , 82);
//     sheet.setColumnWidth(dataColParentCommonId    , 150);
//     sheet.setColumnWidth(dataColDPHierarchy       , 100);
//     sheet.setColumnWidth(dataColTrackedSet        , 100);
//     sheet.setColumnWidth(dataColFreq              , 100);
//     sheet.setColumnWidth(dataColPower             , 100);
//     sheet.setColumnWidth(dataColIntegrated        , 110);
//     sheet.setColumnWidth(dataColTech              , 100);
//     sheet.setColumnWidth(dataColPlanner           , 100);
//     sheet.setColumnWidth(dataColLastUpdatedBy     , 200);
//     sheet.setColumnWidth(dataColLastUpdated       , 150);

//     // Format 'Last Updated' column to 'Plain Text' format
//     const maxRows = sheet.getMaxRows();
//     sheet.getRange(dataStartRow, dataStartCol, dataColLastUpdated, maxRows).setNumberFormat('@');

//     // Format date columns to 'dd-mmm-yy' format
//     const dateColsToFormat = [dataColPlannedStartDate, dataColPlannedEndDate, dataColBudgetStartDate, dataColBudgetEndDate];
//     dateColsToFormat.forEach(function(col) {
//       sheet.getRange(dataStartRow, col, maxRows).setNumberFormat('dd-mmm-yy');
//     });

//     // Format data rows
//     sheet.getRange(dataStartRow + 1, dataStartCol, maxRows, dataColLastUpdated).setHorizontalAlignment('left');
//     sheet.getRange(dataStartRow + 1, dataStartCol, maxRows, dataColLastUpdated).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

//     // Add filter
//     toggleSheetFilter(dataSheetName, dataHeaderRow, true);
//   } catch (error) {
//     Logger.log(error.stack);
//   }
// }

// /**
//  * Setup the 'Input' sheet
//  */
// function setupSheetInput() {
//   try {
//     // Reset sheet
//     resetSheet(inputSheetName);

//     // Get sheet
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName(inputSheetName);

//     // Set Pivot formula
//     const importCell = sheet.getRange('A' + inputHeaderRow);
//     const importFormula = `=ARRAYFORMULA(Data!A1:AA)`;
//     importCell.setFormula(importFormula);
//   } catch (error) {
//     Logger.log(error.stack);
//   }
// }