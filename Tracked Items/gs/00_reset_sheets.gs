// /**
//  * Setup the Admin sheet with proper formatting and structure
//  */
// function setupSheetAdmin() {
//   try {
//     // Reset sheet
//     resetSheet(SHEET_CONFIG.ADMIN.NAME);

//     // Get sheet
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName(SHEET_CONFIG.ADMIN.NAME);

//     // Setup Role Management section
//     setupRoleManagementSection(sheet);

//     // Setup Dropdown Attributes section
//     setupDropdownAttributesSection(sheet);

//     // Set frozen rows and column widths
//     sheet.setFrozenRows(2);
//     sheet.setColumnWidths(1, SHEET_CONFIG.ADMIN.USER_COLS_TOTAL, 200);

//   } catch (error) {
//     Logger.log(error.stack);
//   }
// }

// /**
//  * Setup the Role Management section of Admin sheet
//  * @param {SpreadsheetApp.Sheet} sheet - The Admin sheet
//  */
// function setupRoleManagementSection(sheet) {
//   // Merge and set headers
//   sheet.getRange(1, 1, 1, 2).mergeAcross();
//   sheet.getRange('A1').setValue(SHEET_CONFIG.ADMIN.HEADERS.USERS_LIST);
//   sheet.getRange('A2:B2').setValues([['Admins', 'Users']]);

//   // Apply formatting
//   sheet.getRange('A1').setBackground(SHEET_CONFIG.ADMIN.COLORS.USERS_HEADER);
//   sheet.getRange('A2:B2').setBackground(SHEET_CONFIG.ADMIN.COLORS.SUBHEADER);
//   sheet.getRange('A1:B2')
//        .setHorizontalAlignment('center')
//        .setFontWeight('Bold');
// }

// /**
//  * Setup the Dropdown Attributes section of Admin sheet
//  * @param {SpreadsheetApp.Sheet} sheet - The Admin sheet
//  */
// function setupDropdownAttributesSection(sheet) {
//   const startColumn = 3;
//   const headerCount = ADMIN_DROPDOWN_ATTRIBUTES.length;

//   // Merge and set headers
//   sheet.getRange(1, startColumn, 1, headerCount).mergeAcross();
//   sheet.getRange(1, startColumn).setValue(SHEET_CONFIG.ADMIN.HEADERS.DROPDOWN_ATTRS);
//   sheet.getRange(2, startColumn, 1, headerCount).setValues([ADMIN_DROPDOWN_ATTRIBUTES]);

//   // Apply formatting
//   sheet.getRange(1, startColumn).setBackground(SHEET_CONFIG.ADMIN.COLORS.DROPDOWN_HEADER);
//   sheet.getRange(2, startColumn, 1, headerCount).setBackground(SHEET_CONFIG.ADMIN.COLORS.SUBHEADER);
//   sheet.getRange(1, startColumn, 2, headerCount)
//        .setHorizontalAlignment('center')
//        .setFontWeight('Bold');

//   // Set column formatting
//   sheet.getRange(3, startColumn, sheet.getLastRow(), headerCount)
//        .setHorizontalAlignment('left');
//   sheet.setColumnWidths(startColumn, headerCount, 120);
// }

// /**
//  * Common function to setup and format sheets
//  * @param {string} sheetName - Name of the sheet to setup
//  * @param {Object} config - Configuration object with sheet-specific settings
//  */
// function setupSheet(sheetName, config) {
//   try {
//     // Reset sheet
//     resetSheet(sheetName);

//     // Get sheet
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName(sheetName);

//     // Format header row
//     const headerRange = sheet.getRange(
//       config.startRow,
//       config.startCol,
//       config.headerRow,
//       SHEET_CONFIG.DATA.HEADERS.length
//     );

//     formatHeaderRow(headerRange);

//     // Format user editable columns if specified
//     if (config.userEditableRange) {
//       const userEditableRange = sheet.getRange(
//         config.startRow,
//         SHEET_CONFIG.DATA.COLUMNS.SHORT_NAME,
//         config.headerRow,
//         (SHEET_CONFIG.DATA.COLUMNS.LAST_UPDATED - SHEET_CONFIG.DATA.COLUMNS.SHORT_NAME)
//       );
//       userEditableRange.setBackground('#FFE599');
//     }

//     // Apply sheet-specific configurations
//     if (config.setupFunction) {
//       config.setupFunction(sheet, config);
//     }

//     // Common formatting
//     applyCommonFormatting(sheet, config);

//     // Add filter
//     toggleSheetFilter(sheetName, config.headerRow, true);

//   } catch (error) {
//     Logger.log(error.stack);
//   }
// }

// /**
//  * Format the header row of a sheet
//  * @param {SpreadsheetApp.Range} headerRange - The header range to format
//  */
// function formatHeaderRow(headerRange) {
//   headerRange
//     .setFontWeight('Bold')
//     .setBackground('#C9DAF8')
//     .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
// }

// /**
//  * Apply common formatting to a sheet
//  * @param {SpreadsheetApp.Sheet} sheet - The sheet to format
//  * @param {Object} config - Sheet configuration object
//  */
// function applyCommonFormatting(sheet, config) {
//   const maxRows = sheet.getMaxRows();

//   // Set header row height
//   sheet.setRowHeight(config.headerRow, 50);

//   // Apply column widths
//   const columnWidths = getColumnWidths();
//   Object.entries(columnWidths).forEach(([col, width]) => {
//     sheet.setColumnWidth(parseInt(col), width);
//   });

//   // Format date columns if applicable
//   const dateColumns = [
//     SHEET_CONFIG.DATA.COLUMNS.PLANNED_START_DATE,
//     SHEET_CONFIG.DATA.COLUMNS.PLANNED_END_DATE,
//     SHEET_CONFIG.DATA.COLUMNS.BUDGET_START_DATE,
//     SHEET_CONFIG.DATA.COLUMNS.BUDGET_END_DATE
//   ];

//   dateColumns.forEach(col => {
//     sheet.getRange(config.startRow, col, maxRows)
//          .setNumberFormat('dd-mmm-yy');
//   });

//   // Format data rows
//   sheet.getRange(config.startRow + 1, config.startCol, maxRows, SHEET_CONFIG.DATA.COLUMNS.LAST_UPDATED)
//        .setHorizontalAlignment('left')
//        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

//   // Freeze header row
//   sheet.setFrozenRows(config.headerRow);
// }

// /**
//  * Setup the Data sheet
//  */
// function setupSheetData() {
//   setupSheet(SHEET_CONFIG.DATA.NAME, {
//     startRow: SHEET_CONFIG.DATA.START_ROW,
//     startCol: SHEET_CONFIG.DATA.START_COL,
//     headerRow: SHEET_CONFIG.DATA.HEADER_ROW,
//     userEditableRange: true,
//     setupFunction: (sheet, config) => {
//       sheet.getRange(
//         config.startRow,
//         config.startCol,
//         config.headerRow,
//         SHEET_CONFIG.DATA.HEADERS.length
//       ).setValues([SHEET_CONFIG.DATA.HEADERS]);
//     }
//   });
// }

// /**
//  * Setup the Input sheet
//  */
// function setupSheetInput() {
//   setupSheet(SHEET_CONFIG.INPUT.NAME, {
//     startRow: SHEET_CONFIG.INPUT.START_ROW,
//     startCol: SHEET_CONFIG.INPUT.START_COL,
//     headerRow: SHEET_CONFIG.INPUT.HEADER_ROW,
//     setupFunction: (sheet, config) => {
//       sheet.getRange('A' + config.headerRow)
//            .setFormula(SHEET_CONFIG.INPUT.ARRAY_FORMULA);
//     }
//   });
// }

// /**
//  * Setup the Requests sheet
//  */
// function setupSheetRequests() {
//   try {
//     resetSheet(SHEET_CONFIG.REQUESTS.NAME);

//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName(SHEET_CONFIG.REQUESTS.NAME);

//     // Set headers and formatting
//     const headerRange = sheet.getRange(1, 1, 1, SHEET_CONFIG.REQUESTS.HEADERS.length);
//     headerRange.setValues([SHEET_CONFIG.REQUESTS.HEADERS])
//               .setBackground('#C9DAF8')
//               .setFontWeight('bold')
//               .setHorizontalAlignment('center');

//     // Set column widths
//     SHEET_CONFIG.REQUESTS.COLUMN_WIDTHS.forEach((width, index) => {
//       sheet.setColumnWidth(index + 1, width);
//     });

//     // Freeze header row and add filter
//     sheet.setFrozenRows(1);
//     toggleSheetFilter(SHEET_CONFIG.REQUESTS.NAME, 1, true);

//   } catch (error) {
//     Logger.log(error.stack);
//   }
// }