// Sheet Reset Functions
// ----------------------------------------------------------------------------------------------------------
//
// Run the following function to reset and setup the respective sheets from scratch
// setupSheetData()         - Setup the 'Data' sheet
// setupSheetAdmin()        - Setup the 'Admin' sheet
// setupSheetInput()        - Setup the 'Input' sheet
// setupSheetChangelog()    - Setup the 'Changelog' sheet
//
// ----------------------------------------------------------------------------------------------------------

/**
 * Setup the 'Data' sheet
 */
function setupSheetData() {
  try {
    // Reset sheet
    resetSheet(dataSheetName, 'all');

    // Get sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(dataSheetName);

    // Add empty rows
    const currentRows = sheet.getMaxRows();
    const rowsToAdd = 15000 - currentRows;
    if (rowsToAdd > 0) {
      sheet.insertRowsAfter(currentRows, rowsToAdd);
    }

    // Freeze top row
    sheet.setFrozenRows(dataHeaderRow);

    // Format Header row
    headerRange = sheet.getRange(dataStartRow, dataStartCol, dataHeaderRow, dataColLastUpdated);
    headerRange.setFontWeight('Bold');
    headerRange.setBackground('#C9DAF8');
    headerRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    // Show User editable columns
    userEditableRange = sheet.getRange(dataStartRow, dataColValue, dataHeaderRow, (dataColLastUpdated - dataColValue));
    userEditableRange.setBackground('#FFE599');

    // Set Header row height
    sheet.setRowHeight(dataHeaderRow, 50);

    // Set column widths
    sheet.setColumnWidth(dataColValidityDate      , 85);
    sheet.setColumnWidth(dataColForecastType      , 90);
    sheet.setColumnWidth(dataColFamily            , 120);
    sheet.setColumnWidth(dataColGroupFunction     , 100);
    sheet.setColumnWidth(dataColVendor            , 100);
    sheet.setColumnWidth(dataColShortName         , 100);
    sheet.setColumnWidth(dataColCommonID          , 130);
    sheet.setColumnWidth(dataColCommonDesc        , 100);
    sheet.setColumnWidth(dataColMaterialGroup     , 90);
    sheet.setColumnWidth(dataColPlannedStartDate  , 85);
    sheet.setColumnWidth(dataColPlannedEndDate    , 82);
    sheet.setColumnWidth(dataColIsCurrentlyActive , 95);
    sheet.setColumnWidth(dataColMonth             , 85);
    sheet.setColumnWidth(dataColValue             , 105);
    sheet.setColumnWidth(dataColAction            , 145);
    sheet.setColumnWidth(dataColNotes             , 100);
    sheet.setColumnWidth(dataColUserEmail         , 200);
    sheet.setColumnWidth(dataColLastUpdated       , 120);

    // Format date columns to 'dd-mmm-yy' format
    const maxRows = sheet.getMaxRows();
    const dateColsToFormat = [dataColValidityDate, dataColPlannedStartDate, dataColPlannedEndDate, dataColMonth];
    dateColsToFormat.forEach(function(col) {
      sheet.getRange(dataStartRow, col, maxRows).setNumberFormat('dd-mmm-yy');
    });

    // Format 'Last Updated' column to 'Plain Text' format
    sheet.getRange(1, dataColLastUpdated, maxRows).setNumberFormat('@');

  } catch (error) {
    Logger.log(error.stack);
  }
}

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

    // Add and Format Header Rows
    sheet.getRange(1, 1, 1, 2).mergeAcross();
    sheet.getRange(1, 3, 1, 2).mergeAcross();
    sheet.getRange(1, 5, 1, 2).mergeAcross();

    sheet.getRange('A1').setValue('Users List');
    sheet.getRange('C1').setValue('Add Constraint Items List');
    sheet.getRange('E1').setValue('Remove Constraint Items List');
    sheet.getRange('A2:F2').setValues([['Admins','Users','Months','Common ID','Months','Common ID']]);

    sheet.getRange(1, 1, 2, 6).setHorizontalAlignment('center');
    sheet.getRange(1, 1, 2, 6).setFontWeight('Bold');
    sheet.getRange('A1').setBackground('#FFFF00');
    sheet.getRange('C1').setBackground('#B6D7A8');
    sheet.getRange('E1').setBackground('#F4CCCC');
    sheet.getRange(2, 1, 1, 6).setBackground('#C9DAF8');

    // Set Formulas
    sheet.getRange('C3').setFormula('=SORT(UNIQUE(FILTER(Data!$M$2:$M, TO_DATE(Data!$M$2:$M) >= EOMONTH(TODAY(),-2)+1)),1,TRUE)');
    sheet.getRange('D3').setFormula('=SORT(UNIQUE(Data!$G$2:$G))');
    sheet.getRange('E3').setFormula('=SORT(UNIQUE(FILTER(Data!$M$2:$M,Data!$N$2:$N = "Manual Constraint", TO_DATE(Data!$M$2:$M) >= EOMONTH(TODAY(),-2)+1)),1,TRUE)');
    sheet.getRange('F3').setFormula('=SORT(UNIQUE(FILTER(Data!$G$2:$G,Data!$N$2:$N = "Manual Constraint", TO_DATE(Data!$M$2:$M) >= EOMONTH(TODAY(),-2)+1)),1,FALSE)');

    // Auto Resize columns
    sheet.setColumnWidths(1, 6, 200);

    // Format date columns to 'dd-mmm-yy' format
    const maxRows = sheet.getMaxRows();
    const dateColsToFormat = [adminAddConstraintsMonth, adminRemoveConstraintsMonth];
    dateColsToFormat.forEach(function(col) {
      sheet.getRange(1, col, maxRows).setNumberFormat('dd-mmm-yy');
    });

    // Hide Sheet
    hideSheet(adminSheetName);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Setup the 'Input' sheet
 */
function setupSheetInput() {
  try {
    // Reset sheet
    resetSheet(inputSheetName, 'all');

    // Get sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(inputSheetName);

    // Set Pivot formula
    const pivotCell = sheet.getRange('A' + inputHeaderRow);
    const pivotFormula = `=INDEX(REGEXREPLACE(""&QUERY({Data!A1:O, UPPER(TEXT(Data!M1:M*1, "yyyymmdd♦mmm-yyyy")), IF(ISBLANK(Data!J1:K),"",UPPER(TEXT(Data!J1:K, "dd-mmm-yy")))},
    "SELECT Col3, Col4, Col5, Col6, Col7, Col8, Col9, Col17, Col18, Col12, MAX(Col14)
      WHERE Col7<>''
      GROUP BY Col3, Col4, Col5, Col6, Col7, Col8, Col9, Col17, Col18, Col12
      PIVOT Col16", 1),
    "^(.*♦)", ))`;
    pivotCell.setFormula(pivotFormula);

    // Set Refresh Date formula
    sheet.getRange('A1').setValue('Last Refresh:');
    const refreshDateCell = sheet.getRange('B1');
    const refreshDateFormula = '=IF(COUNT(UNIQUE(FILTER(Data!$A$2:$A,Data!$A$2:$A<>"")))>1,"ERROR!!!",UNIQUE(FILTER(Data!$A$2:$A,Data!$A$2:$A<>"")))';
    refreshDateCell.setFormula(refreshDateFormula);

    // Freeze rows and columns
    sheet.setFrozenRows(inputHeaderRow);
    sheet.setFrozenColumns(inputColPivotDateStart - 1);

    // Get Pivot column end
    const inputColPivotDateEnd = sheet.getLastColumn();

    // Format Refresh Date Row
    sheet.getRange(1, 1, 1, 2).setFontSize(9).setFontWeight('Bold');
    sheet.getRange(1, 1, 1, 2).setBackground('#FFF2CC');

    // Format Header Row
    sheet.getRange(inputHeaderRow, 1, 1, inputColPivotDateEnd).setFontSize(9).setFontWeight('Bold');
    sheet.getRange(1, 1, 1, 2).setBackground('#FFF2CC');

    // Format Header Row
    sheet.getRange(inputHeaderRow, 1, 1, inputColPivotDateEnd).setFontSize(9).setFontWeight('Bold');
    sheet.getRange(inputHeaderRow, 1, 1, inputColPivotDateEnd).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheet.setRowHeight(inputHeaderRow, 40);
    sheet.getRange(inputHeaderRow, 1, 1, inputColPivotDateStart-1).setBackground('#D9EAD3');
    sheet.getRange(inputHeaderRow, inputColPivotDateStart, 1, ((inputColPivotDateEnd - inputColPivotDateStart) + 1)).setBackground('#C9DAF8');

    // Set Wrap for Sheet
    sheet.getRange(inputHeaderRow+1, 1, sheet.getLastRow(), inputColIsCurrentlyActive).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // Set filter
    sheet.getRange(inputHeaderRow, 1, sheet.getLastRow()-(inputHeaderRow-1), sheet.getLastColumn()).createFilter();

    // Set column widths
    sheet.setColumnWidth(inputColFamily             , 86);
    sheet.setColumnWidth(inputColGroupFunction      , 89);
    sheet.setColumnWidth(inputColVendor             , 85);
    sheet.setColumnWidth(inputColShortName          , 180);
    sheet.setColumnWidth(inputColCommonID           , 140);
    sheet.setColumnWidth(inputColCommonDesc         , 100);
    sheet.setColumnWidth(inputColMaterialGroup      , 90);
    sheet.setColumnWidth(inputColPlannedStartDate   , 95);
    sheet.setColumnWidth(inputColPlannedEndDate     , 95);
    sheet.setColumnWidth(inputColIsCurrentlyActive  , 95);
    sheet.setColumnWidths(inputColPivotDateStart, inputColPivotDateEnd, 105);

    // Set Horizontal Alignment
    sheet.getRange(inputHeaderRow+1, inputColPlannedStartDate, sheet.getMaxRows(), inputColPivotDateStart - inputColPlannedStartDate).setHorizontalAlignment('center');

    // Conditional Formatting
    // Set cells with 'Manual Constraint' to Red
    const pivotRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sheet.getRange(inputHeaderRow+1, inputColPivotDateStart, sheet.getMaxRows(), ((inputColPivotDateEnd - inputColPivotDateStart) + 1))])
      .whenTextEqualTo('Manual Constraint')
      .setBackground('#B7E1CD')
      .build();
    const conditionalFormatRules = sheet.getConditionalFormatRules();
    conditionalFormatRules.push(pivotRule);

    // Set error rule for Refresh Date
    const rangeRefreshDate = sheet.getRange('B1');
    const refreshDateRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([rangeRefreshDate])
      .whenTextEqualTo('ERROR!!!')
      .setBackground('#FF0000')
      .build();
    conditionalFormatRules.push(refreshDateRule);

    // Set Rules
    sheet.setConditionalFormatRules(conditionalFormatRules);

  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Setup the 'Changelog' sheet
 */
function setupSheetChangelog() {
  try {
    // Reset sheet
    resetSheet(changelogSheetName);

    // Get sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(changelogSheetName);

    // Set and format Header row
    headerRange = sheet.getRange(1, 1, 1, changelogColLastUpdated);
    headerRange.setValues([['COMMON ID','MONTH','VALUE','ACTION','NOTES','USER EMAIL','LAST UPDATED']]);
    headerRange.setFontWeight('Bold');
    headerRange.setBackground('#C9DAF8');

    // Freeze top row
    sheet.setFrozenRows(1);

    // Set column widths
    sheet.setColumnWidth(changelogColCommonId     , 170);
    sheet.setColumnWidth(changelogColMonth        , 70);
    sheet.setColumnWidth(changelogColValue        , 105);
    sheet.setColumnWidth(changelogColAction       , 145);
    sheet.setColumnWidth(changelogColNotes        , 300);
    sheet.setColumnWidth(changelogColUserEmail    , 200);
    sheet.setColumnWidth(changelogColLastUpdated  , 120);

    // Format date columns to 'Plain Text' format
    const maxRows = sheet.getMaxRows();
    const dateColsToFormat = [changelogColMonth, changelogColLastUpdated];
    dateColsToFormat.forEach(function(col) {
      sheet.getRange(1, col, maxRows).setNumberFormat('@');
    });

    // Set Insert Query formula in cell K1
    // const cellFormula = 'K1';
    // const uploadFormula = `IF(A1<>"","INSERT INTO network_rw.forecast_accuracy_items (common_id, month, value, action, notes, user_email, last_updated) VALUES ('"&A1&"',TO_DATE('01-"&B1&"','dd-Mon-yy'),'"&C1&"','"&D1&"','"&E1&"','"&F1&"',TO_DATE('"&TEXT(G1,"dd-mmm-yy HH:mm:ss")&"','dd-Mon-yy hh24:mi:ss'));","")
    // `;
    // sheet.getRange(cellFormula).setFormula(uploadFormula);
    // sheet.getRange(cellFormula).setFontColor('#efefef');

  } catch (error) {
    Logger.log(error.stack);
  }
}