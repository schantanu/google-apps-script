// Constants
//
// -------------------------------------------------- Data --------------------------------------------------
//
// Misc
const dataSheetName = 'Data';
const dataHeaderRow = 1;
const dataStartRow = 1;
const dataStartCol = 1;

// Array columns (Array columns start at 0)
const dataArrColCommonId = 6;
const dataArrColMonth = 12;
const dataArrColValue = 13;

// Sheet columns
const dataColValidityDate       = 1;
const dataColForecastType       = 2;
const dataColFamily             = 3;
const dataColGroupFunction      = 4;
const dataColVendor             = 5;
const dataColShortName          = 6;
const dataColCommonID           = 7;
const dataColCommonDesc         = 8;
const dataColMaterialGroup      = 9;
const dataColPlannedStartDate   = 10;
const dataColPlannedEndDate     = 11;
const dataColIsCurrentlyActive  = 12;
const dataColMonth              = 13;
const dataColValue              = 14;
const dataColAction             = 15;
const dataColNotes              = 16;
const dataColUserEmail          = 17;
const dataColLastUpdated        = 18;
//
// -------------------------------------------------- Input -------------------------------------------------
//
// Misc
const inputSheetName = 'Input';
const inputHeaderRow = 2;
const inputColPivotDateStart = 11;

// Sheet columns
const inputColFamily             = 1;
const inputColGroupFunction      = 2;
const inputColVendor             = 3;
const inputColShortName          = 4;
const inputColCommonID           = 5;
const inputColCommonDesc         = 6;
const inputColMaterialGroup      = 7;
const inputColPlannedStartDate   = 8;
const inputColPlannedEndDate     = 9;
const inputColIsCurrentlyActive  = 10;
//
// -------------------------------------------------- Admin -------------------------------------------------
//
// Misc
const adminSheetName = 'Admin';
const adminDataRowStart = 3;
const adminUserColsTotal = 2;
const adminColAdmins = 1;
const adminColUsers = 2;
const adminAddConstraintsMonth = 3;
const adminRemoveConstraintsMonth = 5;
const addConstraintsRange = 'C3:D';
const removeConstraintsRange = 'E3:F';
//
// ------------------------------------------------ Changelog -----------------------------------------------
//
// Misc
const changelogSheetName = 'Changelog';
const changelogDataRowStart = 2;
const changelogTotalColumns = 7;

// Sheet columns
const changelogColCommonId    = 1;
const changelogColMonth       = 2;
const changelogColValue       = 3;
const changelogColAction      = 4;
const changelogColNotes       = 5;
const changelogColUserEmail   = 6;
const changelogColLastUpdated = 7;

// Array columns (Array columns start at 0)
const changelogArrColCommonId = 0;
const changelogArrColMonth = 1;
const changelogArrColValue = 2;
const changelogArrColAction = 3;
const changelogArrColNotes = 4;
const changelogArrColUserEmail = 5;
const changelogArrColLastUpdated = 6;

// ------------------------------------------- Sheets Protection --------------------------------------------
//
// Sheet Protection Configuration for User edits
// 'Data'       - N, O, P, Q
// 'Changelog'  - A, B, C, D, E, F
const sheetsEditConfig = [
  { sheetName: 'Data', editRowStart: 2, editColStart: 14, editTotalColumns: 4 },
  { sheetName: 'Changelog', editRowStart: 2, editColStart: 1, editTotalColumns: 6 }
];
//
// ----------------------------------------------------------------------------------------------------------