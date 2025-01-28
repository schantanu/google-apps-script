/**
 * Sheet Configuration Constants
 * Contains all sheet-specific configurations and constants
 */

const inputSheetName = 'Input';
const inputHeaderRow = 1;
const inputStartRow = 1;
const inputStartCol = 1;
const inputArrayFormula = '=ARRAYFORMULA(Data!A1:AA)';

const requestsSheetName = 'Requests';

const dataSheetName = 'Data';
const dataHeaderRow = 1;
const dataStartRow = 1;
const dataStartCol = 1;
const dataColHeaders = [
  'ITEM ID','COMMON ID','FORECASTED S4 MMID','SHORT NAME','COMMON DESCRIPTION',
  'ITEM DESCRIPTION','VENDOR','MATERIAL GROUP','GROUP FUNCTION','FAMILY',
  'SUBFAMILY','PLANNED START DATE','PLANNED END DATE','BUDGETLINEITEM',
  'BUDGETLINEITEM 2','BUDGET START DATE','BUDGET END DATE','PARENT COMMON ID',
  'DP HIERARCHY','TRACKED SET','FREQ','POWER','INTEGRATED','TECH','PLANNER',
  'LAST UPDATED BY','LAST UPDATED'
]

const dataColItemId             = 1;
const dataColCommonId           = 2;
const dataColForecastedS4MMId   = 3;
const dataColShortName          = 4;
const dataColCommonDesc         = 5;
const dataColItemDesc           = 6;
const dataColVendor             = 7;
const dataColMaterialGroup      = 8;
const dataColGroupFunction      = 9;
const dataColFamily             = 10;
const dataColSubFamily          = 11;
const dataColPlannedStartDate   = 12;
const dataColPlannedEndDate     = 13;
const dataColBudgetLineItem     = 14;
const dataColBudgetLineItem2    = 15;
const dataColBudgetStartDate    = 16;
const dataColBudgetEndDate      = 17;
const dataColParentCommonId     = 18;
const dataColDPHierarchy        = 19;
const dataColTrackedSet         = 20;
const dataColFreq               = 21;
const dataColPower              = 22;
const dataColIntegrated         = 23;
const dataColTech               = 24;
const dataColPlanner            = 25;
const dataColLastUpdatedBy      = 26;
const dataColLastUpdated        = 27;

const changelogSheetName = 'Changelog';


const adminSheetName = 'Admin';
const adminDataRowStart = 3;
const adminUserColsTotal = 2;
const adminColAdmins = 1;
const adminColUsers = 2;
// const adminColMaintainers = 3;

const adminColDropdownHeaders = ['MATERIAL GROUP','GROUP FUNCTION','FAMILY','SUBFAMILY','VENDOR','DP HIERARCHY','TRACKED SET','FREQ','POWER','INTEGRATED','TECH'];



// ------------------------------------------- Sheets Protection --------------------------------------------
//
// Sheet Protection Configuration for User edits
// 'Data'       - N, O, P, Q
// 'Changelog'  - A, B, C, D, E, F
const sheetsEditConfig = [
  { sheetName: 'Data', editRowStart: 2, editColStart: 3, editTotalColumns: 4 },
  { sheetName: 'Changelog', editRowStart: 2, editColStart: 1, editTotalColumns: 6 }
];
//
// ----------------------------------------------------------------------------------------------------------


// Attribute configuration for Item and Common ID level attributes
const attributeConfig = {
  itemAttributes: [
    { name: 'Item Attr 1', type: 'text', maxLength: 100 },
    { name: 'Item Attr 2', type: 'text', maxLength: 100 }
  ],
  commonAttributes: [
    { name: 'Common Attr 1', type: 'dropdown', sourceSheet: 'Admin', sourceRange: 'C2:C' },
    { name: 'Common Attr 2', type: 'date' },
    { name: 'Common Attr 3', type: 'date', minDate: 'Common Attr 2' }
  ]
};

// Sheet Names and Basic Configurations
// const SHEET_CONFIG = {
//   ADMIN: {
//     NAME: 'Admin',
//     DATA_START_ROW: 3,
//     USER_COLS_TOTAL: 2,
//     COLUMNS: {
//       ADMINS: 1,
//       USERS: 2
//     },
//     HEADERS: {
//       USERS_LIST: 'Users List',
//       DROPDOWN_ATTRS: 'Dropdown Attributes'
//     },
//     COLORS: {
//       USERS_HEADER: '#FFFF00',
//       DROPDOWN_HEADER: '#B6D7A8',
//       SUBHEADER: '#C9DAF8'
//     }
//   },
//   DATA: {
//     NAME: 'Data',
//     HEADER_ROW: 1,
//     START_ROW: 1,
//     START_COL: 1,
//     COLUMNS: {
//       ITEM_ID: 1,
//       COMMON_ID: 2,
//       FORECASTED_S4MM_ID: 3,
//       SHORT_NAME: 4,
//       COMMON_DESC: 5,
//       ITEM_DESC: 6,
//       VENDOR: 7,
//       MATERIAL_GROUP: 8,
//       GROUP_FUNCTION: 9,
//       FAMILY: 10,
//       SUBFAMILY: 11,
//       PLANNED_START_DATE: 12,
//       PLANNED_END_DATE: 13,
//       BUDGET_LINE_ITEM: 14,
//       BUDGET_LINE_ITEM2: 15,
//       BUDGET_START_DATE: 16,
//       BUDGET_END_DATE: 17,
//       PARENT_COMMON_ID: 18,
//       DP_HIERARCHY: 19,
//       TRACKED_SET: 20,
//       FREQ: 21,
//       POWER: 22,
//       INTEGRATED: 23,
//       TECH: 24,
//       PLANNER: 25,
//       LAST_UPDATED_BY: 26,
//       LAST_UPDATED: 27
//     },
//     HEADERS: [
//       'ITEM ID', 'COMMON ID', 'FORECASTED S4 MMID', 'SHORT NAME',
//       'COMMON DESCRIPTION', 'ITEM DESCRIPTION', 'VENDOR', 'MATERIAL GROUP',
//       'GROUP FUNCTION', 'FAMILY', 'SUBFAMILY', 'PLANNED START DATE',
//       'PLANNED END DATE', 'BUDGETLINEITEM', 'BUDGETLINEITEM 2',
//       'BUDGET START DATE', 'BUDGET END DATE', 'PARENT COMMON ID',
//       'DP HIERARCHY', 'TRACKED SET', 'FREQ', 'POWER', 'INTEGRATED',
//       'TECH', 'PLANNER', 'LAST UPDATED BY', 'LAST UPDATED'
//     ]
//   },
//   INPUT: {
//     NAME: 'Input',
//     HEADER_ROW: 1,
//     START_ROW: 1,
//     START_COL: 1,
//     ARRAY_FORMULA: '=ARRAYFORMULA(Data!A1:AA)'
//   },
//   REQUESTS: {
//     NAME: 'Requests',
//     HEADERS: [
//       'Timestamp', 'User', 'Status', 'ID', 'Attribute Level',
//       'Attribute', 'Current Value', 'Requested Value'
//     ],
//     COLUMN_WIDTHS: [160, 200, 70, 150, 120, 200, 200, 200]
//   },
//   CHANGELOG: {
//     NAME: 'Changelog'
//   }
// };

// /**
//  * Configuration for sheet protection and user edit permissions
//  */
// const PROTECTION_CONFIG = {
//   SHEETS: [
//     {
//       sheetName: 'Data',
//       editRowStart: 2,
//       editColStart: 3,
//       editTotalColumns: 4
//     },
//     {
//       sheetName: 'Changelog',
//       editRowStart: 2,
//       editColStart: 1,
//       editTotalColumns: 6
//     }
//   ]
// };

// /**
//  * Configuration for data attributes at Item and Common ID levels
//  */
// const ATTRIBUTE_CONFIG = {
//   ITEM_LEVEL: [
//     {
//       name: 'Item Attr 1',
//       type: 'text',
//       maxLength: 100
//     },
//     {
//       name: 'Item Attr 2',
//       type: 'text',
//       maxLength: 100
//     }
//   ],
//   COMMON_LEVEL: [
//     {
//       name: 'Common Attr 1',
//       type: 'dropdown',
//       sourceSheet: 'Admin',
//       sourceRange: 'C2:C'
//     },
//     {
//       name: 'Common Attr 2',
//       type: 'date'
//     },
//     {
//       name: 'Common Attr 3',
//       type: 'date',
//       minDate: 'Common Attr 2'
//     }
//   ]
// };

// /**
//  * Configuration for Admin sheet dropdown attributes
//  */
// const ADMIN_DROPDOWN_ATTRIBUTES = [
//   'MATERIAL GROUP',
//   'GROUP FUNCTION',
//   'FAMILY',
//   'SUBFAMILY',
//   'VENDOR',
//   'DP HIERARCHY',
//   'TRACKED SET',
//   'FREQ',
//   'POWER',
//   'INTEGRATED',
//   'TECH'
// ];