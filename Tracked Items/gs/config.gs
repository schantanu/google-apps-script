/**
 * Sheet Configuration Constants
 * Contains all sheet-specific configurations and constants
 */
// Sheet Names and Basic Configurations
const SHEET_CONFIG = {
  ARCHIVED_FIELDS: ['BUDGETLINEITEM2','DP HIERARCHY','TRACKED SET','FREQ','POWER','INTEGRATED','TECH'],
  CHAR_LIMITS: {
    NOTE: 200,
    SHORT_NAME: 100,
    COMMON_DESC: 250,
    ITEM_DESC: 250,
    BUDGETLINEITEM: 150,
    BUDGETLINEITEM2: 150,
    PARENT_COMMON_ID: 50
  },
  ADMIN: {
    NAME: 'Admin',
    POSITIONS: {
      SUBHEADER_ROW: 2,
      DATA_ROW_START: 3,
      USERS_TOTAL_COLS: 2
    },
    COLUMNS: {
      ADMINS: 1,
      USERS: 2
    },
    COLUMN_WIDTH: 200,
    RANGES: {
      USERS_HEADER: 'A1:B1',
      USERS_SUBHEADER: 'A2:B2'
    },
    HEADERS: {
      USERS_HEADER: 'Users List',
      USERS_SUBHEADER: [['Admins','Users']]
    },
    COLORS: {
      USERS_HEADER: '#FFFF00',
      USERS_SUBHEADER: '#C9DAF8'
    }
  },
  REQUESTS: {
    NAME: 'Requests',
    POSITIONS: {
      START_ROW: 1,
      START_COL: 1,
      HEADER_ROW: 1,
      DATA_START_ROW: 2
    },
    HEADERS: [
      'TIMESTAMP', 'USER', 'STATUS', 'ID', 'ATTRIBUTE LEVEL',
      'ATTRIBUTE', 'CURRENT VALUE', 'REQUESTED VALUE', 'NOTES'
    ],
    COLUMN_WIDTHS: [165, 200, 70, 150, 130, 200, 200, 200, 300],
    COLORS: {
      HEADER: '#C9DAF8'
    }
  },
  DROPDOWNS: {
    NAME: 'Dropdowns',
    POSITIONS: {
      HEADER_ROW: 1,
      START_ROW: 1,
      START_COL: 1
    },
    COLUMNS: {
      INDEX: {
        VENDOR          : 1,
        GROUP_FUNCTION  : 2,
        FAMILY          : 3,
        SUBFAMILY       : 4,
        MATERIAL_GROUP  : 5,
        DP_HIERARCHY    : 6,
        TRACKED_SET     : 7,
        FREQ            : 8,
        POWER           : 9,
        INTEGRATED      : 10,
        TECH            : 11
      }
    },
    HEADERS: [
      'VENDOR','GROUP FUNCTION','FAMILY','SUBFAMILY','MATERIAL GROUP',
      'DP HIERARCHY','TRACKED SET','FREQ','POWER','INTEGRATED','TECH'
    ],
    COLUMN_WIDTHS: [150, 150, 150, 150, 130, 125, 125, 100, 100, 100, 100],
    COLORS: {
      HEADER: '#C9DAF8'
    }
  },
  DATA: {
    NAME: 'Data',
    POSITIONS: {
      HEADER_ROW: 1,
      START_ROW: 1,
      START_COL: 1,
      DATA_START_ROW: 2
    },
    COLUMNS: {
      INDEX: {
        ITEM_ID             : 1,
        COMMON_ID           : 2,
        FORECASTED_S4MM_ID  : 3,
        SHORT_NAME          : 4,
        COMMON_DESC         : 5,
        ITEM_DESC           : 6,
        VENDOR              : 7,
        MATERIAL_GROUP      : 8,
        GROUP_FUNCTION      : 9,
        FAMILY              : 10,
        SUBFAMILY           : 11,
        PLANNED_START_DATE  : 12,
        PLANNED_END_DATE    : 13,
        BUDGET_LINE_ITEM    : 14,
        BUDGET_LINE_ITEM2   : 15,
        BUDGET_START_DATE   : 16,
        BUDGET_END_DATE     : 17,
        PARENT_COMMON_ID    : 18,
        DP_HIERARCHY        : 19,
        TRACKED_SET         : 20,
        FREQ                : 21,
        POWER               : 22,
        INTEGRATED          : 23,
        TECH                : 24,
        PLANNER_NAME        : 25,
        LAST_UPDATED_BY     : 26,
        LAST_UPDATED        : 27
      },
      WIDTHS : {
        [INDEX.ITEM_ID]             : 150,
        [INDEX.COMMON_ID]           : 150,
        [INDEX.FORECASTED_S4MM_ID]  : 150,
        [INDEX.SHORT_NAME]          : 200,
        [INDEX.COMMON_DESC]         : 200,
        [INDEX.ITEM_DESC]           : 200,
        [INDEX.VENDOR]              : 150,
        [INDEX.MATERIAL_GROUP]      : 90,
        [INDEX.GROUP_FUNCTION]      : 100,
        [INDEX.FAMILY]              : 100,
        [INDEX.SUBFAMILY]           : 100,
        [INDEX.PLANNED_START_DATE]  : 82,
        [INDEX.PLANNED_END_DATE]    : 82,
        [INDEX.BUDGET_LINE_ITEM]    : 200,
        [INDEX.BUDGET_LINE_ITEM2]   : 150,
        [INDEX.BUDGET_START_DATE]   : 82,
        [INDEX.BUDGET_END_DATE]     : 82,
        [INDEX.PARENT_COMMON_ID]    : 150,
        [INDEX.DP_HIERARCHY]        : 100,
        [INDEX.TRACKED_SET]         : 100,
        [INDEX.FREQ]                : 100,
        [INDEX.POWER]               : 100,
        [INDEX.INTEGRATED]          : 110,
        [INDEX.TECH]                : 100,
        [INDEX.PLANNER_NAME]        : 150,
        [INDEX.LAST_UPDATED_BY]     : 200,
        [INDEX.LAST_UPDATED]        : 150
      }
    },
    HEADERS: [
      'ITEM ID', 'COMMON ID', 'FORECASTED S4 MMID', 'SHORT NAME',
      'COMMON DESCRIPTION', 'ITEM DESCRIPTION', 'VENDOR', 'MATERIAL GROUP',
      'GROUP FUNCTION', 'FAMILY', 'SUBFAMILY', 'PLANNED START DATE',
      'PLANNED END DATE', 'BUDGETLINEITEM', 'BUDGETLINEITEM2',
      'BUDGET START DATE', 'BUDGET END DATE', 'PARENT COMMON ID',
      'DP HIERARCHY', 'TRACKED SET', 'FREQ', 'POWER', 'INTEGRATED',
      'TECH', 'PLANNER NAME', 'LAST UPDATED BY', 'LAST UPDATED'
    ],
    COLORS: {
      HEADER: '#C9DAF8'
    }
  },
  INPUT: {
    NAME: 'Input',
    POSITIONS: {
      HEADER_ROW: 1,
      START_ROW: 1,
      START_COL: 1
    },
    IMPORT: {
      CELL: 'A1',
      FORMULA: '=ARRAYFORMULA(Data!A1:AA)'
    }
  }
};

/**
 * Configuration for sheet protection and user edit permissions
 */
const PROTECTION_CONFIG = {
  SHEETS: [
    {
      sheetName: 'Data',
      editRowStart: 2,
      editColStart: 3,
      editTotalColumns: 4
    },
    {
      sheetName: 'Changelog',
      editRowStart: 2,
      editColStart: 1,
      editTotalColumns: 6
    }
  ]
};

/**
 * Configuration for data attributes at Item and Common ID levels
 */
const ATTRIBUTE_CONFIG = {
  ITEM_LEVEL: [
    { name: 'Item Attr 1', type: 'text', maxLength: 100 },
    { name: 'Item Attr 2', type: 'text', maxLength: 100 }
  ],
  COMMON_LEVEL: [
    { name: 'Common Attr 1', type: 'dropdown', sourceSheet: 'Admin', sourceRange: 'C2:C' },
    { name: 'Common Attr 2', type: 'date' },
    { name: 'Common Attr 3', type: 'date', minDate: 'Common Attr 2' }
  ]
};

/**
 * Configuration for Admin sheet dropdown attributes
 */
const ADMIN_DROPDOWN_ATTRIBUTES = [
  'MATERIAL GROUP',
  'GROUP FUNCTION',
  'FAMILY',
  'SUBFAMILY',
  'VENDOR',
  'DP HIERARCHY',
  'TRACKED SET',
  'FREQ',
  'POWER',
  'INTEGRATED',
  'TECH'
];
