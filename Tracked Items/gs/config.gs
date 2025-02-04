/**
 * Configuration for attributes
 */
const ATTRIBUTE_CONFIG = {
  COMMON_LEVEL_FIELDS: [
    'SHORT NAME','COMMON DESCRIPTION','VENDOR','MATERIAL GROUP','GROUP FUNCTION','FAMILY',
    'SUBFAMILY','PLANNED START DATE','PLANNED END DATE','BUDGETLINEITEM','BUDGETLINEITEM2',
    'BUDGET START DATE','BUDGET END DATE','PARENT COMMON ID','TRACKED SET','FREQ','POWER',
    'INTEGRATED','TECH'
  ],
  ITEM_LEVEL_FIELDS: [
    'ITEM DESCRIPTION','DP HIERARCHY'
  ],
  DATES: [
    'PLANNED START DATE','PLANNED END DATE','BUDGET START DATE','BUDGET END DATE'
  ],
  ARCHIVED_FIELDS: [
    'BUDGETLINEITEM2','DP HIERARCHY','TRACKED SET','FREQ','POWER','INTEGRATED','TECH'
  ],
  CHAR_LIMITS: {
    SHORT_NAME        : 100,
    COMMON_DESC       : 200,
    ITEM_DESC         : 200,
    BUDGETLINEITEM    : 200,
    BUDGETLINEITEM2   : 200,
    PARENT_COMMON_ID  : 200,
    NOTE              : 200,
  },
}

/**
 * Sheet Configuration Constants
 * Contains all sheet-specific configurations and constants
 */
const SHEET_CONFIG = {
  ADMIN: {
    NAME: 'Admin',
    POSITIONS: {
      SUBHEADER_ROW: 2,
      DATA_START_ROW: 3,
      USERS_TOTAL_COLS: 2
    },
    COLUMNS: {
      INDEX: {
        ADMINS: 1,
        USERS: 2
      },
      HEADERS: {
        HEADER: 'Users List',
        SUBHEADER: [['Admins','Users']]
      },
      WIDTH: 200
    },
    RANGES: {
      HEADER: 'A1:B1',
      SUBHEADER: 'A2:B2'
    },
    COLORS: {
      HEADER: '#FFFF00',
      SUBHEADER: '#C9DAF8'
    }
  },
  DROPDOWNS: {
    NAME: 'Dropdowns',
    POSITIONS: {
      HEADER_ROW: 1,
      START_ROW: 1,
      START_COL: 1,
      DATA_START_ROW: 2
    },
    COLUMNS: {
      INDEX: {
        'VENDOR'          : 1,
        'GROUP FUNCTION'  : 2,
        'FAMILY'          : 3,
        'SUBFAMILY'       : 4,
        'MATERIAL GROUP'  : 5,
        'DP HIERARCHY'    : 6,
        'TRACKED SET'     : 7,
        'FREQ'            : 8,
        'POWER'           : 9,
        'INTEGRATED'      : 10,
        'TECH'            : 11
      },
      HEADERS: [
        'VENDOR','GROUP FUNCTION','FAMILY','SUBFAMILY','MATERIAL GROUP',
        'DP HIERARCHY','TRACKED SET','FREQ','POWER','INTEGRATED','TECH'
      ],
      WIDTHS: [150, 150, 150, 150, 130, 125, 125, 100, 100, 100, 100]
    },
    COLORS: {
      HEADER: '#C9DAF8'
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
    COLUMNS: {
      HEADERS: [
        'TIMESTAMP', 'USER', 'STATUS', 'ID', 'ATTRIBUTE LEVEL',
        'ATTRIBUTE', 'CURRENT VALUE', 'REQUESTED VALUE', 'NOTES'
      ],
      WIDTHS: [165, 200, 70, 150, 130, 200, 200, 200, 300]
    },
    COLORS: {
      HEADER: '#C9DAF8'
    }
  },
  INPUT: {
    NAME: 'Input',
    POSITIONS: {
      HEADER_ROW: 1,
      START_ROW: 1,
      START_COL: 1,
      DATA_START_ROW: 2,
    },
    IMPORT: {
      CELL: 'A1',
      FORMULA: '=ARRAYFORMULA(Data!A1:AA)'
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
        'ITEM ID'             : 1,
        'COMMON ID'           : 2,
        'FORECASTED S4 MMID'  : 3,
        'SHORT NAME'          : 4,
        'COMMON DESCRIPTION'  : 5,
        'ITEM DESCRIPTION'    : 6,
        'VENDOR'              : 7,
        'MATERIAL GROUP'      : 8,
        'GROUP FUNCTION'      : 9,
        'FAMILY'              : 10,
        'SUBFAMILY'           : 11,
        'PLANNED START DATE'  : 12,
        'PLANNED END DATE'    : 13,
        'BUDGET LINE ITEM'    : 14,
        'BUDGET LINE ITEM2'   : 15,
        'BUDGET START DATE'   : 16,
        'BUDGET END DATE'     : 17,
        'PARENT COMMON ID'    : 18,
        'DP HIERARCHY'        : 19,
        'TRACKED SET'         : 20,
        'FREQ'                : 21,
        'POWER'               : 22,
        'INTEGRATED'          : 23,
        'TECH'                : 24,
        'PLANNER NAME'        : 25,
        'LAST UPDATED BY'     : 26,
        'LAST UPDATED'        : 27
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
      WIDTHS: [
        150, 150, 150, 200, 200, 200, 150, 90, 100, 100, 100, 82, 82, 200,
        150, 82, 82, 150, 100, 100, 100, 100, 110, 100, 150, 200, 150
      ],
    },
    COLORS: {
      HEADER: '#C9DAF8'
    }
  }
};

/**
 * Configuration for sheet protection and user edit permissions
 */
const PROTECTION_CONFIG = {
  SHEETS: [
    { sheetName: 'Data', editRowStart: 2, editColStart: 3, editTotalColumns: 4 },
    { sheetName: 'Changelog', editRowStart: 2, editColStart: 1, editTotalColumns: 6 }
  ]
};