/**
 * Adds a custom menu to the Google Sheets UI on open.
 */
function onOpen(e) {
  try {
    // Create custom menu
    const ui = SpreadsheetApp.getUi();

    ui.createMenu('✍️ Editor')
        .addItem('📖 User Guide', 'showUserGuide')
      .addSeparator()
        .addItem('✅ Add Constraint Items', 'showAddConstraints')
        .addItem('❎ Remove Constraint Items', 'showRemoveConstraints')
      .addSeparator()
        .addItem('⚙️ Admin Console', 'showAdminConsole')
      .addToUi();

    // Hide Admin Sheet
    hideSheet(adminSheetName);

    // Activate Input Sheet
    activateCell(inputSheetName);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Displays the User Guide as a modal dialog.
 */
function showUserGuide() {
  try {
    // Show User Guide modal
    const html = HtmlService.createHtmlOutputFromFile('user_guide').setWidth(500).setHeight(640);
    SpreadsheetApp.getUi().showModalDialog(html, '📖 User Guide');
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Shows the add constraints sidebar form.
 */
function showAddConstraints() {
  try {
    // Show add constraints sidebar form
    showUserSidebarForm('add');
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Shows the remove constraints sidebar form.
 */
function showRemoveConstraints() {
  try {
    // Show remove constraints sidebar form
    showUserSidebarForm('remove');
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Shows the admin console sidebar.
 */
function showAdminConsole() {
  try {
    // Run only if current user has Admin access
    if (!checkAccess('Admin')) return;

    // Show Sidebar
    showSidebar('sidebar_admin', '⚙️ Admin Console');
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Shows Update data sidebar form.
 */
function showUpdateDataSidebar() {
  try {
    // Run only if current user has Admin access
    if (!checkAccess('Admin')) return;

    // Activate 'Changelog' sheet
    activateCell(changelogSheetName);

    // Get unique Comment Dates
    commentDates = getCommentDates();

    // Show Sidebar form
    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createTemplateFromFile('sidebar_data_update');
    htmlOutput.data = JSON.stringify(commentDates);
    ui.showSidebar(htmlOutput.evaluate().setTitle('Update Data'));
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Shows Get data sidebar.
 */
function showGetDataSidebar() {
  try {
    // Run only if current user has Admin access
    if (!checkAccess('Admin')) return;

    // Activate 'Data' sheet
    activateCell(dataSheetName);

    // Show Sidebar
    showSidebar('sidebar_data_get', 'Get Data');
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Shows the add role sidebar form.
 */
function showAddRoleSidebar() {
  try {
    // Run only if current user has Admin access
    if (!checkAccess('Admin')) return;

    // Activate 'Admin' sheet
    activateCell(adminSheetName);

    // Show Sidebar
    showSidebar('sidebar_role_add', 'Add Role');
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Shows the remove role sidebar form.
 */
function showRemoveRoleSidebar() {
  try {
    // Run only if current user has Admin access
    if (!checkAccess('Admin')) return;

    // Activate 'Admin' sheet
    activateCell(adminSheetName);

    // Show Sidebar
    showSidebar('sidebar_role_remove', 'Remove Role');
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Shows the reset sheet sidebar form.
 */
function showResetSheetSidebar() {
  try {
    // Run only if current user has Admin access
    if (!checkAccess('Admin')) return;

    // Show Sidebar
    showSidebar('sidebar_reset_sheet', 'Reset Sheet');
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Activates the previous month's cell in the 'Input' sheet.
 * Assumes dates are formatted as 'yyyy-M-d' and located in the header row of 'Input' sheet.
 */
function activatePreviousMonthCell() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(inputSheetName);
    const values = sheet.getRange(inputHeaderRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Get previous month
    const now = new Date();
    const previousMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const formattedPreviousMonth = Utilities.formatDate(previousMonth, Session.getScriptTimeZone(), 'MMM-yyyy').toUpperCase();

    // Find the cell with the previous month date
    const columnIndex = values.indexOf(formattedPreviousMonth);

    if (columnIndex !== -1) {
      const arbitaryCell = sheet.getRange(inputHeaderRow, columnIndex + 8).getA1Notation();
      const cell = sheet.getRange(inputHeaderRow, columnIndex + 1).getA1Notation();

      // Activate an arbitary cell first then scroll to previous month cell
      activateCell(inputSheetName, arbitaryCell);
      SpreadsheetApp.flush();
      activateCell(inputSheetName, cell);
    } else {
      Logger.log('Previous month cell not found');
    }
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Shows the sidebar form for adding or removing constraints.
 * @param {string} action - The action to be performed ('add' or 'remove').
 */
function showUserSidebarForm(action) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(adminSheetName);
    let dataRange;

    // Run only if current user has User access
    if (!checkAccess('User')) return;

    // Scroll to previous month cell
    activatePreviousMonthCell();

    // Source dropdown values from the 'Admin' sheet
    if (action === 'add') {
      dataRange = sheet.getRange(addConstraintsRange + sheet.getLastRow()).getValues();
    } else {
      dataRange = sheet.getRange(removeConstraintsRange + sheet.getLastRow()).getValues();
    }

    // Create empty set
    const itemSet = new Set();
    const monthSet = new Set();

    dataRange.forEach(row => {
      const monthVal = row[0];
      const itemVal = row[1];

      // Remove NA values and return null if no values in the Dropdown
      if (itemVal === '#N/A' || monthVal === '#N/A') {
        return;
      }

      // Add unique values to Set
      itemSet.add(itemVal);
      if (monthVal) {
        const date = new Date(monthVal);
        const month = date.toLocaleString('default', { month: 'short' });
        const year = date.getFullYear().toString().slice(-2);
        monthSet.add(`${month}-${year}`);
      }
    });

    // Create object for dropdown selection
    const result = { items: Array.from(itemSet), months: Array.from(monthSet), action };

    // Show Sidebar
    const html = HtmlService.createTemplateFromFile('sidebar_user');
    html.data = JSON.stringify(result);
    SpreadsheetApp.getUi().showSidebar(html.evaluate().setTitle(action === 'add' ? '✅ Add Constraint Items' : '❎ Remove Constraint Items').setWidth(300));
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Attempts to acquire a lock and submits the form data.
 * @param {Array} notes - Demand Planner notes.
 * @param {Array} selectedMonths - The selected months for constraints.
 * @param {Array} selectedItems - The selected items for constraints.
 * @param {boolean} isRemove - Whether the action is to remove constraints.
 */
function tryLockAndSubmit(notes, selectedMonths, selectedItems, isRemove) {
  const lock = LockService.getScriptLock();
  const user = Session.getActiveUser().getEmail();
  const userEmail = user ? user : 'Unknown User';

  try {
    // Wait up to 15 seconds for other processes to release the lock.
    if (lock.tryLock(15000)) {
      try {
        if (isRemove) {
          updateConstraints('', 'Removed Constraint Item', notes, selectedMonths, selectedItems);
        } else {
          updateConstraints('Manual Constraint', 'Added Constraint Item', notes, selectedMonths, selectedItems);
        }
        return true;
      } finally {
        lock.releaseLock();
        console.log(`User: ${userEmail}\nMonths: ${selectedMonths}\nCommon IDs: ${selectedItems}\nNotes: ${notes}`);
      }
    } else {
      // Display failure to acquire lock
      throw new Error('The Sheet is currently being edited by another user. Kindly wait 15-30 seconds and resubmit your selections.');
    }
  } catch (error) {
    // Display error in the client-side `showError` function.
    Logger.log(error);
    throw error;
  }
}

/**
 * Updates constraints in the Data sheet based on the provided parameters.
 * @param {string} notes - Demand Planner notes.
 * @param {string} value - The value to set in the constraint column ('Manual Constraint' or '').
 * @param {string} action - The action performed to be set ('Added Constraint Item' or 'Removed Constraint Item').
 * @param {Array} selectedMonths - The selected months for constraint item.
 * @param {Array} selectedItems - The selected items for constraint.
 */
function updateConstraints(value, action, notes, selectedMonths, selectedItems) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName);
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn());
    const values = dataRange.getValues();

    const updates = [];
    const user = Session.getActiveUser().getEmail();
    const userEmail = user ? user : 'Unknown User';
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yy HH:mm:ss');

    for (let i = 0; i < values.length; i++) {
      const item = values[i][dataArrColCommonId];
      const date = new Date(values[i][dataArrColMonth]);
      const rawMonthYear = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM-yy');

      if (selectedItems.includes(item) && selectedMonths.includes(rawMonthYear)) {
        const currentValue = values[i][dataArrColValue];
        const monthYear = '01-'+rawMonthYear;
        if ((value === '' && currentValue === 'Manual Constraint') || (value === 'Manual Constraint' && currentValue === '')) {
          updates.push({ row: i + 2, value, action, notes, userEmail, timestamp, item, monthYear });
        }
      }
    }

    if (updates.length > 0) {
      applyBatchUpdates(sheet, updates);
      logChanges(updates);
    }
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Applies batch updates to the Data sheet.
 * @param {Sheet} sheet - The sheet where updates will be applied.
 * @param {Array} updates - The updates to be applied.
 */
function applyBatchUpdates(sheet, updates) {
  try {
    updates.forEach(u => {
      sheet.getRange(u.row, dataColValue).setValue(u.value);
      sheet.getRange(u.row, dataColAction).setValue(u.action);
      sheet.getRange(u.row, dataColNotes).setValue(u.notes);
      sheet.getRange(u.row, dataColUserEmail).setValue(u.userEmail);
      sheet.getRange(u.row, dataColLastUpdated).setValue(u.timestamp);
    });
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Logs changes to the Changelog sheet.
 * @param {Array} updates - The updates to be logged.
 */
function logChanges(updates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(changelogSheetName);

    // If sheet does not exist, then create and add column headers
    if (!sheet) {
      sheet = ss.insertSheet(changelogSheetName);
      sheet.appendRow(['COMMON ID','MONTH','VALUE','ACTION','NOTES','USER EMAIL','LAST UPDATED']);
    }

    // Remove header filter, if any
    const headerFilter = sheet.getRange('1:1').getFilter();
    if(headerFilter){
      headerFilter.remove();
    }

    // Write changelog
    const changes = updates.map(u => [u.item, u.monthYear, u.value, u.action, u.notes, u.userEmail, u.timestamp]);
    sheet.getRange(sheet.getLastRow() + 1, 1, changes.length, changes[0].length).setValues(changes);

    // Sort sheet by last_updated descending
    sheet.sort(changelogColLastUpdated, false);

    // Apply header filter again
    sheet.getRange(1, 1, sheet.getLastRow(), changelogColLastUpdated).createFilter();
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Get the unique Comment dates from the 'Changelog' sheet.
 */
function getCommentDates() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(changelogSheetName);

    // Get all Comment dates
    const data = sheet.getRange(changelogDataRowStart, changelogColLastUpdated, sheet.getLastRow())
                      .getValues()
                      .flat()
                      .filter(String);

    // Get unique Comment dates
    const commentDates = [...new Set(data.map(d => new Date(d).toDateString()))]
      .filter(Boolean)
      .map(date => Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd-MMM-yy'))
      .sort((a, b) => new Date(b) - new Date(a));

    return commentDates;
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Generate the MERGE statement based on selected Comment dates from the 'Changelog' sheet
 * @param {string} selectedDates - Dates for which the Merge statement needs to be generated.
 */
function generateMergeStatement(selectedDates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(changelogSheetName);

    // Get data
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(changelogDataRowStart, changelogColCommonId, lastRow, changelogTotalColumns)
      .getValues()
      .filter(row => !row.every(cell => cell === ''))
      .filter(row => {
        if (row[changelogArrColLastUpdated]) {
          const sheetDate = row[changelogArrColLastUpdated].toString().substring(0,9);
          return selectedDates.includes(sheetDate);
        }
        return false;
      });

    // SQL Merge statement
    const mergeSql = `MERGE INTO network_rw.forecast_accuracy_items t
USING (
    ${data.map(row =>
    `SELECT '${row[changelogArrColCommonId]}' common_id, TO_DATE('${row[changelogArrColMonth]}','dd-Mon-yy') month, '${row[changelogArrColValue]}' value, '${row[changelogArrColAction]}' action, '${row[changelogArrColNotes]}' notes, '${row[changelogArrColUserEmail]}' user_email, TO_DATE('${row[changelogArrColLastUpdated]}','dd-Mon-yy hh24:mi:ss') last_updated FROM dual`
  ).join(' UNION ALL\n    ')}
) src
ON (t.common_id = src.common_id AND t.month = src.month AND t.last_updated = src.last_updated)
WHEN NOT MATCHED THEN
    INSERT (common_id, month, value, action, notes, user_email, last_updated)
    VALUES (src.common_id, src.month, src.value, src.action, src.notes, src.user_email, src.last_updated);

COMMIT;`;

    return mergeSql;
  } catch (error) {
    Logger.log(error.stack);
  }
}