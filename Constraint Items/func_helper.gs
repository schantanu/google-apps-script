// Helper Functions
// ----------------------------------------------------------------------------------------------------------
//
// getRoleCol(role)                 - Get column number for the requested role
// checkAccess(role)                - Check if role has corresponding access
// addAccess(role, emailAddresses)  - Add access for corresponding role
// getEmailsByRole(role)            - Get a list of current email addressess for a role.
// removeAccess(role, email)        - Remove access for a role.
// updateSheetsProtection()         - Check if the current user has Admin access
// protectSheets()                  - Updates protection to the Sheet
//
// ----------------------------------------------------------------------------------------------------------

/**
 * Get column number for the requested role
 * @param {string} role - The role to get the column for, Admins or Users.
 */
function getRoleCol(role) {
  try {
    return (role === 'Admin') ? adminColAdmins
      : (role === 'User') ? adminColUsers
      : null;
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Check if the current user has Role privileges.
 * @param {string} role - The role to check privileges for.
 */
function checkAccess(role) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(adminSheetName);
    const ui = SpreadsheetApp.getUi();

    // Get current user email
    const currentUserEmail = Session.getActiveUser().getEmail();

    // Get role column from Admin sheet
    const roleCol = getRoleCol(role);

    // Get list of email addresses to check access privileges against
    const roleEmails = role === 'User'
      ? sheet.getRange(adminDataRowStart, adminColAdmins, sheet.getLastRow(), adminUserColsTotal).getValues().flat().filter(String)
      : sheet.getRange(adminDataRowStart, roleCol, sheet.getLastRow()).getValues().flat().filter(String);

    // Show alert and exit if current user is not admin
    if (!roleEmails.includes(currentUserEmail)) {
      const response = ui.alert(`ERROR: \n\n You do not have ${role} access. Please request ${role} access from the current Admin to run this function.`);
      return false;
    }

    // Return true if current user has role access
    return true;

  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Get a list of current email addressess for a role.
 * @param {string} role - The role type.
 */
function getEmailsByRole(role) {
  try {
    // Get the Admin sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(adminSheetName);

    // Get role column from Admin sheet
    const roleCol = getRoleCol(role);

    // Get emails from the role column
    const emails = sheet.getRange(adminDataRowStart, roleCol, sheet.getLastRow())
                     .getValues()
                     .flat()
                     .filter(email => email);

    return emails;
  } catch (error) {
    Logger.log(error.stack);
    return [];
  }
}

/**
 * Add a role access.
 * @param {string} role - The role type.
 * @param {array} emailAddresses - A list of addresses to give role access to.
 */
function addAccess(role, emailAddresses) {
  try {
    // Run only if current user has admin access
    if (!checkAccess('Admin')) return;

    // Get the Admin sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(adminSheetName);

    // Get role column from Admin sheet
    const roleCol = getRoleCol(role);

    // Fetch all existing emails from the role column
    const existingEmails = getEmailsByRole(role);

    // Split and trim email array
    const emails = emailAddresses.split(',').map(email => email.trim());

    //Create temporary protection to validate user
    const tempProtection = sheet.protect().setDescription('Temporary protection for validation');
    const invalidEmails = [];
    const validEmails = [];

    emails.forEach(email=> {
      if (existingEmails.includes(email)) return;

      try{
        tempProtection.addEditor(email);
        validEmails.push(email);
      } catch (e) {
        invalidEmails.push(email);
      }
    });

    // Delete temporary protection
    tempProtection.remove();

    //Add valid Emails to sheet
    if (validEmails.length > 0) {
      validEmails.forEach(email => {
        var lastRow = getColumnLastRow(adminSheetName, roleCol, adminDataRowStart);
        sheet.getRange(lastRow, roleCol).setValue(email);
      });

      // Apply all pending Spreadsheet changes before proceeding
      SpreadsheetApp.flush();

      // Update Sheet protection
      updateSheetsProtection();
    }

    // Throw error for invalid users
    if (invalidEmails.length > 0){
      throw new Error('Invalid email(s):' + invalidEmails.join(', '));
    }
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Remove access for a role.
 * @param {string} role - The role type.
 * @param {string} email - The user email.
 */
function removeAccess(role, email) {
  try {
    // Run only if current user has admin access
    if (!checkAccess('Admin')) return;

    // Get the Admin sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(adminSheetName);

    // Get role column from Admin sheet
    const roleCol = getRoleCol(role);

    // Get existing user emails
    const data = sheet.getRange(adminDataRowStart, roleCol, sheet.getLastRow())
                      .getValues()
                      .flat()
                      .filter(name => name);

    // Remove the email from existing user emails
    const updatedValues = data.filter(name => name !== email);

    // Update the role column by removing the email
    while (updatedValues.length < data.length) updatedValues.push('');
    sheet.getRange(adminDataRowStart, roleCol, updatedValues.length, 1)
        .setValues(updatedValues.map(value => [value]));

    // Apply all pending Spreadsheet changes before proceeding
    SpreadsheetApp.flush();

    // Update Sheet protection
    updateSheetsProtection();
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Update all the Sheets protection based on customized user access.
 */
function updateSheetsProtection() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(adminSheetName);

    // Get Admins
    const admins = sheet.getRange(adminDataRowStart, adminColAdmins, sheet.getLastRow()).getValues().flat().filter(String);

    // Protect all sheets and give only Admins access
    ss.getSheets().forEach(sheet => {
      sheetName = sheet.getName();

      // Protect the entire sheet and allow only Admins to edit
      const protection = sheet.protect().setDescription('Admin access only');

      // Remove existing editors
      protection.removeEditors(protection.getEditors());

      // Check and add valid admin users
      admins.forEach(email => {
        try {
          // Add editor
          protection.addEditor(email);
          Logger.log(`Access successfully granted to '${sheetName}' sheet for ${email}.`);
        } catch (error) {
          Logger.log(`Invalid admin skipped: ${email}. Error: ${error.message}`);
        }
      });
    });

    // Loop through the array and call the protectSheets function for each sheet
    sheetsEditConfig.forEach(function(sheetConfig) {
      protectSheets(sheetConfig.sheetName, sheetConfig.editRowStart, sheetConfig.editColStart, sheetConfig.editTotalColumns);
    });
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Update all the Sheets protection based on customized user access.
 */
function protectSheets(sheetName, editRowStart, editColStart, editTotalColumns) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    // Throw error if sheet does not exist
    if (!sheet) throw new Error('Sheet with name ' + sheetName + ' not found.');

    // Get Users
    const adminSheet = ss.getSheetByName(adminSheetName);
    const users = adminSheet.getRange(adminDataRowStart, adminColUsers, adminSheet.getLastRow()).getValues().flat().filter(String);

    // Allow users to edit the specified range
    const protection = sheet.protect().setDescription('Admin access only');
    const unprotectedRange = sheet.getRange(editRowStart, editColStart, sheet.getMaxRows(), editTotalColumns);
    protection.setUnprotectedRanges([unprotectedRange]);

    // Check and add valid users
    users.forEach(email => {
      try {
        protection.addEditor(email);
        Logger.log(`Access successfully granted to '${sheetName}' sheet for ${email}.`);
      } catch (error) {
        Logger.log(`Invalid user skipped: ${email}. Error: ${error.message}`);
      }
    });
  } catch (error) {
    Logger.log(error.stack);
  }
}
