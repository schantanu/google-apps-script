/**
 * Adds a custom menu to the Google Sheets UI on open.
 */
function onOpen(e) {
  try {
    // Create custom menu
    const ui = SpreadsheetApp.getUi();

    ui.createMenu('✍️ Editor')
        .addItem('📖 User Guide', 'showUserGuide')
        .addItem('👤 User Console', 'showUserConsole')
      .addSeparator()
        .addItem('⚙️ Admin Console', 'showAdminConsole')
      .addToUi();

    // Hide Admin Sheet
    // hideSheet(adminSheetName);

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
 * Shows the user console sidebar.
 */
function showUserConsole() {
  try {
    // Run only if current user has Admin access
    if (!checkAccess('User')) return;

    // Show Sidebar
    showSidebar('user_console', 'User Console');
  } catch (error) {
    Logger.log(error.stack);
  }
}

// function showAttributeManager(type) {
//   try {
//     // Run only if current user has User access
//     if (!checkAccess('User')) return;

//     // Show Sidebar with type parameter
//     const html = HtmlService.createHtmlOutputFromFile('attr_manager')
//       .setTitle('Update Attributes');

//     // Append initialization script
//     const template = html.getContent();
//     const output = template + `<script>window.onload = function() { initializeForm('${type}'); }</script>`;

//     // Show sidebar
//     SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(output));
//   } catch (error) {
//     Logger.log(error.stack);
//   }
// }

/**
 * Shows the attribute manager sidebar with initialization
 * @param {string} type - The type of attributes to manage ('common' or 'item')
 */
function showAttributeManager(type) {
  try {
    if (!checkAccess('User')) return;

    const html = HtmlService.createTemplateFromFile('attr_manager');
    html.attributeType = type;

    const output = html.evaluate()
      .setTitle('Update Attributes')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    SpreadsheetApp.getUi().showSidebar(output);
  } catch (error) {
    Logger.log(error.stack);
  }
}

function showUserRequests() {
  try {
    // Run only if current user has User access
    if (!checkAccess('User')) return;

    // Show Sidebar
    showSidebar('user_requests', 'My Requests');
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
    showSidebar('admin_console', 'Admin Console');
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