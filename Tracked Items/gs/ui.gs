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

    // Hide sheets
    // hideSheet(SHEET_CONFIG.ADMIN.NAME);
    // hideSheet(SHEET_CONFIG.DROPDOWNS.NAME);
    // hideSheet(SHEET_CONFIG.DATA.NAME);

    // Activate Input Sheet
    activateCell(SHEET_CONFIG.INPUT.NAME);
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Show sidebar form.
 * @param {string} sidebarHtml - The html element to show as sidebar.
 * @param {string} sidebarTitle - The title of the sidebar.
 */
function showSidebar(sidebarHtml, sidebarTitle) {
  try {
    // Show Sidebar
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile(sidebarHtml).setTitle(sidebarTitle);
    ui.showSidebar(html);
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

/**
 * Shows the requests sidebar form.
 */
function showRequestsSidebar() {
  try {
    // Run only if current user has User access
    if (!checkAccess('User')) return;

    // Show Sidebar
    showSidebar('requests', 'Admin Requests');
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

    // Activate 'Requests' sheet
    activateCell(SHEET_CONFIG.REQUESTS.NAME);

    // Get unique Comment Dates
    commentDates = getCommentDates();

    // Show Sidebar form
    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createTemplateFromFile('data_update');
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
    activateCell(SHEET_CONFIG.DATA.NAME);

    // Show Sidebar
    showSidebar('data_get', 'Get Data');
  } catch (error) {
    Logger.log(error.stack);
  }
}

/**
 * Shows the role management sidebar form.
 * @param {string} action - The action to perform ('add' or 'remove')
 */
function showRoleManagementSidebar(action) {
  try {
    // Run only if current user has Admin access
    if (!checkAccess('Admin')) return;

    // Activate 'Admin' sheet
    activateCell(SHEET_CONFIG.ADMIN.NAME);

    // Show Sidebar
    const html = HtmlService.createHtmlOutput(
      HtmlService.createHtmlOutputFromFile('role_management')
        .getContent()
        .replace('let currentAction = \'\';', `let currentAction = '${action}';`)
    )
    .setTitle(action === 'add' ? 'Add Role' : 'Remove Role');

    SpreadsheetApp.getUi().showSidebar(html);
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
    showSidebar('reset_sheets', 'Reset Sheets');
  } catch (error) {
    Logger.log(error.stack);
  }
}