// ============================================================
// WIT Canada — Main Menu
// This is the ONLY onOpen() in the project.
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('WIT Operations')

    // ── 📘 Documentation ────────────────────────────────────
    .addSubMenu(ui.createMenu('📘 Documentation')
      .addItem('System Guide',                   'showFileGuide')
      .addItem('Onboarding SOP — Leads',         'showNewMemberGuideSidebar')
      .addItem('Onboarding Guide — Members  ↗',  'openNewMemberGuideExternal'))

    .addSeparator()

    // ── ⚙️ Actions ──────────────────────────────────────────
    .addSubMenu(ui.createMenu('⚙️ Actions')
      .addItem('Refresh Email Requests',         'buildEmailRequestsReport')
      .addItem('Refresh Photos & Bios',          'syncPhotosAndBios')
      .addItem('Refresh Groups',                 'buildGroupsView'))

    .addSeparator()

    // ── 🔍 Reports ───────────────────────────────────────────
    .addSubMenu(ui.createMenu('🔍 Reports')
      .addItem('Open Email Requests',            'openEmailRequestsReport')
      .addItem('Open Missing Bios',              'openPhotoBioReport')
      .addItem('Open Groups Status',             'openGroupsReport'))

    .addToUi();
}


// ============================================================
// Report navigation helpers
// ============================================================

function openEmailRequestsReport() {
  _openReportSheet(EMAIL_REQUESTS_SHEET);
}

function openPhotoBioReport() {
  _openReportSheet(REPORT_SHEET_NAME);
}

function openGroupsReport() {
  _openReportSheet(GROUPS_VIEW_SHEET);
}

function _openReportSheet(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Report not found',
      `"${name}" hasn't been generated yet.\n\nRun the corresponding Refresh action first.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  ss.setActiveSheet(sheet);
}


// ============================================================
// Group access wrappers (with permission handling)
// ============================================================

function syncAllMembers() {
  runWithAlert(() => {
    try {
      _syncAllMembers();
    } catch (e) {
      if (isPermissionError(e)) {
        showPermissionAlert();
      } else {
        throw e;
      }
    }
  });
}

function syncSelectedMember() {
  runWithAlert(() => {
    try {
      _syncSelectedMember();
    } catch (e) {
      if (isPermissionError(e)) {
        showPermissionAlert();
      } else {
        throw e;
      }
    }
  });
}


// ============================================================
// Helpers (menu-specific)
// ============================================================

function isPermissionError(e) {
  return e.message && e.message.toLowerCase().includes('permission');
}

function showPermissionAlert() {
  SpreadsheetApp.getUi().alert(
    'Permission required',
    'Assigning group access requires the Groups Admin role in Google Workspace.\n\n' +
    'Please contact the Ops Lead — she will run the sync for you:\n' +
    CONFIG.EMAIL.OPS_LEAD,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
