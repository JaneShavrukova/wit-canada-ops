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
      .addItem('Refresh Groups',                 'buildGroupsView')
      .addSeparator()
      .addItem('✍️ Create email signature',      'openSignatureGeneratorForRow'))

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


// ============================================================
// Signature Generator
// ============================================================

function openSignatureGeneratorForRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET.MAIN) {
    showAlert('Wrong sheet', 'Please select a row in the WIT_Members sheet first.');
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row < CONFIG.SHEET.DATA_START_ROW) {
    showAlert('No member selected', 'Please click on a member row first.');
    return;
  }

  const colMap  = getColumnIndexMap(sheet);
  const get     = (header) => safeString(sheet.getRange(row, colMap[normalizeHeader(header)]).getValue());

  const firstName = get(CONFIG.HEADERS.FIRST_NAME);
  const lastName  = get(CONFIG.HEADERS.LAST_NAME);
  const role      = get(CONFIG.HEADERS.ROLE);
  const witEmail  = get(CONFIG.HEADERS.WIT_EMAIL);

  const url = CONFIG.URLS.SIGNATURE_GENERATOR
    + '&firstName=' + encodeURIComponent(firstName)
    + '&lastName='  + encodeURIComponent(lastName)
    + '&role='      + encodeURIComponent(role)
    + '&witEmail='  + encodeURIComponent(witEmail);

  const html = HtmlService
    .createHtmlOutput('<script>window.open("' + url + '","_blank");google.script.host.close();</script>')
    .setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Signature Generator…');
}
