// sidebar.gs
//
// Server-side functions called from sidebar_html.html via google.script.run.
// Every public function starts with checkAdminAccess() — two layers of protection:
//   1. showSidebar() blocks the panel from opening for non-admins.
//   2. Each sidebar function re-checks, so calling google.script.run directly
//      from the browser console is also blocked.
//
// Admin emails are read from ADMIN_EMAILS_CELL in the Configuration sheet
// (comma-separated). Update that cell to change who has access — no code change needed.

// ── Access control ────────────────────────────────────────────────────────────

/**
 * Throws if the current user is not in the admin allowlist.
 * Called at the top of showSidebar() and every sidebar* function.
 */
function checkAdminAccess() {
  const currentEmail = Session.getActiveUser().getEmail();
  if (!currentEmail) {
    throw new Error('Kunne ikke verificere brugerens e-mail. Prøv at genindlæse arket.');
  }

  const cellValue   = getConfigSheet().getRange(ADMIN_EMAILS_CELL).getValue().toString();
  const adminEmails = cellValue.split(',').map(e => e.trim().toLowerCase()).filter(Boolean);

  if (adminEmails.length === 0) {
    throw new Error(`Ingen admin-e-mails konfigureret. Tilføj dem i celle A3/B3 i konfigurationsarket.`);
  }

  if (!adminEmails.includes(currentEmail.toLowerCase())) {
    throw new Error(`Adgang nægtet for ${currentEmail}.`);
  }
}

// ── Sidebar entry point ───────────────────────────────────────────────────────

function showSidebar() {
  try {
    checkAdminAccess();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Adgang nægtet', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('sidebar_html').setTitle('Caféen? Admin').setWidth(300)
  );
}

// ── Skabelon ──────────────────────────────────────────────────────────────────

function sidebarGenerateTemplate() {
  checkAdminAccess();
  generateTemplateCore();
  return 'Skabelon genereret ✓';
}

// ── Frivillige ────────────────────────────────────────────────────────────────

function sidebarCreatePersonalPages() {
  checkAdminAccess();
  const count = createPersonalPagesCore();
  return `Personlige sider oprettet for ${count} frivillige ✓`;
}

function sidebarDeletePersonalPages() {
  checkAdminAccess();
  const count = deletePersonalPagesCore();
  return `${count} personlige sider slettet ✓`;
}

// ── Skema ─────────────────────────────────────────────────────────────────────

function sidebarUpdateDropdowns() {
  checkAdminAccess();
  updateScheduleDropdowns();
  return 'Dropdowns opdateret ✓';
}

function sidebarRefreshHours() {
  checkAdminAccess();
  refreshHours();
  return 'Timer opdateret ✓';
}

function sidebarFillSchedule() {
  checkAdminAccess();
  fillSchedule();
  return 'Skemalægning fuldført ✓';
}

// ── Kalender ──────────────────────────────────────────────────────────────────

function sidebarSyncCalendar() {
  checkAdminAccess();
  const count = syncShiftsToSharedCalendarCore();
  return `${count} kalenderbegivenheder oprettet ✓`;
}

// ── E-mails ───────────────────────────────────────────────────────────────────

function sidebarSendReminders() {
  checkAdminAccess();
  const count = sendReminderEmailsCore();
  return count > 0
    ? `Påmindelser sendt til ${count} frivillige ✓`
    : 'Ingen frivillige mangler at udfylde skemaet ✓';
}

// ── Eksport ───────────────────────────────────────────────────────────────────

function sidebarExportWiki() {
  checkAdminAccess();
  parseToWiki();
  return 'Wiki-fil eksporteret til Google Drive ✓';
}
