// utils.gs

// ── Sheet names ──────────────────────────────────────────────────────────────
const TEMPLATE_SHEET_NAME = "Template";
const SCHEDULE_SHEET_NAME = "Endeligt Skema";
const CONFIG_SHEET_NAME   = "Configuration";

// ── Cell references ──────────────────────────────────────────────────────────
const HOURS_CELL   = "F3";
const DATE_CELL    = "A";
const WEEKDAY_CELL = "B";
const WEEK_CELL    = "C";
const NOTE_CELL    = "D";
const PREFERENCE_CELL = "E";

const PREFERENCE_LIST        = ["Kan", "Åbne", "Lukke", "Kan Ikke"];
const PREFERENCE_LIST_FRIDAY = ["Kan", "Åbne", "Mellem", "Åbne & Mellem", "Mellem & Lukke", "Åbne & Lukke", "Lukke", "Kan Ikke"];

const CLEANING_WEEK_START_CELL = 'B6';
const YEAR_CELL        = 'B4';
const WEEK_RANGE_CELL  = 'B5';
const TOTAL_HOURS_CELL  = 'B18';
const ADMIN_EMAILS_CELL   = 'B3';  // comma-separated — controls sidebar access AND sheet editor permissions
const CALENDAR_ID_CELL    = 'B7';  // shared Google Calendar ID for shift sync
const AVERAGE_HOURS_CELL  = 'B21'; // target hours per volunteer for fair shift distribution

// ── Layout constants — update here if the spreadsheet structure changes ──────
const VOLUNTEER_LIST_START_ROW = 5;
const VOLUNTEER_LIST_COL       = 14;  // column N
const VOLUNTEER_LIST_NUM_COLS  = 7;
const SCHEDULE_START_ROW       = 4;
const SCHEDULE_END_ROW         = 71;
const MAX_VOLUNTEER_HOURS      = 30;

// ── Spreadsheet helpers ──────────────────────────────────────────────────────
const ss          = () => SpreadsheetApp.getActive();
const activeSheet = () => ss().getActiveSheet();

// Lazy sheet getters — throw a descriptive error if a required sheet is missing
// rather than silently returning null and crashing deep inside a function.
function getTemplateSheet() {
  const s = ss().getSheetByName(TEMPLATE_SHEET_NAME);
  if (!s) throw new Error(`Required sheet "${TEMPLATE_SHEET_NAME}" not found.`);
  return s;
}
function getScheduleSheet() {
  const s = ss().getSheetByName(SCHEDULE_SHEET_NAME);
  if (!s) throw new Error(`Required sheet "${SCHEDULE_SHEET_NAME}" not found.`);
  return s;
}
function getConfigSheet() {
  const s = ss().getSheetByName(CONFIG_SHEET_NAME);
  if (!s) throw new Error(`Required sheet "${CONFIG_SHEET_NAME}" not found.`);
  return s;
}

// ── Confirmation dialogs ─────────────────────────────────────────────────────
function confirmTemplateGeneration() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Generating a new template will delete all data!',
    'Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );
  if (result === ui.Button.NO) throw new Error('Template generation cancelled.');
}

function confirmPersonalPageCreation() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Generating new pages will delete all old pages first! Including ALL inputs!',
    'Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );
  if (result === ui.Button.NO) throw new Error('Personal page creation cancelled.');
}

// ── Schedule utilities ───────────────────────────────────────────────────────
function clearSchedule() {
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID_THICK;
  getTemplateSheet().getRange(4, 1, 100, 5).clearContent()
    .setBorder(false, false, false, false, false, false, "black", borderStyle);
  getScheduleSheet().getRange(4, 5, 100, 6).clearContent()
    .setBorder(false, false, false, false, false, false, "black", borderStyle);
}

function getFirstEmptyRowByColumnArray(column, sheet, offset = 0) {
  const values = sheet.getRange(`${column}:${column}`).getValues();
  let ct = offset;
  while (values[ct] && values[ct][0] !== "") ct++;
  return ct + 1;
}

function getVolunteerList() {
  const sheet = getScheduleSheet();
  const firstEmptyRow = getFirstEmptyRowByColumnArray("N", sheet, VOLUNTEER_LIST_START_ROW - 1);
  const numRows = firstEmptyRow - VOLUNTEER_LIST_START_ROW;
  if (numRows <= 0) return [];
  return sheet
    .getRange(VOLUNTEER_LIST_START_ROW, VOLUNTEER_LIST_COL, numRows, VOLUNTEER_LIST_NUM_COLS)
    .getValues()
    .filter(row => row[0] !== "");
}

// ── Shift highlighting ───────────────────────────────────────────────────────
function highlightAvailableShifts(volunteerName) {
  try {
    const sheet = getScheduleSheet();
    sheet.getRange(`E${SCHEDULE_START_ROW}:K${SCHEDULE_END_ROW}`).setBackground(null);

    if (!volunteerName) return;

    const preferences = getUserPreferences(volunteerName);
    const shiftRange  = sheet.getRange(`E${SCHEDULE_START_ROW}:K${SCHEDULE_END_ROW}`);

    for (let row = 0; row < preferences.length; row++) {
      const preference = preferences[row][0];
      if (!preference || preference === 'Kan Ikke') continue;

      let cols = [];
      switch (preference) {
        case 'Åbne':           cols = [1, 2];             break;
        case 'Mellem':         cols = [3, 4];             break;
        case 'Lukke':          cols = [5, 6, 7];          break;
        case 'Kan':            cols = [1, 2, 3, 4, 5, 6, 7]; break;
        case 'Mellem & Lukke': cols = [3, 4, 5, 6, 7];   break;
        case 'Åbne & Mellem':  cols = [1, 2, 3, 4];      break;
        case 'Åbne & Lukke':   cols = [1, 2, 5, 6, 7];   break;
      }
      cols.forEach(col => shiftRange.getCell(row + 1, col).setBackground('#C6EFCE'));
    }
  } catch (error) {
    Logger.log(`Error in highlightAvailableShifts: ${error}`);
  }
}

// ── Wiki export ──────────────────────────────────────────────────────────────
function parseToWiki() {
  const scheduleSheet = getScheduleSheet();
  const configSheet   = getConfigSheet();
  const schedule = scheduleSheet.getRange(`A${SCHEDULE_START_ROW}:K65`).getValues();
  const [openWeek, closeWeek] = configSheet.getRange(WEEK_RANGE_CELL).getValue().toString().split('-').map(Number);
  const tz = ss().getSpreadsheetTimeZone();

  let offset = 0;
  let content = "";
  let week = openWeek;

  do {
    content += "||border=1\n||!Dato||!Dag||!Uge||!Bemærkning||!Åbner||!Mellem||!Lukker||\n";

    for (let j = 0; j < 6; j++) {
      if (offset >= schedule.length || !schedule[offset][0]) { offset++; continue; }

      const d = new Date(schedule[offset][0]);
      d.setHours(12, 0, 0, 0); // normalise to midday — avoids CET↔CEST off-by-one at midnight
      const dateString = Utilities.formatDate(d, tz, "dd-MM-yyyy");

      const [day, weekNumber, comment] = schedule[offset].slice(1, 4);
      const openers = schedule[offset].slice(4, 6).filter(n => n).join(", ");
      const mellem  = schedule[offset].slice(6, 8).filter(n => n).join(", ");
      const closers = schedule[offset].slice(8, 11).filter(n => n).join(", ");
      content += `|| ${dateString} || ${day} || ${weekNumber} || ${comment || ''} || ${openers || ''} || ${mellem || ''} || ${closers || ''} ||\n`;
      offset++;
    }

    content += "\n\n";
    week++;
    if (week > 52) week = 1;
  } while (week !== ((closeWeek + 1) % 52 || 52));

  const fileName = `Schedule ${Utilities.formatDate(new Date(), tz, "yyyy-MM-dd")}.txt`;
  DriveApp.createFile(fileName, content);
  Logger.log(fileName);
}
