// main.gs

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Caféen?')
    .addItem('Åbn Admin Panel', 'showSidebar')
    .addSeparator()
    .addItem('Generate Template',           'generateTemplate')
    .addItem('Create Personal Pages',       'createPersonalPages')
    .addItem('Delete Personal Pages',       'deletePersonalPages')
    .addItem('Refresh Schedule Dropdowns',  'updateScheduleDropdowns')
    .addItem('Refresh Schedule',            'refreshHours')
    .addItem('FillSchedule (Experimental)', 'fillSchedule')
    .addItem('Sync Shifts to Calendar',     'syncShiftsToSharedCalendar')
    .addSeparator()
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu('KU Skema')
    .addItem('Vis MitSkema', 'previewSkema')
    .addToUi();
}

function onEdit(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() === SCHEDULE_SHEET_NAME) {
      if (range.getA1Notation() === 'O1') {
        highlightAvailableShifts(e.value);
      }
      updateSchedule(e);
      refreshHours();
    } else {
      activeSheet().getRange(HOURS_CELL).setValue(getHours());
    }
  } catch (error) {
    Logger.log(`Error in onEdit: ${error}`);
  }
}

// ── Menu entry points (include confirmation dialogs) ─────────────────────────

function generateTemplate() {
  confirmTemplateGeneration();
  generateTemplateCore();
}

function createPersonalPages() {
  confirmPersonalPageCreation();
  createPersonalPagesCore();
}

function deletePersonalPages() {
  const ui     = SpreadsheetApp.getUi();
  const result = ui.alert('Slet personlige sider', 'Er du sikker? Dette kan ikke fortrydes.', ui.ButtonSet.YES_NO);
  if (result !== ui.Button.YES) return;
  deletePersonalPagesCore();
}

// ── Core implementations (no confirmation dialogs, safe to call from sidebar) ─

function generateTemplateCore() {
  try {
    const configSheet   = getConfigSheet();
    let year            = configSheet.getRange(YEAR_CELL).getValue();
    const [openWeek, closeWeek] = configSheet
      .getRange(WEEK_RANGE_CELL).getValue().toString().split('-').map(Number);
    const cleaningWeeks = configSheet
      .getRange(CLEANING_WEEK_START_CELL).getValue().toString().split(',').map(Number);

    clearSchedule();

    let offset = 4;
    if (cleaningWeeks.length > 1) offset = createWeek(offset, cleaningWeeks[0], year);

    let week = openWeek;
    while (true) {
      offset = createWeek(offset, week, year);
      if (week === closeWeek) break;
      week++;
      if (week > 52) { week = 1; year += 1; }
    }

    offset = createDay(offset, closeWeek + 1, 5, year, true);
    offset = createDay(offset, closeWeek + 2, 5, year, true);

    const finalCleaningWeek = cleaningWeeks.length > 1 ? cleaningWeeks[1] : cleaningWeeks[0];
    if (finalCleaningWeek) offset = createWeek(offset, finalCleaningWeek, year);

    getConfigSheet().getRange(TOTAL_HOURS_CELL).setValue(getTotalHours());
  } catch (error) {
    Logger.log(`Error in generateTemplateCore: ${error}`);
    throw error;
  }
}

function createPersonalPagesCore() {
  try {
    const volunteerList = getVolunteerList();
    if (volunteerList.length === 0) throw new Error('Volunteer list is empty.');

    const spreadsheet   = ss();
    const templateSheet = getTemplateSheet();

    volunteerList.forEach(volunteer => {
      const sheetName     = volunteer[0];
      const existingSheet = spreadsheet.getSheetByName(sheetName);
      if (existingSheet) spreadsheet.deleteSheet(existingSheet);

      const newSheet = templateSheet.copyTo(spreadsheet).setName(sheetName);
      newSheet.getRange('A1').setValue(sheetName);
      newSheet.getRange('D:D').setValue(`='${SCHEDULE_SHEET_NAME}'!D:D`);

      const adminEmails = getConfigSheet()
        .getRange(ADMIN_EMAILS_CELL)
        .getValue()
        .toString()
        .split(',')
        .map(e => e.trim())
        .filter(Boolean);

      const permissions = newSheet.protect();
      if (adminEmails.length) permissions.addEditors(adminEmails);

      try {
        const userEmail = volunteer[5];
        if (userEmail) {
          AdminDirectory.Users.get(userEmail); // throws if not found
          permissions.addEditor(userEmail);
        }
      } catch (e) {
        Logger.log(`Error adding editor ${volunteer[5]}: ${e.message}`);
      }

      permissions.setWarningOnly(false);
    });

    return volunteerList.length;
  } catch (error) {
    Logger.log(`Error in createPersonalPagesCore: ${error}`);
    throw error;
  }
}

function deletePersonalPagesCore() {
  try {
    const spreadsheet   = ss();
    const volunteerList = getVolunteerList();
    volunteerList.forEach(volunteer => {
      const sheet = spreadsheet.getSheetByName(volunteer[0]);
      if (sheet) spreadsheet.deleteSheet(sheet);
    });
    return volunteerList.length;
  } catch (error) {
    Logger.log(`Error in deletePersonalPagesCore: ${error}`);
    throw error;
  }
}

// ── Hours ────────────────────────────────────────────────────────────────────

function refreshHours() {
  try {
    const scheduleSheet = getScheduleSheet();
    const configSheet   = getConfigSheet();
    const choices       = scheduleSheet.getRange(4, 2, 100, 10).getValues();
    const dayHoursData  = configSheet.getRange('A9:D16').getValues();
    const cleaningWeeks = configSheet.getRange(CLEANING_WEEK_START_CELL).getValue().toString().split(',');

    const values   = scheduleSheet.getRange('N5:R39').getValues();
    const newHours = values.map(row => getSelectedHours(row[0], choices, dayHoursData, cleaningWeeks));

    scheduleSheet.getRange('P5:P39').setValues(newHours.map(h => [h[0]]));
    scheduleSheet.getRange('R5:R39').setValues(newHours.map(h => [h[1]]));
  } catch (error) {
    Logger.log(`Error in refreshHours: ${error}`);
    throw error;
  }
}
