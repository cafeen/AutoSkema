// schedule.gs

// ── Schedule editing ─────────────────────────────────────────────────────────

function updateSchedule(e) {
  try {
    const scheduleSheet   = getScheduleSheet();
    const configSheet     = getConfigSheet();
    const userChangedCellRange = scheduleSheet.getRange('N5:T36');
    const userChoicesRange     = scheduleSheet.getRange(4, 2, 100, 9);
    const dayHoursData         = configSheet.getRange('A9:D16').getValues();
    const cleaningWeeks        = configSheet.getRange(CLEANING_WEEK_START_CELL).getValue().toString().split(',');

    const valuesData = userChangedCellRange.getValues();
    const userIndex  = valuesData.findIndex(i => i[0] === e.value || i[0] === e.oldValue);

    if (userIndex !== -1) {
      const choices    = userChoicesRange.getValues();
      const updateData = [];

      [e.oldValue, e.value].forEach(user => {
        const hours = getSelectedHours(user, choices, dayHoursData, cleaningWeeks);
        updateData.push([userIndex + 1, 3, hours[0]]);
        updateData.push([userIndex + 1, 5, hours[1]]);
      });

      updateData.forEach(([row, col, val]) => {
        userChangedCellRange.getCell(row, col).setValue(val);
      });
    }
  } catch (error) {
    Logger.log(`Error in updateSchedule: ${error}`);
  }
}

// ── Template generation ──────────────────────────────────────────────────────

// ISO weekdays for a standard café week: Mon–Fri (1–5) then Sun (7), skipping Sat.
const CAFE_WEEKDAYS = [1, 2, 3, 4, 5, 7];

function createWeek(offset, week, year) {
  const templateSheet = getTemplateSheet();
  const scheduleSheet = getScheduleSheet();
  const borderStyle   = SpreadsheetApp.BorderStyle.SOLID_THICK;

  templateSheet.getRange(`A${offset}:E${offset + 5}`)
    .setBorder(true, true, true, true, false, false, 'black', borderStyle);
  scheduleSheet.getRange(`A${offset}:J${offset + 5}`)
    .setBorder(true, true, true, true, false, false, 'black', borderStyle);

  for (const weekday of CAFE_WEEKDAYS) {
    offset = createDay(offset, week, weekday, year);
  }
  return offset;
}

function createDay(offset, week, weekday, year, withBorder = false) {
  const templateSheet = getTemplateSheet();
  const scheduleSheet = getScheduleSheet();
  const dateFormula   = `=DATE(${year}, 1, -3 + 7 * ${week} - WEEKDAY(DATE(${year}, 1, 4), 2) + ${weekday})`;

  templateSheet.getRange(`${DATE_CELL}${offset}`).setValue(dateFormula);
  templateSheet.getRange(`${WEEKDAY_CELL}${offset}`).setValue(`=UPPER(TEXT(${DATE_CELL}${offset}, "ddd"))`);
  templateSheet.getRange(`${WEEK_CELL}${offset}`).setValue(week);
  templateSheet.getRange(`${NOTE_CELL}${offset}`).setValue(`='${SCHEDULE_SHEET_NAME}'!D:D`);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(weekday === 5 ? PREFERENCE_LIST_FRIDAY : PREFERENCE_LIST, true)
    .build();
  templateSheet.getRange(`${PREFERENCE_CELL}${offset}`).setDataValidation(rule);

  if (withBorder) {
    const borderStyle = SpreadsheetApp.BorderStyle.SOLID_THICK;
    templateSheet.getRange(`A${offset}:E${offset + 1}`)
      .setBorder(true, true, true, true, false, false, 'black', borderStyle);
    scheduleSheet.getRange(`A${offset}:J${offset + 1}`)
      .setBorder(true, true, true, true, false, false, 'black', borderStyle);
  }

  return offset + 1;
}

// ── Dropdown refresh ─────────────────────────────────────────────────────────

function updateScheduleDropdowns() {
  try {
    const scheduleSheet = getScheduleSheet();
    const volunteerList = getVolunteerList();
    const userPrefs = volunteerList.map(volunteer => ({
      name:  volunteer[0],
      prefs: getUserPreferences(volunteer[0]),
    }));

    const options  = scheduleSheet.getRange('E4:L71');
    const weekDays = scheduleSheet.getRange('B4:B71').getValues().flat();

    const numRows = options.getNumRows();
    const numCols = options.getNumColumns();
    const newDataValidations = Array.from({ length: numRows }, () => []);

    for (let i = 0; i < numRows; i++) {
      const rowValidations = [];
      const isFriday = weekDays[i] === 'FRI';

      for (let j = 0; j < numCols; j++) {
        let shiftType;
        if (j < 2)      shiftType = 'Åbne';
        else if (j < 4) shiftType = 'Mellem';
        else            shiftType = 'Lukke';

        const availableVolunteers = userPrefs
          .filter(user => {
            const pref = user.prefs[i]?.[0];
            if (!pref || pref === 'Kan Ikke') return false;
            if (pref === 'Kan' || pref === shiftType) return true;
            if (!isFriday) return false;
            if (pref === 'Åbne & Mellem')  return shiftType === 'Åbne'  || shiftType === 'Mellem';
            if (pref === 'Mellem & Lukke') return shiftType === 'Mellem' || shiftType === 'Lukke';
            if (pref === 'Åbne & Lukke')   return shiftType === 'Åbne'  || shiftType === 'Lukke';
            return false;
          })
          .map(user => user.name);

        rowValidations.push(
          SpreadsheetApp.newDataValidation().requireValueInList(availableVolunteers, true).build()
        );
      }
      newDataValidations[i] = rowValidations;
    }

    options.setDataValidations(newDataValidations);
  } catch (error) {
    Logger.log(`Error in updateScheduleDropdowns: ${error}`);
  }
}

// ── User preferences ─────────────────────────────────────────────────────────

function getUserPreferences(user) {
  try {
    const sheet = ss().getSheetByName(user);
    if (!sheet) throw new Error(`Sheet for user ${user} not found.`);
    return sheet.getRange(`E${SCHEDULE_START_ROW}:E${SCHEDULE_END_ROW}`).getValues();
  } catch (error) {
    Logger.log(`Error in getUserPreferences for user ${user}: ${error}`);
    return [];
  }
}

// ── Hours calculation ────────────────────────────────────────────────────────

function getSelectedHours(user, choices, dayHoursData, cleaningWeeks) {
  const result = choices
    .map(day => {
      if (!day.includes(user)) return { totalHours: 0, daysWorked: 0 };

      let weekday = day[0];
      const week  = day[1];
      if (cleaningWeeks.includes(week.toString()) && weekday !== 'SUN') weekday = 'CLEANING WEEK';
      if (!weekday) return { totalHours: 0, daysWorked: 0 };

      const dayHours = dayHoursData.find(x => x[0] === weekday);
      if (!dayHours) return { totalHours: 0, daysWorked: 0 };

      const shiftIndex = day.indexOf(user);
      let shiftHours = 0;
      // Columns: E=3, F=4 → Åbne | G=5, H=6 → Mellem | I=7, J=8, K=9 → Lukke
      if      (shiftIndex === 3 || shiftIndex === 4)                      shiftHours = parseFloat(dayHours[1]) || 0;
      else if (shiftIndex === 5 || shiftIndex === 6)                      shiftHours = parseFloat(dayHours[2]) || 0;
      else if (shiftIndex === 7 || shiftIndex === 8 || shiftIndex === 9)  shiftHours = parseFloat(dayHours[3]) || 0;

      return { totalHours: shiftHours, daysWorked: 1 };
    })
    .reduce(
      (acc, curr) => ({ totalHours: acc.totalHours + curr.totalHours, daysWorked: acc.daysWorked + curr.daysWorked }),
      { totalHours: 0, daysWorked: 0 }
    );

  return [result.totalHours, result.daysWorked];
}

// weekDayHours is a pure reducer — callers read config data once and pass it in
// to avoid repeated spreadsheet API calls inside the loop.
function weekDayHours(total, value, dayHoursData, cleaningWeeksList) {
  let [weekday, week, note, shift] = value;

  if (note === 'LUKKET') return total;

  if (cleaningWeeksList.includes(week.toString()) && weekday !== 'SUN') weekday = 'CLEANING WEEK';

  const dayHours = dayHoursData.find(x => x[0] === weekday);
  if (!dayHours) return total;

  if (weekday === 'SUN' && shift !== "" && shift !== "Kan Ikke") {
    return total + (parseFloat(dayHours[1]) || 0);
  }

  switch (shift) {
    case 'Kan':            return total + (parseFloat(dayHours[1]) + parseFloat(dayHours[2]) + parseFloat(dayHours[3]) || 0);
    case 'Åbne':           return total + (parseFloat(dayHours[1]) || 0);
    case 'Åbne & Mellem':  return total + (parseFloat(dayHours[1]) + parseFloat(dayHours[2]) || 0);
    case 'Mellem':         return total + (parseFloat(dayHours[2]) || 0);
    case 'Mellem & Lukke': return total + (parseFloat(dayHours[2]) + parseFloat(dayHours[3]) || 0);
    case 'Åbne & Lukke':   return total + (parseFloat(dayHours[1]) + parseFloat(dayHours[3]) || 0);
    case 'Lukke':          return total + (parseFloat(dayHours[3]) || 0);
    default:               return total;
  }
}

function getTotalHours() {
  try {
    const configSheet       = getConfigSheet();
    const dayHoursData      = configSheet.getRange('A9:D16').getValues();
    const cleaningWeeksList = configSheet.getRange(CLEANING_WEEK_START_CELL).getValue().toString().split(',');
    const reducer           = (acc, val) => weekDayHours(acc, val, dayHoursData, cleaningWeeksList);

    let days = getTemplateSheet()
      .getRange(4, 2, 100, 4)
      .getValues()
      .filter(a => a[0] !== '');

    const openDays  = days.map(v => { v[3] = 'Åbne';  return v; });
    const closeDays = days.map(v => { v[3] = 'Lukke'; return v; });

    const openHours    = openDays.reduce(reducer, 0);
    const openFridays  = openDays.filter(a => a[0] === 'FRI').reduce(reducer, 0) * 2;
    const closeHours   = closeDays.reduce(reducer, 0);
    const closeFridays = closeDays.filter(a => a[0] === 'FRI').reduce(reducer, 0) * 2;

    return openHours + closeHours + closeFridays + openFridays;
  } catch (error) {
    Logger.log(`Error in getTotalHours: ${error}`);
    return 0;
  }
}

function getHours() {
  try {
    const configSheet       = getConfigSheet();
    const dayHoursData      = configSheet.getRange('A9:D16').getValues();
    const cleaningWeeksList = configSheet.getRange(CLEANING_WEEK_START_CELL).getValue().toString().split(',');
    const reducer           = (acc, val) => weekDayHours(acc, val, dayHoursData, cleaningWeeksList);

    const answers = activeSheet()
      .getRange(4, 2, 100, 4)
      .getValues()
      .filter(a => a[3] !== '');

    return answers.reduce(reducer, 0);
  } catch (error) {
    Logger.log(`Error in getHours: ${error}`);
    return 0;
  }
}

function refreshTotalHours() {
  try {
    getConfigSheet().getRange('B18').setValue(getTotalHours());
  } catch (error) {
    Logger.log(`Error in refreshTotalHours: ${error}`);
  }
}

function getWeekNumberFromIndex(dayIndex) {
  return Math.floor(dayIndex / 6);
}

// ── CSP scheduler ────────────────────────────────────────────────────────────

class Volunteer {
  constructor(name, email) {
    this.name             = name;
    this.email            = email;
    this.prefs            = [];
    this.totalHours       = 0;
    this.assignedShifts   = [];
    this.applicableShifts = []; // pre-computed: shifts this volunteer can work
  }
}

class Shift {
  constructor(dayIndex, shiftType, requiredVolunteers) {
    this.dayIndex           = dayIndex;
    this.shiftType          = shiftType;
    this.requiredVolunteers = requiredVolunteers;
    this.possibleVolunteers = [];
    this.assignedVolunteers = [];
  }
}

// ── Entry point ───────────────────────────────────────────────────────────────

function fillSchedule() {
  try {
    const ctx = buildSchedulingData();
    buildShiftDomains(ctx.shifts, ctx.volunteers);

    // Pre-search summary (these run once, so logging is fine here)
    const unfilledShifts = ctx.shifts.filter(s => s.possibleVolunteers.length === 0);
    Logger.log(`Volunteers: ${ctx.volunteers.length} | Shifts: ${ctx.shifts.length} | Target: ${ctx.targetHours}h | Unfillable at start: ${unfilledShifts.length}`);

    // MCV: hardest-to-fill shifts first
    ctx.shifts.sort((a, b) => a.possibleVolunteers.length - b.possibleVolunteers.length);

    const state = {
      bestAssignment: { assignment: {}, numAssigned: 0 },
      startTime:      Date.now(),
      maxTime:        5 * 60 * 1000,
    };

    const assignment  = {};
    const complete    = backtrack(ctx, state, assignment);
    const elapsed     = ((Date.now() - state.startTime) / 1000).toFixed(1);
    const finalResult = complete ? assignment : state.bestAssignment.assignment;
    const numAssigned = complete ? ctx.shifts.length : state.bestAssignment.numAssigned;

    if (numAssigned > 0) {
      applySchedule(finalResult, ctx.weekDays);
      refreshHours();

      // Hour distribution summary
      const dist = ctx.volunteers
        .map(v => `${v.name}: ${v.totalHours}h`)
        .sort()
        .join('\n');
      const zeroHours = ctx.volunteers.filter(v => v.totalHours === 0).map(v => v.name);
      Logger.log(`Hour distribution after ${elapsed}s:\n${dist}`);

      const zeroMsg = zeroHours.length
        ? `\n\nFrivillige med 0 timer: ${zeroHours.join(', ')}`
        : '';

      SpreadsheetApp.getUi().alert(
        complete
          ? `Skema genereret ✓ (${elapsed}s)${zeroMsg}`
          : `Skema delvist udfyldt — ${numAssigned}/${ctx.shifts.length} vagter besat (${elapsed}s).${zeroMsg}`
      );
    } else {
      SpreadsheetApp.getUi().alert('Ingen gyldigt skema kunne genereres.');
    }
  } catch (error) {
    Logger.log(`Error in fillSchedule: ${error}`);
    SpreadsheetApp.getUi().alert(`En fejl opstod: ${error.message}`);
  }
}

// ── Data building ─────────────────────────────────────────────────────────────

function buildSchedulingData() {
  const scheduleSheet = getScheduleSheet();
  const configSheet   = getConfigSheet();

  const volunteerList = getVolunteerList();
  const volunteers    = volunteerList.map(vol => {
    const v = new Volunteer(vol[0], vol[5]);
    v.prefs = getUserPreferences(v.name);
    return v;
  });

  const dayHoursData = configSheet.getRange('A9:D16').getValues();
  const dayHoursMap  = dayHoursData.reduce((map, row) => {
    map[row[0]] = { open: row[1], mellem: row[2], close: row[3] };
    return map;
  }, {});

  const rawTarget   = configSheet.getRange(AVERAGE_HOURS_CELL).getValue();
  const targetHours = (typeof rawTarget === 'number' && rawTarget > 0) ? rawTarget : MAX_VOLUNTEER_HOURS;

  const weekDays = scheduleSheet.getRange('B4:B66').getValues().flat();

  const shifts = [];
  for (let dayIndex = 0; dayIndex < weekDays.length; dayIndex++) {
    const weekday = weekDays[dayIndex];
    if (!weekday) continue;
    const isFriday = weekday === 'FRI';
    const shiftTypes = isFriday ? ['Åbne', 'Mellem', 'Lukke'] : ['Åbne', 'Lukke'];
    shiftTypes.forEach(shiftType => {
      shifts.push(new Shift(dayIndex, shiftType, shiftType === 'Mellem' ? 2 : (isFriday ? 2 : 1)));
    });
  }

  return { volunteers, shifts, dayHoursMap, weekDays, targetHours };
}

function buildShiftDomains(shifts, volunteers) {
  shifts.forEach(shift => {
    volunteers.forEach(volunteer => {
      const pref = volunteer.prefs[shift.dayIndex]?.[0];
      if (!pref || pref === 'Kan Ikke') return;

      let canWork = pref === 'Kan' || pref === shift.shiftType;
      if (!canWork) {
        if      (pref === 'Åbne & Mellem')  canWork = shift.shiftType === 'Åbne'  || shift.shiftType === 'Mellem';
        else if (pref === 'Mellem & Lukke') canWork = shift.shiftType === 'Mellem' || shift.shiftType === 'Lukke';
        else if (pref === 'Åbne & Lukke')   canWork = shift.shiftType === 'Åbne'  || shift.shiftType === 'Lukke';
      }
      if (canWork) shift.possibleVolunteers.push(volunteer);
    });
  });

  // Reverse index: each volunteer knows which shifts they can work.
  // Used by LCV ordering to avoid O(all shifts) scans inside the hot loop.
  volunteers.forEach(volunteer => {
    volunteer.applicableShifts = shifts.filter(s => s.possibleVolunteers.includes(volunteer));
  });
}

// ── Backtracking search ───────────────────────────────────────────────────────
//
// ctx   — static data:   { shifts, volunteers, dayHoursMap, weekDays, targetHours }
// state — mutable state: { bestAssignment, startTime, maxTime }
//
// Logger.log is intentionally absent from this function and all functions it calls.
// In Apps Script each log call is a synchronous API round-trip (~10–50ms); logging
// inside a tight recursive loop is the primary reason the search can take minutes.

function backtrack(ctx, state, assignment, depth = 0) {
  if (Date.now() - state.startTime > state.maxTime) {
    Logger.log('Time limit reached.');
    return false;
  }

  const numAssigned = Object.values(assignment).filter(arr => arr.length > 0).length;

  if (numAssigned > state.bestAssignment.numAssigned) {
    state.bestAssignment = { assignment: Object.assign({}, assignment), numAssigned };
  }

  if (numAssigned === ctx.shifts.length) return true;

  const shift = ctx.shifts.find(s => !assignment[s.dayIndex + '_' + s.shiftType]);
  if (!shift) return false;

  if (shift.possibleVolunteers.length === 0) {
    assignment[shift.dayIndex + '_' + shift.shiftType] = [];
    return backtrack(ctx, state, assignment, depth + 1);
  }

  const orderedVolunteers = orderVolunteersByFairnessAndLCV(shift, ctx);
  let attemptedAny = false;

  for (const volunteer of orderedVolunteers) {
    if (isConsistent(ctx, volunteer, shift)) {
      assign(ctx, volunteer, shift, assignment);
      if (backtrack(ctx, state, assignment, depth + 1)) return true;
      unassign(ctx, volunteer, shift, assignment);
      attemptedAny = true;
    }
  }

  if (!attemptedAny) {
    assignment[shift.dayIndex + '_' + shift.shiftType] = [];
    return backtrack(ctx, state, assignment, depth + 1);
  }

  return false;
}

// ── Constraint checking ───────────────────────────────────────────────────────

function isConsistent(ctx, volunteer, shift) {
  if (shift.assignedVolunteers.length >= shift.requiredVolunteers) return false;

  const shiftHours = getShiftHours(ctx.dayHoursMap, ctx.weekDays[shift.dayIndex], shift.shiftType);
  if (volunteer.totalHours + shiftHours > MAX_VOLUNTEER_HOURS) return false;

  const currentWeek = getWeekNumberFromIndex(shift.dayIndex);
  for (const s of volunteer.assignedShifts) {
    if (s.dayIndex === shift.dayIndex)                           return false; // same day
    if (Math.abs(s.dayIndex - shift.dayIndex) <= 1)             return false; // consecutive days
    if (getWeekNumberFromIndex(s.dayIndex) === currentWeek)     return false; // same week
  }

  return true;
}

// ── Assignment / unassignment ─────────────────────────────────────────────────

function assign(ctx, volunteer, shift, assignment) {
  shift.assignedVolunteers.push(volunteer);
  volunteer.assignedShifts.push(shift);
  volunteer.totalHours += getShiftHours(ctx.dayHoursMap, ctx.weekDays[shift.dayIndex], shift.shiftType);

  if (shift.assignedVolunteers.length === shift.requiredVolunteers) {
    assignment[shift.dayIndex + '_' + shift.shiftType] = shift.assignedVolunteers.map(v => v.name);
  }
}

function unassign(ctx, volunteer, shift, assignment) {
  shift.assignedVolunteers = shift.assignedVolunteers.filter(v => v !== volunteer);
  volunteer.assignedShifts = volunteer.assignedShifts.filter(s => s !== shift);
  volunteer.totalHours    -= getShiftHours(ctx.dayHoursMap, ctx.weekDays[shift.dayIndex], shift.shiftType);

  if (shift.assignedVolunteers.length < shift.requiredVolunteers) {
    delete assignment[shift.dayIndex + '_' + shift.shiftType];
  } else {
    assignment[shift.dayIndex + '_' + shift.shiftType] = shift.assignedVolunteers.map(v => v.name);
  }
}

// ── Volunteer ordering ────────────────────────────────────────────────────────

// Primary sort: volunteers furthest below their target hours go first (fairness).
// This directly prevents the 0-shift problem — under-scheduled volunteers are
// always tried before those who already have enough hours.
//
// Tiebreak: LCV — volunteers whose assignment leaves more options for other
// shifts go first (fewer constraints removed from remaining search space).
// Uses the pre-computed applicableShifts list to avoid O(all shifts) scans.
function orderVolunteersByFairnessAndLCV(shift, ctx) {
  return [...shift.possibleVolunteers].sort((a, b) => {
    const aDeficit = ctx.targetHours - a.totalHours;
    const bDeficit = ctx.targetHours - b.totalHours;
    if (Math.abs(aDeficit - bDeficit) > 0.01) return bDeficit - aDeficit;

    // LCV tiebreak
    const aOptions = a.applicableShifts.filter(s => s !== shift).length;
    const bOptions = b.applicableShifts.filter(s => s !== shift).length;
    return bOptions - aOptions;
  });
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function getShiftHours(dayHoursMap, weekday, shiftType) {
  const h = dayHoursMap[weekday];
  if (!h) return 0;
  if (shiftType === 'Åbne')   return h.open;
  if (shiftType === 'Mellem') return h.mellem;
  return h.close;
}

// ── Schedule application ─────────────────────────────────────────────────────

function applySchedule(assignment, weekDays) {
  const scheduleSheet  = getScheduleSheet();
  const scheduleUpdates = [];

  for (const key in assignment) {
    const [dayIndexStr, shiftType] = key.split('_');
    const dayIndex       = parseInt(dayIndexStr, 10);
    const volunteerNames = assignment[key];

    volunteerNames.forEach((name, index) => {
      let cellAddress;
      switch (shiftType) {
        case 'Åbne':  cellAddress = `${index === 0 ? 'E' : 'F'}${SCHEDULE_START_ROW + dayIndex}`; break;
        case 'Mellem': cellAddress = `${index === 0 ? 'G' : 'H'}${SCHEDULE_START_ROW + dayIndex}`; break;
        case 'Lukke': cellAddress = `${'IJK'[index]}${SCHEDULE_START_ROW + dayIndex}`; break;
      }
      if (cellAddress) scheduleUpdates.push({ cell: cellAddress, value: name });
    });
  }

  scheduleSheet.getRange(`E${SCHEDULE_START_ROW}:K${SCHEDULE_END_ROW}`).clearContent();
  scheduleUpdates.forEach(({ cell, value }) => scheduleSheet.getRange(cell).setValue(value));
}
