// calendar.gs

// ── Volunteer shift collection ───────────────────────────────────────────────

function getVolunteerShifts() {
  const scheduleData    = getScheduleSheet().getRange('A4:J').getValues();
  const volunteerShifts = {};

  scheduleData.forEach(row => {
    const date = row[0];
    if (!date || !(date instanceof Date) || isNaN(date.getTime())) return;

    const weekday  = row[1];
    const note     = row[3];
    const openers  = row.slice(4, 7).filter(name => name);
    const closers  = row.slice(7, 10).filter(name => name);

    [...openers, ...closers].forEach(volunteerName => {
      const info = getVolunteerInfoByName(volunteerName);
      if (!info?.email) return;

      if (!volunteerShifts[info.email]) volunteerShifts[info.email] = [];
      volunteerShifts[info.email].push({
        date,
        weekday,
        shiftType: openers.includes(volunteerName) ? 'Åbner' : 'Lukker',
        note,
      });
    });
  });

  return volunteerShifts;
}

function getVolunteerInfoByName(name) {
  const volunteer = getVolunteerList().find(vol => vol[0] === name);
  return volunteer ? { name: volunteer[0], email: volunteer[5] } : null;
}

// ── Calendar sync ────────────────────────────────────────────────────────────

function getCalendarId() {
  const id = getConfigSheet().getRange(CALENDAR_ID_CELL).getValue().toString().trim();
  if (!id) throw new Error(`Calendar ID not configured. Add it to cell ${CALENDAR_ID_CELL} in the Configuration sheet.`);
  return id;
}

/**
 * Menu entry point — syncs shifts and shows a summary alert.
 */
function syncShiftsToSharedCalendar() {
  try {
    const count = syncShiftsToSharedCalendarCore();
    SpreadsheetApp.getUi().alert(`${count} vagter synkroniseret til kalenderen.`);
  } catch (error) {
    Logger.log(`Error in syncShiftsToSharedCalendar: ${error}`);
    SpreadsheetApp.getUi().alert(`An error occurred: ${error.message}`);
  }
}

/**
 * Core logic — syncs shifts and returns the number of events created.
 * Safe to call from the sidebar (no UI alerts).
 * @return {number} Number of calendar events created.
 */
function syncShiftsToSharedCalendarCore() {
  const volunteerShifts = getVolunteerShifts();
  const calendar        = CalendarApp.getCalendarById(getCalendarId());

  if (!calendar) throw new Error('Shared calendar not found. Check CALENDAR_ID in calendar.gs.');

  clearCalendarEvents(calendar);

  let created = 0;

  for (const [email, shifts] of Object.entries(volunteerShifts)) {
    shifts.forEach(shift => {
      const eventTitle      = `C? Vagt - ${shift.shiftType}`;
      const eventTimes      = getShiftEventTimes(shift.date, shift.shiftType, shift.weekday);
      const eventDescription = `${shift.note || ''}\n[AutoSkema]`.trim();

      const existingEvents = calendar.getEvents(eventTimes.startTime, eventTimes.endTime, { search: eventTitle });
      const alreadyInvited = existingEvents.some(e => e.getGuestList().some(g => g.getEmail() === email));

      if (!alreadyInvited) {
        calendar.createEvent(eventTitle, eventTimes.startTime, eventTimes.endTime, {
          description: eventDescription,
          guests:      email,
          sendInvites: true,
        });
        Logger.log(`Created event for ${email}: ${eventTitle} on ${shift.date}`);
        created++;
      } else {
        Logger.log(`Event already exists for ${email}: ${eventTitle} on ${shift.date}`);
      }
    });
  }

  return created;
}

// ── Shift time calculation ───────────────────────────────────────────────────

function getShiftEventTimes(date, shiftType, weekday) {
  const startTime = new Date(date);
  const endTime   = new Date(date);

  switch (weekday) {
    case 'WED':
      startTime.setHours(shiftType === 'Åbner' ? 12 : 17, 0, 0, 0);
      endTime.setHours(shiftType === 'Åbner' ? 16 : 20, 0, 0, 0);
      break;
    case 'THU':
      startTime.setHours(shiftType === 'Åbner' ? 12 : 17, 0, 0, 0);
      endTime.setHours(shiftType === 'Åbner' ? 16 : 22, 0, 0, 0);
      break;
    case 'FRI':
      if (shiftType === 'Åbner') {
        startTime.setHours(12, 0, 0, 0);
        endTime.setHours(20, 0, 0, 0);
      } else {
        // Friday closing shift ends at 3am Saturday
        startTime.setHours(20, 0, 0, 0);
        endTime.setDate(endTime.getDate() + 1);
        endTime.setHours(3, 0, 0, 0);
      }
      break;
    case 'SUN':
      startTime.setHours(12, 0, 0, 0);
      endTime.setHours(16, 0, 0, 0);
      break;
    default:
      startTime.setHours(shiftType === 'Åbner' ? 12 : 16, 0, 0, 0);
      endTime.setHours(shiftType === 'Åbner' ? 16 : 20, 0, 0, 0);
      break;
  }

  return { startTime, endTime };
}

// Only deletes events created by this script (identified by the [AutoSkema] tag).
function clearCalendarEvents(calendar) {
  const startDate = new Date();
  const endDate   = new Date();
  endDate.setDate(endDate.getDate() + 90);

  let deleted = 0;
  calendar.getEvents(startDate, endDate).forEach(event => {
    if (event.getDescription().includes('[AutoSkema]')) {
      event.deleteEvent();
      deleted++;
    }
  });

  Logger.log(`Cleared ${deleted} AutoSkema events from ${startDate.toDateString()} to ${endDate.toDateString()}`);
}

// ── KU Skema preview ─────────────────────────────────────────────────────────

const previewSkema = () => {
  const html = `
    <html>
      <body>
        <div style="font-family: Arial, sans-serif; font-size: 14px; margin: 20px;">
          <label for="userid">Indtast KU nummerplade:</label>
          <input type="text" id="userid" style="width: 100%; padding: 8px; margin: 10px 0;"
                 placeholder="Indtast KU-ID">
          <button onclick="submit()"
                  style="padding: 10px 20px; background-color: #2e7d32; color: white; border: none; cursor: pointer; border-radius: 4px;">
            Vis skema
          </button>
        </div>
        <script>
          function submit() {
            const userId = document.getElementById('userid').value.trim();
            if (!userId) { alert('Feltet kan ikke være tomt!'); return; }
            google.script.run.withSuccessHandler(closeDialog).showPreview(userId);
          }
          function closeDialog() { google.script.host.close(); }
        <\/script>
      </body>
    </html>
  `;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(400).setHeight(200),
    'Indtast KU-ID'
  );
};

const showPreview = (userId) => {
  const baseUrl = 'https://gra0007.github.io/modern-cal-embed/iframe.html?ical=https%3A%2F%2Fapi.codetabs.com%2Fv1%2Fproxy%2F%3Fquest%3Dhttps%253A%252F%252Fpersonligtskema.ku.dk%252Fical.asp%253Fobjectclass%253Dstudent%2526id%253DUSER_ID&title=1&nav=1&date=1&view=1&details=0&monstart=1&dview=1&color=%231A73E8&colorbg=%23FFFFFF&colortxt=%23000000&colorsecondarytxt=%23FFFFFF';
  const url     = baseUrl.replace('USER_ID', encodeURIComponent(userId));

  const iframeHtml = `
    <html>
      <head>
        <style>
          html, body { margin: 0; padding: 0; height: 100%; overflow: hidden; }
          iframe { width: 100%; height: 100%; border: none; }
        </style>
      </head>
      <body><iframe src="${url}"></iframe></body>
    </html>
  `;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(iframeHtml).setWidth(800).setHeight(600),
    'MitSkema'
  );
};
