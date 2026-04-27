function fetchClassSchedule(schoolUserId) {
  const url = `https://personligtskema.ku.dk/ical.asp?objectclass=student&id=${schoolUserId}`;
  try {
    const response = UrlFetchApp.fetch(url);
    const icsContent = response.getContentText();

    // Parse the iCalendar data
    const events = parseICS(icsContent);

    // Process events to extract busy times
    const busyTimes = events.map(event => {
      return {
        start: new Date(event.start),
        end: new Date(event.end)
      };
    });

    return busyTimes;
  } catch (error) {
    Logger.log(`Error fetching class schedule: ${error}`);
    return [];
  }
}

function parseICS(icsContent) {
  const events = [];
  // Unfold RFC 5545 folded lines (CRLF or LF followed by whitespace)
  const unfolded = icsContent.replace(/\r?\n[ \t]/g, '');
  const lines = unfolded.split(/\r?\n/);
  let currentEvent = null;

  lines.forEach(line => {
    if (line.startsWith('BEGIN:VEVENT')) {
      currentEvent = {};
    } else if (line.startsWith('END:VEVENT')) {
      if (currentEvent && currentEvent.start && currentEvent.end) {
        events.push(currentEvent);
      }
      currentEvent = null;
    } else if (currentEvent) {
      const colonIdx = line.indexOf(':');
      if (colonIdx === -1) return;
      const rawKey = line.slice(0, colonIdx);
      const value = line.slice(colonIdx + 1).trim();
      // Strip parameters (e.g. DTSTART;TZID=Europe/Copenhagen → DTSTART)
      const key = rawKey.split(';')[0];
      if (key === 'DTSTART') currentEvent.start = value;
      if (key === 'DTEND') currentEvent.end = value;
    }
  });

  return events;
}

function isDateInBusyTimes(date, busyTimes) {
  return busyTimes.some(busy => date >= busy.start && date < busy.end);
}
