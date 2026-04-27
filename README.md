# AutoSkema

Google Apps Script system for automating volunteer shift scheduling at **Caféen?**. Manages shift templates, personal volunteer pages, automatic schedule generation, Google Calendar sync, and email reminders — all from a sidebar panel inside Google Sheets.

---

## Table of contents

1. [Overview](#overview)
2. [File structure](#file-structure)
3. [Setup](#setup)
4. [Local development with clasp](#local-development-with-clasp)
5. [Configuration sheet](#configuration-sheet)
6. [Using the admin sidebar](#using-the-admin-sidebar)
7. [How the scheduler works](#how-the-scheduler-works)
8. [Personal pages](#personal-pages)
9. [Calendar sync](#calendar-sync)
10. [Email reminders](#email-reminders)
11. [KU schedule preview](#ku-schedule-preview)
12. [Troubleshooting](#troubleshooting)

---

## Overview

The spreadsheet has three permanent sheets:

| Sheet | Purpose |
|---|---|
| **Configuration** | All settings: year, week range, shift hours, admin emails |
| **Endeligt Skema** | The live schedule filled in by admins; volunteer list lives here too |
| **Template** | Blank per-volunteer preference sheet; copied when personal pages are created |

When the season starts, an admin:
1. Fills in the Configuration sheet
2. Generates a fresh schedule template
3. Creates personal pages (one per volunteer)
4. Volunteers fill in their availability on their own page
5. Admin runs the auto-scheduler or fills the schedule manually
6. Admin syncs shifts to Google Calendar and sends reminder emails

---

## File structure

```
AutoSkema/
├── main.gs        Entry points for menu actions and spreadsheet events
├── schedule.gs    Template generation, hours calculation, CSP auto-scheduler
├── calendar.gs    Google Calendar sync and KU schedule preview
├── email.gs       Reminder email sending
├── ics.gs         iCalendar (.ics) parser for university class schedules
├── utils.gs       Constants, sheet getters, shared utilities
├── sidebar.gs     Admin sidebar server functions and access control
└── sidebar.html   Admin sidebar UI
```

### utils.gs — constants and shared helpers

All magic numbers and cell references live here. If the spreadsheet layout changes, this is the only file that needs updating.

```
VOLUNTEER_LIST_START_ROW  Row where the volunteer list begins (default 5)
VOLUNTEER_LIST_COL        Column N (14) — where volunteer names are stored
SCHEDULE_START_ROW        Row 4 — top of the schedule grid
SCHEDULE_END_ROW          Row 71 — bottom of the schedule grid
MAX_VOLUNTEER_HOURS       Hard cap on hours any one volunteer can be assigned
AVERAGE_HOURS_CELL        B21 — target hours per volunteer for the auto-scheduler
ADMIN_EMAILS_CELL         B3  — comma-separated list of admin email addresses
```

Sheet references are lazy getters (`getScheduleSheet()`, `getConfigSheet()`, `getTemplateSheet()`)
---

## Setup

### 1. Copy the spreadsheet

The script is bound to the Google Spreadsheet. Open the sheet, go to **Extensions → Apps Script** to view the code.

### 2. Enable required APIs

In Apps Script: **Services → Add a service** and enable:
- **Admin SDK Directory API** — used to verify volunteer email addresses when protecting personal pages

### 3. Set the Calendar ID

In `calendar.gs`, update the constant with your shared Google Calendar ID:

```js
const CALENDAR_ID = 'your-calendar-id@group.calendar.google.com';
```

Find the Calendar ID in Google Calendar under **Settings → [calendar name] → Integrate calendar**.

### 4. Configure admin access

In the **Configuration** sheet, put the label `Admin e-mails` in cell **A3** and a comma-separated list of admin email addresses in **B3**:

```
freja@cafeen.org,kr0ll@cafeen.org,charlotte@cafeen.org
```

Only these users can open the admin sidebar or call any sidebar server function.

### 5. Run `onOpen` once

From the Apps Script editor, run the `onOpen` function manually once to register the custom menu. After that it runs automatically every time the spreadsheet is opened.

---

## Local development with clasp

[clasp](https://github.com/google/clasp) (Command Line Apps Script Projects) is Google's official CLI for Apps Script. It lets you edit `.gs` files locally in any editor, version-control them with Git, and push changes to the live script without touching the browser-based editor.

### Install clasp

```bash
npm install -g @google/clasp
```

Requires Node.js 12+.

### Log in

```bash
clasp login
```

This opens a browser window to authorise clasp with your Google account. Credentials are stored in `~/.clasprc.json`.

### Link the project

The Apps Script project is bound to the spreadsheet. To get the **Script ID**:
1. Open the spreadsheet
2. Go to **Extensions → Apps Script**
3. Click **Project Settings** (gear icon)
4. Copy the **Script ID**

Create a `.clasp.json` file in the root of this repository:

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "."
}
```

> `.clasp.json` contains your personal script ID. It is already listed in `.gitignore` — do not commit it.

### Daily workflow

| Task | Command |
|---|---|
| Pull latest code from Apps Script to local | `clasp pull` |
| Push local changes to Apps Script | `clasp push` |
| Open the script editor in the browser | `clasp open` |
| Watch for local changes and auto-push | `clasp push --watch` |

**Typical session:**

```bash
# Make sure you have the latest before editing
clasp pull

# ... edit .gs / .html files locally ...

# Push when ready
clasp push
```

### What gets pushed

clasp pushes every file in `rootDir` that matches a supported extension. For this project that means all `.gs` files and `sidebar.html`. The `README.md` and `.clasp.json` are ignored by Apps Script.

### appsscript.json

Apps Script stores project metadata (timezone, OAuth scopes, enabled APIs) in `appsscript.json`. After running `clasp pull` for the first time you will see this file appear. Commit it alongside the source files — it defines which Google APIs the script is allowed to call and should be kept in sync with the live project.

A minimal example for this project:

```json
{
  "timeZone": "Europe/Copenhagen",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "AdminDirectory",
        "serviceId": "admin",
        "version": "directory_v1"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

---

## Configuration sheet

| Cell | Contents |
|---|---|
| A3 / B3 | Label / Admin email addresses (comma-separated) — controls sidebar access and personal page editor permissions |
| B4 | Year (e.g. `2025`) |
| B5 | Open–close week range (e.g. `36-2`) |
| B6 | Cleaning week(s) — single week or two comma-separated weeks |
| A7 / B7 | Label / Shared Google Calendar ID for shift sync |
| A9:D16 | Shift hours by weekday — columns: Weekday, Åbne hours, Mellem hours, Lukke hours |
| B18 | Total hours in the schedule (auto-calculated) |
| B21 | **Target hours per volunteer** — used by the auto-scheduler to distribute shifts fairly |

### Shift hours table (A9:D16)

Each row defines how many hours each shift type is worth on a given weekday:

| A (Weekday) | B (Åbne) | C (Mellem) | D (Lukke) |
|---|---|---|---|
| MON | 4 | — | 4 |
| TUE | 4 | — | 4 |
| WED | 4 | — | 3 |
| THU | 4 | — | 5 |
| FRI | 8 | 4 | 7 |
| SUN | 4 | — | — |
| CLEANING WEEK | 4 | — | 4 |

---

## Using the admin sidebar

Open it from **Caféen? → Åbn Admin Panel** in the menu bar.

![Sidebar sections: Skabelon, Frivillige, Skema, Kalender, E-mails, Eksport]

Access is restricted to the emails listed in B3. Any other user sees an "Adgang nægtet" alert and the sidebar never opens. Each server function also re-checks access independently, so the restriction holds even if someone calls `google.script.run` directly from the browser console.

### Sections

**Skabelon**
- **Generer skabelon** — Clears the schedule and template sheets and rebuilds them from the Configuration sheet (week range, cleaning weeks, year). Shows an inline confirmation before running.

**Frivillige**
- **Opret personlige sider** — Creates one sheet per volunteer (named after the volunteer) by copying the Template sheet. Protects each sheet so only that volunteer and the admins can edit it. Requires confirmation.
- **Slet personlige sider** — Permanently deletes all personal pages. Requires confirmation.

**Skema**
- **Opdater dropdowns** — Rebuilds the dropdown menus in the schedule grid so each cell only shows volunteers who are available for that specific shift on that day.
- **Opdater timer** — Recalculates the total hours and days worked for every volunteer and writes the results back to the schedule sheet.
- **Auto-udfyld skema** — Runs the CSP auto-scheduler (see [How the scheduler works](#how-the-scheduler-works)). Can take up to a few seconds; results appear in a popup with a per-volunteer hour summary.

**Kalender**
- **Synkroniser kalender** — Creates Google Calendar events for every assigned shift and sends email invitations to the volunteers. Only deletes events previously created by AutoSkema (tagged `[AutoSkema]` in the event description).

**E-mails**
- **Send påmindelser** — Sends a personalised HTML reminder email to every volunteer whose schedule status is `Nej`.

**Eksport**
- **Eksportér til Wiki** — Saves the schedule as a Wiki-formatted `.txt` file in Google Drive.

### Button states

- **Loading spinner** — shown while a server function is running; all buttons are disabled
- **Green status bar** — shown on success (auto-dismisses after 6 seconds)
- **Red status bar** — shown on error with the error message
- **Inline confirmation panel** — shown for destructive actions before the server call is made

---

## How the scheduler works

`fillSchedule()` in `schedule.gs` uses a **Constraint Satisfaction Problem (CSP)** solver with backtracking search and two heuristics.

### Constraints

| Constraint | Rule |
|---|---|
| Availability | Volunteer must have marked the day as anything other than `Kan Ikke` |
| Shift match | Volunteer's preference must cover the shift type (Åbne / Mellem / Lukke), including compound preferences like `Åbne & Lukke` |
| Max hours | A volunteer's total assigned hours cannot exceed `MAX_VOLUNTEER_HOURS` |
| One shift per week | A volunteer can only work one shift per week |
| No consecutive days | A volunteer cannot work on two rows that are adjacent in the schedule (prevents back-to-back days across week boundaries) |
| Shift capacity | Each shift slot has a required number of volunteers (1 for most, 2 for Friday Mellem) |

### Heuristics

**MCV — Most Constrained Variable**
Shifts are sorted before search begins: shifts with the fewest possible volunteers are assigned first. This catches impossible or near-impossible slots early rather than discovering them deep in the search tree.

**Fairness + LCV — Volunteer ordering**
For each shift, volunteers are ordered by two keys:

1. **Fairness (primary)** — volunteers furthest below their target hours (`B21`) go first. This directly prevents some volunteers ending up with 0 shifts: under-scheduled volunteers are always tried before those who already have enough hours.
2. **LCV tiebreak** — among volunteers with similar hours, prefer the one whose assignment removes the fewest options from remaining shifts (Least Constraining Value). Uses a pre-computed `applicableShifts` list per volunteer to keep this fast.

### Result

If a complete schedule is found, it is applied directly. If the time limit is reached (5 minutes), the best partial schedule found so far is applied. The result popup shows:
- Time taken
- Any volunteers who still have 0 hours after scheduling

---

## Personal pages

Each volunteer has their own sheet (named after them) copied from the **Template** sheet. The sheet is protected so only that volunteer and the admins listed in B3 can edit it.

Column E (rows 4–71) holds the volunteer's availability preferences for each row in the schedule:

| Value | Meaning |
|---|---|
| `Kan` | Available for any shift |
| `Åbne` | Opening shift only |
| `Lukke` | Closing shift only |
| `Mellem` | Middle shift only (Fridays) |
| `Åbne & Lukke` | Either opening or closing (Fridays) |
| `Åbne & Mellem` | Either opening or middle (Fridays) |
| `Mellem & Lukke` | Either middle or closing (Fridays) |
| `Kan Ikke` | Not available |

Fridays have an extended preference list because they have three shift types (Åbne, Mellem, Lukke).

### Highlighting available shifts

When an admin selects a volunteer name in cell **O1** of the schedule sheet, the cells that volunteer is available for are highlighted green. Clearing O1 removes the highlight.

---

## Calendar sync

`syncShiftsToSharedCalendar()` in `calendar.gs`:

1. Reads all assigned volunteer names from the schedule grid
2. Looks up each volunteer's email address from the volunteer list
3. Creates a Google Calendar event for each shift with the correct start/end time
4. Invites the volunteer as a guest (sends an email invitation)
5. Tags each created event with `[AutoSkema]` in the description

On re-sync, only events tagged `[AutoSkema]` are deleted first — manually added events on the same calendar are left untouched.

### Shift times

| Day | Åbner | Lukker |
|---|---|---|
| Mon / Tue | 12:00–16:00 | 16:00–20:00 |
| Wed | 12:00–16:00 | 17:00–20:00 |
| Thu | 12:00–16:00 | 17:00–22:00 |
| Fri | 12:00–20:00 | 20:00–03:00 (Saturday) |
| Sun | 12:00–16:00 | 12:00–16:00 |

---

## Email reminders

`sendReminderEmails()` in `email.gs` sends a personalised HTML email to every volunteer whose status column (index 6 in the volunteer list) is `Nej`.

The email contains a direct link to the volunteer's personal sheet and is sent from the account running the script, displayed as **Caféen?**.

---

## KU schedule preview

**KU Skema → Vis MitSkema** opens a dialog where a volunteer can enter their KU student ID. Their university timetable is then displayed in an embedded calendar view (fetched via `personligtskema.ku.dk`).

`ics.gs` also provides `fetchClassSchedule(userId)` and `isDateInBusyTimes(date, busyTimes)` which can be used to check whether a volunteer has a class clash on a given date — useful for future integration with the availability system.

---

## Troubleshooting

**"Required sheet X not found"**
A sheet has been renamed or deleted. Check that the three sheets are named exactly `Configuration`, `Endeligt Skema`, and `Template`.

**"Adgang nægtet" when opening the sidebar**
Your email is not in the admin list in cell B3 of the Configuration sheet.

**"Ingen admin-e-mails konfigureret"**
Cell B3 in the Configuration sheet is empty. Add at least one email address.

**Calendar sync creates duplicate events**
Events are de-duplicated by checking existing events in the same time window with the same title. If you see duplicates, check that the calendar ID in `calendar.gs` is correct and that previous events were tagged with `[AutoSkema]`.

**Auto-scheduler gives some volunteers 0 hours**
- Check that those volunteers have filled in their personal page — `Kan Ikke` on every row means the scheduler has nowhere to place them.
- Verify that `B21` in the Configuration sheet contains a reasonable target (e.g. `10`). If the cell is empty the scheduler falls back to `MAX_VOLUNTEER_HOURS` as the target, which may be too high to drive the fairness ordering effectively.

**Dates in the Wiki export are off by one**
The export uses the spreadsheet's own timezone (`getSpreadsheetTimeZone()`). If dates still look wrong, verify the spreadsheet timezone in **File → Settings → Time zone** matches Denmark (Europe/Copenhagen).
