// email.gs

/**
 * Menu entry point — runs sendReminderEmailsCore() and shows a summary alert.
 */
function sendReminderEmails() {
  try {
    const count = sendReminderEmailsCore();
    SpreadsheetApp.getUi().alert(`Påmindelser sendt til ${count} frivillige der ikke har udfyldt skemaet.`);
  } catch (error) {
    Logger.log(`Error in sendReminderEmails: ${error}`);
    SpreadsheetApp.getUi().alert(`An error occurred: ${error.message}`);
  }
}

/**
 * Core logic — sends reminder emails and returns the number of emails sent.
 * Safe to call from the sidebar (no UI alerts).
 * @return {number} Number of reminder emails sent.
 */
function sendReminderEmailsCore() {
  const volunteerList        = getVolunteerList();
  const incompleteVolunteers = volunteerList.filter(v => v[6] === 'Nej');

  incompleteVolunteers.forEach(v => {
    Logger.log(`Sending reminder to ${v[0]} (${v[5]})`);
    sendReminderEmail(v[0], v[5]);
  });

  Logger.log(`Reminder emails sent to ${incompleteVolunteers.length} volunteers.`);
  return incompleteVolunteers.length;
}

/**
 * Sends a personalised HTML reminder email to a single volunteer.
 * Logs and returns if the volunteer's sheet is missing — does not throw.
 * @param {string} name  - Volunteer name (must match their sheet name).
 * @param {string} email - Volunteer email address.
 */
function sendReminderEmail(name, email) {
  try {
    const spreadsheet    = SpreadsheetApp.getActiveSpreadsheet();
    const volunteerSheet = spreadsheet.getSheetByName(name);
    if (!volunteerSheet) {
      Logger.log(`Sheet not found for volunteer: ${name}, skipping.`);
      return;
    }

    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/edit#gid=${volunteerSheet.getSheetId()}`;

    const htmlMessage = `
      <div style="font-family: Arial, sans-serif; color: #333;">
        <div style="text-align: center;">
          <img src="https://raw.githubusercontent.com/cafeen/CafeLogo/refs/heads/master/cafeen_logo.svg"
               alt="Caféen? logo" style="max-width: 200px; height: auto;"/>
        </div>
        <p>Hej ${name},</p>
        <p>
          Dette er en venlig påmindelse om at udfylde dine vagtpræferencer for den kommende blok.
          Vi har bemærket, at dit skema enten ikke er udfyldt eller kun er delvist udfyldt.
        </p>
        <p>Brug linket nedenfor for at gå til vagtskemaet:</p>
        <p style="text-align: center;">
          <a href="${spreadsheetUrl}"
             style="background-color: #2e7d32; color: #fff; padding: 10px 20px;
                    text-decoration: none; border-radius: 5px; font-weight: bold;">
            Udfyld vagtskema
          </a>
        </p>
        <p>
          Svar endelig tilbage på denne mail hvis du har problemer med at udfylde vagtskemaet.
          Du er selvfølgelig også velkommen til at skrive på Messenger eller Discord.
        </p>
        <p>Venlig hilsen,<br/>Vagtplanlægnings-ansvarlige</p>
      </div>
    `;

    MailApp.sendEmail({
      to:       email,
      subject:  'C? Reminder: Husk at udfylde vagtskema!',
      htmlBody: htmlMessage,
      name:     'Caféen?',
    });

    Logger.log(`Reminder email sent to ${name} at ${email}.`);
  } catch (error) {
    Logger.log(`Error sending email to ${name} at ${email}: ${error}`);
  }
}
