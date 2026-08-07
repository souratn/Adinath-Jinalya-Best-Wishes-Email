function sendDateBasedEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = 2;
  const numRows = sheet.getLastRow() - 1;

  if (numRows < 1) return;

  // Read 7 columns (A to G: Name, Email, Date, Subject, Status, TimeZone, LeadFlag)
  const dataRange = sheet.getRange(startRow, 1, numRows, 7);
  const data = dataRange.getValues();
  const currentYear = new Date().getFullYear();

  // Load the image from Google Drive using its File ID
  const imageFileId = "1ytp3XYrGVLD6zhgs5SJpD64F-mhTi-2t";
  const imageBlob = DriveApp.getFileById(imageFileId).getBlob().setName("headerImage");

  // Map user inputs to IANA Timezone Identifiers
  const tzMapping = {
    "IST": "Asia/Kolkata",
    "EST": "America/New_York",
    "EDT": "America/New_York",
    "CST": "America/Chicago",
    "PST": "America/Los_Angeles"
  };

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const name = row[0];
    const email = row[1];
    const rawDate = row[2];
    const subject = row[3];
    const status = row[4];
    const rawTimeZone = String(row[5]).trim().toUpperCase();
    const leadFlag = row[6];

    if (!email || !rawDate || status === "Sent") continue;

    const targetTimeZone = tzMapping[rawTimeZone] || rawTimeZone || "Asia/Kolkata";

    // Check if it is 9 AM in target timezone
    const nowInTargetTz = new Date();
    const currentHourInTargetTz = Number(Utilities.formatDate(nowInTargetTz, targetTimeZone, "HH"));

    if (currentHourInTargetTz !== 8) {
      continue;
    }

    // Safely extract Month and Day in target timezone
    const eventDateObj = new Date(rawDate);
    const todayMonthDay = Utilities.formatDate(nowInTargetTz, targetTimeZone, "MM-dd");
    const rowMonthDay = Utilities.formatDate(eventDateObj, targetTimeZone, "MM-dd");

    if (rowMonthDay === todayMonthDay) {
      // Parse Month and Day reliably using target timezone strings to prevent UTC offset shift
      const monthStr = Utilities.formatDate(eventDateObj, targetTimeZone, "MM");
      const dayStr = Utilities.formatDate(eventDateObj, targetTimeZone, "dd");

      // Construct target date at 12:00 PM to avoid daylight saving time boundary shifts
      const currentYearEventDate = new Date(currentYear, Number(monthStr) - 1, Number(dayStr), 12, 0, 0);

      // Calculate days until next Sunday (0 = Sun, 1 = Mon, ..., 6 = Sat)
      const dayOfWeek = currentYearEventDate.getDay();
      
      // If today is Sunday (0), (7 - 0) % 7 = 0 (same day). 
      // Use (7 - dayOfWeek) % 7 to stay on same day if Sunday, or change to ((7 - dayOfWeek) || 7) to force next Sunday.
      const daysUntilSunday = (7 - dayOfWeek) % 7; 

      const nextSunday = new Date(currentYearEventDate);
      nextSunday.setDate(currentYearEventDate.getDate() + daysUntilSunday);

      const formattedEventDate = Utilities.formatDate(nextSunday, targetTimeZone, "EEEE, MMMM d, yyyy");

      // Load and evaluate HTML Template
      const template = HtmlService.createTemplateFromFile('EmailTemplate');
      template.recipientName = name || "Devotee";
      template.eventDate = formattedEventDate;
      template.subject = subject;
      template.leadFlag = leadFlag;
      template.headerImage = true;

      const htmlBody = template.evaluate().getContent();

      // Send email
      MailApp.sendEmail({
        to: email,
        cc: "heemansu@gmail.com",
        subject: subject,
        htmlBody: htmlBody,
        inlineImages: {
          headerImage: imageBlob
        }
      });

      // Mark as Sent in Column E
      sheet.getRange(i + 2, 5).setValue("Sent");
    }
  }
}
