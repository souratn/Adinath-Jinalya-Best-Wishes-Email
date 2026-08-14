function sendDateBasedEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = 2;
  const numRows = sheet.getLastRow() - 1;

  if (numRows < 1) return;

  // Read 10 columns (A to J)
  const dataRange = sheet.getRange(startRow, 1, numRows, 10);
  const data = dataRange.getValues();
  const currentYear = new Date().getFullYear();

  // Load the image from Google Drive using its File ID
  const imageFileId = "1ytp3XYrGVLD6zhgs5SJpD64F-mhTi-2t";
  const imageBlob = DriveApp.getFileById(imageFileId).getBlob().setName("headerImage");

  // Map user inputs to valid IANA Timezone Identifiers
  const tzMapping = {
    "IST": "Asia/Kolkata",
    "EST": "America/New_York",
    "EDT": "America/New_York",
    "CST": "America/Chicago",
    "CDT": "America/Chicago",
    "PST": "America/Los_Angeles",
    "PDT": "America/Los_Angeles"
  };

  const now = new Date();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const name = row[0];
    const email = row[1];
    const rawDate = row[2];
    const subject = row[3];
    const status = row[4];
    const rawTimeZone = String(row[5]).trim().toUpperCase();
    const leadFlag = row[6];

    // Columns H, I, J
    const isKidFlag = String(row[7]).trim().toUpperCase();
    const isKid = isKidFlag === "Y" || isKidFlag === "YES";
    const childName = String(row[8]).trim();
    const gender = String(row[9]).trim().toUpperCase(); // 'F' or 'M'

    if (!email || !rawDate || status === "Sent") continue;

    const targetTimeZone = tzMapping[rawTimeZone] || rawTimeZone || "Asia/Kolkata";

    // Get exact current hour and formatted date in the target time zone
    const currentHourInTargetTz = Number(Utilities.formatDate(now, targetTimeZone, "HH"));

    // Check execution hour (8 AM in target timezone)
    if (currentHourInTargetTz !== 8) {
      continue;
    }

    // Compare Month-Day in the target timezone
    const todayMonthDay = Utilities.formatDate(now, targetTimeZone, "MM-dd");

    // Parse input rawDate properly to avoid UTC conversion shifts
    let eventDateObj;
    if (rawDate instanceof Date) {
      eventDateObj = rawDate;
    } else {
      // Fallback for text strings like 'YYYY-MM-DD' or 'MM/DD/YYYY'
      eventDateObj = new Date(rawDate);
    }

    const rowMonthDay = Utilities.formatDate(eventDateObj, targetTimeZone, "MM-dd");

    if (rowMonthDay === todayMonthDay) {
      const monthStr = Utilities.formatDate(eventDateObj, targetTimeZone, "MM");
      const dayStr = Utilities.formatDate(eventDateObj, targetTimeZone, "dd");

      const currentYearEventDate = new Date(currentYear, Number(monthStr) - 1, Number(dayStr), 12, 0, 0);

      // Calculate upcoming Sunday
      const dayOfWeek = currentYearEventDate.getDay();
      const daysUntilSunday = (7 - dayOfWeek) % 7;

      const nextSunday = new Date(currentYearEventDate);
      nextSunday.setDate(currentYearEventDate.getDate() + daysUntilSunday);

      const formattedEventDate = Utilities.formatDate(nextSunday, targetTimeZone, "EEEE, MMMM d, yyyy");

      // Set pronouns based on Gender
      let pronounSubject = "they";
      let pronounObject = "them";
      let pronounPossessive = "their";

      const isFemale = gender === "F" || gender === "GIRL" || gender === "FEMALE";
      const isMale = gender === "M" || gender === "BOY" || gender === "MALE";

      if (isFemale) {
        pronounSubject = "she";
        pronounObject = "her";
        pronounPossessive = "her";
      } else if (isMale) {
        pronounSubject = "he";
        pronounObject = "him";
        pronounPossessive = "his";
      }

      // Check condition: Hide "Abhishek and " if NOT a kid and female
      const hideAbhishek = !isKid && isFemale;

      // Load and evaluate HTML Template
      const template = HtmlService.createTemplateFromFile('EmailTemplate');
      template.recipientName = name || "Devotee";
      template.eventDate = formattedEventDate;
      template.subject = subject;
      template.leadFlag = leadFlag;
      template.headerImage = true;

      // Kid & Gender dynamic parameters
      template.isKid = isKid;
      template.childName = childName;
      template.pronounSubject = pronounSubject;
      template.pronounObject = pronounObject;
      template.pronounPossessive = pronounPossessive;
      template.hideAbhishek = hideAbhishek;

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