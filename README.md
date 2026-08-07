Automated Date-Based Email Sender (Google Apps Script)
An automated Google Apps Script project that parses dynamic Google Sheets data to trigger date-based event emails (such as Anniversary and Birthday greetings) with personalized content, inline header banners, dynamic upcoming Sunday event dates, and automatic status tracking.

🛠️ Features
Dynamic Text Rendering: Automatically toggles message copy between Anniversary and Birthday greetings based on the email subject line.

Month/Day Annual Matching: Ignores the year stored in the spreadsheet to allow recurring annual triggers (e.g., matching August 4th regardless of original event year).

Automatic Sunday Calculation: Calculates the next coming Sunday date for the current calendar year to invite recipients to lead weekly religious services (Abhishek and Pooja).

Inline Image Embeds: Uses Google Drive File IDs to safely inline-embed banner graphics (cid:headerImage) across email clients.

Status Updates: Marks row entries as "Sent" in column E upon successful delivery to prevent duplicate emails.

📂 Project Structure
Plaintext
├── Code.gs             # Core Google Apps Script logic
└── EmailTemplate.html  # Responsive HTML template with Apps Script scriptlets

🚀 Setup & Configuration
1. Google Sheet Setup
Ensure your active Google Sheet has the following column layout starting from Row 2:

Column,Header Name,Description / Example
A,Name,"Recipient Name (e.g., Souratn and Silky)"
B,Email,Target email address
C,Event Date,"Date string or Date object (e.g., 2021-08-04)"
D,Subject,"Email subject (e.g., Anniversary wishes or Birthday wishes)"
E,Status,Execution status (Sent or blank)
F,Sunday Date,Reserved for Sunday calculation output/tracking

2. Google Apps Script Setup
Open your Google Sheet.

Navigate to Extensions > Apps Script.

Create two files in the editor:

Code.gs: Paste the script logic.

EmailTemplate.html: Create an HTML file with this exact name and paste the template code.

Replace the imageFileId variable in Code.gs with your Google Drive File ID for the header banner:

JavaScript
const imageFileId = "YOUR_GOOGLE_DRIVE_FILE_ID_HERE";


⏰ Automation (Time-Driven Trigger)
To automatically send emails daily:

In Apps Script, click the Triggers icon (clock on the left sidebar).

Click + Add Trigger.

Set the following options:

Function to run: sendDateBasedEmails

Select event source: Time-driven

Type of time-based trigger: Day timer

Select time of day: Choose preferred daily window (e.g., 8am to 9am)

Save permissions.
