# Google Workspace Advance Request Automation

A Google Apps Script automated workflow that transforms standard Google Form submissions into an interactive, multi-step approval process for managers and HR. 

By leveraging Google Sheets, Gmail, and Google's `HtmlService`, this project eliminates manual spreadsheet tracking and allows managers to approve, deny, or adjust financial requests directly from their inbox.

## Key Features

* **Interactive Email Approvals:** Uses `MailApp` to send customized HTML emails with embedded action buttons, allowing managers to process requests with a single click.
* **Custom Web App Integration:** The "Change Amount" button redirects managers to a secure, script-hosted webpage (`doGet`) where they can input and save an adjusted financial figure.
* **Bilingual Communication:** All automated emails, forms, and success pages are dynamically generated in both Latvian and English.
* **Automated HR Sync:** Utilizes an `onEdit` simple trigger to instantly notify the employee once HR officially registers the request in the background spreadsheet.

## Workflow Overview

1. **Submission:** An employee submits a Google Form requesting a financial advance.
2. **Notification:** The `onFormSubmit` trigger fires, generating an email to the direct manager with the request details and three action buttons (Approve, Deny, Change Amount).
3. **Manager Action:** 
   * Clicking *Approve* or *Deny* instantly updates the Google Sheet and emails the employee the final decision.
   * Clicking *Change Amount* opens a custom web app for the manager to input a new value. Upon submission, it updates the Sheet and notifies the employee of the adjusted amount.
4. **HR Processing:** When HR reviews the sheet and changes the status column to "Registered", a background `onEdit` trigger fires and sends a final confirmation email to the employee.

## Setup Instructions

1. **Prepare the Google Sheet:** Create a Google Form for advance requests and link it to a Google Sheet. Ensure your columns match the mapping in the configuration section.
2. **Add the Script:** Go to `Extensions > Apps Script` in your Google Sheet, clear the default code, and paste the contents of `Code.js`.
3. **Deploy the Web App:**
   * Click **Deploy > New deployment**.
   * Select **Web app**.
   * Execute as: **Me**.
   * Who has access: **Anyone**.
   * Copy the generated Web App URL.
4. **Configure the Script:** Paste your Web App URL into the `WEB_APP_URL` variable at the top of the script. Update `SHEET_NAME` to match your exact tab name.
5. **Set the Trigger:** Click the clock icon (Triggers) on the left sidebar and add a new trigger for `sendEmailOnFormSubmission` to run on `From spreadsheet` > `On form submit`.

## Technologies Used

* Google Apps Script (JavaScript)
* Google Sheets API
* MailApp / GmailApp API
* HTML / CSS (for email styling and Web App interface)
