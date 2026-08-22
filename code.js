// --- CONFIGURATION ---
const WEB_APP_URL = "YOUR_WEB_APP_URL_HERE"; 
const SHEET_NAME = "YOUR_SHEET_NAME_HERE"; // e.g., "Form Responses 1"

// Column Mapping (0-based index is not used here; uses column number where A=1, B=2, etc.)
const EMP_EMAIL_COL = 2;    // Column B 
const NAME_COL = 3;         // Column C 
const TYPE_COL = 4;         // Column D 
const AMOUNT_COL = 5;       // Column E 
const MGR_EMAIL_COL = 6;    // Column F 
const STATUS_COL = 7;       // Column G - Manager Status
const ADJUST_AMT_COL = 8;   // Column H - Adjusted Amount
const HR_STATUS_COL = 11;   // Column K - HR Status
const HR_NOTIFIED_COL = 16; // Column P - To track if HR sent email
// ---------------------

function sendEmailOnFormSubmission(e) {
  const sheet = e.source.getSheetByName(SHEET_NAME);
  const row = e.range.getRow();
  
  const empEmail = sheet.getRange(row, EMP_EMAIL_COL).getDisplayValue();
  const name = sheet.getRange(row, NAME_COL).getDisplayValue();
  const advType = sheet.getRange(row, TYPE_COL).getDisplayValue();
  const amount = sheet.getRange(row, AMOUNT_COL).getDisplayValue();
  const mgrEmail = sheet.getRange(row, MGR_EMAIL_COL).getDisplayValue();

  const approveUrl = `${WEB_APP_URL}?action=Approve&row=${row}`;
  const denyUrl = `${WEB_APP_URL}?action=Deny&row=${row}`;
  const adjustUrl = `${WEB_APP_URL}?action=Adjust&row=${row}`;

  const subject = `Nepieciešams apstiprināt / Confirmation required: ${name}`;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.5;">
      <h3 style="color: #2c3e50;">Jauns avansa pieprasījums / New Advance Request</h3>
      <ul>
        <li><strong>Darbinieks / Employee:</strong> ${name}</li>
        <li><strong>Avansa veids / Advance Type:</strong> ${advType}</li>
        <li><strong>Summa / Amount:</strong> ${amount}</li>
      </ul>
      <p>Lūdzu, izvēlieties vienu no iespējām: / Please select an option:</p>
      <br>
      <a href="${approveUrl}" style="background-color: #4CAF50; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; display: inline-block; margin-right: 10px; font-weight: bold;">Apstiprināt / Approve</a>
      <a href="${denyUrl}" style="background-color: #f44336; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; display: inline-block; margin-right: 10px; font-weight: bold;">Noraidīt / Deny</a>
      <a href="${adjustUrl}" style="background-color: #ff9800; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold;">Mainīt summu / Change Amount</a>
      
      <p style="color: #f8f9fa; font-size: 8px; margin-top: 30px;">Ref: ${new Date().getTime()}_${row}</p>
    </div>
  `;

  MailApp.sendEmail({ to: mgrEmail, subject: subject, htmlBody: htmlBody });
}

function doGet(e) {
  const action = e.parameter.action;
  const row = e.parameter.row;

  if (!action || !row) return HtmlService.createHtmlOutput("Invalid request / Nederīgs pieprasījums.");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  
  // 1. If Manager wants to adjust the amount, show them an HTML form to type it in
  // Uses target="_top" so it properly redirects to the success page out of the iframe
  if (action === "Adjust") {
    return HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h2>Pielāgot avansa summu / Adjust Advance Amount</h2>
        <form action="${WEB_APP_URL}" method="GET" target="_top">
          <input type="hidden" name="action" value="SaveAdjusted">
          <input type="hidden" name="row" value="${row}">
          <p>Ievadiet jauno summu / Enter the new amount:</p>
          <input type="text" name="newAmount" required style="padding: 10px; font-size: 16px; width: 200px; margin-bottom: 20px;">
          <br>
          <button type="submit" style="background-color: #4CAF50; color: white; padding: 12px 24px; border: none; border-radius: 4px; font-size: 16px; cursor: pointer;">Saglabāt / Save</button>
        </form>
      </div>
    `);
  }

  // Check if already processed
  const currentStatus = sheet.getRange(row, STATUS_COL).getDisplayValue();
  if (currentStatus !== "") {
     return HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h3 style="color: #f44336;">Šis pieprasījums jau ir apstrādāts. / This request has already been processed.</h3>
        <p>Tagad varat droši aizvērt šo logu. / You may now safely close this tab.</p>
      </div>
     `);
  }

  // 2. Process the action (Approve, Deny, or SaveAdjusted)
  const empEmail = sheet.getRange(row, EMP_EMAIL_COL).getDisplayValue();
  const name = sheet.getRange(row, NAME_COL).getDisplayValue();
  const amount = sheet.getRange(row, AMOUNT_COL).getDisplayValue();
  
  let finalStatus = "";
  let subject = "";
  let bodyInfo = "";
  let introText = ""; 

  if (action === "Approve") {
    finalStatus = "Apstiprināts / Approved";
    subject = "Avansa pieprasījums apstiprināts / Advance request approved";
    introText = "Jūsu avansa pieprasījums ir apstiprināts. / Your advance request has been approved.";
    bodyInfo = `Sākotnējā summa / Original amount: ${amount}`;
    sheet.getRange(row, STATUS_COL).setValue(finalStatus);
  } 
  else if (action === "Deny") {
    finalStatus = "Noraidīts / Declined";
    subject = "Avansa pieprasījums noraidīts / Advance request denied";
    introText = "Jūsu avansa pieprasījums ir noraidīts. / Your advance request has been denied.";
    bodyInfo = `Summa / Amount: ${amount}`;
    sheet.getRange(row, STATUS_COL).setValue(finalStatus);
  } 
  else if (action === "SaveAdjusted") {
    const newAmount = e.parameter.newAmount;
    finalStatus = "Pielāgota summa / Adjusted amount";
    subject = "Avansa pieprasījuma summa pielāgota / Advance request adjusted";
    introText = "Jūsu avansa pieprasījuma summa ir pielāgota. / Your advance request amount has been adjusted.";
    bodyInfo = `Sākotnējā summa / Original amount: ${amount}\nJaunā summa / New amount: ${newAmount}`;
    
    // Write both the status and the new amount to the sheet
    sheet.getRange(row, STATUS_COL).setValue(finalStatus);
    sheet.getRange(row, ADJUST_AMT_COL).setValue(newAmount);
  }

  // Email the Employee
  const body = `Sveiki / Hello ${name},\n\n${introText}\n\n${bodyInfo}\n\nGeros dienos! / Have a great day!`;
  MailApp.sendEmail(empEmail, subject, body);

  return HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
      <h2 style="color: #4CAF50;">Paveikts! / Success!</h2>
      <p>Pieprasījuma statuss veiksmīgi atjaunināts uz / The request status was successfully updated to:<br><br><strong>${finalStatus}</strong></p>
      <p>Darbiniekam ir nosūtīts e-pasts. / An email notification has been sent to the employee.</p>
      <p style="margin-top: 30px; color: #555;">Tagad varat droši aizvērt šo logu. / You may now safely close this tab.</p>
    </div>
  `);
}

// 3. HR Trigger
function onEdit(e) {
  const editedCell = e.range;
  const sheet = e.source.getActiveSheet();
  
  if (sheet.getName() === SHEET_NAME && editedCell.getColumn() === HR_STATUS_COL) {
    const row = editedCell.getRow();
    const statusHR = editedCell.getValue();
    const isNotified = sheet.getRange(row, HR_NOTIFIED_COL).getValue();

    if (statusHR === "Reģistrēts / Registered" && isNotified !== "Sent") {
      const empEmail = sheet.getRange(row, EMP_EMAIL_COL).getDisplayValue();
      const name = sheet.getRange(row, NAME_COL).getDisplayValue();
      const amount = sheet.getRange(row, AMOUNT_COL).getDisplayValue();
      
      const subject = "Avansa pieprasījums ir reģistrēts / Advance request has been registered";
      const body = `Sveiki / Hello ${name},\n\nJūsu avansa pieprasījums (${amount}) ir reģistrēts. / Your advance request has been registered.\n\nPaldies! / Thank you!`;
      
      MailApp.sendEmail(empEmail, subject, body);
      sheet.getRange(row, HR_NOTIFIED_COL).setValue("Sent");
    }
  }
}
