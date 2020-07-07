// This constant is written in column C for rows for which an email has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';

// Email Template
var htmlBody = HtmlService.createHtmlOutputFromFile('mail_template').getContent();

// File Attachment to be changed as required
var file = DriveApp.getFileById('your_file_id');

// Set Salutation in the email and create template
function generateEmailTemplate(firstName) {
  var t = HtmlService.createTemplateFromFile('mail_template');
  t.first_name = firstName;
  return t.evaluate();
}

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmailsToAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = sheet.getLastRow() - 1; // Number of rows to process
  var numCols = sheet.getLastColumn(); // Number of columns to process
  
  // Fetch the entire data excluding the header row
  var dataRange = sheet.getRange(startRow, 1, numRows, numCols);
  
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var firstName = row[0]; // First column
    var lastName = row[1]; // Second column
    var emailAddress = row[2]; // Third column
    var emailSent = row[numCols]; // Last column
    
    htmlBody = generateEmailTemplate(firstName).getContent();
    
    if (emailSent !== EMAIL_SENT) { // Prevents sending duplicates
      MailApp.sendEmail({
        to: emailAddress,
        subject: 'My Email has a subject 😱',
        htmlBody: htmlBody,
        name: 'You got a name, bro?',
        attachments: [file.getAs(MimeType.PDF)]
      });
      sheet.getRange(startRow + i, numCols).setValue(EMAIL_SENT);
      
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
