// Paste this into Google Apps Script (Extensions → Apps Script)
// Then Deploy → New Deployment → Web App → Anyone can access

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  
  // Check for duplicate emails
  var emails = sheet.getRange("A:A").getValues().flat();
  if (emails.includes(data.email)) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "duplicate", message: "Already on the list!" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  sheet.appendRow([data.email, new Date().toISOString()]);
  
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "You're in!" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "FFB Waitlist API" }))
    .setMimeType(ContentService.MimeType.JSON);
}
