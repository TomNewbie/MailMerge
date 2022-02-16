//Add button to spreadsheet to use the function
function addUI() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu("VSA Automatic");
  menu.addItem("Send Email", "sendRealEmail");
  let separator = menu.addSeparator()
  menu.addItem("Get Remaining Quora", "getRemainQuota");
  menu.addItem("Send Test Email", "sendTestEmail")
  menu.addToUi();
}

//Change the content in the interviewSheet, rejectSheet,passSheet, offerSheet based on the resultSheet
function changeQueryContent(e) {
  console.log("edit run")
  if (e.source.getCurrentCell().getA1Notation() === cot_hien_thi) {
    let value = e.range.getValue();
    console.log(value)
    let cell = e.source.getActiveSheet().getRange(displayQuery);
    let nameSheet = e.source.getActiveSheet().getName();
    if (resultSheet.includes(nameSheet)) {
      if (nameSheet === interview) {
        setQueryStatement(cell, value, null);
      }
      else {
        setQueryStatement(cell, value, nameSheet);
      }
    }
  }
}

//Set value from form to corresponding sheet
//The input is:
//- the targeted cell
//- the value want to set 
//- the sheet(interviewSheet, rejectSheet,passSheet, offerSheet)
function setQueryStatement(cell, value, nameSheet) {
  if (nameSheet === null) {
    cell.setValue("SELECT " + value);
  }
  else {
    cell.setValue("SELECT " + value + ` WHERE A LIKE '${nameSheet}'`);
  }
}

//Returns the number of remaining emails a user can send for the rest of the day
function getRemainQuota() {
  var quotaLeft = MailApp.getRemainingDailyQuota();
  SpreadsheetApp.getUi().alert("Your email can sent " + quotaLeft + " emails now");
}

//Send emails to all recipients(rows) in the active sheet
function runSendEmail(testEmail = false) {
  let sheet = spreadsheet.getActiveSheet();
  let nameSheet = sheet.getName();
  let cell = sheet.getRange(cot_data).getValue();
  let column = cell.split(/[\s,]+/);
  // let lastRow = sheet.getRange("C1:C").getValues().filter(String).length + 2;
  let lastRow = sheet.getLastRow();
  if (nameSheet === interview) {
    runData(testEmail, interviewTemplate, interviewSheet, column, lastRow);
  }
  else if (nameSheet === reject) {
    runData(testEmail, rejectTemplate, rejectSheet, column, lastRow);
  }
  else if (nameSheet === pass) {
    runData(testEmail, passTemplate, passSheet, column, lastRow);
  }
  else if (nameSheet === offer) {
    runData(testEmail, offerTemplate, offerSheet, column, lastRow);
  }
  else {
    SpreadsheetApp.getUi().alert("Choose sheet Interview, Pass, Reject or Offer to send the Email");
  }
}

//Go to each row and send email, then check the cell
//The input is:
// - testEmail: Boolean 
// - the template(interviewTemplate,  passTemplate, rejectTemplate, offerTemplate)
// - the sheet (interviewSheet, rejectSheet,passSheet, offerSheet)
// - columns to get data of recipient
// - last row in the sheet
function runData(testEmail, template, sheet, column, lastRow) {
  if (testEmail) {
    sendEmail(template, sheet, column, testRow)
  }
  else {
    for (row; row < lastRow; row++) {
      if (!isMailAlreadySent(sheet, row)) {
        sendEmail(template, sheet, column, row)
      }
    }
  }
}

//Send real emails
function sendRealEmail() {
  runSendEmail()
}

//Send test email
function sendTestEmail() {
  runSendEmail(true)
}