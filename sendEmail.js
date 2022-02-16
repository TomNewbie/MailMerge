let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
let formSheet = spreadsheet.getSheetByName(form);
let interviewSheet = spreadsheet.getSheetByName(interview);
let rejectSheet = spreadsheet.getSheetByName(reject);
let passSheet = spreadsheet.getSheetByName(pass);
let offerSheet = spreadsheet.getSheetByName(offer);
let timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

// function Test() {
//   console.log(formSheet.getLastRow());
// }
//A function to send one email based on the information on one row.
  // The input is:
  // - the template(interviewTemplate,  passTemplate, rejectTemplate, offerTemplate)
  // - the sheet (interviewSheet, rejectSheet,passSheet, offerSheet)
  // - the column
  // - the row: First Row Data in InterviewSheet
function sendEmail(template, sheet, column, row) {
  let data = [];
  //return all data of one recipient depended on the sheet
  for (let i = 0; i < column.length; i++) {
    data.push(sheet.getRange(column[i] + row).getValue());
  } 
  // data[0] is email
  let subject = subjectResult;
  let email = data[0];
  // Format date and translate to Vietnamese of offerTemplate
  if (template === offerTemplate) {
    //Email - Tên - Vị trí cũ - Ban cũ - Vị trí offer - Ban offer - Deadline
    data[6] = Utilities.formatDate(data[6], timezone, `HH:mm, EEEEEEEE, dd/MM/yyyy`);
    data[6] = LanguageApp.translate(data[6], 'en', 'vi');
  }
  // Fill in the passTemplate, rejectTemplate, offerTemplate with the data
  let bodyHtml = replaceTemplate(template, data, 1);

  // Fill in the interviewTemplate with the data
  if (template === interviewTemplate) {
    // Formated date and time to vietnamse
    data[4] = Utilities.formatDate(data[4], timezone, `EEEEEEEE, dd/MM/yyyy`);
    data[4] = LanguageApp.translate(data[4], 'en', 'vi');
    data[5] = Utilities.formatDate(data[5], timezone, "HH:mm ")
    data[8] = Utilities.formatDate(data[8], timezone, `HH:mm, EEEEEEEE, dd/MM/yyyy`);
    data[8] = LanguageApp.translate(data[8], 'en', 'vi');
    data.push(data[8]);
    data.push(data[8]);
    //Vị trí - Email - Họ và tên - Vị trí - Ban - Lịch - Thời gian - Hình thức - Địa điểm - Online - Deadline - Deadline
    data.unshift(data[2]);
    if (data[7] === 'Phỏng vấn online') {
      data[9] = online;
    }
    else {
      data[9] = '<br>';
    }
    subject = replaceTemplate(subjectInterview, data, 0);
    email = data[1]; //data[1] la email
    bodyHtml = replaceTemplate(interviewTemplate, data, 2);
  }

  //Send one email to each recipient (https://developers.google.com/apps-script/reference/mail/mail-app?hl=en#sendemailrecipient,-subject,-body,-options)
  MailApp.sendEmail(email, subject, 'Your email does not support HTML email', {
    cc: ccEmail,
    name: 'VGU Student Association - VSA',
    htmlBody: bodyHtml
  });

  //After sending, check the 'Send Email' cell
  sheet.getRange("A" + row).check();
}

//Replace the placeholder in the template with data of the recipient
//The input is:
//- the template (interviewTemplate,  passTemplate, rejectTemplate, offerTemplate)
//- the data of recipient (một array)
//- the index: thứ tự của họ tên trong array data
function replaceTemplate(template, data, index) {
  while (template.match(/{[a-z]+}/) !== null) {
    let keyword = template.match(/{[a-z]+}/)[0];
    template = template.replace(keyword, data[index]);
    index++;
  }
  return template;
}

//Check if the email is sent or not
//The input is:
//- the sheet (interviewSheet, rejectSheet,passSheet, offerSheet)
//- the specific row wanted to check
function isMailAlreadySent(sheet, row) {
  if (!sheet.getRange("A" + row).isChecked() && !sheet.getRange("D" + row).isBlank()) {
    return false
  }
  return true
}

