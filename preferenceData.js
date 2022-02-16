// naming convention
let form = 'Form Responses';
let pass = 'Pass';
let reject = 'Reject';
let offer = 'Offer'
let interview = 'Interview';    
let resultSheet = [interview, pass, reject, offer];
let ccEmail = '' // mail for cc

let interviewDocLink = '';
let passDocLink = '';
let rejectDocLink = '';
let offerDocLink = ''; 
// url of the code, get the template
let interviewTemplate = DocumentApp.openByUrl(interviewDocLink).getBody().getText();
let passTemplate = DocumentApp.openByUrl(passDocLink).getBody().getText();
let rejectTemplate = DocumentApp.openByUrl(rejectDocLink).getBody().getText();
let offerTemplate = DocumentApp.openByUrl(offerDocLink).getBody().getText();

// subject email
let subjectInterview = 'VSA - Thư mời phỏng vấn - {vitri}';
let subjectResult = 'VSA - Thông báo kết quả vòng Phỏng vấn';

// number row of data
let rowInterview = 6; // First Row Data in InterviewSheet
let test_rowInterview = 3; //The Row number for data test

let row = 6;
let testRow = 4;

// Get and display data
let cot_hien_thi = "C1";
let cot_data = "E1";
let displayQuery = "A2"

let online = 'Khi tham gia phỏng vấn, bạn vui lòng bật camera, chuẩn bị micro sử dụng được và ngồi ở nơi yên tĩnh, tránh gây gián đoạn buổi phỏng vấn.<br><br>'