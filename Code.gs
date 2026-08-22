/**
 * 🧪 TEST RUN FUNCTION: Click "Run" ▶️ on this function to test sending an email!
 */
function testSendEmail() {
  var myEmail = Session.getActiveUser().getEmail() || "your_email@gmail.com";
  
  var sampleName        = "Rahul Sharma";
  var sampleRefId       = "COHORT-2026-853185";
  var sampleCollege     = "IIT Kharagpur";
  var sampleBranch      = "Computer Science & Engineering";
  var sampleYear        = "4th Year";
  var sampleDomain      = "Artificial Intelligence & Machine Learning";
  var sampleLanguages   = "English, Hindi, Bengali";
  
  sendConfirmationEmail(myEmail, sampleName, sampleRefId, sampleCollege, sampleBranch, sampleYear, sampleDomain, sampleLanguages);
}

/**
 * 🌐 WEB APP GET HANDLER (For Health Check / Browser Visit)
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ "status": "online", "message": "IBM x ACCENLEARN WebApp Endpoint Active" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🌐 WEB APP SUBMISSION HANDLER (Handles form submit POST requests)
 */
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var p = e.parameter || {};
    
    // Safely extract parameters from JSON body if e.parameter is empty
    if (e.postData && e.postData.contents) {
      try {
        var jsonParsed = JSON.parse(e.postData.contents);
        p = Object.assign({}, p, jsonParsed);
      } catch (jsonErr) {
        // Continue with e.parameter
      }
    }
    
    // Extract ALL 13 form parameters
    var timestamp     = p.timestamp || new Date().toLocaleString();
    var refId         = p.refId || ("COHORT-2026-" + Math.floor(100000 + Math.random() * 900000));
    var fullName      = p.fullName || "Applicant";
    var phone         = p.phone || "";
    var whatsapp      = p.whatsapp || "";
    var collegeEmail  = p.collegeEmail || "";
    var personalEmail = p.personalEmail || "";
    var state         = p.state || "";
    var collegeName   = p.collegeName || "";
    var branch        = p.branch || "";
    var yearOfStudy   = p.yearOfStudy || p.year || "";
    var domain        = p.domain || "";
    var languages     = p.languages || "";
    
    // Auto-create Header row if empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", "Ref ID", "Full Name", "Phone", "WhatsApp", 
        "College Email", "Personal Email", "State", 
        "College Name", "Branch / Stream", "Year of Study", 
        "Domain", "Languages"
      ]);
    }
    
    // Append full data row to Google Sheet
    sheet.appendRow([
      timestamp, refId, fullName, phone, whatsapp, 
      collegeEmail, personalEmail, state, 
      collegeName, branch, yearOfStudy, 
      domain, languages
    ]);
    
    // Dispatch confirmation email
    var recipientEmail = personalEmail || collegeEmail;
    if (recipientEmail && recipientEmail.indexOf("@") !== -1) {
      sendConfirmationEmail(recipientEmail, fullName, refId, collegeName, branch, yearOfStudy, domain, languages);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success", "refId": refId }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 📧 EMAIL BUILDER & DISPATCHER (High-Contrast Theme for IBM x ACCENLEARN)
 */
function sendConfirmationEmail(email, name, refId, college, branch, year, domain, languages) {
  var subject = "Application Confirmed: IBM x ACCENLEARN 2026 Cohort [" + refId + "]";
  var waGroupUrl = "https://chat.whatsapp.com/FaoL6LzlE5p0ih6PKxubWQ?s=cl&p=a&mlu=4";
  
  var htmlBody = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    'body { font-family: Arial, sans-serif; background-color: #0b0f19; margin: 0; padding: 20px; color: #f1f5f9; }' +
    '.container { max-width: 600px; margin: 0 auto; background-color: #111827; border: 1px solid #1e293b; border-radius: 12px; padding: 24px; }' +
    '.header { background: linear-gradient(135deg, #1d4ed8, #0f172a); padding: 20px; text-align: center; border-radius: 8px; margin-bottom: 20px; }' +
    '.brand { color: #60a5fa; font-size: 11px; font-weight: bold; letter-spacing: 2px; text-transform: uppercase; }' +
    '.title { color: #ffffff; font-size: 22px; margin: 8px 0 0; }' +
    '.ref { background: rgba(30, 41, 59, 0.8); border: 1px dashed #3b82f6; border-radius: 8px; padding: 12px; text-align: center; margin: 20px 0; color: #93c5fd; font-family: monospace; font-size: 14px; }' +
    '.card { background: #1e293b; border-radius: 8px; padding: 18px; margin-bottom: 16px; border: 1px solid #334155; }' +
    '.card h4 { margin: 0 0 12px; color: #60a5fa; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; }' +
    '.row { margin-bottom: 8px; font-size: 14px; color: #e2e8f0; }' +
    '.btn { display: block; background: #22c55e; color: #ffffff !important; text-decoration: none; text-align: center; font-weight: bold; padding: 16px; border-radius: 8px; margin-top: 20px; font-size: 15px; }' +
    '</style></head><body>' +
    '<div class="container">' +
    '<div class="header"><div class="brand">IBM x ACCENLEARN</div><h1 class="title">Application Received 🎉</h1></div>' +
    
    '<p style="color: #ffffff !important; font-size: 17px; font-weight: bold; margin: 16px 0 8px 0;">Dear ' + name + ',</p>' +
    '<p style="color: #cbd5e1 !important; font-size: 14.5px; line-height: 1.6; margin: 0 0 20px 0;">Thank you for applying to the <strong style="color: #ffffff !important;">IBM x ACCENLEARN 2026 Cohort</strong>. Your profile has been successfully logged into our cohort database.</p>' +
    
    '<div class="ref">Application Reference ID: <strong style="color: #ffffff !important;">' + refId + '</strong></div>' +
    
    '<div class="card">' +
    '<h4>Application Summary</h4>' +
    '<div class="row"><strong style="color: #94a3b8;">Selected Track:</strong> <span style="color: #ffffff;">' + (domain || 'Not specified') + '</span></div>' +
    '<div class="row"><strong style="color: #94a3b8;">College Name:</strong> <span style="color: #ffffff;">' + (college || 'Not specified') + '</span></div>' +
    '<div class="row"><strong style="color: #94a3b8;">Branch / Year:</strong> <span style="color: #ffffff;">' + (branch || '') + (year ? ' (' + year + ')' : '') + '</span></div>' +
    '<div class="row"><strong style="color: #94a3b8;">Preferred Language(s):</strong> <span style="color: #ffffff;">' + (languages || 'English') + '</span></div>' +
    '</div>' +
    
    '<div class="card">' +
    '<h4>What Happens Next</h4>' +
    '<div class="row"><span style="color:#22c55e;">✓</span> <strong style="color: #ffffff;">Application Review:</strong> Evaluating profile against requirements (Est. 24h).</div>' +
    '<div class="row"><span style="color:#22c55e;">✓</span> <strong style="color: #ffffff;">WhatsApp Contact:</strong> An onboarding specialist will message you on WhatsApp in your preferred language.</div>' +
    '<div class="row"><span style="color:#22c55e;">✓</span> <strong style="color: #ffffff;">Program Starter Kit:</strong> Full curriculum syllabus will be sent to your registered email.</div>' +
    '</div>' +
    
    '<a href="' + waGroupUrl + '" target="_blank" class="btn">💬 Join Official Student WhatsApp Group</a>' +
    '</div></body></html>';

  MailApp.sendEmail({
    to: email,
    subject: subject,
    name: "IBM x Accenlearn",
    htmlBody: htmlBody
  });
}
