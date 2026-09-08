/**
 * 🧪 TEST RUN FUNCTION: Click "Run" ▶️ on this function to test sending an email!
 */
function testSendEmail() {
  // ⚠️ REPLACE THE EMAIL BELOW WITH YOUR REAL GMAIL ADDRESS:
  var myEmail = "YOUR_ACTUAL_EMAIL@gmail.com"; 
  
  Logger.log("Attempting to send test email to: " + myEmail);
  
  var sampleName        = "Rahul Sharma";
  var sampleRefId       = "IOTIP-2026-85318";
  var sampleCollege     = "IIT Kharagpur";
  var sampleBranch      = "Computer Science & Engineering";
  var sampleYear        = "4th Year";
  var sampleDomain      = "Data Science";
  var sampleLanguages   = "English, Hindi, Bengali";
  var sampleState       = "West Bengal";
  
  sendConfirmationEmail(myEmail, sampleName, sampleRefId, sampleCollege, sampleBranch, sampleYear, sampleDomain, sampleLanguages, sampleState);
  
  Logger.log("✅ Check your inbox! Test email sent successfully to " + myEmail);
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
    var refId         = p.refId || ("IOTIP-2026-" + Math.floor(10000 + Math.random() * 90000));
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
        "College Email", "Personal Email", "State / Region", 
        "College Name", "Branch / Stream", "Year of Study", 
        "Domain Track", "Preferred Languages"
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
      sendConfirmationEmail(recipientEmail, fullName, refId, collegeName, branch, yearOfStudy, domain, languages, state);
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
 * 📧 EMAIL BUILDER & DISPATCHER (IBM x ACCENLEARN Branded)
 */
function sendConfirmationEmail(email, name, refId, college, branch, year, domain, languages, state) {
  // DYNAMIC SUBJECT LINE: Your [Selected Track] Program Registration is Confirmed
  var trackTitle = domain || "IOTIP";
  var subject = "Your " + trackTitle + " Program Registration is Confirmed [" + refId + "]";
  var waGroupUrl = "https://chat.whatsapp.com/FKC4AVeDHav3KPP3789KNf?s=cl&p=a&mlu=4&ilr=4";
  
  var htmlBody = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    'body { font-family: Arial, sans-serif; background-color: #0b0f19; margin: 0; padding: 20px; color: #f1f5f9; }' +
    '.container { max-width: 600px; margin: 0 auto; background-color: #111827; border: 1px solid #1e293b; border-radius: 12px; padding: 24px; }' +
    '.header { background: linear-gradient(135deg, #1d4ed8, #0f172a); padding: 20px; text-align: center; border-radius: 8px; margin-bottom: 20px; border: 1px solid #3b82f6; }' +
    '.brand { color: #60a5fa; font-size: 11px; font-weight: bold; letter-spacing: 2px; text-transform: uppercase; }' +
    '.title { color: #ffffff; font-size: 22px; margin: 8px 0 0; font-weight: 800; }' +
    '.ref { background: rgba(30, 41, 59, 0.8); border: 1px dashed #3b82f6; border-radius: 8px; padding: 12px; text-align: center; margin: 20px 0; color: #93c5fd; font-family: monospace; font-size: 14px; }' +
    '.counsellor-box { background: rgba(34, 197, 94, 0.12); border: 1px solid #22c55e; border-radius: 8px; padding: 14px; text-align: center; margin-bottom: 20px; color: #4ade80; font-size: 15px; font-weight: bold; }' +
    '.card { background: #1e293b; border-radius: 8px; padding: 18px; margin-bottom: 16px; border: 1px solid #334155; }' +
    '.card h4 { margin: 0 0 12px; color: #60a5fa; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; }' +
    '.row { margin-bottom: 8px; font-size: 14px; color: #e2e8f0; line-height: 1.5; }' +
    '.btn { display: block; background: #22c55e; color: #ffffff !important; text-decoration: none; text-align: center; font-weight: bold; padding: 16px; border-radius: 8px; margin-top: 20px; font-size: 15px; }' +
    '.footer-note { text-align: center; color: #64748b; font-size: 11px; margin-top: 20px; font-family: monospace; }' +
    '</style></head><body>' +
    '<div class="container">' +
    '<div class="header"><div class="brand">IBM x ACCENLEARN</div><h1 class="title">Registration Confirmed &#127881;</h1></div>' +
    
    '<p style="color: #ffffff !important; font-size: 17px; font-weight: bold; margin: 16px 0 8px 0;">Dear ' + name + ',</p>' +
    '<p style="color: #cbd5e1 !important; font-size: 14.5px; line-height: 1.6; margin: 0 0 18px 0;">Thank you for registering for the <strong style="color: #ffffff !important;">Industry-Oriented Training & Internship Program (IOTIP) 2026</strong> presented by <strong style="color: #60a5fa !important;">IBM x ACCENLEARN</strong>. Your profile has been successfully verified and logged into our system database.</p>' +
    
    '<div class="ref">Official Registration ID: <strong style="color: #ffffff !important;">' + refId + '</strong></div>' +
    
    '<!-- PROMINENT COUNSELLOR CALLOUT BOX -->' +
    '<div class="counsellor-box">&#128172; Our counsellor will contact you within 24 hours.</div>' +
    
    '<div class="card">' +
    '<h4>Application Summary Dossier</h4>' +
    '<div class="row"><strong style="color: #94a3b8;">Selected Track:</strong> <span style="color: #60a5fa; font-weight: bold;">' + (domain || 'Not specified') + '</span></div>' +
    '<div class="row"><strong style="color: #94a3b8;">College Name:</strong> <span style="color: #ffffff;">' + (college || 'Not specified') + '</span></div>' +
    '<div class="row"><strong style="color: #94a3b8;">Branch / Year:</strong> <span style="color: #ffffff;">' + (branch || '') + (year ? ' (' + year + ')' : '') + '</span></div>' +
    '<div class="row"><strong style="color: #94a3b8;">State / Region:</strong> <span style="color: #ffffff;">' + (state || 'Not specified') + '</span></div>' +
    '<div class="row"><strong style="color: #94a3b8;">Preferred Language(s):</strong> <span style="color: #ffffff;">' + (languages || 'English') + '</span></div>' +
    '</div>' +
    
    '<div class="card">' +
    '<h4>&#128640; Next Onboarding Milestones</h4>' +
    '<div class="row"><span style="color:#22c55e;">&#10004;</span> <strong style="color: #ffffff;">Application Verified & Profile Created:</strong> Your details are locked in the candidate registry.</div>' +
    '<div class="row"><span style="color:#22c55e;">&#10004;</span> <strong style="color: #ffffff;">Counsellor Contact:</strong> Our counsellor will contact you within 24 hours.</div>' +
    '<div class="row"><span style="color:#22c55e;">&#10004;</span> <strong style="color: #ffffff;">Batch & Schedule Confirmation:</strong> Official schedule & orientation pass will be shared during your onboarding call.</div>' +
    '<div class="row"><span style="color:#22c55e;">&#10004;</span> <strong style="color: #ffffff;">Join Student Network:</strong> Connect with cohort members & placement mentors immediately.</div>' +
    '</div>' +
    
    '<a href="' + waGroupUrl + '" target="_blank" class="btn">&#128172; Join Official Student WhatsApp Group</a>' +
    '<div class="footer-note">© 2026 IBM x ACCENLEARN. All Rights Reserved.</div>' +
    '</div></body></html>';

  // Send via GmailApp directly
  GmailApp.sendEmail(email, subject, "", {
    name: "IBM x ACCENLEARN",
    htmlBody: htmlBody
  });
}
