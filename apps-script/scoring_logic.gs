const CONFIG = {
  ROLE_NAME: "[ Role ]  ",
  EXP_MIN: 3,
  EXP_MAX: 7,
  SKILLS_MUST_HAVE: [
    "Skill 1",
    "Skill 2",
    "Skill 3",
    "Skill 4",
    "Skill 5"
  ],
  CORRECT_ANSWERS: {
   "QUESTION": "ANSWER"
  },
  SHORTLIST_THRESHOLD: 84,
  REVIEW_THRESHOLD: 48,
  HR_EMAIL: "[ HR_EMAIL ]"
};
function findBestKeyMatch(data, scriptQuestion, minPrefix) {
  minPrefix = minPrefix || 25;
  if (data[scriptQuestion] !== undefined) return scriptQuestion;
  let bestMatch = null;
  let bestMatchLen = 0;
  for (const dataKey in data) {
    let i = 0;
    while (i < scriptQuestion.length && i < dataKey.length && scriptQuestion[i] === dataKey[i]) i++;
    if (i > bestMatchLen && i >= minPrefix) {
      bestMatchLen = i;
      bestMatch = dataKey;
    }
  }
  return bestMatch;
}
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const data = e.namedValues;
  let score = 0;

  const expAnswer = String((data["Total years of work experience in [ Role ]  "] || [""])[0]).trim();
  if (expAnswer === "Less than 3 years") score += 0;
  else if (expAnswer === "3-7 years") score += 10;
  else if (expAnswer === "8 years") score += 10;
  else if (expAnswer === "more than 8 years") score += 7;

  const notice = String((data["NOTICE PERIOD"] || [""])[0]).trim();
  if (notice.indexOf("15") > -1 || notice.indexOf("Immediate") > -1) score += 10;
  else if (notice.indexOf("30") > -1) score += 7;
  else if (notice.indexOf("50") > -1) score += 4;

   const skills = String((data["SKILLS  (select all you have used  )"] || [""])[0]).trim();
  CONFIG.SKILLS_MUST_HAVE.forEach(s => {
    if (skills.indexOf(s) !== -1) score += 5;
  });

  

 for (const question in CONFIG.CORRECT_ANSWERS) {
  const matchedKey = findBestKeyMatch(data, question);
  if (matchedKey) {
    const answer = String((data[matchedKey] || [""])[0]).trim();
    if (answer === CONFIG.CORRECT_ANSWERS[question]) score +=4;
  }
}


  const bucket = score >= CONFIG.SHORTLIST_THRESHOLD ? "SHORTLIST"
               : score >= CONFIG.REVIEW_THRESHOLD ? "REVIEW" : "REJECT";

  sheet.getRange(row, 30).setValue(score);
  sheet.getRange(row, 31).setValue(bucket);

  const candidateEmail = String((data["EMAIL"] || [""])[0]).trim();
  const name = String((data["FULL NAME"] || ["Candidate"])[0]).trim();

  if (candidateEmail) {
    if (bucket === "SHORTLIST") {
      GmailApp.sendEmail(candidateEmail,
        "You have been shortlisted — [ Role ]   role at abc@company Pvt. Ltd.",
        "Hi " + name + ",\n\n" +
        "Congratulations! You have been shortlisted for the [ Role ]   role at abc@company Pvt. Ltd.\n\n" +
        "WHAT HAPPENS NEXT:\n" +
        "1. Our HR team reviews today's shortlist this evening.\n" +
        "2. You will receive a calendar invite with your interview slot.\n" +
        "3. The interview is a 15-30 minute phone call interview.\n\n" +
        "NO ACTION NEEDED FROM YOU RIGHT NOW.\n" +
        "Please do NOT reply to this email — it is automated.\n\n" +
        "If you have an URGENT query (e.g. you cannot attend any of the slots you picked), email [ HR_EMAIL ] directly.\n\n" +
        "Looking forward to speaking with you,\n" +
        "abc@company Hiring Team\n" +
        "abc@company");
    } else if (bucket === "REVIEW") {
      GmailApp.sendEmail(candidateEmail,
        "Application received — abc@company Pvt. Ltd.",
        "Hi " + name + ",\n\n" +
        "Thank you for applying to the [ Role ]   role at abc@company Pvt. Ltd.\n\n" +
        "WHAT HAPPENS NEXT:\n" +
        "Your application is being reviewed by our HR team.\n" +
        "You will hear back from us within 3-5 working days.\n\n" +
        "NO ACTION NEEDED FROM YOU RIGHT NOW.\n" +
        "Please do NOT reply to this email — it is automated.\n\n" +
        "For urgent queries only, contact [ HR_EMAIL ].\n\n" +
        "Best regards,\n" +
        "abc@company Hiring Team");
    } else if (bucket === "REJECT") {
      GmailApp.sendEmail(candidateEmail,
        "Thank you for applying — abc@company Pvt. Ltd.",
        "Hi " + name + ",\n\n" +
        "Thank you for your interest in the [ Role ]   role at abc@company Pvt. Ltd.\n\n" +
        "After reviewing your application, we have decided to move forward with profiles that more closely match our current needs. We will keep your profile in our database for future openings that may be a better fit.\n\n" +
        "We genuinely appreciate the time you invested in your application.\n\n" +
        "Please do NOT reply to this email — it is automated.\n\n" +
        "Best wishes,\n" +
        "abc@company Hiring Team");
    }
  }

 
}

function sendDailyDigest() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const tsCol = 0;
  const nameCol = headers.indexOf("FULL NAME");
  const phoneCol = headers.indexOf("PHONE NUMBER");
  const emailCol = headers.indexOf("EMAIL");
  const resumeCol = headers.indexOf("RESUME");
  const ctcCol = headers.indexOf("EXPECTED CTC (in LPA)");
  const noticeCol = headers.indexOf("NOTICE PERIOD");
  const locCol = headers.indexOf("LOCATION");
  const slotsCol = headers.indexOf("Available interview slots — next week");
  const scoreCol = headers.indexOf("SCORE");
  const bucketCol = headers.indexOf("BUCKET");
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayStr = Utilities.formatDate(today, "Asia/Kolkata", "dd MMM yyyy");
  
  const todayRows = data.slice(1).filter(r => {
    const ts = new Date(r[tsCol]);
    ts.setHours(0, 0, 0, 0);
    return ts.getTime() === today.getTime();
  });
  
  const shortlist = todayRows.filter(r => r[bucketCol] === "SHORTLIST").sort((a, b) => b[scoreCol] - a[scoreCol]);
  const reviewlist = todayRows.filter(r => r[bucketCol] === "REVIEW").sort((a, b) => b[scoreCol] - a[scoreCol]);
  
  const subject = "Daily Digest — " + shortlist.length + " SHORTLIST + " + reviewlist.length + " REVIEW — " + todayStr + " (" + CONFIG.ROLE_NAME + ")";
  
  // Plain text fallback
  let plainBody = "Daily Digest — " + CONFIG.ROLE_NAME + "\nDate: " + todayStr + "\n";
  plainBody += "SHORTLIST: " + shortlist.length + " | REVIEW: " + reviewlist.length + "\n\n";
  
   //HTML body
  let htmlBody = "<div style='font-family:Arial,sans-serif;'>";
  htmlBody += "<h2 style='color:#0a66c2;margin-bottom:5px;'>Daily Digest — " + CONFIG.ROLE_NAME + "</h2>";
  htmlBody += "<p style='color:#555;margin-top:0;'><b>Date:</b> " + todayStr + " &nbsp;|&nbsp; <b>SHORTLIST:</b> " + shortlist.length + " &nbsp;|&nbsp; <b>REVIEW:</b> " + reviewlist.length + "</p>";
  
  // ===== SECTION 1 — SHORTLIST =====
  htmlBody += "<h3 style='color:#0f7b3a;margin-top:30px;border-bottom:2px solid #0f7b3a;padding-bottom:5px;'>Section 1 — SHORTLIST (" + shortlist.length + ")</h3>";
  
  if (shortlist.length === 0) {
    htmlBody += "<p style='color:#888;'>No shortlisted candidates today.</p>";
  } else {
    htmlBody += buildCandidateTable(shortlist, nameCol, phoneCol, emailCol, ctcCol, noticeCol, locCol, slotsCol, scoreCol, resumeCol, "#0f7b3a", "#d4edda");
  }
  
  // ===== SECTION 2 — REVIEW =====
  htmlBody += "<h3 style='color:#b8860b;margin-top:30px;border-bottom:2px solid #b8860b;padding-bottom:5px;'> Section 2 — REVIEW (" + reviewlist.length + ")</h3>";
  htmlBody += "<p style='color:#666;font-size:12px;'>These need your manual triage — move to shortlist or reject.</p>";
  
  if (reviewlist.length === 0) {
    htmlBody += "<p style='color:#888;'>No review-bucket candidates today.</p>";
  } else {
    htmlBody += buildCandidateTable(reviewlist, nameCol, phoneCol, emailCol, ctcCol, noticeCol, locCol, slotsCol, scoreCol, resumeCol, "#b8860b", "#fff3cd");
  }
  
  htmlBody += "<p style='margin-top:30px;font-size:12px;color:#555;'><b>Next step:</b> Open Claude Desktop and say <i>\"Book interviews for today's SHORTLIST\"</i> to auto-schedule calendar invites.</p>";
  htmlBody += "<p style='color:#888;font-size:11px;margin-top:30px;'>— company Hiring Bot | Auto-generated 9 PM digest</p>";
  htmlBody += "</div>";
  
  // Plain text rows
  if (shortlist.length > 0) {
    plainBody += "\n=== SHORTLIST ===\n";
    shortlist.forEach((r, i) => {
      plainBody += (i + 1) + ". " + r[nameCol] + " | " + r[scoreCol] + "/120 | " + r[phoneCol] + "\n";
    });
  }
  if (reviewlist.length > 0) {
    plainBody += "\n=== REVIEW ===\n";
    reviewlist.forEach((r, i) => {
      plainBody += (i + 1) + ". " + r[nameCol] + " | " + r[scoreCol] + "/120 | " + r[phoneCol] + "\n";
    });
  }
  
  GmailApp.sendEmail(CONFIG.HR_EMAIL, subject, plainBody, {
    htmlBody: htmlBody,
    name: "company Hiring Bot"
  });
}

function buildCandidateTable(rows, nameCol, phoneCol, emailCol, ctcCol, noticeCol, locCol, slotsCol, scoreCol, resumeCol, headerColor, scoreBgColor) {
  let html = "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse:collapse;font-size:13px;width:100%;'>";
  html += "<thead><tr style='background-color:" + headerColor + ";color:white;'>";
  html += "<th>#</th><th>Name</th><th>Score</th><th>Phone</th><th>Email</th><th>Expected CTC</th><th>Notice</th><th>Location</th><th>Preferred Slot</th><th>Resume</th>";
  html += "</tr></thead><tbody>";
  
  rows.forEach((r, i) => {
    const rowColor = i % 2 === 0 ? "#ffffff" : "#fafafa";
    html += "<tr style='background-color:" + rowColor + ";'>";
    html += "<td>" + (i + 1) + "</td>";
    html += "<td><b>" + r[nameCol] + "</b></td>";
    html += "<td style='background-color:" + scoreBgColor + ";text-align:center;font-weight:bold;'>" + r[scoreCol] + "/120</td>";
    html += "<td>" + r[phoneCol] + "</td>";
    html += "<td>" + r[emailCol] + "</td>";
    html += "<td>₹" + r[ctcCol] + " LPA</td>";
    html += "<td>" + r[noticeCol] + "</td>";
    html += "<td>" + r[locCol] + "</td>";
    html += "<td>" + r[slotsCol] + "</td>";
    html += "<td><a href='" + r[resumeCol] + "' target='_blank'>View</a></td>";
    html += "</tr>";
  });
  
  html += "</tbody></table>";
  return html;
}

// TRIGGERS — SET THESE UP IN APPS SCRIPT
// 1. onFormSubmit
//    Type: From spreadsheet
//    Event: On form submit
//    Runs: onFormSubmit function
//
// 2. sendDailyDigest
//    Type: Time-driven
//    Hour timer: Every day between 9 PM - 10 PM
//    Runs: sendDailyDigest function
