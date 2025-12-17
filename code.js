// ---------- Code.gs ----------
function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  template.logo = getLogoAsBase64();
  return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

function getMasterData() {
  const ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1u6Des7QBsdTbUij1_wgA04iOPVdTI74agE91FPQB8AA/edit");
  const sh = ss.getSheetByName("Attendance_master");
  const data = sh.getDataRange().getValues();
  return data; // array of arrays
}

function getLogoAsBase64() {
  var file = DriveApp.getFileById('14ZuFdSbND36frZ0a3oZaa3hWSeBcgAcO');
  var blob = file.getBlob();
  var data = Utilities.base64Encode(blob.getBytes());
  return 'data:' + blob.getContentType() + ';base64,' + data;
}

/**
 * Detect shift based on current script timezone
 * Shift1: 06:00 - 08:15  -> t in minutes 360 .. 495
 * Shift2: 16:00 - 16:45  -> t in minutes 960 .. 1005
 * Shift3: 00:15 - 01:00  -> t in minutes 15 .. 60
 */
function detectShift() {
  const now = new Date();
  const tz = Session.getScriptTimeZone(); // use script timezone
  const hh = Number(Utilities.formatDate(now, tz, "HH"));
  const mm = Number(Utilities.formatDate(now, tz, "mm"));
  const t = hh * 60 + mm;

  if (t >= 360 && t <= 495) return "Shift1";
  if (t >= 960 && t <= 1005) return "Shift2";
  if (t >= 15 && t <= 60) return "Shift3";
  return ""; // no shift matched
}

function saveToDB(rows) {
  const ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1u6Des7QBsdTbUij1_wgA04iOPVdTI74agE91FPQB8AA/edit");
  const db = ss.getSheetByName("DATABASE");

  // get existing reference list
  const data = db.getDataRange().getValues();
  const existingRefs = data.map(r => r[r.length - 1]); // last column

  // check for duplicates
  for (let i = 0; i < rows.length; i++) {
    let ref = rows[i][rows[i].length - 1];  // reference value
    if (existingRefs.includes(ref)) {
      return "DUPLICATE|" + ref;
    }
  }

  // no duplicates → append rows
  rows.forEach(r => db.appendRow(r));
  return "OK";
}

function sendReportAndMoveDataOutput() {
  const SS_ID = '1u6Des7QBsdTbUij1_wgA04iOPVdTI74agE91FPQB8AA';
  const MAIL_SHEET = 'Output';
  const DB_SHEET = 'DATABASE';
  const DB_CUM_SHEET = 'DB_CUMMULATIVE';
  const EMAIL_TO = 'vasan.ped@delphitvs.com';
  const NO_REPLY = 'noreply@delphitvs.com';

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss"
  );
  const subject = `Output PLAN Vs ACTUAL – ${timestamp}`;

  const ss = SpreadsheetApp.openById(SS_ID);
  const mailSheet = ss.getSheetByName(MAIL_SHEET);
  const dbSheet = ss.getSheetByName(DB_SHEET);
  const dbCumSheet = ss.getSheetByName(DB_CUM_SHEET);

  const mailValues = mailSheet
    .getRange(1, 1, mailSheet.getLastRow(), mailSheet.getLastColumn())
    .getValues();

  let htmlBody = `
    <p><b>Output PLAN Vs ACTUAL</b></p>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">
  `;

  mailValues.forEach((row, rIndex) => {
    htmlBody += "<tr>";
    row.forEach(cell => {
      htmlBody += rIndex === 0
        ? `<th style="background:#f2f2f2;">${cell}</th>`
        : `<td>${cell}</td>`;
    });
    htmlBody += "</tr>";
  });

  htmlBody += `
    </table>
    <br>
    <p style="font-size:12px;color:#666;">
      ⚠️ This is an automated email from DTVS system.<br>
      Please do not reply to this email.
    </p>
  `;

  // PDF export
  const pdfUrl =
    "https://docs.google.com/spreadsheets/d/" + SS_ID + "/export?" +
    "format=pdf" +
    "&portrait=false" +
    "&size=A4" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false" +
    "&gid=" + mailSheet.getSheetId();

  const token = ScriptApp.getOAuthToken();
  const pdf = UrlFetchApp.fetch(pdfUrl, {
    headers: { Authorization: "Bearer " + token }
  }).getBlob().setName(subject + ".pdf");

  // ✅ NO-REPLY MAIL
  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: subject,
    htmlBody: htmlBody,
    attachments: [pdf],
    name: "DTVS No-Reply",
    replyTo: NO_REPLY
  });


}

function sendReportAndMoveDataOutputFinal() {
  const SS_ID = '1u6Des7QBsdTbUij1_wgA04iOPVdTI74agE91FPQB8AA';
  const MAIL_SHEET = 'Output';
  const DB_SHEET = 'DATABASE';
  const DB_CUM_SHEET = 'DB_CUMMULATIVE';
  const EMAIL_TO = 'vasan.ped@delphitvs.com';
  const NO_REPLY = 'noreply@delphitvs.com';

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss"
  );
  const subject = `Output PLAN Vs ACTUAL – ${timestamp}`;

  const ss = SpreadsheetApp.openById(SS_ID);
  const mailSheet = ss.getSheetByName(MAIL_SHEET);
  const dbSheet = ss.getSheetByName(DB_SHEET);
  const dbCumSheet = ss.getSheetByName(DB_CUM_SHEET);

  const mailValues = mailSheet
    .getRange(1, 1, mailSheet.getLastRow(), mailSheet.getLastColumn())
    .getValues();

  let htmlBody = `
    <p><b>Output PLAN Vs ACTUAL</b></p>
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">
  `;

  mailValues.forEach((row, rIndex) => {
    htmlBody += "<tr>";
    row.forEach(cell => {
      htmlBody += rIndex === 0
        ? `<th style="background:#f2f2f2;">${cell}</th>`
        : `<td>${cell}</td>`;
    });
    htmlBody += "</tr>";
  });

  htmlBody += `
    </table>
    <br>
    <p style="font-size:12px;color:#666;">
      ⚠️ This is an automated email from DTVS system.<br>
      Please do not reply to this email.
    </p>
  `;

  // PDF export
  const pdfUrl =
    "https://docs.google.com/spreadsheets/d/" + SS_ID + "/export?" +
    "format=pdf" +
    "&portrait=false" +
    "&size=A4" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false" +
    "&gid=" + mailSheet.getSheetId();

  const token = ScriptApp.getOAuthToken();
  const pdf = UrlFetchApp.fetch(pdfUrl, {
    headers: { Authorization: "Bearer " + token }
  }).getBlob().setName(subject + ".pdf");

  // ✅ NO-REPLY MAIL
  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: subject,
    htmlBody: htmlBody,
    attachments: [pdf],
    name: "DTVS No-Reply",
    replyTo: NO_REPLY
  });

  // Move DB data
  const lastRow = dbSheet.getLastRow();
  if (lastRow > 1) {
    const dbRange = dbSheet.getRange(2, 1, lastRow - 1, 11);
    const dbValues = dbRange.getValues();

    if (dbValues.length > 0) {
      const destRow = dbCumSheet.getLastRow() + 1;
      dbCumSheet.getRange(destRow, 1, dbValues.length, 11).setValues(dbValues);
      dbRange.clearContent();
    }
  }

  SpreadsheetApp.getActive()
    .toast("No-Reply mail sent successfully; DATABASE data moved.");
}


/**
 * Run this function ONCE manually to create all triggers
 */
function createDailyMailTriggers() {

  // 🔥 Remove existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    ScriptApp.deleteTrigger(t);
  });

  // --- 08:00 Triggers ---
  createTrigger_("sendReportAndMoveDataService", 8, 0);
  createTrigger_("sendReportAndMoveDataApu", 8, 0);

  // --- 08:15 Triggers ---
  createTrigger_("sendReportAndMoveDataService", 8, 15);
  createTrigger_("sendReportAndMoveDataApu", 8, 15);

  // --- 08:30 Trigger ---
  createTrigger_("sendReportAndMoveDataOutputFinal", 8, 45);

  Logger.log("✅ All daily triggers created successfully");
}

/**
 * Helper to create time-based trigger
 */
function createTrigger_(functionName, hour, minute) {
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .nearMinute(minute)
    .create();
}
