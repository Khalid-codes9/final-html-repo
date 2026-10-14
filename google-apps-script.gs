/**
 * Explained Academy - registration backend
 * Paste this whole file into the Apps Script editor of a Google Sheet
 * (Extensions > Apps Script), then deploy it as a Web app.
 *
 * For each registration it:
 *  1. issues a unique, sequential candidate number (EA-2026-0001, 0002, ...)
 *  2. saves the candidate's photo in a Google Drive folder
 *  3. adds a row to the "Registrations" sheet
 *  4. emails the candidate a confirmation (if they gave an email)
 *  5. emails you a notice (to the owner of this script)
 */

var SHEET_NAME  = "Registrations";
var FOLDER_NAME = "Explained Academy Photos";
var ACADEMY     = "Explained Academy";
var CLOSES      = new Date(2026, 9, 24, 23, 59, 59); // 24 Oct 2026, end of day
var ADMIN_EMAIL = "";   // optional: put the email that should get notices; blank = owner of this script

var HEADERS = ["Date", "Candidate No.", "Name", "Gender", "Phone", "Email", "Category",
               "Programme", "Fee", "Core subjects", "Electives", "Notes", "Photo link"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (new Date() > CLOSES) return reply({ ok: false, error: "closed" });

    var d = JSON.parse(e.postData.contents);
    var name = clean(d.name, 120), phone = clean(d.phone, 20);
    if (!name || !/^0\d{9}$/.test(phone)) return reply({ ok: false, error: "invalid" });

    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty("counter") || 0) + 1;
    props.setProperty("counter", String(n));
    var ref = "EA-" + new Date().getFullYear() + "-" + ("0000" + n).slice(-4);

    var photoUrl = savePhoto(d.photo, ref);

    var sheet = getSheet();
    var core = (d.core || []).map(function (x) { return clean(x, 60); }).join(", ");
    var electives = (d.electives || []).map(function (x) { return clean(x, 80); }).join(", ");
    sheet.appendRow([new Date(), ref, name, clean(d.gender, 10), phone, clean(d.email, 120),
      clean(d.type, 40), clean(d.programme, 60), clean(d.fee, 30), core, electives,
      clean(d.notes, 500), photoUrl]);

    var email = clean(d.email, 120);
    if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      MailApp.sendEmail(email, ACADEMY + " registration received (" + ref + ")",
        "Hello " + name + ",\n\nThank you for registering with " + ACADEMY + ".\n\n" +
        "Candidate number: " + ref + "\nProgramme: " + clean(d.programme, 60) +
        (d.fee ? "\nFee: " + clean(d.fee, 30) : "") +
        "\nCore subjects: " + core + "\nElectives: " + electives +
        "\n\nPlease keep this number and bring your printed slip to the academy office.\n\n" + ACADEMY);
    }
    MailApp.sendEmail(ADMIN_EMAIL || Session.getEffectiveUser().getEmail(),
      "New registration: " + name + " (" + ref + ")",
      name + " | " + phone + " | " + clean(d.programme, 60) + "\nElectives: " + electives);

    return reply({ ok: true, ref: ref });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

function savePhoto(dataUrl, ref) {
  var m = String(dataUrl || "").match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+\/=]+)$/);
  if (!m) return "";
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], ref + ".jpg");
  return folder.createFile(blob).getUrl();
}

// Trim, limit length, and stop spreadsheet formulas (=, +, -, @) being run
function clean(v, max) {
  var t = String(v == null ? "" : v).trim().slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
