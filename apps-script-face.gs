// Face login: add these pieces to your existing Google Apps Script (the one at BACKEND).
//
// 1) In the "Registrations" sheet, add a column headed "Face" at the end.
//    In your registration handler, save the template the page sends (data.face):
//        row.push(data.face ? JSON.stringify(data.face) : "");
//
// 2) In doPost, before your other actions:
//        if (data.action === "face_verify") return json_(faceVerify_(data));
//    (json_ is your existing helper that returns ContentService JSON. Use yours if it has another name.)
//
// 3) The real security check: in your exam login action, before you accept the login code, add
//        if (!faceMatches_(findFace_(data.email), data.face)) return json_({ ok: false, error: "Face check failed. Please try again." });
//    and have exam.html send `face` (from EAFace.capture) with the login request.
//    A check that only happens in the browser can be skipped by anyone who edits the page;
//    a check here, on the server, cannot.

var FACE_SHEET = "Registrations";   // the sheet tab that holds registrations
var FACE_MATCH = 0.5;               // keep the same as MATCH in face-auth.js

function faceDistance_(a, b) {
  if (!a || !b || a.length !== b.length) return Infinity;
  var s = 0;
  for (var i = 0; i < a.length; i++) { var t = a[i] - b[i]; s += t * t; }
  return Math.sqrt(s);
}

function faceMatches_(savedJson, given) {
  if (!savedJson || !Array.isArray(given) || given.length !== 128) return false;
  var saved;
  try { saved = JSON.parse(savedJson); } catch (e) { return false; }
  return faceDistance_(saved, given) <= FACE_MATCH;
}

// Finds the saved face template by email or candidate number (Ref)
function findFace_(id) {
  id = String(id || "").trim().toLowerCase();
  if (!id) return "";
  var rows = SpreadsheetApp.getActive().getSheetByName(FACE_SHEET).getDataRange().getValues();
  var head = rows[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var cEmail = head.indexOf("email"), cRef = head.indexOf("ref"), cFace = head.indexOf("face");
  if (cFace < 0) return "";
  for (var r = rows.length - 1; r > 0; r--) {
    if ((cEmail > -1 && String(rows[r][cEmail]).trim().toLowerCase() === id) ||
        (cRef > -1 && String(rows[r][cRef]).trim().toLowerCase() === id)) return rows[r][cFace];
  }
  return "";
}

// Limits guesses: 5 failed face checks per candidate per 15 minutes
function faceVerify_(data) {
  var id = String(data.id || "").trim().toLowerCase();
  var cache = CacheService.getScriptCache(), key = "face-fail:" + id;
  var fails = Number(cache.get(key) || 0);
  if (fails >= 5) return { ok: false, error: "Too many attempts. Please wait 15 minutes and try again." };
  var saved = findFace_(id);
  if (!saved) return { ok: false, error: "No face login is set up for this email or candidate number." };
  if (!faceMatches_(saved, data.face)) {
    cache.put(key, String(fails + 1), 900);
    return { ok: false, error: "Your face did not match the face set up at registration. Please try again in good light." };
  }
  cache.remove(key);
  return { ok: true };
}
