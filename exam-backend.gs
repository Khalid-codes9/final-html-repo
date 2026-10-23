/*
  Explained Academy Exam Portal backend (Google Apps Script)

  Sheet "Students" columns (row 1 = headings):
    A StudentID | B Name | C Email | D Programme | E FeePaid (GHS) | F OTP (auto)
  Sheet "Questions" columns (row 1 = headings):
    A ID | B Subject | C Year | D Type (Past or Exam) | E Question
    F OptionA | G OptionB | H OptionC | I OptionD | J Answer (A/B/C/D) | K Explanation
*/

const FEES = {
  'Pure Science': 7500, 'Home Economics': 7500, 'General Art': 7000,
  'Business Studies': 6500, 'Visual Art': 6500, 'Technical Studies': 8000,
  'Agricultural Science': 7000
};
const MIN_RATIO = 0.5;     // half of the fees
const EXAM_MINUTES = 60;   // length of a timed exam

const sheet = n => SpreadsheetApp.getActive().getSheetByName(n);
const eligible = (prog, paid) => FEES[prog] && Number(paid) >= FEES[prog] * MIN_RATIO;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Exam Portal')
    .addItem('Generate OTPs for eligible students', 'makeOtps').addToUi();
}

// Run from the sheet menu after you update the FeePaid column.
function makeOtps() {
  const s = sheet('Students'), d = s.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < d.length; i++) {
    const [id, name, email, prog, paid, otp] = d[i];
    if (id && !otp && eligible(prog, paid)) {
      const code = String(Math.floor(100000 + Math.random() * 900000));
      s.getRange(i + 1, 6).setValue(code);
      if (email) MailApp.sendEmail(email, 'Your Explained Academy exam OTP',
        'Hello ' + name + ',\n\nStudent ID: ' + id + '\nOTP: ' + code +
        '\n\nUse these to log in to the Exam Portal.');
      n++;
    }
  }
  SpreadsheetApp.getUi().alert(n + ' new OTP(s) generated.');
}

function doPost(e) {
  let out;
  try { out = handle(JSON.parse(e.postData.contents)); }
  catch (err) { out = { error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function handle(r) {
  const cache = CacheService.getScriptCache();
  const all = sheet('Questions').getDataRange().getValues().slice(1);

  if (r.action === 'login') {
    const row = sheet('Students').getDataRange().getValues().slice(1)
      .find(x => String(x[0]).trim().toLowerCase() === String(r.id).trim().toLowerCase());
    if (!row || !row[5] || String(row[5]) !== String(r.otp).trim())
      throw new Error('Wrong Student ID or OTP.');
    if (!eligible(row[3], row[4]))
      throw new Error('Your fees payment is below half. Please see the school office.');
    const token = Utilities.getUuid();
    cache.put(token, String(row[0]), 10800); // 3 hours
    return { token, name: row[1], subjects: [...new Set(all.filter(x => x[1]).map(x => x[1]))] };
  }

  if (!r.token || !cache.get(r.token)) throw new Error('Session expired. Please log in again.');
  const qs = all.filter(x => x[1] === r.subject && x[3] === r.type);

  if (r.action === 'questions')
    return { minutes: EXAM_MINUTES,
      questions: qs.map(x => ({ id: x[0], year: x[2], q: x[4], o: [x[5], x[6], x[7], x[8]] })) };

  if (r.action === 'submit') {
    let score = 0;
    const review = qs.map(x => {
      const correct = String(x[9]).trim().toUpperCase(), yours = (r.answers || {})[x[0]] || '';
      if (yours === correct) score++;
      return { q: x[4], correct, yours, why: x[10] };
    });
    // Exam mode: only the score is returned, so answers are not leaked.
    return { score, total: qs.length, review: r.type === 'Exam' ? [] : review };
  }
  throw new Error('Unknown action.');
}
