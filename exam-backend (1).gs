/*
  Explained Academy Exam Portal backend (Google Apps Script)

  Sheet "Students" columns (row 1 = headings):
    A StudentID | B Name | C Email | D Programme | E FeePaid (GHS) | F OTP (auto)
    G SubjectsOffered (optional: comma-separated; if empty the student gets every subject of the programme)
  Sheet "Questions" columns (row 1 = headings):
    A ID | B Subject | C Year | D Type (Past or Exam) | E Question
    F OptionA | G OptionB | H OptionC | I OptionD | J Answer (A/B/C/D) | K Explanation
*/

const FEES = {
  'Pure Science': 7500, 'Home Economics': 7500, 'General Art': 7000,
  'Business Studies': 6500, 'Visual Art': 6500, 'Technical Studies': 8000,
  'Agricultural Science': 7000
};
// Subjects tested for each programme. Edit this list to add or remove a subject.
const PROGRAMME_SUBJECTS = {
  'Pure Science': ['Aviation and Aerospace Engineering','Engineering','Manufacturing Engineering','Further Mathematics/Mathematics','Additional Mathematics','Chemistry','Robotics','Biology','Geography','Physics','Computing','Information and Communication Technology','Biomedical Science'],
  'Home Economics': ['Foods and Nutrition','Economics','Management-In-Living','Clothing and Textiles','Graphic Design','General Knowledge-In-Art'],
  'General Art': ['Dagaare','Dagbani','Dangme','Ewe','Fante','Ga','Kasem','Nzema','Twi (Akuapem)','Twi (Asante)','Performing Arts','Arabic','Music','French','Spanish','Government','History','Economics','Religious and Moral Education','Literature-in-English','Christian Religious Studies','West African Traditional Religion','Islamic Religious Studies','Computing','Geography','Additional Mathematics'],
  'Business Studies': ['Principles of Cost Accounting','Business Studies - Accounting','Business Management','Further Mathematics/Mathematics','Economics','Financial Accounting','Clerical Office Duties','Computing'],
  'Visual Art': ['Ceramics','Further Mathematics/Mathematics','Basketry','Graphic Design','Jewelry','Leatherwork','Picture Making','Sculpture','General Knowledge-In-Art'],
  'Technical Studies': ['Electronics','Art and Design','Applied Technology (Wood Technology)','Applied Technology (Automotive)','Applied Technology (Metal Technology)','Applied Technology (Building Construction and Wood Technology)','Auto Mechanics','Physical Education and Health','Design and Communication Technology','Applied Technology (Electrical & Electronic Technology)','Woodwork','Applied Technology (Automotive & Metal Technology)','Applied Electricity','Metalwork','Typewriting','Technical Drawing'],
  'Agricultural Science': ['Animal Husbandry','Crop Husbandry and Horticulture','Agriculture Science','Fisheries','Forestry','General Agriculture']
};
const MIN_RATIO = 0.5;     // half of the fees
const EXAM_MINUTES = 60;   // length of a timed exam

const sheet = n => SpreadsheetApp.getActive().getSheetByName(n);
const eligible = (prog, paid) => FEES[prog] && Number(paid) >= FEES[prog] * MIN_RATIO;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Exam Portal')
    .addItem('Set up sheets and subject lists', 'setupSheets')
    .addItem('Generate OTPs for eligible students', 'makeOtps').addToUi();
}

// Subjects a student can be tested on: their own list (column G) or their programme's list.
function subjectsFor(row) {
  const own = String(row[6] || '').split(',').map(x => x.trim()).filter(Boolean);
  return own.length ? own : (PROGRAMME_SUBJECTS[row[3]] || []);
}

// Run once: creates the three sheets, a Subjects reference tab, and dropdowns that stop spelling mistakes.
function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  const mk = (n, h) => { const t = ss.getSheetByName(n) || ss.insertSheet(n); if (!t.getLastRow()) t.appendRow(h); return t; };
  const st = mk('Students', ['StudentID','Name','Email','Programme','FeePaid','OTP','SubjectsOffered (optional)']);
  const q = mk('Questions', ['ID','Subject','Year','Type','Question','A','B','C','D','Answer','Explanation']);
  const ref = mk('Subjects', ['Programme','Subject']);
  if (ref.getLastRow() < 2) {
    const rows = [];
    Object.keys(PROGRAMME_SUBJECTS).forEach(p => PROGRAMME_SUBJECTS[p].forEach(x => rows.push([p, x])));
    ref.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  const dv = list => SpreadsheetApp.newDataValidation().requireValueInList(list, true).build();
  q.getRange('B2:B2000').setDataValidation(dv([...new Set(Object.values(PROGRAMME_SUBJECTS).flat())]));
  q.getRange('D2:D2000').setDataValidation(dv(['Past', 'Exam']));
  st.getRange('D2:D2000').setDataValidation(dv(Object.keys(FEES)));
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
    return { token, name: row[1], subjects: subjectsFor(row), programme: row[3] };
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
