/**
 * Volunteer sign-up backend for the Green Theatres Roadmap page.
 * Bound to a Google Sheet. Receives sign-ups, stores them and emails you.
 * The site only learns which actions are taken: no names leave the sheet.
 * Also counts completed quiz rounds (quiz.html) in a "Quiz" tab: no personal data.
 *
 * SET THIS: the inbox that should receive sign-up emails.
 */
const NOTIFY_EMAIL = 'YOUR-OUTLOOK-ADDRESS@outlook.com';
const SHEET_NAME = 'Signups';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Timestamp', 'Action #', 'Action', 'Name', 'Role', 'Email', 'Note', 'Counts as taken']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(v, max) { return String(v || '').replace(/[<>]/g, '').trim().slice(0, max); }

function quizSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Quiz');
  if (!sh) {
    sh = ss.insertSheet('Quiz');
    sh.appendRow(['Timestamp', 'Mode', 'Score', 'Out of', 'Retry of missed']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// GET ?quiz=1 → { count: 123 }  (completed quiz rounds)
// GET ?list=1 → { taken: [3, 9, ...] }  (action numbers only)
function doGet(e) {
  if (e && e.parameter && e.parameter.quiz) return json_({ count: Math.max(0, quizSheet_().getLastRow() - 1) });
  const rows = sheet_().getDataRange().getValues().slice(1);
  const taken = {};
  rows.forEach(r => { if (String(r[7]).toUpperCase() !== 'NO') taken[r[1]] = true; });  // type NO in column H to free an action
  return json_({ taken: Object.keys(taken).map(Number).sort((a, b) => a - b) });
}

// POST form fields: action, actionTitle, name, role, email, note, website (honeypot)
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return json_({ ok: true });                  // bot filled the hidden field
  if (p.kind === 'quiz') {                                     // anonymous quiz completion
    const mode = ['practice', 'induction'].indexOf(p.mode) >= 0 ? p.mode : 'practice';
    const outOf = Math.min(Math.max(parseInt(p.outOf, 10) || 0, 1), 50);
    const score = Math.min(Math.max(parseInt(p.score, 10) || 0, 0), outOf);
    const qlock = LockService.getScriptLock();
    qlock.waitLock(10000);
    try {
      const qs = quizSheet_();
      qs.appendRow([new Date(), mode, score, outOf, p.retry === '1' ? 'Yes' : '']);
      return json_({ ok: true, count: qs.getLastRow() - 1 });
    } finally {
      qlock.releaseLock();
    }
  }
  const name = clean_(p.name, 60), email = clean_(p.email, 120);
  const action = parseInt(p.action, 10);
  if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !(action >= 1 && action <= 99)) {
    return json_({ ok: false, error: 'invalid' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const title = clean_(p.actionTitle, 200), role = clean_(p.role, 60), note = clean_(p.note, 400);
    sh.appendRow([new Date(), action, title, name, role, email, note, '']);
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      replyTo: email,
      subject: 'Roadmap volunteer: #' + action + ' ' + title,
      body: name + ' has volunteered for roadmap action #' + action + ': ' + title + '\n\n' +
            'Role: ' + (role || 'not given') + '\nEmail: ' + email + '\n' +
            (note ? '\nNote: ' + note + '\n' : '') +
            '\nReply to this email to contact them. The action now shows as Work in progress on the site. To make it available again, type NO in column H of the sheet.'
    });
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}
