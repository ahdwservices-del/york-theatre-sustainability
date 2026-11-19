/**
 * Volunteer sign-up backend for the Green Theatres Roadmap page.
 * Bound to a Google Sheet. Receives sign-ups, stores them, emails you,
 * and returns public names (first name + surname initial) for the site.
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
    sh.appendRow(['Timestamp', 'Action #', 'Action', 'Name', 'Role', 'Email', 'Note', 'Show on site', 'Approved']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function publicName_(name) {
  const p = String(name).trim().split(/\s+/);
  return p[0] + (p.length > 1 ? ' ' + p[p.length - 1].charAt(0).toUpperCase() + '.' : '');
}

function clean_(v, max) { return String(v || '').replace(/[<>]/g, '').trim().slice(0, max); }

// GET ?list=1 → { volunteers: { "3": ["Jane D."], ... } }
function doGet(e) {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const out = {};
  rows.forEach(r => {
    const show = String(r[7]).toLowerCase() === 'yes';
    const approved = String(r[8]).toUpperCase() !== 'NO';   // type NO in column I to hide someone
    if (show && approved) (out[r[1]] = out[r[1]] || []).push(publicName_(r[3]));
  });
  return json_({ volunteers: out });
}

// POST form fields: action, actionTitle, name, role, email, note, show, website (honeypot)
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return json_({ ok: true });                  // bot filled the hidden field
  const name = clean_(p.name, 60), email = clean_(p.email, 120);
  const action = parseInt(p.action, 10);
  if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !(action >= 1 && action <= 41)) {
    return json_({ ok: false, error: 'invalid' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const title = clean_(p.actionTitle, 200), role = clean_(p.role, 60), note = clean_(p.note, 400);
    const show = p.show === 'yes' ? 'yes' : 'no';
    sh.appendRow([new Date(), action, title, name, role, email, note, show, '']);
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      replyTo: email,
      subject: 'Roadmap volunteer: #' + action + ' ' + title,
      body: name + ' has volunteered for roadmap action #' + action + ': ' + title + '\n\n' +
            'Role: ' + (role || 'not given') + '\nEmail: ' + email + '\n' +
            'Show name on site: ' + show + '\n' + (note ? '\nNote: ' + note + '\n' : '') +
            '\nReply to this email to contact them. To hide them from the site, type NO in column I of the sheet.'
    });
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}
