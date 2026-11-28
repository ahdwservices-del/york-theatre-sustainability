/**
 * Volunteer sign-up backend for the Green Theatres Roadmap page.
 * Bound to a Google Sheet. Receives sign-ups, stores them and emails you.
 * The site only learns which actions are taken: no names leave the sheet.
 * Also counts completed quiz rounds (quiz.html) in a "Quiz" tab: no personal data.
 * Also stores anonymous staff feedback (feedback.html) in a "Feedback" tab: date only, no time, no identity.
 * Also logs the drug waste audit (drug-waste.html) in a "Drug waste" tab: site, theatre, drug, time. No identity.
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

function feedbackSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Feedback');
  if (!sh) {
    sh = ss.insertSheet('Feedback');
    sh.appendRow(['Date', 'Topic', 'Where', 'What got in the way', 'What would make it easier']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Drug waste audit (drug-waste.html): one row per drug (and propofol/remifentanil mg band) per entry. No personal data.
const DW_SITES = {
  'York': ['1','2','3','4','5','6','7','8','9','10','11','12','DC1','DC2','DC3','DC4','DC5','DC6','Eye 1','Eye 2'],
  'Scarborough': ['1','2','3','4','5','6'],
  'Bridlington': ['1','2']
};
const DW_DRUGS = ['Propofol', 'Remifentanil', 'Ephedrine', 'Metaraminol', 'Atropine', 'Glycopyrrolate', 'Rocuronium', 'Atracurium', 'Suxamethonium'];
const DW_GROUP = { Propofol: 'Induction agent', Remifentanil: 'Opioid', Ephedrine: 'Vasopressor', Metaraminol: 'Vasopressor',
                   Atropine: 'Anticholinergic', Glycopyrrolate: 'Anticholinergic',
                   Rocuronium: 'Neuromuscular blocker', Atracurium: 'Neuromuscular blocker', Suxamethonium: 'Neuromuscular blocker' };
const DW_AMOUNTS = { Propofol: ['<50 mg', '50–100 mg', '>100 mg'],
                     Remifentanil: ['<0.5 mg', '0.5–1 mg', '1–1.5 mg', '>1.5 mg'] };
const DW_REASONS = ['Drawn up "just in case"', 'Case cancelled or changed', 'Contaminated or dropped', 'Expired / time limit', 'Not required', 'Vial bigger than needed', 'Other'];
const DW_HEADERS = ['Received', 'Logged at', 'Date', 'Time', 'Weekday', 'Site', 'Theatre', 'Drug', 'Drug group',
                    'Amount (mg)', 'Syringes', 'Reason', 'Entry ID', 'Given', 'Form version'];
const DW_ID_COL = 13;                                               // 'Entry ID' column (1-based); 'Given' added after it
const DW_GIVEN = ['None given', 'Part used'];

// ?test=1 on the page sends test:'1' -> separate TEST tabs, so devices can be checked without touching real data
function drugSheet_(test) {
  const name = 'Drug waste' + (test ? ' TEST' : '');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(DW_HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

// End-of-list check (drug-waste.html#end-of-list): the denominator. One row per list.
// The five emergency drugs are counted ONLY here; the waste log covers propofol, remifentanil,
// rocuronium, atracurium and other drugs, so nothing is counted twice.
const DL_DRUGS = ['Ephedrine', 'Metaraminol', 'Atropine', 'Glycopyrrolate', 'Suxamethonium'];
const DL_HEADERS = ['Received', 'Logged at', 'Date', 'Time', 'Weekday', 'Site', 'Theatre', 'Cases']
  .concat(...DL_DRUGS.map(d => [d + ' drawn up', d + ' used', d + ' never given']))
  .concat(['Entry ID', 'Never-given reason', 'Form version']);
const DL_ID_COL = 8 + DL_DRUGS.length * 3 + 1;                     // 'Entry ID' (1-based); reason added after it
const DL_REASONS = ['Drawn up "just in case"', 'Case cancelled or changed', 'Contaminated or dropped', 'Other'];

function listSheet_(test) {
  const name = 'Drug waste lists' + (test ? ' TEST' : '');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(DL_HEADERS); sh.setFrozenRows(1); }
  return sh;
}

function drugList_(p) {
  const sheetSafe = s => /^[=+\-@]/.test(s) ? "'" + s : s;
  const site = Object.prototype.hasOwnProperty.call(DW_SITES, p.site) ? p.site : '';
  if (!site) return json_({ ok: false, error: 'invalid' });
  let theatre = clean_(p.theatre, 40);
  if (DW_SITES[site].indexOf(theatre) < 0) theatre = 'Other: ' + sheetSafe(theatre);
  const cases = Math.min(Math.max(parseInt(p.cases, 10) || 0, 0), 40);
  let counts;
  try { counts = JSON.parse(p.counts || '{}'); } catch (x) { return json_({ ok: false, error: 'invalid' }); }
  if (!counts || typeof counts !== 'object') return json_({ ok: false, error: 'invalid' });
  const id = clean_(p.id, 20).replace(/[^a-z0-9]/gi, '');
  const now = new Date();
  let at = new Date(p.ts);
  if (isNaN(at) || at > new Date(now.getTime() + 5 * 6e4) || at < new Date(now.getTime() - 14 * 864e5)) at = now;
  const tz = Session.getScriptTimeZone();
  const row = [now, at, Utilities.formatDate(at, tz, 'yyyy-MM-dd'), Utilities.formatDate(at, tz, 'HH:mm'),
               Utilities.formatDate(at, tz, 'EEE'), site, theatre, cases];
  DL_DRUGS.forEach(d => {
    const v = Array.isArray(counts[d]) ? counts[d] : [0, 0];
    const drawn = Math.min(Math.max(parseInt(v[0], 10) || 0, 0), 30);
    const used = Math.min(Math.max(parseInt(v[1], 10) || 0, 0), drawn);
    row.push(drawn, used, drawn - used);
  });
  row.push(id, DL_REASONS.indexOf(p.reason) >= 0 ? p.reason : '', clean_(p.v, 10));

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = listSheet_(p.test === '1');
    if (id && sh.getLastRow() > 1) {
      const dup = sh.getRange(2, DL_ID_COL, sh.getLastRow() - 1, 1)
                    .createTextFinder(id).matchEntireCell(true).findNext();
      if (dup) return json_({ ok: true, duplicate: true });
    }
    sh.getRange(sh.getLastRow() + 1, 1, 1, DL_HEADERS.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

function drugWaste_(p) {
  const sheetSafe = s => /^[=+\-@]/.test(s) ? "'" + s : s;          // no formula injection
  const site = Object.prototype.hasOwnProperty.call(DW_SITES, p.site) ? p.site : '';
  if (!site) return json_({ ok: false, error: 'invalid' });
  let theatre = clean_(p.theatre, 40);
  if (DW_SITES[site].indexOf(theatre) < 0) theatre = 'Other: ' + sheetSafe(theatre);
  let items;
  try { items = JSON.parse(p.items || '[]'); } catch (x) { return json_({ ok: false, error: 'invalid' }); }
  if (!Array.isArray(items) || !items.length || items.length > 12) return json_({ ok: false, error: 'invalid' });
  const id = clean_(p.id, 20).replace(/[^a-z0-9]/gi, '');

  // time the button was pressed (allow up to 14 days of offline queueing); otherwise time received
  const now = new Date();
  let at = new Date(p.ts);
  if (isNaN(at) || at > new Date(now.getTime() + 5 * 6e4) || at < new Date(now.getTime() - 14 * 864e5)) at = now;
  const tz = Session.getScriptTimeZone();

  const rows = [];
  items.forEach(it => {
    const n = Math.min(Math.max(parseInt(it && it.n, 10) || 0, 0), 50);
    if (!n) return;
    let drug = clean_(it.drug, 50);
    if (DW_DRUGS.indexOf(drug) < 0) drug = 'Other: ' + sheetSafe(drug.replace(/^Other:\s*/, ''));
    const vol = DW_AMOUNTS[drug] && DW_AMOUNTS[drug].indexOf(it.vol) >= 0 ? it.vol : '';
    rows.push([now, at, Utilities.formatDate(at, tz, 'yyyy-MM-dd'), Utilities.formatDate(at, tz, 'HH:mm'),
               Utilities.formatDate(at, tz, 'EEE'), site, theatre, drug, DW_GROUP[drug] || 'Other', vol, n,
               DW_REASONS.indexOf(p.reason) >= 0 ? p.reason : '', id,
               DW_GIVEN.indexOf(p.given) >= 0 ? p.given : '', clean_(p.v, 10)]);
  });
  if (!rows.length) return json_({ ok: false, error: 'invalid' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = drugSheet_(p.test === '1');
    if (id && sh.getLastRow() > 1) {                                // ignore re-sent entries
      const dup = sh.getRange(2, DW_ID_COL, sh.getLastRow() - 1, 1)
                    .createTextFinder(id).matchEntireCell(true).findNext();
      if (dup) return json_({ ok: true, duplicate: true });
    }
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, DW_HEADERS.length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// ---- Drug waste audit: monitoring and analysis ----

// Run DAILY during the audit (Triggers -> Add trigger -> dailyDrugWasteSummary -> Day timer, 7-8pm).
// Emails yesterday-to-now totals by site and flags patterns that look like errors or fake entries.
function dailyDrugWasteSummary() {
  const since = Date.now() - 864e5;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const w = ss.getSheetByName('Drug waste'), l = ss.getSheetByName('Drug waste lists');
  const wRows = w ? w.getDataRange().getValues().slice(1).filter(r => new Date(r[0]).getTime() >= since) : [];
  const lRows = l ? l.getDataRange().getValues().slice(1).filter(r => new Date(r[0]).getTime() >= since) : [];
  const lines = [], flags = [];
  Object.keys(DW_SITES).forEach(site => {
    const ws = wRows.filter(r => r[5] === site), ls = lRows.filter(r => r[5] === site);
    const syr = ws.reduce((a, r) => a + Number(r[10] || 0), 0);
    let drawn = 0, never = 0;
    ls.forEach(r => DL_DRUGS.forEach((d, i) => { drawn += Number(r[8 + i * 3] || 0); never += Number(r[10 + i * 3] || 0); }));
    lines.push(site + ': ' + ls.length + ' end-of-list checks (' + ls.reduce((a, r) => a + Number(r[7] || 0), 0) + ' patients), ' +
               'emergency syringes ' + never + ' never given of ' + drawn + ' drawn up; ' + syr + ' other syringes logged as waste.');
  });
  // bursts: more than 5 submissions from one theatre within 10 minutes
  const byTheatre = {};
  wRows.concat(lRows).forEach(r => { const k = r[5] + ' / ' + r[6]; (byTheatre[k] = byTheatre[k] || []).push(new Date(r[0]).getTime()); });
  Object.keys(byTheatre).forEach(k => {
    const t = byTheatre[k].sort();
    for (let i = 5; i < t.length; i++) if (t[i] - t[i - 5] < 10 * 6e4) { flags.push('Burst of entries from ' + k); break; }
  });
  lRows.forEach(r => { if (Number(r[7]) > 15) flags.push('List with ' + r[7] + ' patients at ' + r[5] + ' / ' + r[6]); });
  wRows.forEach(r => { if (Number(r[10]) > 10) flags.push(r[10] + ' syringes of ' + r[7] + ' in one entry at ' + r[5] + ' / ' + r[6]); });
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: 'Drug waste audit: last 24 h (' + lRows.length + ' lists, ' + wRows.length + ' waste entries)',
    body: lines.join('\n') + '\n\n' + (flags.length ? 'Check these:\n- ' + flags.join('\n- ') : 'Nothing unusual.') +
          '\n\nFull totals: Drug waste summary tab.'
  });
}

// Run ONCE (select setupDrugWasteSummary -> Run). Builds a 'Drug waste summary' tab of live formulas:
// emergency-drug waste rates by site (end-of-list check), other waste by drug and site, and reasons.
function setupDrugWasteSummary() {
  drugSheet_(false); listSheet_(false);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Drug waste summary') || ss.insertSheet('Drug waste summary');
  sh.clear();
  const col = n => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
  const L = "'Drug waste lists'!", W = "'Drug waste'!", sites = Object.keys(DW_SITES);
  const out = [];
  out.push(['Emergency drugs: end-of-list check', '', '', '', '']);
  out.push(['', ...sites, 'All sites']);
  out.push(['Lists checked', ...sites.map(s => '=COUNTIF(' + L + 'F:F,"' + s + '")'), '=COUNTA(' + L + 'F2:F)']);
  out.push(['Patients', ...sites.map(s => '=SUMIFS(' + L + 'H:H,' + L + 'F:F,"' + s + '")'), '=SUM(' + L + 'H2:H)']);
  DL_DRUGS.forEach((d, i) => {
    const dr = col(8 + i * 3), nv = col(10 + i * 3);
    out.push([d + ': % never given',
      ...sites.map(s => '=IFERROR(SUMIFS(' + L + nv + ':' + nv + ',' + L + 'F:F,"' + s + '")/SUMIFS(' + L + dr + ':' + dr + ',' + L + 'F:F,"' + s + '"),"")'),
      '=IFERROR(SUM(' + L + nv + '2:' + nv + ')/SUM(' + L + dr + '2:' + dr + '),"")']);
    out.push([d + ': never given / drawn up',
      ...sites.map(s => '=SUMIFS(' + L + nv + ':' + nv + ',' + L + 'F:F,"' + s + '")&" / "&SUMIFS(' + L + dr + ':' + dr + ',' + L + 'F:F,"' + s + '")'),
      '=SUM(' + L + nv + '2:' + nv + ')&" / "&SUM(' + L + dr + '2:' + dr + ')']);
  });
  out.push(['', '', '', '', '']);
  out.push(['Other drugs: syringes logged as waste', ...sites, 'All sites']);
  ['Propofol', 'Remifentanil', 'Rocuronium', 'Atracurium'].forEach(d => out.push([d,
    ...sites.map(s => '=SUMIFS(' + W + 'K:K,' + W + 'H:H,"' + d + '",' + W + 'F:F,"' + s + '")'),
    '=SUMIFS(' + W + 'K:K,' + W + 'H:H,"' + d + '")']));
  out.push(['Other drugs (any)', ...sites.map(s => '=SUMIFS(' + W + 'K:K,' + W + 'H:H,"Other*",' + W + 'F:F,"' + s + '")'), '=SUMIFS(' + W + 'K:K,' + W + 'H:H,"Other*")']);
  out.push(['', '', '', '', '']);
  out.push(['Reasons (syringes, waste log)', '', '', '', 'All sites']);
  DW_REASONS.forEach(r => out.push([r, '', '', '', '=SUMIFS(' + W + 'K:K,' + W + 'L:L,"' + r.replace(/"/g, '""') + '")']));
  out.push(['Reasons (lists with never-given emergency drugs)', '', '', '', '']);
  DL_REASONS.forEach(r => out.push([r, '', '', '', '=COUNTIF(' + L + col(DL_HEADERS.length - 2) + ':' + col(DL_HEADERS.length - 2) + ',"' + r.replace(/"/g, '""') + '")']));
  sh.getRange(1, 1, out.length, 5).setValues(out);
  sh.getRange('A:A').setFontWeight('bold');
  sh.setColumnWidth(1, 300);
  out.forEach((r, i) => { if (/% never given/.test(r[0])) sh.getRange(i + 1, 2, 1, 4).setNumberFormat('0%'); });
}

// Optional: run weekly from a time-driven trigger (Triggers → Add trigger → sendFeedbackDigest → Week timer).
// A weekly batch, not one email per message, so the timing of an email can't hint at who sent it.
function sendFeedbackDigest() {
  const rows = feedbackSheet_().getDataRange().getValues().slice(1);
  const since = Date.now() - 7 * 864e5;
  const recent = rows.filter(r => new Date(r[0]).getTime() >= since);
  if (!recent.length) return;
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: 'Theatre sustainability feedback: ' + recent.length + ' new this week',
    body: recent.map(r => '[' + r[1] + (r[2] ? ', ' + r[2] : '') + ']\nWhat got in the way: ' + r[3] + (r[4] ? '\nWhat would help: ' + r[4] : '')).join('\n\n') +
          '\n\nAll feedback is in the Feedback tab of the sign-ups sheet.'
  });
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
  if (p.kind === 'drugwaste') return drugWaste_(p);
  if (p.kind === 'drugwaste_list') return drugList_(p);         // end-of-list counts: no personal data           // drug waste audit: no personal data
  if (p.kind === 'feedback') {                                 // anonymous feedback: date only
    const topics = ['Waste and bins', 'Anaesthetic gases', 'Drugs and propofol', 'Reusable kit', 'Energy and lights', 'Something else'];
    const places = ['Anaesthetic room', 'Theatre', 'Recovery', 'Somewhere else'];
    const issue = clean_(p.issue, 1500);
    if (issue.length < 5) return json_({ ok: false, error: 'invalid' });
    const flock = LockService.getScriptLock();
    flock.waitLock(10000);
    try {
      feedbackSheet_().appendRow([
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        topics.indexOf(p.topic) >= 0 ? p.topic : 'Something else',
        places.indexOf(p.where) >= 0 ? p.where : '',
        issue, clean_(p.idea, 800)
      ]);
    } finally {
      flock.releaseLock();
    }
    return json_({ ok: true });
  }
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
