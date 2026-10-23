/**
 * twinly — application form backend.
 * Every request from the site form is written as a new row in this Google Sheet
 * and sent to your Gmail as a notification.
 *
 * Setup: see INSTRUCTIONS.md in the repository.
 */

// Where notifications go. Leave empty to use the Google account that owns the script.
const NOTIFY_EMAIL = '';

const SHEET_NAME = 'Leads';
const COLUMNS = ['Date', 'Name', 'Email', 'Contact', 'Plan', 'Instagram / TikTok', 'Message', 'Page'];

function doPost(e) {
  const p = (e && e.parameter) || {};

  // honeypot: bots fill the hidden "website" field, people don't
  if (p.website) return json_({ ok: true });

  if (!p.name || !p.email) return json_({ ok: false, error: 'Name and email are required' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const row = [
      new Date(),
      clean_(p.name),
      clean_(p.email),
      clean_(p.contact),
      clean_(p.plan),
      clean_(p.social),
      clean_(p.message),
      clean_(p.page),
    ];
    sheet_().appendRow(row);
    notify_(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Opening the Web app URL in a browser shows this, handy to check the deployment works.
function doGet() {
  return json_({ ok: true, message: 'twinly form backend is running' });
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNS);
    sh.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function notify_(row) {
  const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  const [date, name, email, contact, plan, social, message, page] = row;
  const lines = [
    ['Name', name], ['Email', email], ['Contact', contact], ['Plan', plan],
    ['Instagram / TikTok', social], ['Message', message], ['Page', page],
  ];
  const html =
    '<h2 style="margin:0 0 12px">New request from the site</h2>' +
    '<table cellpadding="6" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">' +
    lines.map(([k, v]) => '<tr><td style="color:#888">' + k + '</td><td><b>' + esc_(v || '—') + '</b></td></tr>').join('') +
    '</table>' +
    '<p style="color:#888;font-size:12px">' + date + '</p>';
  MailApp.sendEmail({
    to: to,
    subject: 'New twinly request: ' + name + ' (' + (plan || 'no plan') + ')',
    htmlBody: html,
    replyTo: email,
  });
}

function clean_(v) {
  // a leading = + - @ would make Sheets treat the text as a formula
  return String(v || '').trim().slice(0, 2000).replace(/^[=+\-@]/, "'$&");
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Run this once from the editor (Run ▶ testSetup) to grant permissions and get a test row + email.
function testSetup() {
  doPost({ parameter: { name: 'Test', email: Session.getEffectiveUser().getEmail(), plan: 'Pro ($1,000 / month)', message: 'Test request', page: 'manual test' } });
}
