/**
 * Gmail Attachments -> Drive + Sheets + Daily Digest (Google Apps Script)
 *
 * Every 15 minutes:  find new emails with attachments, save each attachment to
 *                    Drive (one subfolder per month), log it in a Google Sheet
 *                    with sender, subject, PO number and a Drive link, and
 *                    label the thread so it is not processed twice.
 * Every evening:     email a digest of what arrived today, grouped by sender.
 *
 * Setup: paste into a new Apps Script project, fill CONFIG, run setupTriggers()
 * once and approve the Gmail / Drive / Sheets permissions.
 */

const CONFIG = {
  QUERY: 'has:attachment -in:chats newer_than:7d',
  PROCESSED_LABEL: 'auto-filed',
  DRIVE_FOLDER_NAME: 'Email Attachments',
  SHEET_ID: '',            // leave blank to create "Attachment Log" on first run
  SHEET_NAME: 'Log',
  DIGEST_TO: '',           // leave blank to send to yourself
  DIGEST_HOUR: 18,         // 0-23, in the script's time zone
  ALLOWED_EXT: ['pdf', 'xlsx', 'xls', 'csv', 'docx', 'doc'],
  MAX_THREADS: 50,
};

const HEADERS = ['Received', 'From', 'Subject', 'PO number', 'File', 'Size (KB)', 'Drive link', 'Message ID'];

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------
function processInbox() {
  const label = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL) ||
                GmailApp.createLabel(CONFIG.PROCESSED_LABEL);
  const query = `${CONFIG.QUERY} -label:${CONFIG.PROCESSED_LABEL}`;
  const threads = GmailApp.search(query, 0, CONFIG.MAX_THREADS);
  if (threads.length === 0) return;

  const root = getOrCreateFolder_(CONFIG.DRIVE_FOLDER_NAME);
  const sheet = getLogSheet_();
  const seen = new Set(sheet.getLastRow() > 1
    ? sheet.getRange(2, HEADERS.length, sheet.getLastRow() - 1, 1).getValues().flat()
    : []);
  const tz = Session.getScriptTimeZone();
  const rows = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const id = msg.getId();
      if (seen.has(id)) return;
      const received = msg.getDate();
      const from = msg.getFrom();
      const subject = msg.getSubject();
      const po = extractPoNumber(subject + ' ' + msg.getPlainBody().slice(0, 5000));
      const month = Utilities.formatDate(received, tz, 'yyyy-MM');
      const day = Utilities.formatDate(received, tz, 'yyyy-MM-dd');
      const folder = getOrCreateSubfolder_(root, month);

      msg.getAttachments({ includeInlineImages: false })
        .filter(att => isAllowed(att.getName(), CONFIG.ALLOWED_EXT))
        .forEach(att => {
          const file = folder.createFile(att.copyBlob())
            .setName(buildFileName(day, from, att.getName()));
          rows.push([received, from, subject, po, file.getName(),
                     Math.round(att.getSize() / 1024), file.getUrl(), id]);
        });
      seen.add(id);
    });
    thread.addLabel(label);
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
  }
}

function sendDailyDigest() {
  const sheet = getLogSheet_();
  if (sheet.getLastRow() < 2) return;
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues()
    .filter(r => Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd') === today)
    .map(r => ({ from: r[1], subject: r[2], po: r[3], file: r[4], link: r[6] }));

  const digest = buildDigest(rows, today);
  const to = CONFIG.DIGEST_TO || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({ to, subject: digest.subject, body: digest.text });
}

function setupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['processInbox', 'sendDailyDigest'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processInbox').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('sendDailyDigest').timeBased().atHour(CONFIG.DIGEST_HOUR).everyDays(1).create();
  processInbox();   // first run now, so permissions are granted while you watch
}

// ---------------------------------------------------------------------------
// Google services helpers
// ---------------------------------------------------------------------------
function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function getOrCreateSubfolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getLogSheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = CONFIG.SHEET_ID || props.getProperty('SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('Attachment Log');
  if (!id) props.setProperty('SHEET_ID', ss.getId());
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// Pure helpers (no Google services - unit-tested with Node, see tests/)
// ---------------------------------------------------------------------------
/** "PO# 45-1180", "P.O. Number: A12345", "purchase order 778812" -> the ID, else "". */
function extractPoNumber(text) {
  const m = String(text || '').match(
    /\b(?:P\.?\s?O\.?|purchase\s+order)\s*(?:#|no\.?|number)?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\-\/]{3,})/i);
  return m ? m[1].replace(/[\-\/]+$/, '').toUpperCase() : '';
}

/** "Jane Doe <jane@acme.com>" -> "Jane Doe"; "jane@acme.com" -> "jane". */
function senderName(from) {
  const s = String(from || '').trim();
  const named = s.match(/^"?([^"<]+?)"?\s*<[^>]+>$/);
  if (named) return named[1].trim();
  const email = s.match(/<?([^<>\s@]+)@[^<>\s]+>?/);
  return email ? email[1] : s;
}

function sanitizeFilename(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120) || 'attachment';
}

function extOf(name) {
  const m = String(name || '').match(/\.([A-Za-z0-9]+)$/);
  return m ? m[1].toLowerCase() : '';
}

function isAllowed(name, allowed) {
  return allowed.includes(extOf(name));
}

/** "2026-11-18", "Jane Doe <j@x.com>", "Quote.pdf" -> "2026-11-18 Jane Doe - Quote.pdf" */
function buildFileName(day, from, original) {
  return sanitizeFilename(`${day} ${senderName(from)} - ${original}`);
}

/** rows: [{from, subject, po, file, link}] -> {subject, text} */
function buildDigest(rows, dayLabel) {
  if (!rows.length) {
    return { subject: `Attachments ${dayLabel}: none`, text: `No new attachments on ${dayLabel}.` };
  }
  const bySender = new Map();
  rows.forEach(r => {
    const key = senderName(r.from);
    if (!bySender.has(key)) bySender.set(key, []);
    bySender.get(key).push(r);
  });
  const pos = [...new Set(rows.map(r => r.po).filter(Boolean))];
  const lines = [`${rows.length} attachment(s) from ${bySender.size} sender(s) on ${dayLabel}.`];
  if (pos.length) lines.push(`PO numbers: ${pos.join(', ')}`);
  [...bySender.entries()]
    .sort((a, b) => b[1].length - a[1].length)
    .forEach(([sender, list]) => {
      lines.push('', `${sender} (${list.length})`);
      list.forEach(r => lines.push(`  - ${r.file}${r.po ? ` [PO ${r.po}]` : ''}\n    ${r.link}`));
    });
  return { subject: `Attachments ${dayLabel}: ${rows.length} file(s), ${bySender.size} sender(s)`,
           text: lines.join('\n') };
}

// Lets Node's test runner import the pure helpers; ignored by Apps Script.
if (typeof module !== 'undefined') {
  module.exports = { extractPoNumber, senderName, sanitizeFilename, extOf, isAllowed, buildFileName, buildDigest };
}
