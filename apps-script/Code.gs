/**
 * gặp — event registrations
 * Paste this into the Apps Script editor of the "gặp — Event registrations" Google Sheet
 * (Extensions → Apps Script), then Deploy → New deployment → Web app
 * (Execute as: Me, Who has access: Anyone). Paste the web app URL into events/register.js.
 *
 * Tabs it expects (already set up in the Sheet):
 *   Events        — one row per event. Column A "Event ID" must match data-event on the page.
 *   Registrations — one row per sign-up, written by this script.
 *
 * The public page never sees numbers. It only gets a state:
 *   open → filling → almost → waitlist (full) → closed
 */

var EVENTS = 'Events';
var REGS = 'Registrations';

// When the page says "filling fast" / "almost full" (share of capacity taken)
var FILLING_AT = 0.5;
var ALMOST_AT = 0.85;

var FROM_NAME = 'gặp';
var REPLY_TO = 'hello@gapevents.vn';
var TZ = 'Asia/Ho_Chi_Minh';

// Registrations columns (1-based)
var C = { ts:1, event:2, name:3, email:4, phone:5, ig:6, party:7, heard:8, notes:9, list:10, status:11, checkin:12 };

/* ---------- web endpoints ---------- */

function doGet(e) {
  var id = (e && e.parameter && e.parameter.event || '').trim();
  var ev = getEvent_(id);
  if (!ev) return json_({ ok:false, error:'unknown_event' });
  return json_({ ok:true, state: stateFor_(ev), maxParty: ev.maxParty });
}

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok:false, error:'bad_request' }); }

  // Honeypot: bots fill the hidden "website" field. Pretend it worked.
  if (d.website) return json_({ ok:true, result:'confirmed' });

  var id = String(d.event || '').trim();
  var name = clean_(d.name, 80);
  var email = clean_(d.email, 120).toLowerCase();
  var phone = clean_(d.phone, 40);
  if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json_({ ok:false, error:'missing_fields' });

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ev = getEvent_(id);
    if (!ev) return json_({ ok:false, error:'unknown_event' });
    if (ev.closed) return json_({ ok:false, error:'closed' });

    var party = Math.max(1, Math.min(ev.maxParty, parseInt(d.party, 10) || 1));
    var sheet = ss_().getSheetByName(REGS);
    var rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, C.checkin).getValues() : [];

    var taken = 0;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r[C.event-1]) !== id) continue;
      if (String(r[C.email-1]).toLowerCase() === email && r[C.status-1] !== 'Cancelled') {
        return json_({ ok:true, result:'duplicate', status: r[C.status-1] === 'Waitlist' ? 'waitlist' : 'confirmed' });
      }
      if (r[C.status-1] === 'Confirmed') taken += Number(r[C.party-1]) || 1;
    }

    var status = (ev.capacity && taken + party > ev.capacity) ? 'Waitlist' : 'Confirmed';
    var row = [];
    row[C.ts-1] = new Date();
    row[C.event-1] = id;
    row[C.name-1] = text_(name);
    row[C.email-1] = email;
    row[C.phone-1] = text_(phone);
    row[C.ig-1] = text_(clean_(d.instagram, 60));
    row[C.party-1] = party;
    row[C.heard-1] = text_(clean_(d.heard, 60));
    row[C.notes-1] = text_(clean_(d.notes, 500));
    row[C.list-1] = d.list ? 'Yes' : 'No';
    row[C.status-1] = status;
    row[C.checkin-1] = false;
    sheet.appendRow(row);
    sheet.getRange(sheet.getLastRow(), C.checkin).insertCheckboxes();
  } finally {
    lock.releaseLock();
  }

  try { sendEmail_(ev, { name:name, email:email, party:party }, status); } catch (err) { console.error(err); }

  return json_({ ok:true, result: status === 'Waitlist' ? 'waitlist' : 'confirmed', state: stateFor_(getEvent_(id)) });
}

/* ---------- Sheet menu: move people off the waitlist ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('gặp')
    .addItem('Confirm selected guests + email them', 'confirmSelected')
    .addItem('Resend confirmation to selected guests', 'resendSelected')
    .addToUi();
}

function confirmSelected() { emailSelected_(true); }
function resendSelected() { emailSelected_(false); }

function emailSelected_(promote) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  if (sheet.getName() !== REGS) { ui.alert('Select rows on the Registrations tab first.'); return; }
  var range = sheet.getActiveRange();
  var sent = 0;
  for (var r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) {
    if (r < 2) continue;
    var v = sheet.getRange(r, 1, 1, C.checkin).getValues()[0];
    var ev = getEvent_(String(v[C.event-1]));
    if (!ev || !v[C.email-1]) continue;
    if (promote) sheet.getRange(r, C.status).setValue('Confirmed');
    var status = promote ? 'Confirmed' : v[C.status-1];
    if (status !== 'Confirmed' && status !== 'Waitlist') continue;
    sendEmail_(ev, { name:v[C.name-1], email:v[C.email-1], party:v[C.party-1] }, status, promote);
    sent++;
  }
  ui.alert(sent + (sent === 1 ? ' email' : ' emails') + ' sent.');
}

/* ---------- helpers ---------- */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getEvent_(id) {
  if (!id) return null;
  var sheet = ss_().getSheetByName(EVENTS);
  var vals = sheet.getDataRange().getValues();
  var h = vals[0].map(function (x) { return String(x).trim().toLowerCase(); });
  function col(name) { return h.indexOf(name); }
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i];
    if (String(row[col('event id')]).trim() !== id) continue;
    var ev = {
      id: id,
      name: row[col('event name')],
      when: row[col('when (shown in email)')],
      start: row[col('start')],
      end: row[col('end')],
      venue: row[col('venue')],
      map: row[col('map link')],
      capacity: Number(row[col('capacity')]) || 0,
      closed: String(row[col('registration')]).trim().toLowerCase() === 'closed',
      maxParty: Math.max(1, Number(row[col('max per sign-up')]) || 1),
      note: row[col('note for confirmation email')],
      page: row[col('page link')]
    };
    ev.taken = countTaken_(id);
    return ev;
  }
  return null;
}

function countTaken_(id) {
  var sheet = ss_().getSheetByName(REGS);
  if (sheet.getLastRow() < 2) return 0;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, C.status).getValues();
  var n = 0;
  rows.forEach(function (r) {
    if (String(r[C.event-1]) === id && r[C.status-1] === 'Confirmed') n += Number(r[C.party-1]) || 1;
  });
  return n;
}

function stateFor_(ev) {
  if (ev.closed) return 'closed';
  if (!ev.capacity) return 'open';
  var share = ev.taken / ev.capacity;
  if (share >= 1) return 'waitlist';
  if (share >= ALMOST_AT) return 'almost';
  if (share >= FILLING_AT) return 'filling';
  return 'open';
}

function clean_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max); }

// Sheets treats text starting with = + - @ as a formula (e.g. +84 phone numbers), so store it as plain text.
function text_(v) { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; }

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[c]; }); }

function calLink_(ev) {
  if (!(ev.start instanceof Date)) return '';
  var end = ev.end instanceof Date ? ev.end : new Date(ev.start.getTime() + 3 * 3600 * 1000);
  var f = function (d) { return Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'"); };
  return 'https://calendar.google.com/calendar/render?action=TEMPLATE'
    + '&text=' + encodeURIComponent(ev.name)
    + '&dates=' + f(ev.start) + '/' + f(end)
    + '&location=' + encodeURIComponent(ev.venue || '')
    + '&details=' + encodeURIComponent('See you there! ' + (ev.page || 'https://gapevents.vn'));
}

function whenText_(ev) {
  if (ev.when) return String(ev.when);
  if (ev.start instanceof Date) return Utilities.formatDate(ev.start, TZ, "EEEE d MMMM, h:mm a");
  return '';
}

function sendEmail_(ev, g, status, promoted) {
  var first = String(g.name).split(' ')[0];
  var waitlist = status === 'Waitlist';
  var subject = waitlist
    ? "You're on the waitlist: " + ev.name
    : (promoted ? 'Good news, a spot opened up: ' : "You're in: ") + ev.name;

  var lead = waitlist
    ? "This one is fully booked, so we've popped you on the waitlist. If a spot opens up, you'll be the first to hear from us."
    : (promoted
        ? "A spot just opened up and it's yours. We can't wait to see you there."
        : "Your spot is saved. We can't wait to share a table with you.");

  var rows = '';
  var when = whenText_(ev);
  if (when) rows += row_('When', esc_(when));
  if (ev.venue) rows += row_('Where', esc_(ev.venue) + (ev.map ? ' · <a href="' + esc_(ev.map) + '" style="color:#C1592F">Map</a>' : ''));
  if (Number(g.party) > 1) rows += row_('Party', esc_(g.party) + ' people');

  var cal = !waitlist ? calLink_(ev) : '';
  var html =
    '<div style="background:#F8EEE7;padding:32px 16px;font-family:Helvetica,Arial,sans-serif;color:#3A1F14">' +
    '<div style="max-width:520px;margin:0 auto;background:#FFF8F3;border-radius:18px;padding:32px 28px">' +
      '<p style="margin:0 0 6px;font-size:12px;letter-spacing:.16em;text-transform:uppercase;color:#8C2A1E;font-weight:bold">' + esc_(ev.name) + '</p>' +
      '<h1 style="margin:0 0 16px;font-size:28px;line-height:1.15">' + (waitlist ? "You're on the waitlist." : 'Hẹn gặp bạn, ' + esc_(first) + '!') + '</h1>' +
      '<p style="margin:0 0 20px;font-size:16px;line-height:1.6;color:#6B5045">' + lead + '</p>' +
      (rows ? '<table style="width:100%;border-collapse:collapse;margin:0 0 20px;font-size:15px">' + rows + '</table>' : '') +
      (ev.note && !waitlist ? '<p style="margin:0 0 20px;font-size:15px;line-height:1.6">' + esc_(ev.note).replace(/\n/g, '<br>') + '</p>' : '') +
      (cal ? '<p style="margin:0 0 24px"><a href="' + cal + '" style="display:inline-block;background:#3A1F14;color:#F8EEE7;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:999px">Add to calendar</a></p>' : '') +
      '<p style="margin:0;font-size:14px;line-height:1.6;color:#6B5045">' +
        (waitlist ? '' : "Can't make it anymore? Just reply to this email so we can pass your spot on. ") +
        'Questions? Reply here or DM us on <a href="https://www.instagram.com/gap.vietnam/" style="color:#C1592F">Instagram</a>.</p>' +
    '</div>' +
    '<p style="text-align:center;font-size:13px;color:#6B5045;margin:18px 0 0">gặp · Connection begins with a seat at the table.</p>' +
    '</div>';

  MailApp.sendEmail({ to: g.email, subject: subject, htmlBody: html, name: FROM_NAME, replyTo: REPLY_TO });
}

function row_(k, v) {
  return '<tr><td style="padding:8px 0;border-top:1px solid #E4D3C7;color:#6B5045;width:70px;vertical-align:top">' + k +
    '</td><td style="padding:8px 0;border-top:1px solid #E4D3C7;font-weight:bold">' + v + '</td></tr>';
}

/* Run once from the editor to approve email + Sheet permissions and send yourself a sample. */
function testEmail() {
  var sheet = ss_().getSheetByName(EVENTS);
  var id = String(sheet.getRange(2, 1).getValue());
  var ev = getEvent_(id);
  if (!ev) throw new Error('Add an event on the Events tab first.');
  sendEmail_(ev, { name: 'gặp Test', email: Session.getActiveUser().getEmail(), party: 1 }, 'Confirmed');
}
