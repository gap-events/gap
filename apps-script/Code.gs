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
 *
 * QR check-in
 *   Every sign-up gets a ticket code (Registrations column M). Confirmed guests get it as a QR
 *   in their email. Staff scan it at gapevents.vn/checkin, which ticks "Checked in" and stamps
 *   the time (column N). The check-in page needs the staff PIN, kept in Project Settings →
 *   Script properties → STAFF_PIN (never in this file, because this file is public on GitHub).
 *   Emailing the QR fetches the image from quickchart.io, so the first run asks for permission
 *   to "connect to an external service".
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
var C = { ts:1, event:2, name:3, email:4, phone:5, ig:6, party:7, heard:8, notes:9, list:10, status:11, checkin:12, ticket:13, checkedAt:14 };
var LAST = 14;
var HEADERS = { ticket: 'Ticket', checkedAt: 'Checked in at' };

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

  // "Join the gặp community" page (gapevents.vn/community).
  if (d.action === 'join') {
    try { return json_(joinCommunity_(d)); }
    catch (err) { console.error(err); return json_({ ok:false, error:'server', detail: String(err && err.message || err) }); }
  }

  // Staff check-in calls (from gapevents.vn/checkin) carry an action and the staff PIN.
  if (d.action) return json_(staff_(d));

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
    var rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, C.status).getValues() : [];

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
    var ticket = newTicket_();
    row[C.ticket-1] = ticket;
    row[C.checkedAt-1] = '';
    ensureHeaders_(sheet);
    sheet.appendRow(row);
    sheet.getRange(sheet.getLastRow(), C.checkin).insertCheckboxes();
  } finally {
    lock.releaseLock();
  }

  try { sendEmail_(ev, { name:name, email:email, party:party, ticket:ticket }, status); } catch (err) { console.error(err); }

  // The ticket goes back only on a brand-new confirmed sign-up (never on 'duplicate'), so nobody can
  // fetch someone else's ticket by typing their email.
  var out = { ok:true, result: status === 'Waitlist' ? 'waitlist' : 'confirmed', state: stateFor_(getEvent_(id)) };
  if (status === 'Confirmed') out.ticket = ticket;
  return json_(out);
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
  ensureHeaders_(sheet);
  var range = sheet.getActiveRange();
  var sent = 0;
  for (var r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) {
    if (r < 2) continue;
    var v = sheet.getRange(r, 1, 1, LAST).getValues()[0];
    var ev = getEvent_(String(v[C.event-1]));
    if (!ev || !v[C.email-1]) continue;
    if (promote) sheet.getRange(r, C.status).setValue('Confirmed');
    var status = promote ? 'Confirmed' : v[C.status-1];
    if (status !== 'Confirmed' && status !== 'Waitlist') continue;
    // Rows added by hand (or before QR tickets existed) get a ticket now.
    var ticket = String(v[C.ticket-1] || '').trim();
    if (!ticket) { ticket = newTicket_(); sheet.getRange(r, C.ticket).setValue(ticket); }
    sendEmail_(ev, { name:v[C.name-1], email:v[C.email-1], party:v[C.party-1], ticket:ticket }, status, promote);
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

// Sheets treats text starting with = + - @ as a formula (e.g. +84 phone numbers), and turns digits into a
// number (dropping the leading 0 of Vietnamese phone numbers), so store both as plain text.
function text_(v) { v = String(v == null ? '' : v); return /^[=+\-@0-9]/.test(v) ? "'" + v : v; }

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

  // Confirmed guests get their QR ticket. If the QR image can't be made, the code still goes out as text.
  var ticketHtml = '', inline = {};
  if (!waitlist && g.ticket) {
    var qr = qrBlob_(g.ticket);
    if (qr) inline.ticketqr = qr;
    ticketHtml =
      '<div style="margin:0 0 24px;padding:22px 16px;border:1px solid #E4D3C7;border-radius:14px;background:#FFFFFF;text-align:center">' +
        '<p style="margin:0 0 12px;font-size:12px;letter-spacing:.16em;text-transform:uppercase;color:#8C2A1E;font-weight:bold">Your ticket</p>' +
        (qr ? '<img src="cid:ticketqr" width="200" height="200" alt="QR ticket ' + esc_(g.ticket) + '" style="display:block;margin:0 auto 12px;width:200px;height:200px">' : '') +
        '<p style="margin:0 0 6px;font-family:Menlo,Consolas,monospace;font-size:20px;letter-spacing:.12em;font-weight:bold;color:#3A1F14">' + esc_(g.ticket) + '</p>' +
        '<p style="margin:0;font-size:14px;line-height:1.5;color:#6B5045">Show this at the door and we\'ll scan you in' +
          (Number(g.party) > 1 ? ' (one code covers your whole party)' : '') + '.</p>' +
      '</div>';
  }

  var html =
    '<div style="background:#F8EEE7;padding:32px 16px;font-family:Helvetica,Arial,sans-serif;color:#3A1F14">' +
    '<div style="max-width:520px;margin:0 auto;background:#FFF8F3;border-radius:18px;padding:32px 28px">' +
      '<p style="margin:0 0 6px;font-size:12px;letter-spacing:.16em;text-transform:uppercase;color:#8C2A1E;font-weight:bold">' + esc_(ev.name) + '</p>' +
      '<h1 style="margin:0 0 16px;font-size:28px;line-height:1.15">' + (waitlist ? "You're on the waitlist." : 'Hi ' + esc_(first)) + '</h1>' +
      '<p style="margin:0 0 20px;font-size:16px;line-height:1.6;color:#6B5045">' + lead + '</p>' +
      ticketHtml +
      (rows ? '<table style="width:100%;border-collapse:collapse;margin:0 0 20px;font-size:15px">' + rows + '</table>' : '') +
      (ev.note && !waitlist ? '<p style="margin:0 0 20px;font-size:15px;line-height:1.6">' + esc_(ev.note).replace(/\n/g, '<br>') + '</p>' : '') +
      (cal ? '<p style="margin:0 0 24px"><a href="' + cal + '" style="display:inline-block;background:#3A1F14;color:#F8EEE7;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:999px">Add to calendar</a></p>' : '') +
      '<p style="margin:0;font-size:14px;line-height:1.6;color:#6B5045">' +
        (waitlist ? '' : "Can't make it anymore? Just reply to this email so we can pass your spot on. ") +
        'Questions? Reply here or DM us on <a href="https://www.instagram.com/gap.vietnam/" style="color:#C1592F">Instagram</a>.</p>' +
    '</div>' +
    '<p style="text-align:center;font-size:13px;color:#6B5045;margin:18px 0 0">gặp · Connection begins with a seat at the table.</p>' +
    '</div>';

  var msg = { to: g.email, subject: subject, htmlBody: html, name: FROM_NAME, replyTo: REPLY_TO };
  if (inline.ticketqr) msg.inlineImages = inline;
  MailApp.sendEmail(msg);
}

/* ---------- tickets ---------- */

// 10 characters from 32 symbols with no look-alikes (no 0/O, 1/I), shown as XXXXX-XXXXX.
function newTicket_() {
  var a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now());
  var s = '';
  for (var i = 0; i < 10; i++) s += a[(bytes[i] + 256) % 32];
  return s.slice(0, 5) + '-' + s.slice(5);
}

function normTicket_(t) {
  var s = String(t || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  return s.length === 10 ? s.slice(0, 5) + '-' + s.slice(5) : '';
}

function qrBlob_(text) {
  var urls = [
    'https://quickchart.io/qr?size=480&margin=2&ecLevel=M&format=png&text=' + encodeURIComponent(text),
    'https://api.qrserver.com/v1/create-qr-code/?size=480x480&margin=8&format=png&data=' + encodeURIComponent(text)
  ];
  for (var i = 0; i < urls.length; i++) {
    try {
      var r = UrlFetchApp.fetch(urls[i], { muteHttpExceptions: true });
      if (r.getResponseCode() === 200) return r.getBlob().setName('ticket.png').setContentType('image/png');
    } catch (err) { console.error(err); }
  }
  return null;
}

function ensureHeaders_(sheet) {
  var h = sheet.getRange(1, C.ticket, 1, 2).getValues()[0];
  if (h[0] !== HEADERS.ticket || h[1] !== HEADERS.checkedAt) sheet.getRange(1, C.ticket, 1, 2).setValues([[HEADERS.ticket, HEADERS.checkedAt]]);
}

/* ---------- the gặp community (gapevents.vn/community) ---------- */

// Sign-ups go into the same tab the old Google Form filled ("EOI Form Responses" in the main gặp
// planning spreadsheet), so existing lists keep working. Its ID lives in Script properties as
// COMMUNITY_SHEET_ID. Columns are matched by the first line of each header, so their order can change.
var COMMUNITY_TAB = 'EOI Form Responses';
var COMMUNITY_COLS = {
  ts: 'Timestamp', name: 'Full Name', email: 'Email', phone: 'Phone Number (inc country code)',
  gender: 'Gender', age: 'Age', passport: 'What passport(s) do you travel on?', drink: 'Will you be joining us for a drink?',
  diet: 'Anything we should know about before planning the menu?', celebrate: 'Is there anything worth raising a glass to?',
  handle: 'Social Media Handle', loop: 'Stay in the loop?', outreached: 'Outreached'
};
var COMMUNITY_CHOICES = {
  gender: ['Male', 'Female', 'Prefer not to say'],
  age: ['18 - 29', '30 - 45', '46+'],
  drink: ["Yes, I'll be drinking", "No thanks, I'll stick to non-alcoholic options", "Maybe, I'll decide on the night"],
  loop: ['Yes, keep me in the loop!', 'No thanks, not right now.']
};

function joinCommunity_(d) {
  if (d.website) return { ok:true };   // honeypot
  var v = {
    name: clean_(d.name, 80), email: clean_(d.email, 120).toLowerCase(), phone: clean_(d.phone, 40),
    handle: clean_(d.handle, 60), gender: clean_(d.gender, 40), age: clean_(d.age, 20), passport: clean_(d.passport, 120),
    drink: clean_(d.drink, 80), diet: clean_(d.diet, 500), celebrate: clean_(d.celebrate, 500), loop: clean_(d.loop, 60)
  };
  if (!v.name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.email) || v.phone.replace(/\D/g, '').length < 8 || !v.handle ||
      !v.passport || !v.diet) return { ok:false, error:'missing_fields' };
  for (var k in COMMUNITY_CHOICES) if (COMMUNITY_CHOICES[k].indexOf(v[k]) < 0) return { ok:false, error:'missing_fields' };

  var id = PropertiesService.getScriptProperties().getProperty('COMMUNITY_SHEET_ID');
  if (!id) return { ok:false, error:'not_configured' };
  var sheet = SpreadsheetApp.openById(id).getSheetByName(COMMUNITY_TAB);
  if (!sheet) return { ok:false, error:'not_configured' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var heads = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) {
      return String(h).split('\n')[0].trim().toLowerCase();
    });
    var row = heads.map(function () { return ''; });
    var put = function (key, val) { var i = heads.indexOf(COMMUNITY_COLS[key].toLowerCase()); if (i > -1) row[i] = val; };
    put('ts', new Date());
    put('name', text_(v.name)); put('email', v.email); put('phone', text_(v.phone)); put('handle', text_(v.handle));
    put('gender', v.gender); put('age', v.age); put('passport', text_(v.passport)); put('drink', v.drink);
    put('diet', text_(v.diet)); put('celebrate', text_(v.celebrate)); put('loop', v.loop);
    put('outreached', false);
    sheet.appendRow(row);
    // The tab is a Sheets table whose "Outreached" column is already a tickbox column, so no
    // checkbox needs adding here (and adding one is refused on typed columns).
  } finally {
    lock.releaseLock();
  }
  return { ok:true };
}

/* ---------- staff check-in (gapevents.vn/checkin) ---------- */

function staff_(d) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pin_fails') || 0);
  if (fails >= 20) return { ok:false, error:'locked' };
  var pin = PropertiesService.getScriptProperties().getProperty('STAFF_PIN');
  if (!pin) return { ok:false, error:'no_pin_set' };
  if (String(d.pin || '') !== String(pin)) {
    cache.put('pin_fails', String(fails + 1), 600);
    return { ok:false, error:'bad_pin' };
  }

  if (d.action === 'events') return { ok:true, events: listEvents_() };

  var id = String(d.event || '').trim();
  if (!getEvent_(id)) return { ok:false, error:'unknown_event' };

  if (d.action === 'stats') return { ok:true, stats: stats_(id) };
  if (d.action === 'search') return { ok:true, results: search_(id, d.q), stats: stats_(id) };
  if (d.action === 'checkin') return checkin_(id, d.ref, true);
  if (d.action === 'undo') return checkin_(id, d.ref, false);
  return { ok:false, error:'bad_action' };
}

function listEvents_() {
  var vals = ss_().getSheetByName(EVENTS).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    if (!String(vals[i][0]).trim()) continue;
    out.push({ id: String(vals[i][0]).trim(), name: String(vals[i][1] || vals[i][0]) });
  }
  return out;
}

function regRows_() {
  var sheet = ss_().getSheetByName(REGS);
  var n = sheet.getLastRow() - 1;
  return { sheet: sheet, rows: n > 0 ? sheet.getRange(2, 1, n, LAST).getValues() : [] };
}

function stats_(id) {
  var rows = regRows_().rows, seats = 0, inSeats = 0, guests = 0, inGuests = 0;
  rows.forEach(function (r) {
    if (String(r[C.event-1]) !== id || r[C.status-1] !== 'Confirmed') return;
    var p = Number(r[C.party-1]) || 1;
    guests++; seats += p;
    if (r[C.checkin-1] === true) { inGuests++; inSeats += p; }
  });
  return { guests: guests, seats: seats, inGuests: inGuests, inSeats: inSeats };
}

function guest_(r, rowNum) {
  var at = r[C.checkedAt-1];
  return {
    ref: String(r[C.ticket-1] || '') || ('row:' + rowNum),
    ticket: String(r[C.ticket-1] || ''),
    name: String(r[C.name-1] || ''),
    party: Number(r[C.party-1]) || 1,
    handle: String(r[C.ig-1] || ''),
    status: String(r[C.status-1] || ''),
    checkedIn: r[C.checkin-1] === true,
    at: at instanceof Date ? Utilities.formatDate(at, TZ, 'h:mm a') : ''
  };
}

function search_(id, q) {
  q = String(q || '').toLowerCase().trim();
  if (q.length < 2) return [];
  var digits = q.replace(/\D/g, '');
  var rows = regRows_().rows, out = [];
  for (var i = 0; i < rows.length && out.length < 20; i++) {
    var r = rows[i];
    if (String(r[C.event-1]) !== id) continue;
    var hay = [r[C.name-1], r[C.email-1], r[C.ig-1], r[C.ticket-1]].join(' ').toLowerCase();
    var phone = String(r[C.phone-1]).replace(/\D/g, '');
    if (hay.indexOf(q) > -1 || (digits.length >= 4 && phone.indexOf(digits) > -1)) out.push(guest_(r, i + 2));
  }
  return out;
}

function checkin_(id, ref, on) {
  ref = String(ref || '').trim();
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var data = regRows_(), rows = data.rows, idx = -1;
    var m = /^row:(\d+)$/.exec(ref);
    if (m) {
      idx = Number(m[1]) - 2;
      if (idx < 0 || idx >= rows.length) idx = -1;
    } else {
      var t = normTicket_(ref);
      if (t) for (var i = 0; i < rows.length; i++) if (normTicket_(rows[i][C.ticket-1]) === t) { idx = i; break; }
    }
    if (idx < 0) return { ok:true, result:'not_found', stats: stats_(id) };

    var r = rows[idx], rowNum = idx + 2, g = guest_(r, rowNum);
    if (String(r[C.event-1]) !== id) {
      var other = getEvent_(String(r[C.event-1]));
      return { ok:true, result:'wrong_event', guest:g, otherEvent: other ? String(other.name) : String(r[C.event-1]), stats: stats_(id) };
    }
    if (on) {
      if (g.status !== 'Confirmed') return { ok:true, result:'not_confirmed', guest:g, stats: stats_(id) };
      if (g.checkedIn) return { ok:true, result:'already', guest:g, stats: stats_(id) };
      var now = new Date();
      data.sheet.getRange(rowNum, C.checkin).setValue(true);
      data.sheet.getRange(rowNum, C.checkedAt).setValue(now);
      g.checkedIn = true; g.at = Utilities.formatDate(now, TZ, 'h:mm a');
      return { ok:true, result:'checked_in', guest:g, stats: stats_(id) };
    }
    data.sheet.getRange(rowNum, C.checkin).setValue(false);
    data.sheet.getRange(rowNum, C.checkedAt).setValue('');
    g.checkedIn = false; g.at = '';
    return { ok:true, result:'undone', guest:g, stats: stats_(id) };
  } finally {
    lock.releaseLock();
  }
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
  sendEmail_(ev, { name: 'gặp Test', email: Session.getActiveUser().getEmail(), party: 1, ticket: newTicket_() }, 'Confirmed');
}

/* Run once from the editor after setting COMMUNITY_SHEET_ID, to approve access to the community spreadsheet. */
function authorizeCommunity() {
  var id = PropertiesService.getScriptProperties().getProperty('COMMUNITY_SHEET_ID');
  var sheet = SpreadsheetApp.openById(id).getSheetByName(COMMUNITY_TAB);
  console.log('Community tab found: ' + !!sheet + ', rows: ' + (sheet ? sheet.getLastRow() : 0));
}
