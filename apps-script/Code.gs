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
var C = { ts:1, event:2, name:3, email:4, phone:5, ig:6, party:7, heard:8, notes:9, list:10, status:11, checkin:12, ticket:13, checkedAt:14, diet:15, celebrate:16 };
var LAST = 16;
// Columns M onwards, added by the script when missing
var EXTRA_HEADERS = ['Ticket', 'Checked in at', 'Dietary', 'Raising a glass to'];

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
    row[C.diet-1] = text_(clean_(d.diet, 500));
    row[C.celebrate-1] = text_(clean_(d.celebrate, 500));
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
    .addSeparator()
    .addItem('Preview gallery (in Drive)', 'previewGallery')
    .addItem('Publish gallery', 'publishGallery')
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

  // Sent through Gmail rather than MailApp: Hotmail/Outlook silently drop MailApp tickets from
  // gapevents.vn, but accept the same email sent from Gmail.
  var opts = { htmlBody: html, name: FROM_NAME, replyTo: REPLY_TO };
  if (inline.ticketqr) opts.inlineImages = inline;
  var text = html.replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|h1|tr|div)>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/\n{3,}/g, '\n\n').trim();
  GmailApp.sendEmail(g.email, subject, text, opts);
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
  var n = EXTRA_HEADERS.length, h = sheet.getRange(1, C.ticket, 1, n).getValues()[0];
  if (h.join('|') !== EXTRA_HEADERS.join('|')) sheet.getRange(1, C.ticket, 1, n).setValues([EXTRA_HEADERS]);
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
  handle: 'Social Media Handle', loop: 'Stay in the loop?', outreached: 'Outreached', heard: 'Heard about us'
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
    drink: clean_(d.drink, 80), diet: clean_(d.diet, 500), celebrate: clean_(d.celebrate, 500), loop: clean_(d.loop, 60),
    heard: clean_(d.heard, 80)
  };
  if (!v.name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.email) || v.phone.replace(/\D/g, '').length < 8 || !v.handle ||
      !v.passport) return { ok:false, error:'missing_fields' };
  // Drinks, dietary and celebrations are asked per event now, so they're optional here.
  for (var k in COMMUNITY_CHOICES) {
    if (k === 'drink' && !v.drink) continue;
    if (COMMUNITY_CHOICES[k].indexOf(v[k]) < 0) return { ok:false, error:'missing_fields' };
  }

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
    put('diet', text_(v.diet)); put('celebrate', text_(v.celebrate)); put('loop', v.loop); put('heard', text_(v.heard));
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

// A search shorter than 2 characters returns the whole guest list for the event (so staff can
// scroll it at the door): not-yet-arrived guests first, then A–Z by name.
function search_(id, q) {
  q = String(q || '').toLowerCase().trim();
  var all = q.length < 2;
  var digits = q.replace(/\D/g, '');
  var rows = regRows_().rows, out = [];
  for (var i = 0; i < rows.length && (all || out.length < 20); i++) {
    var r = rows[i];
    if (String(r[C.event-1]) !== id) continue;
    if (!all) {
      var hay = [r[C.name-1], r[C.email-1], r[C.ig-1], r[C.ticket-1]].join(' ').toLowerCase();
      var phone = String(r[C.phone-1]).replace(/\D/g, '');
      if (hay.indexOf(q) < 0 && !(digits.length >= 4 && phone.indexOf(digits) > -1)) continue;
    }
    out.push(guest_(r, i + 2));
  }
  if (all) out.sort(function (a, b) {
    return (a.checkedIn - b.checkedIn) || String(a.name).localeCompare(String(b.name), 'vi', { sensitivity: 'base' });
  });
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


/* ---------- photo gallery (gapevents.vn/gallery) ----------
   The "Gallery" tab lists albums: Album | Date (shown) | Sort date | Venue | Drive folder | Show.
   Rows with the same Album name are merged. For each folder the photos come from a "Website"
   subfolder if there is one, else "Final Photos", else the folder itself (raw folders are skipped).
   Up to 40 per album, spread evenly. Each photo is resized to 1600px by Drive and committed to
   GitHub (images/events/<album>/<file id>.jpg) with images/events/albums.json, which Netlify then
   publishes. Needs a GitHub token in Script properties → GITHUB_TOKEN (contents read & write on the
   gap repo). Big runs carry on by themselves a minute later until everything is published.

   Review first: "Preview gallery" fills the Drive folder "gặp — Website gallery" with one subfolder per
   album, holding shortcuts to the photos that would go up. Delete or add shortcuts there to change an
   album; once an album's preview folder has photos in it, publishing uses exactly those. */

var GALLERY_TAB = 'Gallery';
var GALLERY_MAX = 40;
var GALLERY_PX = 1600;
var GITHUB_REPO = 'gap-events/gap';
var GITHUB_BRANCH = 'main';
var GALLERY_DIR = 'images/events';
var GALLERY_JSON = 'images/events/albums.json';
var GALLERY_BUDGET_MS = 4.5 * 60 * 1000;
var GALLERY_ROOT_NAME = 'gặp — Website gallery';

function publishGallery() { publishGallery_(true); }
function continueGallery() { publishGallery_(false); }

function publishGallery_(fromMenu) {
  var ui = fromMenu ? SpreadsheetApp.getUi() : null;
  var say = function (m) { if (ui) ui.alert(m); else console.log(m); };
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) return say('Add a GitHub token first: Extensions → Apps Script → Project Settings → Script properties → GITHUB_TOKEN.');

  var cache = CacheService.getScriptCache();
  if (cache.get('gallery_running')) return say('The gallery is already publishing. Give it a few minutes.');
  cache.put('gallery_running', '1', 360);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'continueGallery') ScriptApp.deleteTrigger(t); });

  try {
    var start = Date.now();
    var sheet = ss_().getSheetByName(GALLERY_TAB);
    var albums = galleryAlbums_(sheet);
    var gh = github_(token);
    var headSha = gh('get', '/git/ref/heads/' + GITHUB_BRANCH).object.sha;
    var baseTree = gh('get', '/git/commits/' + headSha).tree.sha;

    var existing = {}, old = gh('raw', '/contents/' + GALLERY_JSON + '?ref=' + GITHUB_BRANCH);
    if (old && old.albums) old.albums.forEach(function (a) { a.photos.forEach(function (p) { existing[p.id] = p; }); });

    var tree = [], used = {}, pending = 0, added = 0, failed = 0;
    albums.forEach(function (a) {
      a.photos = [];
      a.pick.forEach(function (f) {
        var path = GALLERY_DIR + '/' + a.slug + '/' + f.id + '.jpg';
        var ex = existing[f.id];
        if (ex && ex.src === '/' + path) { used[path] = 1; a.photos.push(ex); return; }
        if (Date.now() - start > GALLERY_BUDGET_MS) { pending++; return; }
        var img = galleryImage_(f.id);
        if (!img) { failed++; return; }
        var blob = gh('post', '/git/blobs', { content: Utilities.base64Encode(img.bytes), encoding: 'base64' });
        tree.push({ path: path, mode: '100644', type: 'blob', sha: blob.sha });
        used[path] = 1; added++;
        a.photos.push({ id: f.id, src: '/' + path, w: img.w, h: img.h });
      });
    });

    // Remove photos that are no longer picked (only once everything new is in).
    var removed = 0;
    if (!pending) Object.keys(existing).forEach(function (id) {
      var path = existing[id].src.replace(/^\//, '');
      if (!used[path]) { tree.push({ path: path, mode: '100644', type: 'blob', sha: null }); removed++; }
    });

    var manifest = {
      updated: new Date().toISOString(),
      albums: albums.filter(function (a) { return a.photos.length; }).map(function (a) {
        return { slug: a.slug, title: a.title, date: a.date, sort: a.sort, venue: a.venue, cover: a.photos[0].src, photos: a.photos };
      })
    };
    var changed = added || removed || !old || JSON.stringify(old.albums) !== JSON.stringify(manifest.albums);
    if (changed) {
      tree.push({ path: GALLERY_JSON, mode: '100644', type: 'blob', content: JSON.stringify(manifest) });
      var newTree = gh('post', '/git/trees', { base_tree: baseTree, tree: tree });
      var commit = gh('post', '/git/commits', { message: 'Gallery: ' + added + ' added, ' + removed + ' removed' + (pending ? ' (more to come)' : ''), tree: newTree.sha, parents: [headSha] });
      gh('patch', '/git/refs/heads/' + GITHUB_BRANCH, { sha: commit.sha });
    }

    // Report back on the Gallery tab (column G) and in a message.
    var counts = {};
    albums.forEach(function (a) { counts[a.title] = a.photos.length + ' photos' + (a.source ? ' from ' + a.source : ''); });
    var rows = sheet.getLastRow() - 1;
    if (rows > 0) {
      sheet.getRange(1, 7).setValue('Published');
      var vals = sheet.getRange(2, 1, rows, 6).getValues().map(function (r) {
        return [String(r[5]).trim().toLowerCase() === 'yes' && counts[String(r[0]).trim()] ? counts[String(r[0]).trim()] : ''];
      });
      sheet.getRange(2, 7, rows, 1).setValues(vals);
    }

    var msg = added + ' photos added, ' + removed + ' removed across ' + manifest.albums.length + ' albums.' +
      (failed ? ' ' + failed + " couldn't be read from Drive and were skipped." : '');
    if (pending) {
      ScriptApp.newTrigger('continueGallery').timeBased().after(60 * 1000).create();
      msg += ' ' + pending + ' photos still to go: carrying on automatically in a minute. The site updates as each batch finishes.';
    } else {
      msg += changed ? ' The website updates in about a minute.' : ' Nothing new to publish.';
    }
    sheet.getRange(1, 8).setValue('Last run: ' + Utilities.formatDate(new Date(), TZ, 'd MMM, h:mm a') + '. ' + msg);
    say(msg);
  } catch (err) {
    console.error(err);
    say('Publishing stopped: ' + (err && err.message || err));
  } finally {
    cache.remove('gallery_running');
  }
}

function galleryAlbums_(sheet) {
  if (!sheet) throw new Error('There is no "' + GALLERY_TAB + '" tab.');
  var vals = sheet.getDataRange().getValues(), by = {}, order = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i], title = String(r[0]).trim(), link = String(r[4]).trim();
    if (!title || !link || String(r[5]).trim().toLowerCase() !== 'yes') continue;
    var m = /folders\/([\w-]+)/.exec(link) || /id=([\w-]+)/.exec(link) || /^([\w-]{20,})$/.exec(link);
    if (!m) continue;
    if (!by[title]) {
      var sort = r[2] instanceof Date ? Utilities.formatDate(r[2], TZ, 'yyyy-MM-dd') : String(r[2]).trim();
      by[title] = { title: title, date: String(r[1]).trim(), sort: sort, venue: String(r[3]).trim(), folders: [] };
      order.push(title);
    }
    by[title].folders.push(m[1]);
  }
  return order.map(function (t) {
    var a = by[t];
    a.slug = slug_(t + (a.sort ? '-' + a.sort.slice(0, 7) : ''));
    var files = [], sources = [];
    a.folders.forEach(function (id) { var got = galleryFolderPhotos_(id); files = files.concat(got.files); if (got.source) sources.push(got.source); });
    a.source = sources.filter(function (s, i) { return sources.indexOf(s) === i; }).join(', ');
    a.pick = spread_(files, GALLERY_MAX);
    var curated = curatedPhotos_(t);
    if (curated && curated.length) { a.pick = curated; a.source = 'your Website gallery folder'; a.curated = true; }
    return a;
  }).sort(function (x, y) { return String(y.sort).localeCompare(String(x.sort)); });
}

// Website → Final Photos → the folder itself → an edited-looking subfolder → any subfolder (never raw).
function galleryFolderPhotos_(id) {
  var folder = DriveApp.getFolderById(id), subs = [], it = folder.getFolders();
  while (it.hasNext()) subs.push(it.next());
  var named = function (re) { return subs.filter(function (f) { return re.test(f.getName()); })[0]; };
  var tries = [named(/^website$/i), named(/^final/i), folder, named(/final|edit|colou?r|graded|select|photos?/i)]
    .concat(subs.filter(function (f) { return !/raw/i.test(f.getName()); }));
  for (var i = 0; i < tries.length; i++) {
    var f = tries[i]; if (!f) continue;
    var files = [], fi = f.getFiles();
    while (fi.hasNext()) { var x = fi.next(); if (/^image\//.test(x.getMimeType())) files.push({ id: x.getId(), name: x.getName() }); }
    if (files.length) {
      files.sort(function (a, b) { return a.name.localeCompare(b.name, undefined, { numeric: true }); });
      return { files: files, source: f.getId() === id ? '' : f.getName() };
    }
  }
  return { files: [], source: '' };
}

function spread_(list, max) {
  if (list.length <= max) return list;
  var out = [];
  for (var i = 0; i < max; i++) out.push(list[Math.round(i * (list.length - 1) / (max - 1))]);
  return out;
}

// Drive makes the resized copy (longest side GALLERY_PX); we keep its size for the page layout.
function galleryImage_(id) {
  var auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  try {
    var meta = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id +
      '?fields=thumbnailLink,imageMediaMetadata(width,height,rotation)&supportsAllDrives=true', { headers: auth }).getContentText());
    var m = meta.imageMediaMetadata || {}, w = m.width || 0, h = m.height || 0;
    if (m.rotation === 1 || m.rotation === 3) { var t = w; w = h; h = t; }
    var urls = [];
    if (meta.thumbnailLink) urls.push(meta.thumbnailLink.replace(/=s\d+(-[^&?]*)?$/, '') + '=s' + GALLERY_PX);
    urls.push('https://drive.google.com/thumbnail?id=' + id + '&sz=s' + GALLERY_PX);
    for (var i = 0; i < urls.length; i++) {
      var r = UrlFetchApp.fetch(urls[i], { headers: auth, muteHttpExceptions: true });
      if (r.getResponseCode() !== 200 || !/^image\//.test(r.getHeaders()['Content-Type'] || 'image/')) continue;
      var bytes = r.getBlob().getBytes();
      if (bytes.length < 2000) continue;
      var scale = w && h ? Math.min(1, GALLERY_PX / Math.max(w, h)) : 1;
      return { bytes: bytes, w: Math.round((w || 4) * scale), h: Math.round((h || 3) * scale) };
    }
  } catch (err) { console.error('gallery image ' + id + ': ' + err); }
  return null;
}

function github_(token) {
  var base = 'https://api.github.com/repos/' + GITHUB_REPO;
  return function (method, path, body) {
    var raw = method === 'raw';
    var r = UrlFetchApp.fetch(base + path, {
      method: raw ? 'get' : method, muteHttpExceptions: true, contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token, Accept: raw ? 'application/vnd.github.raw+json' : 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' },
      payload: body ? JSON.stringify(body) : undefined
    });
    var code = r.getResponseCode();
    if (raw && code === 404) return null;
    if (code >= 300) throw new Error('GitHub ' + code + ' on ' + path + ': ' + r.getContentText().slice(0, 200));
    return JSON.parse(r.getContentText());
  };
}

function slug_(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[đĐ]/g, 'd')
    .toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60);
}


/* ---------- gallery preview in Drive ---------- */

function previewGallery() {
  var ui = SpreadsheetApp.getUi();
  try {
    var root = galleryRoot_(true), sheet = ss_().getSheetByName(GALLERY_TAB);
    var albums = galleryAlbums_(sheet), made = [], kept = [];
    albums.forEach(function (a) {
      if (a.curated) { kept.push(a.title + ' (' + a.pick.length + ')'); return; }
      var folder = subfolder_(root, a.title, true);
      a.pick.forEach(function (f) { folder.createShortcut(f.id); });
      made.push(a.title + ' (' + a.pick.length + ')');
    });
    ui.alert('Gallery preview ready in Google Drive → "' + GALLERY_ROOT_NAME + '".\n\n' +
      (made.length ? 'New: ' + made.join(', ') + '.\n' : '') +
      (kept.length ? 'Already reviewed, left as they are: ' + kept.join(', ') + '.\n' : '') +
      '\nOpen each album folder to check the photos. Delete a shortcut to drop a photo, or add one ' +
      '(right-click a photo → Organise → Add shortcut). When you are happy, use gặp → Publish gallery.\n\n' +
      'To start an album again, delete its folder there and run Preview gallery again.');
  } catch (err) {
    console.error(err);
    ui.alert('Preview stopped: ' + (err && err.message || err));
  }
}

function galleryRoot_(create) {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('GALLERY_FOLDER_ID');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var it = DriveApp.getRootFolder().getFoldersByName(GALLERY_ROOT_NAME);
  var root = it.hasNext() ? it.next() : (create ? DriveApp.getRootFolder().createFolder(GALLERY_ROOT_NAME) : null);
  if (root) props.setProperty('GALLERY_FOLDER_ID', root.getId());
  return root;
}

function subfolder_(parent, name, create) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return create ? parent.createFolder(name) : null;
}

// Photos in an album's preview folder (shortcuts or real images), in name order; null if none yet.
function curatedPhotos_(title) {
  var root = galleryRoot_(false); if (!root) return null;
  var folder = subfolder_(root, title, false); if (!folder) return null;
  var out = [], seen = {}, it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next(), id = null;
    if (f.getMimeType() === 'application/vnd.google-apps.shortcut') {
      try { if (/^image\//.test(f.getTargetMimeType())) id = f.getTargetId(); } catch (e) {}
    } else if (/^image\//.test(f.getMimeType())) id = f.getId();
    if (id && !seen[id] && !f.isTrashed()) { seen[id] = 1; out.push({ id: id, name: f.getName() }); }
  }
  out.sort(function (a, b) { return a.name.localeCompare(b.name, undefined, { numeric: true }); });
  return out.length ? out : null;
}
