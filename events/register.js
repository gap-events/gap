/* gặp event registration — shared by every page in /events.
   Paste the Apps Script web app URL below once (it stays the same for every event). */
var GAP_ENDPOINT = 'https://script.google.com/macros/s/AKfycbzlPGw92vSj4_C283lq1nihXIx8jFQvJoYLYxQ3UGIAina9lzN4ldoZTScmR-u5UY1p/exec';

(function () {
  var form = document.querySelector('form[data-event]');
  if (!form) return;
  var EVENT = form.dataset.event;
  var pill = document.querySelector('.pill');
  var btn = form.querySelector('button[type=submit]');
  var msg = form.querySelector('.msg');
  var lede = document.querySelector('.card .lede');
  var partyWrap = form.querySelector('.party');
  var party = form.querySelector('select[name=party]');
  var card = document.getElementById('register');
  var done = document.getElementById('done');
  var closed = document.getElementById('closed');
  var head = document.querySelector('.card-head');
  var defaultBtn = btn.textContent;

  var STATES = {
    open: 'Spots are limited',
    filling: 'Half the spots are gone',
    almost: 'Almost fully booked',
    waitlist: 'Fully booked · join the waitlist',
    closed: 'Registrations closed'
  };

  function setState(state) {
    if (!STATES[state]) return;
    if (pill) { pill.dataset.state = state; pill.textContent = STATES[state]; }
    document.querySelectorAll('[data-cta]').forEach(function (a) {
      a.textContent = state === 'waitlist' ? 'Join the waitlist ↓' : a.dataset.cta;
      a.hidden = state === 'closed';
    });
    if (state === 'closed') { if (!done.hidden) return; form.hidden = true; if (lede) lede.hidden = true; if (head) head.hidden = true; closed.hidden = false; return; }
    if (state === 'waitlist') {
      btn.textContent = defaultBtn = 'Join the waitlist';
      if (lede) lede.textContent = "We're fully booked, but plans change. Leave your details and we'll let you know the moment a spot opens up.";
    }
  }

  function setParty(max) {
    max = Math.max(1, Math.min(10, parseInt(max, 10) || 1));
    if (max < 2) { partyWrap.hidden = true; return; }
    party.innerHTML = '';
    for (var n = 1; n <= max; n++) {
      var o = document.createElement('option');
      o.value = n; o.textContent = n === 1 ? 'Just me' : 'Me + ' + (n - 1);
      party.appendChild(o);
    }
    partyWrap.hidden = false;
  }

  setParty(form.dataset.maxParty);

  var heardSel = form.querySelector('select[name=heard]'), heardOther = null;
  if (heardSel) {
    heardOther = document.createElement('input');
    heardOther.name = 'heard_other'; heardOther.id = 'heard_other'; heardOther.maxLength = 50;
    heardOther.placeholder = 'Tell us where'; heardOther.setAttribute('aria-label', 'Where did you hear about us?');
    heardOther.hidden = true; heardOther.style.marginTop = '8px';
    heardSel.parentNode.appendChild(heardOther);
    heardSel.addEventListener('change', function () {
      var on = heardSel.value === 'Other'; heardOther.hidden = !on; if (on) heardOther.focus();
    });
  }
  if (pill) { pill.dataset.state = 'open'; pill.textContent = STATES.open; }

  if (GAP_ENDPOINT) {
    fetch(GAP_ENDPOINT + '?event=' + encodeURIComponent(EVENT))
      .then(function (r) { return r.json(); })
      .then(function (d) { if (d.ok) { setState(d.state); if (d.maxParty) setParty(d.maxParty); } })
      .catch(function () {});
  }

  function fail(text, field) {
    msg.className = 'msg err'; msg.textContent = text;
    if (field) { field.setAttribute('aria-invalid', 'true'); field.focus(); }
  }

  form.addEventListener('input', function (e) { e.target.removeAttribute('aria-invalid'); });

  form.addEventListener('submit', function (e) {
    e.preventDefault();
    msg.textContent = '';
    var f = form.elements;
    // Optional fields can be left off any event page.
    function val(n) { return f[n] ? String(f[n].value).trim() : ''; }
    var name = f.name.value.trim(), email = f.email.value.trim(), phone = f.phone.value.trim();
    if (!name) return fail("What should we call you?", f.name);
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return fail('That email looks a little off. Try again?', f.email);
    if (phone.replace(/\D/g, '').length < 8) return fail("We'll need a phone number (Zalo or WhatsApp) for event-day updates.", f.phone);
    if (!GAP_ENDPOINT) return fail("Sign-ups aren't switched on yet. Email us at hello@gapevents.vn and we'll save you a spot.");

    btn.disabled = true; btn.textContent = 'Saving your spot…';
    var data = {
      event: EVENT, name: name, email: email, phone: phone,
      instagram: val('instagram'),
      party: partyWrap.hidden ? 1 : party.value,
      heard: (val('heard') === 'Other' && val('heard_other')) ? 'Other: ' + val('heard_other') : val('heard'), notes: val('notes'),
      list: !!(f.list && f.list.checked), website: val('website')
    };

    fetch(GAP_ENDPOINT, { method: 'POST', body: JSON.stringify(data) })
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (!d.ok) {
          if (d.error === 'closed') { setState('closed'); return; }
          throw new Error(d.error);
        }
        var wait = d.result === 'waitlist' || (d.result === 'duplicate' && d.status === 'waitlist');
        var first = name.split(' ')[0];
        done.querySelector('h2').textContent = d.result === 'duplicate' ? "You're already on the list, " + first + '.' : wait ? "You're on the waitlist, " + first + '.' : "You're in, " + first + '!';
        done.querySelector('p').textContent = d.result === 'duplicate'
          ? "Looks like you're already signed up for this one. Check your inbox for the details, or email hello@gapevents.vn if anything's changed."
          : wait
            ? "We'll email you at " + email + " as soon as a spot opens up."
            : "Your spot is saved. We've sent the details to " + email + ' (check your promotions folder if it hides). Hẹn gặp bạn!';
        form.hidden = true; if (lede) lede.hidden = true; if (head) head.hidden = true; done.hidden = false;
        card.scrollIntoView({ behavior: 'smooth', block: 'start' });
        if (d.state) setState(d.state);
      })
      .catch(function () {
        btn.disabled = false; btn.textContent = defaultBtn;
        fail("Something went wrong on our end. Try again, or email hello@gapevents.vn and we'll sort you out.");
      });
  });
})();
