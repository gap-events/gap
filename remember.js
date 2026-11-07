/* gặp "remember me": keeps a guest's contact details in their own browser so the next sign-up
   is pre-filled. Nothing here is sent anywhere; it lives in localStorage on their device only.

   Use: add data-remember to a <form>, plus an optional <input type="checkbox" name="remember">.
   The form's own script dispatches a "gap:saved" event on the form after a successful submit. */
(function () {
  var KEY = 'gap_guest';
  // Stored field → names it may have on different forms
  var FIELDS = { name: ['name'], email: ['email'], phone: ['phone'], handle: ['handle', 'instagram'] };

  function load() { try { return JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { return null; } }
  function save(d) { try { localStorage.setItem(KEY, JSON.stringify(d)); } catch (e) {} }
  function clear() { try { localStorage.removeItem(KEY); } catch (e) {} }
  function field(form, key) {
    var names = FIELDS[key];
    for (var i = 0; i < names.length; i++) if (form.elements[names[i]]) return form.elements[names[i]];
    return null;
  }

  document.querySelectorAll('form[data-remember]').forEach(function (form) {
    var box = form.elements.remember;
    var saved = load();

    if (saved && saved.email) {
      var filled = false;
      Object.keys(FIELDS).forEach(function (k) {
        var el = field(form, k);
        if (el && !el.value && saved[k]) { el.value = saved[k]; filled = true; }
      });
      if (filled) {
        var note = document.createElement('p');
        note.className = 'remembered';
        note.innerHTML = '<span></span> <button type="button" class="linkish">Not you? Clear my details</button>';
        note.querySelector('span').textContent = 'Welcome back' + (saved.name ? ', ' + String(saved.name).split(' ')[0] : '') + '! We\'ve filled in your details.';
        form.insertBefore(note, form.firstChild);
        note.querySelector('button').addEventListener('click', function () {
          clear();
          Object.keys(FIELDS).forEach(function (k) { var el = field(form, k); if (el) el.value = ''; });
          note.remove();
          var first = field(form, 'name'); if (first) first.focus();
        });
      }
    }

    form.addEventListener('gap:saved', function () {
      if (box && !box.checked) { clear(); return; }
      var d = {};
      Object.keys(FIELDS).forEach(function (k) { var el = field(form, k); if (el && el.value.trim()) d[k] = el.value.trim(); });
      if (d.email) save(d);
    });
  });
})();
