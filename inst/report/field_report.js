// forplotR field report: find a tag ("1043") or a subplot ("S37", "N3");
// tapping a point on a subplot map highlights its table row and vice versa.
// No network, no libraries; the report also works with scripts disabled.
(function () {
  var body = document.body;
  var prefix = (body.getAttribute('data-sub-prefix') || '').toLowerCase();
  var esc = function (s) { return window.CSS && CSS.escape ? CSS.escape(s) : String(s).replace(/["\\]/g, '\\$&'); };
  var clear = function () {
    var on = document.querySelectorAll('.hl');
    for (var i = 0; i < on.length; i++) on[i].classList.remove('hl');
  };
  var mark = function (row) {
    clear();
    if (!row) return null;
    row.classList.add('hl');
    var sec = row.closest('.sub');
    var tag = row.getAttribute('data-tag');
    var g = sec && tag ? sec.querySelector('.t[data-tag="' + esc(tag) + '"]') : null;
    if (g) g.classList.add('hl');
    return row;
  };
  var go = function (q) {
    q = String(q || '').trim();
    if (!q) return;
    var low = q.toLowerCase();
    var sub = null;
    if (prefix && low.indexOf(prefix) === 0 && /^\d+$/.test(low.slice(prefix.length).trim())) {
      sub = document.getElementById('s-' + low.slice(prefix.length).trim());
    } else if (!prefix) {
      sub = document.getElementById('s-' + q.toUpperCase());
    }
    if (sub) { location.hash = sub.id; return; }
    var row = mark(document.querySelector('.trees tr[data-tag="' + esc(q) + '"]'));
    if (row) {
      history.replaceState(null, '', '#' + row.id);
      row.scrollIntoView({ block: 'center' });
      return;
    }
    var f = document.getElementById('find');
    f.setCustomValidity(body.getAttribute('data-not-found') || 'Not found');
    f.reportValidity();
    setTimeout(function () { f.setCustomValidity(''); }, 1500);
  };
  var f = document.getElementById('find');
  if (f) {
    f.addEventListener('keydown', function (e) { if (e.key === 'Enter') go(f.value); });
    f.addEventListener('search', function () { go(f.value); });
  }
  document.addEventListener('click', function (e) {
    var g = e.target.closest('.map .t');
    if (g) {
      var sec = g.closest('.sub');
      var row = sec && sec.querySelector('.trees tr[data-tag="' + esc(g.getAttribute('data-tag')) + '"]');
      if (mark(row)) row.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      return;
    }
    var a = e.target.closest('a[href^="#r"]');
    if (a) { setTimeout(function () { mark(document.getElementById(a.getAttribute('href').slice(1))); }, 0); return; }
    var tr = e.target.closest('.trees tr[id]');
    if (tr) mark(tr);
  });
  if (/^#r\d+$/.test(location.hash)) mark(document.getElementById(location.hash.slice(1)));
})();
