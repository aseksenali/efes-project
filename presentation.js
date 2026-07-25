/* Efes Promo Automation — presentation engine */
(function () {
  'use strict';

  var stage = document.getElementById('stage');
  var viewport = document.getElementById('viewport');
  var slides = Array.prototype.slice.call(document.querySelectorAll('.slide'));
  var total = slides.length;
  var idx = 0;

  /* ---- scale stage to viewport ---- */
  function scaleStage() {
    var pad = 0;
    var s = Math.min(
      (window.innerWidth - pad) / 1280,
      (window.innerHeight - pad) / 720
    );
    s = Math.max(s, 0.1);
    stage.style.transform = 'scale(' + s + ')';
  }
  window.addEventListener('resize', scaleStage);

  /* ---- chrome elements ---- */
  var progress = document.getElementById('progress');
  var counter = document.getElementById('counter');
  var dotsWrap = document.getElementById('dots');
  var hint = document.getElementById('hint');

  /* build dot nav */
  slides.forEach(function (sl, i) {
    var d = document.createElement('div');
    d.className = 'dot-nav';
    d.title = (sl.dataset.title || ('Slide ' + (i + 1)));
    d.addEventListener('click', function () { go(i); });
    dotsWrap.appendChild(d);
  });
  var dots = Array.prototype.slice.call(dotsWrap.children);

  /* ---- navigation ---- */
  function render() {
    slides.forEach(function (sl, i) {
      sl.classList.toggle('active', i === idx);
      sl.classList.toggle('prev', i < idx);
    });
    dots.forEach(function (d, i) { d.classList.toggle('on', i === idx); });
    progress.style.width = ((idx + 1) / total * 100) + '%';
    counter.innerHTML = '<b>' + String(idx + 1).padStart(2, '0') + '</b> / ' + String(total).padStart(2, '0');
    if (location.hash !== '#' + (idx + 1)) {
      history.replaceState(null, '', '#' + (idx + 1));
    }
    syncOverview();
  }
  function go(i) {
    idx = Math.max(0, Math.min(total - 1, i));
    render();
  }
  function next() { if (idx < total - 1) go(idx + 1); }
  function prev() { if (idx > 0) go(idx - 1); }

  document.getElementById('nav-next').addEventListener('click', next);
  document.getElementById('nav-prev').addEventListener('click', prev);

  /* ---- keyboard ---- */
  document.addEventListener('keydown', function (e) {
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    var k = e.key;
    if (k === 'ArrowRight' || k === 'ArrowDown' || k === 'PageDown' || k === ' ') {
      e.preventDefault(); next();
    } else if (k === 'ArrowLeft' || k === 'ArrowUp' || k === 'PageUp') {
      e.preventDefault(); prev();
    } else if (k === 'Home') { e.preventDefault(); go(0); }
    else if (k === 'End') { e.preventDefault(); go(total - 1); }
    else if (k === 'o' || k === 'O' || k === 'Escape') { e.preventDefault(); toggleOverview(); }
    else if (k === 'f' || k === 'F') {
      e.preventDefault();
      if (!document.fullscreenElement) {
        (document.documentElement.requestFullscreen || function () {}).call(document.documentElement);
      } else { document.exitFullscreen(); }
    } else if (/^[0-9]$/.test(k)) {
      var n = parseInt(k, 10);
      if (n >= 1 && n <= total) go(n - 1);
    }
  });

  /* ---- touch swipe ---- */
  var tx = 0, ty = 0;
  viewport.addEventListener('touchstart', function (e) {
    tx = e.changedTouches[0].clientX; ty = e.changedTouches[0].clientY;
  }, { passive: true });
  viewport.addEventListener('touchend', function (e) {
    var dx = e.changedTouches[0].clientX - tx;
    var dy = e.changedTouches[0].clientY - ty;
    if (Math.abs(dx) > 55 && Math.abs(dx) > Math.abs(dy)) {
      if (dx < 0) next(); else prev();
    }
  }, { passive: true });

  /* ---- overview grid ---- */
  var overview = document.getElementById('overview');
  var ovGrid = document.getElementById('ov-grid');
  var ovBuilt = false;
  function buildOverview() {
    slides.forEach(function (sl, i) {
      var cell = document.createElement('div');
      cell.className = 'ov-cell';
      var thumb = document.createElement('div');
      thumb.className = 'ov-thumb';
      var clone = sl.cloneNode(true);
      clone.classList.add('active');
      clone.style.opacity = '1';
      clone.style.visibility = 'visible';
      clone.style.position = 'absolute';
      clone.style.top = '0'; clone.style.left = '0';
      clone.style.transformOrigin = 'top left';
      clone.style.transition = 'none';
      Array.prototype.slice.call(clone.querySelectorAll('.anim, .stagger > *')).forEach(function (e) {
        e.style.opacity = '1'; e.style.transform = 'none'; e.style.transition = 'none';
      });
      var frames = clone.querySelectorAll('iframe');
      Array.prototype.slice.call(frames).forEach(function (f) {
        var ph = document.createElement('div');
        ph.style.cssText = 'width:100%;height:100%;background:#0c1626;display:grid;place-items:center;color:#5ca3db;font-size:40px;font-weight:800';
        ph.textContent = 'LIVE';
        f.parentNode.replaceChild(ph, f);
      });
      thumb.appendChild(clone);
      cell.appendChild(thumb);
      var cap = document.createElement('div');
      cap.className = 'ov-cap';
      cap.innerHTML = '<b>' + String(i + 1).padStart(2, '0') + '</b> ' + (sl.dataset.title || '');
      cell.appendChild(cap);
      cell.addEventListener('click', function () { go(i); toggleOverview(); });
      ovGrid.appendChild(cell);
    });
    ovBuilt = true;
    sizeOverview();
  }
  function sizeOverview() {
    var cells = ovGrid.querySelectorAll('.ov-thumb');
    Array.prototype.slice.call(cells).forEach(function (th) {
      var w = th.clientWidth;
      var clone = th.querySelector('.slide');
      if (clone) { clone.style.transform = 'scale(' + (w / 1280) + ')'; }
    });
  }
  window.addEventListener('resize', function () { if (ovBuilt) sizeOverview(); });
  function syncOverview() {
    if (!ovBuilt) return;
    var cells = ovGrid.querySelectorAll('.ov-cell');
    Array.prototype.slice.call(cells).forEach(function (c, i) {
      c.classList.toggle('cur', i === idx);
    });
  }
  function toggleOverview() {
    if (!ovBuilt) buildOverview();
    overview.classList.toggle('show');
    if (overview.classList.contains('show')) { sizeOverview(); syncOverview(); }
  }
  document.getElementById('ov-close').addEventListener('click', toggleOverview);

  /* ---- init ---- */
  var startHash = parseInt((location.hash || '').replace('#', ''), 10);
  if (startHash >= 1 && startHash <= total) idx = startHash - 1;
  scaleStage();
  render();
  setTimeout(function () { if (hint) hint.style.opacity = '0'; }, 6000);
})();
