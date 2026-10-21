/* SchoolFlow phone layout: below 760px the sidebar becomes a slide-out drawer
   opened from a menu button in the top bar. */
'use strict';
(function () {
  const open = (on) => document.body.classList.toggle('nav-open', on);
  function ensure() {
    const topbar = document.querySelector('.topbar');
    if (topbar && !topbar.querySelector('.menu-btn')) {
      const btn = document.createElement('button');
      btn.className = 'menu-btn';
      btn.type = 'button';
      btn.setAttribute('aria-label', 'Open menu');
      btn.innerHTML = '<svg viewBox="0 0 24 24"><path d="M4 7h16M4 12h16M4 17h16"/></svg>';
      btn.addEventListener('click', () => open(!document.body.classList.contains('nav-open')));
      topbar.prepend(btn);
    }
    if (!document.querySelector('.nav-scrim')) {
      const scrim = document.createElement('div');
      scrim.className = 'nav-scrim';
      scrim.addEventListener('click', () => open(false));
      document.body.appendChild(scrim);
    }
  }
  new MutationObserver(ensure).observe(document.getElementById('app'), { childList: true });
  document.addEventListener('click', (e) => { if (e.target.closest('.sidebar [data-action="go"]')) open(false); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') open(false); });
  window.matchMedia('(min-width: 761px)').addEventListener('change', (m) => { if (m.matches) open(false); });
  ensure();
})();
