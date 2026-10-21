/* SchoolFlow interface. Talks to Python through window.pywebview.api (see
   Program/schoolflow_api.py). Pages are rendered from state with template
   strings; clicks are handled by delegation on [data-action] elements. */
'use strict';

window.__errors = [];
window.addEventListener('error', (e) => window.__errors.push(String(e.message)));
window.addEventListener('unhandledrejection', (e) => window.__errors.push(String(e.reason && e.reason.message || e.reason)));

// ------------------------------------------------------------ Helpers

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const esc = (value) => String(value ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

const ICONS = {
  search: '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
  home: '<path d="M3.5 10.5 12 3.5l8.5 7"/><path d="M5.5 9v11h13V9"/><path d="M10 20v-5.5h4V20"/>',
  sliders: '<path d="M4 6h9M17 6h3M4 12h3M11 12h9M4 18h11M19 18h1"/><circle cx="15" cy="6" r="2"/><circle cx="9" cy="12" r="2"/><circle cx="17" cy="18" r="2"/>',
  calendar: '<rect x="3.5" y="5" width="17" height="15" rx="2"/><path d="M3.5 10h17M8 3v4M16 3v4"/>',
  hash: '<path d="M5 9h14M5 15h14M10 4 8 20M16 4l-2 16"/>',
  sigma: '<path d="M18 6.5V4H6l6.5 8L6 20h12v-2.5"/>',
  clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
  file: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>',
  table: '<rect x="3.5" y="4.5" width="17" height="15" rx="2"/><path d="M3.5 9.5h17M3.5 14.5h17M9.5 9.5v10"/>',
  layers: '<path d="m12 3.5 8.5 4.5-8.5 4.5L3.5 8z"/><path d="m3.5 12.5 8.5 4.5 8.5-4.5"/>',
  type: '<path d="M5 7V4.5h14V7M12 4.5v15M9 19.5h6"/>',
  user: '<circle cx="12" cy="8" r="3.5"/><path d="M5 20a7 7 0 0 1 14 0"/>',
  users: '<circle cx="9" cy="8" r="3.2"/><path d="M3 19.5a6 6 0 0 1 12 0"/><path d="M16 4.8a3.2 3.2 0 0 1 0 6.4M18 13.6a6 6 0 0 1 3 5.9"/>',
  list: '<path d="M9 6h11M9 12h11M9 18h11"/><circle cx="4.5" cy="6" r=".8"/><circle cx="4.5" cy="12" r=".8"/><circle cx="4.5" cy="18" r=".8"/>',
  down: '<path d="m6.5 9.5 5.5 5.5 5.5-5.5"/>',
  up: '<path d="m6.5 14.5 5.5-5.5 5.5 5.5"/>',
  right: '<path d="m9.5 6.5 5.5 5.5-5.5 5.5"/>',
  dright: '<path d="m6 17 5-5-5-5M13 17l5-5-5-5"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  minus: '<path d="M5 12h14"/>',
  x: '<path d="M6.5 6.5l11 11M17.5 6.5l-11 11"/>',
  check: '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
  alert: '<path d="M12 3.5 2.8 19.5h18.4z"/><path d="M12 10v4.5M12 17v.01"/>',
  zap: '<path d="M13 2.5 4.5 13.5H11l-1 8 8.5-11H12z"/>',
  external: '<path d="M14 4h6v6M20 4l-9 9"/><path d="M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/>',
  folder: '<path d="M3.5 7a2 2 0 0 1 2-2h4l2 2.5h7a2 2 0 0 1 2 2V17a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2z"/>',
  refresh: '<path d="M19.5 10.5A7.8 7.8 0 0 0 5.6 7M4.5 3.5V7.5h4"/><path d="M4.5 13.5A7.8 7.8 0 0 0 18.4 17M19.5 20.5v-4h-4"/>',
  scissors: '<circle cx="6" cy="6.5" r="2.5"/><circle cx="6" cy="17.5" r="2.5"/><path d="M8.2 7.8 20 18M8.2 16.2 20 6"/>',
  trash: '<path d="M4.5 7h15M9.5 7V4.5h5V7M6.5 7l.9 12.5h9.2L17.5 7"/>',
  copy: '<rect x="8.5" y="8.5" width="11.5" height="11.5" rx="2"/><path d="M15.5 8.5V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v7.5a2 2 0 0 0 2 2h2.5"/>',
  filter: '<path d="M4 5h16l-6 7.5V19l-4 1.5v-8z"/>',
  sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2.5v2M12 19.5v2M4.6 4.6 6 6M18 18l1.4 1.4M2.5 12h2M19.5 12h2M4.6 19.4 6 18M18 6l1.4-1.4"/>',
  moon: '<path d="M19.5 14.5A8 8 0 0 1 9.5 4.5a8 8 0 1 0 10 10z"/>',
  monitor: '<rect x="3" y="4.5" width="18" height="12" rx="2"/><path d="M8.5 20h7M12 16.5V20"/>',
  arrow: '<path d="M5 12h14M13 6l6 6-6 6"/>',
  pencil: '<path d="M4 20h4L19 9l-4-4L4 16z"/><path d="m13.5 6.5 4 4"/>',
  grid: '<rect x="4" y="4" width="7" height="7" rx="1.5"/><rect x="13" y="4" width="7" height="7" rx="1.5"/><rect x="4" y="13" width="7" height="7" rx="1.5"/><rect x="13" y="13" width="7" height="7" rx="1.5"/>',
  at: '<circle cx="12" cy="12" r="3.5"/><path d="M15.5 12v1.5a2.5 2.5 0 0 0 5 0V12a8.5 8.5 0 1 0-3.3 6.7"/>',
  more: '<circle cx="5.5" cy="12" r="1"/><circle cx="12" cy="12" r="1"/><circle cx="18.5" cy="12" r="1"/>',
  check2: '<path d="M4 12.5 8 16.5 15.5 9M12 16l1 1 7.5-7.5"/>',
  timer: '<circle cx="12" cy="13" r="7.5"/><path d="M12 9.5V13l2 1.5M9.5 2.5h5"/>',
  undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10a6 6 0 0 1 0 12h-3"/>',
};
const icon = (name, cls = '') => `<svg class="i ${cls}" viewBox="0 0 24 24">${ICONS[name] || ''}</svg>`;

const TAG_COLORS = ['blue', 'green', 'orange', 'purple', 'pink', 'yellow', 'brown', 'red', 'gray'];
function colorFor(name) {
  let hash = 2166136261;
  const text = String(name || '').toLowerCase();
  for (let i = 0; i < text.length; i++) { hash ^= text.charCodeAt(i); hash = Math.imul(hash, 16777619); }
  return TAG_COLORS[Math.abs(hash) % TAG_COLORS.length];
}
const initials = (name) => String(name || '').split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0]).join('').toUpperCase();
const avatar = (name) => `<span class="avatar b-${colorFor(name)}">${esc(initials(name))}</span>`;
const basename = (path) => String(path || '').split(/[\\/]/).pop();
// Escape a path and let it wrap only after folder separators.
const pathHtml = (path) => esc(path).replace(/\\/g, '\\<wbr>');
// Keep the last folder and file name of a long path: "…\Exports\school_timetable.xlsx".
const shortPath = (path) => { const parts = String(path || '').split(/[\\/]/); return parts.length > 3 ? '…\\' + parts.slice(-2).join('\\') : String(path || ''); };
const plural = (n, word, many) => `${n} ${n === 1 ? word : (many || word + 's')}`;

function when(seconds) {
  if (!seconds) return '';
  const date = new Date(seconds * 1000);
  const now = new Date();
  const time = date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  if (date.toDateString() === now.toDateString()) return `Today at ${time}`;
  const yesterday = new Date(now); yesterday.setDate(now.getDate() - 1);
  if (date.toDateString() === yesterday.toDateString()) return `Yesterday at ${time}`;
  return `${date.toLocaleDateString([], { day: 'numeric', month: 'short', year: 'numeric' })}, ${time}`;
}
const clock = (s) => `${Math.floor(s / 60)}:${String(Math.floor(s % 60)).padStart(2, '0')}`;

// ------------------------------------------------------------ Python bridge

let API = null;
async function call(name, ...args) {
  const result = await API[name](...args);
  if (result && result.ok === false) throw new Error(result.error || 'Something went wrong.');
  return result;
}

// ------------------------------------------------------------ State

const S = {
  boot: null,
  settings: null,
  page: 'home',
  summary: null,
  req: { loaded: false, rows: [], fieldnames: [], path: '', dirty: false, changed: new Set(), grade: 'All', division: 'All', q: '', view: 'table', peek: null },
  gen: { state: 'idle', log: [], logCount: 0, result: null, elapsed: 0, timer: null },
  tt: { data: null, mode: 'classes', current: { classes: null, teachers: null }, q: '', loading: false },
};

const PAGES = {
  home: { name: 'Home', icon: 'home' },
  requirements: { name: 'Requirements', icon: 'table' },
  timetables: { name: 'Timetables', icon: 'calendar' },
  settings: { name: 'Settings', icon: 'sliders' },
};
const ORDER = ['home', 'requirements', 'timetables', 'settings'];

// ------------------------------------------------------------ Theme

const darkQuery = window.matchMedia('(prefers-color-scheme: dark)');
function applyTheme() {
  const pref = S.settings?.theme || 'system';
  const dark = pref === 'dark' || (pref === 'system' && darkQuery.matches);
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
}
darkQuery.addEventListener('change', applyTheme);

async function saveSettings(patch) {
  Object.assign(S.settings, patch);
  const result = await call('save_settings', patch);
  S.settings = result.settings;
}

// ------------------------------------------------------------ Shell

function renderShell() {
  $('#app').innerHTML = `
    <aside class="sidebar">
      <div class="workspace" data-action="go" data-page="home">
        <div class="mark">S</div>
        <div><div class="name">SchoolFlow</div><div class="sub">Timetable planner</div></div>
      </div>
      <nav class="nav">
        <div class="nav-item" data-action="go" data-page="home" data-nav="home"><span class="emoji">${icon('home')}</span><span class="label">Home</span><span class="side" id="nav-gen"></span></div>
        <div class="nav-item" data-action="go" data-page="requirements" data-nav="requirements"><span class="emoji">${icon('table')}</span><span class="label">Requirements</span><span class="side" id="nav-req"></span></div>
        <div class="nav-item" data-action="go" data-page="timetables" data-nav="timetables"><span class="emoji">${icon('calendar')}</span><span class="label">Timetables</span></div>
        <div class="nav-item" data-action="go" data-page="settings" data-nav="settings"><span class="emoji">${icon('sliders')}</span><span class="label">Settings</span></div>
      </nav>
      <div class="sidebar-foot">
        <span class="grow">${esc({ source: 'Source', portable: 'Portable', installed: 'Installed', web: 'Web', online: 'Online' }[S.boot.mode] || S.boot.mode)} · v${esc(S.boot.version)}</span>
        <button class="btn ghost icon" data-action="cycle-theme" title="Change theme" id="theme-btn"></button>
      </div>
    </aside>
    <main class="main">
      <header class="topbar"><div class="crumbs" id="crumbs"></div><div class="status-pill" id="status"></div></header>
      <div class="view">
        ${ORDER.map((key) => `<section class="page ${['requirements', 'timetables'].includes(key) ? 'fill' : ''}" id="page-${key}"></section>`).join('')}
      </div>
    </main>`;
  renderNavMeta();
  renderThemeButton();
}

function renderThemeButton() {
  const pref = S.settings.theme;
  $('#theme-btn').innerHTML = icon(pref === 'dark' ? 'moon' : pref === 'light' ? 'sun' : 'monitor');
}

function renderNavMeta() {
  const req = $('#nav-req');
  if (req) req.innerHTML = S.req.dirty ? '<span class="dot" title="Unsaved changes"></span>' : (S.summary?.requirements ? `<span class="count">${S.summary.requirements.rows}</span>` : '');
  const gen = $('#nav-gen');
  if (gen) gen.innerHTML = S.gen.state === 'running' ? '<span class="dot busy"></span>' : '';
  const status = $('#status');
  if (!status) return;
  if (S.gen.state === 'running') {
    status.className = 'status-pill busy';
    status.innerHTML = `<span class="spinner"></span>Generating · ${clock(S.gen.elapsed)}`;
  } else if (S.req.dirty) {
    status.className = 'status-pill';
    status.innerHTML = `<span class="unsaved">Unsaved changes</span>`;
  } else {
    status.className = 'status-pill';
    status.innerHTML = S.summary?.workbook ? `Generated ${esc(when(S.summary.workbook.modified)).replace(/^Today at/, 'today at')}` : '';
  }
}

function renderCrumbs() {
  const page = PAGES[S.page];
  const parts = [`<span class="crumb ${pageCrumb() ? '' : 'current'}" data-action="go" data-page="${S.page}">${icon(page.icon)}${esc(page.name)}</span>`];
  const extra = pageCrumb();
  if (extra) parts.push(`<span class="crumb-sep">/</span><span class="crumb current">${esc(extra)}</span>`);
  $('#crumbs').innerHTML = parts.join('');
  document.title = (extra ? extra + ' · ' : '') + page.name + ' — SchoolFlow';
}
function pageCrumb() {
  if (S.page === 'requirements' && S.req.peek) return S.req.peek.values.subject || 'New row';
  if (S.page === 'timetables' && S.tt.data && !S.tt.data.missing) {
    const grid = currentGrid();
    if (grid) return S.tt.mode === 'classes' ? `Class ${grid.label}` : grid.label;
  }
  return '';
}

async function go(page, opts = {}) {
  if (page === 'generate') page = 'home';  // Generate lives on Home
  if (!PAGES[page]) return;
  closeMenu();
  S.page = page;
  $$('.nav-item[data-nav]').forEach((el) => el.classList.toggle('active', el.dataset.nav === page));
  $$('.page').forEach((el) => el.classList.toggle('active', el.id === `page-${page}`));
  renderCrumbs();
  await RENDER[page](opts);
  renderCrumbs();
}

async function refreshSummary() {
  try { S.summary = await call('summary'); } catch (error) { S.summary = null; }
  renderNavMeta();
}

// ------------------------------------------------------------ Menus, modals, toasts

let openMenuState = null;
function closeMenu() {
  if (openMenuState) { openMenuState.el.remove(); openMenuState.onClose?.(); openMenuState = null; }
}

/* items: {label, value, emoji, icon, avatar, hint, checked, disabled, danger, header, sep, html} */
function openMenu(anchor, items, onPick, opts = {}) {
  closeMenu();
  const el = document.createElement('div');
  el.className = 'menu ' + (opts.className || '');
  if (opts.width) el.style.width = opts.width + 'px';
  el.innerHTML = (opts.header || '') + `<div class="ml">${menuItems(items)}</div>` + (opts.footer || '');
  document.body.appendChild(el);
  positionPopover(el, anchor, opts.align);
  openMenuState = { el, items, onPick, onClose: opts.onClose };
  el.addEventListener('mousedown', (e) => e.stopPropagation());
  el.addEventListener('click', (e) => {
    const item = e.target.closest('.mi');
    if (!item || item.classList.contains('disabled')) return;
    const index = Number(item.dataset.i);
    const picked = openMenuState.items[index];
    closeMenu();
    onPick(picked.value, picked);
  });
  return el;
}
function menuItems(items) {
  return items.map((item, i) => {
    if (item.sep) return '<div class="msep"></div>';
    if (item.header) return `<div class="mh">${item.html || esc(item.header)}</div>`;
    const lead = item.icon ? icon(item.icon) : item.tagColor ? `<span class="dot-c b-${item.tagColor}"></span>` : '';
    return `<div class="mi ${item.disabled ? 'disabled' : ''} ${item.danger ? 'danger' : ''} ${item.className || ''}" data-i="${i}">
      ${lead}<span class="grow">${item.html || esc(item.label)}</span>
      ${item.hint ? `<span class="faint small">${esc(item.hint)}</span>` : ''}${item.checked ? `<span class="check">${icon('check')}</span>` : ''}</div>`;
  }).join('');
}
function positionPopover(el, anchor, align = 'left') {
  const rect = anchor.getBoundingClientRect();
  const width = el.offsetWidth;
  const height = el.offsetHeight;
  let left = align === 'right' ? rect.right - width : rect.left;
  let top = rect.bottom + 4;
  if (top + height > window.innerHeight - 10) top = Math.max(10, rect.top - height - 4);
  left = Math.max(10, Math.min(left, window.innerWidth - width - 10));
  el.style.left = left + 'px';
  el.style.top = top + 'px';
}
document.addEventListener('mousedown', (e) => { if (openMenuState && !openMenuState.el.contains(e.target)) closeMenu(); });

function modal(html, opts = {}) {
  const scrim = document.createElement('div');
  scrim.className = 'scrim';
  scrim.innerHTML = `<div class="modal ${opts.wide ? 'wide' : ''}">${html}</div>`;
  document.body.appendChild(scrim);
  const close = () => { scrim.remove(); modalStack.splice(modalStack.indexOf(close), 1); opts.onClose?.(); };
  modalStack.push(close);
  scrim.addEventListener('mousedown', (e) => { if (e.target === scrim) close(); });
  return { el: scrim.firstElementChild, close };
}
const modalStack = [];

function confirmDialog({ title, text, confirm = 'Confirm', danger = false }) {
  return new Promise((resolve) => {
    let done = false;
    const m = modal(`
      <div class="modal-body"><h3>${esc(title)}</h3><p>${esc(text)}</p></div>
      <div class="modal-foot"><button class="btn secondary" data-r="0">Cancel</button><button class="btn ${danger ? 'danger-solid' : 'primary'}" data-r="1">${esc(confirm)}</button></div>`,
      { onClose: () => { if (!done) resolve(false); } });
    m.el.addEventListener('click', (e) => {
      const b = e.target.closest('[data-r]');
      if (!b) return;
      done = true; m.close(); resolve(b.dataset.r === '1');
    });
    setTimeout(() => m.el.querySelector('[data-r="1"]').focus(), 20);
  });
}

function infoDialog({ title, text, tone = '' }) {
  const m = modal(`<div class="modal-body"><h3>${esc(title)}</h3>
    <div class="${tone ? `callout ${tone}` : ''}" style="${tone ? 'margin-top:10px' : ''}"><p style="color:inherit" class="selectable">${esc(text)}</p></div></div>
    <div class="modal-foot"><button class="btn primary" data-close>OK</button></div>`);
  m.el.querySelector('[data-close]').addEventListener('click', m.close);
  setTimeout(() => m.el.querySelector('[data-close]').focus(), 20);
}

function toast(text, kind = '', opts = {}) {
  const el = document.createElement('div');
  el.className = 'toast';
  el.innerHTML = `${['⚠️', '⛔'].includes(kind) ? `<span class="te">${icon('alert')}</span>` : ''}<span>${esc(text)}</span>${opts.action ? `<button class="btn">${esc(opts.action.label)}</button>` : ''}`;
  if (opts.action) el.querySelector('.btn').addEventListener('click', () => { opts.action.run(); dismiss(); });
  $('#toasts').appendChild(el);
  const dismiss = () => { el.classList.add('leaving'); setTimeout(() => el.remove(), 220); };
  setTimeout(dismiss, opts.duration || 2800);
}
const fail = (error) => toast(error.message || String(error), '⚠️', { duration: 5200 });

// ------------------------------------------------------------ Home

const RENDER = {};

function prop(iconName, name, value, opts = {}) {
  return `<div class="prop"><div class="prop-name">${icon(iconName)}${esc(name)}</div>
    <div class="prop-value ${opts.click ? 'click' : ''}" ${opts.attrs || ''}>${value}</div></div>`;
}

// ------------------------------------------------------------ Requirements

const REQ_FIELDS = ['class', 'division', 'subject', 'teacher', 'periods_per_week'];

async function loadRequirements(force = false) {
  if (S.req.loaded && !force) return;
  const data = await call('requirements');
  S.req.rows = data.rows;
  S.req.fieldnames = data.fieldnames;
  S.req.path = data.path;
  S.req.missing = !!data.missing;
  S.req.changed = new Set();
  S.req.loaded = true;
  S.req.peek = null;
  setDirty(false);
}

function setDirty(dirty) {
  S.req.dirty = dirty;
  API.set_dirty(dirty);
  renderNavMeta();
  renderReqStats();
  const save = $('#req-save');
  if (save) { save.className = `btn sm ${dirty ? 'primary' : 'secondary'}`; save.disabled = !dirty; }
  const foot = $('#req-unsaved');
  if (foot) foot.innerHTML = dirty ? `<span class="unsaved">${plural(S.req.changed.size, 'unsaved change')}</span>` : '';
}

function visibleRows() {
  const { rows, grade, division, q } = S.req;
  const term = q.trim().toLowerCase();
  const out = [];
  rows.forEach((row, index) => {
    if (grade !== 'All' && row.class !== grade) return;
    if (division !== 'All' && row.division !== division) return;
    if (term && !`${row.subject} ${row.teacher} ${row.class}-${row.division}`.toLowerCase().includes(term)) return;
    out.push(index);
  });
  return out;
}

RENDER.requirements = async () => {
  const page = $('#page-requirements');
  try { await loadRequirements(); } catch (error) { fail(error); }
  page.innerHTML = `
    <div class="db-head req-head">
      <div class="req-top">
        <div>
          <div class="eyebrow">School data</div>
          <div class="db-title"><h1>Requirements</h1></div>
          <div class="db-lead">Every subject a class division takes, who teaches it, and how many periods a week.</div>
          <div class="req-file">
            <span class="rf-label">Requirements file</span>
            <span class="chip" title="${esc(S.req.path)}">${icon('file', 'sm')}${esc(basename(S.req.path))}</span>
            <button class="btn accent sm" data-action="choose-req">${icon('folder', 'sm')}Change file…</button>
          </div>
        </div>
        <div class="req-stats" id="req-stats"></div>
      </div>
      <div class="grade-pills" id="grade-pills"></div>
    </div>
    <div class="toolbar">
      <div class="tabs">
        <div class="tab ${S.req.view === 'table' ? 'on' : ''}" data-action="req-view" data-view="table">${icon('table')}All rows <span class="n" id="req-n"></span></div>
        <div class="tab ${S.req.view === 'group' ? 'on' : ''}" data-action="req-view" data-view="group">${icon('layers')}By class</div>
      </div>
      <div class="spacer"></div>
      <div class="select" id="f-div" data-action="filter-div"></div>
      <label class="search">${icon('search', 'sm')}<input id="req-q" placeholder="Search…" value="${esc(S.req.q)}"></label>
      <div class="sep"></div>
      <button class="btn ghost sm" data-action="check-req">${icon('check2')}Check data</button>
      <button class="btn sm secondary" id="req-save" data-action="save-req">Save</button>
      <button class="btn sm primary" data-action="new-row">${icon('plus')}New</button>
    </div>
    <div class="db-body"><div class="db-scroll" id="req-scroll"></div><div id="peek-host"></div></div>`;
  $('#req-q').addEventListener('input', (e) => { S.req.q = e.target.value; renderReqTable(); });
  renderFilters();
  renderReqTable();
  renderPeek();
  setDirty(S.req.dirty);
};

function renderFilters() {
  const d = $('#f-div');
  if (!d) return;
  d.className = `select ${S.req.division !== 'All' ? 'active-filter' : ''}`;
  d.innerHTML = `${icon('filter', 'sm')}${S.req.division === 'All' ? 'Division' : `Division ${esc(S.req.division)}`}${icon('down', 'sm')}`;
  const pills = $('#grade-pills');
  if (pills) {
    pills.innerHTML = `<span class="gp-label">Class</span>` + ['All', ...S.boot.classes].map((g) =>
      `<button class="grade-pill ${S.req.grade === g ? 'on' : ''}" data-action="grade-pill" data-v="${g}">${g === 'All' ? 'All classes' : g}</button>`).join('');
  }
}

// Live summary of the rows being edited, including teachers given two subjects.
function renderReqStats() {
  const host = $('#req-stats');
  if (!host) return;
  const rows = S.req.rows;
  const subjectsByTeacher = new Map();
  rows.forEach((r) => {
    const key = r.teacher.toLowerCase();
    if (!subjectsByTeacher.has(key)) subjectsByTeacher.set(key, new Set());
    subjectsByTeacher.get(key).add(r.subject.toLowerCase());
  });
  const conflicts = [...subjectsByTeacher.values()].filter((s) => s.size > 1).length;
  const tile = (tint, iconName, value, label, extra = '') =>
    `<div class="rq-tile ${extra}"><div class="chip-i tint-${tint}">${icon(iconName, 'lg')}</div><div><b>${value}</b><span>${label}</span></div></div>`;
  host.innerHTML =
    tile('blue', 'table', rows.length, 'rows') +
    tile('purple', 'grid', new Set(rows.map((r) => `${r.class}-${r.division}`)).size, 'divisions') +
    tile('green', 'users', subjectsByTeacher.size, 'teachers') +
    tile('orange', 'layers', new Set(rows.map((r) => r.subject.toLowerCase())).size, 'subjects') +
    (conflicts
      ? tile('orange', 'alert', conflicts, conflicts === 1 ? 'teacher has 2+ subjects' : 'teachers have 2+ subjects', 'warn')
      : tile('green', 'check', 'All good', 'one subject per teacher', 'ok'));
}

const maxPeriods = () => Math.max(1, ...S.req.rows.map((r) => Number(r.periods_per_week) || 0));

function rowCells(row, max) {
  const n = Number(row.periods_per_week) || 0;
  return `<td class="title"><span class="subj-pill b-${colorFor(row.subject)}">${esc(row.subject)}</span></td>
    <td><span class="cls-chip">${esc(row.class)}</span></td><td><span class="cls-chip div">${esc(row.division)}</span></td>
    <td><div class="cellflex">${avatar(row.teacher)}<span>${esc(row.teacher)}</span></div></td>
    <td class="num"><div class="pcell"><span class="pb"><i style="width:${Math.round((n / max) * 100)}%"></i></span><b>${esc(row.periods_per_week)}</b></div></td>`;
}

function renderReqTable() {
  const host = $('#req-scroll');
  if (!host) return;
  const indexes = visibleRows();
  $('#req-n').textContent = indexes.length;
  if (S.req.missing) {
    host.innerHTML = `<div class="empty-state"><div><h3>No requirements file</h3><p>${esc(S.req.path)} was not found. Upload a CSV or restore the sample data in Settings.</p>
      <button class="btn primary" data-action="choose-req">Upload a CSV…</button></div></div>`;
    return;
  }
  const sel = S.req.peek?.index;
  const head = `<thead><tr>
    <th style="width:34%"><div class="th">Subject</div></th>
    <th style="width:9%"><div class="th">Class</div></th>
    <th style="width:10%"><div class="th">Division</div></th>
    <th><div class="th">Teacher</div></th>
    <th class="num" style="width:14%"><div class="th">Periods / week</div></th></tr></thead>`;
  let body = '';
  const max = maxPeriods();
  const capacity = S.settings.days.length * S.settings.periods_per_day;
  const rowHtml = (i) => `<tr data-action="open-row" data-i="${i}" class="${i === sel ? 'sel' : ''} ${S.req.changed.has(i) ? 'changed' : ''}">${rowCells(S.req.rows[i], max)}</tr>`;
  if (S.req.view === 'group') {
    const groups = new Map();
    indexes.forEach((i) => {
      const r = S.req.rows[i];
      const key = `${r.class}-${r.division}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(i);
    });
    const sorted = [...groups.entries()].sort((a, b) => {
      const [ga, da] = a[0].split('-'); const [gb, db] = b[0].split('-');
      return (Number(ga) - Number(gb)) || da.localeCompare(db);
    });
    for (const [key, list] of sorted) {
      const total = list.reduce((sum, i) => sum + (Number(S.req.rows[i].periods_per_week) || 0), 0);
      const fits = total === capacity;
      body += `<tr class="group"><td colspan="5"><div class="gl"><span class="gname">Class ${esc(key)}</span><span class="gc">${plural(list.length, 'subject')} · ${total} periods</span>
        <span class="gmeter"><i class="${fits ? '' : 'adj'}" style="width:${Math.min(100, Math.round((total / capacity) * 100))}%"></i></span>
        <span class="gbadge ${fits ? 'ok' : 'adj'}" title="${fits ? 'Exactly fills the week' : `The week has ${capacity} slots; repeats are adjusted automatically`}">${fits ? `Fills ${capacity}` : `${total > capacity ? '+' : ''}${total - capacity} vs ${capacity}`}</span></div></td></tr>`;
      body += list.map(rowHtml).join('');
    }
  } else {
    body = indexes.map(rowHtml).join('');
  }
  const empty = indexes.length ? '' : `<div class="empty-state" style="height:auto;padding:48px"><div><h3>No matching rows</h3><p>Try a different search or clear the filters.</p></div></div>`;
  host.innerHTML = `<table class="grid">${head}<tbody>${body}</tbody></table>${empty}
    <div class="new-row" data-action="new-row">${icon('plus', 'sm')}New</div>
    <div class="table-foot"><span><b>COUNT</b>&nbsp; ${indexes.length}${indexes.length !== S.req.rows.length ? ` of ${S.req.rows.length}` : ''}</span><span id="req-unsaved"></span></div>`;
  setDirty(S.req.dirty);
}

function openPeek(index, values) {
  if (values === undefined) {
    const row = S.req.rows[index];
    values = Object.fromEntries(REQ_FIELDS.map((f) => [f, row[f] || '']));
  }
  S.req.peek = { index, values: { periods_per_week: '1', ...values } };
  renderPeek();
  renderReqTable();
  renderCrumbs();
}
function closePeek() {
  S.req.peek = null;
  renderPeek();
  renderReqTable();
  renderCrumbs();
}

function teacherSubjects(teacher, ignoreIndex) {
  const key = teacher.trim().toLowerCase();
  const subjects = new Set();
  S.req.rows.forEach((r, i) => { if (i !== ignoreIndex && r.teacher.toLowerCase() === key && r.subject) subjects.add(r.subject); });
  return [...subjects];
}

function renderPeek() {
  const host = $('#peek-host');
  if (!host) return;
  const p = S.req.peek;
  host.parentElement.classList.toggle('has-peek', !!p);
  if (!p) { host.innerHTML = ''; return; }
  const v = p.values;
  const isNew = p.index === null;
  host.innerHTML = `
    <div class="peek">
      <div class="peek-bar">
        <button class="btn ghost icon" data-action="close-peek" title="Close (Esc)">${icon('dright')}</button>
        <button class="btn ghost icon" data-action="peek-step" data-step="-1" title="Previous row" ${isNew ? 'disabled' : ''}>${icon('up')}</button>
        <button class="btn ghost icon" data-action="peek-step" data-step="1" title="Next row" ${isNew ? 'disabled' : ''}>${icon('down')}</button>
        <span class="grow"></span>
        <button class="btn ghost sm" data-action="dup-row" ${isNew ? 'disabled' : ''}>${icon('copy')}Duplicate</button>
        <button class="btn ghost icon" data-action="delete-row" title="Delete row" ${isNew ? 'disabled' : ''}>${icon('trash')}</button>
      </div>
      <div class="peek-body">
        <div class="pk-icon b-${colorFor(v.subject)}" id="pk-icon">${icon('layers', 'xl')}</div>
        <input class="title-input" id="pk-subject" placeholder="Subject name" value="${esc(v.subject)}" autocomplete="off">
        <div class="suggest" id="pk-suggest"></div>
        <div class="props">
          ${prop('list', 'Class', `<div class="select bordered" data-action="pk-class">${v.class ? `Class ${esc(v.class)}` : '<span class="placeholder">Choose</span>'}${icon('down', 'sm')}</div>`)}
          ${prop('list', 'Division', `<div class="select bordered" data-action="pk-div">${v.division ? esc(v.division) : '<span class="placeholder">Choose</span>'}${icon('down', 'sm')}</div>`)}
          ${prop('user', 'Teacher', `<div class="combo"><input id="pk-teacher" placeholder="Type or pick a teacher" value="${esc(v.teacher)}" autocomplete="off"></div>`)}
          ${prop('hash', 'Periods / week', `<div class="stepper"><button data-action="pk-step" data-d="-1">${icon('minus', 'sm')}</button><span class="val" id="pk-periods">${esc(v.periods_per_week)}</span><button data-action="pk-step" data-d="1">${icon('plus', 'sm')}</button></div>`)}
        </div>
        <div id="pk-warning"></div>
      </div>
      <div class="peek-foot">
        <span class="faint small">Ctrl+Enter to ${isNew ? 'add' : 'apply'}</span><span class="grow"></span>
        <button class="btn secondary" data-action="close-peek">Cancel</button>
        <button class="btn primary" data-action="apply-peek">${isNew ? 'Add row' : 'Apply'}</button>
      </div>
    </div>`;
  const subject = $('#pk-subject');
  subject.addEventListener('input', () => {
    v.subject = subject.value;
    $('#pk-icon').className = `pk-icon b-${colorFor(v.subject)}`;
    renderPeekHints(); renderCrumbs();
  });
  const teacher = $('#pk-teacher');
  teacher.addEventListener('input', () => { v.teacher = teacher.value; renderPeekHints(); teacherMenu(teacher); });
  teacher.addEventListener('focus', () => teacherMenu(teacher));
  renderPeekHints();
  if (isNew) setTimeout(() => subject.focus(), 30);
}

function teacherMenu(input) {
  const term = input.value.trim().toLowerCase();
  const subjectByTeacher = new Map();
  S.req.rows.forEach((r) => { if (r.teacher && !subjectByTeacher.has(r.teacher)) subjectByTeacher.set(r.teacher, r.subject); });
  let names = [...subjectByTeacher.keys()].sort((a, b) => a.localeCompare(b));
  const subject = S.req.peek.values.subject.trim().toLowerCase();
  // Teachers of the chosen subject first.
  names.sort((a, b) => ((subjectByTeacher.get(b) || '').toLowerCase() === subject) - ((subjectByTeacher.get(a) || '').toLowerCase() === subject));
  names = names.filter((n) => !term || n.toLowerCase().includes(term)).slice(0, 40);
  if (!names.length || (names.length === 1 && names[0].toLowerCase() === term)) { closeMenu(); return; }
  const host = input.parentElement;
  openMenu(host, names.map((n) => ({ label: n, value: n, hint: subjectByTeacher.get(n) })), (value) => {
    input.value = value; S.req.peek.values.teacher = value; renderPeekHints();
  }, { width: Math.max(280, host.offsetWidth) });
  setTimeout(() => input.focus(), 0);
}

function renderPeekHints() {
  const p = S.req.peek;
  if (!p) return;
  const v = p.values;
  const known = [...new Set(S.req.rows.map((r) => r.subject).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  const term = v.subject.trim().toLowerCase();
  let list = term ? known.filter((s) => s.toLowerCase().startsWith(term) && s.toLowerCase() !== term) : known;
  list = list.slice(0, 8);
  $('#pk-suggest').innerHTML = list.map((s) => `<span class="tag click c-gray" data-action="pk-pick-subject" data-v="${esc(s)}">${esc(s)}</span>`).join('');
  const clash = v.teacher.trim() ? teacherSubjects(v.teacher, p.index).filter((s) => s.toLowerCase() !== term) : [];
  $('#pk-warning').innerHTML = clash.length ? `<div class="callout yellow" style="margin-top:18px"><div class="cb"><div class="ct">One subject per teacher</div>
    ${esc(v.teacher)} already teaches ${esc(clash.join(', '))}. The scheduler will reject this row.</div></div>` : '';
}

function applyPeek() {
  const p = S.req.peek;
  if (!p) return;
  const v = Object.fromEntries(REQ_FIELDS.map((f) => [f, String(p.values[f] || '').trim()]));
  v.division = v.division.toUpperCase();
  const missing = REQ_FIELDS.filter((f) => !v[f]).map((f) => f.replace(/_/g, ' '));
  if (missing.length) return toast('Fill in: ' + missing.join(', '));
  const n = Number(v.periods_per_week);
  if (!Number.isInteger(n) || n < 1 || n > 99) return toast('Periods per week must be 1–99.');
  const key = `${v.class}|${v.division}|${v.subject.toLowerCase()}|${v.teacher.toLowerCase()}`;
  const dup = S.req.rows.findIndex((r, i) => i !== p.index && `${r.class}|${r.division}|${r.subject.toLowerCase()}|${r.teacher.toLowerCase()}` === key);
  if (dup >= 0) return toast('That class, subject and teacher already has a row.', '⚠️');
  let index = p.index;
  if (index === null) {
    const row = Object.fromEntries(S.req.fieldnames.map((f) => [f, '']));
    S.req.rows.push(Object.assign(row, v));
    index = S.req.rows.length - 1;
    toast(`Added ${v.subject} for ${v.class}-${v.division}`);
  } else {
    Object.assign(S.req.rows[index], v);
  }
  S.req.changed.add(index);
  setDirty(true);
  openPeek(index);
  const tr = $(`tr[data-i="${index}"]`);
  tr?.scrollIntoView({ block: 'nearest' });
}

async function deleteRow() {
  const p = S.req.peek;
  if (!p || p.index === null) return;
  const row = S.req.rows[p.index];
  const ok = await confirmDialog({ title: 'Delete this row?', text: `${row.subject} · ${row.teacher} · Class ${row.class}-${row.division}\nThe file changes when you press Save.`, confirm: 'Delete', danger: true });
  if (!ok) return;
  const removed = S.req.rows.splice(p.index, 1)[0];
  const at = p.index;
  S.req.changed = new Set([...S.req.changed].filter((i) => i !== at).map((i) => (i > at ? i - 1 : i)));
  setDirty(true);
  closePeek();
  toast('Row deleted', '', { action: { label: 'Undo', run: () => { S.req.rows.splice(at, 0, removed); S.req.changed.add(at); setDirty(true); renderReqTable(); } } });
}

async function saveRequirements() {
  if (!S.req.dirty) return true;
  try {
    await call('save_requirements', S.req.fieldnames, S.req.rows);
    S.req.changed = new Set();
    setDirty(false);
    renderReqTable();
    refreshSummary();
    toast('Requirements saved');
    return true;
  } catch (error) { fail(error); return false; }
}

async function checkRequirements() {
  const button = $('[data-action="check-req"]');
  button.classList.add('loading'); button.innerHTML = '<span class="spinner"></span>Checking…';
  try {
    const r = await call('check_requirements', S.req.fieldnames, S.req.rows);
    infoDialog({ title: 'Data looks good', text: `Valid for ${S.settings.days.length} days × ${S.settings.periods_per_day} periods.\n\n${r.lessons} lessons to schedule · ${r.teachers} teachers · ${r.adjusted_divisions} divisions with automatic repeat adjustments.` });
  } catch (error) {
    infoDialog({ title: 'Data needs attention', text: error.message, tone: 'yellow' });
  } finally {
    button.classList.remove('loading'); button.innerHTML = `${icon('check2')}Check data`;
  }
}

async function guardUnsaved() {
  if (!S.req.dirty) return true;
  const ok = await confirmDialog({ title: 'Save requirement changes?', text: `You have ${plural(S.req.changed.size, 'unsaved change')}. They need to be saved before generating.`, confirm: 'Save and continue' });
  return ok && saveRequirements();
}

// ------------------------------------------------------------ Generate

function minPeriods(dayCount) { return S.boot.minPeriods[String(dayCount)] ?? 12; }

RENDER.home = async () => {
  const page = $('#page-home');
  if (!S.summary) await refreshSummary();
  const st = S.settings;
  const req = S.summary?.requirements;
  const wb = S.summary?.workbook;
  const hour = new Date().getHours();
  const hello = hour < 12 ? 'Good morning' : hour < 17 ? 'Good afternoon' : 'Good evening';
  const min = minPeriods(st.days.length);
  const capacity = st.days.length * st.periods_per_day;
  const rule = st.days.length === 2 ? 'Two school days need exactly 12 periods' : `${st.days.length} days · choose ${min}–12 periods`;
  // Week preview: one column per weekday, one row per period; selected days light up.
  let preview = DAYS.map((d) => `<span class="wh2 ${st.days.includes(d) ? 'on' : ''}">${d.slice(0, 2)}</span>`).join('');
  for (let p = 0; p < st.periods_per_day; p++) {
    preview += DAYS.map((d) => `<i class="${st.days.includes(d) ? 'on' : ''}"></i>`).join('');
  }
  const card = (tint, iconName, title, sub, body) => `
    <div class="set-card"><div class="sc-h"><div class="chip-i tint-${tint}">${icon(iconName, 'lg')}</div>
      <div><div class="sc-t">${title}</div><div class="sc-d">${sub}</div></div></div>${body}</div>`;
  page.innerHTML = `
    <div class="doc">
      <section class="hero gen-hero">
        <div class="hm">
          <div class="hello">${hello} 👋</div>
          <h1 class="page-title">School timetable</h1>
          <p class="page-lead">Pick the school week, then build clash-free timetables for every class and teacher in one go.</p>
          <div class="ha">
            <button class="btn primary big" id="gen-btn" data-action="generate">${icon('zap')}Generate timetable</button>
            <button class="btn secondary big" data-action="go" data-page="timetables" ${wb ? '' : 'disabled'}>${icon('calendar')}Browse timetables</button>
            <span class="faint">or press <span class="kbd">F5</span></span>
          </div>
          <div class="hero-meta">${wb ? `Last generated ${esc(when(wb.modified)).replace(/^Today/, 'today').replace(/^Yesterday/, 'yesterday')} · <span class="link" data-action="open-path" data-path="${esc(wb.path)}">${esc(basename(wb.path))}</span>` : 'No timetable generated yet'}</div>
        </div>
        <div class="week-card">
          <div class="wl">Week preview</div>
          <div class="wn">${capacity} <small>periods / division</small></div>
          <div class="wkgrid">${preview}</div>
        </div>
      </section>

      <div class="cols wide-left">
        <div class="gen-settings">
          ${card('blue', 'calendar', 'School days', rule,
            `<div class="day-pills">${DAYS.map((d) => `<button class="day-pill ${st.days.includes(d) ? 'on' : ''}" data-action="toggle-day" data-day="${d}" title="${d}">${d.slice(0, 3)}</button>`).join('')}</div>`)}
          <div class="set-row2">
            ${card('purple', 'hash', 'Periods per day', `${min}–12 for ${st.days.length} days`,
              `<div class="big-stepper"><button data-action="periods-step" data-d="-1" ${st.periods_per_day <= min ? 'disabled' : ''}>${icon('minus')}</button>
                <span class="val">${st.periods_per_day}</span>
                <button data-action="periods-step" data-d="1" ${st.periods_per_day >= 12 ? 'disabled' : ''}>${icon('plus')}</button></div>`)}
            ${card('orange', 'timer', 'Search time limit', 'How long the solver may look',
              `<div class="select bordered" data-action="pick-limit">${st.solver_time_limit} seconds${icon('down', 'sm')}</div>`)}
          </div>
          ${card('green', 'file', 'Files', 'What goes in and where the workbook is saved', `
            <div class="out-row"><span class="k">Requirements</span><span class="link path" data-action="go" data-page="requirements">${esc(basename(st.input_path))}</span></div>
            <div class="out-row"><span class="k">Save to</span><span class="link path" data-action="open-folder" title="${esc(st.output_path)}">${esc(shortPath(st.output_path))}</span><button class="btn accent sm" data-action="choose-output">${icon('pencil', 'sm')}Rename…</button></div>
            <div class="out-row"><span class="k">Download when done</span><div class="switch ${st.open_after_generate ? 'on' : ''}" data-action="toggle-open"></div><span class="faint small">Download the workbook afterwards</span></div>`)}
        </div>
        <div>
          <div id="gen-status"></div>
          <div class="panel-title" style="margin-top:22px">Activity</div>
          <div class="code" id="gen-log"></div>
        </div>
      </div>
      ${overviewHtml(req)}
    </div>`;
  renderGenStatus();
  renderGenLog();
};

// School overview shown under the generator: counts from the requirements file.
function overviewHtml(req) {
  return `
      <h2 class="h" style="margin-top:36px">Overview</h2>
      <div class="gallery">
        ${[['grid', 'blue', req?.divisions, 'Class divisions'], ['users', 'green', req?.teachers, 'Teachers'],
           ['layers', 'purple', req?.subjects, 'Subjects'], ['clock', 'orange', req?.periods, 'Periods requested']]
          .map(([i, c, v, l]) => `<div class="stat"><div class="chip-i tint-${c}">${icon(i, 'xl')}</div><div><div class="sv">${v ?? '—'}</div><div class="sl">${l}</div></div></div>`).join('')}
      </div>`;
}

// Which of the three run steps has been reached (0 = checking … 3 = all done).
function genStage() {
  if (S.gen.state === 'done') return 3;
  const texts = S.gen.log.map((l) => l.text);
  if (texts.some((t) => t.startsWith('Schedule found'))) return 2;
  if (texts.some((t) => t.startsWith('Input checked'))) return 1;
  return 0;
}

function renderGenStatus() {
  const host = $('#gen-status');
  const button = $('#gen-btn');
  if (button) {
    button.disabled = S.gen.state === 'running';
    button.innerHTML = S.gen.state === 'running' ? '<span class="spinner"></span>Generating…' : `${icon('zap')}Generate timetable`;
  }
  if (!host) return;
  const g = S.gen;
  const stage = genStage();
  const steps = [
    ['Check requirements', 'Validate rows and balance subject repeats'],
    ['Place every lesson', 'Find slots with no class or teacher clashes'],
    ['Write the workbook', 'Save class and teacher timetables to Excel'],
  ];
  const pipeline = steps.map(([title, desc], i) => {
    let state = 'pending';
    if (g.state === 'idle') state = 'pending';
    else if (i < stage) state = 'done';
    else if (i === stage) state = g.state === 'error' ? 'error' : g.state === 'running' ? 'active' : 'done';
    const mark = state === 'done' ? icon('check', 'sm') : state === 'error' ? icon('x', 'sm') : state === 'active' ? '<span class="spinner"></span>' : i + 1;
    return `<div class="pipe-step ${state}"><span class="dot">${mark}</span><div><div class="pt2">${title}</div><div class="pd2">${desc}</div></div></div>`;
  }).join('');
  const head = {
    idle: ['tint-blue', 'zap', 'Ready when you are', 'Requirements are checked first, so you see exactly what to fix.'],
    running: ['tint-blue', 'timer', 'Generating…', 'This usually takes a few seconds; large schools can take a minute.'],
    done: ['tint-green', 'check', 'Timetable ready', 'Every class period is filled and no teacher is double-booked.'],
    error: ['tint-orange', 'alert', "Couldn't generate the timetable", 'See what needs to change below.'],
  }[g.state];
  let extra = '';
  if (g.state === 'running') extra = '<div class="progress"></div>';
  if (g.state === 'done') {
    const r = g.result;
    extra = `<div class="res-tiles">
        <div class="res-tile"><b>${r.lessons}</b><span>lessons</span></div><div class="res-tile"><b>${r.divisions}</b><span>divisions</span></div>
        <div class="res-tile"><b>${r.teachers}</b><span>teachers</span></div><div class="res-tile"><b>${r.seconds}s</b><span>to build</span></div></div>
      ${r.adjusted_divisions ? `<div class="faint small" style="margin-top:10px">${plural(r.adjusted_divisions, 'division')} had subject repeats adjusted to fill the week.</div>` : ''}
      <div class="run-actions"><button class="btn primary" data-action="go" data-page="timetables">${icon('calendar')}Open timetables</button>
        <button class="btn secondary" data-action="open-path" data-path="${esc(r.output)}">${icon('external')}Download .xlsx</button></div>`;
  }
  if (g.state === 'error') {
    extra = `<pre class="err selectable">${esc(g.result.error)}</pre><div class="faint small" style="margin-top:8px">${esc(g.result.settings || '')}</div>
      <div class="run-actions"><button class="btn secondary" data-action="go" data-page="requirements">${icon('table')}Review requirements</button>
        <button class="btn ghost" data-action="generate">${icon('refresh')}Try again</button></div>`;
  }
  host.innerHTML = `<div class="run-card ${g.state}">
      <div class="run-head"><div class="chip-i ${head[0]}">${icon(head[1], 'lg')}</div>
        <div><div class="rt">${head[2]}</div><div class="rs2">${head[3]}</div></div>
        ${g.state === 'running' || g.state === 'done' ? `<span class="clock" id="gen-clock">${clock(g.elapsed)}</span>` : ''}</div>
      <div class="pipe">${pipeline}</div>${extra}</div>`;
}

function renderGenLog() {
  const host = $('#gen-log');
  if (!host) return;
  host.innerHTML = S.gen.log.length
    ? S.gen.log.map((l) => `<div><span class="t">${esc(l.time)}</span><span class="${esc(l.kind)}">${esc(l.text)}</span></div>`).join('')
    : '<span class="empty">Nothing yet — progress appears here while a timetable is generated.</span>';
  host.scrollTop = host.scrollHeight;
}

async function startGenerate() {
  if (S.gen.state === 'running') return;
  if (!(await guardUnsaved())) return;
  try {
    await call('generate');
  } catch (error) { return fail(error); }
  S.gen = { state: 'running', log: [], logCount: 0, result: null, elapsed: 0, timer: null };
  renderGenStatus(); renderGenLog(); renderNavMeta();
  S.gen.timer = setInterval(pollJob, 400);
}

async function pollJob() {
  let job;
  try { job = await call('job', S.gen.logCount); } catch (error) { return; }
  S.gen.elapsed = job.elapsed;
  if (job.log.length) {
    const before = genStage();
    S.gen.log.push(...job.log); S.gen.logCount = job.logCount; renderGenLog();
    if (genStage() !== before) renderGenStatus();
  }
  const c = $('#gen-clock'); if (c) c.textContent = clock(job.elapsed);
  renderNavMeta();
  if (!job.running && job.result) {
    clearInterval(S.gen.timer);
    S.gen.result = job.result;
    S.gen.state = job.result.ok ? 'done' : 'error';
    renderGenStatus(); renderNavMeta();
    await refreshSummary();
    S.tt.data = null;
    if (S.page === 'home') await RENDER.home();
    if (job.result.ok) toast(`Timetable generated in ${job.result.seconds}s`, '', { action: { label: 'View', run: () => go('timetables') } });
    else toast('Generation needs attention', '⚠️');
  }
}

// ------------------------------------------------------------ Timetables

function currentGrid() {
  const data = S.tt.data;
  if (!data || data.missing) return null;
  const list = data[S.tt.mode];
  return list.find((g) => g.sheet === S.tt.current[S.tt.mode]) || list[0] || null;
}

async function loadTimetables(force = false) {
  if (S.tt.data && !force) return;
  S.tt.data = await call('timetables');
}

RENDER.timetables = async (opts = {}) => {
  const page = $('#page-timetables');
  if (!S.tt.data) {
    page.innerHTML = `<div class="db-head"><div class="db-title"><h1>Timetables</h1></div><div class="db-lead">Loading…</div></div>
      <div style="padding:20px 96px"><div class="skeleton" style="height:420px"></div></div>`;
  }
  try { await loadTimetables(opts.force); } catch (error) { fail(error); S.tt.data = { missing: true, error: error.message }; }
  const data = S.tt.data;
  if (data.missing) {
    page.innerHTML = `<div class="db-head"><div class="db-title"><h1>Timetables</h1></div></div>
      <div class="empty-state"><div><h3>${data.error ? "Couldn't open the workbook" : 'No timetable yet'}</h3>
      <p>${esc(data.error || 'Generate a timetable and every class and teacher schedule will appear here, ready to review and edit.')}</p>
      <button class="btn primary" data-action="go" data-page="generate">${icon('zap')}Generate a timetable</button></div></div>`;
    renderCrumbs();
    return;
  }
  page.innerHTML = `
    <div class="db-head">
      <div class="db-title"><h1>Timetables</h1></div>
      <div class="db-lead">Every class and teacher in the generated workbook.
        <span class="chip" data-action="open-path" data-path="${esc(data.path)}" title="Download workbook">${icon('file', 'sm')}${esc(basename(data.path))}</span>
        <span class="faint small">Updated ${esc(when(data.modified)).replace(/^Today/, 'today').replace(/^Yesterday/, 'yesterday')}</span></div>
    </div>
    <div class="toolbar">
      <div class="tabs">
        <div class="tab ${S.tt.mode === 'classes' ? 'on' : ''}" data-action="tt-mode" data-mode="classes">${icon('grid')}Classes <span class="n">${data.classes.length}</span></div>
        <div class="tab ${S.tt.mode === 'teachers' ? 'on' : ''}" data-action="tt-mode" data-mode="teachers">${icon('users')}Teachers <span class="n">${data.teachers.length}</span></div>
      </div>
      <div class="spacer"></div>
      <label class="search">${icon('search', 'sm')}<input id="tt-q" placeholder="Filter…" value="${esc(S.tt.q)}"></label>
      <div class="sep"></div>
      <button class="btn accent sm" data-action="split">${icon('scissors')}Split periods</button>
      <button class="btn accent sm" data-action="open-path" data-path="${esc(data.path)}">${icon('external')}Download .xlsx</button>
      <button class="btn ghost icon" data-action="tt-reload" title="Reload">${icon('refresh')}</button>
    </div>
    <div class="tt-body"><div class="tt-list" id="tt-list"></div><div class="tt-view" id="tt-view"></div></div>`;
  $('#tt-q').addEventListener('input', (e) => { S.tt.q = e.target.value; renderTtList(); });
  renderTtList();
  renderTtView();
};

function renderTtList() {
  const host = $('#tt-list');
  if (!host) return;
  const data = S.tt.data;
  const term = S.tt.q.trim().toLowerCase();
  const current = currentGrid();
  if (S.tt.mode === 'classes') {
    const byGrade = new Map();
    data.classes.forEach((g) => {
      const [grade] = g.label.split('-');
      if (term && !(`class ${g.label} grade ${grade}`).includes(term)) return;
      if (!byGrade.has(grade)) byGrade.set(grade, []);
      byGrade.get(grade).push(g);
    });
    host.innerHTML = [...byGrade.entries()].map(([grade, list]) => `
      <div class="grade-row"><span class="gn">Grade ${esc(grade)}</span>${list.map((g) => `<span class="class-chip ${current && g.sheet === current.sheet ? 'on' : ''}" data-action="tt-pick" data-sheet="${esc(g.sheet)}" title="Class ${esc(g.label)}">${esc(g.label.split('-')[1])}</span>`).join('')}</div>`).join('')
      || '<div class="list-count">No classes match.</div>';
  } else {
    const list = data.teachers.filter((g) => !term || `${g.label} ${g.subject}`.toLowerCase().includes(term));
    host.innerHTML = list.map((g) => `
      <div class="person ${current && g.sheet === current.sheet ? 'on' : ''}" data-action="tt-pick" data-sheet="${esc(g.sheet)}">
        <div class="pt"><div class="pn">${esc(g.label)}</div><div class="ps">${esc(g.subject || '')}</div></div></div>`).join('')
      + `<div class="list-count">${list.length} of ${data.teachers.length} teachers</div>`;
  }
}

function renderTtView(flash) {
  const host = $('#tt-view');
  if (!host) return;
  const grid = currentGrid();
  if (!grid) { host.innerHTML = ''; return; }
  const isClass = S.tt.mode === 'classes';
  S.tt.current[S.tt.mode] = grid.sheet;
  const lessons = grid.cells.flat().filter((c) => c.trim()).length;
  const slots = grid.days.length * grid.periods.length;
  const counts = new Map();
  grid.cells.flat().forEach((c) => {
    if (!c.trim()) return;
    const parts = c.split('\n');
    const subject = isClass ? parts[0] : (parts[1] || '');
    counts.set(subject, (counts.get(subject) || 0) + 1);
  });
  const classesTaught = isClass ? [] : [...new Set(grid.cells.flat().filter((c) => c.trim()).map((c) => c.split('\n')[0]))].sort();
  const title = isClass ? `Class ${grid.label}` : grid.label;
  const meta = isClass
    ? `${lessons} of ${slots} periods · ${plural(counts.size, 'subject')}<span class="faint">· Click a period to change it</span>`
    : `${esc(grid.subject || 'Teacher')} · ${plural(lessons, 'lesson')} · ${slots - lessons} free<span class="faint">· Teaches</span>${classesTaught.map((c) => `<span class="link" data-action="open-class" data-sheet="${esc(c)}">${esc(c)}</span>`).join(', ')}`;
  const dense = grid.periods.length > 7;
  const columns = `${dense ? 96 : 104}px repeat(${grid.periods.length}, minmax(${dense ? 100 : 118}px, 1fr))`;
  let cells = `<div class="wh"></div>` + grid.periods.map((p, i) => `<div class="wh">${esc(p.replace(/^Period\s*/i, 'P'))}</div>`).join('');
  grid.days.forEach((day, d) => {
    cells += `<div class="wd">${esc(day)}</div>`;
    grid.periods.forEach((_p, pi) => {
      const value = grid.cells[d]?.[pi] || '';
      const parts = value.split('\n');
      const first = (parts[0] || '').trim();
      const second = (parts[1] || '').trim();
      const subject = isClass ? first : second;
      const attrs = isClass ? `data-action="edit-lesson" data-d="${d}" data-p="${pi}"` : (value.trim() ? `data-action="open-class" data-sheet="${esc(first)}"` : '');
      const flashCls = flash && flash[0] === d && flash[1] === pi ? 'flash' : '';
      cells += `<div class="wc ${isClass || value.trim() ? 'editable' : ''} ${flashCls}" ${attrs}>${value.trim()
        ? `<div class="lesson b-${colorFor(subject)}"><span class="ls">${esc(first)}</span><span class="lt">${esc(second)}</span></div>`
        : `<div class="free">${isClass ? '+' : ''}</div>`}</div>`;
    });
  });
  host.innerHTML = `
    <div class="tt-head"><h2>${esc(title)}</h2><span class="grow"></span>
      ${isClass ? '' : `<span class="faint small">Click a lesson to open its class</span>`}</div>
    <div class="tt-meta">${meta}</div>
    <div class="week ${dense ? 'dense' : ''}" style="grid-template-columns:${columns}">${cells}</div>`;
  renderCrumbs();
}

async function openSheet(mode, sheet) {
  S.tt.mode = mode;
  S.tt.current[mode] = sheet;
  S.tt.q = '';
  if (S.page !== 'timetables') await go('timetables');
  else { await RENDER.timetables(); }
  $('.class-chip.on, .person.on')?.scrollIntoView({ block: 'nearest' });
}

async function editLesson(cell) {
  const grid = currentGrid();
  const d = Number(cell.dataset.d), p = Number(cell.dataset.p);
  let options;
  cell.classList.add('saving');
  try { options = await call('lesson_options', grid.sheet, d, p); } catch (error) { cell.classList.remove('saving'); return fail(error); }
  cell.classList.remove('saving');
  const cur = options.current;
  const build = (term) => {
    const items = [];
    const t = term.trim().toLowerCase();
    options.subjects.forEach((s) => {
      const teachers = s.teachers.filter((x) => !t || s.subject.toLowerCase().includes(t) || x.name.toLowerCase().includes(t));
      if (!teachers.length) return;
      items.push({ header: s.subject });
      teachers.forEach((x) => {
        const isCurrent = x.name === cur.teacher && s.subject === cur.subject;
        items.push({ label: x.name, value: { subject: s.subject, teacher: x.name }, disabled: !!x.busy,
          className: `${x.busy ? 'blocked' : ''} ${isCurrent ? 'current' : ''}`, hint: x.busy ? `Busy · ${x.busy}` : '', checked: isCurrent });
      });
    });
    if (cur.subject) items.push({ sep: true }, { label: 'Clear this period', icon: 'trash', danger: true, value: { subject: '', teacher: '' } });
    return items;
  };
  const header = `<div class="ph"><div class="t">${esc(options.day)} · ${esc(options.period)}</div>
    <div class="s">Class ${esc(options.class)}${cur.subject ? ` · now ${esc(cur.subject)} with ${esc(cur.teacher)}` : ' · empty period'}</div></div>
    <div class="pf"><input id="pick-q" placeholder="Find a subject or teacher…" autocomplete="off"></div>`;
  const onPick = async (value) => {
    if (value.subject === cur.subject && value.teacher === cur.teacher) return;
    cell.classList.add('saving');
    cell.innerHTML = `<div class="free"><span class="spinner"></span></div>`;
    try {
      await call('set_lesson', grid.sheet, d, p, value.subject, value.teacher);
      await loadTimetables(true);
      renderTtView([d, p]);
      refreshSummary();
      toast(value.subject ? `${options.day} ${options.period} → ${value.subject} · ${value.teacher}` : `${options.day} ${options.period} cleared`, '');
    } catch (error) {
      renderTtView();
      infoDialog({ title: "Can't make that change", text: error.message, tone: 'red' });
    }
  };
  const menu = openMenu(cell, build(''), onPick, { className: 'picker', header, width: 340 });
  const input = $('#pick-q', menu);
  input.focus();
  input.addEventListener('input', () => {
    $('.ml', menu).innerHTML = menuItems(openMenuState.items = build(input.value));
  });
}

async function openSplit() {
  const grades = S.boot.classes;
  const f = {
    grade: grades[grades.length - 1], subject: '', teacher: '', division: 'All divisions', periods: 1, options: [],
    mode: 'existing',     // 'existing' = a teacher who already teaches the subject, 'new' = type a name
    candidates: null,     // null = not loaded yet, 'loading', or [{name, lessons, free}]
    existing: '', newTeacher: '',
  };
  const m = modal(`<div class="modal-body"><h3>Split subject periods</h3>
      <p>Move some of a teacher's lessons in one grade to another teacher of the same subject. Teacher sheets and requirement counts update together.</p>
      <div id="split-form"></div></div>
      <div class="modal-foot"><button class="btn secondary" data-x="cancel">Cancel</button><button class="btn primary" data-x="go" id="split-go">Transfer periods</button></div>`, { wide: true });

  const target = () => (f.mode === 'existing' ? f.existing : f.newTeacher.trim());
  const chosenCandidate = () => (Array.isArray(f.candidates) ? f.candidates.find((c) => c.name === f.existing) : null);

  const render = () => {
    const subject = f.options.find((s) => s.subject === f.subject);
    const teacher = subject?.teachers.find((t) => t.name === f.teacher);
    const divisions = teacher ? teacher.divisions : {};
    let maxPeriods = f.division === 'All divisions' ? Math.min(...Object.values(divisions), 99) : (divisions[f.division] || 1);
    const cand = chosenCandidate();
    if (f.mode === 'existing' && cand) maxPeriods = Math.min(maxPeriods, cand.free);
    f.periods = Math.max(1, Math.min(f.periods, maxPeriods || 1));

    let receiver = '';
    if (teacher) {
      const toggle = `<div class="segmented" style="margin-bottom:10px">
          <button class="${f.mode === 'existing' ? 'on' : ''}" data-x="mode" data-v="existing">${icon('users', 'sm')}Teacher already teaching ${esc(f.subject)}</button>
          <button class="${f.mode === 'new' ? 'on' : ''}" data-x="mode" data-v="new">${icon('plus', 'sm')}New teacher</button></div>`;
      let field;
      if (f.mode === 'new') {
        field = `<input class="input" id="split-new" placeholder="Name of the new ${esc(f.subject)} teacher" value="${esc(f.newTeacher)}">`;
      } else if (f.candidates === 'loading' || f.candidates === null) {
        field = `<div class="select bordered" style="width:100%"><span class="spinner"></span><span class="placeholder">Finding ${esc(f.subject)} teachers…</span></div>`;
      } else if (!f.candidates.length) {
        field = `<div class="callout yellow"><div class="cb"><div class="ct">No other ${esc(f.subject)} teachers</div>Nobody else teaches ${esc(f.subject)} yet. Choose <b>New teacher</b> to add one.</div></div>`;
      } else {
        field = `<div class="select bordered" data-x="existing" style="width:100%">${cand ? `${avatar(cand.name)}${esc(cand.name)}<span class="faint small" style="margin-left:6px">free for ${cand.free} · ${cand.lessons} lessons now</span>`
          : '<span class="placeholder">Choose a teacher</span>'}<span style="flex:1"></span>${icon('down', 'sm')}</div>`;
      }
      receiver = `<div class="form-row"><label>Give the periods to</label>${toggle}${field}</div>`;
    }

    $('#split-form', m.el).innerHTML = `
      <div style="display:flex;gap:12px">
        <div class="form-row" style="flex:1"><label>Grade</label><div class="select bordered" data-x="grade" style="width:100%">Grade ${esc(f.grade)}<span style="flex:1"></span>${icon('down', 'sm')}</div></div>
        <div class="form-row" style="flex:2"><label>Subject</label><div class="select bordered" data-x="subject" style="width:100%">${f.subject ? esc(f.subject) : '<span class="placeholder">Choose a subject</span>'}<span style="flex:1"></span>${icon('down', 'sm')}</div></div>
      </div>
      <div class="form-row"><label>Current teacher</label><div class="select bordered" data-x="teacher" style="width:100%">${f.teacher ? esc(f.teacher) : '<span class="placeholder">Choose a teacher</span>'}<span style="flex:1"></span>${icon('down', 'sm')}</div></div>
      ${teacher ? `<div class="form-row"><label>Divisions</label><div class="division-picks">
        ${['All divisions', ...Object.keys(divisions)].map((dv) => `<span class="tag click ${f.division === dv ? 'c-blue' : 'off'}" data-x="div" data-v="${esc(dv)}">${dv === 'All divisions' ? 'All' : `${esc(f.grade)}-${esc(dv)}`}${dv === 'All divisions' ? '' : ` <span style="opacity:.6">${divisions[dv]}</span>`}</span>`).join('')}</div></div>` : ''}
      ${receiver}
      ${teacher ? `<div class="form-row"><label>Periods per division</label><div class="stepper"><button data-x="step" data-d="-1">${icon('minus', 'sm')}</button><span class="val">${f.periods}</span><button data-x="step" data-d="1">${icon('plus', 'sm')}</button></div></div>` : ''}
      <div class="form-row help">${teacher
        ? `Up to ${maxPeriods} per division${f.mode === 'existing' && cand ? ` (${esc(cand.name)} is free for ${cand.free})` : ''}. Lessons that would clash are refused, and nothing is saved unless the whole split works.`
        : 'Pick a subject and its current teacher to see the divisions they teach.'}</div>`;
    const input = $('#split-new', m.el);
    if (input) input.addEventListener('input', (e) => { f.newTeacher = e.target.value; $('#split-go', m.el).disabled = !target(); });
    $('#split-go', m.el).disabled = !(f.subject && f.teacher && target() && !(f.mode === 'existing' && cand && cand.free < 1));
  };

  // Other teachers of this subject, with how many of the periods each is free for.
  const loadCandidates = async () => {
    f.candidates = 'loading'; f.existing = '';
    render();
    try {
      f.candidates = (await call('split_candidates', f.grade, f.subject, f.teacher, f.division)).teachers;
      const best = f.candidates.find((c) => c.free > 0);
      if (best) f.existing = best.name;
    } catch (error) { f.candidates = []; fail(error); }
    render();
  };

  const loadOptions = async () => {
    $('#split-form', m.el).innerHTML = '<div class="skeleton" style="height:180px;margin-top:14px"></div>';
    try { f.options = (await call('split_options', f.grade)).subjects; } catch (error) { f.options = []; fail(error); }
    f.subject = ''; f.teacher = ''; f.division = 'All divisions'; f.candidates = null; f.existing = '';
    render();
  };

  m.el.addEventListener('click', async (e) => {
    const x = e.target.closest('[data-x]');
    if (!x) return;
    const kind = x.dataset.x;
    if (kind === 'cancel') m.close();
    if (kind === 'grade') openMenu(x, grades.map((g) => ({ label: `Grade ${g}`, value: g, checked: g === f.grade })), (v) => { f.grade = v; loadOptions(); });
    if (kind === 'subject') openMenu(x, f.options.map((s) => ({ label: s.subject, value: s.subject, tagColor: colorFor(s.subject), checked: s.subject === f.subject })), (v) => { f.subject = v; f.teacher = ''; f.division = 'All divisions'; f.candidates = null; render(); });
    if (kind === 'teacher') {
      const s = f.options.find((o) => o.subject === f.subject);
      if (!s) return toast('Choose a subject first');
      openMenu(x, s.teachers.map((t) => ({ label: t.name, value: t.name, hint: Object.keys(t.divisions).map((d) => `${f.grade}-${d}`).join(', '), checked: t.name === f.teacher })), (v) => { f.teacher = v; f.division = 'All divisions'; loadCandidates(); });
    }
    if (kind === 'div') { f.division = x.dataset.v; loadCandidates(); }
    if (kind === 'mode') { f.mode = x.dataset.v; render(); if (f.mode === 'new') $('#split-new', m.el)?.focus(); }
    if (kind === 'existing' && Array.isArray(f.candidates)) {
      openMenu(x, f.candidates.map((c) => ({
        label: c.name, value: c.name, disabled: c.free < 1, checked: c.name === f.existing,
        hint: c.free < 1 ? 'busy at those times' : `free for ${c.free} · ${c.lessons} lessons now`,
      })), (v) => { f.existing = v; render(); }, { width: 420 });
    }
    if (kind === 'step') { f.periods += Number(x.dataset.d); render(); }
    if (kind === 'go') {
      const to = target();
      if (!to) return toast(f.mode === 'new' ? "Enter the new teacher's name" : 'Choose a teacher');
      x.disabled = true; x.innerHTML = '<span class="spinner"></span>Transferring…';
      try {
        const r = await call('split', f.grade, f.subject, f.teacher, to, f.periods, f.division);
        m.close();
        await loadTimetables(true); renderTtList(); renderTtView();
        toast(`Moved ${r.periods_transferred ?? ''} ${f.subject} periods from ${f.teacher} to ${r.new_teacher || to}`);
      } catch (error) {
        x.disabled = false; x.innerHTML = 'Transfer periods';
        infoDialog({ title: "Couldn't split the periods", text: error.message, tone: 'red' });
      }
    }
  });
  loadOptions();
}

// ------------------------------------------------------------ Settings

RENDER.settings = async () => {
  const st = S.settings, b = S.boot;
  const modeText = { source: 'Running from source — data is kept in the project folder.', portable: 'Portable — data is kept beside SchoolFlow.exe, so the folder can move between computers.', installed: 'Installed — your school data is kept in Documents\\SchoolFlow.', web: 'Web — your school data is kept on the SchoolFlow server, in the folder below. Use Download to take a copy.', online: 'Online — your school data is saved in this browser on this device only. The server works on it for a moment and keeps nothing. Download a copy to back it up or move it to another computer.' }[b.mode] || '';
  const row = (name, desc, control = '') => `<div class="setting"><div class="st"><div class="sn">${name}</div><div class="sd selectable">${desc}</div></div>${control}</div>`;
  $('#page-settings').innerHTML = `
    <div class="doc">
      <h1 class="page-title">Settings</h1>
      <p class="page-lead">Appearance, where your school data lives, and how long the scheduler may search.</p>
      <div class="cols">
      <div>
      <div class="section-title" style="margin-top:0">Appearance</div>
      ${row('Theme', 'Follow your device, or always use light or dark.', `<div class="segmented">${[['system', 'monitor', 'System'], ['light', 'sun', 'Light'], ['dark', 'moon', 'Dark']].map(([v, i, l]) => `<button class="${st.theme === v ? 'on' : ''}" data-action="set-theme" data-v="${v}">${icon(i, 'sm')}${l}</button>`).join('')}</div>`)}
      <div class="section-title">Data</div>
      ${row('Storage', esc(modeText))}
      ${row('Requirements file', pathHtml(st.input_path), `<button class="btn secondary sm" data-action="choose-req">Upload CSV…</button>`)}
      ${row('Requirements folder', pathHtml(b.requirementsFolder), `<button class="btn secondary sm" data-action="open-path" data-path="${esc(b.requirementsFolder)}">${icon('folder')}Download</button>`)}
      ${row('Exports folder', pathHtml(b.exportsFolder), `<button class="btn secondary sm" data-action="open-path" data-path="${esc(b.exportsFolder)}">${icon('folder')}Download</button>`)}
      <div class="section-title">Scheduler</div>
      ${row('Search time limit', 'How long the solver may look for a timetable. Large schools or tight weeks may need more.', `<div class="select bordered" data-action="pick-limit">${st.solver_time_limit} seconds${icon('down', 'sm')}</div>`)}
      <div class="section-title">Sample data</div>
      ${row('Restore sample requirements', 'Replace the working requirements file with the sample school that ships with SchoolFlow.', `<button class="btn danger sm" data-action="restore-sample">Restore…</button>`)}
      </div>
      <div>
      <div class="section-title" style="margin-top:0">Keyboard shortcuts</div>
      ${[['Switch pages', ['Ctrl', '1 – 4']], ['Save requirements', ['Ctrl', 'S']], ['New requirement row', ['Ctrl', 'N']], ['Generate timetable', ['F5']], ['Close panels and dialogs', ['Esc']]]
        .map(([l, k]) => `<div class="shortcut"><span>${l}</span><span class="keys">${k.map((x) => `<span class="kbd">${x}</span>`).join('')}</span></div>`).join('')}
      <div class="section-title">About</div>
      ${row('Version', `SchoolFlow ${esc(b.version)}`)}
      ${row('Engines', esc(b.engines))}
      ${row('Preferences file', pathHtml(b.settingsFile))}
      </div>
      </div>
    </div>`;
};

function pickLimit(anchor) {
  openMenu(anchor, [30, 60, 120, 180, 240].map((v) => ({ label: `${v} seconds`, value: v, checked: v === S.settings.solver_time_limit })), async (v) => {
    await saveSettings({ solver_time_limit: v });
    RENDER[S.page]();
    toast(`Time limit set to ${v} seconds`);
  });
}

// ------------------------------------------------------------ Actions

function cycleTheme() {
  const next = { system: 'light', light: 'dark', dark: 'system' }[S.settings.theme] || 'system';
  saveSettings({ theme: next }).then(() => { applyTheme(); renderThemeButton(); if (S.page === 'settings') RENDER.settings(); toast(`Theme: ${next}`); });
}
const openPath = (path) => call('open_path', path).catch(fail);

const ACTIONS = {
  go: (el) => go(el.dataset.page),
  'cycle-theme': cycleTheme,
  'set-theme': (el) => saveSettings({ theme: el.dataset.v }).then(() => { applyTheme(); renderThemeButton(); RENDER.settings(); }),
  'open-path': (el) => openPath(el.dataset.path),
  'open-folder': () => { const p = S.settings.output_path; openPath(p.replace(/[\\/][^\\/]*$/, '')); },
  'choose-req': async () => {
    if (S.req.dirty && !(await confirmDialog({ title: 'Discard unsaved changes?', text: 'Switching files discards your unsaved requirement changes.', confirm: 'Discard', danger: true }))) return;
    try {
      const r = await call('choose_requirements');
      if (!r.changed) return;
      S.settings.input_path = r.path; S.req.loaded = false; setDirty(false);
      await refreshSummary(); RENDER[S.page](); toast(`Using ${basename(r.path)}`);
    } catch (error) { fail(error); }
  },
  'choose-output': async () => {
    try { const r = await call('choose_output'); if (r.changed) { S.settings.output_path = r.path; window.__sfOutput = r.path; S.tt.data = null; await refreshSummary(); RENDER.home(); } } catch (error) { fail(error); }
  },
  'restore-sample': async () => {
    if (!(await confirmDialog({ title: 'Restore the sample school?', text: `This replaces ${S.boot.requirementsFolder}\\requirements.csv with the sample data. Your current requirements in that file will be lost.`, confirm: 'Restore sample', danger: true }))) return;
    try { const r = await call('restore_sample'); S.settings.input_path = r.path; S.req.loaded = false; setDirty(false); await refreshSummary(); RENDER.settings(); toast('Sample requirements restored'); } catch (error) { fail(error); }
  },

  // Requirements
  'req-view': (el) => { S.req.view = el.dataset.view; $$('[data-action="req-view"]').forEach((t) => t.classList.toggle('on', t === el)); renderReqTable(); },
  'grade-pill': (el) => { S.req.grade = el.dataset.v; renderFilters(); renderReqTable(); $('#req-scroll').scrollTop = 0; },
  'filter-div': (el) => openMenu(el, [{ label: 'All divisions', value: 'All', checked: S.req.division === 'All' }, { sep: true }, ...S.boot.divisions.map((d) => ({ label: `Division ${d}`, value: d, checked: S.req.division === d }))], (v) => { S.req.division = v; renderFilters(); renderReqTable(); }),
  'open-row': (el) => openPeek(Number(el.dataset.i)),
  'new-row': () => openPeek(null, { class: S.req.grade !== 'All' ? S.req.grade : '', division: S.req.division !== 'All' ? S.req.division : '' }),
  'close-peek': closePeek,
  'peek-step': (el) => {
    const list = visibleRows();
    const at = list.indexOf(S.req.peek.index);
    const next = list[at + Number(el.dataset.step)];
    if (next !== undefined) { openPeek(next); $(`tr[data-i="${next}"]`)?.scrollIntoView({ block: 'nearest' }); }
  },
  'dup-row': () => { const v = { ...S.req.peek.values }; openPeek(null, v); },
  'delete-row': deleteRow,
  'apply-peek': applyPeek,
  'pk-class': (el) => openMenu(el, S.boot.classes.map((g) => ({ label: `Class ${g}`, value: g, checked: S.req.peek.values.class === g })), (v) => { S.req.peek.values.class = v; el.innerHTML = `Class ${esc(v)}${icon('down', 'sm')}`; }),
  'pk-div': (el) => openMenu(el, S.boot.divisions.map((d) => ({ label: d, value: d, checked: S.req.peek.values.division === d })), (v) => { S.req.peek.values.division = v; el.innerHTML = `${esc(v)}${icon('down', 'sm')}`; }),
  'pk-step': (el) => { const v = S.req.peek.values; v.periods_per_week = String(Math.max(1, Math.min(99, (Number(v.periods_per_week) || 1) + Number(el.dataset.d)))); $('#pk-periods').textContent = v.periods_per_week; },
  'pk-pick-subject': (el) => { S.req.peek.values.subject = el.dataset.v; $('#pk-subject').value = el.dataset.v; renderPeekHints(); renderCrumbs(); },
  'save-req': saveRequirements,
  'check-req': checkRequirements,

  // Generate
  'toggle-day': (el) => {
    if (S.gen.state === 'running') return;
    const day = el.dataset.day;
    let days = S.settings.days.includes(day) ? S.settings.days.filter((d) => d !== day) : [...S.settings.days, day];
    if (days.length < 2) return toast('Keep at least two school days');
    days = DAYS.filter((d) => days.includes(d));
    const min = minPeriods(days.length);
    const periods = Math.max(min, S.settings.periods_per_day);
    saveSettings({ days, periods_per_day: periods }).then(() => RENDER.home());
  },
  'periods-step': (el) => {
    if (S.gen.state === 'running') return;
    const min = minPeriods(S.settings.days.length);
    const next = Math.max(min, Math.min(12, S.settings.periods_per_day + Number(el.dataset.d)));
    if (next !== S.settings.periods_per_day) saveSettings({ periods_per_day: next }).then(() => RENDER.home());
  },
  'pick-limit': (el) => pickLimit(el),
  'toggle-open': (el) => { el.classList.toggle('on'); saveSettings({ open_after_generate: el.classList.contains('on') }); },
  generate: () => startGenerate(),

  // Timetables
  'tt-mode': (el) => { S.tt.mode = el.dataset.mode; $$('[data-action="tt-mode"]').forEach((t) => t.classList.toggle('on', t === el)); renderTtList(); renderTtView(); },
  'tt-pick': (el) => { S.tt.current[S.tt.mode] = el.dataset.sheet; renderTtList(); renderTtView(); $('#tt-view').scrollTop = 0; },
  'tt-reload': () => { S.tt.data = null; RENDER.timetables({ force: true }).then(() => toast('Timetables reloaded')); },
  'open-class': (el) => openSheet('classes', el.dataset.sheet),
  'edit-lesson': (el) => editLesson(el),
  split: () => openSplit(),
};

document.addEventListener('click', (e) => {
  const el = e.target.closest('[data-action]');
  if (!el || el.disabled) return;
  const handler = ACTIONS[el.dataset.action];
  if (handler) { e.preventDefault(); handler(el, e); }
});

document.addEventListener('keydown', (e) => {
  const key = e.key.toLowerCase();
  if (e.key === 'F5' || (e.ctrlKey && key === 'r')) { e.preventDefault(); if (e.key === 'F5') { go('generate'); startGenerate(); } return; }
  if (e.key === 'Escape') {
    if (openMenuState) { closeMenu(); return; }
    if (modalStack.length) { modalStack[modalStack.length - 1](); return; }
    if (S.page === 'requirements' && S.req.peek) { closePeek(); return; }
  }
  if (openMenuState && ['ArrowDown', 'ArrowUp', 'Enter'].includes(e.key)) {
    const items = $$('.mi:not(.disabled)', openMenuState.el);
    if (!items.length) return;
    let i = items.findIndex((x) => x.classList.contains('focus'));
    if (e.key === 'Enter') { if (i >= 0) { items[i].click(); e.preventDefault(); } return; }
    i = e.key === 'ArrowDown' ? Math.min(items.length - 1, i + 1) : Math.max(0, i - 1);
    items.forEach((x, j) => x.classList.toggle('focus', j === i));
    items[i].scrollIntoView({ block: 'nearest' });
    e.preventDefault();
    return;
  }
  if (modalStack.length) return;
  if (e.ctrlKey && /^[1-4]$/.test(e.key)) { e.preventDefault(); go(ORDER[Number(e.key) - 1]); return; }
  if (e.ctrlKey && key === 's') { e.preventDefault(); if (S.page === 'requirements') saveRequirements(); return; }
  if (e.ctrlKey && key === 'n') { e.preventDefault(); go('requirements').then(() => ACTIONS['new-row']()); return; }
  if (e.ctrlKey && e.key === 'Enter' && S.req.peek && S.page === 'requirements') { e.preventDefault(); applyPeek(); }
});
document.addEventListener('contextmenu', (e) => { if (!e.target.closest('input, textarea, .selectable')) e.preventDefault(); });

// ------------------------------------------------------------ Start

async function start() {
  API = window.pywebview.api;
  S.boot = await call('bootstrap');
  S.settings = S.boot.settings;
  applyTheme();
  renderShell();
  await refreshSummary();
  await go('home');
  const boot = $('#boot');
  boot.style.opacity = '0';
  setTimeout(() => boot.remove(), 220);
  // Warm the timetable cache so Ctrl+K can find classes and teachers immediately.
  loadTimetables().catch(() => {});
}

if (window.pywebview && window.pywebview.api) start();
else window.addEventListener('pywebviewready', start);
