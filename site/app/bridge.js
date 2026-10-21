/* SchoolFlow web bridge. The desktop app talks to Python through
   window.pywebview.api; here the same calls go to the SchoolFlow server.
   The school's files (requirements CSV and timetable workbook) live in this
   browser and travel with each call; the server keeps nothing. */
'use strict';

(function () {
  const KEY = 'schoolflow.v1';
  const DEFAULT_SETTINGS = { days: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'], periods_per_day: 7, solver_time_limit: 180, open_after_generate: false, theme: 'system' };
  const REQUIRED = ['class', 'division', 'subject', 'teacher', 'periods_per_week'];

  // ---- Browser storage (wrapped: private windows may refuse it)
  let store = null;
  try { store = JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { store = null; }
  store = Object.assign({ settings: {}, requirements: null, reqName: 'requirements.csv', workbook: null, wbName: 'school_timetable.xlsx', wbModified: 0 }, store || {});
  store.settings = Object.assign({}, DEFAULT_SETTINGS, store.settings);
  let storageWarned = false;
  function persist() {
    try { localStorage.setItem(KEY, JSON.stringify(store)); }
    catch (e) { if (!storageWarned) { storageWarned = true; console.warn('SchoolFlow could not save to browser storage; data will be lost when the tab closes.'); } }
  }

  const reqPath = () => 'Requirements/' + store.reqName;
  const wbPath = () => 'Exports/' + store.wbName;
  const publicSettings = () => Object.assign({}, store.settings, { input_path: reqPath(), output_path: wbPath() });

  // ---- Server calls
  async function post(body) {
    let response;
    try {
      response = await fetch('/api/call', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    } catch (error) {
      return { ok: false, error: 'Cannot reach the SchoolFlow server. Check your internet connection and try again.' };
    }
    if (response.status === 504) return { ok: false, error: 'The server ran out of time. Lower the search time limit in Settings, or simplify the requirements, and try again.' };
    try { return await response.json(); } catch (error) { return { ok: false, error: 'The server sent an unexpected reply (' + response.status + ').' }; }
  }

  async function remote(method, args) {
    const result = await post({
      method, args: args || [],
      files: { requirements: store.requirements, workbook: store.workbook, workbookModified: store.wbModified },
      settings: Object.assign({}, store.settings, { requirements_name: store.reqName, workbook_name: store.wbName }),
    });
    const files = result && result._files;
    if (files) {
      if (typeof files.requirements === 'string') store.requirements = files.requirements;
      if (files.workbook) { store.workbook = files.workbook; store.wbModified = files.workbookModified || Date.now() / 1000; }
      if (files.requirements !== undefined || files.workbook) persist();
      delete result._files;
    }
    return result;
  }

  // ---- Downloads
  function save(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }
  function downloadWorkbook() {
    if (!store.workbook) return { ok: false, error: 'Generate a timetable first; there is no workbook to download yet.' };
    const bytes = Uint8Array.from(atob(store.workbook), (c) => c.charCodeAt(0));
    save(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), store.wbName);
    return { ok: true };
  }
  function downloadRequirements() {
    if (store.requirements == null) return { ok: false, error: 'There is no requirements file yet.' };
    save(new Blob(['﻿' + store.requirements], { type: 'text/csv' }), store.reqName);
    return { ok: true };
  }

  function pickFile(accept) {
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file'; input.accept = accept; input.style.display = 'none';
      input.addEventListener('change', () => { resolve(input.files[0] || null); input.remove(); });
      window.addEventListener('focus', () => setTimeout(() => { if (!input.files.length) { resolve(null); input.remove(); } }, 800), { once: true });
      document.body.appendChild(input);
      input.click();
    });
  }

  // ---- Generation runs as one long request; the app polls job() meanwhile.
  const now = () => new Date().toTimeString().slice(0, 8);
  let job = { running: false, log: [], result: null, started: 0, finished: 0 };

  let dirty = false;
  window.addEventListener('beforeunload', (e) => { if (dirty || job.running) { e.preventDefault(); e.returnValue = ''; } });

  const local = {
    async bootstrap() {
      const info = await post({ method: 'info' });
      if (!info || info.ok === false) return info;
      if (store.requirements == null) {
        const s = await post({ method: 'sample' });
        if (s && s.ok) { store.requirements = s.text; persist(); }
      }
      return {
        ok: true, version: info.version, mode: 'online', engines: info.engines,
        classes: info.classes, divisions: info.divisions, minPeriods: info.minPeriods,
        dataFolder: 'This browser', requirementsFolder: 'Requirements', exportsFolder: 'Exports',
        settingsFile: 'This browser (local storage)', settings: publicSettings(),
      };
    },
    async save_settings(patch) {
      for (const [key, value] of Object.entries(patch || {})) {
        if (key in DEFAULT_SETTINGS && typeof value === typeof DEFAULT_SETTINGS[key]) store.settings[key] = value;
      }
      persist();
      return { ok: true, settings: publicSettings() };
    },
    async choose_requirements() {
      const file = await pickFile('.csv,text/csv');
      if (!file) return { ok: true, changed: false };
      let text = await file.text();
      text = text.replace(/^﻿/, '');
      const header = (text.split(/\r?\n/)[0] || '').split(',').map((h) => h.trim().toLowerCase());
      const missing = REQUIRED.filter((c) => !header.includes(c));
      if (missing.length) return { ok: false, error: 'The CSV is missing these columns: ' + missing.join(', ') };
      store.requirements = text;
      store.reqName = file.name.replace(/[^\w .()-]/g, '_') || 'requirements.csv';
      persist();
      return { ok: true, changed: true, path: reqPath() };
    },
    async choose_output() {
      const name = window.prompt('Workbook file name:', store.wbName);
      if (!name || !name.trim()) return { ok: true, changed: false };
      store.wbName = name.trim().replace(/[^\w .()-]/g, '_').replace(/(\.xlsx)?$/i, '.xlsx');
      persist();
      return { ok: true, changed: true, path: wbPath() };
    },
    async restore_sample() {
      const s = await post({ method: 'sample' });
      if (!s || !s.ok) return s;
      store.requirements = s.text; store.reqName = 'requirements.csv'; persist();
      return { ok: true, path: reqPath() };
    },
    async open_path(path) {
      const p = String(path || '');
      if (p === 'Requirements' || /\.csv$/i.test(p)) return downloadRequirements();
      return downloadWorkbook();
    },
    async set_dirty(value) { dirty = !!value; return true; },
    async generate() {
      if (job.running) return { ok: false, error: 'A timetable is already being generated.' };
      job = { running: true, log: [{ time: now(), text: 'Sent to the SchoolFlow server. Solving…', kind: 'accent' }], result: null, started: Date.now(), finished: 0 };
      remote('generate').then((r) => {
        if (r && Array.isArray(r.log)) job.log.push(...r.log);
        let result = r && r.result;
        if (!result) {
          result = { ok: false, error: (r && r.error) || 'Something went wrong.' };
          job.log.push({ time: now(), text: result.error, kind: 'error' });
        }
        job.result = result; job.finished = Date.now(); job.running = false;
        if (result.ok && store.settings.open_after_generate) downloadWorkbook();
      });
      return { ok: true, started: true };
    },
    async job(since) {
      const elapsed = ((job.finished || Date.now()) - job.started) / 1000;
      return { ok: true, running: job.running, elapsed: job.started ? Math.round(elapsed * 10) / 10 : 0, log: job.log.slice(Number(since) || 0), logCount: job.log.length, result: job.result };
    },
  };

  const api = new Proxy({}, {
    get(_target, name) {
      if (typeof name !== 'string') return undefined;
      if (local[name]) return local[name];
      return (...args) => remote(name, args);
    },
  });

  window.pywebview = { api };
})();
