// app.js — Sheets Insight screens: Upload, Overview, and for each tab Summary / Problems / Rows (Cards or Table).
// Data is only ever READ. Nothing is written back to Google Drive or the original Excel files.
(function () {
  'use strict';
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  // Icons are cosmetic: if the icon file failed to load, the page must still render.
  const I = (name, cls) => { try { return typeof window.icon === 'function' ? window.icon(name, cls) : ''; } catch (e) { return ''; } };
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const REFRESH_MS = 5 * 60 * 1000;
  const PAGE = 100;        // table rows per page
  const CARD_STEP = 48;    // cards shown per "Show more"

  const pref = (k, d) => { try { return localStorage.getItem(k) ?? d; } catch (e) { return d; } };
  const setPref = (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* ignore */ } };

  const state = {
    sources: [],        // { id, name, kind, file?, status, error }
    files: [],          // { id, name, addedAt, sheets: [model] }
    screen: 'start',    // 'start' = upload screen (shown first), 'app' = dashboard
    sel: null,          // { fileId, sheet } — null = Overview
    open: new Set(),    // files expanded in the sidebar
    view: 'dashboard',  // dashboard | health | data
    rowsView: pref('sd-rows-view', 'cards') === 'table' ? 'table' : 'cards',
    filters: {}, search: '', issue: '', measure: undefined, sort: null, page: 0, cardLimit: CARD_STEP,
    fromIssue: null,    // { title } when rows were opened from a problem
    locate: null,       // row number to scroll to in the table
    lastSync: null,
    canSave: true,      // false on read-only hosting (e.g. Vercel): uploads stay in this browser
    locked: 0,          // shared Google sheets waiting behind the website password
    unlocked: false,
    misconfigured: false,
  };
  async function loadInfo() {
    try {
      const r = await fetch('/api/info');
      if (!r.ok) return;
      const j = await r.json();
      state.canSave = j.canSave !== false;
      state.locked = Number(j.locked) || 0;
      state.unlocked = !!j.unlocked;
      state.misconfigured = !!j.misconfigured;
    } catch (e) { /* keep defaults */ }
  }
  async function unlock(password) {
    const btn = $('#unlockBtn');
    if (btn) { btn.disabled = true; btn.textContent = 'Checking…'; }
    try {
      const r = await fetch('/api/login', { method: 'POST', headers: { 'X-App': 'sheets-dashboard', 'Content-Type': 'application/json' }, body: JSON.stringify({ password }) });
      const j = await r.json().catch(() => ({}));
      if (!r.ok) { toast(j.error || 'Could not unlock.'); return; }
      await loadInfo();
      await syncAll();
      toast('Unlocked — shared Google sheets are loading');
    } catch (e) { toast('Could not unlock.'); }
    finally { const b = $('#unlockBtn'); if (b) { b.disabled = false; b.textContent = 'Unlock'; } }
  }
  async function lockAgain() {
    await fetch('/api/logout', { method: 'POST', headers: { 'X-App': 'sheets-dashboard' } }).catch(() => {});
    state.files = state.files.filter((f) => !state.sources.some((s) => s.id === f.id && s.kind === 'drive'));
    for (const s of state.sources) if (s.kind === 'drive') store.del(s.id);
    await loadInfo();
    await syncAll();
    toast('Locked — shared sheets are hidden on this browser');
  }
  function lockCardHtml() {
    if (state.misconfigured) return `<div class="notice">${I('triangle-alert', 'sm')}<span>Shared Google sheets are set up on this website but no password is set, so they are hidden. Add <b>APP_PASSWORD</b> in the website settings.</span></div>`;
    if (!state.locked) return '';
    return `<form class="card lock-card" id="unlockForm">
      <span class="card-icon">${I('cloud')}</span>
      <div class="grow"><div class="fname">${plural(state.locked, 'shared Google sheet')} protected by a password</div>
        <div class="fmeta">Enter the password once — this browser remembers it for 30 days.</div></div>
      <input type="password" id="unlockPw" class="input" placeholder="Password" autocomplete="current-password" aria-label="Password" required>
      <button class="btn primary" id="unlockBtn" type="submit">Unlock</button>
    </form>`;
  }
  let charts = [];
  let syncing = false;

  // ---------- formatting ----------
  const nf = new Intl.NumberFormat('en-IN', { maximumFractionDigits: 2 });
  function short(n) { // 1,23,45,678 -> 1.23 Cr ; 2,80,000 -> 2.8 L
    const a = Math.abs(n);
    if (a >= 1e7) return nf.format(+(n / 1e7).toFixed(2)) + ' Cr';
    if (a >= 1e5) return nf.format(+(n / 1e5).toFixed(2)) + ' L';
    return nf.format(Math.round(n * 100) / 100);
  }
  const isMoney = (col) => col && col.role === 'amount' && !/qty|quantity|stock|ltr|litre|kg|days|points/i.test(col.name);
  const money = (col, n) => (isMoney(col) ? '₹' + short(n) : short(n));
  const fmtDate = (d) => `${String(d.getDate()).padStart(2, '0')} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
  // Numbers get commas only in number columns; codes like PO / invoice numbers show as typed.
  const fmtCell = (v, col) => (v === null || v === undefined ? '' : v instanceof Date ? fmtDate(v)
    : typeof v === 'number' ? (!col || col.type === 'number' ? (isMoney(col) ? '₹' + nf.format(v) : nf.format(v)) : String(v)) : String(v));
  const totalLabel = (n) => (/^(total|net|tot)\b/i.test(n) ? n : `Total ${n}`);
  const plural = (n, one, many) => `${nf.format(n)} ${n === 1 ? one : many || one + 's'}`;
  const shortTitle = (t) => t.replace(/ \(column [A-Z]+\)/, '');
  function ago(t) {
    const s = Math.round((Date.now() - t) / 1000);
    if (s < 60) return 'just now';
    if (s < 3600) return `${Math.round(s / 60)} min ago`;
    return new Date(t).toLocaleString('en-IN', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' });
  }
  function toast(msg, ms = 3500) {
    const t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(() => (t.hidden = true), ms);
  }

  // ---------- problem counts / status colour ----------
  const errCount = (s) => s.issues.filter((i) => i.sev === 'err').length;
  const warnCount = (s) => s.issues.filter((i) => i.sev === 'warn').length;
  const statusColor = (s) => (errCount(s) ? 'red' : warnCount(s) ? 'amber' : 'green');
  function statusText(s) {
    const e = errCount(s), w = warnCount(s);
    if (e) return `${e} to fix`;
    if (w) return `${w} to check`;
    return 'No problems';
  }
  const SEV = { err: { label: 'Critical', color: 'red', icon: 'octagon-alert' }, warn: { label: 'Check', color: 'amber', icon: 'triangle-alert' }, info: { label: 'Suggestion', color: 'grey', icon: 'lightbulb' } };

  // Per-row lookup of problems: row -> { sev, cells: Map(ci -> {sev,title,sug,why,fix}), issues: Set(idx) }
  function rowIndex(sheet) {
    if (sheet._ix) return sheet._ix;
    const rank = { err: 3, warn: 2, info: 1 };
    const byRow = new Map();
    const get = (r) => { let e = byRow.get(r); if (!e) { e = { sev: null, cells: new Map(), issues: new Set() }; byRow.set(r, e); } return e; };
    sheet.issues.forEach((it, idx) => {
      const touch = (e) => { e.issues.add(idx); if (it.sev !== 'info' && (!e.sev || rank[it.sev] > rank[e.sev])) e.sev = it.sev; };
      for (const [r, ci, sug] of it.cells) {
        const e = get(r); touch(e);
        const cur = e.cells.get(ci);
        if (!cur || rank[it.sev] > rank[cur.sev]) e.cells.set(ci, { sev: it.sev, title: shortTitle(it.title), sug, why: it.why, fix: it.fix, idx });
      }
      if (!it.cells.length) for (const r of it.rows) touch(get(r));
    });
    sheet._ix = byRow;
    return byRow;
  }

  // ---------- local cache of the last copy (this PC only) ----------
  const store = {
    db: null,
    async open() {
      if (this.db) return this.db;
      this.db = await new Promise((res, rej) => {
        const r = indexedDB.open('sheets-dashboard', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('files', { keyPath: 'id' });
        r.onsuccess = () => res(r.result);
        r.onerror = () => rej(r.error);
      });
      return this.db;
    },
    async run(mode, fn) {
      try {
        const db = await this.open();
        return await new Promise((res, rej) => {
          const tx = db.transaction('files', mode);
          const req = fn(tx.objectStore('files'));
          tx.oncomplete = () => res(req && req.result);
          tx.onerror = () => rej(tx.error);
        });
      } catch (e) { return null; }
    },
    all() { return this.run('readonly', (s) => s.getAll()); },
    put(rec) { return this.run('readwrite', (s) => s.put(rec)); },
    del(id) { return this.run('readwrite', (s) => s.delete(id)); },
  };

  // ---------- reading workbooks ----------
  function parseWorkbook(buf) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const out = [];
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      if (!ws || !ws['!ref']) continue;
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: true });
      const model = Analyze.analyzeSheet(name, aoa);
      model.title = Analyze.prettyName(model.name);
      if (model.rows.length) out.push(model);
    }
    return out;
  }

  // ---------- progress between uploads ----------
  // For each file we keep a small summary of the current and the previous version
  // (rows, problems, health). A new version = the file's content changed.
  const PROGRESS_KEY = 'sd-progress';
  function loadProgress() { try { return JSON.parse(localStorage.getItem(PROGRESS_KEY)) || {}; } catch (e) { return {}; } }
  function saveProgress(p) { try { localStorage.setItem(PROGRESS_KEY, JSON.stringify(p)); } catch (e) { /* ignore */ } }
  let progress = loadProgress();

  function hashBuf(buf) { // FNV-1a — just to notice "this file changed"
    const b = new Uint8Array(buf);
    let h = 0x811c9dc5;
    for (let i = 0; i < b.length; i++) { h ^= b[i]; h = Math.imul(h, 0x01000193); }
    return (h >>> 0).toString(16) + ':' + b.length;
  }
  function snapshot(sheets) {
    const tabs = {};
    for (const s of sheets) {
      const cells = {};
      for (const it of s.issues) if (it.sev !== 'info') cells[it.title] = Math.max(it.cells.length, it.rows.length);
      tabs[s.name] = { rows: s.rows.length, score: s.score, cells };
    }
    return tabs;
  }
  function fileHealth(tabs) {
    let w = 0, sum = 0;
    for (const k in tabs) { const t = tabs[k]; const wt = Math.max(1, t.rows); w += wt; sum += t.score * wt; }
    return w ? Math.round(sum / w) : 100;
  }
  function recordVersion(id, buf, sheets) {
    const hash = hashBuf(buf);
    const p = progress[id] || {};
    if (p.cur && p.cur.hash === hash) return; // same content as last time
    progress[id] = { prev: p.cur || null, cur: { hash, at: Date.now(), tabs: snapshot(sheets) } };
    saveProgress(progress);
  }
  // Compare previous vs current version (whole file, or one tab).
  function changes(id, tabName) {
    const p = progress[id];
    if (!p || !p.prev || !p.cur) return null;
    const pick = (v) => (tabName ? (v.tabs[tabName] ? { [tabName]: v.tabs[tabName] } : {}) : v.tabs);
    const A = pick(p.prev), B = pick(p.cur);
    let rowsA = 0, rowsB = 0, fixed = 0, added = 0;
    for (const k in A) rowsA += A[k].rows;
    for (const k in B) rowsB += B[k].rows;
    for (const k of new Set([...Object.keys(A), ...Object.keys(B)])) {
      const a = (A[k] && A[k].cells) || {}, b = (B[k] && B[k].cells) || {};
      for (const t of new Set([...Object.keys(a), ...Object.keys(b)])) {
        const d = (b[t] || 0) - (a[t] || 0);
        if (d < 0) fixed -= d; else added += d;
      }
    }
    return { since: p.prev.at, rows: rowsB - rowsA, fixed, added, healthA: fileHealth(A), healthB: fileHealth(B) };
  }
  function changeLine(c) {
    if (!c) return '';
    const parts = [];
    if (c.rows) parts.push(`${c.rows > 0 ? '+' : ''}${nf.format(c.rows)} rows`);
    if (c.fixed) parts.push(`<span class="up">${plural(c.fixed, 'problem')} fixed</span>`);
    if (c.added) parts.push(`<span class="down">${plural(c.added, 'new problem')}</span>`);
    if (!parts.length) parts.push('no changes to the data');
    const hd = c.healthB - c.healthA;
    parts.push(`health ${c.healthA} → <b>${c.healthB}</b>${hd ? ` <span class="${hd > 0 ? 'up' : 'down'}">${hd > 0 ? '▲' : '▼'}${Math.abs(hd)}</span>` : ''}`);
    return parts.join(' · ');
  }

  // ---------- Fix list (Excel checklist she downloads and works through) ----------
  function fixItems(file, sheet) {
    const items = [];
    for (const it of sheet.issues) {
      if (it.sev === 'info') continue;
      const pri = it.sev === 'err' ? '1 · Fix' : '2 · Check';
      if (!it.cells.length) {
        for (const r of it.rows) items.push([pri, file.name, sheet.name, r, '—', '', it.title, '', '', it.fix]);
        continue;
      }
      for (const [r, ci, sug] of it.cells) {
        const col = sheet.cols[ci];
        const row = sheet.rows.find((x) => x.r === r);
        const isRow = /^Duplicate rows/.test(it.title);
        const now = row && !isRow ? (row.raw[ci] instanceof Date ? fmtDate(row.raw[ci]) : row.raw[ci] ?? '(empty)') : isRow ? '(whole row)' : '';
        items.push([pri, file.name, sheet.name, r, isRow ? '—' : col.letter, isRow ? '' : (col.raw || '(no heading)').trim(), shortTitle(it.title), String(now), sug || '', it.fix]);
      }
    }
    return items;
  }
  function nameItems(file, sheet) {
    return (sheet.names || []).map((n) => [file.name, sheet.name, n.where, n.current, n.suggested, n.why]);
  }
  const allTabs = () => state.files.flatMap((f) => f.sheets.map((sheet) => ({ file: f, sheet })));
  function countFixItems(list) { return list.reduce((s, x) => s + fixItems(x.file, x.sheet).length, 0); }

  function downloadFixList(list, label) {
    if (typeof XLSX === 'undefined') return toast('Could not create the file.');
    const head = ['Done ✓', 'Priority', 'File', 'Tab', 'Row', 'Column', 'Heading', 'Problem', "What's there now", 'Suggested fix', 'How to fix'];
    let rows = list.flatMap((x) => fixItems(x.file, x.sheet));
    rows.sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]) || a[2].localeCompare(b[2]) || a[3] - b[3]);
    rows = rows.map((r) => ['', ...r]);
    const names = list.flatMap((x) => nameItems(x.file, x.sheet));
    const wb = XLSX.utils.book_new();
    const how = [
      ['How to use this fix list'], [''],
      ['1. Open the "Fix list" sheet. Each line is one cell to look at in your Google Sheet.'],
      ['2. Go to the File + Tab, find the Row and Column, and compare with "What\'s there now".'],
      ['3. Make the change in Google Sheets (this app never changes your sheets).'],
      ['4. Type ✓ in the "Done" column so you know where you stopped.'],
      ['5. Start with "1 · Fix" lines — they make totals wrong. "2 · Check" lines may be fine.'],
      ['6. The "Clearer names" sheet suggests better tab and column names. Optional, but makes sheets easier to read.'],
      ['7. Next time, download the sheet again and upload it — the app shows how many problems you fixed.'],
      [''], [`Made on ${new Date().toLocaleString('en-IN')} · ${rows.length} items to look at · ${names.length} name suggestions`],
    ];
    const wsHow = XLSX.utils.aoa_to_sheet(how);
    wsHow['!cols'] = [{ wch: 110 }];
    XLSX.utils.book_append_sheet(wb, wsHow, 'How to use');
    const ws = XLSX.utils.aoa_to_sheet([head, ...rows]);
    ws['!cols'] = [6, 10, 24, 24, 6, 8, 22, 40, 30, 45, 55].map((wch) => ({ wch }));
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: head.length - 1 } }) };
    XLSX.utils.book_append_sheet(wb, ws, 'Fix list');
    const wsN = XLSX.utils.aoa_to_sheet([['File', 'Tab', 'Where', 'Current name', 'Suggested name', 'Why'], ...names]);
    wsN['!cols'] = [24, 26, 12, 34, 38, 36].map((wch) => ({ wch }));
    XLSX.utils.book_append_sheet(wb, wsN, 'Clearer names');
    const d = new Date();
    const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    const safe = String(label).replace(/[\\/:*?"<>|]+/g, ' ').trim();
    try { XLSX.writeFile(wb, `Fix list - ${safe} - ${stamp}.xlsx`); toast(`Fix list downloaded (${nf.format(rows.length)} items)`); }
    catch (e) { console.error('fix list failed'); toast('Could not create the file.'); }
  }

  function addRecord(rec, { save = true } = {}) {
    let sheets;
    try { sheets = parseWorkbook(rec.buf); } catch (e) { console.error('could not read workbook'); return false; }
    if (!sheets.length) return false;
    if (save) { store.put(rec); recordVersion(rec.id, rec.buf, sheets); }
    const file = { id: rec.id, name: rec.name, addedAt: rec.addedAt, sheets };
    const i = state.files.findIndex((f) => f.id === rec.id);
    if (i >= 0) state.files[i] = file; else state.files.push(file);
    const order = state.sources.map((s) => s.id);
    state.files.sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id));
    if (state.sel && !findSel()) state.sel = null; // that tab disappeared -> Overview
    return true;
  }

  // ---------- fetching (read-only) ----------
  // Files uploaded on a site that cannot store files (e.g. the Vercel website) are kept
  // only in this browser (IndexedDB) and never sent to the server. Their ids start with "dev:".
  const isDev = (id) => String(id).startsWith('dev:');
  async function loadSourceList() {
    let list = null;
    try { const r = await fetch('/api/sources'); if (r.ok) list = await r.json(); } catch (e) { /* offline */ }
    const recs = (await store.all()) || [];
    if (!list) { if (!recs.some((r) => isDev(r.id))) return; list = state.sources.filter((s) => !isDev(s.id)); }
    const devs = recs.filter((r) => isDev(r.id)).sort((a, b) => a.name.localeCompare(b.name))
      .map((r) => ({ id: r.id, name: r.name, kind: 'device', file: r.file }));
    list = [...devs, ...list];
    const old = new Map(state.sources.map((s) => [s.id, s]));
    state.sources = list.map((s) => ({ status: 'loading', error: '', ...(old.get(s.id) || {}), ...s }));
    const ids = new Set(list.map((s) => s.id));
    state.files = state.files.filter((f) => ids.has(f.id));
    for (const rec of recs) if (!ids.has(rec.id)) store.del(rec.id);
    let gone = false;
    for (const id of Object.keys(progress)) if (!ids.has(id)) { delete progress[id]; gone = true; }
    if (gone) saveProgress(progress);
  }

  async function fetchSource(src) {
    if (src.kind === 'device') { // already in this browser — just make sure it is loaded
      if (state.files.some((f) => f.id === src.id)) { src.status = 'ok'; return; }
      const rec = ((await store.all()) || []).find((r) => r.id === src.id);
      if (rec && addRecord(rec, { save: false })) { src.status = 'ok'; src.error = ''; }
      else { src.status = 'error'; src.error = 'This file could not be read. Upload it again.'; }
      return;
    }
    try {
      const r = await fetch(src.kind === 'local' ? '/api/local?file=' + encodeURIComponent(src.file) : '/api/drive?id=' + encodeURIComponent(src.id));
      if (!r.ok) {
        const j = await r.json().catch(() => ({}));
        src.status = 'error'; src.error = j.error || 'Could not open this file.';
        return;
      }
      const buf = await r.arrayBuffer();
      const fileName = decodeURIComponent(r.headers.get('X-File-Name') || '').replace(/\.(xlsx|xls|xlsm)$/i, '');
      if (addRecord({ id: src.id, name: fileName || src.name || 'Sheet', addedAt: Date.now(), buf })) { src.status = 'ok'; src.error = ''; }
      else { src.status = 'error'; src.error = 'This file has no readable data.'; }
    } catch (e) {
      src.status = 'error'; src.error = src.kind === 'local' ? 'Could not read this file.' : 'No internet — showing the last copy.';
    }
  }

  function setRefreshBtn(busy) {
    const b = $('#refreshBtn');
    b.disabled = busy;
    b.classList.toggle('spin', busy);
    b.innerHTML = `${I('refresh-cw', 'sm')}<span class="lbl">${busy ? 'Refreshing…' : 'Refresh'}</span>`;
  }
  async function syncAll() {
    if (syncing) return;
    syncing = true;
    setRefreshBtn(true);
    await loadSourceList();
    await Promise.all(state.sources.map(fetchSource));
    syncing = false;
    setRefreshBtn(false);
    state.lastSync = Date.now();
    // Don't redraw while she is typing in a search box.
    const typing = document.activeElement && ['searchBox', 'gsearch'].includes(document.activeElement.id);
    if (typing) { renderNav(); updateStatus(); return; }
    render();
  }

  function updateStatus() {
    const bad = state.sources.filter((s) => s.status === 'error').length;
    $('#syncStatus').innerHTML = state.lastSync
      ? `Updated ${esc(ago(state.lastSync))}${bad ? ` · <span class="bad">${bad} file${bad > 1 ? 's' : ''} not available</span>` : ''}`
      : 'Getting the latest data…';
  }

  // ---------- selection / navigation ----------
  function findSel() {
    if (!state.sel) return null;
    const file = state.files.find((f) => f.id === state.sel.fileId);
    const sheet = file && file.sheets.find((s) => s.name === state.sel.sheet);
    return sheet ? { file, sheet } : null;
  }
  function resetFilters() { state.filters = {}; state.search = ''; state.issue = ''; state.sort = null; state.page = 0; state.cardLimit = CARD_STEP; state.fromIssue = null; }
  function openTab(fileId, sheetName, view, opts = {}) {
    const same = state.sel && state.sel.fileId === fileId && state.sel.sheet === sheetName;
    state.screen = 'app';
    state.sel = { fileId, sheet: sheetName };
    state.open.add(fileId);
    if (!same) { resetFilters(); state.measure = undefined; }
    state.view = view || 'dashboard';
    if (opts.rowsView) setRowsView(opts.rowsView, false);
    if (opts.locate) state.locate = opts.locate;
    closeMobileNav();
    render();
    if (!opts.locate) $('#main').scrollTop = 0;
  }
  function goHome() { state.screen = 'app'; state.sel = null; closeMobileNav(); render(); $('#main').scrollTop = 0; }
  function goStart() { state.screen = 'start'; closeMobileNav(); render(); $('#main').scrollTop = 0; }
  function setRowsView(v, rerender = true) {
    state.rowsView = v === 'table' ? 'table' : 'cards';
    setPref('sd-rows-view', state.rowsView);
    if (rerender) render();
  }

  // ---------- upload (saved into the app's data folder on this PC) ----------
  const APP_HEADERS = { 'X-App': 'sheets-dashboard' };
  // replaceFile = name of the file in the data folder that this upload should replace.
  async function uploadFiles(list, replaceFile) {
    const all = [...list];
    const files = all.filter((f) => /\.(xlsx|xlsm|xls|csv)$/i.test(f.name) && !f.name.startsWith('~$'));
    if (!files.length) { toast(all.length ? 'No Excel files found there (.xlsx, .xls or .csv).' : 'Please choose Excel files.'); return; }
    const drop = $('#drop');
    drop.classList.add('busy');
    $('.drop-title').textContent = `Uploading ${plural(files.length, 'file')}…`;
    const failed = [];
    const toDevice = !state.canSave || (replaceFile && isDev(replaceFile));
    for (const f of files) {
      if (toDevice) { // keep it in this browser only
        try {
          const name = f.name.replace(/\s*\(\d+\)(?=\.[a-z]+$)/i, ''); // "Sales Tracker (1).xlsx" replaces "Sales Tracker.xlsx"
          const id = replaceFile && isDev(replaceFile) ? replaceFile : 'dev:' + name;
          const prev = state.sources.find((s) => s.id === id);
          const rec = { id, name: prev ? prev.name : name.replace(/\.(xlsx|xlsm|xls|csv)$/i, ''), file: name, addedAt: Date.now(), buf: await f.arrayBuffer() };
          if (!state.sources.some((s) => s.id === id)) state.sources.unshift({ id, name: rec.name, kind: 'device', file: name, status: 'ok', error: '' });
          if (addRecord(rec, { save: false })) {
            await store.put(rec); // wait until it is safely stored before the list refreshes
            const file = state.files.find((x) => x.id === id);
            if (file) recordVersion(id, rec.buf, file.sheets);
          } else {
            failed.push(`${f.name}: no readable data in this file`);
            if (!prev) state.sources = state.sources.filter((s) => s.id !== id);
          }
        } catch (e) { failed.push(`${f.name}: could not be read`); }
        continue;
      }
      try {
        const q = '?name=' + encodeURIComponent(f.name) + (replaceFile ? '&replace=' + encodeURIComponent(replaceFile) : '');
        const r = await fetch('/api/upload' + q, { method: 'POST', headers: APP_HEADERS, body: f });
        if (!r.ok) { const j = await r.json().catch(() => ({})); failed.push(`${f.name}: ${j.error || 'upload failed'}`); }
      } catch (e) { failed.push(`${f.name}: upload failed`); }
    }
    drop.classList.remove('busy');
    $('.drop-title').textContent = 'Drag your Excel files here';
    await syncAll();
    const ok = files.length - failed.length;
    if (failed.length) toast((ok ? `${plural(ok, 'file')} uploaded. ` : '') + failed.join(' · '), 8000);
    else toast(replaceFile ? `“${replaceFile.replace(/^dev:/, '')}” replaced with the new copy` : `${plural(ok, 'file')} ${toDevice ? 'added — kept in this browser only' : 'uploaded'}`);
    const skipped = all.length - files.length;
    if (skipped > 0 && !failed.length && all.length > 1) setTimeout(() => toast(`${plural(skipped, 'other file')} in that folder were not Excel files and were skipped`), 3600);
  }

  async function removeSource(src) {
    const isLink = src.kind === 'drive';
    const name = (state.files.find((f) => f.id === src.id) || src).name;
    if (src.kind === 'device') {
      if (!confirm(`Remove “${name}” from this browser?\n\nOnly the copy kept in this browser is removed. Your original file is not touched.`)) return;
      await store.del(src.id);
      delete progress[src.id]; saveProgress(progress);
      state.files = state.files.filter((f) => f.id !== src.id);
      if (state.sel && state.sel.fileId === src.id) state.sel = null;
      await syncAll();
      toast(`Removed “${name}”`);
      return;
    }
    const msg = isLink
      ? `Remove “${name}” from the dashboard?\n\nOnly the link is taken off this list. The Google Sheet itself is not touched.`
      : `Remove “${name}” from the dashboard?\n\nThe app's copy is moved to a “_removed” folder (not deleted). Your original file is not touched.`;
    if (!confirm(msg)) return;
    try {
      const r = await fetch('/api/remove?id=' + encodeURIComponent(src.id), { method: 'POST', headers: APP_HEADERS });
      const j = await r.json().catch(() => ({}));
      if (!r.ok) { toast(j.error || 'Could not remove it.'); return; }
      if (state.sel && state.sel.fileId === src.id) state.sel = null;
      await syncAll();
      toast(`Removed “${name}”`);
    } catch (e) { toast('Could not remove it.'); }
  }

  function renderStart() {
    const rows = state.sources.map((src) => {
      const f = state.files.find((x) => x.id === src.id);
      const isLocal = src.kind === 'local' || src.kind === 'device';
      const ch = f ? changes(f.id) : null;
      const p = f && progress[f.id] && progress[f.id].cur;
      const where = src.kind === 'device' ? ' · kept in this browser only' : src.kind === 'drive' ? ' · from Google Sheets' : '';
      const meta = f ? `${plural(f.sheets.length, 'tab')} · updated ${ago(f.addedAt)}${p ? ` · health ${fileHealth(p.tabs)}` : ''}${ch && ch.fixed ? ` · ${plural(ch.fixed, 'problem')} fixed since last upload` : ''}${where}`
        : src.status === 'error' ? src.error : 'Loading…';
      return `<div class="card file-row ${!f && src.status === 'error' ? 'err' : ''}">
        <span class="card-icon">${I(isLocal ? 'file-spreadsheet' : 'cloud')}</span>
        <div class="grow"><div class="fname">${esc(f ? f.name : src.name)}</div><div class="fmeta">${esc(meta)}</div></div>
        <div class="acts">
          ${isLocal ? `<button class="btn small" data-replace="${esc(src.kind === 'device' ? src.id : src.file)}" type="button" title="Upload a newer copy of this file">${I('replace', 'sm')}Replace</button>` : ''}
          ${src.kind === 'drive' && !state.canSave ? '' : `<button class="btn small" data-remove="${esc(src.id)}" type="button" title="Take this file off the dashboard">${I('trash-2', 'sm')}Remove</button>`}
        </div>
      </div>`;
    }).join('');
    $('#startList').innerHTML = lockCardHtml() + (state.sources.length ? `<div class="list-head">Files in the dashboard (${state.sources.length})</div>${rows}` : '')
      + (state.unlocked ? `<div class="note"><button class="link-btn" data-lock type="button">${I('x', 'sm')}Lock shared sheets on this browser</button></div>` : '');
    $('#startNote').textContent = state.canSave
      ? 'Files stay on this computer. A file with the same name — even “Sales Tracker (1).xlsx” — replaces the older copy. The original Google Sheet is never changed.'
      : 'Private: files are read and kept in this browser on this computer only — they are never uploaded to any server. A file with the same name replaces the older copy. Use the same browser next time to see them again.';
    $('#openDash').disabled = !state.files.length;
  }

  // ---------- sidebar ----------
  function renderNav() {
    const sel = state.sel;
    const onStart = state.screen === 'start';
    const totalErr = allTabs().reduce((s, x) => s + errCount(x.sheet), 0);
    let html = `
      <button class="nav-item ${onStart ? 'active' : ''}" data-upload type="button" title="Upload files"><span class="nav-ic">${I('upload')}</span><span class="label">Upload files</span></button>
      <button class="nav-item ${!onStart && !sel ? 'active' : ''}" data-home type="button" title="Overview"><span class="nav-ic">${I('layout-dashboard')}</span><span class="label">Overview</span>${totalErr ? `<span class="nav-badge" title="${totalErr} problems to fix">${totalErr}</span>` : ''}</button>`;
    if (state.sources.length) html += '<div class="nav-section">Spreadsheets</div>';
    for (const src of state.sources) {
      const f = state.files.find((x) => x.id === src.id);
      const title = f ? f.name : src.name;
      const isOpen = state.open.has(src.id);
      html += `<button class="nav-item nav-file ${isOpen ? 'open' : ''}" data-toggle="${esc(src.id)}" type="button" aria-expanded="${isOpen}">
        <span class="nav-ic chev">${I('chevron-right', 'sm')}</span><span class="label" title="${esc(title)}">${esc(title)}</span></button>`;
      if (!f && src.status === 'error') html += `<div class="nav-error">${esc(src.error)}</div>`;
      if (f && isOpen) {
        html += '<div class="nav-tabs">' + f.sheets.map((s) => `
          <button class="nav-item ${!onStart && sel && sel.fileId === f.id && sel.sheet === s.name ? 'active' : ''}" data-file="${esc(f.id)}" data-sheet="${esc(s.name)}" type="button" title="${esc(s.title)} — ${esc(statusText(s))}">
            <span class="dot ${statusColor(s)}"></span><span class="label">${esc(s.title)}</span>
          </button>`).join('') + '</div>';
      }
    }
    $('#nav').innerHTML = html;

    const n = countFixItems(allTabs());
    const card = $('#sideCard');
    card.hidden = !n || (!onStart && !sel); // the Overview has its own "Next step" card
    if (n) card.innerHTML = `<h3>Clean up your sheets</h3><p>${nf.format(n)} cells to review, each with a suggested fix.</p>
      <button class="btn" data-fixall type="button">${I('download', 'sm')}Download fix list</button>`;
  }

  // ---------- headline number for a tab ----------
  function measureOptions(sheet) {
    const pri = { amount: 0, qty: 1 };
    return sheet.cols.map((c, i) => ({ c, i }))
      .filter((x) => x.c.type === 'number' && x.c.role !== 'price' && x.c.fill >= 0.2)
      .sort((a, b) => (pri[a.c.role] ?? 5) - (pri[b.c.role] ?? 5) || b.c.fill - a.c.fill);
  }
  function defaultMeasure(sheet) {
    const opts = measureOptions(sheet);
    // Prefer a "Total / Net / Value" money column over helper columns like "Difference".
    const best = opts.find((x) => x.c.role === 'amount' && /^total\b|\bnet\b|value|amount|amt/i.test(x.c.name) && !/difference|tax|gst/i.test(x.c.name))
      || opts.find((x) => x.c.role === 'amount' || x.c.role === 'qty');
    return best ? best.i : null;
  }
  function headline(sheet) {
    if (sheet.targets) {
      const tot = sheet.targets.people.find((p) => p.isTotal) || sheet.targets.people[0];
      const s = personStats(tot, sheet.targets.months);
      return { num: `${Math.round(s.pct * 100)}%`, lbl: `of target achieved${s.lastM !== null ? ` (Jan–${MONTHS[s.lastM]})` : ''}` };
    }
    const m = defaultMeasure(sheet);
    if (m !== null) {
      let sum = 0;
      for (const r of sheet.rows) if (typeof r.v[m] === 'number') sum += r.v[m];
      return { num: money(sheet.cols[m], sum), lbl: `${totalLabel(sheet.cols[m].name)} · ${plural(sheet.rows.length, 'row')}` };
    }
    return { num: nf.format(sheet.rows.length), lbl: sheet.rows.length === 1 ? 'row' : 'rows' };
  }

  // ---------- OVERVIEW ----------
  function renderHome() {
    const el = $('#view-home');
    const tabs = allTabs();
    if (!tabs.length) {
      const failed = state.sources.length && state.sources.every((s) => s.status === 'error');
      el.innerHTML = failed
        ? `<div class="card empty"><div class="card-icon">${I('triangle-alert', 'lg')}</div><h2>Couldn't open the sheets</h2><p>${esc(state.sources[0].error)}</p><button class="btn" data-refresh type="button">${I('refresh-cw', 'sm')}Try again</button></div>`
        : !state.sources.length
          ? `<div class="card empty"><div class="card-icon">${I('upload', 'lg')}</div><h2>No spreadsheets yet</h2><p>Upload your Excel files to see the overview.</p><button class="btn primary" data-upload type="button">${I('upload', 'sm')}Upload files</button></div>`
          : `<div class="card empty"><div class="spinner"></div><h2>Getting the latest data…</h2><p>This takes a few seconds.</p></div>`;
      return;
    }
    const toFix = tabs.reduce((s, x) => s + errCount(x.sheet), 0);
    const toCheck = tabs.reduce((s, x) => s + warnCount(x.sheet), 0);
    const tabsErr = tabs.filter((x) => errCount(x.sheet)).length;
    const tabsWarn = tabs.filter((x) => !errCount(x.sheet) && warnCount(x.sheet)).length;
    const totalRows = tabs.reduce((s, x) => s + x.sheet.rows.length, 0);
    const health = fileHealth(Object.fromEntries(tabs.map((x, k) => [k, { rows: x.sheet.rows.length, score: x.sheet.score }])));
    const nItems = countFixItems(tabs);
    // Overall change since last upload, across files that have a previous version.
    const chs = state.files.map((f) => changes(f.id)).filter(Boolean);
    const fixedTotal = chs.reduce((s, c) => s + c.fixed, 0);
    const bad = state.sources.filter((s) => s.status === 'error');

    // What needs attention: the biggest critical issues first, then warnings.
    const att = tabs.flatMap(({ file, sheet }) => sheet.issues.map((it, idx) => ({ file, sheet, it, idx })))
      .filter((x) => x.it.sev !== 'info')
      .sort((a, b) => (a.it.sev === 'err' ? 0 : 1) - (b.it.sev === 'err' ? 0 : 1) || Math.max(b.it.cells.length, b.it.rows.length) - Math.max(a.it.cells.length, a.it.rows.length))
      .slice(0, 6);

    let html = bad.map((s) => `<div class="notice">${I('triangle-alert', 'sm')}<span><b>${esc(s.name)}</b> could not be opened: ${esc(s.error)}</span></div>`).join('');
    if (state.locked) html += `<div class="change-banner" style="margin:0 0 16px">${I('cloud', 'sm')}<span style="flex:1">${plural(state.locked, 'shared Google sheet')} (like the Shell Order Tracker) ${state.locked === 1 ? 'is' : 'are'} protected by a password.</span><button class="btn small" data-upload type="button">Unlock</button></div>`;
    html += `
      <div class="grid kpis">
        <div class="card kpi">
          <div class="card-head"><span class="card-icon">${I('activity')}</span><div class="card-title">Data health</div></div>
          <div class="big">${health}<small>/ 100</small>${fixedTotal ? `<span class="delta">${fixedTotal} fixed</span>` : ''}</div>
          <div class="meter"><i style="width:${health}%"></i></div>
        </div>
        <div class="card kpi">
          <div class="card-head"><span class="card-icon" style="color:var(--red)">${I('octagon-alert')}</span><div class="card-title">Critical issues<small>affect totals and reports</small></div></div>
          <div class="big" style="color:${toFix ? 'var(--red)' : 'inherit'}">${toFix}</div>
          <div class="sub">${toFix ? `in ${plural(tabsErr, 'tab')}` : 'Nothing critical'}</div>
        </div>
        <div class="card kpi">
          <div class="card-head"><span class="card-icon" style="color:var(--amber)">${I('triangle-alert')}</span><div class="card-title">Warnings<small>values to verify</small></div></div>
          <div class="big">${toCheck}</div>
          <div class="sub">${toCheck ? `${plural(tabsWarn + tabsErr, 'tab')} affected` : 'Nothing to check'}</div>
        </div>
        <div class="card kpi dark">
          <div class="card-head"><span class="card-icon">${I('files')}</span><div class="card-title">Spreadsheets<small>${plural(tabs.length, 'tab')}</small></div></div>
          <div class="big">${state.files.length}</div>
          <div class="sub">${nf.format(totalRows)} rows checked</div>
        </div>
      </div>

      <div class="home-row">
        <div class="card">
          <div class="card-head"><span class="card-icon">${I('list-filter')}</span><div class="card-title">Needs attention<small>Biggest problems first — open one to see the exact rows</small></div></div>
          <div class="attention-list">${att.length ? att.map((x) => {
            const s = SEV[x.it.sev];
            const n = Math.max(x.it.cells.length, x.it.rows.length);
            return `<div class="att">
              <span class="sev-ic ${s.color}">${I(s.icon, 'sm')}</span>
              <div class="body"><div class="t" title="${esc(x.it.title)}">${esc(shortTitle(x.it.title))}</div><div class="s">${esc(x.sheet.title)} · ${esc(x.file.name)} · ${plural(n, x.it.cells.length ? 'cell' : 'row')}</div></div>
              <button class="btn small" data-inspect="${esc(x.file.id)}|${esc(x.sheet.name)}|${x.idx}" type="button">${I('eye', 'sm')}Inspect</button>
            </div>`;
          }).join('') : `<div class="empty" style="padding:24px">${I('badge-check', 'lg')}<p>No problems found. Nice work.</p></div>`}</div>
        </div>
        <div class="card next-card">
          <span class="chip" style="background:rgba(0,0,0,.08);align-self:flex-start;margin-bottom:12px">${I('sparkles', 'sm')}Next step</span>
          <h3>${nItems ? 'Work through the fix list' : 'All clear'}</h3>
          <p>${nItems ? `${nf.format(nItems)} cells to review, each with a suggested fix.` : 'No cells need fixing right now.'}</p>
          <ol><li>Download the fix list</li><li>Correct the cells in Google Sheets</li><li>Upload the new copy to see your progress</li></ol>
          <button class="btn" data-fixall type="button" ${nItems ? '' : 'disabled'}>${I('download', 'sm')}Download fix list</button>
        </div>
      </div>`;

    for (const f of state.files) {
      const c = changes(f.id);
      html += `<div class="section-head"><h2>${esc(f.name)}</h2><span class="muted">${plural(f.sheets.length, 'tab')} · updated ${esc(ago(f.addedAt))}</span>
        <span class="muted">${c ? `· Since last upload: ${changeLine(c)}` : '· First upload — progress shows after the next one'}</span></div>
        <div class="grid tiles">`;
      for (const s of f.sheets) {
        const h = headline(s);
        const col = statusColor(s);
        html += `
          <button class="card tile" data-file="${esc(f.id)}" data-sheet="${esc(s.name)}" type="button">
            <div class="tile-top"><span class="tile-name" title="${esc(s.title)}">${esc(s.title)}</span>${I('arrow-up-right', 'sm')}</div>
            <div class="tile-num">${esc(h.num)}</div>
            <div class="tile-lbl" title="${esc(h.lbl)}">${esc(h.lbl)}</div>
            <div class="tile-foot"><span class="chip ${col}"><span class="dot ${col}"></span>${esc(statusText(s))}</span>
              <span class="hbar" title="Data health ${s.score}/100"><i style="width:${s.score}%"></i></span></div>
          </button>`;
      }
      html += '</div>';
    }
    el.innerHTML = html;
  }

  // ---------- filtering (shared by Summary, Cards and Table) ----------
  function categoryCols(sheet) {
    const pri = { status: 0, person: 1, customer: 2, product: 3, group: 4 };
    return sheet.cols.map((c, i) => ({ c, i }))
      .filter((x) => (x.c.type === 'category' || (x.c.type === 'text' && pri[x.c.role] !== undefined)) && x.c.fill >= 0.3)
      .sort((a, b) => (pri[a.c.role] ?? 9) - (pri[b.c.role] ?? 9) || (a.c.type === 'category' ? 0 : 1) - (b.c.type === 'category' ? 0 : 1) || b.c.fill - a.c.fill);
  }
  const rowText = (sheet, r) => (r._t || (r._t = (r.r + ' | ' + r.v.map((v, i) => fmtCell(v, sheet.cols[i])).join(' | ')).toLowerCase()));
  function matchIssue(sheet, r) {
    if (!state.issue) return true;
    const e = rowIndex(sheet).get(r.r);
    if (state.issue === 'ok') return !e || !e.sev;
    if (!e) return false;
    if (state.issue === 'any') return !!e.sev;
    if (state.issue === 'err') return e.sev === 'err';
    if (state.issue === 'warn') return e.sev === 'warn';
    if (state.issue.startsWith('t:')) return e.issues.has(Number(state.issue.slice(2)));
    return true;
  }
  function visibleRows(sheet) {
    const q = state.search.trim().toLowerCase();
    return sheet.rows.filter((r) => {
      if (!matchIssue(sheet, r)) return false;
      for (const k in state.filters) { const v = r.v[k]; if (v === null || Analyze.normKey(v) !== state.filters[k]) return false; }
      if (q && !rowText(sheet, r).includes(q)) return false;
      return true;
    });
  }
  // Groups ignore case/space differences; label = most common spelling.
  function groupBy(rows, ci, measure) {
    const m = new Map();
    for (const r of rows) {
      const v = r.v[ci];
      if (v === null || v === undefined || v === '') continue;
      const k = Analyze.normKey(v instanceof Date ? fmtDate(v) : v);
      let g = m.get(k);
      if (!g) { g = { key: k, total: 0, labels: new Map() }; m.set(k, g); }
      g.labels.set(String(v), (g.labels.get(String(v)) || 0) + 1);
      g.total += measure === null ? 1 : typeof r.v[measure] === 'number' ? r.v[measure] : 0;
    }
    return [...m.values()].map((g) => ({ key: g.key, label: [...g.labels.entries()].sort((a, b) => b[1] - a[1])[0][0], total: g.total }));
  }
  function issueOptions(sheet) {
    const opts = [['', 'All rows'], ['any', 'Rows with problems'], ['err', 'Critical only'], ['warn', 'Warnings only'], ['ok', 'Rows without problems']];
    sheet.issues.forEach((it, idx) => { if (it.cells.length || it.rows.length) opts.push([`t:${idx}`, `${SEV[it.sev].label}: ${shortTitle(it.title)}`]); });
    return opts;
  }
  function issueLabel(sheet) {
    const o = issueOptions(sheet).find((x) => x[0] === state.issue);
    return o ? o[1] : '';
  }
  function filterCols(sheet) { return categoryCols(sheet).filter((x) => groupBy(sheet.rows, x.i, null).length <= 60).slice(0, 3); }

  function renderFilters(sheet) {
    const el = $('#filters');
    const fcols = filterCols(sheet);
    el.innerHTML = `
      <label class="fsearch">${I('search', 'sm')}<span class="sr-only">Search this tab</span>
        <input type="search" id="searchBox" placeholder="Search customer, product, row no., date…" value="${esc(state.search)}"></label>
      <select id="issueSel" aria-label="Filter by problem">
        ${issueOptions(sheet).map(([v, l]) => `<option value="${esc(v)}" ${state.issue === v ? 'selected' : ''}>${esc(l.length > 60 ? l.slice(0, 59) + '…' : l)}</option>`).join('')}
      </select>
      ${fcols.map(({ c, i }) => {
        const opts = groupBy(sheet.rows, i, null).sort((a, b) => a.label.localeCompare(b.label));
        return `<select data-filter="${i}" aria-label="Filter by ${esc(c.name)}">
          <option value="">All ${esc(c.name)}</option>
          ${opts.map((o) => `<option value="${esc(o.key)}" ${state.filters[i] === o.key ? 'selected' : ''}>${esc(o.label.slice(0, 50))}</option>`).join('')}
        </select>`;
      }).join('')}`;
    el.hidden = false;
  }
  // Result count + active filter chips + Clear all (shown under the filters).
  function filterSummary(sheet, shown) {
    const chips = [];
    if (state.search) chips.push(['search', `Search: “${state.search}”`]);
    if (state.issue) chips.push(['issue', `Problem: ${issueLabel(sheet)}`]);
    for (const k in state.filters) {
      const g = groupBy(sheet.rows, Number(k), null).find((x) => x.key === state.filters[k]);
      chips.push([`f:${k}`, `${sheet.cols[k].name}: ${g ? g.label : state.filters[k]}`]);
    }
    return `<div class="filter-bar2">
      <span class="result-count">Showing <b>${nf.format(shown)}</b> of ${plural(sheet.rows.length, 'row')}</span>
      ${chips.map(([k, l]) => `<span class="fchip">${esc(l.length > 70 ? l.slice(0, 69) + '…' : l)}<button type="button" data-unfilter="${esc(k)}" aria-label="Remove filter ${esc(l)}" title="Remove this filter">${I('x', 'sm')}</button></span>`).join('')}
      ${chips.length ? `<button class="link-btn" data-clearall type="button">${I('filter-x', 'sm')}Clear all</button>` : ''}
    </div>`;
  }
  const hasFilters = () => !!(state.search || state.issue || Object.keys(state.filters).length);
  function emptyState() {
    return `<div class="card empty"><div class="card-icon">${I('search', 'lg')}</div><h2>No rows match</h2>
      <p>Try a different search, or clear the filters.</p>
      ${hasFilters() ? `<button class="btn" data-clearall type="button">${I('filter-x', 'sm')}Clear filters</button>` : ''}</div>`;
  }

  // ---------- charts ----------
  function palette() {
    const cs = getComputedStyle(document.documentElement);
    const v = (n) => cs.getPropertyValue(n).trim();
    const dark = document.documentElement.dataset.theme === 'dark';
    return { lav: v('--lavender'), lime: v('--lime'), ink: dark ? '#5a5a66' : v('--side'), text: v('--muted'), grid: v('--line'),
      series: [v('--lavender'), dark ? '#6b6b78' : '#202023', '#9f93ea', dark ? '#8f8f9c' : '#5c5c66'] };
  }
  function destroyCharts() { charts.forEach((c) => { try { c.destroy(); } catch (e) { /* ignore */ } }); charts = []; }
  function makeChart(canvas, cfg) {
    if (!canvas || typeof Chart === 'undefined') return; // charts must never break the page
    try { charts.push(new Chart(canvas, cfg)); } catch (e) { console.warn('chart failed'); }
  }
  const anim = () => (matchMedia('(prefers-reduced-motion: reduce)').matches ? false : { duration: 240 });
  function barCfg(labels, data, color, horizontal, col) {
    const p = palette();
    const fmt = (v) => (col ? money(col, v) : nf.format(v));
    const trim = (l) => (l.length > 26 ? l.slice(0, 25) + '…' : l);
    return {
      type: 'bar',
      data: { labels, datasets: [{ data, backgroundColor: color, borderRadius: 8, borderSkipped: false, maxBarThickness: 30 }] },
      options: {
        indexAxis: horizontal ? 'y' : 'x', maintainAspectRatio: false, animation: anim(),
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: (ctx) => ' ' + (col ? money(col, ctx.raw) : nf.format(ctx.raw) + ' rows') } } },
        scales: horizontal ? {
          x: { grid: { color: p.grid }, border: { display: false }, ticks: { color: p.text, maxRotation: 0, callback: (v) => fmt(v) } },
          y: { grid: { display: false }, border: { display: false }, ticks: { color: p.text, callback: function (v) { return trim(this.getLabelForValue(v)); } } },
        } : {
          x: { grid: { display: false }, border: { display: false }, ticks: { color: p.text, maxRotation: 0, autoSkip: true } },
          y: { grid: { color: p.grid }, border: { display: false }, ticks: { color: p.text, callback: (v) => fmt(v) } },
        },
      },
    };
  }

  // ---------- SUMMARY: target tabs ----------
  const pctColor = (p) => (p >= 1 ? 'green' : p >= 0.8 ? 'amber' : 'red');
  function personStats(p, months) {
    // "To date" = up to the last month with a real (non-zero) achievement.
    const lastIdx = months.reduce((li, m, k) => (p.ach[m] ? k : li), -1);
    const done = months.slice(0, lastIdx + 1);
    const t = done.reduce((s, m) => s + (p.target[m] || 0), 0);
    const a = done.reduce((s, m) => s + (p.ach[m] || 0), 0);
    const year = p.yearTarget ?? months.reduce((s, m) => s + (p.target[m] || 0), 0);
    return { t, a, pct: t ? a / t : 0, year, lastM: lastIdx >= 0 ? months[lastIdx] : null };
  }
  function renderTargets(sheet) {
    const el = $('#view-dashboard');
    const T = sheet.targets, p = palette();
    const unit = T.unit ? T.unit.replace(/^vol,?\s*/i, '').trim() : '';
    const people = T.people.filter((x) => !x.isTotal);
    const total = T.people.find((x) => x.isTotal) || people[0];
    const cards = [total, ...(people.length > 1 ? people : [])].filter(Boolean);
    const n2 = (v) => nf.format(+v.toFixed(2));
    el.innerHTML = `
      <div class="grid kpis">${cards.map((x, k) => {
        const s = personStats(x, T.months);
        const c = pctColor(s.pct);
        return `<div class="card kpi ${k === 0 ? 'dark' : ''}">
          <div class="card-head"><span class="card-icon">${I(x.isTotal ? 'target' : 'user')}</span><div class="card-title">${esc(x.isTotal ? 'Whole team' : x.name)}${x.role && !x.isTotal ? `<small>${esc(x.role)}</small>` : '<small>all people</small>'}</div></div>
          <div class="big">${Math.round(s.pct * 100)}<small>%</small> <span class="chip ${c}" style="vertical-align:middle">${s.pct >= 1 ? 'On target' : 'Behind'}</span></div>
          <div class="sub">${n2(s.a)} of ${n2(s.t)} ${esc(unit)} so far${s.lastM !== null ? ` (Jan–${MONTHS[s.lastM]})` : ''} · year target ${n2(s.year)}</div>
        </div>`;
      }).join('')}</div>
      <div class="grid charts">
        <div class="card wide">
          <div class="card-head"><span class="card-icon">${I('chart-column')}</span><div class="card-title">Target vs achieved, month by month<small>${esc(total.isTotal ? 'whole team' : total.name)}${unit ? ' · ' + esc(unit) : ''}</small></div>
            <div class="legend"><span><i style="background:var(--line-2)"></i>Target</span><span><i style="background:var(--lavender)"></i>Achieved</span><span><i style="background:var(--lime)"></i>Hit target</span></div></div>
          <div class="chart-box"><canvas id="tg-total" aria-label="Target versus achieved by month"></canvas></div></div>
        ${people.length > 1 ? people.map((x, k) => `<div class="card"><div class="card-head"><span class="card-icon">${I('user')}</span><div class="card-title">${esc(x.name)}<small>${esc(x.role)}</small></div></div><div class="chart-box"><canvas id="tg-p${k}"></canvas></div></div>`).join('') : ''}
      </div>`;
    const cfg = (x) => ({
      type: 'bar',
      data: {
        labels: T.months.map((m) => MONTHS[m]),
        datasets: [
          { label: 'Target', data: T.months.map((m) => x.target[m] ?? null), backgroundColor: p.grid, borderRadius: 8, borderSkipped: false, maxBarThickness: 22 },
          { label: 'Achieved', data: T.months.map((m) => x.ach[m] ?? null), borderRadius: 8, borderSkipped: false, maxBarThickness: 22,
            backgroundColor: T.months.map((m) => (x.ach[m] !== undefined && x.ach[m] >= (x.target[m] || 0) ? p.lime : p.lav)) },
        ],
      },
      options: {
        maintainAspectRatio: false, animation: anim(),
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { afterBody: (items) => { const m = T.months[items[0].dataIndex]; const t = x.target[m], a = x.ach[m]; return t && a !== undefined ? `${Math.round((a / t) * 100)}% of target` : ''; } } },
        },
        scales: { x: { grid: { display: false }, border: { display: false }, ticks: { color: p.text } }, y: { grid: { color: p.grid }, border: { display: false }, ticks: { color: p.text } } },
      },
    });
    makeChart($('#tg-total'), cfg(total));
    if (people.length > 1) people.forEach((x, k) => makeChart($('#tg-p' + k), cfg(x)));
  }

  // ---------- SUMMARY: normal tabs ----------
  function renderDashboard(sheet) {
    destroyCharts();
    if (sheet.targets) return renderTargets(sheet);
    const el = $('#view-dashboard');
    if (state.measure === undefined) state.measure = defaultMeasure(sheet);
    const rows = visibleRows(sheet);
    const meas = state.measure;
    const mCol = meas === null ? null : sheet.cols[meas];
    const p = palette();
    const valueOf = (r) => (meas === null ? 1 : typeof r.v[meas] === 'number' ? r.v[meas] : 0);

    const kpis = [{ icon: 'rows-3', lbl: 'Rows', num: nf.format(rows.length), sub: rows.length !== sheet.rows.length ? `out of ${nf.format(sheet.rows.length)}` : 'entries in this tab' }];
    for (const { c, i } of measureOptions(sheet).slice(0, 2)) {
      let sum = 0;
      for (const r of rows) if (typeof r.v[i] === 'number') sum += r.v[i];
      kpis.push({ icon: isMoney(c) ? 'indian-rupee' : 'boxes', lbl: totalLabel(c.name), num: money(c, sum), sub: isMoney(c) ? '₹' + nf.format(Math.round(sum)) : 'added up' });
    }
    const dateCol = sheet.cols.map((c, i) => ({ c, i })).filter((x) => x.c.type === 'date').sort((a, b) => b.c.fill - a.c.fill)[0];
    let byMonth = null;
    if (dateCol) {
      byMonth = new Map();
      for (const r of rows) {
        const d = r.v[dateCol.i];
        if (!(d instanceof Date)) continue;
        const k = d.getFullYear() * 12 + d.getMonth();
        byMonth.set(k, (byMonth.get(k) || 0) + valueOf(r));
      }
      const keys = [...byMonth.keys()].sort((a, b) => a - b);
      if (keys.length) {
        const last = keys[keys.length - 1], cur = byMonth.get(last), prev = byMonth.get(last - 1) || 0;
        const ch = prev ? (cur - prev) / prev : null;
        kpis.push({
          icon: 'calendar', lbl: `Latest month · ${MONTHS[last % 12]} ${Math.floor(last / 12)}`,
          num: mCol ? money(mCol, cur) : plural(cur, 'row'),
          subHtml: ch === null ? 'no data the month before' : `<span class="${ch >= 0 ? 'up' : 'down'}">${ch >= 0 ? '▲' : '▼'} ${Math.abs(Math.round(ch * 100))}%</span> vs ${MONTHS[(last - 1) % 12]}`,
        });
      }
    }
    const cust = sheet.cols.findIndex((c) => c.role === 'customer');
    if (cust >= 0 && kpis.length < 4) kpis.push({ icon: 'user', lbl: sheet.cols[cust].name, num: nf.format(groupBy(rows, cust, null).length), sub: 'different names' });

    const breakdowns = categoryCols(sheet)
      .filter((x, k, arr) => !x.c.role || arr.findIndex((y) => y.c.role === x.c.role) === k) // one chart per kind
      .filter((x) => groupBy(sheet.rows, x.i, null).length > 1)
      .slice(0, 4);
    const mOpts = measureOptions(sheet);
    const what = mCol ? mCol.name : 'Number of rows';

    el.innerHTML = `
      ${filterSummary(sheet, rows.length)}
      <div class="grid kpis">${kpis.map((k) => `
        <div class="card kpi"><div class="card-head"><span class="card-icon">${I(k.icon)}</span><div class="card-title" title="${esc(k.lbl)}">${esc(k.lbl)}</div></div>
        <div class="big" style="font-size:32px">${esc(k.num)}</div><div class="sub">${k.subHtml || esc(k.sub)}</div></div>`).join('')}
      </div>
      ${!rows.length ? emptyState() : `
      ${mOpts.length ? `<div class="measure"><label for="measureSel">Charts show</label>
        <select id="measureSel">
          ${mOpts.map(({ c, i }) => `<option value="${i}" ${meas === i ? 'selected' : ''}>${esc(c.name)}</option>`).join('')}
          <option value="" ${meas === null ? 'selected' : ''}>Number of rows</option>
        </select></div>` : ''}
      <div class="grid charts">
        ${byMonth && byMonth.size ? `<div class="card wide"><div class="card-head"><span class="card-icon">${I('chart-column')}</span><div class="card-title">${esc(what)} by month<small>based on ${esc(dateCol.c.name)} · latest month highlighted</small></div></div><div class="chart-box"><canvas id="ch-trend"></canvas></div></div>` : ''}
        ${breakdowns.map(({ c }, k) => `<div class="card"><div class="card-head"><span class="card-icon">${I(c.role === 'person' || c.role === 'customer' ? 'user' : c.role === 'product' ? 'package' : 'tag')}</span><div class="card-title">Top 10 by ${esc(c.name)}</div></div><div class="chart-box"><canvas id="ch-b${k}"></canvas></div></div>`).join('')}
        ${!(byMonth && byMonth.size) && !breakdowns.length ? `<div class="card wide muted">This tab has no columns suitable for charts. Open <b>Rows</b> to see everything.</div>` : ''}
      </div>`}`;

    if (byMonth && byMonth.size) {
      const keys = [...byMonth.keys()].sort((a, b) => a - b);
      const end = keys[keys.length - 1], start = Math.max(keys[0], end - 35);
      const labels = [], data = [];
      for (let k = start; k <= end; k++) { labels.push(`${MONTHS[k % 12]} ${String(Math.floor(k / 12)).slice(2)}`); data.push(byMonth.get(k) || 0); }
      makeChart($('#ch-trend'), barCfg(labels, data, data.map((_, i) => (i === data.length - 1 ? p.lime : p.lav)), false, mCol));
    }
    breakdowns.forEach(({ i }, k) => {
      const groups = groupBy(rows, i, meas).sort((a, b) => b.total - a.total);
      const top = groups.slice(0, 10);
      const rest = groups.slice(10).reduce((s, g) => s + g.total, 0);
      const labels = top.map((g) => g.label), data = top.map((g) => g.total);
      if (rest) { labels.push(`All others (${groups.length - 10})`); data.push(rest); }
      makeChart($('#ch-b' + k), barCfg(labels, data, p.series[k % p.series.length], true, mCol));
    });
  }

  // ---------- PROBLEMS ----------
  function rowsText(rows) { return rows.length > 12 ? rows.slice(0, 12).join(', ') + ` … +${rows.length - 12} more` : rows.join(', '); }
  function issueCard(sheet, it, idx) {
    const s = SEV[it.sev];
    const colsHit = [...new Set(it.cells.map((c) => c[1]))].map((ci) => sheet.cols[ci]).filter(Boolean);
    const where = (colsHit.length && !/^Duplicate rows/.test(it.title)
      ? colsHit.slice(0, 3).map((c) => `${c.name} (column ${c.letter})`).join(', ') + (colsHit.length > 3 ? ` +${colsHit.length - 3} more` : '') + ' · '
      : '') + (it.rows.length ? `rows ${rowsText(it.rows)}` : '');
    const n = it.rows.length;
    return `
      <article class="card issue ${s.color}">
        <div class="issue-top">
          <span class="sev-ic ${s.color}">${I(s.icon, 'sm')}</span>
          <div class="issue-title">${esc(shortTitle(it.title))}</div>
          ${n ? `<button class="btn small" data-inspect-here="${idx}" type="button">${I('eye', 'sm')}Inspect ${plural(n, 'row')}</button>` : ''}
        </div>
        <div class="issue-grid">
          <div><div class="k">${I('circle-help', 'sm')}What's wrong</div><div class="v">${esc(it.detail)}</div></div>
          <div><div class="k">${I('locate', 'sm')}Where</div><div class="v">${esc(where || '—')}</div></div>
          <div><div class="k">${I('info', 'sm')}Why it matters</div><div class="v">${esc(it.why || '—')}</div></div>
          <div><div class="k">${I('wrench', 'sm')}What to do</div><div class="v">${esc(it.fix || '—')}</div></div>
        </div>
      </article>`;
  }
  function renderHealth(sheet) {
    const el = $('#view-health');
    const list = sheet.issues.map((it, idx) => ({ it, idx }));
    const errs = list.filter((x) => x.it.sev === 'err'), warns = list.filter((x) => x.it.sev === 'warn'), infos = list.filter((x) => x.it.sev === 'info');
    const color = errs.length ? 'red' : warns.length ? 'amber' : 'green';
    const title = errs.length ? `${plural(errs.length, 'critical issue')} to fix` : warns.length ? 'Looks fine — a few things to verify' : 'Everything looks good';
    const sel = findSel();
    const nItems = fixItems(sel.file, sheet).length;
    const ch = changes(sel.file.id, sheet.name);
    const names = sheet.names || [];
    let html = `
      <div class="card">
        <div class="verdict">
          <span class="sev-ic ${color === 'green' ? 'grey' : color}" style="${color === 'green' ? 'background:var(--green-soft);color:var(--green)' : ''}">${I(errs.length ? 'octagon-alert' : warns.length ? 'triangle-alert' : 'circle-check', 'lg')}</span>
          <div style="flex:1;min-width:220px">
            <div class="verdict-title">${esc(title)}</div>
            <div class="muted">${plural(sheet.rows.length, 'row')} checked · data health ${sheet.score}/100 · row numbers match the sheet.</div>
          </div>
          ${nItems || names.length ? `<button class="btn primary" data-fixtab type="button">${I('download', 'sm')}Fix list for this tab</button>` : ''}
        </div>
        ${ch ? `<div class="change-banner">${I('trending-up', 'sm')}<span>Since the last upload (${esc(ago(ch.since))}): ${changeLine(ch)}</span></div>` : ''}
      </div>`;
    const group = (icon, color2, h, sub, items) => items.length
      ? `<div class="group-head"><span class="sev-ic ${color2}" style="width:30px;height:30px">${I(icon, 'sm')}</span><h2>${h}</h2><span class="muted">${sub}</span></div>${items.map((x) => issueCard(sheet, x.it, x.idx)).join('')}`
      : '';
    html += group('octagon-alert', 'red', `Critical issues (${errs.length})`, 'likely to make totals and reports wrong', errs);
    html += group('triangle-alert', 'amber', `Warnings (${warns.length})`, 'values to verify — may be fine', warns);
    if (infos.length || names.length) {
      html += `<div class="group-head"><span class="sev-ic grey" style="width:30px;height:30px">${I('lightbulb', 'sm')}</span><h2>Suggestions (${infos.length + (names.length ? 1 : 0)})</h2><span class="muted">optional clean-up — nothing is wrong</span></div>`;
      html += infos.map((x) => issueCard(sheet, x.it, x.idx)).join('');
      if (names.length) {
        html += `<div class="card issue grey">
          <div class="issue-top"><span class="sev-ic grey">${I('pen-line', 'sm')}</span><div class="issue-title">Clearer names for ${plural(names.length, 'tab or column', 'tabs and columns')}</div></div>
          <p class="muted" style="margin:8px 0 6px;font-size:13.5px">Rename these in the sheet so everyone understands them. Optional.</p>
          <div style="overflow-x:auto"><table class="names"><thead><tr><th>Where</th><th>Now</th><th>Better</th><th>Why</th></tr></thead>
          <tbody>${names.map((n) => `<tr><td class="why">${esc(n.where)}</td><td class="old">${esc(n.current)}</td><td class="new">${esc(n.suggested)}</td><td class="why">${esc(n.why)}</td></tr>`).join('')}</tbody></table></div>
        </div>`;
      }
    }
    el.innerHTML = html;
  }

  // ---------- ROWS: shared ----------
  // Which columns a row card shows (chosen once per tab from the column roles).
  function cardLayout(sheet) {
    if (sheet._card) return sheet._card;
    const cols = sheet.cols.map((c, i) => ({ c, i })).filter((x) => x.c.type !== 'empty');
    const best = (pred) => cols.filter(pred).sort((a, b) => b.c.fill - a.c.fill)[0];
    const named = (x) => x.c.type === 'text' || x.c.type === 'category'; // names, not codes or numbers
    const title = best((x) => x.c.role === 'customer' && named(x)) || best((x) => x.c.role === 'product' && named(x)) || best(named);
    const sub = best((x) => x.c.role === 'product' && named(x) && (!title || x.i !== title.i));
    const used = new Set([title && title.i, sub && sub.i]);
    const meta = [];
    const add = (x, icon) => { if (x && !used.has(x.i) && meta.length < 6) { meta.push({ i: x.i, icon }); used.add(x.i); } };
    cols.filter((x) => x.c.type === 'date').sort((a, b) => b.c.fill - a.c.fill).slice(0, 2).forEach((x) => add(x, 'calendar'));
    add(best((x) => x.c.role === 'person'), 'user');
    add(best((x) => x.c.type !== 'date' && /\b(po|invoice|order|bill)\b/i.test(x.c.name)), 'hash');
    add(best((x) => x.c.role === 'qty'), 'boxes');
    const m = defaultMeasure(sheet);
    if (m !== null && isMoney(sheet.cols[m])) add({ i: m, c: sheet.cols[m] }, 'indian-rupee');
    add(best((x) => x.c.role === 'status'), 'tag');
    for (const x of cols) { if (meta.length >= 4) break; add(x, 'file-text'); } // small tabs: fill up
    sheet._card = { title: title ? title.i : null, sub: sub ? sub.i : null, meta };
    return sheet._card;
  }
  function cellView(sheet, r, ci) {
    const v = r.v[ci];
    return v === null && r.raw[ci] != null ? String(r.raw[ci]) : fmtCell(v, sheet.cols[ci]);
  }
  function sortedRows(sheet) {
    let rows = visibleRows(sheet);
    if (state.sort) {
      const { ci, dir } = state.sort;
      const val = (r) => { const v = r.v[ci]; return v instanceof Date ? v.getTime() : v; };
      rows = rows.slice().sort((a, b) => {
        const x = val(a), y = val(b);
        if (x === null || x === undefined) return 1;
        if (y === null || y === undefined) return -1;
        if (typeof x === 'number' && typeof y === 'number') return (x - y) * dir;
        return String(x).localeCompare(String(y), 'en', { numeric: true }) * dir;
      });
    }
    return rows;
  }

  function renderRows(sheet) {
    const el = $('#view-data');
    // If asked to locate a row that the current filters hide, clear filters so it can be shown.
    if (state.locate && !visibleRows(sheet).some((r) => r.r === state.locate)) {
      if (sheet.rows.some((r) => r.r === state.locate)) { resetFilters(); renderFilters(sheet); toast(`Filters were cleared to show row ${state.locate}`); }
      else { toast(`Row ${state.locate} is not in this tab's data`); state.locate = null; }
    }
    const rows = sortedRows(sheet);
    const head = `
      ${state.fromIssue ? `<div class="focus-banner">${I('list-filter', 'sm')}<span class="grow">Showing rows for <b>${esc(state.fromIssue.title)}</b></span>
        <button class="btn small" data-backissues type="button">${I('arrow-left', 'sm')}Back to problems</button>
        <button class="btn small" data-clearall type="button">${I('filter-x', 'sm')}Show all rows</button></div>` : ''}
      <div class="rows-bar">
        ${filterSummary(sheet, rows.length).replace('class="filter-bar2"', 'class="filter-bar2" style="margin:0"')}
        <div class="seg small" role="tablist" aria-label="Choose layout">
          <button type="button" role="tab" data-rowsview="cards" class="${state.rowsView === 'cards' ? 'active' : ''}" aria-selected="${state.rowsView === 'cards'}">${I('layout-grid', 'sm')}Cards</button>
          <button type="button" role="tab" data-rowsview="table" class="${state.rowsView === 'table' ? 'active' : ''}" aria-selected="${state.rowsView === 'table'}">${I('table-2', 'sm')}Table</button>
        </div>
      </div>`;
    if (!rows.length) { el.innerHTML = head + emptyState(); return; }
    el.innerHTML = head + (state.rowsView === 'cards' ? cardsHtml(sheet, rows) : tableHtml(sheet, rows));
    if (state.rowsView === 'table' && state.locate) {
      const tr = el.querySelector(`tr[data-r="${state.locate}"]`);
      if (tr) {
        tr.classList.add('located');
        requestAnimationFrame(() => { tr.scrollIntoView({ block: 'center', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' }); tr.focus({ preventScroll: true }); });
      }
      state.locate = null;
    }
  }

  // ---------- ROWS: cards ----------
  function cardsHtml(sheet, rows) {
    const L = cardLayout(sheet);
    const ix = rowIndex(sheet);
    const shown = rows.slice(0, state.cardLimit);
    const cards = shown.map((r) => {
      const e = ix.get(r.r);
      const sev = e && e.sev;
      const title = L.title !== null ? cellView(sheet, r, L.title) : '';
      const sub = L.sub !== null ? cellView(sheet, r, L.sub) : '';
      const fields = L.meta.slice();
      // A flagged field is always shown, even if it is not one of the usual card fields.
      if (e) for (const [ci, c] of e.cells) if (c.sev !== 'info' && ci !== L.title && ci !== L.sub && !fields.some((f) => f.i === ci)) fields.unshift({ i: ci, icon: 'triangle-alert' });
      const firstIssue = e ? [...e.cells.values()].filter((c) => c.sev !== 'info').sort((a, b) => (a.sev === 'err' ? 0 : 1) - (b.sev === 'err' ? 0 : 1))[0] : null;
      const issueOnlyRow = e && !firstIssue && e.sev ? sheet.issues[[...e.issues].find((k) => sheet.issues[k].sev !== 'info')] : null;
      const titleFlag = e && L.title !== null && e.cells.get(L.title);
      return `<article class="card rcard ${sev ? 'flag-' + SEV[sev].color : ''}">
        <div class="rcard-top">
          <span class="rownum" title="Row number in the sheet">${I('hash', 'sm')}Row ${r.r}</span>
          ${sev ? `<span class="chip ${SEV[sev].color}">${I(SEV[sev].icon, 'sm')}${sev === 'err' ? 'Needs fixing' : 'Check'}</span>` : '<span class="chip green">' + I('circle-check', 'sm') + 'OK</span>'}
        </div>
        <div>
          <div class="rcard-title" style="${titleFlag && titleFlag.sev !== 'info' ? `color:var(--${titleFlag.sev === 'err' ? 'red' : 'amber'})` : ''}" title="${esc(title)}">${esc(title || '(no ' + (L.title !== null ? sheet.cols[L.title].name : 'name') + ')')}</div>
          ${sub ? `<div class="rcard-sub" title="${esc(sub)}">${I('package', 'sm')} ${esc(sub)}</div>` : ''}
        </div>
        <div class="rmeta">${fields.filter((f) => (e && e.cells.has(f.i)) || cellView(sheet, r, f.i) !== '').slice(0, 6).map((f) => {
          const c = sheet.cols[f.i];
          const fl = e && e.cells.get(f.i);
          const cls = fl && fl.sev === 'err' ? 'bad' : fl && fl.sev === 'warn' ? 'warn' : '';
          const val = cellView(sheet, r, f.i) || '—';
          return `<div class="f ${cls}" ${fl && fl.sev !== 'info' ? `title="${esc(fl.title)}"` : ''}><div class="k">${I(f.icon, 'sm')}${esc(c.name)}</div><div class="v" title="${esc(val)}">${esc(val)}</div></div>`;
        }).join('')}</div>
        ${firstIssue ? `<div class="rissue ${SEV[firstIssue.sev].color}">${I(SEV[firstIssue.sev].icon, 'sm')}<span>${esc(firstIssue.title)}</span></div>`
          : issueOnlyRow ? `<div class="rissue ${SEV[issueOnlyRow.sev].color}">${I(SEV[issueOnlyRow.sev].icon, 'sm')}<span>${esc(shortTitle(issueOnlyRow.title))}</span></div>` : ''}
        <div class="rcard-acts">
          <button class="btn small" data-details="${r.r}" type="button">${I('eye', 'sm')}Details</button>
          <button class="btn small" data-intable="${r.r}" type="button">${I('table-2', 'sm')}View in table</button>
        </div>
      </article>`;
    }).join('');
    const more = rows.length > shown.length
      ? `<div class="more-wrap"><button class="btn" data-more type="button">Show ${nf.format(Math.min(CARD_STEP, rows.length - shown.length))} more <span class="muted">(${nf.format(shown.length)} of ${nf.format(rows.length)} shown)</span></button></div>` : '';
    return `<div class="row-cards">${cards}</div>${more}`;
  }

  // ---------- ROWS: table ----------
  function tableHtml(sheet, rows) {
    const pages = Math.max(1, Math.ceil(rows.length / PAGE));
    if (state.locate) { const k = rows.findIndex((r) => r.r === state.locate); if (k >= 0) state.page = Math.floor(k / PAGE); }
    state.page = Math.min(state.page, pages - 1);
    const slice = rows.slice(state.page * PAGE, state.page * PAGE + PAGE);
    const ix = rowIndex(sheet);
    // When filtered by one problem, put its column(s) first so they are visible without scrolling.
    let focusCols = [];
    if (state.issue.startsWith('t:')) { const it = sheet.issues[Number(state.issue.slice(2))]; if (it) focusCols = [...new Set(it.cells.map((c) => c[1]))]; }
    const order = [...focusCols, ...sheet.cols.map((_, i) => i).filter((i) => !focusCols.includes(i))];
    return `
      <div class="table-legend"><span><i style="background:var(--red-soft);border:1px solid var(--red)"></i>Critical</span><span><i style="background:var(--amber-soft);border:1px solid var(--amber)"></i>Check</span><span>Hover a coloured cell to see why · click a heading to sort</span></div>
      <div class="table-wrap" tabindex="0" aria-label="Rows of ${esc(sheet.title)}"><table class="data">
        <thead><tr><th class="rn" scope="col">Row</th>${order.map((i) => {
          const c = sheet.cols[i];
          const s = state.sort && state.sort.ci === i;
          return `<th scope="col" class="${s ? 'sorted' : ''}" aria-sort="${s ? (state.sort.dir > 0 ? 'ascending' : 'descending') : 'none'}"><button type="button" data-sort="${i}" title="Column ${c.letter} — click to sort">${esc(c.name)}${s ? I(state.sort.dir > 0 ? 'chevron-down' : 'chevron-right', 'sm') : ''}</button></th>`;
        }).join('')}</tr></thead>
        <tbody>${slice.map((r) => {
          const e = ix.get(r.r);
          return `<tr data-r="${r.r}" tabindex="-1"><td class="rn">${r.r}</td>${order.map((i) => {
            const fl = e && e.cells.get(i);
            const flag = fl && fl.sev !== 'info';
            const shown = cellView(sheet, r, i);
            const cls = (typeof r.v[i] === 'number' && sheet.cols[i].type === 'number' ? 'num ' : '') + (flag ? 'c-' + fl.sev : '');
            const tip = flag ? `${fl.title}${fl.sug ? ' — Suggested: ' + fl.sug : ''}` : shown;
            return `<td class="${cls}" title="${esc(tip)}"${flag ? ` aria-label="${esc(shown || 'empty')}: ${esc(fl.title)}"` : ''}>${esc(shown)}</td>`;
          }).join('')}</tr>`;
        }).join('')}</tbody>
      </table></div>
      <div class="pager">${plural(rows.length, 'row')}
        ${pages > 1 ? `<button class="btn small" data-page="-1" ${state.page === 0 ? 'disabled' : ''} type="button">${I('chevron-left', 'sm')}Previous</button> Page ${state.page + 1} of ${pages}
        <button class="btn small" data-page="1" ${state.page >= pages - 1 ? 'disabled' : ''} type="button">Next${I('chevron-right', 'sm')}</button>` : ''}
      </div>`;
  }

  // ---------- details drawer ----------
  let lastFocus = null;
  function openDrawer(sheet, rn) {
    const r = sheet.rows.find((x) => x.r === rn);
    if (!r) return;
    const e = rowIndex(sheet).get(rn);
    const L = cardLayout(sheet);
    $('#drawerKicker').textContent = `${sheet.title} · Row ${rn}`;
    $('#drawerTitle').textContent = (L.title !== null && cellView(sheet, r, L.title)) || `Row ${rn}`;
    const rowLevel = e ? [...e.issues].map((k) => sheet.issues[k]).filter((it) => !it.cells.length && it.sev !== 'info') : [];
    $('#drawerBody').innerHTML = rowLevel.map((it) => `<div class="dfield"><div class="why"><b>${esc(shortTitle(it.title))}.</b> ${esc(it.fix)}</div></div>`).join('')
      + sheet.cols.map((c, i) => {
        const fl = e && e.cells.get(i);
        const flag = fl && fl.sev !== 'info';
        const val = cellView(sheet, r, i);
        return `<div class="dfield ${flag ? (fl.sev === 'err' ? 'bad' : 'warn') : ''}">
          <div class="k">${esc(c.name)} · column ${c.letter}</div><div class="v">${esc(val || '—')}</div>
          ${flag ? `<div class="why">${esc(fl.title)}${fl.why ? `. ${esc(fl.why)}` : ''}${fl.sug ? `<br><b>Suggested:</b> ${esc(fl.sug)}` : ''}</div>` : ''}
        </div>`;
      }).join('');
    $('#drawerFoot').innerHTML = `<button class="btn primary" data-intable="${rn}" type="button">${I('table-2', 'sm')}View in table</button><button class="btn" data-closedrawer type="button">Close</button>`;
    lastFocus = document.activeElement;
    $('#drawer').classList.add('open');
    $('#drawer').setAttribute('aria-hidden', 'false');
    const scrim = $('#scrim'); scrim.hidden = false; scrim.classList.add('over-drawer');
    setTimeout(() => $('#drawerClose').focus(), 50);
  }
  function closeDrawer() {
    if (!$('#drawer').classList.contains('open')) return;
    $('#drawer').classList.remove('open');
    $('#drawer').setAttribute('aria-hidden', 'true');
    const scrim = $('#scrim'); scrim.classList.remove('over-drawer');
    if (!document.body.classList.contains('nav-open')) scrim.hidden = true;
    if (lastFocus && lastFocus.focus) lastFocus.focus();
  }

  // ---------- global search ----------
  let gHits = [];
  function globalSearch(q) {
    const box = $('#gresults');
    q = q.trim().toLowerCase();
    if (q.length < 2) { box.hidden = true; gHits = []; return; }
    const tabs = [], rows = [];
    for (const { file, sheet } of allTabs()) {
      if (sheet.title.toLowerCase().includes(q) || sheet.name.toLowerCase().includes(q) || file.name.toLowerCase().includes(q)) tabs.push({ file, sheet });
      if (rows.length >= 8) continue;
      const L = cardLayout(sheet);
      let perTab = 0;
      for (const r of sheet.rows) {
        if (rows.length >= 8 || perTab >= 3) break; // spread results across tabs
        if (rowText(sheet, r).includes(q)) {
          perTab++;
          const t = L.title !== null ? cellView(sheet, r, L.title) : '';
          const hit = r.v.map((v, i) => cellView(sheet, r, i)).find((s) => s && s.toLowerCase().includes(q)) || '';
          rows.push({ file, sheet, r: r.r, t, hit });
        }
      }
    }
    gHits = [...tabs.slice(0, 5).map((x) => ({ type: 'tab', ...x })), ...rows.map((x) => ({ type: 'row', ...x }))];
    let html = '';
    if (tabs.length) html += `<div class="gr-head">Tabs</div>` + tabs.slice(0, 5).map((x, k) => `<button class="gr-item" data-g="${k}" type="button" role="option">${I('file-spreadsheet', 'sm')}<span><span class="t">${esc(x.sheet.title)}</span><br><span class="s">${esc(x.file.name)} · ${plural(x.sheet.rows.length, 'row')}</span></span></button>`).join('');
    const off = Math.min(tabs.length, 5);
    if (rows.length) html += `<div class="gr-head">Rows</div>` + rows.map((x, k) => `<button class="gr-item" data-g="${off + k}" type="button" role="option">${I('rows-3', 'sm')}<span><span class="t">Row ${x.r}${x.t ? ' · ' + esc(x.t) : ''}</span><br><span class="s">${esc(x.sheet.title)} — ${esc(x.hit)}</span></span></button>`).join('');
    box.innerHTML = html || `<div class="gr-empty">Nothing found for “${esc(q)}”.</div>`;
    box.hidden = false;
  }
  function pickHit(k) {
    const h = gHits[k];
    if (!h) return;
    $('#gresults').hidden = true;
    $('#gsearch').value = '';
    if (h.type === 'tab') openTab(h.file.id, h.sheet.name, 'dashboard');
    else openTab(h.file.id, h.sheet.name, 'data', { rowsView: 'table', locate: h.r });
  }

  // ---------- mobile navigation ----------
  const isMobile = () => matchMedia('(max-width: 900px)').matches;
  function openMobileNav() { document.body.classList.add('nav-open'); $('#scrim').hidden = false; }
  function closeMobileNav() { if (!document.body.classList.contains('nav-open')) return; document.body.classList.remove('nav-open'); if (!$('#drawer').classList.contains('open')) $('#scrim').hidden = true; }

  // ---------- main render ----------
  function render() {
    renderNav();
    updateStatus();
    const start = state.screen === 'start';
    const sel = start ? null : findSel();
    const home = !start && !sel;
    $('#view-start').hidden = !start;
    $('#view-home').hidden = !home;
    ['dashboard', 'health', 'data'].forEach((v) => ($('#view-' + v).hidden = !sel || state.view !== v));
    $('#tabs').hidden = !sel;
    $('#backBtn').hidden = !sel;
    if (!sel) { destroyCharts(); $('#filters').hidden = true; }
    if (start) {
      $('#title').textContent = 'Upload your files';
      $('#subtitle').textContent = 'Add the latest Excel downloads, then open the dashboard.';
      renderStart();
      return;
    }
    if (home) {
      $('#title').textContent = 'Overview';
      $('#subtitle').textContent = "What's happening in your data, what needs attention, and what to do next.";
      renderHome();
      return;
    }
    const { file, sheet } = sel;
    $('#title').textContent = sheet.title;
    $('#subtitle').textContent = `${file.name} · ${plural(sheet.rows.length, 'row')} · data health ${sheet.score}/100`;
    const e = errCount(sheet), w = warnCount(sheet);
    $$('#tabs button').forEach((b) => {
      const v = b.dataset.view;
      b.classList.toggle('active', v === state.view);
      b.setAttribute('aria-selected', v === state.view);
      b.innerHTML = v === 'dashboard' ? `${I('chart-column', 'sm')}Summary`
        : v === 'health' ? `${I('triangle-alert', 'sm')}Problems <span class="count ${e ? 'red' : w ? 'amber' : ''}">${e || w || '✓'}</span>`
          : `${I('rows-3', 'sm')}Rows`;
    });
    if (state.view === 'health') { $('#filters').hidden = true; destroyCharts(); renderHealth(sheet); }
    else if (state.view === 'dashboard' && sheet.targets) { $('#filters').hidden = true; renderDashboard(sheet); }
    else { renderFilters(sheet); if (state.view === 'dashboard') renderDashboard(sheet); else { destroyCharts(); renderRows(sheet); } }
  }
  function rerenderBody() {
    const sel = findSel(); if (!sel) return;
    if (state.view === 'dashboard') renderDashboard(sel.sheet); else if (state.view === 'data') renderRows(sel.sheet);
  }

  // ---------- events ----------
  function inspectIssue(fileId, sheetName, idx) {
    const file = state.files.find((f) => f.id === fileId);
    const sheet = file && file.sheets.find((s) => s.name === sheetName);
    const it = sheet && sheet.issues[idx];
    if (!it) return;
    openTab(fileId, sheetName, 'data');
    resetFilters();
    state.issue = `t:${idx}`;
    state.fromIssue = { title: shortTitle(it.title) };
    render();
    $('#main').scrollTop = 0;
  }

  function wire() {
    // static icons
    $('#gsearchIc').innerHTML = I('search', 'sm');
    $('#menuBtn').innerHTML = I('menu');
    $('#drawerClose').innerHTML = I('x');
    $('#dropIcon').innerHTML = I('upload', 'lg');
    $('#chooseFilesLbl').innerHTML = `${I('file-spreadsheet', 'sm')}Choose files`;
    $('#chooseFolderLbl').innerHTML = `${I('folder-open', 'sm')}Choose a folder`;
    $('#openDash').innerHTML = `Open dashboard ${I('arrow-up-right', 'sm')}`;
    $('#backBtn').innerHTML = `${I('chevron-left', 'sm')}Overview`;
    setRefreshBtn(false);

    $('#refreshBtn').addEventListener('click', () => syncAll());
    $('#backBtn').addEventListener('click', goHome);
    $('#openDash').addEventListener('click', goHome);

    // Collapsible sidebar (remembered). Charts resize to the new width.
    const setNav = (collapsed) => {
      document.body.classList.toggle('nav-collapsed', collapsed);
      const b = $('#navToggle');
      b.innerHTML = I(collapsed ? 'panel-left-open' : 'panel-left-close');
      b.title = collapsed ? 'Expand menu' : 'Collapse menu';
      b.setAttribute('aria-label', b.title);
      setPref('sd-nav', collapsed ? 'collapsed' : 'open');
      setTimeout(() => charts.forEach((c) => { try { c.resize(); } catch (e) { /* ignore */ } }), 220);
    };
    const navPref = pref('sd-nav', null);
    setNav(navPref ? navPref === 'collapsed' : window.innerWidth < 1200 && window.innerWidth > 900);
    $('#navToggle').addEventListener('click', () => setNav(!document.body.classList.contains('nav-collapsed')));
    $('#menuBtn').addEventListener('click', openMobileNav);
    $('#scrim').addEventListener('click', () => { closeDrawer(); closeMobileNav(); });

    // Theme
    const paintTheme = () => {
      const dark = document.documentElement.dataset.theme === 'dark';
      $('#themeIcon').innerHTML = I(dark ? 'sun' : 'moon');
      $('#themeLabel').textContent = dark ? 'Light mode' : 'Dark mode';
    };
    paintTheme();
    $('#themeBtn').addEventListener('click', () => {
      const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
      document.documentElement.dataset.theme = next;
      setPref('sd-theme', next);
      paintTheme();
      render();
    });

    // Upload: click to choose, or drag files anywhere onto the upload screen.
    const drop = $('#drop');
    $('#fileInput').addEventListener('change', (e) => { uploadFiles(e.target.files); e.target.value = ''; });
    $('#folderInput').addEventListener('change', (e) => { uploadFiles(e.target.files); e.target.value = ''; });
    let replaceTarget = null;
    $('#replaceInput').addEventListener('change', (e) => { if (e.target.files.length) uploadFiles(e.target.files, replaceTarget); e.target.value = ''; });
    $('#startList').addEventListener('submit', (e) => {
      if (e.target.id !== 'unlockForm') return;
      e.preventDefault();
      const pw = $('#unlockPw').value;
      if (pw) unlock(pw);
    });
    $('#startList').addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      if (b.dataset.lock !== undefined) return lockAgain();
      if (b.dataset.replace) { replaceTarget = b.dataset.replace; $('#replaceInput').click(); return; }
      if (b.dataset.remove) { const src = state.sources.find((s) => s.id === b.dataset.remove); if (src) removeSource(src); }
    });
    const onStart = () => state.screen === 'start';
    ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); if (onStart()) drop.classList.add('over'); }));
    document.addEventListener('dragleave', (e) => { if (!e.relatedTarget) drop.classList.remove('over'); });
    document.addEventListener('drop', (e) => {
      e.preventDefault(); drop.classList.remove('over');
      if (!e.dataTransfer || !e.dataTransfer.files.length) return;
      if (!onStart()) goStart();
      uploadFiles(e.dataTransfer.files);
    });

    // Sidebar
    $('#side').addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      if (b.dataset.upload !== undefined) return goStart();
      if (b.dataset.home !== undefined) return goHome();
      if (b.dataset.fixall !== undefined) return downloadFixList(allTabs(), 'all files');
      if (b.dataset.toggle) {
        if (document.body.classList.contains('nav-collapsed') && !isMobile()) { goHome(); return; }
        state.open.has(b.dataset.toggle) ? state.open.delete(b.dataset.toggle) : state.open.add(b.dataset.toggle);
        return renderNav();
      }
      if (b.dataset.sheet) openTab(b.dataset.file, b.dataset.sheet);
    });

    // Global search
    const gs = $('#gsearch');
    let gt;
    let gk = -1;
    gs.addEventListener('input', () => { clearTimeout(gt); gk = -1; gt = setTimeout(() => globalSearch(gs.value), 180); });
    gs.addEventListener('keydown', (e) => {
      const items = $$('#gresults .gr-item');
      if (e.key === 'Escape') { $('#gresults').hidden = true; gs.blur(); return; }
      if (!items.length) return;
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        gk = (gk + (e.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length;
        items.forEach((x, k) => x.classList.toggle('hl', k === gk));
        items[gk].scrollIntoView({ block: 'nearest' });
      }
      if (e.key === 'Enter') { e.preventDefault(); pickHit(Number(items[Math.max(0, gk)].dataset.g)); }
    });
    $('#gresults').addEventListener('click', (e) => { const b = e.target.closest('[data-g]'); if (b) pickHit(Number(b.dataset.g)); });
    document.addEventListener('click', (e) => { if (!e.target.closest('.gsearch')) $('#gresults').hidden = true; });

    // Page tabs
    $('#tabs').addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      state.view = b.dataset.view;
      if (state.view !== 'data') state.fromIssue = null;
      render();
    });

    // Filters (shared)
    $('#filters').addEventListener('change', (e) => {
      if (e.target.id === 'issueSel') { state.issue = e.target.value; state.fromIssue = null; }
      else if (e.target.dataset.filter !== undefined) { if (e.target.value) state.filters[e.target.dataset.filter] = e.target.value; else delete state.filters[e.target.dataset.filter]; }
      else return;
      state.page = 0; state.cardLimit = CARD_STEP; rerenderBody();
    });
    $('#filters').addEventListener('input', (e) => {
      if (e.target.id !== 'searchBox') return;
      state.search = e.target.value; state.page = 0; state.cardLimit = CARD_STEP;
      clearTimeout(wire._t);
      wire._t = setTimeout(rerenderBody, 200);
    });

    // Clicks inside the main area (delegated)
    $('#main').addEventListener('click', (e) => {
      const b = e.target.closest('button, [data-file]');
      if (!b) return;
      const d = b.dataset;
      const sel = findSel();
      if (d.upload !== undefined) return goStart();
      if (d.refresh !== undefined) return syncAll();
      if (d.fixall !== undefined) return downloadFixList(allTabs(), 'all files');
      if (d.fixtab !== undefined && sel) return downloadFixList([{ file: sel.file, sheet: sel.sheet }], sel.sheet.title);
      if (d.inspect) { const [f, s, i] = d.inspect.split('|'); return inspectIssue(f, s, Number(i)); }
      if (d.inspectHere !== undefined && sel) return inspectIssue(sel.file.id, sel.sheet.name, Number(d.inspectHere));
      if (d.backissues !== undefined) { state.view = 'health'; state.fromIssue = null; return render(); }
      if (d.clearall !== undefined) { resetFilters(); return render(); }
      if (d.unfilter) {
        if (d.unfilter === 'search') state.search = '';
        else if (d.unfilter === 'issue') { state.issue = ''; state.fromIssue = null; }
        else delete state.filters[d.unfilter.slice(2)];
        state.page = 0; return render();
      }
      if (d.rowsview) return setRowsView(d.rowsview);
      if (d.more !== undefined) { state.cardLimit += CARD_STEP; return rerenderBody(); }
      if (d.details && sel) return openDrawer(sel.sheet, Number(d.details));
      if (d.intable && sel) { state.locate = Number(d.intable); return setRowsView('table'); }
      if (d.sort !== undefined) {
        const ci = Number(d.sort);
        state.sort = state.sort && state.sort.ci === ci ? (state.sort.dir > 0 ? { ci, dir: -1 } : null) : { ci, dir: 1 };
        return rerenderBody();
      }
      if (d.page) { state.page += Number(d.page); rerenderBody(); const w = $('.table-wrap'); if (w) w.scrollTop = 0; return; }
      if (d.file && d.sheet) return openTab(d.file, d.sheet);
    });
    $('#view-dashboard').addEventListener('change', (e) => {
      if (e.target.id !== 'measureSel') return;
      state.measure = e.target.value === '' ? null : Number(e.target.value);
      const sel = findSel(); if (sel) renderDashboard(sel.sheet);
    });

    // Drawer
    $('#drawer').addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      if (b.id === 'drawerClose' || b.dataset.closedrawer !== undefined) return closeDrawer();
      if (b.dataset.intable) { closeDrawer(); state.view = 'data'; state.locate = Number(b.dataset.intable); setRowsView('table'); }
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { closeDrawer(); closeMobileNav(); }
      if (e.key === '/' && !/input|select|textarea/i.test(document.activeElement.tagName)) { e.preventDefault(); $('#gsearch').focus(); }
    });

    // Keep "Updated x min ago" fresh.
    setInterval(updateStatus, 30000);
  }

  async function init() {
    const t = pref('sd-theme', null); if (t) document.documentElement.dataset.theme = t;
    wire();
    await loadInfo();
    await loadSourceList();
    render();
    // Show the last copy instantly (works offline), then get fresh data.
    for (const rec of (await store.all()) || []) addRecord(rec, { save: false });
    if (state.files.length) render();
    await syncAll();
    setInterval(() => { if (!document.hidden) syncAll(); }, REFRESH_MS);
    document.addEventListener('visibilitychange', () => { if (!document.hidden && state.lastSync && Date.now() - state.lastSync > REFRESH_MS) syncAll(); });
  }

  init();
})();
