// app.js — the screens: Home (all tabs at a glance), and for each tab: Summary, Problems, Table.
// Data is only ever READ. Nothing is written back to Google Drive or the Excel files.
(function () {
  'use strict';
  const $ = (s, el = document) => el.querySelector(s);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const REFRESH_MS = 5 * 60 * 1000;

  const state = {
    sources: [],        // { id, name, kind, file?, status, error }
    files: [],          // { id, name, addedAt, sheets: [model] }
    screen: 'start',    // 'start' = upload screen (shown first), 'app' = dashboard
    sel: null,          // { fileId, sheet } — null = Home
    open: new Set(),    // files expanded in the sidebar
    view: 'dashboard',
    filters: {}, search: '', measure: undefined, focus: null, sort: null, page: 0,
    lastSync: null,
  };
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
    : typeof v === 'number' ? (!col || col.type === 'number' ? nf.format(v) : String(v)) : String(v));
  const totalLabel = (n) => (/^(total|net|tot)\b/i.test(n) ? n : `Total ${n}`);
  const plural = (n, one, many) => `${nf.format(n)} ${n === 1 ? one : many || one + 's'}`;
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
    let rowsA = 0, rowsB = 0, fixed = 0, added = 0, probA = 0, probB = 0;
    for (const k in A) { rowsA += A[k].rows; for (const t in A[k].cells) probA += A[k].cells[t]; }
    for (const k in B) { rowsB += B[k].rows; for (const t in B[k].cells) probB += B[k].cells[t]; }
    for (const k of new Set([...Object.keys(A), ...Object.keys(B)])) {
      const a = (A[k] && A[k].cells) || {}, b = (B[k] && B[k].cells) || {};
      for (const t of new Set([...Object.keys(a), ...Object.keys(b)])) {
        const d = (b[t] || 0) - (a[t] || 0);
        if (d < 0) fixed -= d; else added += d;
      }
    }
    return { since: p.prev.at, rows: rowsB - rowsA, fixed, added, probA, probB, healthA: fileHealth(A), healthB: fileHealth(B) };
  }
  function changeLine(c) {
    if (!c) return '';
    const parts = [];
    if (c.rows) parts.push(`${c.rows > 0 ? '+' : ''}${nf.format(c.rows)} rows`);
    if (c.fixed) parts.push(`<span class="up">✓ ${plural(c.fixed, 'problem')} fixed</span>`);
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
        items.push([pri, file.name, sheet.name, r, isRow ? '—' : col.letter, isRow ? '' : (col.raw || '(no heading)').trim(), it.title.replace(/ \(column [A-Z]+\)/, ''), String(now), sug || '', it.fix]);
      }
    }
    return items;
  }
  function nameItems(file, sheet) {
    return (sheet.names || []).map((n) => [file.name, sheet.name, n.where, n.current, n.suggested, n.why]);
  }
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
      ['How to use this fix list'],
      [''],
      ['1. Open the "Fix list" sheet. Each line is one cell to look at in your Google Sheet.'],
      ['2. Go to the File + Tab, find the Row and Column, and compare with "What\'s there now".'],
      ['3. Make the change in Google Sheets (this app never changes your sheets).'],
      ['4. Type ✓ in the "Done" column so you know where you stopped.'],
      ['5. Start with "1 · Fix" lines — they make totals wrong. "2 · Check" lines may be fine.'],
      ['6. The "Clearer names" sheet suggests better tab and column names. Optional, but makes sheets easier to read.'],
      ['7. Next time, download the sheet again and upload it — the app shows how many problems you fixed.'],
      [''],
      [`Made on ${new Date().toLocaleString('en-IN')} · ${rows.length} items to look at · ${names.length} name suggestions`],
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
    try { XLSX.writeFile(wb, `Fix list - ${safe} - ${stamp}.xlsx`); toast(`✓ Fix list downloaded (${rows.length} items)`); }
    catch (e) { console.error(e); toast('Could not create the file.'); }
  }

  function addRecord(rec, { save = true } = {}) {
    let sheets;
    try { sheets = parseWorkbook(rec.buf); } catch (e) { console.error(e); return false; }
    if (!sheets.length) return false;
    if (save) { store.put(rec); recordVersion(rec.id, rec.buf, sheets); }
    const file = { id: rec.id, name: rec.name, addedAt: rec.addedAt, sheets };
    const i = state.files.findIndex((f) => f.id === rec.id);
    if (i >= 0) state.files[i] = file; else state.files.push(file);
    const order = state.sources.map((s) => s.id);
    state.files.sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id));
    if (state.sel && !findSel()) state.sel = null; // that tab disappeared -> Home
    return true;
  }

  // ---------- fetching (read-only) ----------
  async function loadSourceList() {
    let list;
    try { list = await (await fetch('/api/sources')).json(); } catch (e) { return; }
    const old = new Map(state.sources.map((s) => [s.id, s]));
    state.sources = list.map((s) => ({ status: 'loading', error: '', ...(old.get(s.id) || {}), ...s }));
    const ids = new Set(list.map((s) => s.id));
    state.files = state.files.filter((f) => ids.has(f.id));
    for (const rec of (await store.all()) || []) if (!ids.has(rec.id)) store.del(rec.id);
    let gone = false;
    for (const id of Object.keys(progress)) if (!ids.has(id)) { delete progress[id]; gone = true; }
    if (gone) saveProgress(progress);
  }

  async function fetchSource(src) {
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

  async function syncAll() {
    if (syncing) return;
    syncing = true;
    const btn = $('#refreshBtn');
    btn.disabled = true; btn.textContent = 'Refreshing…';
    await loadSourceList();
    await Promise.all(state.sources.map(fetchSource));
    syncing = false;
    btn.disabled = false; btn.textContent = '↻ Refresh';
    state.lastSync = Date.now();
    // Don't redraw while she is typing in the search box.
    if (document.activeElement && document.activeElement.id === 'searchBox') { renderNav(); updateStatus(); return; }
    render();
  }

  function updateStatus() {
    const bad = state.sources.filter((s) => s.status === 'error').length;
    $('#syncStatus').innerHTML = state.lastSync
      ? `Updated ${esc(ago(state.lastSync))}${bad ? ` · <span style="color:var(--red)">${bad} file${bad > 1 ? 's' : ''} not available</span>` : ''}<br>Refreshes by itself every 5 min`
      : 'Getting the latest data…';
  }

  // ---------- selection ----------
  function findSel() {
    if (!state.sel) return null;
    const file = state.files.find((f) => f.id === state.sel.fileId);
    const sheet = file && file.sheets.find((s) => s.name === state.sel.sheet);
    return sheet ? { file, sheet } : null;
  }
  function openTab(fileId, sheetName, view) {
    state.screen = 'app';
    state.sel = { fileId, sheet: sheetName };
    state.open.add(fileId);
    state.filters = {}; state.search = ''; state.focus = null; state.sort = null; state.page = 0; state.measure = undefined;
    state.view = view || 'dashboard';
    render();
    $('#main').scrollTop = 0;
  }
  function goHome() { state.screen = 'app'; state.sel = null; render(); $('#main').scrollTop = 0; }
  function goStart() { state.screen = 'start'; render(); $('#main').scrollTop = 0; }

  // ---------- upload (saved into the app's data folder on this PC) ----------
  async function uploadFiles(list) {
    const files = [...list].filter((f) => /\.(xlsx|xlsm|xls|csv)$/i.test(f.name));
    if (!files.length) { toast('Please choose Excel files (.xlsx, .xls or .csv).'); return; }
    const drop = $('#drop');
    drop.classList.add('busy');
    $('.drop-title').textContent = `Uploading ${plural(files.length, 'file')}…`;
    const failed = [];
    for (const f of files) {
      try {
        const r = await fetch('/api/upload?name=' + encodeURIComponent(f.name), { method: 'POST', headers: { 'X-App': 'sheets-dashboard' }, body: f });
        if (!r.ok) { const j = await r.json().catch(() => ({})); failed.push(`${f.name}: ${j.error || 'upload failed'}`); }
      } catch (e) { failed.push(`${f.name}: upload failed`); }
    }
    drop.classList.remove('busy');
    $('.drop-title').textContent = 'Drag your Excel files here';
    await syncAll();
    const ok = files.length - failed.length;
    if (failed.length) toast(failed.join(' · '), 7000);
    else toast(`✓ ${plural(ok, 'file')} uploaded`);
  }

  function renderStart() {
    const rows = state.sources.map((src) => {
      const f = state.files.find((x) => x.id === src.id);
      const icon = src.kind === 'local' ? '📄' : '☁';
      const ch = f ? changes(f.id) : null;
      const p = f && progress[f.id] && progress[f.id].cur;
      const meta = f ? `${plural(f.sheets.length, 'tab')} · updated ${ago(f.addedAt)}${p ? ` · health ${fileHealth(p.tabs)}` : ''}${ch && ch.fixed ? ` · ✓ ${plural(ch.fixed, 'problem')} fixed since last upload` : ''}`
        : src.status === 'error' ? src.error : 'Loading…';
      return `<div class="card file-row ${!f && src.status === 'error' ? 'err' : ''}">
        <span class="ficon">${icon}</span>
        <div class="grow"><div class="fname">${esc(f ? f.name : src.name)}</div><div class="fmeta">${esc(meta)}</div></div>
        ${f ? '<span class="chip green" style="margin:0">Ready</span>' : ''}
      </div>`;
    }).join('');
    $('#startList').innerHTML = state.sources.length ? `<h3>Files in the dashboard</h3>${rows}` : '';
    $('#openDash').disabled = !state.files.length;
  }

  // ---------- sidebar ----------
  function renderNav() {
    const sel = state.sel;
    const onStart = state.screen === 'start';
    let html = `<button class="nav-item nav-upload ${onStart ? 'active' : ''}" data-upload type="button">⤒ <span class="label">Upload files</span></button>
      <button class="nav-item nav-home ${!onStart && !sel ? 'active' : ''}" data-home type="button">🏠 <span class="label">Home</span></button>`;
    for (const src of state.sources) {
      const f = state.files.find((x) => x.id === src.id);
      const title = f ? f.name : src.name;
      const isOpen = state.open.has(src.id);
      html += `<button class="nav-item nav-file ${isOpen ? 'open' : ''}" data-toggle="${esc(src.id)}" type="button">
        <span class="chev">▶</span><span class="label" title="${esc(title)}">${esc(title)}</span></button>`;
      if (!f && src.status === 'error') html += `<div class="nav-error">${esc(src.error)}</div>`;
      if (f && isOpen) {
        html += '<div class="nav-tabs">' + f.sheets.map((s) => `
          <button class="nav-item ${!onStart && sel && sel.fileId === f.id && sel.sheet === s.name ? 'active' : ''}" data-file="${esc(f.id)}" data-sheet="${esc(s.name)}" type="button" title="${esc(statusText(s))}">
            <span class="dot ${statusColor(s)}"></span><span class="label">${esc(s.title)}</span>
          </button>`).join('') + '</div>';
      }
    }
    $('#nav').innerHTML = html;
  }

  // ---------- headline number for a tab (used on Home tiles) ----------
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

  // ---------- HOME ----------
  function renderHome() {
    const el = $('#view-home');
    const allSheets = state.files.flatMap((f) => f.sheets);
    if (!allSheets.length) {
      const failed = state.sources.length && state.sources.every((s) => s.status === 'error');
      el.innerHTML = failed
        ? `<div class="empty-state"><h2>Couldn't open the sheets</h2><p>${esc(state.sources[0].error)}</p><p>Press ↻ Refresh to try again.</p></div>`
        : '<div class="empty-state"><div class="spinner"></div><h2>Getting the latest data…</h2><p>This takes a few seconds.</p></div>';
      return;
    }
    const toFix = allSheets.reduce((s, x) => s + errCount(x), 0);
    const toCheck = allSheets.reduce((s, x) => s + warnCount(x), 0);
    const bad = state.sources.filter((s) => s.status === 'error');
    let html = bad.map((s) => `<div class="notice">⚠ <b>${esc(s.name)}</b> could not be opened: ${esc(s.error)}</div>`).join('');
    const all = state.files.flatMap((f) => f.sheets.map((sheet) => ({ file: f, sheet })));
    const health = fileHealth(Object.fromEntries(all.map((x, k) => [k, { rows: x.sheet.rows.length, score: x.sheet.score }])));
    const nItems = countFixItems(all);
    html += `
      <div class="grid strip">
        <div class="card"><div class="big" style="color:${health >= 90 ? 'var(--green)' : health >= 75 ? 'var(--amber)' : 'var(--red)'}">${health}<span style="font-size:15px;color:var(--muted)"> / 100</span></div><div class="lbl">data health score</div></div>
        <div class="card"><div class="big" style="color:${toFix ? 'var(--red)' : 'var(--green)'}">${toFix}</div><div class="lbl">problems to fix</div></div>
        <div class="card"><div class="big" style="color:${toCheck ? 'var(--amber)' : 'var(--green)'}">${toCheck}</div><div class="lbl">things to double-check</div></div>
        <div class="card coach">
          <div class="lbl" style="margin-bottom:8px">Want to clean up the sheets?</div>
          <button class="btn primary" id="fixAll" type="button" ${nItems ? '' : 'disabled'}>⬇ Download fix list</button>
          <div class="lbl" style="font-size:12.5px;margin-top:6px">${nItems ? `${nf.format(nItems)} cells to look at, with suggested fixes` : 'Nothing to fix 🎉'}</div>
        </div>
      </div>`;
    for (const f of state.files) {
      const c = changes(f.id);
      html += `<div class="section-title">${esc(f.name)}</div>
        <div class="progress-line">${c ? `Since the last upload (${esc(ago(c.since))}): ${changeLine(c)}` : 'First upload of this file — next time you upload a newer copy, you\'ll see here how many problems were fixed.'}</div>
        <div class="grid tiles">`;
      for (const s of f.sheets) {
        const h = headline(s);
        const c = statusColor(s);
        html += `
          <button class="card tile" data-file="${esc(f.id)}" data-sheet="${esc(s.name)}" type="button">
            <div class="tile-name" title="${esc(s.title)}">${esc(s.title)}</div>
            <div class="tile-num">${esc(h.num)}</div>
            <div class="tile-lbl" title="${esc(h.lbl)}">${esc(h.lbl)}</div>
            <span class="chip ${c}"><span class="dot ${c}"></span>${esc(statusText(s))}</span>
          </button>`;
      }
      html += '</div>';
    }
    el.innerHTML = html;
  }

  // ---------- filters ----------
  function categoryCols(sheet) {
    const pri = { status: 0, person: 1, customer: 2, product: 3, group: 4 };
    return sheet.cols.map((c, i) => ({ c, i }))
      .filter((x) => (x.c.type === 'category' || (x.c.type === 'text' && pri[x.c.role] !== undefined)) && x.c.fill >= 0.3)
      .sort((a, b) => (pri[a.c.role] ?? 9) - (pri[b.c.role] ?? 9) || (a.c.type === 'category' ? 0 : 1) - (b.c.type === 'category' ? 0 : 1) || b.c.fill - a.c.fill);
  }
  function visibleRows(sheet) {
    const q = state.search.trim().toLowerCase();
    return sheet.rows.filter((r) => {
      if (state.focus && !state.focus.rows.has(r.r)) return false;
      for (const k in state.filters) { const v = r.v[k]; if (v === null || Analyze.normKey(v) !== state.filters[k]) return false; }
      if (q && !r.v.some((v) => v !== null && fmtCell(v).toLowerCase().includes(q))) return false;
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
  function renderFilters(sheet) {
    const el = $('#filters');
    const fcols = categoryCols(sheet).filter((x) => groupBy(sheet.rows, x.i, null).length <= 60).slice(0, 3);
    el.innerHTML = `
      <span class="lbl">Show only</span>
      ${fcols.map(({ c, i }) => {
        const opts = groupBy(sheet.rows, i, null).sort((a, b) => a.label.localeCompare(b.label));
        return `<select data-filter="${i}" title="${esc(c.name)}">
          <option value="">All ${esc(c.name)}</option>
          ${opts.map((o) => `<option value="${esc(o.key)}" ${state.filters[i] === o.key ? 'selected' : ''}>${esc(o.label.slice(0, 50))}</option>`).join('')}
        </select>`;
      }).join('')}
      <input type="search" id="searchBox" placeholder="🔍 Search…" value="${esc(state.search)}">
      ${Object.keys(state.filters).length || state.search ? '<button class="link" id="clearFilters" type="button">Clear</button>' : ''}`;
    el.hidden = false;
  }

  // ---------- charts ----------
  function palette() {
    const cs = getComputedStyle(document.documentElement);
    const dark = document.documentElement.dataset.theme === 'dark';
    return {
      accent: cs.getPropertyValue('--accent').trim(),
      green: cs.getPropertyValue('--green').trim(),
      text: cs.getPropertyValue('--muted').trim(),
      grid: cs.getPropertyValue('--line').trim(),
      series: dark ? ['#6b9bf2', '#3cc1a8', '#e6b04e', '#a98af0'] : ['#2f6fde', '#16a394', '#e08a2e', '#8a5cd6'],
    };
  }
  function destroyCharts() { charts.forEach((c) => { try { c.destroy(); } catch (e) { /* ignore */ } }); charts = []; }
  function makeChart(canvas, cfg) {
    if (!canvas || typeof Chart === 'undefined') return; // charts must never break the page
    try { charts.push(new Chart(canvas, cfg)); } catch (e) { console.warn('chart failed', e); }
  }
  function barCfg(labels, data, color, horizontal, col) {
    const p = palette();
    const fmt = (v) => (col ? money(col, v) : nf.format(v));
    const trim = (l) => (l.length > 26 ? l.slice(0, 25) + '…' : l);
    return {
      type: 'bar',
      data: { labels, datasets: [{ data, backgroundColor: color, borderRadius: 5, maxBarThickness: 32 }] },
      options: {
        indexAxis: horizontal ? 'y' : 'x', maintainAspectRatio: false, animation: { duration: 200 },
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: (ctx) => ' ' + (col ? money(col, ctx.raw) : nf.format(ctx.raw) + ' rows') } } },
        scales: horizontal ? {
          x: { grid: { color: p.grid }, ticks: { color: p.text, maxRotation: 0, callback: (v) => fmt(v) } },
          y: { grid: { color: 'transparent' }, ticks: { color: p.text, callback: function (v) { return trim(this.getLabelForValue(v)); } } },
        } : {
          x: { grid: { color: 'transparent' }, ticks: { color: p.text, maxRotation: 0, autoSkip: true } },
          y: { grid: { color: p.grid }, ticks: { color: p.text, callback: (v) => fmt(v) } },
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
      <div class="grid kpis">${cards.map((x) => {
        const s = personStats(x, T.months);
        const c = pctColor(s.pct);
        return `<div class="card">
          <div class="kpi-lbl">${esc(x.isTotal ? 'Whole team' : x.name)}${x.role && !x.isTotal ? ` · ${esc(x.role)}` : ''}</div>
          <div class="kpi-num">${Math.round(s.pct * 100)}% <span class="chip ${c}" style="margin:0">${s.pct >= 1 ? 'On target' : 'Behind'}</span></div>
          <div class="kpi-sub">${n2(s.a)} of ${n2(s.t)} ${esc(unit)} so far${s.lastM !== null ? ` (Jan–${MONTHS[s.lastM]})` : ''}</div>
          <div class="kpi-sub">Full year target: ${n2(s.year)} ${esc(unit)}</div>
        </div>`;
      }).join('')}</div>
      <div class="grid charts">
        <div class="card wide"><h3>Target vs achieved, month by month <small>${esc(total.isTotal ? 'whole team' : total.name)}${unit ? ' · ' + esc(unit) : ''}</small></h3><div class="chart-box"><canvas id="tg-total"></canvas></div></div>
        ${people.length > 1 ? people.map((x, k) => `<div class="card"><h3>${esc(x.name)} <small>${esc(x.role)}</small></h3><div class="chart-box"><canvas id="tg-p${k}"></canvas></div></div>`).join('') : ''}
      </div>`;
    const cfg = (x) => ({
      type: 'bar',
      data: {
        labels: T.months.map((m) => MONTHS[m]),
        datasets: [
          { label: 'Target', data: T.months.map((m) => x.target[m] ?? null), backgroundColor: p.grid, borderRadius: 5, maxBarThickness: 24 },
          { label: 'Achieved', data: T.months.map((m) => x.ach[m] ?? null), borderRadius: 5, maxBarThickness: 24,
            backgroundColor: T.months.map((m) => (x.ach[m] !== undefined && x.ach[m] >= (x.target[m] || 0) ? p.green : p.accent)) },
        ],
      },
      options: {
        maintainAspectRatio: false, animation: { duration: 200 },
        plugins: {
          legend: { labels: { color: p.text, boxWidth: 12 } },
          tooltip: { callbacks: { afterBody: (items) => { const m = T.months[items[0].dataIndex]; const t = x.target[m], a = x.ach[m]; return t && a !== undefined ? `${Math.round((a / t) * 100)}% of target` : ''; } } },
        },
        scales: { x: { grid: { color: 'transparent' }, ticks: { color: p.text } }, y: { grid: { color: p.grid }, ticks: { color: p.text } } },
      },
    });
    makeChart($('#tg-total'), cfg(total));
    people.forEach((x, k) => { if (people.length > 1) makeChart($('#tg-p' + k), cfg(x)); });
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

    const kpis = [{ lbl: 'Rows', num: nf.format(rows.length), sub: rows.length !== sheet.rows.length ? `out of ${nf.format(sheet.rows.length)}` : 'entries in this tab' }];
    for (const { c, i } of measureOptions(sheet).slice(0, 2)) {
      let sum = 0;
      for (const r of rows) if (typeof r.v[i] === 'number') sum += r.v[i];
      kpis.push({ lbl: totalLabel(c.name), num: money(c, sum), sub: isMoney(c) ? nf.format(Math.round(sum)) : 'added up' });
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
          lbl: `Latest month (${MONTHS[last % 12]} ${Math.floor(last / 12)})`,
          num: mCol ? money(mCol, cur) : plural(cur, 'row'),
          subHtml: ch === null ? 'no data the month before' : `<span class="${ch >= 0 ? 'up' : 'down'}">${ch >= 0 ? '▲' : '▼'} ${Math.abs(Math.round(ch * 100))}%</span> vs ${MONTHS[(last - 1) % 12]}`,
        });
      }
    }
    const cust = sheet.cols.findIndex((c) => c.role === 'customer');
    if (cust >= 0 && kpis.length < 5) kpis.push({ lbl: sheet.cols[cust].name, num: nf.format(groupBy(rows, cust, null).length), sub: 'different names' });

    const breakdowns = categoryCols(sheet)
      .filter((x, k, arr) => !x.c.role || arr.findIndex((y) => y.c.role === x.c.role) === k) // one chart per kind
      .filter((x) => groupBy(sheet.rows, x.i, null).length > 1)
      .slice(0, 4);
    const mOpts = measureOptions(sheet);
    const what = mCol ? mCol.name : 'Number of rows';

    el.innerHTML = `
      <div class="grid kpis">${kpis.map((k) => `
        <div class="card"><div class="kpi-lbl" title="${esc(k.lbl)}">${esc(k.lbl)}</div><div class="kpi-num">${esc(k.num)}</div><div class="kpi-sub">${k.subHtml || esc(k.sub)}</div></div>`).join('')}
      </div>
      ${mOpts.length ? `<div class="measure">Charts are showing
        <select id="measureSel">
          ${mOpts.map(({ c, i }) => `<option value="${i}" ${meas === i ? 'selected' : ''}>${esc(c.name)}</option>`).join('')}
          <option value="" ${meas === null ? 'selected' : ''}>Number of rows</option>
        </select></div>` : ''}
      <div class="grid charts">
        ${byMonth && byMonth.size ? `<div class="card wide"><h3>${esc(what)} by month <small>based on ${esc(dateCol.c.name)}</small></h3><div class="chart-box"><canvas id="ch-trend"></canvas></div></div>` : ''}
        ${breakdowns.map(({ c }, k) => `<div class="card"><h3>Top 10 by ${esc(c.name)}</h3><div class="chart-box"><canvas id="ch-b${k}"></canvas></div></div>`).join('')}
        ${!(byMonth && byMonth.size) && !breakdowns.length ? '<div class="card wide muted">This tab has no columns suitable for charts. Open the <b>Table</b> to see everything.</div>' : ''}
      </div>`;

    if (byMonth && byMonth.size) {
      const keys = [...byMonth.keys()].sort((a, b) => a - b);
      const end = keys[keys.length - 1], start = Math.max(keys[0], end - 35);
      const labels = [], data = [];
      for (let k = start; k <= end; k++) { labels.push(`${MONTHS[k % 12]} ${String(Math.floor(k / 12)).slice(2)}`); data.push(byMonth.get(k) || 0); }
      makeChart($('#ch-trend'), barCfg(labels, data, p.accent, false, mCol));
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
  function rowsText(rows) { return rows.length > 15 ? rows.slice(0, 15).join(', ') + ` … and ${rows.length - 15} more` : rows.join(', '); }
  function problemCard(it, idx, color) {
    return `
      <div class="card problem ${color}">
        <div class="problem-body">
          <div class="problem-title">${esc(it.title)}</div>
          <div class="problem-detail">${esc(it.detail)}</div>
          ${it.fix ? `<div class="problem-fix"><b>How to fix:</b> ${esc(it.fix)}</div>` : ''}
          ${it.rows.length ? `<div class="problem-rows">Row ${esc(rowsText(it.rows))}</div>` : ''}
        </div>
        ${it.cells.length ? `<button class="btn small" data-focus="${idx}" type="button">Show me →</button>` : ''}
      </div>`;
  }
  function renderHealth(sheet) {
    const el = $('#view-health');
    const list = sheet.issues.map((it, idx) => ({ it, idx }));
    const errs = list.filter((x) => x.it.sev === 'err'), warns = list.filter((x) => x.it.sev === 'warn'), infos = list.filter((x) => x.it.sev === 'info');
    const c = errs.length ? 'red' : warns.length ? 'amber' : 'green';
    const title = errs.length ? `${plural(errs.length, 'problem')} to fix` : warns.length ? 'Looks fine — a few things to double-check' : 'Everything looks good';
    const sel = findSel();
    const nItems = fixItems(sel.file, sheet).length;
    const ch = changes(sel.file.id, sheet.name);
    let html = `
      <div class="verdict">
        <div class="verdict-icon ${c}">${errs.length ? '!' : warns.length ? '?' : '✓'}</div>
        <div style="flex:1">
          <div class="verdict-title">${esc(title)}</div>
          <div class="muted">We checked ${plural(sheet.rows.length, 'row')}. Row numbers are the same as in the sheet, so you can go straight to them.</div>
        </div>
        ${nItems || (sheet.names || []).length ? `<button class="btn" id="fixTab" type="button">⬇ Fix list for this tab</button>` : ''}
      </div>
      ${ch ? `<div class="change-banner">Since the last upload (${esc(ago(ch.since))}): ${changeLine(ch)}</div>` : ''}`;
    if (errs.length) html += `<div class="problem-group"><span class="dot red"></span>Fix these — they make totals wrong</div>` + errs.map((x) => problemCard(x.it, x.idx, 'red')).join('');
    if (warns.length) html += `<div class="problem-group"><span class="dot amber"></span>Double-check these</div>` + warns.map((x) => problemCard(x.it, x.idx, 'amber')).join('');
    const names = sheet.names || [];
    if (names.length) {
      html += `<div class="problem-group">✎ Clearer names <span class="muted" style="font-weight:400;font-size:13.5px">— optional: rename these in the sheet so everyone understands them</span></div>
        <div class="card" style="padding:4px 8px"><table class="names-table">
          <thead><tr><th>Where</th><th>Now</th><th>Better</th><th>Why</th></tr></thead>
          <tbody>${names.map((n) => `<tr><td class="why">${esc(n.where)}</td><td class="old">${esc(n.current)}</td><td class="new">${esc(n.suggested)}</td><td class="why">${esc(n.why)}</td></tr>`).join('')}</tbody>
        </table></div>`;
    }
    if (infos.length) html += `<details class="tidy"><summary>Tidy-up suggestions (${infos.length}) — optional, nothing is wrong</summary>${infos.map((x) => problemCard(x.it, x.idx, '')).join('')}</details>`;
    el.innerHTML = html;
  }

  // ---------- TABLE ----------
  const PAGE = 100;
  function renderData(sheet) {
    const el = $('#view-data');
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
    const pages = Math.max(1, Math.ceil(rows.length / PAGE));
    state.page = Math.min(state.page, pages - 1);
    const slice = rows.slice(state.page * PAGE, state.page * PAGE + PAGE);
    const ci = sheet.cellIssues;
    // In "Show me" mode, put the problem column(s) first so they are visible without scrolling.
    const focusCols = state.focus && state.focus.cols ? state.focus.cols : [];
    const order = [...focusCols, ...sheet.cols.map((_, i) => i).filter((i) => !focusCols.includes(i))];
    el.innerHTML = `
      ${state.focus ? `<div class="banner">Showing only the rows for: <b>${esc(state.focus.title)}</b><button class="btn small" id="clearFocus" type="button">Show all rows</button></div>` : ''}
      <div class="legend"><span><i style="background:var(--red-soft);border:1px solid var(--red)"></i>needs fixing</span><span><i style="background:var(--amber-soft);border:1px solid var(--amber)"></i>double-check</span><span>Click a heading to sort.</span></div>
      <div class="table-wrap"><table>
        <thead><tr><th class="rownum" title="Row number in the sheet">Row</th>${order.map((i) => { const c = sheet.cols[i]; return `<th data-sort="${i}" title="Column ${c.letter}">${esc(c.name)} ${state.sort && state.sort.ci === i ? `<span class="arrow">${state.sort.dir > 0 ? '▲' : '▼'}</span>` : ''}</th>`; }).join('')}</tr></thead>
        <tbody>${slice.map((r) => `<tr><td class="rownum">${r.r}</td>${order.map((i) => {
          const v = r.v[i];
          const iss = ci[r.r + ':' + i];
          const shown = iss && v === null && r.raw[i] != null ? String(r.raw[i]) : fmtCell(v, sheet.cols[i]);
          const cls = (typeof v === 'number' && sheet.cols[i].type === 'number' ? 'num ' : '') + (iss && iss.sev !== 'info' ? 'c-' + iss.sev : '');
          return `<td class="${cls}" title="${esc(iss && iss.sev !== 'info' ? iss.title + ' — ' : '')}${esc(shown)}">${esc(shown)}</td>`;
        }).join('')}</tr>`).join('') || `<tr><td class="rownum"></td><td colspan="${sheet.cols.length}" class="muted">Nothing matches.</td></tr>`}</tbody>
      </table></div>
      <div class="pager">${plural(rows.length, 'row')}
        ${pages > 1 ? `<button class="btn small" data-page="-1" ${state.page === 0 ? 'disabled' : ''} type="button">‹ Previous</button> Page ${state.page + 1} of ${pages}
        <button class="btn small" data-page="1" ${state.page >= pages - 1 ? 'disabled' : ''} type="button">Next ›</button>` : ''}
      </div>`;
  }

  // ---------- main render ----------
  function render() {
    renderNav();
    updateStatus();
    const start = state.screen === 'start';
    $('#view-start').hidden = !start;
    if (start) {
      destroyCharts();
      ['home', 'dashboard', 'health', 'data'].forEach((v) => ($('#view-' + v).hidden = true));
      $('#tabs').hidden = true; $('#backBtn').hidden = true; $('#filters').hidden = true;
      $('#title').textContent = 'Upload your files';
      $('#subtitle').textContent = 'Add the latest Excel files, then open the dashboard.';
      renderStart();
      return;
    }
    const sel = findSel();
    const home = !sel;
    $('#view-home').hidden = !home;
    $('#tabs').hidden = home;
    $('#backBtn').hidden = home;
    ['dashboard', 'health', 'data'].forEach((v) => ($('#view-' + v).hidden = home || state.view !== v));
    if (home) {
      destroyCharts();
      $('#filters').hidden = true;
      $('#title').textContent = 'Overview';
      $('#subtitle').textContent = 'Every tab at a glance. Click a box to open it.';
      renderHome();
      return;
    }
    const { file, sheet } = sel;
    $('#title').textContent = sheet.title;
    $('#subtitle').textContent = `${file.name} · ${plural(sheet.rows.length, 'row')}`;
    const cnt = $('#issueCount');
    const e = errCount(sheet), w = warnCount(sheet);
    cnt.textContent = e || w || '✓';
    cnt.className = 'count ' + (e ? 'red' : w ? 'amber' : '');
    document.querySelectorAll('#tabs button').forEach((b) => b.classList.toggle('active', b.dataset.view === state.view));
    if (state.view === 'health') { $('#filters').hidden = true; renderHealth(sheet); }
    else if (state.view === 'dashboard' && sheet.targets) { $('#filters').hidden = true; renderDashboard(sheet); }
    else { renderFilters(sheet); state.view === 'dashboard' ? renderDashboard(sheet) : renderData(sheet); }
  }

  // ---------- events ----------
  function wire() {
    $('#refreshBtn').addEventListener('click', () => syncAll());
    $('#backBtn').addEventListener('click', goHome);
    $('#openDash').addEventListener('click', goHome);

    // Upload: click to choose, or drag files anywhere onto the upload screen.
    const drop = $('#drop');
    $('#fileInput').addEventListener('change', (e) => { uploadFiles(e.target.files); e.target.value = ''; });
    const onStart = () => state.screen === 'start';
    ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); if (onStart()) drop.classList.add('over'); }));
    document.addEventListener('dragleave', (e) => { if (!e.relatedTarget) drop.classList.remove('over'); });
    document.addEventListener('drop', (e) => {
      e.preventDefault(); drop.classList.remove('over');
      if (!e.dataTransfer || !e.dataTransfer.files.length) return;
      if (!onStart()) goStart();
      uploadFiles(e.dataTransfer.files);
    });

    $('#nav').addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      if (b.dataset.upload !== undefined) return goStart();
      if (b.dataset.home !== undefined) return goHome();
      if (b.dataset.toggle) { state.open.has(b.dataset.toggle) ? state.open.delete(b.dataset.toggle) : state.open.add(b.dataset.toggle); return renderNav(); }
      if (b.dataset.sheet) openTab(b.dataset.file, b.dataset.sheet);
    });
    $('#view-home').addEventListener('click', (e) => {
      const t = e.target.closest('.tile');
      if (t) openTab(t.dataset.file, t.dataset.sheet);
    });
    $('#tabs').addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      state.view = b.dataset.view; state.focus = null;
      render();
    });

    $('#filters').addEventListener('change', (e) => {
      const s = e.target.closest('select[data-filter]');
      if (!s) return;
      if (s.value) state.filters[s.dataset.filter] = s.value; else delete state.filters[s.dataset.filter];
      state.page = 0; render();
    });
    $('#filters').addEventListener('input', (e) => {
      if (e.target.id !== 'searchBox') return;
      state.search = e.target.value; state.page = 0;
      clearTimeout(wire._t);
      wire._t = setTimeout(() => { const sel = findSel(); if (!sel) return; state.view === 'dashboard' ? renderDashboard(sel.sheet) : renderData(sel.sheet); }, 200);
    });
    $('#filters').addEventListener('click', (e) => {
      if (e.target.id === 'clearFilters') { state.filters = {}; state.search = ''; state.page = 0; render(); }
    });

    $('#view-dashboard').addEventListener('change', (e) => {
      if (e.target.id !== 'measureSel') return;
      state.measure = e.target.value === '' ? null : Number(e.target.value);
      const sel = findSel(); if (sel) renderDashboard(sel.sheet);
    });
    $('#view-home').addEventListener('click', (e) => {
      if (e.target.id !== 'fixAll') return;
      downloadFixList(state.files.flatMap((f) => f.sheets.map((sheet) => ({ file: f, sheet }))), 'all files');
    });
    $('#view-health').addEventListener('click', (e) => {
      if (e.target.id === 'fixTab') { const sel = findSel(); if (sel) downloadFixList([{ file: sel.file, sheet: sel.sheet }], sel.sheet.title); return; }
      const b = e.target.closest('[data-focus]');
      if (!b) return;
      const sel = findSel(); if (!sel) return;
      const it = sel.sheet.issues[Number(b.dataset.focus)];
      state.focus = { title: it.title, rows: new Set(it.rows), cols: [...new Set(it.cells.map((c) => c[1]))] };
      state.filters = {}; state.search = ''; state.page = 0; state.sort = null;
      state.view = 'data';
      render();
      $('#main').scrollTop = 0;
    });
    $('#view-data').addEventListener('click', (e) => {
      if (e.target.id === 'clearFocus') { state.focus = null; state.page = 0; return render(); }
      const th = e.target.closest('th[data-sort]');
      if (th) {
        const ci = Number(th.dataset.sort);
        state.sort = state.sort && state.sort.ci === ci ? (state.sort.dir > 0 ? { ci, dir: -1 } : null) : { ci, dir: 1 };
        return render();
      }
      const pg = e.target.closest('[data-page]');
      if (pg) { state.page += Number(pg.dataset.page); render(); $('.table-wrap').scrollTop = 0; }
    });

    $('#themeBtn').addEventListener('click', () => {
      const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
      document.documentElement.dataset.theme = next;
      try { localStorage.setItem('sd-theme', next); } catch (e) { /* ignore */ }
      render();
    });
    // Keep "Updated x min ago" fresh.
    setInterval(updateStatus, 30000);
  }

  async function init() {
    try { const t = localStorage.getItem('sd-theme'); if (t) document.documentElement.dataset.theme = t; } catch (e) { /* ignore */ }
    wire();
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
