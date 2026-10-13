// analyze.js — turns a raw Excel tab into clean rows + finds data problems.
// Pure functions only (no page/DOM), so it can be tested from Node too.
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Analyze = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, sept: 8, oct: 9, nov: 10, dec: 11 };
  const PLACEHOLDERS = new Set(['-', '--', 'na', 'n/a', 'nil', 'none', '#n/a', '.']);

  // ---------- small helpers ----------
  const isBlank = (v) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  const isPlaceholder = (v) => typeof v === 'string' && PLACEHOLDERS.has(v.trim().toLowerCase());
  const tidy = (s) => String(s).replace(/\s+/g, ' ').trim();
  // Key used to spot "same thing written differently": case, spaces, dots, commas.
  const normKey = (s) => String(s).toLowerCase().replace(/[.,'"()\-_/]+/g, ' ').replace(/\s+/g, ' ').trim();

  // "Amount_Uploaded" -> "Amount Uploaded", "CUSTOMER NAME" -> "Customer Name" (keeps PO, DSR, SKU…)
  const ACRONYMS = new Set(['PO', 'DSR', 'SKU', 'GRN', 'GST', 'SGST', 'CGST', 'IGST', 'UOM', 'HSN', 'MB', 'SI', 'SO', 'ID', 'KG', 'L', 'MRP', 'YTD', 'OBAM', 'DIP', 'JISPL', 'SF', 'HP', 'ATF', 'MOQ', 'Q1', 'Q2', 'Q3', 'Q4', 'KL', 'LTR', 'SPAR', 'P', 'DVP', 'OEM']);
  function prettyName(s) {
    s = tidy(String(s).replace(/_/g, ' ')
      .replace(/([a-z])([A-Z])/g, '$1 $2')          // camelCase
      .replace(/([A-Z]+)([A-Z][a-z]{2,})/g, '$1 $2') // GRNApproved (but not SKUs)
      .replace(/([A-Za-z])-(\d)/g, '$1 $2'));
    return s.replace(/[A-Za-z][A-Za-z0-9.]*/g, (w) => {
      const key = w.replace(/\./g, '').toUpperCase();
      if (ACRONYMS.has(key)) return key === w.replace(/\./g, '') ? w : key;
      if (w === w.toUpperCase() || w === w.toLowerCase()) return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
      return w.charAt(0).toUpperCase() + w.slice(1);
    });
  }

  function colLetter(i) {
    let s = '';
    for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
    return s;
  }

  function toNum(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    if (typeof v !== 'string') return null;
    let s = v.trim().replace(/[₹,\s]/g, '').replace(/^rs\.?/i, '');
    let pct = false;
    if (s.endsWith('%')) { pct = true; s = s.slice(0, -1); }
    if (!/^[-+]?(\d+\.?\d*|\.\d+)$/.test(s)) return null;
    const n = parseFloat(s);
    return pct ? n / 100 : n;
  }

  function mkDate(y, m, d) {
    if (y < 100) y += 2000;
    if (m < 0 || m > 11 || d < 1 || d > 31) return null;
    const dt = new Date(y, m, d);
    return dt.getMonth() === m ? dt : null; // rejects 31-Feb etc.
  }

  // Indian style: day comes before month.
  function parseDateStr(s) {
    s = String(s).trim().toLowerCase();
    if (!s || s.length > 25) return null;
    let m;
    if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return mkDate(+m[1], +m[2] - 1, +m[3]);
    if ((m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})(\s+\d{1,2}:\d{2}(:\d{2})?)?$/))) return mkDate(+m[3], +m[2] - 1, +m[1]);
    if ((m = s.match(/^(\d{1,2})(?:st|nd|rd|th)?[\s\-\/.]*([a-z]{3,9})[\s\-\/.,]*(\d{2,4})$/)) && MONTHS[m[2].slice(0, 3)] !== undefined)
      return mkDate(+m[3], MONTHS[m[2].slice(0, 3)], +m[1]);
    if ((m = s.match(/^([a-z]{3,9})[\s\-\/.]*(\d{1,2}),?\s*(\d{4})$/)) && MONTHS[m[1].slice(0, 3)] !== undefined)
      return mkDate(+m[3], MONTHS[m[1].slice(0, 3)], +m[2]);
    return null;
  }

  function toDate(v) {
    if (v instanceof Date) return isNaN(v) ? null : v;
    if (typeof v === 'number' && v > 20000 && v < 80000) { // Excel serial date
      return new Date(Math.round((v - 25569) * 86400000) + new Date().getTimezoneOffset() * 60000);
    }
    if (typeof v === 'string') return parseDateStr(v);
    return null;
  }

  // ---------- header + column understanding ----------
  function looksLikeLabel(v) {
    return typeof v === 'string' && v.trim().length > 0 && v.trim().length <= 60 && toNum(v) === null && !parseDateStr(v);
  }

  function findHeader(aoa) {
    let best = -1, bestScore = 0;
    const lim = Math.min(aoa.length, 20);
    for (let i = 0; i < lim; i++) {
      const row = aoa[i] || [];
      let labels = 0, filled = 0;
      for (const v of row) { if (isBlank(v)) continue; filled++; if (looksLikeLabel(v)) labels++; }
      if (labels < 2) continue;
      const below = (aoa[i + 1] || []).filter((x) => !isBlank(x)).length + (aoa[i + 2] || []).filter((x) => !isBlank(x)).length;
      if (below < 2) continue;
      const score = labels + (labels / filled >= 0.8 ? 2 : 0);
      if (score > bestScore) { bestScore = score; best = i; }
    }
    return best < 0 ? 0 : best;
  }

  function roleOf(h) {
    h = h.toLowerCase().replace(/_/g, ' ');
    if (/description|remark|comment|note/.test(h)) return /description/.test(h) ? 'product' : null;
    if (/(price|rate)/.test(h) && !/total/.test(h)) return 'price';
    if (/customer|client|party|buyer/.test(h)) return 'customer';
    if (/\bdsr\b|sales ?(person|man)|executive|assigned|owner/.test(h)) return 'person';
    if (/status|stage/.test(h)) return 'status';
    if (/date|\bdt\b/.test(h)) return 'date';
    if (/(price|rate)/.test(h) && !/total/.test(h)) return 'price';
    if (/qty|quantity|no\. ?of|\bnos\b|stock|balance|sold|purchased|barrel|bucket|l\/kg|ltr|litre/.test(h)) return 'qty';
    if (/total|amount|\bamt\b|value|freight|charges|cost|difference|payout|incentive/.test(h)) return 'amount';
    if (/product|item|description|grade|material/.test(h)) return 'product';
    if (/type|category|uom|quarter|pack|month|segment|zone|region|city|location/.test(h)) return 'group';
    return null;
  }

  function inferType(name, vals) {
    const n = vals.length;
    if (!n) return 'empty';
    let num = 0, date = 0;
    for (const v of vals) {
      if (v instanceof Date) date++;
      else if (toNum(v) !== null) num++;
      else if (typeof v === 'string' && parseDateStr(v)) date++;
    }
    const h = name.toLowerCase().replace(/_/g, ' ');
    if (/date|\bdt\b/.test(h) && date > 0 && (date + num) / n >= 0.5) return 'date';
    if (date / n >= 0.6) return 'date';
    if (num / n >= 0.6) {
      const idLike = /(\bno\b|\bno\.|\bnob\b|number|code|\bsku\b|\bid\b|invoice|\bpo\b|\bsr\b|serial|\bbill\b|\bhsn\b|\bgst\b no)/.test(h);
      const measureLike = /qty|quantity|amount|total|price|rate|no\. ?of|value|freight|days|potential/.test(h);
      return idLike && !measureLike ? 'id' : 'number';
    }
    const distinct = new Set(vals.map((v) => normKey(v))).size;
    if (distinct <= 30 || distinct / n <= 0.35) return 'category';
    return 'text';
  }

  // ---------- main: one tab -> model ----------
  function analyzeSheet(name, aoa, opts) {
    opts = opts || {};
    const today = opts.today || new Date();
    const h = findHeader(aoa);
    const headerRow = aoa[h] || [];
    let width = 0;
    for (let i = h; i < aoa.length; i++) if (aoa[i] && aoa[i].length > width) width = aoa[i].length;

    // Collect data rows (skip blank rows, repeated header rows and lone subtotal rows).
    const headerKeys = headerRow.map((v) => (isBlank(v) ? '' : normKey(v)));
    const headerFilled = headerKeys.filter(Boolean).length;
    const headerSet = new Set(headerKeys.filter(Boolean));
    const rawRows = [];
    const restarts = []; // header-like rows in a different layout = a new table starts here
    let repeatedHeaders = 0, totalRows = 0;
    for (let i = h + 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const filled = [];
      for (let c = 0; c < width; c++) if (!isBlank(row[c])) filled.push(c);
      if (!filled.length) continue;
      const sameAsHeader = filled.filter((c) => headerKeys[c] && normKey(row[c]) === headerKeys[c]).length;
      if (headerFilled >= 3 && sameAsHeader >= Math.max(3, headerFilled * 0.6)) { repeatedHeaders++; continue; }
      const headerWords = filled.filter((c) => typeof row[c] === 'string' && headerSet.has(normKey(row[c]))).length;
      if (headerFilled >= 3 && headerWords >= 3) { restarts.push(i + 1); repeatedHeaders++; continue; }
      if (width >= 5 && filled.length <= 2 && filled.every((c) => toNum(row[c]) !== null)) { totalRows++; continue; }
      rawRows.push({ r: i + 1, raw: row });
    }

    // Columns that have a header or any data.
    const cols = [];
    const usedNames = {};
    for (let c = 0; c < width; c++) {
      const vals = [];
      for (const rr of rawRows) { const v = rr.raw[c]; if (!isBlank(v) && !isPlaceholder(v)) vals.push(v); }
      const hv = headerRow[c];
      if (isBlank(hv) && !vals.length) continue;
      let nm = isBlank(hv) ? `Column ${colLetter(c)}` : prettyName(hv);
      if (usedNames[nm]) nm = `${nm} (${colLetter(c)})`;
      usedNames[nm] = 1;
      const type = inferType(nm, vals);
      let role = roleOf(nm);
      if (role === 'date' && type !== 'date') role = null;
      if ((role === 'qty' || role === 'amount' || role === 'price') && type !== 'number') role = null;
      cols.push({ idx: c, letter: colLetter(c), name: nm, raw: isBlank(hv) ? '' : String(hv), type, role, fill: rawRows.length ? vals.length / rawRows.length : 0 });
    }

    // Normalised values: numbers -> Number, dates -> Date, text -> trimmed string.
    const rows = rawRows.map((rr) => {
      const v = cols.map((col) => {
        const x = rr.raw[col.idx];
        if (isBlank(x) || isPlaceholder(x)) return null;
        if (col.type === 'number') { const n = toNum(x); return n === null ? tidy(x) : n; }
        if (col.type === 'date') { const d = toDate(x); return d || tidy(x); }
        if (x instanceof Date) return x;
        return typeof x === 'string' ? tidy(x) : x;
      });
      return { r: rr.r, raw: cols.map((col) => rr.raw[col.idx]), v };
    });

    const model = { name: tidy(name) || name, headerRow: h + 1, cols, rows, restarts, skipped: { repeatedHeaders, totalRows } };
    model.targets = findTargets(aoa, h);
    model.names = nameSuggestions(model, name);
    model.issues = findIssues(model, today);
    model.cellIssues = {};
    const rank = { err: 3, warn: 2, info: 1 };
    for (const iss of model.issues) {
      for (const cell of iss.cells) {
        const k = cell[0] + ':' + cell[1];
        const cur = model.cellIssues[k];
        if (!cur || rank[iss.sev] > rank[cur.sev]) model.cellIssues[k] = { sev: iss.sev, title: iss.title };
      }
    }
    model.score = healthScore(model);
    return model;
  }

  // ---------- "Target vs Achievement" tables ----------
  // Layout: a label row ("Sales In Target Q1", "... Achivement Q1") above a month header row,
  // then one row per person (Role, Name) and usually a Total row.
  function findTargets(aoa, h) {
    if (h < 1) return null;
    const head = aoa[h] || [], above = aoa[h - 1] || [];
    const items = [];
    let label = '';
    for (let c = 0; c < head.length; c++) {
      if (!isBlank(above[c])) label = String(above[c]).toLowerCase();
      if (isBlank(head[c])) continue;
      const hv = String(head[c]).toLowerCase().replace(/sales\s*(in|out)/, '').trim();
      const mm = hv.match(/^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*$/);
      if (!mm) continue;
      const kind = /target|tgt/.test(label) ? 'target' : /ach/.test(label) ? 'ach' : null;
      if (kind) items.push({ c, kind, m: MONTHS[mm[1]] });
    }
    if (items.filter((i) => i.kind === 'target').length < 3 || items.filter((i) => i.kind === 'ach').length < 3) return null;

    const hk = head.map((v) => (isBlank(v) ? '' : String(v).toLowerCase()));
    let nameCol = hk.findIndex((v) => v === 'name');
    const roleCol = hk.findIndex((v) => v === 'role');
    if (nameCol < 0) nameCol = 1;
    const ytdT = hk.findIndex((v) => /ytd/.test(v) && /tgt|target/.test(v));
    const ytdA = hk.findIndex((v) => /ytd/.test(v) && /ach/.test(v));

    let unit = '';
    for (let i = 0; i < h - 1; i++) {
      const first = (aoa[i] || []).find((v) => !isBlank(v));
      if (typeof first === 'string' && first.length <= 20) { unit = first.trim(); break; }
    }

    const months = [...new Set(items.map((i) => i.m))].sort((a, b) => a - b);
    const people = [];
    for (let i = h + 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      if (row.every(isBlank)) { if (people.length) break; continue; }
      const role = roleCol >= 0 && !isBlank(row[roleCol]) ? tidy(row[roleCol]) : '';
      const name = !isBlank(row[nameCol]) ? tidy(row[nameCol]) : role;
      if (!name) continue;
      const target = {}, ach = {};
      for (const it of items) {
        const n = toNum(row[it.c]);
        if (n === null) continue;
        (it.kind === 'target' ? target : ach)[it.m] = n;
      }
      people.push({
        role, name, isTotal: /^total/i.test(role || name), row: i + 1, target, ach,
        yearTarget: ytdT >= 0 ? toNum(row[ytdT]) : null, ytdAch: ytdA >= 0 ? toNum(row[ytdA]) : null,
      });
    }
    return people.length ? { unit, months, people } : null;
  }

  // ---------- suggested values for the fix list ----------
  const ddmmyyyy = (d) => `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
  function suggestDate(s) {
    s = String(s);
    const parts = s.match(/\d{1,2}[\/.\-]+\d{1,2}[\/.\-]+\d{2,4}/g) || [];
    if (parts.length >= 2) {
      const d = parseDateStr(parts[0].replace(/([\/.\-])+/g, '$1'));
      return `Two dates in one cell — keep one${d ? ` (e.g. ${ddmmyyyy(d)})` : ''}, put the other on its own row`;
    }
    const d = parseDateStr(s.replace(/([\/.\-])+/g, '$1').replace(/\s+/g, ' ').trim());
    if (d) return ddmmyyyy(d);
    return /[a-z]{3,}/i.test(s) ? `This is a note, not a date — leave the date blank and move “${s}” to a Remarks column` : 'Retype as DD/MM/YYYY';
  }
  function suggestOddDate(d, today) {
    const swapped = d.getDate() <= 12 ? new Date(d.getFullYear(), d.getDate() - 1, d.getMonth() + 1) : null;
    if (swapped && swapped <= today && swapped.getFullYear() >= 2015) return `${ddmmyyyy(swapped)}? (day and month may be swapped)`;
    return `Check the year (${ddmmyyyy(d)})`;
  }
  function suggestNumber(s) {
    s = String(s);
    const m = s.match(/^\s*([\d,]+(?:\.\d+)?)\s*([a-z]+)?\s*$/i);
    if (m) return `${m[1].replace(/,/g, '')}${m[2] ? ` (put “${m[2]}” in the column heading)` : ''}`;
    return `Leave blank, write “${s}” in a Remarks column`;
  }

  // ---------- problem finder ----------
  function fmtNum(n) { return Number.isInteger(n) ? n.toLocaleString('en-IN') : n.toLocaleString('en-IN', { maximumFractionDigits: 2 }); }
  function rowList(rows) {
    const u = [...new Set(rows)].sort((a, b) => a - b);
    return u.length > 12 ? u.slice(0, 12).join(', ') + ` … (+${u.length - 12} more)` : u.join(', ');
  }

  function findIssues(m, today) {
    const issues = [];
    const add = (sev, title, detail, cells) => {
      if (!cells.length) return;
      issues.push({ sev, title, detail, cells, rows: [...new Set(cells.map((c) => c[0]))].sort((a, b) => a - b) });
    };
    const N = m.rows.length;
    if (!N) return issues;

    if (m.restarts.length) {
      issues.push({
        sev: 'warn', title: 'Headings repeated partway down the tab',
        detail: `The column headings appear again at row ${m.restarts.slice(0, 5).join(', ')}${m.restarts.length > 5 ? ' …' : ''}, sometimes in different columns. If the columns below are in a different order, the checks below may mix them up — keeping one table per tab avoids this.`,
        cells: [], rows: m.restarts.slice(),
      });
    }

    m.cols.forEach((col, ci) => {
      const label = `“${col.name}” (column ${col.letter})`;

      // 1. Text typed into a number column.
      if (col.type === 'number') {
        const bad = m.rows.filter((r) => typeof r.v[ci] === 'string');
        if (bad.length) {
          const ex = [...new Set(bad.map((r) => r.v[ci]))].slice(0, 5).map((s) => `“${s}”`).join(', ');
          add('err', `Text in a number column: ${label}`, `${bad.length} cell(s) are not numbers, so totals ignore them. Examples: ${ex}.`, bad.map((r) => [r.r, ci, suggestNumber(r.v[ci])]));
        }
        if (col.role === 'qty' || col.role === 'amount' || col.role === 'price') {
          const neg = m.rows.filter((r) => typeof r.v[ci] === 'number' && r.v[ci] < 0);
          add('warn', `Negative values in ${label}`, `${neg.length} cell(s) are below zero — check if that is intended.`, neg.map((r) => [r.r, ci, `Check: should it be ${fmtNum(Math.abs(r.v[ci]))}?`]));
        }
      }

      // 2. Dates that are unreadable or out of range.
      if (col.type === 'date') {
        const bad = m.rows.filter((r) => typeof r.v[ci] === 'string');
        if (bad.length) {
          const ex = [...new Set(bad.map((r) => r.v[ci]))].slice(0, 5).map((s) => `“${s}”`).join(', ');
          add('err', `Unreadable dates in ${label}`, `${bad.length} cell(s) are not a valid date. Examples: ${ex}.`, bad.map((r) => [r.r, ci, suggestDate(r.v[ci])]));
        }
        const future = new Date(today.getTime() + 60 * 86400000);
        const odd = m.rows.filter((r) => r.v[ci] instanceof Date && (r.v[ci] > future || r.v[ci].getFullYear() < 2015));
        if (odd.length) {
          const ex = odd.slice(0, 4).map((r) => r.v[ci].toLocaleDateString('en-GB')).join(', ');
          add('warn', `Suspicious dates in ${label}`, `${odd.length} date(s) are far in the future or before 2015 (${ex}). Often day and month are swapped or the year is mistyped.`, odd.map((r) => [r.r, ci, suggestOddDate(r.v[ci], today)]));
        }
      }

      // 3. Missing values in a column that is normally filled.
      const important = ['customer', 'person', 'status', 'product', 'date', 'qty', 'amount'].includes(col.role);
      if (col.type !== 'empty' && col.fill < 1 && ((important && col.fill >= 0.6) || col.fill >= 0.85)) {
        const blanks = m.rows.filter((r) => r.v[ci] === null);
        add('warn', `Missing ${label}`, `${blanks.length} of ${N} rows are empty here while the rest are filled.`, blanks.map((r) => [r.r, ci, 'Fill in (or leave if not known yet)']));
      }

      // 4. Same name written in different ways.
      if ((col.type === 'category' || col.type === 'text') && ['customer', 'person', 'status', 'product', 'group'].includes(col.role || (col.type === 'category' ? 'group' : ''))) {
        const groups = new Map();
        m.rows.forEach((r) => {
          const val = r.v[ci];
          if (typeof val !== 'string') return;
          const k = normKey(val);
          if (!groups.has(k)) groups.set(k, new Map());
          const g = groups.get(k);
          g.set(val, (g.get(val) || []).concat(r.r));
        });
        const msgs = [], cells = [];
        for (const g of groups.values()) {
          if (g.size < 2) continue;
          const sorted = [...g.entries()].sort((a, b) => b[1].length - a[1].length);
          msgs.push(`“${sorted[0][0]}” also written as ${sorted.slice(1, 4).map((e) => `“${e[0]}”`).join(', ')}`);
          sorted.slice(1).forEach((e) => e[1].forEach((rn) => cells.push([rn, ci, sorted[0][0]])));
        }
        if (cells.length) add('warn', `Same name spelled differently in ${label}`, msgs.slice(0, 4).join('; ') + (msgs.length > 4 ? `; +${msgs.length - 4} more` : '') + '. Reports count these as separate names.', cells);
      }
    });

    // 5. Totals that don't add up (C should equal A × B, or A − B).
    const numCols = m.cols.map((c, i) => ({ c, i })).filter((x) => x.c.type === 'number' && x.c.fill >= 0.2).sort((a, b) => b.c.fill - a.c.fill).slice(0, 14);
    const tol = (exp, got) => Math.abs(exp - got) <= Math.max(1, Math.abs(got) * 0.01);
    const usedTriples = new Set();
    for (const target of numCols) {
      if (target.c.role === 'price') continue;
      let best = null;
      for (let a = 0; a < numCols.length; a++) for (let b = a; b < numCols.length; b++) {
        const A = numCols[a], B = numCols[b];
        if (A.i === target.i || B.i === target.i) continue;
        for (const op of ['×', '−']) {
          if (op === '−' && a === b) continue;
          const pairs = op === '×' ? [[A, B]] : [[A, B], [B, A]];
          for (const [X, Y] of pairs) {
            let n = 0, ok = 0, trivial = 0; const bad = [];
            for (const r of m.rows) {
              const x = r.v[X.i], y = r.v[Y.i], t = r.v[target.i];
              if (typeof x !== 'number' || typeof y !== 'number' || typeof t !== 'number') continue;
              const exp = op === '×' ? x * y : x - y;
              if (t === 0 && exp === 0) continue; // empty-ish rows prove nothing
              if (op === '−' && y === 0) continue; // "A − 0 = A" proves nothing
              n++;
              if (op === '×' && (x === 1 || y === 1)) trivial++;
              if (op === '−' && (x === 0 || Math.abs(y) < Math.abs(t) * 0.02)) trivial++; // tiny B hides inside rounding
              if (tol(exp, t)) ok++; else bad.push({ r: r.r, exp, t });
            }
            const key = [X.i, Y.i, target.i].sort().join(',');
            if (n >= 5 && ok >= 5 && trivial < n * 0.5 && ok / n >= 0.6 && bad.length && !usedTriples.has(key) && (!best || ok / n > best.rate)) {
              best = { rate: ok / n, X, Y, op, bad, n };
            }
          }
        }
      }
      if (best) {
        usedTriples.add([best.X.i, best.Y.i, target.i].sort().join(','));
        const formula = `${best.X.c.name} ${best.op} ${best.Y.c.name}`;
        const ex = best.bad.slice(0, 3).map((b) => `row ${b.r}: has ${fmtNum(b.t)}, expected ${fmtNum(b.exp)}`).join('; ');
        add('err', `“${target.c.name}” doesn't match ${formula}`,
          `In ${Math.round(best.rate * 100)}% of rows ${target.c.name} = ${formula}, but not in these ${best.bad.length}. ${ex}.`,
          best.bad.map((b) => [b.r, target.i, `${fmtNum(b.exp)} (= ${formula}) — or correct the other number`]));
      }
    }

    // 6. Exact duplicate rows.
    const seen = new Map();
    m.rows.forEach((r) => {
      const k = r.v.map((x) => (x instanceof Date ? x.getTime() : x === null ? '' : String(x).toLowerCase())).join('\u0001');
      seen.set(k, (seen.get(k) || []).concat(r.r));
    });
    const dupCells = [], dupMsgs = [];
    for (const rs of seen.values()) if (rs.length > 1 && m.cols.length > 1) { dupMsgs.push(`rows ${rs.join(' & ')}`); rs.forEach((rn, k) => dupCells.push([rn, 0, k === 0 ? `Same as row ${rs.slice(1).join(', ')} — keep one` : `Copy of row ${rs[0]} — delete if typed twice`])); }
    if (dupCells.length) add('warn', 'Duplicate rows', `${dupMsgs.length} row(s) appear more than once with exactly the same values: ${dupMsgs.slice(0, 6).join(', ')}${dupMsgs.length > 6 ? ' …' : ''}.`, dupCells);

    // 7. Cosmetic: extra spaces / line breaks inside cells.
    const spaceCells = [], breakCells = [];
    m.rows.forEach((r) => r.raw.forEach((x, ci) => {
      if (typeof x !== 'string' || !x.trim()) return;
      if (/\n/.test(x.trim())) breakCells.push([r.r, ci]);
      else if (x !== x.trim() || /\s{2,}/.test(x)) spaceCells.push([r.r, ci]);
    }));
    if (spaceCells.length) {
      const colsHit = [...new Set(spaceCells.map((c) => m.cols[c[1]].name))].slice(0, 5).join(', ');
      add('info', 'Extra spaces in text', `${spaceCells.length} cell(s) have spaces at the start/end or double spaces (in ${colsHit}). Harmless to read, but can break lookups and matching.`, spaceCells);
    }
    if (breakCells.length) add('info', 'Line breaks inside cells', `${breakCells.length} cell(s) contain a line break (pressed Alt+Enter), which can hide part of the value.`, breakCells);

    // A plain "how to fix" line for each kind of problem.
    const FIXES = [
      [/^Text in a number column/, 'Type only numbers in this column. Put notes like “Cancelled” in a separate Remarks column.'],
      [/^Negative values/, 'Check whether the minus sign is correct.'],
      [/^Unreadable dates/, 'Type one date per cell, as DD/MM/YYYY (for example 25/03/2024).'],
      [/^Suspicious dates/, 'Check the year, and that the day and month are not swapped.'],
      [/^Missing/, 'Fill in the empty cells — or ignore this if the value is not known yet.'],
      [/^Same name spelled/, 'Pick one spelling and use it everywhere, so totals add up under one name.'],
      [/doesn't match/, 'Re-check the numbers in these rows and correct whichever one is wrong.'],
      [/^Duplicate rows/, 'If the same entry was typed twice, delete the extra copy.'],
      [/^Extra spaces/, 'Optional: remove extra spaces (Find & Replace, or the TRIM formula).'],
      [/^Line breaks/, 'Optional: edit the cell and keep everything on one line.'],
      [/^Headings repeated/, 'Keep one table per tab — move each extra table to its own tab.'],
    ];
    issues.forEach((it) => { const f = FIXES.find(([re]) => re.test(it.title)); it.fix = f ? f[1] : ''; });

    const order = { err: 0, warn: 1, info: 2 };
    issues.sort((a, b) => order[a.sev] - order[b.sev] || b.cells.length - a.cells.length);
    return issues;
  }

  // ---------- clearer names for tabs and columns ----------
  const SPELLING = {
    quatation: 'Quotation', quotaion: 'Quotation', qoutation: 'Quotation', discription: 'Description', descripton: 'Description',
    achivement: 'Achievement', acheivement: 'Achievement', indsutrial: 'Industrial', recieved: 'Received', ammount: 'Amount',
    adress: 'Address', quantiy: 'Quantity', invocie: 'Invoice', custmer: 'Customer', costumer: 'Customer', stell: 'Steel',
    nob: 'No.', tgt: 'Target', avl: 'Available', dt: 'Date', ltrs: 'Litres', ltr: 'Litres',
    tot: 'Total', amt: 'Amount', ach: 'Achieved', itemcode: 'Item Code',
  };
  const SMALL = new Set(['of', 'to', 'in', 'on', 'and', 'for', 'by', 'be', 'due', 'the', 'a', 'at', 'per', 'from', 'or', 'with']);
  function cleanName(raw) {
    let s = prettyName(raw)
      .replace(/(\d+)\s*(st|nd|rd|th)\b/gi, (_, n, o) => n + o.toLowerCase()) // 13 TH / 23RD -> 13th / 23rd
      .replace(/\b(Report)\s+\d+$/i, '$1');                                   // "...Report 17" -> "...Report"
    s = s.replace(/[A-Za-z]+\.?/g, (w, at) => {
      const k = w.replace(/\.$/, '').toLowerCase();
      if (SPELLING[k]) return SPELLING[k]; // "Amt." -> "Amount", "nob" -> "No."
      return at > 0 && SMALL.has(k) && !w.endsWith('.') ? k : w; // "No. of Barrel", "Price to be Uploaded"
    });
    return s.replace(/\.\.+/g, '.').replace(/\s+/g, ' ').trim();
  }
  function reasons(raw, clean) {
    const why = [];
    const r = String(raw);
    if (r !== r.trim() || /\s{2,}/.test(r)) why.push('extra spaces');
    if (/_/.test(r)) why.push('underscores');
    if (Object.keys(SPELLING).some((k) => new RegExp(`\\b${k}\\b`, 'i').test(r))) why.push('spelling / short form');
    if (/[A-Z]{4,}/.test(r.replace(/[^A-Za-z ]/g, '')) && clean !== r.trim()) why.push('ALL CAPITALS');
    if (/\b(Report)\s+\d+$/i.test(prettyName(r))) why.push('extra number');
    return why;
  }
  function nameSuggestions(m, rawTab) {
    const out = [];
    const tabClean = cleanName(rawTab);
    const tabWhy = reasons(rawTab, tabClean);
    if (tabWhy.length) out.push({ what: 'Tab', where: 'Tab name', current: String(rawTab), suggested: tabClean, why: tabWhy.join(', ') });
    const seen = new Map();
    for (const col of m.cols) {
      const where = `Column ${col.letter}`;
      if (!col.raw) {
        if (col.fill >= 0.2 && !m.targets) out.push({ what: 'Column', where, current: '(no heading)', suggested: 'Add a heading that says what this column holds', why: 'no heading' });
        continue;
      }
      let clean = cleanName(col.raw);
      const why = reasons(col.raw, clean);
      // Money columns: say so in the heading.
      const money = col.role === 'amount' && col.type === 'number' && !/qty|quantity|stock|litre|ltr|kg|days|points|potential/i.test(col.raw);
      if (money && !/₹|\brs\b|inr|rupee/i.test(clean)) {
        clean = clean.replace(/^Total\s+(.+?)\s+Rate$/i, '$1 Value') + ' (₹)';
        why.push('add the unit (₹)');
      }
      const key = normKey(col.raw);
      if (seen.has(key) && !m.targets) {
        why.push(`same name as column ${seen.get(key)}`);
        clean = `${clean} (2)`;
      } else seen.set(key, col.letter);
      // Only suggest when there is a real reason (not just one capital letter in normal text).
      if (why.length && clean !== col.raw) out.push({ what: 'Column', where, current: col.raw, suggested: clean, why: why.join(', ') });
    }
    return out;
  }

  function healthScore(m) {
    if (!m.rows.length) return 100;
    const total = m.rows.length * Math.max(1, m.cols.length);
    let bad = 0;
    for (const k in m.cellIssues) bad += m.cellIssues[k].sev === 'err' ? 3 : m.cellIssues[k].sev === 'warn' ? 1 : 0.1;
    return Math.max(0, Math.min(100, Math.round(100 - (bad / total) * 100 * 4)));
  }

  return { analyzeSheet, prettyName, toNum, toDate, parseDateStr, normKey, colLetter, findHeader };
});
