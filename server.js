// Tiny local web server for the Sheets Dashboard.
// - Listens on 127.0.0.1 only: nothing on the network can reach it.
// - Serves the app's own files from /public.
// - Uploaded Excel files are saved only into ./data on this PC (never sent anywhere else).
// - /api/drive only DOWNLOADS (read-only) a Google Drive file the owner has
//   shared by link. Nothing is ever written to Google Drive.
// - /api/upload, /api/remove change only this app's own list/copies,
//   and only when the request comes from the app's own page.
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = Number(process.env.PORT) || 5180;
const ROOT = path.join(__dirname, 'public');
const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
};
const CSP = "default-src 'self'; img-src 'self' data: blob:; style-src 'self' 'unsafe-inline'; connect-src 'self'";

function sendJson(res, code, obj) {
  res.writeHead(code, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(obj));
}

// The Google Drive files this dashboard shows.
// - On a PC: sources.json (git-ignored).
// - On hosting (Vercel): the SOURCES environment variable, e.g.
//   [{"id":"<drive file id>","name":"Shell Order Tracker"}] — kept in the host's private settings.
function parseSources(text) {
  try {
    const v = JSON.parse(text);
    return Array.isArray(v) ? v.filter((s) => s && /^[A-Za-z0-9_-]{20,100}$/.test(s.id)) : [];
  } catch (e) {
    // Also accept a plain list of links / ids separated by commas or new lines.
    return String(text).split(/[\s,]+/).map((t) => (t.match(/\/d\/([A-Za-z0-9_-]{20,100})/) || t.match(/^([A-Za-z0-9_-]{20,100})$/) || [])[1])
      .filter(Boolean).map((id) => ({ id, name: 'Google sheet' }));
  }
}
function sourcesFromFile() {
  try { return parseSources(fs.readFileSync(path.join(__dirname, 'sources.json'), 'utf8')); } catch (e) { return []; }
}
function loadSources() {
  return process.env.SOURCES ? parseSources(process.env.SOURCES) : sourcesFromFile();
}

// ---------- password for shared Google sheets on a public website ----------
// Google-sheet data is only given out when either:
//  - it comes from sources.json on this PC and no password is set (the PC app), or
//  - APP_PASSWORD is set and the browser has unlocked it (the website).
// SOURCES without APP_PASSWORD is never served: that would publish the data to anyone.
const crypto = require('crypto');
const APP_PASSWORD = process.env.APP_PASSWORD || '';
const authToken = () => crypto.createHash('sha256').update('sheets-insight:' + APP_PASSWORD).digest('hex');
function hasAuthCookie(req) {
  const m = (req.headers.cookie || '').match(/(?:^|;\s*)sd_auth=([a-f0-9]{64})/);
  return !!m && crypto.timingSafeEqual(Buffer.from(m[1]), Buffer.from(authToken()));
}
function driveAllowed(req) {
  if (APP_PASSWORD) return hasAuthCookie(req);
  return !process.env.SOURCES; // PC app with a local sources.json
}
function handleLogin(req, res) {
  if (!fromApp(req)) return sendJson(res, 403, { error: 'Not allowed.' });
  if (!APP_PASSWORD) return sendJson(res, 400, { error: 'No password is set up for this site.' });
  let body = '';
  req.on('data', (c) => { body += c; if (body.length > 2000) req.destroy(); });
  req.on('end', () => {
    let pw = '';
    try { pw = String(JSON.parse(body).password || ''); } catch (e) { /* ignore */ }
    const a = crypto.createHash('sha256').update(pw).digest();
    const b = crypto.createHash('sha256').update(APP_PASSWORD).digest();
    if (!crypto.timingSafeEqual(a, b)) {
      return setTimeout(() => sendJson(res, 401, { error: 'Wrong password.' }), 800); // slow down guessing
    }
    const secure = (req.headers['x-forwarded-proto'] || '').includes('https') ? '; Secure' : '';
    res.writeHead(200, {
      'Content-Type': 'application/json', 'Cache-Control': 'no-store',
      'Set-Cookie': `sd_auth=${authToken()}; Path=/; HttpOnly; SameSite=Strict; Max-Age=${30 * 86400}${secure}`,
    });
    res.end(JSON.stringify({ ok: true }));
  });
}
function handleLogout(req, res) {
  if (!fromApp(req)) return sendJson(res, 403, { error: 'Not allowed.' });
  res.writeHead(200, { 'Content-Type': 'application/json', 'Set-Cookie': 'sd_auth=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0' });
  res.end('{"ok":true}');
}

// Excel files placed in the data folder are shown too. Replace a file there to update it.
const DATA_DIR = path.join(__dirname, 'data');
function localFiles() {
  try {
    return fs.readdirSync(DATA_DIR)
      .filter((f) => /\.(xlsx|xlsm|xls|csv)$/i.test(f) && !f.startsWith('~$') && !f.startsWith('.'))
      .map((f) => ({ id: 'local:' + f, name: f.replace(/\.(xlsx|xlsm|xls|csv)$/i, ''), kind: 'local', file: f,
        mtime: fs.statSync(path.join(DATA_DIR, f)).mtimeMs }))
      .sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) { return []; }
}
function allSources(req) {
  const drive = driveAllowed(req) ? loadSources().map((s) => ({ id: s.id, name: s.name || 'Google sheet', kind: 'drive' })) : [];
  return [...localFiles(), ...drive];
}
// How many shared sheets are waiting behind the password (0 = none / already unlocked).
const lockedCount = (req) => (APP_PASSWORD && !hasAuthCookie(req) ? loadSources().length : 0);

function handleLocal(req, res, url) {
  const name = url.searchParams.get('file') || '';
  const hit = localFiles().find((f) => f.file === name); // only files actually in data/
  if (!hit) return sendJson(res, 404, { error: 'File not found in the data folder.' });
  fs.readFile(path.join(DATA_DIR, hit.file), (err, buf) => {
    if (err) return sendJson(res, 500, { error: 'Could not read the file. Is it open in Excel?' });
    res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Cache-Control': 'no-store', 'X-File-Name': encodeURIComponent(hit.name) });
    res.end(buf);
  });
}

async function fetchDrive(id) {
  // Uploaded Excel files and native Google Sheets live behind different URLs.
  const urls = [
    `https://drive.usercontent.google.com/download?id=${id}&export=download&confirm=t`,
    `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`,
  ];
  for (const url of urls) {
    const r = await fetch(url, { redirect: 'follow' });
    const type = r.headers.get('content-type') || '';
    if (r.ok && !type.includes('text/html')) {
      const cd = r.headers.get('content-disposition') || '';
      const m = cd.match(/filename\*=UTF-8''([^;]+)/i) || cd.match(/filename="([^"]+)"/i);
      return { buf: Buffer.from(await r.arrayBuffer()), name: m ? decodeURIComponent(m[1]) : '' };
    }
  }
  return null;
}

async function handleDrive(req, res, url) {
  const id = url.searchParams.get('id') || '';
  // Only listed files, and only for a browser that is allowed to see them.
  if (!driveAllowed(req)) return sendJson(res, 401, { error: 'Enter the password to load the shared Google sheets.' });
  if (!loadSources().some((s) => s.id === id)) return sendJson(res, 400, { error: 'Unknown file.' });
  try {
    const got = await fetchDrive(id);
    if (!got) {
      return sendJson(res, 403, {
        error: 'This file is private (Restricted). The owner needs to set Share → Anyone with the link → Viewer.',
      });
    }
    res.writeHead(200, {
      'Content-Type': 'application/octet-stream',
      'Cache-Control': 'no-store',
      'X-File-Name': encodeURIComponent(got.name),
    });
    res.end(got.buf);
  } catch (e) {
    sendJson(res, 502, { error: 'Could not reach Google Drive. Check the internet connection.' });
  }
}

// A Google Sheets link pasted by the user in the page. The link is kept in the user's browser;
// the server only downloads a read-only copy and passes it through — nothing is stored here.
// Requires the app's header so other websites cannot use this as a download proxy.
const MAX_LINK = 50 * 1024 * 1024;
async function handleLink(req, res, url) {
  if (!fromApp(req)) return sendJson(res, 403, { error: 'Not allowed.' });
  const id = url.searchParams.get('id') || '';
  if (!/^[A-Za-z0-9_-]{20,100}$/.test(id)) return sendJson(res, 400, { error: 'That does not look like a Google Sheets link.' });
  try {
    const got = await fetchDrive(id);
    if (!got) {
      return sendJson(res, 403, { error: 'This sheet is private. In Google Sheets: Share → General access → “Anyone with the link” → Viewer, then try again.' });
    }
    if (got.buf.length > MAX_LINK) return sendJson(res, 413, { error: 'That sheet is too big (max 50 MB).' });
    res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Cache-Control': 'no-store', 'X-File-Name': encodeURIComponent(got.name) });
    res.end(got.buf);
  } catch (e) {
    sendJson(res, 502, { error: 'Could not reach Google Drive. Check the internet connection.' });
  }
}

// Only the app's own page may change files (blocks other websites from posting to localhost).
// The custom X-App header is the real protection: a browser will not let another website send it
// (it needs CORS permission this server never gives). Origin, when present, must be this machine.
function fromApp(req) {
  const origin = req.headers.origin || '';
  const okOrigin = !origin || origin === 'null' || /^https?:\/\/(localhost|127\.0\.0\.1|\[::1\])(:\d+)?$/i.test(origin)
    || origin.replace(/^https?:\/\//i, '') === req.headers.host;
  const ok = req.headers['x-app'] === 'sheets-dashboard' && okOrigin;
  if (!ok) console.warn(`Refused ${req.method} ${req.url.split('?')[0]} — origin: ${origin || '(none)'}, app header: ${req.headers['x-app'] ? 'yes' : 'no'}`);
  return ok;
}
const cleanFileName = (s) => path.basename(String(s || '')).replace(/[<>:"/\\|?*\x00-\x1f]/g, '').trim();

// Save an uploaded Excel file into data/ (same name = replaces the older copy).
// "Sales Tracker (1).xlsx" (a second download of the same sheet) replaces "Sales Tracker.xlsx".
// ?replace=<file in data/> forces which file it replaces (the "Replace" button).
const MAX_UPLOAD = 50 * 1024 * 1024;
function handleUpload(req, res, url) {
  if (!checkCanSave()) return sendJson(res, 501, { error: 'This website cannot store files. Reload the page — files will be kept in your browser instead.' });
  if (!fromApp(req)) return sendJson(res, 403, { error: 'Upload blocked: please open the app from the “Start Dashboard” shortcut (http://localhost:5180) and try again.' });
  let name = cleanFileName(url.searchParams.get('name'));
  if (!/\.(xlsx|xlsm|xls|csv)$/i.test(name) || name.startsWith('~$')) return sendJson(res, 400, { error: 'Please choose an Excel file (.xlsx, .xls or .csv).' });
  name = name.replace(/\s*\(\d+\)(?=\.[a-z]+$)/i, ''); // drop the browser's " (1)"
  const replace = cleanFileName(url.searchParams.get('replace'));
  if (replace) {
    if (!localFiles().some((f) => f.file === replace)) return sendJson(res, 404, { error: 'The file to replace was not found.' });
    if (path.extname(replace).toLowerCase() !== path.extname(name).toLowerCase()) {
      return sendJson(res, 400, { error: `Please choose a ${path.extname(replace)} file to replace “${replace}”.` });
    }
    name = replace;
  }
  const chunks = [];
  let size = 0, tooBig = false;
  req.on('data', (c) => { size += c.length; if (size > MAX_UPLOAD) tooBig = true; else chunks.push(c); });
  req.on('end', () => {
    if (tooBig) return sendJson(res, 413, { error: 'That file is too big (max 50 MB).' });
    try {
      fs.mkdirSync(DATA_DIR, { recursive: true });
      const tmp = path.join(DATA_DIR, `.${name}.uploading`);
      fs.writeFileSync(tmp, Buffer.concat(chunks));
      fs.renameSync(tmp, path.join(DATA_DIR, name)); // swap in only when fully written
      sendJson(res, 200, { ok: true, name });
    } catch (e) {
      sendJson(res, 500, { error: 'Could not save the file. If it is open in Excel, close it and try again.' });
    }
  });
}

// Remove a file from the dashboard. Uploaded copies are moved to data/_removed (not deleted),
// links are just taken off the list. The original Google Sheet is never touched.
function handleRemove(req, res, url) {
  if (!fromApp(req)) return sendJson(res, 403, { error: 'Not allowed.' });
  const id = url.searchParams.get('id') || '';
  if (id.startsWith('local:')) {
    const hit = localFiles().find((f) => f.id === id);
    if (!hit) return sendJson(res, 404, { error: 'File not found.' });
    try {
      const bin = path.join(DATA_DIR, '_removed');
      fs.mkdirSync(bin, { recursive: true });
      const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
      fs.renameSync(path.join(DATA_DIR, hit.file), path.join(bin, `${stamp} ${hit.file}`));
      return sendJson(res, 200, { ok: true });
    } catch (e) { return sendJson(res, 500, { error: 'Could not remove the file. If it is open in Excel, close it and try again.' }); }
  }
  if (process.env.SOURCES) return sendJson(res, 400, { error: 'This sheet is set up in the website settings (Vercel → Environment Variables → SOURCES). Remove it there.' });
  const list = loadSources();
  if (!list.some((s) => s.id === id)) return sendJson(res, 404, { error: 'Link not found.' });
  saveSources(list.filter((s) => s.id !== id));
  sendJson(res, 200, { ok: true });
}

// Can this server keep uploaded files? (No on read-only hosting such as Vercel —
// then the page keeps uploads in the visitor's own browser instead.)
let canSave = null;
function checkCanSave() {
  if (canSave !== null) return canSave;
  if (process.env.READONLY === '1') return (canSave = false); // test switch: behave like read-only hosting
  try {
    fs.mkdirSync(DATA_DIR, { recursive: true });
    const probe = path.join(DATA_DIR, '.write-test');
    fs.writeFileSync(probe, 'ok');
    fs.unlinkSync(probe);
    canSave = true;
  } catch (e) { canSave = false; }
  return canSave;
}

function saveSources(list) {
  fs.writeFileSync(path.join(__dirname, 'sources.json'), JSON.stringify(list, null, 2) + '\n');
}

http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  if (req.method === 'POST' && url.pathname === '/api/upload') return handleUpload(req, res, url);
  if (req.method === 'POST' && url.pathname === '/api/remove') return handleRemove(req, res, url);
  if (req.method === 'POST' && url.pathname === '/api/login') return handleLogin(req, res);
  if (req.method === 'POST' && url.pathname === '/api/logout') return handleLogout(req, res);
  if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405); return res.end(); }
  if (url.pathname === '/api/drive') return handleDrive(req, res, url);
  if (url.pathname === '/api/link') return handleLink(req, res, url);
  if (url.pathname === '/api/local') return handleLocal(req, res, url);
  if (url.pathname === '/api/sources') return sendJson(res, 200, allSources(req));
  if (url.pathname === '/api/info') {
    return sendJson(res, 200, {
      canSave: checkCanSave(),
      locked: lockedCount(req),                    // shared sheets waiting for the password
      unlocked: !!APP_PASSWORD && hasAuthCookie(req),
      misconfigured: !!process.env.SOURCES && !APP_PASSWORD, // SOURCES set but no password: not served
    });
  }

  let rel = decodeURIComponent(url.pathname);
  if (rel === '/') rel = '/index.html';
  const file = path.normalize(path.join(ROOT, rel));
  if (!file.startsWith(ROOT)) { res.writeHead(403); return res.end(); }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); return res.end('Not found'); }
    res.writeHead(200, {
      'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream',
      'Cache-Control': 'no-store',
      'Content-Security-Policy': CSP,
    });
    res.end(data);
  });
}).listen(PORT, '127.0.0.1', () => {
  console.log(`Sheets Dashboard running at http://localhost:${PORT}`);
});
