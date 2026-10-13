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

// The Google Drive files this dashboard shows (edit sources.json to change them).
function loadSources() {
  try { return JSON.parse(fs.readFileSync(path.join(__dirname, 'sources.json'), 'utf8')); }
  catch (e) { return []; }
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
function allSources() {
  return [...localFiles(), ...loadSources().map((s) => ({ ...s, kind: 'drive' }))];
}

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
  // Only the files listed in sources.json can be fetched.
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

// Only the app's own page may change files (blocks other websites from posting to localhost).
function fromApp(req) {
  const origin = req.headers.origin || '';
  return /^http:\/\/(localhost|127\.0\.0\.1):\d+$/.test(origin) && req.headers['x-app'] === 'sheets-dashboard';
}
const cleanFileName = (s) => path.basename(String(s || '')).replace(/[<>:"/\\|?*\x00-\x1f]/g, '').trim();

// Save an uploaded Excel file into data/ (same name = replaces the older copy).
// "Sales Tracker (1).xlsx" (a second download of the same sheet) replaces "Sales Tracker.xlsx".
// ?replace=<file in data/> forces which file it replaces (the "Replace" button).
const MAX_UPLOAD = 50 * 1024 * 1024;
function handleUpload(req, res, url) {
  if (!fromApp(req)) return sendJson(res, 403, { error: 'Not allowed.' });
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
  const list = loadSources();
  if (!list.some((s) => s.id === id)) return sendJson(res, 404, { error: 'Link not found.' });
  saveSources(list.filter((s) => s.id !== id));
  sendJson(res, 200, { ok: true });
}

function saveSources(list) {
  fs.writeFileSync(path.join(__dirname, 'sources.json'), JSON.stringify(list, null, 2) + '\n');
}

http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  if (req.method === 'POST' && url.pathname === '/api/upload') return handleUpload(req, res, url);
  if (req.method === 'POST' && url.pathname === '/api/remove') return handleRemove(req, res, url);
  if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405); return res.end(); }
  if (url.pathname === '/api/drive') return handleDrive(req, res, url);
  if (url.pathname === '/api/local') return handleLocal(req, res, url);
  if (url.pathname === '/api/sources') return sendJson(res, 200, allSources());

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
