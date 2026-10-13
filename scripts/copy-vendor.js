// Copies the two browser libraries into public/vendor so the app runs fully offline.
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const out = path.join(root, 'public', 'vendor');
fs.mkdirSync(out, { recursive: true });
const files = [
  ['node_modules/xlsx/dist/xlsx.full.min.js', 'xlsx.full.min.js'],
  ['node_modules/chart.js/dist/chart.umd.js', 'chart.umd.js'],
];
for (const [src, dst] of files) {
  const from = path.join(root, src);
  if (!fs.existsSync(from)) { console.warn('missing', src); continue; }
  fs.copyFileSync(from, path.join(out, dst));
  console.log('vendor:', dst);
}
