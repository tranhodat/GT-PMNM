// Kiểm thử xuất Word trong Node: node test/export.test.js
const fs = require('fs'), path = require('path');
const { createRequire } = require('module');
const D = createRequire((process.env.DOCX_DIR || process.cwd()) + '/package.json')('docx');
const E = require('../js/export.js');

function jpegSize(buf) {
  let i = 2;
  while (i < buf.length) {
    if (buf[i] !== 0xFF) { i++; continue; }
    const m = buf[i + 1];
    if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) return { height: buf.readUInt16BE(i + 5), width: buf.readUInt16BE(i + 7) };
    i += 2 + buf.readUInt16BE(i + 2);
  }
  return { width: 1400, height: 800 };
}
const root = path.join(__dirname, '..');
async function loader(src) {
  const buf = fs.readFileSync(path.join(root, src));
  const s = jpegSize(buf);
  return { data: new Uint8Array(buf), width: s.width, height: s.height };
}
(async () => {
  const nums = (process.argv[2] || '1,2,3,4,5,6').split(',').map(Number);
  const chapters = nums.map(n => JSON.parse(fs.readFileSync(path.join(root, 'content', `chuong-${n}.json`), 'utf8')));
  const doc = await E.build(D, chapters, loader, { title: 'Test' });
  const buf = await D.Packer.toBuffer(doc);
  const out = path.join(process.env.TEMP || '/tmp', 'test-giaotrinh.docx');
  fs.writeFileSync(out, buf);
  console.log('OK', out, buf.length, 'bytes');
})().catch(e => { console.error('FAIL', e); process.exit(1); });
