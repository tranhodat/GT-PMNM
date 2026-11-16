/* Hàm serverless (Vercel, Node 18+): ghi nội dung chương và ảnh vào kho GitHub qua REST API.
   Biến môi trường (Vercel → Settings → Environment Variables):
     GITHUB_TOKEN   token cá nhân (fine-grained) có quyền Contents: Read and write cho kho này
     GITHUB_REPO    ví dụ: ten-nguoi-dung/giao-trinh-pmnm
     GITHUB_BRANCH  nhánh ghi (mặc định: main)
     EDIT_PASSWORD  mật khẩu người biên tập (gửi trong header x-edit-password) */
'use strict';
const crypto = require('crypto');

function same(a, b) {
  const x = Buffer.from(String(a || '')), y = Buffer.from(String(b || ''));
  return x.length === y.length && crypto.timingSafeEqual(x, y);
}

async function gh(path, opts, token) {
  const r = await fetch('https://api.github.com' + path, Object.assign({}, opts, {
    headers: Object.assign({ Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28', 'User-Agent': 'giao-trinh-pmnm-editor' }, (opts && opts.headers) || {})
  }));
  const text = await r.text(); let json = null; try { json = text ? JSON.parse(text) : null; } catch (e) { json = null; }
  return { status: r.status, ok: r.ok, json };
}

async function putFile(repo, branch, token, path, contentB64, message) {
  const url = '/repos/' + repo + '/contents/' + path.split('/').map(encodeURIComponent).join('/');
  const cur = await gh(url + '?ref=' + encodeURIComponent(branch), { method: 'GET' }, token);
  const body = { message, content: contentB64, branch };
  if (cur.ok && cur.json && cur.json.sha) body.sha = cur.json.sha;
  const put = await gh(url, { method: 'PUT', body: JSON.stringify(body), headers: { 'Content-Type': 'application/json' } }, token);
  if (!put.ok) throw new Error('GitHub: ' + ((put.json && put.json.message) || put.status));
  return put.json && put.json.commit && put.json.commit.sha;
}

module.exports = async function handler(req, res) {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  const { GITHUB_TOKEN, GITHUB_REPO, EDIT_PASSWORD } = process.env;
  const branch = process.env.GITHUB_BRANCH || 'main';
  const configured = !!(GITHUB_TOKEN && GITHUB_REPO && EDIT_PASSWORD);

  if (req.method === 'GET') return res.status(200).json({ configured, branch });
  if (req.method !== 'POST') return res.status(405).json({ error: 'Phương thức không được hỗ trợ' });
  if (!configured) return res.status(503).json({ error: 'Máy chủ chưa cấu hình GITHUB_TOKEN, GITHUB_REPO và EDIT_PASSWORD' });
  if (!same(req.headers['x-edit-password'], EDIT_PASSWORD)) return res.status(401).json({ error: 'Sai mật khẩu chỉnh sửa' });

  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { return res.status(400).json({ error: 'JSON không hợp lệ' }); } }
  if (!body || typeof body !== 'object') return res.status(400).json({ error: 'Thiếu nội dung yêu cầu' });

  try {
    if (body.action === 'image') {
      const id = String(body.id || '');
      if (!/^[a-z0-9]{3,40}$/.test(id)) return res.status(400).json({ error: 'Mã ảnh không hợp lệ' });
      const data = String(body.data || '');
      if (!/^[A-Za-z0-9+/=]+$/.test(data) || data.length > 3.2e6) return res.status(400).json({ error: 'Ảnh không hợp lệ hoặc quá lớn (tối đa khoảng 2,4 MB)' });
      const path = 'assets/img/uploads/' + id + '.jpg';
      const commit = await putFile(GITHUB_REPO, branch, GITHUB_TOKEN, path, data, 'Thêm ảnh ' + id + ' từ trình biên tập');
      return res.status(200).json({ ok: true, path, commit });
    }
    if (body.action === 'chapter') {
      const n = Number(body.number);
      if (!Number.isInteger(n) || n < 1 || n > 6) return res.status(400).json({ error: 'Số chương không hợp lệ' });
      const d = body.data;
      if (!d || !Array.isArray(d.blocks) || d.number !== n) return res.status(400).json({ error: 'Dữ liệu chương không hợp lệ' });
      const json = JSON.stringify(d, null, 1);
      if (json.length > 3.5e6) return res.status(413).json({ error: 'Nội dung chương quá lớn' });
      const path = 'content/chuong-' + n + '.json';
      const commit = await putFile(GITHUB_REPO, branch, GITHUB_TOKEN, path, Buffer.from(json, 'utf8').toString('base64'), 'Cập nhật chương ' + n + ' từ trình biên tập');
      return res.status(200).json({ ok: true, path, commit });
    }
    return res.status(400).json({ error: 'Hành động không hợp lệ' });
  } catch (e) {
    return res.status(502).json({ error: String(e.message || e) });
  }
};
