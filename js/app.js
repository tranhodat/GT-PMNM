/* Trình biên tập giáo trình: hiển thị, chỉnh sửa trực tiếp, lưu cục bộ, đồng bộ GitHub và xuất Word. */
(function () {
  'use strict';
  var M = window.Markup, E = window.DocxExport, S = window.Store;
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  var NCH = 6;
  var KINDS = { nho: 'Ghi nhớ', luuy: 'Lưu ý', vidu: 'Ví dụ', mo: 'Mở rộng', tinhhuong: 'Tình huống' };
  var TYPES = { h2: 'Tiêu đề mục', h3: 'Tiêu đề mục con', h4: 'Tiêu đề cấp 4', p: 'Đoạn văn', list: 'Danh sách', callout: 'Hộp nhấn mạnh',
                code: 'Khối mã', quote: 'Trích dẫn mẫu', table: 'Bảng', figure: 'Hình', refs: 'Tài liệu tham khảo' };
  var PT = (function () { var d = document.createElement('div'); try { d.contentEditable = 'plaintext-only'; } catch (e) { return 'true'; } return d.contentEditable === 'plaintext-only' ? 'plaintext-only' : 'true'; })();

  var st = { ch: {}, cur: 1, edited: {}, urls: {}, saveT: null, citeT: null, range: null, rangeEl: null, insertAfter: -1, figTarget: -1, saved: null, loadErr: null };

  /* ---------- tiện ích ---------- */
  function h(tag, cls, html) { var e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; }
  var esc = M.esc;
  function toast(msg, ms) { var t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(function () { t.hidden = true; }, ms || 2600); }
  function setStatus(msg, cls) { var s = $('#status'); s.textContent = msg; s.className = 'status ' + (cls || ''); }
  function cur() { return st.ch[st.cur]; }
  function rid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 5); }
  function words(ch) {
    var n = 0; ch.blocks.forEach(function (b) { M.blockStrings(b).concat(b.t === 'h2' || b.t === 'h3' || b.t === 'h4' ? [b.text] : []).forEach(function (s) { n += String(s || '').trim().split(/\s+/).filter(Boolean).length; }); });
    return n;
  }
  function modal(html, onReady) { $('#modalBody').innerHTML = html; $('#modal').hidden = false; if (onReady) onReady($('#modalBody')); }
  function closeModal() { $('#modal').hidden = true; $('#modalBody').innerHTML = ''; }

  /* ---------- nạp / lưu ---------- */
  function fetchChapter(n) {
    return fetch('content/chuong-' + n + '.json?v=' + Date.now()).then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); });
  }
  function loadAll() {
    var jobs = [];
    for (var n = 1; n <= NCH; n++) (function (n) {
      jobs.push(S.get('chapters', n).then(function (local) {
        if (local && local.blocks) { st.ch[n] = local; st.edited[n] = true; return; }
        return fetchChapter(n).then(function (d) { st.ch[n] = d; });
      }));
    })(n);
    return Promise.all(jobs);
  }
  function markDirty() {
    st.edited[st.cur] = true; setStatus('Đang lưu…', 'busy'); renderSideDots();
    clearTimeout(st.saveT); st.saveT = setTimeout(saveNow, 700);
  }
  function saveNow() {
    clearTimeout(st.saveT); var n = st.cur, ch = st.ch[n];
    return S.put('chapters', n, JSON.parse(JSON.stringify(ch))).then(function () {
      st.saved = new Date(); statusOk();
    }).catch(function () { setStatus('Không lưu được vào trình duyệt (hết dung lượng?)', 'err'); });
  }
  function statusOk() {
    var ch = cur(); var t = st.saved ? st.saved.toTimeString().slice(0, 5) : '';
    setStatus('Chương ' + ch.number + ' · ' + words(ch).toLocaleString('vi-VN') + ' từ' + (st.edited[ch.number] ? ' · đã lưu trên trình duyệt ' + t : ' · bản gốc'), 'ok');
  }

  /* ---------- thanh bên ---------- */
  function renderSide() {
    var nav = $('#chapters'); nav.innerHTML = '';
    for (var n = 1; n <= NCH; n++) {
      var c = st.ch[n]; if (!c) continue;
      var b = h('button', n === st.cur ? 'on' : '', '<span class="n">' + n + '</span>' + esc(c.titleText || c.title) + '<span class="dot" title="Đã chỉnh sửa">' + (st.edited[n] ? '●' : '') + '</span>');
      b.dataset.n = n; nav.appendChild(b);
    }
  }
  function renderSideDots() { $$('#chapters button').forEach(function (b) { var d = $('.dot', b); if (d) d.textContent = st.edited[+b.dataset.n] ? '●' : ''; }); }
  function renderOutline() {
    var o = $('#outline'); o.innerHTML = '';
    cur().blocks.forEach(function (b, i) {
      if (b.t !== 'h2' && b.t !== 'h3') return;
      var a = h('a', b.t === 'h3' ? 'l3' : '', esc(b.text || '(chưa có tiêu đề)')); a.href = '#b' + i; a.dataset.i = i; o.appendChild(a);
    });
  }

  /* ---------- hiển thị khối ---------- */
  function ctl(b) {
    var c = h('div', 'ctl');
    c.innerHTML = '<span class="tag">' + (b.t === 'callout' ? KINDS[b.kind] || 'Hộp' : TYPES[b.t] || b.t) + '</span>' +
      '<button data-act="up" title="Chuyển lên" aria-label="Chuyển lên">▲</button><button data-act="down" title="Chuyển xuống" aria-label="Chuyển xuống">▼</button>' +
      '<button data-act="dup" title="Nhân đôi" aria-label="Nhân đôi">⧉</button><button data-act="del" class="del" title="Xóa khối" aria-label="Xóa khối">✕</button>';
    return c;
  }
  function editable(cls, md, map, ph, extra) {
    var d = h('div', cls, M.mdToHtml(md, map)); d.contentEditable = 'true'; d.dataset.ph = ph || ''; d.dataset.f = extra || 'text'; d.spellcheck = true; d.lang = 'vi'; return d;
  }
  function plain(cls, text, ph, f) { var d = h('div', cls); d.textContent = text || ''; d.contentEditable = PT; d.dataset.ph = ph || ''; d.dataset.f = f || 'text'; d.lang = 'vi'; return d; }
  function imgSrc(src) { if (!src) return ''; if (src.indexOf('idb:') === 0) return st.urls[src.slice(4)] || ''; return src; }

  function blockEl(b, i, map, num) {
    var sec = h('section', 'blk'); sec.dataset.i = i; sec.dataset.t = b.t; sec.id = 'b' + i;
    sec.appendChild(ctl(b));
    switch (b.t) {
      case 'h2': case 'h3': case 'h4': sec.appendChild(plain(b.t, b.text, 'Tiêu đề')); break;
      case 'p': { var p = editable('p' + (b.indent === false ? ' noind' : '') + (b.left ? ' left' : ''), b.text, map, 'Nhập nội dung…'); sec.appendChild(p); break; }
      case 'quote': sec.appendChild(editable('quote', b.text, map, 'Nội dung trích dẫn mẫu…')); break;
      case 'list': {
        var ul = h('ul', 'list');
        (b.items || []).forEach(function (it, j) { ul.appendChild(liEl(it, j, b.ordered, map)); });
        sec.appendChild(ul); var add = h('button', 'mini', '+ Thêm mục'); add.dataset.act = 'li-add'; sec.appendChild(add); break;
      }
      case 'callout': {
        var box = h('div', 'callout k-' + b.kind), hd = h('div', 'ch');
        var sel = h('select'); sel.dataset.f = 'kind'; Object.keys(KINDS).forEach(function (k) { var o = h('option', null, KINDS[k]); o.value = k; if (k === b.kind) o.selected = true; sel.appendChild(o); });
        var ti = h('input'); ti.type = 'text'; ti.placeholder = 'Tiêu đề (không bắt buộc)'; ti.value = b.title || ''; ti.dataset.f = 'title';
        hd.appendChild(sel); hd.appendChild(ti); box.appendChild(hd);
        (b.texts || []).forEach(function (t) { box.appendChild(editable('ct', t, map, 'Nội dung…', 'ct')); });
        var ca = h('button', 'mini', '+ Thêm đoạn'); ca.dataset.act = 'ct-add'; box.appendChild(ca); sec.appendChild(box); break;
      }
      case 'code': {
        var cb = h('div', 'codebox'), t1 = h('input'); t1.type = 'text'; t1.placeholder = 'Tiêu đề khối mã (không bắt buộc)'; t1.value = b.title || ''; t1.dataset.f = 'title';
        var ta = h('textarea'); ta.value = b.text || ''; ta.rows = Math.max(2, String(b.text || '').split('\n').length); ta.spellcheck = false; ta.dataset.f = 'text';
        cb.appendChild(t1); cb.appendChild(ta); sec.appendChild(cb); break;
      }
      case 'table': sec.appendChild(tableEl(b, map, num)); break;
      case 'figure': sec.appendChild(figureEl(b, map, num)); break;
      case 'refs': sec.appendChild(refsEl()); break;
    }
    var ins = h('button', 'ins', '+ Thêm nội dung bên dưới'); ins.dataset.act = 'ins'; sec.appendChild(ins);
    return sec;
  }
  function liEl(it, j, ordered, map) {
    var li = h('li'); li.appendChild(h('span', 'mk', ordered ? (j + 1) + '.' : '•'));
    li.appendChild(editable('li', it, map, 'Nội dung mục…', 'li'));
    var rm = h('button', 'rm', '✕'); rm.dataset.act = 'li-rm'; rm.title = 'Xóa mục'; rm.setAttribute('aria-label', 'Xóa mục'); li.appendChild(rm); return li;
  }
  function capEl(label, md, map) {
    var cp = h('div', 'cap'); var b = h('b', null, esc(label || '')); b.className = 'lbl'; cp.appendChild(b); cp.appendChild(document.createTextNode(' '));
    var e = editable('ce', md, map, 'Chú thích…', 'cap'); cp.appendChild(e); return cp;
  }
  function tableEl(b, map, num) {
    var w = h('div', 'tbl'); w.appendChild(capEl(num, b.caption, map));
    var t = h('table'), tr = h('tr'), n = (b.header || []).length;
    (b.header || []).forEach(function (x) { var th = h('th'); th.contentEditable = PT; th.dataset.f = 'th'; th.textContent = x; tr.appendChild(th); });
    var thead = h('thead'); thead.appendChild(tr); t.appendChild(thead);
    var tb = h('tbody');
    (b.rows || []).forEach(function (r) {
      var row = h('tr');
      for (var i = 0; i < n; i++) { var td = h('td', null, M.mdToHtml(r[i] == null ? '' : r[i], map)); td.contentEditable = 'true'; td.dataset.f = 'td'; td.lang = 'vi'; row.appendChild(td); }
      tb.appendChild(row);
    });
    t.appendChild(tb); w.appendChild(t);
    var tools = h('div', 'tools');
    [['row+', '+ Hàng'], ['row-', '− Hàng'], ['col+', '+ Cột'], ['col-', '− Cột']].forEach(function (x) { var bt = h('button', 'mini', x[1]); bt.dataset.act = x[0]; tools.appendChild(bt); });
    w.appendChild(tools); return w;
  }
  function figureEl(b, map, num) {
    var f = h('div', 'fig'), u = imgSrc(b.src);
    if (b.src) { var im = h('img'); im.alt = b.caption || 'Hình minh họa'; im.dataset.src = b.src; if (u) im.src = u; im.style.width = Math.min(100, (b.width || 14) / 16 * 100) + '%'; f.appendChild(im); }
    else { var em = h('div', 'empty', 'Chưa có ảnh. Nhấn “Chọn ảnh” để tải ảnh lên.'); f.appendChild(em); }
    var tl = h('div', 'ftools');
    var up = h('button', 'mini', b.src ? 'Thay ảnh' : 'Chọn ảnh'); up.dataset.act = 'fig-up'; tl.appendChild(up);
    var sel = h('select'); sel.dataset.f = 'width'; sel.setAttribute('aria-label', 'Chiều rộng ảnh khi xuất Word');
    [8, 10, 12, 14, 15.5, 16].forEach(function (x) { var o = h('option', null, 'Rộng ' + x + ' cm'); o.value = x; if (+b.width === x) o.selected = true; sel.appendChild(o); });
    if (![8, 10, 12, 14, 15.5, 16].some(function (x) { return +b.width === x; })) { var o2 = h('option', null, 'Rộng ' + b.width + ' cm'); o2.value = b.width; o2.selected = true; sel.appendChild(o2); }
    tl.appendChild(sel); f.appendChild(tl);
    f.appendChild(capEl(b.src ? num : '', b.caption, map)); return f;
  }
  function refsEl() {
    var ch = cur(), ord = M.computeOrder(ch), w = h('div', 'refsblk');
    var ol = h('ol'); ord.order.forEach(function (k) { ol.appendChild(h('li', null, M.mdToHtml(ch.refs[k], {}))); });
    w.appendChild(ol);
    w.appendChild(h('p', 'hint', 'Danh sách này tự tạo từ các trích dẫn [n] trong chương. Dùng nút “Tài liệu tham khảo” để thêm, sửa hoặc xóa nguồn.'));
    return w;
  }

  /* ---------- hiển thị chương ---------- */
  function renderDoc(keepScroll) {
    var y = keepScroll ? window.scrollY : 0;
    var ch = cur(), ord = M.computeOrder(ch), nums = M.numbering(ch), doc = $('#doc'); doc.innerHTML = '';
    doc.appendChild(h('div', 'chap-no', 'CHƯƠNG ' + ch.number));
    var t = plain('chap-title', ch.titleText || ch.title, 'Tên chương', 'ctitle'); doc.appendChild(t);
    ch.blocks.forEach(function (b, i) { doc.appendChild(blockEl(b, i, ord.map, nums[i])); });
    if (!ch.blocks.length) { var e = h('p', 'hint', 'Chương chưa có nội dung. Dùng “+ Thêm nội dung” trên thanh công cụ.'); doc.appendChild(e); }
    renderSide(); renderOutline(); renderRefs(); loadIdbImages(); statusOk();
    $$('#doc textarea').forEach(autoGrow);
    if (keepScroll) window.scrollTo(0, y);
  }
  function rerenderBlock(i) {
    var ch = cur(), ord = M.computeOrder(ch), nums = M.numbering(ch), old = $('#b' + i);
    var nw = blockEl(ch.blocks[i], i, ord.map, nums[i]); old.parentNode.replaceChild(nw, old); loadIdbImages(); $$('textarea', nw).forEach(autoGrow); return nw;
  }
  function autoGrow(t) { t.style.height = 'auto'; t.style.height = (t.scrollHeight + 2) + 'px'; }
  function loadIdbImages() {
    $$('#doc img[data-src^="idb:"]').forEach(function (im) {
      if (im.src) return; var id = im.dataset.src.slice(4);
      S.get('images', id).then(function (blob) { if (!blob) return; st.urls[id] = URL.createObjectURL(blob); im.src = st.urls[id]; });
    });
  }
  function refreshCites() {
    var ch = cur(), ord = M.computeOrder(ch), nums = M.numbering(ch);
    $$('#doc .cite').forEach(function (c) { c.textContent = M.resolveCites('[@' + (c.dataset.keys || '').split(',').join(',@') + ']', ord.map); });
    $$('#doc .blk').forEach(function (sec) { var i = +sec.dataset.i; var l = $('.cap .lbl', sec); if (l && nums[i] !== undefined) l.textContent = nums[i]; });
    var rb = $('#doc .refsblk'); if (rb) { var nr = refsEl(); rb.parentNode.replaceChild(nr, rb); }
    renderRefs();
  }

  /* ---------- đồng bộ DOM -> dữ liệu ---------- */
  function md(el) { return M.htmlToMd(el).replace(/\s+/g, ' ').trim(); }
  function syncBlock(sec) {
    var ch = cur(), i = +sec.dataset.i, b = ch.blocks[i]; if (!b) return;
    var f = function (n) { return $('[data-f="' + n + '"]', sec); };
    switch (b.t) {
      case 'h2': case 'h3': case 'h4': b.text = f('text').textContent.replace(/\s+/g, ' ').trim(); break;
      case 'p': case 'quote': b.text = md(f('text')); break;
      case 'list': b.items = $$('.li', sec).map(md); break;
      case 'callout': b.kind = f('kind').value; b.title = f('title').value.trim(); b.texts = $$('.ct', sec).map(md); $('.callout', sec).className = 'callout k-' + b.kind; $('.ctl .tag', sec).textContent = KINDS[b.kind]; break;
      case 'code': b.text = f('text').value; b.title = f('title').value.trim(); break;
      case 'table':
        b.caption = md($('.cap [data-f="cap"]', sec)); b.header = $$('th', sec).map(function (x) { return x.textContent.replace(/\s+/g, ' ').trim(); });
        b.rows = $$('tbody tr', sec).map(function (r) { return $$('td', r).map(md); }); break;
      case 'figure': b.caption = md($('.cap [data-f="cap"]', sec)); b.width = +f('width').value; var im = $('img', sec); if (im) im.style.width = Math.min(100, b.width / 16 * 100) + '%'; break;
    }
  }
  function onInput(e) {
    var t = e.target;
    if (t.dataset && t.dataset.f === 'ctitle') { var ch = cur(); ch.titleText = t.textContent.replace(/\s+/g, ' ').trim(); ch.title = ch.titleText.toUpperCase(); markDirty(); return; }
    var sec = t.closest && t.closest('.blk'); if (!sec) return;
    if (t.tagName === 'TEXTAREA') autoGrow(t);
    syncBlock(sec); markDirty();
    var b = cur().blocks[+sec.dataset.i];
    clearTimeout(st.citeT); st.citeT = setTimeout(function () { refreshCites(); if (b.t === 'h2' || b.t === 'h3') renderOutline(); statusOk(); }, 500);
  }

  /* ---------- thao tác khối ---------- */
  function newBlock(type) {
    switch (type) {
      case 'p': return { t: 'p', text: '', indent: true, left: false };
      case 'h2': case 'h3': case 'h4': return { t: type, text: '' };
      case 'ul': return { t: 'list', ordered: false, items: [''] };
      case 'ol': return { t: 'list', ordered: true, items: [''] };
      case 'nho': case 'luuy': case 'vidu': case 'mo': case 'tinhhuong': return { t: 'callout', kind: type, title: '', texts: [''] };
      case 'code': return { t: 'code', text: '', title: '' };
      case 'quote': return { t: 'quote', text: '' };
      case 'table': return { t: 'table', caption: '', header: ['Cột 1', 'Cột 2', 'Cột 3'], rows: [['', '', ''], ['', '', '']], widths: [5.3, 5.3, 5.4], font: 12, firstBold: true, center: [] };
      case 'figure': return { t: 'figure', src: '', caption: '', width: 14 };
    }
  }
  function mutate(fn, focusIdx) {
    fn(cur().blocks); markDirty(); renderDoc(true);
    if (focusIdx != null) {
      var sec = $('#b' + focusIdx); if (sec) { sec.scrollIntoView({ block: 'center', behavior: 'smooth' }); var f = $('[contenteditable="true"],[contenteditable="plaintext-only"],textarea,input', sec); if (f && f.focus) f.focus(); }
    }
  }
  function insertBlock(type, after) {
    var nb = newBlock(type); if (!nb) return; var at = (after == null || after < -1) ? cur().blocks.length - 1 : after;
    mutate(function (bl) { bl.splice(at + 1, 0, nb); }, at + 1);
    if (type === 'figure') { st.figTarget = at + 1; setTimeout(function () { $('#fileImage').click(); }, 200); }
  }
  function onClick(e) {
    var btn = e.target.closest('[data-act]'); if (!btn) return;
    var sec = btn.closest('.blk'); var i = sec ? +sec.dataset.i : -1, ch = cur(), b = ch.blocks[i];
    switch (btn.dataset.act) {
      case 'up': if (i > 0) mutate(function (bl) { var x = bl.splice(i, 1)[0]; bl.splice(i - 1, 0, x); }, i - 1); break;
      case 'down': if (i < ch.blocks.length - 1) mutate(function (bl) { var x = bl.splice(i, 1)[0]; bl.splice(i + 1, 0, x); }, i + 1); break;
      case 'dup': mutate(function (bl) { bl.splice(i + 1, 0, JSON.parse(JSON.stringify(bl[i]))); }, i + 1); break;
      case 'del': if (confirm('Xóa khối “' + (TYPES[b.t] || b.t) + '” này?')) mutate(function (bl) { bl.splice(i, 1); }); break;
      case 'ins': openAddMenu(btn, i); break;
      case 'li-add': b.items.push(''); markDirty(); rerenderBlock(i); focusLast($('#b' + i + ' .li:last-of-type')); break;
      case 'li-rm': var li = btn.closest('li'); if (b.items.length > 1) { b.items.splice($$('li', sec).indexOf(li), 1); markDirty(); rerenderBlock(i); } break;
      case 'ct-add': b.texts.push(''); markDirty(); var ns = rerenderBlock(i); var cts = $$('.ct', ns); if (cts.length) cts[cts.length - 1].focus(); break;
      case 'row+': b.rows.push(b.header.map(function () { return ''; })); markDirty(); rerenderBlock(i); break;
      case 'row-': if (b.rows.length > 1) { b.rows.pop(); markDirty(); rerenderBlock(i); } break;
      case 'col+': b.header.push('Cột ' + (b.header.length + 1)); b.rows.forEach(function (r) { r.push(''); }); equalWidths(b); markDirty(); rerenderBlock(i); break;
      case 'col-': if (b.header.length > 1) { b.header.pop(); b.rows.forEach(function (r) { r.pop(); }); equalWidths(b); markDirty(); rerenderBlock(i); } break;
      case 'fig-up': st.figTarget = i; $('#fileImage').click(); break;
    }
  }
  function equalWidths(b) { var n = b.header.length; b.widths = b.header.map(function () { return +(16 / n).toFixed(2); }); }
  function focusLast(el) { if (el && el.focus) { el.focus(); } }
  function onKeydown(e) {
    var t = e.target, f = t.dataset && t.dataset.f;
    if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
      if (['th', 'td', 'cap', 'ctitle'].indexOf(f) >= 0 || t.classList.contains('h2') || t.classList.contains('h3') || t.classList.contains('h4')) { e.preventDefault(); t.blur(); return; }
      var sec = t.closest('.blk'); if (!sec) return; var i = +sec.dataset.i, b = cur().blocks[i];
      if (f === 'text' && b.t === 'p') { e.preventDefault(); insertBlock('p', i); return; }
      if (f === 'li') { e.preventDefault(); var j = $$('.li', sec).indexOf(t); b.items.splice(j + 1, 0, ''); markDirty(); var ns = rerenderBlock(i); var l = $$('.li', ns)[j + 1]; if (l) l.focus(); return; }
      if (f === 'ct') { e.preventDefault(); var j2 = $$('.ct', sec).indexOf(t); b.texts.splice(j2 + 1, 0, ''); markDirty(); var ns2 = rerenderBlock(i); var c2 = $$('.ct', ns2)[j2 + 1]; if (c2) c2.focus(); return; }
    }
    if (e.key === 'Backspace' && f === 'li' && !t.textContent) {
      var sec2 = t.closest('.blk'), i2 = +sec2.dataset.i, b2 = cur().blocks[i2], j3 = $$('.li', sec2).indexOf(t);
      if (b2.items.length > 1) { e.preventDefault(); b2.items.splice(j3, 1); markDirty(); var ns3 = rerenderBlock(i2); var l3 = $$('.li', ns3)[Math.max(0, j3 - 1)]; if (l3) l3.focus(); }
    }
  }
  function onPaste(e) {
    var t = e.target; if (t.tagName === 'TEXTAREA' || t.tagName === 'INPUT') return;
    if (!t.isContentEditable) return;
    e.preventDefault(); var txt = (e.clipboardData || window.clipboardData).getData('text/plain').replace(/\s*\n\s*/g, ' ');
    document.execCommand('insertText', false, txt);
  }

  /* ---------- thanh định dạng & trích dẫn ---------- */
  function trackSelection() {
    var s = window.getSelection(); if (!s.rangeCount) return;
    var r = s.getRangeAt(0), n = r.commonAncestorContainer; n = n.nodeType === 3 ? n.parentNode : n;
    var ed = n.closest && n.closest('[contenteditable="true"]'); if (ed && $('#doc').contains(ed)) { st.range = r.cloneRange(); st.rangeEl = ed; }
  }
  function restoreSelection() {
    if (!st.rangeEl || !document.body.contains(st.rangeEl)) return false;
    st.rangeEl.focus(); var s = window.getSelection(); s.removeAllRanges(); s.addRange(st.range); return true;
  }
  function fmt(cmd) {
    if (!restoreSelection()) { toast('Hãy bấm vào đoạn văn cần định dạng trước.'); return; }
    if (cmd === 'bold') document.execCommand('bold');
    else if (cmd === 'italic') document.execCommand('italic');
    else if (cmd === 'clear') document.execCommand('removeFormat');
    else if (cmd === 'code') {
      var sel = window.getSelection().toString(); if (!sel) { toast('Hãy chọn đoạn chữ cần định dạng thành mã.'); return; }
      document.execCommand('insertHTML', false, '<code>' + esc(sel) + '</code>');
    } else if (cmd === 'cite') openCite();
    st.rangeEl && st.rangeEl.dispatchEvent(new Event('input', { bubbles: true }));
  }
  function openCite() {
    var ch = cur(), keys = Object.keys(ch.refs || {}), ord = M.computeOrder(ch);
    var rows = keys.map(function (k) { return '<label><input type="checkbox" value="' + esc(k) + '"><span><b>' + (ord.map[k] ? '[' + ord.map[k] + '] ' : '') + '</b>' + esc(ch.refs[k]).slice(0, 220) + '</span></label>'; }).join('');
    modal('<h3>Chèn trích dẫn</h3><p>Chọn một hoặc nhiều nguồn trong danh mục của chương, hoặc thêm nguồn mới.</p>' +
      '<input type="search" id="citeQ" placeholder="Tìm nguồn…" aria-label="Tìm nguồn"><div class="citelist" id="citeList">' + (rows || '<p style="padding:10px">Chương chưa có tài liệu tham khảo.</p>') + '</div>' +
      '<label for="citeNew">Thêm nguồn mới (APA: tác giả, năm, tên tài liệu, nơi công bố, URL)</label><input type="text" id="citeNew" placeholder="Ví dụ: Open Source Initiative. (n.d.). The Open Source Definition. https://opensource.org/osd">' +
      '<div class="row"><button class="btn" id="citeCancel">Hủy</button><button class="btn primary" id="citeOk">Chèn</button></div>', function (m) {
        $('#citeQ', m).addEventListener('input', function () { var q = this.value.toLowerCase(); $$('#citeList label', m).forEach(function (l) { l.style.display = l.textContent.toLowerCase().indexOf(q) >= 0 ? '' : 'none'; }); });
        $('#citeCancel', m).onclick = closeModal;
        $('#citeOk', m).onclick = function () {
          var picked = $$('#citeList input:checked', m).map(function (c) { return c.value; }), nw = $('#citeNew', m).value.trim();
          if (nw) { var k = 'r' + rid(); ch.refs[k] = nw; picked.push(k); }
          if (!picked.length) { toast('Chưa chọn nguồn nào.'); return; }
          closeModal(); insertCite(picked);
        };
      });
  }
  function insertCite(keys) {
    if (!restoreSelection()) { toast('Không xác định được vị trí chèn.'); return; }
    var span = h('span', 'cite'); span.contentEditable = 'false'; span.dataset.keys = keys.join(','); span.textContent = '[?]';
    var s = window.getSelection(), r = s.getRangeAt(0); r.deleteContents(); r.insertNode(span);
    var sp = document.createTextNode(' '); span.after(sp); r.setStartAfter(sp); r.collapse(true); s.removeAllRanges(); s.addRange(r);
    st.rangeEl.dispatchEvent(new Event('input', { bubbles: true })); markDirty(); setTimeout(refreshCites, 50);
  }

  /* ---------- tài liệu tham khảo ---------- */
  function renderRefs() {
    var ch = cur(), ord = M.computeOrder(ch), box = $('#refsList'); if (!box || $('#refsPanel').hidden) return; box.innerHTML = '';
    var keys = ord.order.concat(Object.keys(ch.refs || {}).filter(function (k) { return !(k in ord.map); }));
    keys.forEach(function (k) {
      var it = h('div', 'refitem'), hd = h('div', 'rk'); hd.innerHTML = ord.map[k] ? '<span class="num">[' + ord.map[k] + ']</span>' : '<span class="badge">chưa được trích dẫn</span>';
      var del = h('button', null, 'Xóa'); del.onclick = function () { if (ord.map[k]) { toast('Nguồn đang được trích dẫn trong chương; hãy xóa trích dẫn trước.'); return; } delete ch.refs[k]; markDirty(); renderRefs(); }; hd.appendChild(del);
      var tx = h('div', 'rt', M.mdToHtml(ch.refs[k], {})); tx.contentEditable = PT; tx.dataset.k = k; tx.spellcheck = false; tx.title = 'Bấm để sửa. Dùng *chữ nghiêng* cho tên tài liệu.';
      tx.addEventListener('focus', function () { tx.textContent = ch.refs[k]; });
      tx.addEventListener('blur', function () { ch.refs[k] = tx.textContent.trim(); tx.innerHTML = M.mdToHtml(ch.refs[k], {}); });
      tx.addEventListener('input', function () { ch.refs[k] = tx.textContent.trim(); markDirty(); clearTimeout(st.citeT); st.citeT = setTimeout(function () { var rb = $('#doc .refsblk'); if (rb) rb.parentNode.replaceChild(refsEl(), rb); }, 500); });
      it.appendChild(hd); it.appendChild(tx); box.appendChild(it);
    });
    if (!keys.length) box.appendChild(h('p', 'hint', 'Chương chưa có nguồn nào.'));
  }
  function toggleRefs(show) {
    var p = $('#refsPanel'); p.hidden = !(show == null ? p.hidden : show); $('.layout').classList.toggle('with-refs', !p.hidden); renderRefs();
  }

  /* ---------- menu thêm nội dung ---------- */
  function openAddMenu(anchor, after) {
    var m = $('#addMenu'); st.insertAfter = after;
    var items = [['Văn bản', [['p', 'Đoạn văn'], ['h2', 'Tiêu đề mục (1.1)'], ['h3', 'Tiêu đề mục con (1.1.1)'], ['h4', 'Tiêu đề cấp 4 (a)']]],
      ['Danh sách', [['ul', 'Danh sách gạch đầu dòng'], ['ol', 'Danh sách đánh số']]],
      ['Hộp nhấn mạnh', [['nho', 'Ghi nhớ'], ['luuy', 'Lưu ý'], ['vidu', 'Ví dụ'], ['tinhhuong', 'Tình huống'], ['mo', 'Mở rộng']]],
      ['Khác', [['table', 'Bảng'], ['figure', 'Hình ảnh (tải ảnh lên)'], ['code', 'Khối mã lệnh'], ['quote', 'Trích dẫn mẫu / câu lệnh mẫu']]]];
    m.innerHTML = items.map(function (g) { return '<div class="grp">' + g[0] + '</div>' + g[1].map(function (x) { return '<button data-new="' + x[0] + '">' + x[1] + '</button>'; }).join(''); }).join('');
    var r = anchor.getBoundingClientRect(); m.hidden = false;
    m.style.left = Math.min(r.left, window.innerWidth - 260) + 'px'; m.style.top = Math.min(r.bottom + 4, window.innerHeight - Math.min(m.offsetHeight, 420) - 10) + 'px';
  }
  function hideMenus() { $('#addMenu').hidden = true; $('#moreMenu').hidden = true; }

  /* ---------- ảnh ---------- */
  function shrink(file) {
    return createImageBitmap(file).then(function (bmp) {
      var w = bmp.width, hh = bmp.height, max = 1400; if (w > max) { hh = Math.round(hh * max / w); w = max; }
      var c = document.createElement('canvas'); c.width = w; c.height = hh; var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, w, hh); x.drawImage(bmp, 0, 0, w, hh);
      return new Promise(function (res) { c.toBlob(res, 'image/jpeg', 0.88); });
    });
  }
  function onImageChosen() {
    var f = this.files[0]; this.value = ''; if (!f || st.figTarget < 0) return;
    if (!/^image\//.test(f.type)) { toast('Hãy chọn một tệp ảnh.'); return; }
    shrink(f).then(function (blob) {
      var id = 'u' + rid(); return S.put('images', id, blob).then(function () {
        st.urls[id] = URL.createObjectURL(blob); var b = cur().blocks[st.figTarget]; b.src = 'idb:' + id;
        if (!b.caption) b.caption = 'Chú thích hình'; markDirty(); renderDoc(true); toast('Đã thêm ảnh. Hãy sửa chú thích bên dưới ảnh.');
      });
    }).catch(function () { toast('Không đọc được ảnh này.'); });
  }

  /* ---------- xuất Word ---------- */
  function blobOf(src) {
    if (src.indexOf('idb:') === 0) return S.get('images', src.slice(4)).then(function (b) { if (!b) throw new Error('missing'); return b; });
    return fetch(src).then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.blob(); });
  }
  function loader(src) {
    return blobOf(src).then(function (blob) {
      return Promise.all([blob.arrayBuffer(), createImageBitmap(blob)]).then(function (x) { return { data: new Uint8Array(x[0]), width: x[1].width, height: x[1].height }; });
    });
  }
  function exportWord(nums) {
    var btns = $$('#btnExportOne,#btnExportAll'); btns.forEach(function (b) { b.disabled = true; });
    setStatus('Đang tạo file Word…', 'busy');
    var chapters = nums.map(function (n) { return st.ch[n]; });
    return E.build(window.docx, chapters, loader, { title: nums.length === 1 ? 'Chương ' + nums[0] : 'Giáo trình Phần mềm mã nguồn mở' })
      .then(function (doc) { return window.docx.Packer.toBlob(doc); })
      .then(function (blob) {
        var a = document.createElement('a'); a.href = URL.createObjectURL(blob);
        a.download = nums.length === 1 ? 'Chuong-' + nums[0] + '-Giao-trinh-PMNM.docx' : 'Giao-trinh-Phan-mem-ma-nguon-mo.docx';
        document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
        toast('Đã xuất file Word: ' + a.download); statusOk();
      })
      .catch(function (e) { console.error(e); setStatus('Lỗi khi xuất Word: ' + e.message, 'err'); toast('Không xuất được file Word.'); })
      .then(function () { btns.forEach(function (b) { b.disabled = false; }); });
  }

  /* ---------- đồng bộ GitHub (hàm serverless trên Vercel) ---------- */
  function api(body) {
    return fetch('/api/save', { method: 'POST', headers: { 'Content-Type': 'application/json', 'x-edit-password': sessionStorage.getItem('pmnm-pw') || '' }, body: JSON.stringify(body) })
      .then(function (r) { return r.json().then(function (j) { if (!r.ok) throw new Error(j.error || ('HTTP ' + r.status)); return j; }); });
  }
  function b64(blob) { return new Promise(function (res) { var fr = new FileReader(); fr.onload = function () { res(String(fr.result).split(',')[1]); }; fr.readAsDataURL(blob); }); }
  function pushChapter(n) {
    var ch = st.ch[n];
    var imgs = ch.blocks.filter(function (b) { return b.t === 'figure' && b.src && b.src.indexOf('idb:') === 0; });
    var chain = Promise.resolve();
    imgs.forEach(function (b) {
      chain = chain.then(function () { return S.get('images', b.src.slice(4)); }).then(function (blob) { return b64(blob); })
        .then(function (data) { return api({ action: 'image', id: b.src.slice(4), data: data }); }).then(function (r) { b.src = r.path; });
    });
    return chain.then(function () { return api({ action: 'chapter', number: n, data: ch }); }).then(function (r) { return S.put('chapters', n, JSON.parse(JSON.stringify(ch))).then(function () { return r; }); });
  }
  function openGithub() {
    modal('<h3>Lưu lên GitHub</h3><p>Chương đã chỉnh sửa sẽ được ghi vào kho GitHub của dự án (thư mục <code>content/</code>) và Vercel tự triển khai lại. Cần cấu hình biến môi trường trên Vercel (xem README).</p>' +
      '<p id="ghState" class="hint">Đang kiểm tra kết nối…</p><label for="ghPw">Mật khẩu chỉnh sửa (EDIT_PASSWORD)</label><input type="password" id="ghPw" autocomplete="current-password" value="' + esc(sessionStorage.getItem('pmnm-pw') || '') + '">' +
      '<div class="row"><button class="btn" id="ghClose">Đóng</button><button class="btn" id="ghAll">Lưu mọi chương đã sửa</button><button class="btn primary" id="ghOne">Lưu chương ' + st.cur + '</button></div><p id="ghMsg"></p>', function (m) {
        fetch('/api/save').then(function (r) { return r.json(); }).then(function (j) { $('#ghState', m).textContent = j.configured ? 'Đã kết nối kho GitHub (nhánh ' + j.branch + ').' : 'Máy chủ chưa được cấu hình GITHUB_TOKEN, GITHUB_REPO, EDIT_PASSWORD.'; })
          .catch(function () { $('#ghState', m).textContent = 'Không gọi được /api/save (chỉ hoạt động khi đã triển khai trên Vercel).'; });
        $('#ghClose', m).onclick = closeModal;
        function run(list) {
          sessionStorage.setItem('pmnm-pw', $('#ghPw', m).value); var msg = $('#ghMsg', m); msg.textContent = 'Đang lưu…';
          var chain = Promise.resolve(); list.forEach(function (n) { chain = chain.then(function () { msg.textContent = 'Đang lưu chương ' + n + '…'; return pushChapter(n); }); });
          chain.then(function () { msg.textContent = 'Đã lưu lên GitHub. Vercel sẽ triển khai lại trong ít phút.'; renderDoc(true); }).catch(function (e) { msg.textContent = 'Lỗi: ' + e.message; });
        }
        $('#ghOne', m).onclick = function () { run([st.cur]); };
        $('#ghAll', m).onclick = function () { var l = []; for (var n = 1; n <= NCH; n++) if (st.edited[n]) l.push(n); if (!l.length) { $('#ghMsg', m).textContent = 'Chưa có chương nào được chỉnh sửa.'; return; } run(l); };
      });
  }

  /* ---------- sao lưu / khôi phục ---------- */
  function backup() {
    var out = { version: 1, created: new Date().toISOString(), chapters: {}, images: {} }, ids = [];
    for (var n = 1; n <= NCH; n++) { out.chapters[n] = st.ch[n]; st.ch[n].blocks.forEach(function (b) { if (b.t === 'figure' && b.src && b.src.indexOf('idb:') === 0) ids.push(b.src.slice(4)); }); }
    var chain = Promise.resolve(); ids.forEach(function (id) { chain = chain.then(function () { return S.get('images', id); }).then(function (bl) { return bl ? b64(bl) : null; }).then(function (d) { if (d) out.images[id] = d; }); });
    chain.then(function () {
      var blob = new Blob([JSON.stringify(out)], { type: 'application/json' }), a = document.createElement('a'); a.href = URL.createObjectURL(blob);
      a.download = 'sao-luu-giao-trinh-pmnm-' + new Date().toISOString().slice(0, 10) + '.json'; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
      toast('Đã tải bản sao lưu.');
    });
  }
  function importBackup(file) {
    var fr = new FileReader();
    fr.onload = function () {
      try {
        var d = JSON.parse(fr.result); if (!d.chapters) throw new Error('Tệp không đúng định dạng');
        var chain = Promise.resolve();
        Object.keys(d.images || {}).forEach(function (id) { chain = chain.then(function () { return fetch('data:image/jpeg;base64,' + d.images[id]).then(function (r) { return r.blob(); }); }).then(function (bl) { return S.put('images', id, bl); }); });
        Object.keys(d.chapters).forEach(function (n) { chain = chain.then(function () { st.ch[n] = d.chapters[n]; st.edited[n] = true; return S.put('chapters', +n, d.chapters[n]); }); });
        chain.then(function () { st.urls = {}; renderDoc(); toast('Đã nhập bản sao lưu.'); });
      } catch (e) { toast('Không nhập được: ' + e.message); }
    };
    fr.readAsText(file);
  }
  function resetChapter() {
    if (!confirm('Khôi phục chương ' + st.cur + ' về bản gốc của website? Mọi chỉnh sửa trên trình duyệt đối với chương này sẽ mất.')) return;
    S.del('chapters', st.cur).then(function () { return fetchChapter(st.cur); }).then(function (d) { st.ch[st.cur] = d; delete st.edited[st.cur]; renderDoc(); toast('Đã khôi phục bản gốc.'); });
  }
  function help() {
    modal('<h3>Hướng dẫn nhanh</h3><p><b>Sửa nội dung:</b> bấm trực tiếp vào đoạn văn, tiêu đề, ô bảng hoặc chú thích để gõ. Mọi thay đổi tự lưu trên trình duyệt này.</p>' +
      '<p><b>Thêm nội dung:</b> di chuột vào một khối rồi bấm “+ Thêm nội dung bên dưới”, hoặc dùng thanh công cụ phía trên. Có thể thêm đoạn văn, tiêu đề, danh sách, hộp nhấn mạnh, bảng, khối mã và hình ảnh (tải ảnh từ máy).</p>' +
      '<p><b>Sắp xếp:</b> dùng ▲ ▼ ở góc khối để di chuyển, ⧉ để nhân đôi, ✕ để xóa.</p>' +
      '<p><b>Định dạng:</b> chọn chữ rồi bấm B, I hoặc &lt;/&gt;. Nút “Trích dẫn [n]” chèn số trích dẫn; số được đánh tự động theo thứ tự xuất hiện trong chương.</p>' +
      '<p><b>Xuất Word:</b> dùng nút “Xuất Word chương này” hoặc “Xuất cả giáo trình”. Định dạng theo mục “Định dạng Word” của rules.md (A4, lề 3/2/2/2 cm, Times New Roman 13, giãn dòng 1,3…).</p>' +
      '<p><b>Lưu lâu dài:</b> dữ liệu nằm trên trình duyệt; dùng “Lưu lên GitHub” hoặc “Tải bản sao lưu” để không mất nội dung.</p><div class="row"><button class="btn primary" id="hOk">Đã hiểu</button></div>', function (m) { $('#hOk', m).onclick = closeModal; });
  }

  /* ---------- khởi tạo ---------- */
  function bind() {
    var doc = $('#doc');
    doc.addEventListener('input', onInput); doc.addEventListener('change', onInput); doc.addEventListener('click', onClick);
    doc.addEventListener('keydown', onKeydown); doc.addEventListener('paste', onPaste);
    document.addEventListener('selectionchange', trackSelection);
    $('#fmtbar').addEventListener('mousedown', function (e) { if (e.target.closest('button')) e.preventDefault(); });
    $('#fmtbar').addEventListener('click', function (e) {
      var b = e.target.closest('button'); if (!b) return;
      if (b.dataset.cmd) fmt(b.dataset.cmd);
      else if (b.dataset.add === 'p') { var i = st.rangeEl ? +st.rangeEl.closest('.blk').dataset.i : cur().blocks.length - 1; insertBlock('p', i); }
      else if (b.dataset.add === 'menu') { var i2 = st.rangeEl && st.rangeEl.closest('.blk') ? +st.rangeEl.closest('.blk').dataset.i : cur().blocks.length - 1; openAddMenu(b, i2); }
    });
    $('#addMenu').addEventListener('click', function (e) { var b = e.target.closest('[data-new]'); if (!b) return; var a = st.insertAfter; hideMenus(); insertBlock(b.dataset.new, a); });
    $('#moreMenu').addEventListener('click', function (e) {
      var b = e.target.closest('[data-more]'); if (!b) return; hideMenus();
      ({ github: openGithub, backup: backup, import: function () { $('#fileImport').click(); }, reset: resetChapter, help: help })[b.dataset.more]();
    });
    document.addEventListener('mousedown', function (e) { if (!e.target.closest('.menu') && !e.target.closest('[data-act="ins"]') && !e.target.closest('[data-add="menu"]') && e.target.id !== 'btnMore') hideMenus(); });
    $('#btnMore').addEventListener('click', function () { var m = $('#moreMenu'), r = this.getBoundingClientRect(); m.hidden = !m.hidden; m.style.left = Math.max(8, r.right - 240) + 'px'; m.style.top = (r.bottom + 4) + 'px'; });
    $('#chapters').addEventListener('click', function (e) { var b = e.target.closest('button'); if (!b) return; st.cur = +b.dataset.n; renderDoc(); window.scrollTo(0, 0); $('#side').classList.remove('open'); });
    $('#outline').addEventListener('click', function (e) { var a = e.target.closest('a'); if (!a) return; e.preventDefault(); var s = $('#b' + a.dataset.i); if (s) s.scrollIntoView({ behavior: 'smooth', block: 'start' }); $('#side').classList.remove('open'); });
    $('#btnExportOne').addEventListener('click', function () { exportWord([st.cur]); });
    $('#btnExportAll').addEventListener('click', function () { exportWord([1, 2, 3, 4, 5, 6]); });
    $('#btnRefs').addEventListener('click', function () { toggleRefs(); });
    $('#btnRefsClose').addEventListener('click', function () { toggleRefs(false); });
    $('#btnRefAdd').addEventListener('click', function () { var k = 'r' + rid(); cur().refs[k] = ''; markDirty(); renderRefs(); var el = $('.rt[data-k="' + k + '"]'); if (el) el.focus(); });
    $('#btnSide').addEventListener('click', function () { $('#side').classList.toggle('open'); });
    $('#fileImage').addEventListener('change', onImageChosen);
    $('#fileImport').addEventListener('change', function () { if (this.files[0]) importBackup(this.files[0]); this.value = ''; });
    $('#modal').addEventListener('mousedown', function (e) { if (e.target.id === 'modal') closeModal(); });
    document.addEventListener('keydown', function (e) {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); saveNow(); toast('Đã lưu trên trình duyệt.'); }
      if (e.key === 'Escape') { hideMenus(); closeModal(); }
    });
    window.addEventListener('beforeunload', function () { if (st.saveT) saveNow(); });
  }

  function start() {
    bind();
    loadAll().then(function () { st.cur = 1; renderDoc(); }).catch(function (e) {
      var loc = location.protocol === 'file:';
      $('#doc').innerHTML = '<p class="hint">Không tải được nội dung chương' + (loc ? ': hãy mở trang qua máy chủ web (ví dụ chạy <code>npx serve .</code>) hoặc dùng bản đã triển khai trên Vercel.' : '.') + ' (' + esc(e.message) + ')</p>';
      setStatus('Không tải được nội dung', 'err');
    });
  }
  document.readyState === 'loading' ? document.addEventListener('DOMContentLoaded', start) : start();
})();
