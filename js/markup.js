/* Xử lý định dạng trong dòng, trích dẫn [@khóa] và đánh số hình/bảng.
   Dùng chung cho trình soạn thảo (trình duyệt) và xuất Word (trình duyệt/Node). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Markup = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var CITE_RE = /\[@([A-Za-z0-9_,@ ]+)\]/g;
  var INLINE_RE = /(\*\*.+?\*\*|`.+?`|\*[^*\s][^*]*?\*)/;

  function keysOf(m) {
    return m.split(',').map(function (s) { return s.trim().replace(/^@/, ''); }).filter(Boolean);
  }

  /* Gom mọi chuỗi có thể chứa trích dẫn của một khối, theo thứ tự xuất hiện */
  function blockStrings(b) {
    var out = [];
    switch (b.t) {
      case 'p': case 'quote': out.push(b.text); break;
      case 'list': (b.items || []).forEach(function (x) { out.push(x); }); break;
      case 'callout': out.push(b.title || ''); (b.texts || []).forEach(function (x) { out.push(x); }); break;
      case 'table':
        out.push(b.caption || ''); (b.header || []).forEach(function (x) { out.push(x); });
        (b.rows || []).forEach(function (r) { r.forEach(function (x) { out.push(x); }); }); break;
      case 'figure': out.push(b.caption || ''); break;
      default: break;
    }
    return out;
  }

  /* Thứ tự xuất hiện đầu tiên của các khóa trích dẫn trong chương -> {khóa: số} */
  /* Danh mục tài liệu tham khảo chung cho cả giáo trình: {order:[khóa], refs:{khóa: chuỗi}}.
     Số [n] là vị trí của khóa trong order (cố định, thống nhất giữa các chương). */
  var GLOBAL = null;
  function setGlobalRefs(g) { GLOBAL = g; }
  function globalRefs() { return GLOBAL; }

  function computeOrder(chapter) {
    var order = [], map = {};
    if (GLOBAL && !chapter.refs) {
      var cited = {};
      GLOBAL.order.forEach(function (k, i) { map[k] = i + 1; });
      (chapter.blocks || []).forEach(function (b) {
        blockStrings(b).forEach(function (s) {
          var m; CITE_RE.lastIndex = 0;
          while ((m = CITE_RE.exec(s || ''))) keysOf(m[1]).forEach(function (k) { if (k in map) cited[k] = 1; });
        });
      });
      GLOBAL.order.forEach(function (k) { if (cited[k]) order.push(k); });
      return { order: order, map: map };
    }
    (chapter.blocks || []).forEach(function (b) {
      blockStrings(b).forEach(function (s) {
        var m; CITE_RE.lastIndex = 0;
        while ((m = CITE_RE.exec(s || ''))) {
          keysOf(m[1]).forEach(function (k) {
            if (!(k in map) && chapter.refs && (k in chapter.refs)) { order.push(k); map[k] = order.length; }
          });
        }
      });
    });
    return { order: order, map: map };
  }

  /* [@a,@b] -> [1], [2] (số nhỏ đến lớn); khóa lạ -> [?] */
  function resolveCites(text, map) {
    return String(text || '').replace(CITE_RE, function (_, g) {
      var nums = [];
      keysOf(g).forEach(function (k) { var n = map[k]; nums.push(n ? n : '?'); });
      nums.sort(function (a, b) { return (a === '?') - (b === '?') || a - b; });
      var uniq = nums.filter(function (v, i) { return nums.indexOf(v) === i; });
      return uniq.map(function (n) { return '[' + n + ']'; }).join(', ');
    });
  }

  /* Tách chuỗi thành các đoạn định dạng {t, b, i, c} sau khi đã đổi trích dẫn */
  function parseInline(text, map) {
    var s = resolveCites(text, map || {});
    var parts = s.split(INLINE_RE), out = [];
    parts.forEach(function (t) {
      if (!t) return;
      if (t.length > 4 && t.slice(0, 2) === '**' && t.slice(-2) === '**') out.push({ t: t.slice(2, -2), b: true });
      else if (t.length > 2 && t[0] === '`' && t.slice(-1) === '`') out.push({ t: t.slice(1, -1), c: true });
      else if (t.length > 2 && t[0] === '*' && t.slice(-1) === '*') out.push({ t: t.slice(1, -1), i: true });
      else out.push({ t: t });
    });
    return out;
  }

  /* Đánh số Hình n.m, Bảng n.m theo thứ tự trong chương */
  function numbering(chapter) {
    var f = 0, t = 0, res = [];
    (chapter.blocks || []).forEach(function (b) {
      if (b.t === 'figure') { if (b.src) { f++; res.push('Hình ' + chapter.number + '.' + f + '.'); } else res.push(''); }
      else if (b.t === 'table') { t++; res.push('Bảng ' + chapter.number + '.' + t + '.'); }
      else res.push('');
    });
    return res;
  }

  /* ---------- Phần chỉ chạy trong trình duyệt (DOM) ---------- */
  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

  function mdToHtml(text, map) {
    var s = String(text || '');
    var parts = s.split(/(\[@[A-Za-z0-9_,@ ]+\]|\*\*.+?\*\*|`.+?`|\*[^*\s][^*]*?\*)/);
    return parts.map(function (t) {
      if (!t) return '';
      if (/^\[@/.test(t)) {
        var ks = keysOf(t.slice(2, -1)).filter(Boolean);
        var lab = resolveCites(t, map || {});
        return '<span class="cite" contenteditable="false" data-keys="' + esc(ks.join(',')) + '">' + esc(lab) + '</span>';
      }
      if (t.length > 4 && t.slice(0, 2) === '**' && t.slice(-2) === '**') return '<b>' + esc(t.slice(2, -2)) + '</b>';
      if (t.length > 2 && t[0] === '`' && t.slice(-1) === '`') return '<code>' + esc(t.slice(1, -1)) + '</code>';
      if (t.length > 2 && t[0] === '*' && t.slice(-1) === '*') return '<i>' + esc(t.slice(1, -1)) + '</i>';
      return esc(t);
    }).join('');
  }

  function htmlToMd(el) {
    var out = '';
    el.childNodes.forEach(function (n) {
      if (n.nodeType === 3) { out += n.nodeValue.replace(/ /g, ' '); return; }
      if (n.nodeType !== 1) return;
      var tag = n.tagName.toLowerCase(), inner;
      if (tag === 'span' && n.classList.contains('cite')) { out += '[@' + (n.getAttribute('data-keys') || '').split(',').join(',@') + ']'; return; }
      inner = htmlToMd(n);
      if (tag === 'b' || tag === 'strong') out += inner.trim() ? '**' + inner + '**' : inner;
      else if (tag === 'i' || tag === 'em') out += inner.trim() ? '*' + inner + '*' : inner;
      else if (tag === 'code') out += inner.trim() ? '`' + inner.replace(/`/g, '') + '`' : inner;
      else if (tag === 'br') out += ' ';
      else if (tag === 'div' || tag === 'p') out += (out && !/\s$/.test(out) ? ' ' : '') + inner;
      else out += inner;
    });
    return out;
  }

  return { setGlobalRefs: setGlobalRefs, globalRefs: globalRefs, computeOrder: computeOrder, resolveCites: resolveCites, parseInline: parseInline, numbering: numbering,
           mdToHtml: mdToHtml, htmlToMd: htmlToMd, blockStrings: blockStrings, keysOf: keysOf, esc: esc };
});
