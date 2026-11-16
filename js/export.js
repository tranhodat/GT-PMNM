/* Xuất Word theo mục "Định dạng Word" của rules.md:
   A4; lề trái 3 cm, phải/trên/dưới 2 cm; Times New Roman 13 pt; giãn dòng 1,3; sau đoạn 6 pt; căn đều;
   tiêu đề chương 15/20 pt; Heading 2/3/4; đầu trang/chân trang; bảng, hình, hộp nhấn mạnh theo quy định. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./markup.js'));
  else root.DocxExport = factory(root.Markup);
})(typeof self !== 'undefined' ? self : this, function (M) {
  'use strict';

  var CM = 567;                       // 1 cm = 567 twips
  var TEXT_W = 9071;                  // 16 cm
  var NAVY = '1F3864', BLUE = '2E75B6', GREY = '595959';
  var CALLOUT = {
    nho: ['GHI NHỚ', 'E8F0FA', '2E75B6', '1F4E8C'],
    luuy: ['LƯU Ý', 'FFF4DC', 'BF9000', '7F6000'],
    vidu: ['VÍ DỤ', 'E7F4EA', '2E8B57', '1E6B3F'],
    mo: ['MỞ RỘNG', 'F0F0F0', '7F7F7F', '404040'],
    tinhhuong: ['TÌNH HUỐNG', 'F4E9F5', '7B3F8C', '5B2A6B']
  };

  function build(D, chapters, loadImage, opts) {
    opts = opts || {};
    var FONT = 'Times New Roman';
    var AL = D.AlignmentType, BS = D.BorderStyle;

    /* ---- đoạn chữ có định dạng trong dòng ---- */
    function runs(text, map, o) {
      o = o || {};
      return M.parseInline(text, map).map(function (r) {
        var p = { text: r.t, font: FONT, size: o.size || 26 };
        if (o.color) p.color = o.color;
        if (o.bold || r.b) p.bold = true;
        if (o.italics || r.i) p.italics = true;
        if (r.c) { p.font = 'Consolas'; p.size = Math.max(18, (o.size || 26) - 3); p.shading = { type: D.ShadingType.CLEAR, fill: 'EFEFEF', color: 'auto' }; }
        return new D.TextRun(p);
      });
    }
    function para(children, o) {
      o = o || {};
      var p = { children: children, alignment: o.align || AL.JUSTIFIED, spacing: { line: o.line || 312, after: o.after == null ? 120 : o.after, before: o.before || 0 } };
      if (o.indent) p.indent = o.indent;
      if (o.keepNext) p.keepNext = true;
      if (o.keepLines) p.keepLines = true;
      if (o.heading) p.heading = o.heading;
      if (o.border) p.border = o.border;
      if (o.tabStops) p.tabStops = o.tabStops;
      return new D.Paragraph(p);
    }
    function border(color, size, space) { return { style: BS.SINGLE, size: size, color: color, space: space || 0 }; }
    function none() { return { style: BS.NONE, size: 0, color: 'FFFFFF' }; }
    function spacer() { return para([new D.TextRun({ text: '', size: 8 })], { after: 80, line: 200 }); }

    function listParas(items, ordered, map) {
      var keep = items.length <= 12;
      return items.map(function (it, i) {
        var mark = ordered ? (i + 1) + '.' : '•';
        var ch = [new D.TextRun({ children: [mark, new D.Tab()], bold: true, color: BLUE, font: FONT, size: 26 })].concat(runs(it, map));
        return para(ch, { align: AL.JUSTIFIED, after: 60, indent: { left: 567, hanging: 283 },
          tabStops: [{ type: D.TabStopType.LEFT, position: 567 }], keepNext: keep && i < items.length - 1 });
      });
    }

    function cellBorders(top, left, bottom, right) { return { top: top, left: left, bottom: bottom, right: right }; }

    function box(children, fill, borders, margins) {
      return new D.Table({
        width: { size: TEXT_W, type: D.WidthType.DXA }, columnWidths: [TEXT_W], layout: D.TableLayoutType.FIXED, alignment: AL.CENTER,
        rows: [new D.TableRow({ cantSplit: true, children: [new D.TableCell({
          width: { size: TEXT_W, type: D.WidthType.DXA },
          shading: { type: D.ShadingType.CLEAR, fill: fill, color: 'auto' },
          borders: borders, margins: margins || { top: 100, bottom: 100, left: 180, right: 160 }, children: children })] })]
      });
    }

    function calloutBlock(b, map) {
      var d = CALLOUT[b.kind] || CALLOUT.nho;
      var label = d[0] + (b.title ? ' – ' + b.title : '');
      var kids = [para([new D.TextRun({ text: label, bold: true, size: 23, color: d[3], font: FONT })], { align: AL.LEFT, after: 60, keepNext: true, line: 276 })];
      (b.texts || []).forEach(function (t) {
        if (t.indexOf('- ') === 0) {
          kids.push(para([new D.TextRun({ text: '•  ', bold: true, size: 24, font: FONT })].concat(runs(t.slice(2), map, { size: 24 })),
            { align: AL.LEFT, after: 60, line: 276, indent: { left: 340, hanging: 227 } }));
        } else kids.push(para(runs(t, map, { size: 24 }), { after: 60, line: 276 }));
      });
      var bd = cellBorders(border(d[2], 4), border(d[2], 36), border(d[2], 4), border(d[2], 4));
      return [box(kids, d[1], bd), spacer()];
    }

    function codeBlock(b) {
      var kids = [];
      if (b.title) kids.push(para([new D.TextRun({ text: b.title, bold: true, size: 21, color: GREY, font: FONT })], { align: AL.LEFT, after: 40, line: 240, keepNext: true }));
      String(b.text || '').split('\n').forEach(function (ln) {
        kids.push(para([new D.TextRun({ text: ln === '' ? ' ' : ln, font: 'Consolas', size: 21 })], { align: AL.LEFT, after: 0, line: 252, keepNext: true }));
      });
      var bd = cellBorders(border('D0D0D0', 4), border('7F7F7F', 24), border('D0D0D0', 4), border('D0D0D0', 4));
      return [box(kids, 'F3F3F3', bd, { top: 80, bottom: 80, left: 180, right: 140 }), spacer()];
    }

    function tableBlock(b, label, map) {
      var out = [];
      out.push(para([new D.TextRun({ text: label + ' ', bold: true, size: 24, color: NAVY, font: FONT })].concat(runs(b.caption || '', map, { size: 24, color: NAVY })),
        { align: AL.CENTER, keepNext: true, before: 160, after: 80, line: 276 }));
      var n = (b.header || []).length || 1;
      var w = (b.widths && b.widths.length === n) ? b.widths : (b.header || []).map(function () { return 16 / n; });
      var sum = w.reduce(function (a, c) { return a + c; }, 0);
      var tw = w.map(function (x) { return Math.round(x / sum * TEXT_W); });
      var size = Math.round((b.font || 12) * 2);
      var center = b.center || [];
      function cell(text, i, head, odd) {
        var kids = String(text).split('\n').map(function (ln) {
          var rr = head ? [new D.TextRun({ text: ln.replace(/\*\*/g, ''), bold: true, size: size, color: 'FFFFFF', font: FONT })]
                        : runs(ln, map, { size: size, bold: (i === 0 && b.firstBold !== false) });
          return para(rr, { align: (head || center.indexOf(i) >= 0) ? AL.CENTER : AL.LEFT, after: 0, line: 264 });
        });
        var bd = head ? cellBorders(border(NAVY, 6), border('FFFFFF', 4), border(NAVY, 6), border('FFFFFF', 4))
                      : cellBorders(border('BFC9D9', 4), border('BFC9D9', 4), border('BFC9D9', 4), border('BFC9D9', 4));
        return new D.TableCell({ width: { size: tw[i], type: D.WidthType.DXA },
          shading: { type: D.ShadingType.CLEAR, fill: head ? NAVY : (odd ? 'F2F6FB' : 'FFFFFF'), color: 'auto' },
          borders: bd, margins: { top: 80, bottom: 80, left: 140, right: 140 }, children: kids });
      }
      var rows = [new D.TableRow({ tableHeader: true, cantSplit: true, children: (b.header || []).map(function (h, i) { return cell(h, i, true); }) })];
      (b.rows || []).forEach(function (r, ri) {
        var cells = []; for (var i = 0; i < n; i++) cells.push(cell(r[i] == null ? '' : r[i], i, false, ri % 2 === 1));
        rows.push(new D.TableRow({ cantSplit: true, children: cells }));
      });
      out.push(new D.Table({ width: { size: TEXT_W, type: D.WidthType.DXA }, columnWidths: tw, layout: D.TableLayoutType.FIXED, alignment: AL.CENTER, rows: rows }));
      out.push(spacer());
      return out;
    }

    async function figureBlock(b, label, map) {
      var out = [], img = null;
      try { img = await loadImage(b.src); } catch (e) { img = null; }
      if (img) {
        var wcm = Math.min(b.width || 14.5, 16), wpx = Math.round(wcm * 37.795), hpx = Math.round(wpx * img.height / img.width);
        out.push(para([new D.ImageRun({ data: img.data, transformation: { width: wpx, height: hpx } })], { align: AL.CENTER, keepNext: true, before: 120, after: 40 }));
      } else {
        out.push(para([new D.TextRun({ text: '[Không tải được hình: ' + (b.src || '') + ']', italics: true, color: 'C00000', font: FONT, size: 22 })], { align: AL.CENTER, keepNext: true }));
      }
      out.push(para([new D.TextRun({ text: label + ' ', bold: true, size: 24, color: NAVY, font: FONT })].concat(runs(b.caption || '', map, { size: 24, color: NAVY })),
        { align: AL.CENTER, after: 160, line: 276 }));
      return out;
    }

    async function section(ch) {
      var ord = M.computeOrder(ch), map = ord.map, nums = M.numbering(ch), kids = [];
      kids.push(new D.Paragraph({
        heading: D.HeadingLevel.HEADING_1, spacing: { before: 0, after: 240, line: 276 }, alignment: AL.LEFT,
        border: { bottom: { style: BS.SINGLE, size: 18, color: NAVY, space: 4 } },
        children: [new D.TextRun({ text: 'CHƯƠNG ' + ch.number, bold: true, size: 30, color: BLUE, font: FONT }),
                   new D.TextRun({ text: String(ch.title || ch.titleText || '').toUpperCase(), bold: true, size: 40, color: NAVY, font: FONT, break: 1 })] }));
      var blocks = ch.blocks || [];
      for (var i = 0; i < blocks.length; i++) {
        var b = blocks[i];
        switch (b.t) {
          case 'h2': kids.push(para([new D.TextRun({ text: String(b.text).toUpperCase(), bold: true, size: 30, color: NAVY, font: FONT })],
              { heading: D.HeadingLevel.HEADING_2, align: AL.LEFT, keepNext: true, before: 360, after: 160, line: 276 })); break;
          case 'h3': kids.push(para([new D.TextRun({ text: b.text, bold: true, size: 27, color: BLUE, font: FONT })],
              { heading: D.HeadingLevel.HEADING_3, align: AL.LEFT, keepNext: true, before: 240, after: 120, line: 276 })); break;
          case 'h4': kids.push(para([new D.TextRun({ text: b.text, bold: true, italics: true, size: 26, color: '262626', font: FONT })],
              { heading: D.HeadingLevel.HEADING_4, align: AL.LEFT, keepNext: true, before: 160, after: 80, line: 276 })); break;
          case 'p': kids.push(para(runs(b.text, map), { align: b.left ? AL.LEFT : AL.JUSTIFIED, indent: (b.indent === false || b.left) ? undefined : { firstLine: 567 }, keepLines: false })); break;
          case 'list': kids = kids.concat(listParas(b.items || [], !!b.ordered, map)); break;
          case 'callout': kids = kids.concat(calloutBlock(b, map)); break;
          case 'code': kids = kids.concat(codeBlock(b)); break;
          case 'quote': kids.push(para(runs(b.text, map, { italics: true }), { align: AL.LEFT, indent: { left: 680, right: 340 }, after: 160, keepLines: true,
              border: { left: { style: BS.SINGLE, size: 18, color: '9DB7D5', space: 8 } } })); break;
          case 'table': kids = kids.concat(tableBlock(b, nums[i], map)); break;
          case 'figure': if (b.src) kids = kids.concat(await figureBlock(b, nums[i], map)); break;
          case 'refs':
            ord.order.forEach(function (k, j) {
              kids.push(para([new D.TextRun({ children: ['[' + (j + 1) + ']', new D.Tab()], bold: true, size: 24, font: FONT })].concat(runs(ch.refs[k], {}, { size: 24 })),
                { align: AL.LEFT, after: 100, line: 276, indent: { left: 567, hanging: 567 }, tabStops: [{ type: D.TabStopType.LEFT, position: 567 }] }));
            }); break;
          default: break;
        }
      }
      return {
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1701, right: 1134, header: 567, footer: 567 } } },
        headers: { default: new D.Header({ children: [new D.Paragraph({
          children: [new D.TextRun({ text: 'Giáo trình Phần mềm mã nguồn mở', italics: true, size: 21, color: GREY, font: FONT }),
                     new D.TextRun({ children: [new D.Tab(), 'Chương ' + ch.number], size: 21, color: GREY, font: FONT })],
          tabStops: [{ type: D.TabStopType.RIGHT, position: TEXT_W }],
          border: { bottom: { style: BS.SINGLE, size: 6, color: BLUE, space: 1 } } })] }) },
        footers: { default: new D.Footer({ children: [new D.Paragraph({ alignment: AL.CENTER,
          children: [new D.TextRun({ children: [D.PageNumber.CURRENT], size: 22, font: FONT })] })] }) },
        children: kids
      };
    }

    return (async function () {
      var secs = [];
      for (var i = 0; i < chapters.length; i++) secs.push(await section(chapters[i]));
      return new D.Document({
        creator: 'Giáo trình Phần mềm mã nguồn mở', title: opts.title || 'Giáo trình Phần mềm mã nguồn mở',
        styles: { default: {
          document: { run: { font: FONT, size: 26 }, paragraph: { spacing: { line: 312, after: 120 }, alignment: AL.JUSTIFIED } },
          heading1: { run: { font: FONT, size: 40, bold: true, color: NAVY }, paragraph: { spacing: { before: 0, after: 240 } } },
          heading2: { run: { font: FONT, size: 30, bold: true, color: NAVY }, paragraph: { spacing: { before: 360, after: 160 } } },
          heading3: { run: { font: FONT, size: 27, bold: true, color: BLUE }, paragraph: { spacing: { before: 240, after: 120 } } },
          heading4: { run: { font: FONT, size: 26, bold: true, italics: true, color: '262626' }, paragraph: { spacing: { before: 160, after: 80 } } }
        } },
        sections: secs
      });
    })();
  }

  return { build: build, CALLOUT: CALLOUT };
});
