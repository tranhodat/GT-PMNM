# -*- coding: utf-8 -*-
"""Chạy lại các tập lệnh dựng 6 chương nhưng ghi nội dung thành JSON có cấu trúc cho website."""
import sys, os, re, json, shutil, runpy, io
SCR = r'C:\Users\dat\AppData\Local\Temp\claude\6940572195de8611\46db4457-24e6-4970-9a38-646d9e25da88\scratchpad'
SITE = r'e:\VLUTE\BAI GIANG\PHAN MEM MA NGUON MO\GIAO TRINH PMNM\website'
sys.path.insert(0, SCR)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')
from PIL import Image
import gt

os.makedirs(SITE + r'\content', exist_ok=True)
os.makedirs(SITE + r'\assets\img', exist_ok=True)


class RecBook(gt.Book):
    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.blocks = []; self.meta = {'number': int(a[0]), 'title': a[1]}; self.nfig = 0

    def _rec(self, **d): self.blocks.append(d)

    def title_text(self, text):
        self.meta['titleText'] = text; super().title_text(text)

    def band(self, text): self._rec(t='h2', text=text, band=True); return super().band(text)
    def h2(self, t): self._rec(t='h2', text=t); return super().h2(t)
    def h3(self, t): self._rec(t='h3', text=t); return super().h3(t)
    def h4(self, t): self._rec(t='h4', text=t); return super().h4(t)

    def p(self, text, align=None, indent=True, max_bold=1):
        seen = [0]
        def kf(m):
            seen[0] += 1
            return m.group(0) if seen[0] <= max_bold else m.group(1)
        self._rec(t='p', text=re.sub(r'\*\*(.+?)\*\*', kf, text), indent=bool(indent), left=(align == 'left'))
        return super().p(text, align, indent, max_bold)

    def bullets(self, items, numbered=False, style_char='•', keep=True, maxkeep=8):
        self._rec(t='list', ordered=bool(numbered), items=list(items))
        return super().bullets(items, numbered, style_char, keep, maxkeep)

    def callout(self, kind, text, title=None):
        self._rec(t='callout', kind=kind, title=title or '', texts=list(text) if isinstance(text, list) else [text])
        return super().callout(kind, text, title)

    def code(self, text, title=None):
        self._rec(t='code', text=text, title=title or ''); return super().code(text, title)

    def quote(self, text): self._rec(t='quote', text=text); return super().quote(text)

    def table(self, caption, header, rows, widths, source=None, first_col_bold=True, font=12, align_center_cols=()):
        self._rec(t='table', caption=caption, header=list(header), rows=[list(r) for r in rows], widths=list(widths),
                  font=font, firstBold=bool(first_col_bold), center=list(align_center_cols))
        return super().table(caption, header, rows, widths, source, first_col_bold, font, align_center_cols)

    def figure(self, path, caption, source, width_cm=15.5):
        self.nfig += 1
        name = f'chuong{self.meta["number"]}-hinh{self.nfig}.jpg'
        im = Image.open(path)
        if im.mode in ('RGBA', 'LA', 'P'):
            im = im.convert('RGBA'); bg = Image.new('RGBA', im.size, (255, 255, 255, 255)); bg.alpha_composite(im); im = bg
        im = im.convert('RGB')
        if im.width > 1400: im = im.resize((1400, int(im.height * 1400 / im.width)))
        im.save(SITE + '\\assets\\img\\' + name, quality=88)
        self._rec(t='figure', src='assets/img/' + name, caption=caption, width=width_cm)
        return super().figure(path, caption, source, width_cm)

    def reference_list(self):
        self._rec(t='refs')
        self.refs = {k: self.REFS[k] for k in self.order}
        return super().reference_list()

    def save(self, path):
        data = dict(self.meta); data['blocks'] = self.blocks; data['refs'] = getattr(self, 'refs', {}); data['order'] = list(self.order)
        out = SITE + r'\content\chuong-%d.json' % self.meta['number']
        json.dump(data, open(out, 'w', encoding='utf8'), ensure_ascii=False, indent=1)
        print('JSON', out, len(self.blocks), 'blocks', len(data['refs']), 'refs')


gt.Book = RecBook

# Chương 1: tập lệnh riêng
runpy.run_path(SCR + r'\ch1.py', run_name='ch1')
# Chương 2, 3: bộ chuyển đổi conv
import conv, cfg2, cfg3
conv.Book = RecBook
conv.build(2, cfg2.CFG); conv.build(3, cfg3.CFG)
# Chương 4–6: conv2
import conv2, cfg4, cfg5, cfg6
conv2.Book = RecBook
conv2.build(4, cfg4.CFG); conv2.build(5, cfg5.CFG); conv2.build(6, cfg6.CFG)
