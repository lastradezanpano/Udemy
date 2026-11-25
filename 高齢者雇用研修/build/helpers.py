"""高齢者雇用研修スライド生成用の共通部品（python-pptx）"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

FONT = "メイリオ"
W, H = 13.333, 7.5

# 配色：濃いティール（物流・安全の信頼感）＋オレンジ（注意喚起）
DARK = RGBColor(0x14, 0x3D, 0x4A)
TEAL = RGBColor(0x1F, 0x6F, 0x7F)
TINT = RGBColor(0xE8, 0xF1, 0xF3)
ACC = RGBColor(0xE0, 0x7A, 0x1F)
ACC_T = RGBColor(0xFC, 0xEE, 0xDF)
TXT = RGBColor(0x22, 0x2B, 0x30)
MUTED = RGBColor(0x5E, 0x6B, 0x72)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NG = RGBColor(0xB0, 0x3A, 0x2E)
NG_T = RGBColor(0xF8, 0xE5, 0xE2)
OK = RGBColor(0x2E, 0x7D, 0x4F)
OK_T = RGBColor(0xE3, 0xF2, 0xE8)


# ---- pptx-font-autofit スキルの近似計算 ----
def _units(text):
    return sum(1.0 if ord(c) < 128 else 2.0 for c in text)


def fit_font_size(text, w_pt, h_pt, max_size=24, min_size=14, step=1,
                  line_spacing=1.3, ratio=0.55):
    for size in range(max_size, min_size - 1, -step):
        per_line = max(1.0, (w_pt - 14) / (size * ratio))
        lines = 0
        for para in text.split("\n"):
            lines += max(1, -(-int(_units(para)) // int(per_line)))
        if lines * size * line_spacing <= h_pt - 8:
            return size
    return min_size


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(W)
        self.prs.slide_height = Inches(H)
        self.blank = self.prs.slide_layouts[6]
        self.n = 0

    # ---------- 基本部品 ----------
    def slide(self, notes=""):
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        if notes:
            s.notes_slide.notes_text_frame.text = notes.strip()
        return s

    def text(self, s, x, y, w, h, text, size=None, color=TXT, bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, max_size=24,
             min_size=14, spacing=1.25):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        for m in ("margin_left", "margin_right"):
            setattr(tf, m, Inches(0.05))
        tf.margin_top = tf.margin_bottom = Inches(0.03)
        if size is None:
            size = fit_font_size(text, w * 72, h * 72, max_size, min_size,
                                 line_spacing=spacing)
        for i, line in enumerate(text.split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = Pt(round(size * spacing))  # 行間をpt固定（メイリオの行高に左右されない）
            r = p.add_run()
            r.text = line
            f = r.font
            f.name = FONT
            f.size = Pt(size)
            f.bold = bold
            f.color.rgb = color
            rPr = r._r.get_or_add_rPr()
            for tag in ("a:ea",):
                from pptx.oxml.ns import qn
                el = rPr.find(qn(tag))
                if el is None:
                    el = rPr.makeelement(qn(tag), {})
                    rPr.append(el)
                el.set("typeface", FONT)
        return tb

    def box(self, s, x, y, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
            line=None, radius=0.08):
        sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = line
            sh.line.width = Pt(1.5)
        sh.shadow.inherit = False
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            sh.adjustments[0] = radius
        return sh

    def circle_num(self, s, x, y, d, label, fill=TEAL, color=WHITE, size=20):
        c = self.box(s, x, y, d, d, fill, MSO_SHAPE.OVAL)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = label
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = True
        r.font.color.rgb = color
        return c

    def header(self, s, title, kicker=None):
        """白背景の内容スライドの見出し（全スライド同位置）"""
        if kicker:
            self.text(s, 0.6, 0.35, 9, 0.4, kicker, size=14, color=ACC, bold=True)
        self.text(s, 0.6, 0.7, 12.1, 0.8, title, size=30, color=DARK, bold=True)
        self.text(s, 12.1, 7.0, 0.8, 0.35, str(self.n), size=11, color=MUTED,
                  align=PP_ALIGN.RIGHT)

    def source(self, s, text):
        self.text(s, 0.6, 6.92, 11.3, 0.45, text, size=11, color=MUTED)

    # ---------- スライド型 ----------
    def title_slide(self, title, sub, notes):
        s = self.slide(notes)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = DARK
        self.circle_num(s, 0.8, 1.2, 0.9, "研修", fill=ACC, size=16)
        self.text(s, 0.8, 2.4, 11.5, 2.0, title, size=40, color=WHITE, bold=True,
                  spacing=1.2)
        self.text(s, 0.8, 4.5, 11.5, 1.2, sub, size=20, color=RGBColor(0xCF, 0xE3, 0xE7))
        return s

    def section_slide(self, no, title, lead, notes):
        s = self.slide(notes)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = DARK
        self.circle_num(s, 0.8, 2.0, 1.3, no, fill=ACC, size=32)
        self.text(s, 2.4, 1.85, 10.3, 1.9, title, size=34, color=WHITE, bold=True,
                  spacing=1.3)
        self.text(s, 2.4, 4.0, 10.3, 1.4, lead, size=20, color=RGBColor(0xCF, 0xE3, 0xE7))
        self.text(s, 12.1, 7.0, 0.8, 0.35, str(self.n), size=11,
                  color=RGBColor(0xCF, 0xE3, 0xE7), align=PP_ALIGN.RIGHT)
        return s

    def cards(self, s, items, top=1.75, bottom=6.75, cols=None, numbered=True,
              fill=TINT, head_color=DARK, body_max=24):
        """items: [(見出し, 本文)]"""
        n = len(items)
        cols = cols or (n if n <= 4 else (3 if n in (5, 6) else 4))
        rows = -(-n // cols)
        gap = 0.3
        cw = (W - 1.2 - gap * (cols - 1)) / cols
        ch = (bottom - top - gap * (rows - 1)) / rows
        for i, (head, body) in enumerate(items):
            r, c = divmod(i, cols)
            x = 0.6 + c * (cw + gap)
            y = top + r * (ch + gap)
            self.box(s, x, y, cw, ch, fill)
            hx = x + 0.25
            if numbered:
                self.circle_num(s, x + 0.25, y + 0.25, 0.55, str(i + 1), size=18)
                hx = x + 0.95
            self.text(s, hx, y + 0.22, x + cw - hx - 0.15, 0.65, head, color=head_color,
                      bold=True, max_size=21, min_size=16, anchor=MSO_ANCHOR.MIDDLE)
            self.text(s, x + 0.25, y + 0.95, cw - 0.45, ch - 1.05, body,
                      max_size=body_max, min_size=15)
        return s

    def compare(self, s, left_head, left_items, right_head, right_items,
                top=1.75, bottom=6.75, lfill=NG_T, lcol=NG, rfill=OK_T, rcol=OK,
                lmark="×", rmark="○"):
        cw = (W - 1.2 - 0.4) / 2
        for k, (head, items, fill, col, mark) in enumerate(
                [(left_head, left_items, lfill, lcol, lmark),
                 (right_head, right_items, rfill, rcol, rmark)]):
            x = 0.6 + k * (cw + 0.4)
            self.box(s, x, top, cw, bottom - top, fill)
            self.circle_num(s, x + 0.25, top + 0.22, 0.6, mark, fill=col, size=22)
            self.text(s, x + 1.0, top + 0.22, cw - 1.2, 0.6, head, size=21, color=col,
                      bold=True, anchor=MSO_ANCHOR.MIDDLE)
            body = "\n".join("・" + t for t in items)
            self.text(s, x + 0.3, top + 1.0, cw - 0.55, bottom - top - 1.15, body,
                      max_size=22, min_size=15, spacing=1.4)
        return s

    def steps(self, s, items, top=1.9, h=3.4, note=None):
        """横向きの手順フロー items: [(見出し, 本文)]"""
        n = len(items)
        gap = 0.45
        cw = (W - 1.2 - gap * (n - 1)) / n
        for i, (head, body) in enumerate(items):
            x = 0.6 + i * (cw + gap)
            self.box(s, x, top, cw, h, TINT)
            self.circle_num(s, x + cw / 2 - 0.35, top + 0.25, 0.7, str(i + 1),
                            fill=ACC if i == n - 1 else TEAL, size=22)
            self.text(s, x + 0.15, top + 1.1, cw - 0.3, 0.7, head, color=DARK, bold=True,
                      align=PP_ALIGN.CENTER, max_size=21, min_size=16)
            self.text(s, x + 0.2, top + 1.85, cw - 0.4, h - 1.95, body, max_size=20,
                      min_size=15)
            if i < n - 1:
                a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + cw + 0.07),
                                       Inches(top + h / 2 - 0.2), Inches(0.31),
                                       Inches(0.4))
                a.fill.solid()
                a.fill.fore_color.rgb = ACC
                a.line.fill.background()
        if note:
            self.box(s, 0.6, top + h + 0.3, W - 1.2, 6.75 - top - h - 0.3, ACC_T)
            self.text(s, 0.85, top + h + 0.38, W - 1.7, 6.75 - top - h - 0.46, note,
                      max_size=20, min_size=14, anchor=MSO_ANCHOR.MIDDLE)
        return s

    def stats(self, s, items, top=1.8, h=2.6):
        """大きな数字 items: [(数値, 単位, 説明)]"""
        n = len(items)
        gap = 0.3
        cw = (W - 1.2 - gap * (n - 1)) / n
        for i, (num, unit, label) in enumerate(items):
            x = 0.6 + i * (cw + gap)
            self.box(s, x, top, cw, h, TINT)
            tb = s.shapes.add_textbox(Inches(x + 0.1), Inches(top + 0.2),
                                      Inches(cw - 0.2), Inches(1.2))
            tf = tb.text_frame
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.line_spacing = Pt(60)
            for t, sz, col in ((num, 50, ACC), (unit, 22, DARK)):
                r = p.add_run()
                r.text = t
                r.font.name = FONT
                r.font.size = Pt(sz)
                r.font.bold = True
                r.font.color.rgb = col
            self.text(s, x + 0.2, top + 1.4, cw - 0.4, h - 1.5, label,
                      align=PP_ALIGN.CENTER, max_size=20, min_size=15)
        return s

    def table(self, s, rows, col_w, top=1.75, row_h=0.62, head_fill=DARK,
              size=16):
        nr, nc = len(rows), len(rows[0])
        gt = s.shapes.add_table(nr, nc, Inches(0.6), Inches(top),
                                Inches(sum(col_w)), Inches(row_h * nr))
        t = gt.table
        for j, w in enumerate(col_w):
            t.columns[j].width = Inches(w)
        for i, row in enumerate(rows):
            t.rows[i].height = Inches(row_h)
            for j, val in enumerate(row):
                cell = t.cell(i, j)
                cell.fill.solid()
                cell.fill.fore_color.rgb = head_fill if i == 0 else (
                    TINT if i % 2 == 0 else WHITE)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = cell.margin_right = Inches(0.12)
                tf = cell.text_frame
                tf.word_wrap = True
                for k, line in enumerate(val.split("\n")):
                    p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                    p.line_spacing = Pt(round(size * 1.25))
                    r = p.add_run()
                    r.text = line
                    r.font.name = FONT
                    r.font.size = Pt(size)
                    r.font.bold = (i == 0 or j == 0)
                    r.font.color.rgb = WHITE if i == 0 else (DARK if j == 0 else TXT)
        return gt

    def callout(self, s, text, y=5.75, h=1.0, fill=ACC_T):
        self.box(s, 0.6, y, W - 1.2, h, fill)
        self.circle_num(s, 0.8, y + h / 2 - 0.3, 0.6, "!", fill=ACC, size=22)
        self.text(s, 1.6, y + 0.08, W - 2.4, h - 0.16, text, max_size=20, min_size=14,
                  anchor=MSO_ANCHOR.MIDDLE)

    def save(self, path):
        self.prs.save(path)
