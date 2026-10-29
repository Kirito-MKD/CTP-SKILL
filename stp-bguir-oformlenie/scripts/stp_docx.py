"""Вёрстка docx по СТП 01-2024 (БГУИР) на python-docx, с учётом замечаний преподавателя.

    from stp_docx import StpDoc
    d = StpDoc()
    d.title_lab(faculty=..., department=..., work_no=5, topic=..., discipline=...,
                student="И.О. Фамилия", group="461403", teacher="И.О. Фамилия", teacher_label="Проверила")
    d.unnumbered_heading("ЦЕЛЬ РАБОТЫ", new_page=False)
    d.text("Изучить ...")
    d.heading("1", "Исходные данные")
    ...
    d.save("Отчёт.docx")

Проверки:  python stp_docx.py check Отчёт.docx [Отчёт.pdf]
Пример:    python stp_docx.py demo demo.docx
"""
import os
import re
import subprocess
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

NBSP = "\u00a0"
DASH = "\u2013"
ZWSP = "\u200b"
DASH_SPACING = "70"         # межзнаковый интервал после тире, 1/20 пт: 3,5 пт = ширина пробела Times 14 пт
TEXT_W = 16.5               # ширина текста, см: 21 − 3 − 1,5
TITLE_LINES = 40            # строк по 18 пт на полезной высоте листа (257 мм)
APPENDIX_LETTERS = "АБВГДЕЖИКЛМНПРСТУФХЦШЩЭЮЯ"    # без Ё, З, Й, О, Ч, Ъ, Ы, Ь
UNNUMBERED = ("СОДЕРЖАНИЕ", "РЕФЕРАТ", "ПЕРЕЧЕНЬ УСЛОВНЫХ ОБОЗНАЧЕНИЙ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ",
              "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "ЦЕЛЬ РАБОТЫ", "ВЫВОД", "ВЫВОДЫ")

L, C, R, J = (WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER,
              WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.JUSTIFY)
ALIGN = {"l": L, "c": C, "r": R, "j": J}

# « – » внутри текста и дефис между заглавными/цифрами в обозначениях (C9200L-24P, 1000BASE-T)
_TOKEN = re.compile(r"(?<=\S) " + DASH + r" |(?<=[A-Z0-9])-(?=[A-Z0-9])")
_FONTS = {}


def text_pt(s, bold=False, size=14):
    """Ширина строки в Times New Roman, пт (по метрикам шрифта Windows; без него — оценка)."""
    if bold not in _FONTS:
        try:
            from PIL import ImageFont
            name = "timesbd.ttf" if bold else "times.ttf"
            _FONTS[bold] = ImageFont.truetype(os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", name), 1000)
        except Exception:
            _FONTS[bold] = None
    f = _FONTS[bold]
    return f.getlength(s) * size / 1000 if f else len(s) * size * (0.55 if bold else 0.5)


def _lines(s, bold=False, size=14, width_cm=TEXT_W):
    return max(1, -(-int(text_pt(s, bold, size)) // int(width_cm / 2.54 * 72)))


class StpDoc:
    def __init__(self, by_section=True):
        """by_section=True — рисунки, таблицы, формулы нумеруются по разделам (2.1); False — сквозная (1, 2, 3)."""
        self.doc = Document()
        st = self.doc.styles["Normal"]
        st.font.name = "Times New Roman"
        fonts = st.element.get_or_add_rPr().get_or_add_rFonts()
        for a in ("w:ascii", "w:eastAsia", "w:hAnsi", "w:cs"):
            fonts.set(qn(a), "Times New Roman")
        st.font.size = Pt(14)
        pf = st.paragraph_format
        pf.line_spacing = Pt(18)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.first_line_indent = Cm(1.25)
        pf.alignment = J
        pf.space_before = pf.space_after = Pt(0)
        pf.widow_control = True
        self._page(self.doc.sections[0])
        self.by_section = by_section
        self.chapter = None
        self.count = {"fig": 0, "tab": 0, "form": 0}
        self.last = None
        self.pending_break = False
        self._title_used = 0

    # ------------------------------------------------------------------ низкий уровень
    @staticmethod
    def _page(sec):
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        sec.top_margin = sec.bottom_margin = Cm(2)
        sec.left_margin, sec.right_margin = Cm(3), Cm(1.5)

    def _par(self, align=J, indent=Cm(1.25)):
        p = self.doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.first_line_indent = indent
        if self.pending_break:
            p.paragraph_format.page_break_before = True
            self.pending_break = False
        return p

    @staticmethod
    def _fmt(r, bold=False, italic=False, size=None):
        if bold:
            r.bold = True
        if italic:
            r.italic = True
        if size:
            r.font.size = Pt(size)

    def _runs(self, p, text, bold=False, italic=False, size=None):
        """Текст абзаца с типографикой из замечаний:
        «– пункт» — неразрывный пробел после тире (одинаковый зазор во всех пунктах);
        «слово – слово» — неразрывный пробел перед тире, пробела после нет, интервал даёт разрядка тире
        (обычный пробел при выключке по ширине растягивается); дефис в обозначениях — неразрывный."""
        text = str(text)
        if text.startswith(DASH + " "):
            text = DASH + NBSP + text[2:]
        pos = 0
        for m in _TOKEN.finditer(text):
            dash = m.group() != "-"
            head = text[pos:m.start()] + (NBSP if dash else "")
            if head:
                self._fmt(p.add_run(head), bold, italic, size)
            r = p.add_run()
            if dash:
                sp = OxmlElement("w:spacing")
                sp.set(qn("w:val"), DASH_SPACING)
                r._r.get_or_add_rPr().append(sp)      # до b/i/sz: python-docx вставит их в порядке схемы
                r.add_text(DASH)
            else:
                r._r.append(OxmlElement("w:noBreakHyphen"))
            self._fmt(r, bold, italic, size)
            pos = m.end()
        if pos < len(text):
            self._fmt(p.add_run(text[pos:]), bold, italic, size)
        return p

    def _next(self, kind):
        self.count[kind] += 1
        n = self.count[kind]
        if self.chapter and (self.by_section or self.chapter.isalpha()):
            return f"{self.chapter}.{n}"
        return str(n)

    # ------------------------------------------------------------------ текст
    def blank(self, keep=False):
        """Пробельная строка: одна (две подряд не ставятся), высотой 12 пт, не в начале страницы."""
        if self.last == "blank" or self.pending_break:
            return None
        p = self._par(indent=0)
        p.paragraph_format.line_spacing = Pt(12)
        p.paragraph_format.keep_with_next = keep
        self.last = "blank"
        return p

    def text(self, s, keep=False, align=J):
        p = self._runs(self._par(align), s)
        p.paragraph_format.keep_with_next = keep
        self.last = "text"
        return p

    def items(self, lines):
        """Простое перечисление: «– пункт;», у последнего «.»."""
        for i, s in enumerate(lines):
            s = s.rstrip(" ;.")
            self.text(DASH + " " + s + ("." if i == len(lines) - 1 else ";"), keep=(i == 0))

    def numbered(self, lines):
        """Сложное перечисление (пункты — предложения): «1 Текст.»."""
        for i, s in enumerate(lines, 1):
            s = s.rstrip()
            self.text(f"{i} {s[0].upper() + s[1:]}" + ("" if s.endswith(".") else "."))

    def code(self, lines, size=10):
        """Команды/конфигурация: моноширинный шрифт, без выключки, отступ 1,25 см."""
        for i, s in enumerate(lines):
            p = self._par(L, 0)
            pf = p.paragraph_format
            pf.left_indent = Cm(1.25)
            pf.line_spacing = Pt(size * 1.25)
            pf.keep_with_next = i < len(lines) - 1
            r = p.add_run(s)
            r.font.size = Pt(size)
            r.font.name = "Courier New"
            r._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:hAnsi"), "Courier New")
        self.last = "text"

    # ------------------------------------------------------------------ заголовки
    def unnumbered_heading(self, title, new_page=True):
        """ЦЕЛЬ РАБОТЫ, ВВЕДЕНИЕ, ВЫВОД, ЗАКЛЮЧЕНИЕ… — по центру, без отступа, полужирным, прописными."""
        if new_page:
            self.pending_break = True
        elif self.last not in (None, "blank"):
            self.blank(keep=True)
        p = self._runs(self._par(C, 0), title.upper(), bold=True)
        p.paragraph_format.keep_with_next = True
        self.last = "heading"
        self.blank(keep=True)
        self.last = "heading"
        self.chapter = None
        return p

    def heading(self, num, title, new_page=None):
        """«1 РАЗДЕЛ», «1.1 Подраздел», «1.1.1 Пункт»: с абзацного отступа 1,25 см, после номера табуляция,
        перенос названия — под первой буквой названия (2,5 см). Пустая строка снизу; у подзаголовка и сверху,
        кроме случая сразу после заголовка раздела. new_page по умолчанию — только у раздела."""
        level = num.count(".") + 1
        if new_page is None:
            new_page = level == 1
        if new_page:
            self.pending_break = True
        elif self.last not in (None, "blank", "heading"):
            self.blank(keep=True)
        p = self._par(L, Cm(-1.25))
        pf = p.paragraph_format
        pf.left_indent = Cm(2.5)
        pf.tab_stops.add_tab_stop(Cm(2.5))
        pf.keep_with_next = True
        self._runs(p, f"{num}\t{title.upper() if level == 1 else title}", bold=True)
        self.last = "heading"
        self.blank(keep=True)
        self.last = "heading"
        if level == 1:
            self.chapter = num
            self.count = dict.fromkeys(self.count, 0)
        return p

    # ------------------------------------------------------------------ рисунки, таблицы, формулы
    def figure(self, path, caption, width_cm=15.0, max_h_cm=21.0, legend=None):
        """Рисунок по центру с одинарным интервалом (иначе обрежется), подпись «Рисунок N – Название» без точки.
        Возвращает номер — сошлись на него в тексте ДО вызова."""
        num = self._next("fig")
        self.blank(keep=True)
        p = self._par(C, 0)
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.keep_with_next = True
        try:
            from PIL import Image
            w, h = Image.open(path).size
            width_cm = min(width_cm, TEXT_W, max_h_cm * w / h)
        except Exception:
            pass
        p.add_run().add_picture(path, width=Cm(width_cm))
        self.last = "picture"
        self.blank(keep=True)
        if legend:
            q = self._runs(self._par(C, 0), legend)
            q.paragraph_format.keep_with_next = True
        cap = self._runs(self._par(C, 0), f"Рисунок {num} – {caption.rstrip('.')}")
        cap.paragraph_format.keep_together = True
        self.last = "caption"
        self.blank()
        return num

    def table(self, caption, header, rows, widths_cm, aligns=None, size=12, split_at=()):
        """Таблица: заголовок по левому краю (перенос — под началом названия), шапка граф повторяется,
        строки не разрываются, пустые ячейки — «–». split_at — номера строк данных, с которых начинается
        следующая часть: она идёт с новой страницы под «Продолжение таблицы N» с повтором шапки."""
        num = self._next("tab")
        aligns = aligns or "l" + "c" * (len(header) - 1)
        bounds = [0, *split_at, len(rows)]
        for i in range(len(bounds) - 1):
            if i == 0:
                self.blank(keep=True)
                title = f"Таблица {num} – {caption.rstrip('.')}"
            else:
                self.pending_break = True
                title = f"Продолжение таблицы {num}"
            t = self._par(L, 0)
            if i == 0:
                hang = Pt(text_pt(f"Таблица {num} {DASH} "))
                t.paragraph_format.left_indent = hang
                t.paragraph_format.first_line_indent = -hang
            t.paragraph_format.keep_with_next = True
            self._runs(t, title)
            self._table_body(header, rows[bounds[i]:bounds[i + 1]], widths_cm, aligns, size)
        self.last = "table"
        self.blank()
        return num

    def _table_body(self, header, rows, widths_cm, aligns, size):
        tbl = self.doc.add_table(rows=1 + len(rows), cols=len(header))
        tbl.style = self.doc.styles["Table Grid"]
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        for gc, w in zip(tbl._tbl.tblGrid.findall(qn("w:gridCol")), widths_cm):
            gc.set(qn("w:w"), str(round(w / 2.54 * 1440)))
        data = [header] + rows
        for ri, cells in enumerate(data):
            row = tbl.rows[ri]
            trpr = row._tr.get_or_add_trPr()
            trpr.append(OxmlElement("w:cantSplit"))
            if ri == 0:
                trpr.append(OxmlElement("w:tblHeader"))
            for ci, val in enumerate(cells):
                cell = row.cells[ci]
                cell.width = Cm(widths_cm[ci])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p = cell.paragraphs[0]
                pf = p.paragraph_format
                pf.first_line_indent = 0
                pf.line_spacing = Pt(size + 2)
                pf.keep_with_next = ri < len(data) - 1
                p.alignment = C if ri == 0 else ALIGN[aligns[ci]]
                val = str(val).strip() or DASH
                self._runs(p, val, size=size)

    def formula(self, expr, where=(), num=True):
        """Формула по центру, номер «(N)» у правого края, пустая строка сверху и снизу; затем «где» +
        табуляция, символы друг под другом: where = [("P", "мощность, Вт"), ...]."""
        n = self._next("form") if num else ""
        self.blank(keep=True)
        p = self._par(L, 0)
        pf = p.paragraph_format
        pf.keep_with_next = True
        pf.tab_stops.add_tab_stop(Cm(TEXT_W / 2), WD_TAB_ALIGNMENT.CENTER)
        pf.tab_stops.add_tab_stop(Cm(TEXT_W), WD_TAB_ALIGNMENT.RIGHT)
        self._runs(p, f"\t{expr}" + (f"\t({n})" if n else ""))
        self.last = "formula"
        self.blank(keep=bool(where))
        for i, (sym, desc) in enumerate(where):
            end = "." if i == len(where) - 1 else ";"
            q = self._par(L, Cm(-0.95) if i == 0 else 0)
            q.paragraph_format.left_indent = Cm(0.95)
            q.paragraph_format.tab_stops.add_tab_stop(Cm(0.95))
            q.paragraph_format.keep_with_next = i < len(where) - 1
            self._runs(q, ("где\t" if i == 0 else "") + f"{sym} – {desc.rstrip(';.')}{end}")
            self.last = "text"
        return n

    # ------------------------------------------------------------------ источники, приложения, содержание
    def sources(self, items, new_page=True):
        """«[1] …» в порядке первого упоминания. В длинных URL после «/» — пробел нулевой ширины."""
        self.unnumbered_heading("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", new_page=new_page)
        for i, s in enumerate(items, 1):
            s = re.sub(r"https?://\S+", lambda m: m.group().replace("/", "/" + ZWSP), s.strip())
            self.text(f"[{i}] {s}")

    def appendix(self, letter, title, kind="обязательное"):
        assert letter in APPENDIX_LETTERS, f"буква {letter} для приложений не используется"
        self.pending_break = True
        for s in (f"ПРИЛОЖЕНИЕ {letter}", f"({kind})", title):
            p = self._runs(self._par(C, 0), s, bold=True)
            p.paragraph_format.keep_with_next = True
        self.last = "heading"
        self.blank(keep=True)
        self.last = "heading"
        self.chapter = letter
        self.count = dict.fromkeys(self.count, 0)

    def toc(self, entries, new_page=True):
        """Содержание: entries = [(уровень 1–3, «2.1 Название», страница), ...].
        Номера страниц бери из экспортированного PDF и пересобирай, пока не совпадут."""
        self.unnumbered_heading("СОДЕРЖАНИЕ", new_page=new_page)
        indent = {1: 0, 2: Cm(0.5), 3: Cm(1.25)}
        for level, title, page in entries:
            p = self._par(L, 0)
            pf = p.paragraph_format
            pf.left_indent = indent[level]
            pf.right_indent = Cm(0.8)                # перенос строки не заходит под номер страницы
            pf.tab_stops.add_tab_stop(Cm(TEXT_W), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            self._runs(p, f"{title}\t{page}")
        self.last = "text"

    # ------------------------------------------------------------------ титульные листы и нумерация
    def _tline(self, s="", align=C, bold=False, italic=False, size=None, indent=0, left=None):
        p = self._par(align, indent)
        if left is not None:
            p.paragraph_format.left_indent = Cm(left)
        if s:
            self._runs(p, s, bold=bold, italic=italic, size=size)
        width = TEXT_W - (left or 0)
        self._title_used += sum(_lines(part, bold, size or 14, width) for part in s.split("\t")) if s else 1
        return p

    def _gap(self, n):
        for _ in range(max(0, n)):
            self._tline()

    def _signature(self, left, right, bold=False):
        p = self._runs(self._par(L, 0), left + "\t" + right, bold=bold)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W), WD_TAB_ALIGNMENT.RIGHT)
        self._title_used += 1

    def _university(self):
        for s in ("Министерство образования Республики Беларусь", "Учреждение образования",
                  "БЕЛОРУССКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ", "ИНФОРМАТИКИ И РАДИОЭЛЕКТРОНИКИ"):
            self._tline(s)

    def _city(self, city, year):
        self._gap(TITLE_LINES - self._title_used - 1)
        self._tline(f"{city} {year}")
        self.start_page_numbers()

    def title_lab(self, faculty, department, work_no, topic, discipline, student, group, teacher,
                  teacher_label="Проверил", kind="ОТЧЁТ", city="Минск", year=None):
        """Титульный лист отчёта по ЛР (без номера страницы). student/teacher — «И.О. Фамилия»."""
        self._title_used = 0
        self._university()
        self._gap(2)
        self._tline(faculty, L)
        self._tline(department, L)
        self._gap(6)
        self._tline(kind, bold=True, size=16)
        self._tline(f"по лабораторной работе № {work_no}")
        self._gap(1)
        self._tline(topic.upper(), bold=True)
        self._gap(1)
        self._tline(f"по дисциплине «{discipline}»")
        self._gap(7)
        self._signature(f"Выполнил студент гр. {group}", student)
        self._signature(teacher_label, teacher)
        self._city(city, year or _year())

    def title_kp(self, faculty, department, head_title, head, topic, discipline, code, student, teacher,
                 kind="курсового проекта", city="Минск", year=None):
        """Титульный лист пояснительной записки КП по образцу кафедры.
        head_title — «Заведующий кафедрой ЗИ», head — «И.О. Фамилия», code — «БГУИР КП 1-98 01 02 099 ПЗ»."""
        self._title_used = 0
        self._university()
        self._gap(1)
        self._tline(faculty, L)
        self._tline(department, L)
        self._gap(2)
        self._tline("К защите допустить:", L, italic=True, left=9.5)
        self._tline(head_title, L, left=9.5)
        self._tline("____________ " + head, L, left=9.5)
        self._gap(1)
        self._tline("ПОЯСНИТЕЛЬНАЯ ЗАПИСКА")
        self._tline(kind)
        self._tline("на тему")
        self._tline(topic.upper(), bold=True, size=16)
        self._tline(f"по дисциплине «{discipline}»")
        self._tline(code)
        self._gap(5)
        self._signature("Студент", student)
        self._signature("Руководитель", teacher)
        self._city(city, year or _year())

    def start_page_numbers(self):
        """Новый раздел документа с номером страницы справа внизу; на титульном листе номера нет,
        но он входит в нумерацию."""
        sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        self._page(sec)
        sec.footer.is_linked_to_previous = False      # иначе номер появится и на титуле
        fp = sec.footer.paragraphs[0]
        fp.alignment = R
        fp.paragraph_format.first_line_indent = 0
        for kind, instr in (("begin", None), ("instr", " PAGE "), ("separate", None), ("end", None)):
            r = fp.add_run()._r
            if instr:
                el = OxmlElement("w:instrText")
                el.set(qn("xml:space"), "preserve")
                el.text = instr
            else:
                el = OxmlElement("w:fldChar")
                el.set(qn("w:fldCharType"), kind)
            r.append(el)
        self.last = None

    def save(self, path):
        self.doc.save(path)
        return path


def _year():
    import datetime
    return datetime.date.today().year


# ==================================================================== экспорт и проверки
def export_pdf(docx_path):
    """docx → pdf: Word через COM (Windows), иначе LibreOffice. Возвращает путь к PDF или None."""
    docx_path = os.path.abspath(docx_path)
    pdf = os.path.splitext(docx_path)[0] + ".pdf"
    if os.name == "nt":
        ps = (f"$w = New-Object -ComObject Word.Application; $w.Visible = $false; "
              f"$d = $w.Documents.Open('{docx_path}'); $d.Fields.Update() | Out-Null; "
              f"$d.ExportAsFixedFormat('{pdf}', 17); $d.Close(0); $w.Quit()")
        if subprocess.run(["powershell", "-NoProfile", "-Command", ps]).returncode == 0 and os.path.exists(pdf):
            return pdf
    for exe in ("soffice", "libreoffice"):
        try:
            subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(docx_path),
                            docx_path], check=True, capture_output=True)
            return pdf
        except Exception:
            continue
    return None


def _norm(s):
    s = s.replace(NBSP, " ").replace(ZWSP, "").replace("\u2011", "-")
    s = re.sub(" " + DASH + r"(?=\S)", " " + DASH + " ", s)      # «слово –слово» (разрядка вместо пробела)
    return re.sub(r" +", " ", s).strip()


UNITS = r"(мм|см|м|км|Вт|кВт|В|А|Гц|ГГц|МГц|дБ|дБм|ГБ|МБ|ТБ|Гбит/с|Мбит/с|пт|с|мин|ч|%|шт|руб|BYN|U|°)"


def check_docx(path):
    """Текстовые правила СТП и замечаний. Возвращает список строк «ОШИБКА/ПРОВЕРЬ: …»."""
    doc = Document(path)
    out = []
    pars = doc.paragraphs
    texts = [_norm(p.text) for p in pars]
    start = next((i for i, t in enumerate(texts) if t in UNNUMBERED), 0)
    body = "\n".join(texts[start:])
    figs, tabs = [], []
    empty_run = 0
    for i in range(start, len(pars)):
        t, p = texts[i], pars[i]
        has_img = bool(p._p.xpath(".//w:drawing"))
        if not t and not has_img and not p._p.xpath("./w:pPr/w:sectPr"):
            empty_run += 1
            if empty_run == 2:
                out.append(f"ОШИБКА: две пустые строки подряд перед «{texts[i + 1][:40] if i + 1 < len(texts) else ''}»")
            continue
        empty_run = 0
        m = re.match(r"^(Рисунок|Таблица) ([А-Я]\.\d+|\d+(?:\.\d+)?)\b(.*)$", t)
        if m:
            kind, num, rest = m.groups()
            (figs if kind == "Рисунок" else tabs).append(num)
            if not re.match(r"^ " + DASH + r" [А-ЯЁA-Z0-9«]", rest):
                out.append(f"ОШИБКА: подпись «{t[:50]}» — нужно «{kind} {num} – Название» с прописной")
            if t.endswith("."):
                out.append(f"ОШИБКА: точка в конце подписи «{t[:50]}»")
        runs = [r for r in p.runs if r.text.strip()]
        if runs and all(r.bold for r in runs) and len(t) < 120 and t.endswith("."):
            out.append(f"ОШИБКА: точка в конце заголовка «{t[:50]}»")
        if re.search(r"\s-\s", t):
            out.append(f"ОШИБКА: дефис вместо тире: «{t[:60]}»")
        if t.startswith(DASH) and not p.text.startswith(DASH + NBSP):
            out.append(f"ОШИБКА: после тире в пункте перечисления нужен неразрывный пробел: «{t[:40]}»")
        if not m and not re.match(r"^(\[\d+\]|\d+(\.\d+)*\s)", t):
            for w in re.finditer(r"(?<![\d.,–-])\b([1-9])\s+(?!" + UNITS + r"\b)([а-яё]+)", t):
                out.append(f"ПРОВЕРЬ: число без единиц словами? «…{t[max(0, w.start() - 15):w.end() + 10]}…»")
    for num in figs:
        if not re.search(r"рисун\w*\s+" + re.escape(num) + r"(?![\d.]\d)", re.sub(r"^Рисунок .*$", "", body, flags=re.M)):
            out.append(f"ОШИБКА: нет ссылки в тексте на рисунок {num}")
    for num in tabs:
        if not re.search(r"таблиц\w*\s+" + re.escape(num) + r"(?![\d.]\d)", re.sub(r"^(Продолжение )?[Тт]аблиц.*$", "", body, flags=re.M)):
            out.append(f"ОШИБКА: нет ссылки в тексте на таблицу {num}")
    for letter in re.findall(r"^ПРИЛОЖЕНИЕ ([А-Я])$", body, flags=re.M):
        if letter not in APPENDIX_LETTERS:
            out.append(f"ОШИБКА: буква приложения {letter} не используется")
        if not re.search(r"приложени\w*\s+" + letter + r"\b", body):
            out.append(f"ОШИБКА: нет ссылки в тексте на приложение {letter}")
    for tbl in doc.tables:
        head = [_norm(c.text) for c in tbl.rows[0].cells]
        if any(re.search(r"№\s*п/п|^№$", h) for h in head):
            out.append(f"ОШИБКА: графа «№ п/п» в таблице «{head[0][:30]}»")
        if any("\u00ad" in h or re.search(r"[а-я]-$", h) for h in head):
            out.append(f"ОШИБКА: перенос в шапке графы таблицы «{head[0][:30]}»")
        for row in tbl.rows[1:]:
            if any(not c.text.strip() for c in row.cells) and any(c.text.strip() for c in row.cells):
                out.append(f"ОШИБКА: пустая ячейка в таблице «{head[0][:30]}» — поставь «–»")
                break
    if not re.search(r"\[\d+\]", body):
        out.append("ПРОВЕРЬ: в тексте нет ни одной ссылки на источник «[N]»")
    return out


def check_pdf(pdf, top_mm=20, bottom_mm=20, min_fill=0.25):
    """Заполнение страниц (> 25 % полезной высоты, без колонтитула) и «Продолжение таблицы» на стыках."""
    import pymupdf
    out = []
    d = pymupdf.open(pdf)
    mm = 72 / 25.4
    for i, page in enumerate(d):
        h = page.rect.height
        if page.rect.width > page.rect.height or abs(page.rect.width - 595) > 5:
            continue                                   # листы-чертежи и альбомные страницы
        top, bottom = top_mm * mm, h - bottom_mm * mm
        ys = [b[3] for b in page.get_text("blocks") if b[4].strip() and b[1] < bottom - 2]
        ys += [r.y1 for img in page.get_images() for r in page.get_image_rects(img[0]) if r.y0 < bottom]
        ys += [dr["rect"].y1 for dr in page.get_drawings() if dr["rect"].y0 < bottom - 2]
        fill = (max(ys) - top) / (bottom - top) if ys else 0
        if i > 0 and fill < min_fill:
            out.append(f"ПРОВЕРЬ: страница {i + 1} заполнена на {fill:.0%} — допустимо только для последней "
                       f"страницы раздела")
        text = page.get_text()
        prev = d[i - 1].get_text() if i else ""
        lines = [s.strip() for s in text.splitlines() if s.strip()]
        if i and lines and prev and re.search(r"Таблица \S+", prev) and page.get_drawings() and \
                not any(s.startswith("Продолжение таблицы") for s in lines[:3]):
            first_draw = min((dr["rect"].y0 for dr in page.get_drawings()), default=h)
            first_text = min((b[1] for b in page.get_text("blocks") if b[4].strip()), default=h)
            if first_draw <= first_text + 2:
                out.append(f"ОШИБКА: страница {i + 1} начинается с таблицы без «Продолжение таблицы N»")
    return out


# ==================================================================== пример
def demo(path):
    from PIL import Image, ImageDraw
    img = os.path.join(os.path.dirname(os.path.abspath(path)), "_demo_scheme.png")
    im = Image.new("RGB", (1200, 700), "white")
    dr = ImageDraw.Draw(im)
    dr.rectangle((40, 40, 1160, 660), outline="black", width=4)
    dr.line((40, 350, 1160, 350), fill="black", width=3)
    im.save(img)

    d = StpDoc()
    d.title_lab(faculty="Факультет информационной безопасности", department="Кафедра защиты информации",
                work_no=5, topic="Построение виртуальных локальных сетей",
                discipline="Основы построения локальных сетей", student="И.О. Фамилия", group="000000",
                teacher="И.О. Фамилия", teacher_label="Проверил")
    d.unnumbered_heading("ЦЕЛЬ РАБОТЫ", new_page=False)
    d.text("Изучить принципы построения виртуальных локальных сетей и настроить маршрутизацию между ними.")
    d.heading("1", "Исходные данные", new_page=False)
    d.text("Схема сети приведена на рисунке 1.1, план адресации – в таблице 1.1. Сеть содержит три коммутатора "
           "C9200L-24P-4X-E и маршрутизатор с портами 1000BASE-T [1].")
    d.items(["коммутаторы уровня доступа", "маршрутизатор", "точки доступа"])
    d.figure(img, "Схема сети")
    d.table("План адресации виртуальных сетей", ["Номер VLAN", "Имя VLAN", "Сеть"],
            [["5", "Management", "10.100.5.0/24"], ["6", "Urban", ""], ["7", "Architecture", "10.100.7.0/24"]],
            [3.5, 5.0, 6.0], split_at=[2])
    d.heading("1.1", "Расчёт мощности источника питания")
    d.text("Мощность определяется по формуле (1.1).")
    d.formula("P = U · I", [("P", "мощность, Вт"), ("U", "напряжение, В"), ("I", "сила тока, А")])
    d.text("Полный расчёт приведён в приложении А.")
    d.unnumbered_heading("ВЫВОД")
    d.text("Настроены четыре виртуальные сети, связь между ними проверена.")
    d.sources(["Белоусова, Е. С. Основы построения локальных сетей. Лабораторный практикум : учебно-методическое "
               "пособие / Е. С. Белоусова. – Минск : БГУИР, 2020. – 103 с."])
    d.appendix("А", "Расчёт мощности")
    d.text("Расчёт по формуле (А.1).")
    d.formula("P = 230 · 0,5 = 115 Вт")
    d.save(path)
    return path


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "demo":
        out = demo(sys.argv[2])
        pdf = export_pdf(out)
        problems = check_docx(out) + (check_pdf(pdf) if pdf else ["ПРОВЕРЬ: PDF не создан — нет Word/LibreOffice"])
        print(out, pdf or "")
        print("\n".join(problems) or "Замечаний нет")
    elif len(sys.argv) >= 3 and sys.argv[1] == "check":
        problems = check_docx(sys.argv[2])
        if len(sys.argv) >= 4:
            problems += check_pdf(sys.argv[3])
        print("\n".join(problems) or "Замечаний нет")
    else:
        print(__doc__)
