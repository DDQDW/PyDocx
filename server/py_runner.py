import csv
import io
import json
import os
import re
import sys
import traceback
import types


def _extract_python_blocks_from_markdown(raw_code):
    blocks = re.findall(r"```(?:python|py)?\s*\n([\s\S]*?)```", str(raw_code), re.IGNORECASE)
    return "\n\n".join(b.strip() for b in blocks if b.strip())


def _parse_csv(csv_string):
    if not csv_string or not str(csv_string).strip():
        return []
    reader = csv.reader(io.StringIO(str(csv_string)))
    return [row for row in reader]


def _escape_csv_cell(value):
    s = str(value or "")
    if any(ch in s for ch in (",", '"', "\n")):
        return '"' + s.replace('"', '""') + '"'
    return s


def _parse_layout_from_text(raw_text, source_filename=""):
    trimmed = str(raw_text or "").strip()
    if not trimmed:
        return "", "", "", ""
    lower_name = str(source_filename or "").lower()
    if lower_name.endswith(".json") or (trimmed.startswith("{") and trimmed.endswith("}")):
        try:
            parsed = json.loads(trimmed)
            if isinstance(parsed, dict):
                list_val = parsed.get("listItems") or parsed.get("list") or parsed.get("items") or ""
                list_str = "\n".join(str(x) for x in list_val) if isinstance(list_val, list) else str(list_val)
                table_val = parsed.get("tableCsv") or parsed.get("table") or ""
                if isinstance(table_val, list):
                    rows = []
                    for r in table_val:
                        rows.append(",".join(_escape_csv_cell(c) for c in r) if isinstance(r, list) else str(r))
                    table_str = "\n".join(rows)
                else:
                    table_str = str(table_val)
                return str(parsed.get("title", "")), str(parsed.get("text") or parsed.get("body") or ""), list_str, table_str
        except Exception:
            pass
    if lower_name.endswith(".csv"):
        base_title = os.path.splitext(os.path.basename(source_filename))[0]
        return base_title, "", "", trimmed
    title = ""
    text_lines = []
    list_lines = []
    table_rows = []
    in_table_block = False
    for raw_line in str(raw_text).splitlines():
        line = raw_line.rstrip()
        stripped = line.strip()
        if re.match(r"^TABLE:\s*$", stripped, re.IGNORECASE):
            in_table_block = True
            continue
        if in_table_block:
            if not stripped:
                in_table_block = False
                continue
            if stripped.startswith("|") and stripped.endswith("|"):
                if re.match(r"^\|[\s:\-|]+\|$", stripped):
                    continue
                cells = [_escape_csv_cell(c.strip()) for c in stripped[1:-1].split("|")]
                table_rows.append(",".join(cells))
            else:
                table_rows.append(stripped)
            continue
        title_match = re.match(r"^(?:TITLE|Заголовок)\s*:\s*(.+)$", stripped, re.IGNORECASE)
        if title_match and not title:
            title = title_match.group(1).strip()
            continue
        h1_match = re.match(r"^#\s+(.+)$", stripped)
        if h1_match and not title:
            title = h1_match.group(1).strip()
            continue
        sub_match = re.match(r"^#{2,6}\s+(.+)$", stripped)
        if sub_match:
            text_lines.append(sub_match.group(1).strip())
            continue
        if stripped.startswith("|") and stripped.endswith("|") and len(stripped) > 2:
            if not re.sub(r"[\s|:\-]", "", stripped):
                continue
            cells = [_escape_csv_cell(c.strip()) for c in stripped[1:-1].split("|")]
            table_rows.append(",".join(cells))
            continue
        bullet_match = re.match(r"^(?:[-*•]|\d+[.)])\s+(.+)$", stripped)
        if bullet_match:
            list_lines.append(bullet_match.group(1).strip())
            continue
        text_lines.append(line)
    if not title and source_filename:
        title = os.path.splitext(os.path.basename(source_filename))[0]
    return title, "\n".join(text_lines).strip(), "\n".join(list_lines), "\n".join(table_rows)


class _Length(int):
    """Stores twips (1/20 pt), mimicking python-docx Length."""
    @property
    def pt(self):
        return float(self) / 20.0

    @property
    def cm(self):
        return float(self) / 566.929133858

    @property
    def mm(self):
        return float(self) / 56.6929133858

    @property
    def inches(self):
        return float(self) / 1440.0

    @property
    def twips(self):
        return int(self)


def _Pt(val):
    return _Length(round(float(val) * 20.0))


def _Cm(val):
    return _Length(round(float(val) * 566.929133858))


def _Mm(val):
    return _Length(round(float(val) * 56.6929133858))


def _Inches(val):
    return _Length(round(float(val) * 1440.0))


def _to_pt(val):
    if val is None:
        return None
    if isinstance(val, _Length):
        return val.pt
    return float(val)


def _to_twips(val):
    if val is None:
        return None
    if isinstance(val, _Length):
        return int(val)
    return round(float(val) * 20.0)


def _install_shims(artifacts_out, work_dir):
    # --- docx shim ---
    docx_mod = types.ModuleType("docx")
    docx_shared = types.ModuleType("docx.shared")
    docx_enum = types.ModuleType("docx.enum")
    docx_enum_text = types.ModuleType("docx.enum.text")
    docx_enum_table = types.ModuleType("docx.enum.table")
    docx_enum_section = types.ModuleType("docx.enum.section")
    docx_oxml = types.ModuleType("docx.oxml")
    docx_oxml_ns = types.ModuleType("docx.oxml.ns")

    class _RGBColor(tuple):
        def __new__(cls, r, g, b):
            return super().__new__(cls, (int(r) & 255, int(g) & 255, int(b) & 255))

        def __str__(self):
            return f"{self[0]:02X}{self[1]:02X}{self[2]:02X}"

    docx_shared.Pt = _Pt
    docx_shared.Cm = _Cm
    docx_shared.Mm = _Mm
    docx_shared.Inches = _Inches
    docx_shared.Twips = lambda v: _Length(int(v))
    docx_shared.RGBColor = _RGBColor

    class _AlignEnum:
        LEFT = "LEFT"
        CENTER = "CENTER"
        RIGHT = "RIGHT"
        JUSTIFY = "JUSTIFY"
        DISTRIBUTE = "JUSTIFY"

    docx_enum_text.WD_ALIGN_PARAGRAPH = _AlignEnum
    docx_enum_text.WD_PARAGRAPH_ALIGNMENT = _AlignEnum

    class _TableAlignEnum:
        LEFT = "LEFT"
        CENTER = "CENTER"
        RIGHT = "RIGHT"

    docx_enum_table.WD_TABLE_ALIGNMENT = _TableAlignEnum

    class _SectionOrientEnum:
        PORTRAIT = "PORTRAIT"
        LANDSCAPE = "LANDSCAPE"

    docx_enum_section.WD_ORIENT = _SectionOrientEnum
    docx_enum_section.WD_ORIENTATION = _SectionOrientEnum

    class _AttrDict(dict):
        def set(self, key, val):
            self[str(key)] = val

    class _RPr:
        def __init__(self):
            self.rFonts = _AttrDict()

    class _Element:
        def __init__(self):
            self.rPr = _RPr()

        def set(self, key, val):
            pass

        def append(self, child):
            pass

        def get_or_add_rPr(self):
            return self.rPr

    def _qn(tag):
        return tag

    docx_oxml_ns.qn = _qn
    docx_oxml_ns.nsdecls = lambda *args: ""
    docx_oxml.OxmlElement = lambda name: _Element()
    docx_oxml.parse_xml = lambda s: _Element()

    class _ColorFormat:
        def __init__(self):
            self.rgb = None

    class _Font:
        def __init__(self):
            self.name = None
            self.size = None
            self.bold = None
            self.italic = None
            self.underline = None
            self.color = _ColorFormat()

    class _Style:
        def __init__(self, name):
            self.name = name
            self.font = _Font()
            self.element = _Element()

    class _ParagraphFormat:
        def __init__(self):
            self.space_before = None
            self.space_after = None
            self.first_line_indent = None
            self.left_indent = None
            self.right_indent = None
            self.line_spacing = None
            self.keep_with_next = False
            self.page_break_before = False

    class _Run:
        def __init__(self, text="", style=None):
            self.text = str(text or "")
            self.style = style
            self.bold = None
            self.italic = None
            self.underline = None
            self.font = _Font()

        def add_break(self, *args, **kwargs):
            self.text += "\n"

    class _Paragraph:
        def __init__(self, text="", style=None):
            self.style = style
            self.alignment = None
            self.paragraph_format = _ParagraphFormat()
            self.runs = []
            self.heading_level = None
            if text:
                self.add_run(text)

        @property
        def text(self):
            return "".join(r.text for r in self.runs)

        @text.setter
        def text(self, val):
            self.runs = [_Run(val)] if val else []

        def add_run(self, text="", style=None):
            r = _Run(text, style=style)
            self.runs.append(r)
            return r

    class _Section:
        def __init__(self):
            self.page_width = _Cm(21.0)
            self.page_height = _Cm(29.7)
            self.top_margin = _Cm(2.0)
            self.bottom_margin = _Cm(2.0)
            self.left_margin = _Cm(2.5)
            self.right_margin = _Cm(1.5)
            self.orientation = _SectionOrientEnum.PORTRAIT

    class _Cell:
        def __init__(self):
            self.paragraphs = [_Paragraph()]

        @property
        def text(self):
            return "\n".join(p.text for p in self.paragraphs)

        @text.setter
        def text(self, val):
            self.paragraphs = [_Paragraph(str(val or ""))]

        def add_paragraph(self, text="", style=None):
            p = _Paragraph(text, style=style)
            self.paragraphs.append(p)
            return p

    class _Row:
        def __init__(self, cols):
            self.cells = [_Cell() for _ in range(cols)]

    class _Table:
        def __init__(self, rows, cols, style=None):
            self.rows = [_Row(cols) for _ in range(rows)]
            self.style = style
            self.alignment = None

        def cell(self, r, c):
            return self.rows[r].cells[c]

        def add_row(self):
            cols = len(self.rows[0].cells) if self.rows else 1
            row = _Row(cols)
            self.rows.append(row)
            return row

    class _Document:
        def __init__(self, path=None):
            self.sections = [_Section()]
            self.styles = {
                "Normal": _Style("Normal"),
                "Heading 1": _Style("Heading 1"),
                "Heading 2": _Style("Heading 2"),
                "Heading 3": _Style("Heading 3"),
                "List Bullet": _Style("List Bullet"),
            }
            self.paragraphs = []
            self.tables = []
            self._items = []

        def add_heading(self, text="", level=1):
            p = _Paragraph(text, style=f"Heading {level}")
            p.heading_level = int(level)
            self.paragraphs.append(p)
            self._items.append(("paragraph", p))
            return p

        def add_paragraph(self, text="", style=None):
            p = _Paragraph(text, style=style)
            self.paragraphs.append(p)
            self._items.append(("paragraph", p))
            return p

        def add_table(self, rows, cols, style=None):
            t = _Table(rows, cols, style=style)
            self.tables.append(t)
            self._items.append(("table", t))
            return t

        def add_page_break(self):
            self._items.append(("page_break", None))

        def _serialize(self):
            normal = self.styles.get("Normal")
            default_font = None
            default_size_pt = None
            if normal:
                default_font = normal.font.name or normal.element.rPr.rFonts.get("w:ascii")
                default_size_pt = _to_pt(normal.font.size)

            sec = self.sections[0] if self.sections else _Section()
            blocks = []
            for kind, obj in self._items:
                if kind == "page_break":
                    blocks.append({"type": "page_break"})
                elif kind == "paragraph":
                    style_name = obj.style.name if hasattr(obj.style, "name") else (str(obj.style) if obj.style else None)
                    pf = obj.paragraph_format
                    runs_json = []
                    for r in obj.runs:
                        runs_json.append(
                            {
                                "text": r.text,
                                "bold": r.bold if r.bold is not None else r.font.bold,
                                "italic": r.italic if r.italic is not None else r.font.italic,
                                "underline": bool(r.underline if r.underline is not None else r.font.underline),
                                "fontName": r.font.name or default_font,
                                "fontSizePt": _to_pt(r.font.size) or default_size_pt,
                                "colorHex": str(r.font.color.rgb) if r.font.color.rgb else None,
                            }
                        )
                    blocks.append(
                        {
                            "type": "paragraph",
                            "style": style_name,
                            "headingLevel": obj.heading_level,
                            "alignment": obj.alignment,
                            "spaceBeforeTwips": _to_twips(pf.space_before),
                            "spaceAfterTwips": _to_twips(pf.space_after),
                            "firstLineIndentTwips": _to_twips(pf.first_line_indent),
                            "leftIndentTwips": _to_twips(pf.left_indent),
                            "rightIndentTwips": _to_twips(pf.right_indent),
                            "lineSpacing": float(pf.line_spacing) if pf.line_spacing is not None and not isinstance(pf.line_spacing, _Length) else None,
                            "lineSpacingPt": _to_pt(pf.line_spacing) if isinstance(pf.line_spacing, _Length) else None,
                            "pageBreakBefore": bool(pf.page_break_before),
                            "runs": runs_json,
                        }
                    )
                elif kind == "table":
                    rows_json = []
                    for row in obj.rows:
                        rows_json.append([cell.text for cell in row.cells])
                    blocks.append({"type": "table", "rows": rows_json})

            return {
                "format": "docx_ir",
                "defaultFont": default_font or "Times New Roman",
                "defaultSizePt": default_size_pt or 12,
                "section": {
                    "pageWidthTwips": _to_twips(sec.page_width),
                    "pageHeightTwips": _to_twips(sec.page_height),
                    "topMarginTwips": _to_twips(sec.top_margin),
                    "bottomMarginTwips": _to_twips(sec.bottom_margin),
                    "leftMarginTwips": _to_twips(sec.left_margin),
                    "rightMarginTwips": _to_twips(sec.right_margin),
                },
                "blocks": blocks,
            }

        def save(self, target):
            spec = self._serialize()
            if isinstance(target, (str, os.PathLike)):
                fname = os.path.basename(str(target)) or "document.docx"
                if not fname.lower().endswith(".docx"):
                    fname += ".docx"
                artifacts_out.append({"kind": "docx_ir", "filename": fname, "spec": spec})
            elif hasattr(target, "write"):
                target.write(json.dumps(spec).encode("utf-8"))

    docx_mod.Document = _Document
    docx_mod.shared = docx_shared
    docx_mod.enum = docx_enum
    docx_mod.oxml = docx_oxml

    sys.modules["docx"] = docx_mod
    sys.modules["docx.shared"] = docx_shared
    sys.modules["docx.enum"] = docx_enum
    sys.modules["docx.enum.text"] = docx_enum_text
    sys.modules["docx.enum.table"] = docx_enum_table
    sys.modules["docx.enum.section"] = docx_enum_section
    sys.modules["docx.oxml"] = docx_oxml
    sys.modules["docx.oxml.ns"] = docx_oxml_ns

    # --- pptx shim ---
    pptx_mod = types.ModuleType("pptx")
    pptx_util = types.ModuleType("pptx.util")
    pptx_util.Pt = _Pt
    pptx_util.Cm = _Cm
    pptx_util.Inches = _Inches

    class _PptxParagraph:
        def __init__(self, text=""):
            self.text = str(text or "")
            self.level = 0

    class _PptxTextFrame:
        def __init__(self):
            self.paragraphs = []
            self._text = ""

        @property
        def text(self):
            return self._text

        @text.setter
        def text(self, val):
            self._text = str(val or "")

        def add_paragraph(self):
            p = _PptxParagraph()
            self.paragraphs.append(p)
            return p

    class _PptxShape:
        def __init__(self):
            self.text_frame = _PptxTextFrame()

        @property
        def text(self):
            return self.text_frame.text

        @text.setter
        def text(self, val):
            self.text_frame.text = str(val or "")

    class _PptxShapes:
        def __init__(self):
            self.title = _PptxShape()

    class _PptxSlide:
        def __init__(self):
            self.shapes = _PptxShapes()
            self.placeholders = {0: self.shapes.title, 1: _PptxShape()}

    class _PptxSlides(list):
        def add_slide(self, layout):
            s = _PptxSlide()
            self.append(s)
            return s

    class _Presentation:
        def __init__(self):
            self.slide_layouts = [i for i in range(10)]
            self.slides = _PptxSlides()

        def save(self, target):
            slides_spec = []
            for s in self.slides:
                body_lines = []
                if s.placeholders[1].text_frame.text:
                    body_lines.append(s.placeholders[1].text_frame.text)
                for p in s.placeholders[1].text_frame.paragraphs:
                    if p.text:
                        body_lines.append(p.text)
                slides_spec.append(
                    {"title": s.shapes.title.text, "bodyLines": body_lines}
                )
            fname = os.path.basename(str(target)) if isinstance(target, (str, os.PathLike)) else "presentation.pptx"
            if not fname.lower().endswith(".pptx"):
                fname += ".pptx"
            artifacts_out.append({"kind": "pptx_ir", "filename": fname, "slides": slides_spec})

    pptx_mod.Presentation = _Presentation
    pptx_mod.util = pptx_util
    sys.modules["pptx"] = pptx_mod
    sys.modules["pptx.util"] = pptx_util

    # --- openpyxl shim ---
    openpyxl_mod = types.ModuleType("openpyxl")
    openpyxl_styles = types.ModuleType("openpyxl.styles")
    openpyxl_styles.Alignment = lambda **kwargs: kwargs
    openpyxl_styles.Font = lambda **kwargs: kwargs

    class _XlsxCell:
        def __init__(self):
            self.value = None
            self.alignment = None
            self.font = None

    class _XlsxSheet:
        def __init__(self, title="Sheet1"):
            self.title = title
            self._cells = {}
            self._max_row = 0
            self._max_col = 0

        def cell(self, row, column, value=None):
            c = self._cells.setdefault((row, column), _XlsxCell())
            if value is not None:
                c.value = value
            self._max_row = max(self._max_row, row)
            self._max_col = max(self._max_col, column)
            return c

        def append(self, iterable):
            r = self._max_row + 1
            for idx, val in enumerate(iterable, start=1):
                self.cell(r, idx, val)

    class _Workbook:
        def __init__(self):
            self.active = _XlsxSheet("Sheet1")
            self.sheets = [self.active]

        def create_sheet(self, title="Sheet"):
            s = _XlsxSheet(title)
            self.sheets.append(s)
            return s

        def save(self, target):
            rows = []
            s = self.active
            for r in range(1, s._max_row + 1):
                row_vals = []
                for c in range(1, s._max_col + 1):
                    cell = s._cells.get((r, c))
                    row_vals.append("" if cell is None or cell.value is None else str(cell.value))
                rows.append(row_vals)
            fname = os.path.basename(str(target)) if isinstance(target, (str, os.PathLike)) else "table.xlsx"
            if not fname.lower().endswith(".xlsx"):
                fname += ".xlsx"
            artifacts_out.append({"kind": "xlsx_ir", "filename": fname, "rows": rows})

    openpyxl_mod.Workbook = _Workbook
    openpyxl_mod.styles = openpyxl_styles
    sys.modules["openpyxl"] = openpyxl_mod
    sys.modules["openpyxl.styles"] = openpyxl_styles

    # --- fpdf shim ---
    fpdf_mod = types.ModuleType("fpdf")

    class _FPDF:
        def __init__(self):
            self._lines = []

        def add_page(self):
            pass

        def add_font(self, *args, **kwargs):
            pass

        def set_font(self, *args, **kwargs):
            pass

        def set_font_size(self, *args, **kwargs):
            pass

        def ln(self, *args, **kwargs):
            self._lines.append("")

        def cell(self, w, h=0, txt="", *args, **kwargs):
            if txt:
                self._lines.append(str(txt))

        def multi_cell(self, w, h=0, txt="", *args, **kwargs):
            if txt:
                self._lines.append(str(txt))

        def output(self, name=None):
            spec = {"kind": "pdf_ir", "filename": os.path.basename(str(name)) if name else "document.pdf", "lines": list(self._lines)}
            if name:
                artifacts_out.append(spec)
                return b""
            return json.dumps(spec).encode("utf-8")

    fpdf_mod.FPDF = _FPDF
    sys.modules["fpdf"] = fpdf_mod


def _build_helpers(artifacts_out, work_dir, source_text):
    def create_docx(title, text, list_items, table_csv, filename):
        fname = os.path.basename(str(filename).strip()) or "document.docx"
        if not fname.lower().endswith(".docx"):
            fname += ".docx"
        artifacts_out.append(
            {
                "kind": "helper",
                "format": "docx",
                "filename": fname,
                "title": str(title or ""),
                "text": str(text or ""),
                "listItems": str(list_items or ""),
                "tableCsv": str(table_csv or ""),
            }
        )
        return f"Файл {fname} сохранен в Downloads"

    def create_pptx(title, text, list_items, filename):
        fname = os.path.basename(str(filename).strip()) or "presentation.pptx"
        if not fname.lower().endswith(".pptx"):
            fname += ".pptx"
        artifacts_out.append(
            {
                "kind": "helper",
                "format": "pptx",
                "filename": fname,
                "title": str(title or ""),
                "text": str(text or ""),
                "listItems": str(list_items or ""),
                "tableCsv": "",
            }
        )
        return f"Файл {fname} сохранен в Downloads"

    def create_xlsx(table_csv, filename):
        fname = os.path.basename(str(filename).strip()) or "table.xlsx"
        if not fname.lower().endswith(".xlsx"):
            fname += ".xlsx"
        artifacts_out.append(
            {
                "kind": "helper",
                "format": "xlsx",
                "filename": fname,
                "title": "",
                "text": "",
                "listItems": "",
                "tableCsv": str(table_csv or ""),
            }
        )
        return f"Файл {fname} сохранен в Downloads"

    def create_pdf(title, text, list_items, filename):
        fname = os.path.basename(str(filename).strip()) or "document.pdf"
        if not fname.lower().endswith(".pdf"):
            fname += ".pdf"
        artifacts_out.append(
            {
                "kind": "helper",
                "format": "pdf",
                "filename": fname,
                "title": str(title or ""),
                "text": str(text or ""),
                "listItems": str(list_items or ""),
                "tableCsv": "",
            }
        )
        return f"Файл {fname} сохранен в Downloads"

    def create_txt(title, text, list_items, table_csv, filename):
        fname = os.path.basename(str(filename).strip()) or "document.txt"
        if not fname.lower().endswith(".txt"):
            fname += ".txt"
        artifacts_out.append(
            {
                "kind": "helper",
                "format": "txt",
                "filename": fname,
                "title": str(title or ""),
                "text": str(text or ""),
                "listItems": str(list_items or ""),
                "tableCsv": str(table_csv or ""),
            }
        )
        return f"Файл {fname} сохранен в Downloads"

    def create_md(title, text, list_items, table_csv, filename):
        fname = os.path.basename(str(filename).strip()) or "document.md"
        if not fname.lower().endswith(".md"):
            fname += ".md"
        artifacts_out.append(
            {
                "kind": "helper",
                "format": "md",
                "filename": fname,
                "title": str(title or ""),
                "text": str(text or ""),
                "listItems": str(list_items or ""),
                "tableCsv": str(table_csv or ""),
            }
        )
        return f"Файл {fname} сохранен в Downloads"

    def create_html(title, text, list_items, table_csv, filename):
        fname = os.path.basename(str(filename).strip()) or "document.html"
        if not fname.lower().endswith(".html"):
            fname += ".html"
        artifacts_out.append(
            {
                "kind": "helper",
                "format": "html",
                "filename": fname,
                "title": str(title or ""),
                "text": str(text or ""),
                "listItems": str(list_items or ""),
                "tableCsv": str(table_csv or ""),
            }
        )
        return f"Файл {fname} сохранен в Downloads"

    def create_from_text(raw_text, fmt="docx", filename="layout"):
        t, txt, lst, tbl = _parse_layout_from_text(raw_text)
        if not filename.endswith(f".{fmt}"):
            filename = f"{filename}.{fmt}"
        if fmt == "docx":
            return create_docx(t, txt, lst, tbl, filename)
        if fmt == "pptx":
            return create_pptx(t, txt, lst, filename)
        if fmt == "xlsx":
            return create_xlsx(tbl, filename)
        if fmt == "pdf":
            return create_pdf(t, txt, lst, filename)
        if fmt == "txt":
            return create_txt(t, txt, lst, tbl, filename)
        if fmt == "md":
            return create_md(t, txt, lst, tbl, filename)
        if fmt == "html":
            return create_html(t, txt, lst, tbl, filename)
        raise ValueError(f"Неизвестный формат: {fmt}")

    def save_file(filename, content_bytes_or_path, is_binary=True):
        safe_name = os.path.basename(str(filename).strip()) or "document"
        if is_binary:
            raw = (
                bytes(content_bytes_or_path)
                if isinstance(content_bytes_or_path, (bytes, bytearray))
                else str(content_bytes_or_path).encode("utf-8")
            )
        else:
            with open(content_bytes_or_path, "rb") as fp:
                raw = fp.read()
        out_path = os.path.join(work_dir, safe_name)
        with open(out_path, "wb") as fp:
            fp.write(raw)
        return out_path

    return {
        "__name__": "__main__",
        "__file__": os.path.join(work_dir, "script.py"),
        "create_docx": create_docx,
        "create_pptx": create_pptx,
        "create_xlsx": create_xlsx,
        "create_pdf": create_pdf,
        "create_txt": create_txt,
        "create_md": create_md,
        "create_html": create_html,
        "create_from_text": create_from_text,
        "parse_layout": _parse_layout_from_text,
        "save_file": save_file,
        "save_file_to_download": save_file,
        "SAVE_DIR": work_dir,
        "DOWNLOAD_DIR": work_dir,
        "SOURCE_FILE": "",
        "SOURCE_DIR": work_dir,
        "SOURCE_TEXT": source_text,
        "os": os,
        "json": json,
        "csv": csv,
        "re": re,
    }


def main():
    real_stdout = sys.stdout
    payload = json.loads(sys.stdin.read())
    code = str(payload.get("code") or "")
    work_dir = str(payload.get("workDir") or os.getcwd())
    os.chdir(work_dir)

    artifacts_out = []
    _install_shims(artifacts_out, work_dir)
    namespace = _build_helpers(artifacts_out, work_dir, code)
    logs = []
    namespace["print"] = lambda *args, **kwargs: logs.append(f"[PRINT] {' '.join(map(str, args))}")

    code_to_run = code
    try:
        compile(code_to_run, "<script>", "exec")
    except SyntaxError as se:
        extracted = _extract_python_blocks_from_markdown(code_to_run)
        if extracted:
            logs.append("ℹ️ Обнаружены блоки кода в Markdown — выполняю код из блоков.")
            code_to_run = extracted
        else:
            real_stdout.write(
                json.dumps(
                    {
                        "ok": False,
                        "syntaxError": True,
                        "error": f"SyntaxError: {se.msg} (line {se.lineno})",
                        "logs": logs,
                        "artifacts": [],
                    },
                    ensure_ascii=False,
                )
            )
            return

    captured = io.StringIO()
    sys.stdout = captured
    try:
        exec(code_to_run, namespace)
        extra_out = captured.getvalue().strip()
        if extra_out:
            logs.append(f"Вывод скрипта:\n{extra_out}")
        res = {
            "ok": True,
            "syntaxError": False,
            "logs": logs,
            "artifacts": artifacts_out,
        }
    except Exception as exc:
        logs.append(f"❌ Ошибка: {exc}\n{traceback.format_exc()}")
        res = {
            "ok": False,
            "syntaxError": False,
            "error": str(exc),
            "logs": logs,
            "artifacts": artifacts_out,
        }
    finally:
        sys.stdout = real_stdout

    real_stdout.write(json.dumps(res, ensure_ascii=False))


if __name__ == "__main__":
    main()
