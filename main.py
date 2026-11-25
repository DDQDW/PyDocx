import csv
import html
import io
import json
import os
import re
import shutil
import sys
import tempfile
import threading
import traceback

import flet as ft
import openpyxl
from docx import Document
from fpdf import FPDF
from openpyxl.styles import Alignment, Font
from pptx import Presentation

try:
    import flet_permission_handler as fph
except ImportError:
    fph = None

DEFAULT_REL_SAVE_DIR = "Documents/PyDocx"
ANDROID_STORAGE_ROOT = "/storage/emulated/0"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(APP_DIR, "assets")
FALLBACK_DIR = os.path.join(APP_DIR, "generated_docs")
FONT_PATH = os.path.join(ASSETS_DIR, "DejaVuSans.ttf")
ICON_PATH = os.path.join(ASSETS_DIR, "icon.png")
GENERATED_EXTENSIONS = (".docx", ".pptx", ".xlsx", ".pdf", ".txt", ".md", ".html", ".csv", ".json")
SCRIPT_ALLOWED_EXTENSIONS = ["py", "txt", "md", "json", "csv", "yaml", "yml", "js", "log"]

PRIMARY_COLOR = "#1976d2"
SURFACE_COLOR = "#ffffff"
WIN_BG_COLOR = "#f3f3f3"
BORDER_COLOR = "#c4c7c5"
DIVIDER_COLOR = "#e0e2e0"
CARD_BORDER_COLOR = "#e5e5e5"
TEXT_COLOR = "#1f1f1f"
MUTED_COLOR = "#5f6368"
BTN_BG = "#f0f4fa"
LOG_BG = "#f6f8fa"
DROPZONE_BG = "#fafbfc"

SAMPLE_LAYOUT_TEXT = """# Отчет по проекту PyDocx
TITLE: Отчет по проекту PyDocx

Ниже приведены ключевые результаты работы и сводная таблица показателей.

## Основные задачи
- Поддержка верстки из текстовых файлов (.txt, .md, .json, .csv, .yaml)
- Экспорт в Word, PowerPoint, Excel, PDF, Markdown и HTML
- Выполнение пользовательских скриптов Python

## Сводная таблица
| Модуль | Статус | Формат |
| --- | --- | --- |
| Свой скрипт | Готово | PY/TXT/MD/JSON |
| Текстовая верстка | Готово | DOCX/PPTX/XLSX/PDF |
| Настройки | Готово | Android/Windows |"""

_created_files_lock = threading.Lock()
_recent_created_files = []


def _record_created_file(filename, data_bytes, saved_path, in_target_dir):
    with _created_files_lock:
        _recent_created_files.append(
            {
                "filename": filename,
                "bytes": data_bytes,
                "path": saved_path,
                "in_target_dir": in_target_dir,
            }
        )


def _pop_recent_created_files():
    with _created_files_lock:
        items = list(_recent_created_files)
        _recent_created_files.clear()
        return items


def ensure_dir(path_dir):
    if path_dir and not os.path.exists(path_dir):
        os.makedirs(path_dir, exist_ok=True)


def _is_cache_or_temp_dir(path_dir):
    if not path_dir:
        return True
    norm = os.path.normpath(path_dir).replace("\\", "/").lower()
    tmp_norm = os.path.normpath(tempfile.gettempdir()).replace("\\", "/").lower()
    if norm.startswith(tmp_norm):
        return True
    if "/cache/" in norm or norm.endswith("/cache") or "/data/user/" in norm or "/data/data/" in norm:
        return True
    return False


def _settings_file_candidates():
    home = os.path.expanduser("~")
    return [
        os.path.join(home, ".pydocx_settings.json"),
        os.path.join(APP_DIR, "pydocx_settings.json"),
    ]


def load_app_settings():
    defaults = {
        "default_dir": DEFAULT_REL_SAVE_DIR,
        "save_next_to_source": True,
    }
    for path in _settings_file_candidates():
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    return {
                        "default_dir": str(loaded.get("default_dir") or DEFAULT_REL_SAVE_DIR).strip()
                        or DEFAULT_REL_SAVE_DIR,
                        "save_next_to_source": bool(loaded.get("save_next_to_source", True)),
                    }
        except Exception:
            continue
    return defaults


def save_app_settings(settings):
    payload = {
        "default_dir": str(settings.get("default_dir") or DEFAULT_REL_SAVE_DIR).strip()
        or DEFAULT_REL_SAVE_DIR,
        "save_next_to_source": bool(settings.get("save_next_to_source", True)),
    }
    for path in _settings_file_candidates():
        try:
            ensure_dir(os.path.dirname(path))
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            continue
    return False


def resolve_save_dir(setting_value, is_android=False):
    raw = str(setting_value or "").strip() or DEFAULT_REL_SAVE_DIR
    if os.path.isabs(raw):
        return os.path.normpath(raw)
    parts = [p for p in raw.replace("\\", "/").split("/") if p and p != "."]
    if is_android:
        return os.path.join(ANDROID_STORAGE_ROOT, *parts)
    return os.path.join(os.path.expanduser("~"), *parts)


def parse_csv(csv_string):
    if not csv_string or not str(csv_string).strip():
        return []
    reader = csv.reader(io.StringIO(str(csv_string)))
    return [row for row in reader]


def save_file_to_target(filename, content_bytes_or_path, is_binary=True, target_dir=None):
    safe_name = os.path.basename(str(filename).strip()) or "document"
    if is_binary:
        data_bytes = (
            bytes(content_bytes_or_path)
            if isinstance(content_bytes_or_path, (bytes, bytearray))
            else str(content_bytes_or_path).encode("utf-8")
        )
    else:
        with open(content_bytes_or_path, "rb") as f:
            data_bytes = f.read()

    preferred_dir = target_dir or resolve_save_dir(DEFAULT_REL_SAVE_DIR, is_android=False)
    in_target_dir = False
    saved_path = None
    try:
        ensure_dir(preferred_dir)
        candidate = os.path.join(preferred_dir, safe_name)
        with open(candidate, "wb") as f:
            f.write(data_bytes)
        saved_path = candidate
        in_target_dir = True
    except Exception:
        ensure_dir(FALLBACK_DIR)
        candidate = os.path.join(FALLBACK_DIR, safe_name)
        with open(candidate, "wb") as f:
            f.write(data_bytes)
        saved_path = candidate
        in_target_dir = False

    _record_created_file(safe_name, data_bytes, saved_path, in_target_dir)
    return saved_path


def _format_save_message(filename, saved_path, in_target_dir):
    if in_target_dir:
        return f"Файл {filename} сохранён: {saved_path}"
    return (
        f"Файл {filename} создан во внутренней папке ({saved_path}). "
        "Нет прямого доступа к папке — нажмите «Сохранить ... в...»."
    )


def _escape_csv_cell(value):
    s = str(value or "")
    if any(ch in s for ch in (",", '"', "\n")):
        return '"' + s.replace('"', '""') + '"'
    return s


def parse_layout_from_text(raw_text, source_filename=""):
    trimmed = str(raw_text or "").strip()
    if not trimmed:
        return "", "", "", ""

    lower_name = str(source_filename or "").lower()

    if lower_name.endswith(".json") or (trimmed.startswith("{") and trimmed.endswith("}")):
        try:
            parsed = json.loads(trimmed)
            if isinstance(parsed, dict):
                list_val = (
                    parsed.get("listItems")
                    or parsed.get("list")
                    or parsed.get("items")
                    or ""
                )
                if isinstance(list_val, list):
                    list_str = "\n".join(str(x) for x in list_val)
                else:
                    list_str = str(list_val)

                table_val = parsed.get("tableCsv") or parsed.get("table") or ""
                if isinstance(table_val, list):
                    rows = []
                    for r in table_val:
                        if isinstance(r, list):
                            rows.append(",".join(_escape_csv_cell(c) for c in r))
                        else:
                            rows.append(str(r))
                    table_str = "\n".join(rows)
                else:
                    table_str = str(table_val)

                return (
                    str(parsed.get("title", "")),
                    str(parsed.get("text") or parsed.get("body") or ""),
                    list_str,
                    table_str,
                )
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

        title_match = re.match(
            r"^(?:TITLE|Заголовок)\s*:\s*(.+)$", stripped, re.IGNORECASE
        )
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

    return (
        title,
        "\n".join(text_lines).strip(),
        "\n".join(list_lines),
        "\n".join(table_rows),
    )


def create_docx(title, text, list_items, table_csv, filename, target_dir=None):
    doc = Document()
    if title:
        doc.add_heading(str(title), level=1)
    if text:
        doc.add_paragraph(str(text))
    if list_items:
        for item in str(list_items).splitlines():
            if item.strip():
                doc.add_paragraph(item.strip(), style="List Bullet")
    table_data = parse_csv(table_csv)
    if table_data:
        rows = len(table_data)
        cols = len(table_data[0]) if rows > 0 else 0
        if cols > 0:
            table = doc.add_table(rows=rows, cols=cols)
            table.style = "Table Grid"
            for i, row in enumerate(table_data):
                for j, cell_value in enumerate(row):
                    if j < cols:
                        table.cell(i, j).text = str(cell_value)
    buf = io.BytesIO()
    doc.save(buf)
    saved_path = save_file_to_target(filename, buf.getvalue(), is_binary=True, target_dir=target_dir)
    in_target = target_dir is not None and os.path.dirname(saved_path) == os.path.normpath(target_dir)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_target)


def create_pptx(title, text, list_items, filename, target_dir=None):
    prs = Presentation()
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = str(title or "")
    body_shape = slide.placeholders[1]
    tf = body_shape.text_frame
    tf.text = str(text or "")
    if list_items:
        for item in str(list_items).splitlines():
            if item.strip():
                p = tf.add_paragraph()
                p.text = item.strip()
                p.level = 1
    buf = io.BytesIO()
    prs.save(buf)
    saved_path = save_file_to_target(filename, buf.getvalue(), is_binary=True, target_dir=target_dir)
    in_target = target_dir is not None and os.path.dirname(saved_path) == os.path.normpath(target_dir)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_target)


def create_xlsx(table_csv, filename, target_dir=None):
    table_data = parse_csv(table_csv)
    if not table_data:
        raise ValueError("Нет данных для таблицы Excel.")
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, row in enumerate(table_data):
        for j, value in enumerate(row):
            cell = ws.cell(row=i + 1, column=j + 1, value=value)
            cell.alignment = Alignment(wrap_text=True)
            if i == 0:
                cell.font = Font(bold=True)
    buf = io.BytesIO()
    wb.save(buf)
    saved_path = save_file_to_target(filename, buf.getvalue(), is_binary=True, target_dir=target_dir)
    in_target = target_dir is not None and os.path.dirname(saved_path) == os.path.normpath(target_dir)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_target)


def _is_valid_ttf(font_path):
    if not os.path.exists(font_path):
        return False
    try:
        with open(font_path, "rb") as f:
            header = f.read(4)
        return header in (b"\x00\x01\x00\x00", b"OTTO", b"true")
    except Exception:
        return False


def create_pdf(title, text, list_items, filename, target_dir=None):
    pdf = FPDF()
    pdf.add_page()
    if _is_valid_ttf(FONT_PATH):
        pdf.add_font("DejaVu", "", FONT_PATH)
        pdf.set_font("DejaVu", size=16)
    else:
        pdf.set_font("Helvetica", size=16)
    pdf.cell(0, 10, str(title or ""), ln=True, align="C")
    pdf.ln(10)
    pdf.set_font_size(12)
    pdf.multi_cell(0, 8, str(text or ""))
    if list_items:
        for item in str(list_items).splitlines():
            if item.strip():
                pdf.multi_cell(0, 8, f"• {item.strip()}")
    pdf_bytes = bytes(pdf.output())
    saved_path = save_file_to_target(filename, pdf_bytes, is_binary=True, target_dir=target_dir)
    in_target = target_dir is not None and os.path.dirname(saved_path) == os.path.normpath(target_dir)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_target)


def create_txt(title, text, list_items, table_csv, filename, target_dir=None):
    parts = []
    if title:
        parts.append(str(title).upper())
        parts.append("=" * min(max(len(str(title)), 3), 60))
        parts.append("")
    if text:
        parts.append(str(text))
        parts.append("")
    if list_items:
        for item in str(list_items).splitlines():
            if item.strip():
                parts.append(f"• {item.strip()}")
        parts.append("")
    for row in parse_csv(table_csv):
        parts.append("\t|\t".join(row))
    content = "\n".join(parts).strip() + "\n"
    saved_path = save_file_to_target(filename, content.encode("utf-8"), is_binary=True, target_dir=target_dir)
    in_target = target_dir is not None and os.path.dirname(saved_path) == os.path.normpath(target_dir)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_target)


def create_md(title, text, list_items, table_csv, filename, target_dir=None):
    parts = []
    if title:
        parts.append(f"# {title}\n")
    if text:
        parts.append(f"{text}\n")
    if list_items:
        for item in str(list_items).splitlines():
            if item.strip():
                parts.append(f"- {item.strip()}")
        parts.append("")
    table_data = parse_csv(table_csv)
    if table_data:
        cols = len(table_data[0])
        parts.append("| " + " | ".join(table_data[0]) + " |")
        parts.append("| " + " | ".join(["---"] * cols) + " |")
        for row in table_data[1:]:
            padded = [row[i] if i < len(row) else "" for i in range(cols)]
            parts.append("| " + " | ".join(padded) + " |")
    content = "\n".join(parts).strip() + "\n"
    saved_path = save_file_to_target(filename, content.encode("utf-8"), is_binary=True, target_dir=target_dir)
    in_target = target_dir is not None and os.path.dirname(saved_path) == os.path.normpath(target_dir)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_target)


def create_html(title, text, list_items, table_csv, filename, target_dir=None):
    sections = []
    if title:
        sections.append(f"<h1>{html.escape(str(title))}</h1>")
    if text:
        for para in str(text).splitlines():
            if para.strip():
                sections.append(f"<p>{html.escape(para)}</p>")
    if list_items:
        items = [
            f"<li>{html.escape(i.strip())}</li>"
            for i in str(list_items).splitlines()
            if i.strip()
        ]
        if items:
            sections.append("<ul>" + "".join(items) + "</ul>")
    table_data = parse_csv(table_csv)
    if table_data:
        rows_html = []
        for idx, row in enumerate(table_data):
            tag = "th" if idx == 0 else "td"
            rows_html.append(
                "<tr>" + "".join(f"<{tag}>{html.escape(c)}</{tag}>" for c in row) + "</tr>"
            )
        sections.append("<table>" + "".join(rows_html) + "</table>")

    doc_html = (
        f'<!doctype html><html lang="ru"><head><meta charset="UTF-8">'
        f"<title>{html.escape(str(title or filename))}</title>"
        f"<style>body{{font-family:sans-serif;max-width:800px;margin:40px auto;padding:0 16px}}"
        f"table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #ccc;padding:6px 8px}}</style>"
        f"</head><body>{''.join(sections)}</body></html>"
    )
    saved_path = save_file_to_target(filename, doc_html.encode("utf-8"), is_binary=True, target_dir=target_dir)
    in_target = target_dir is not None and os.path.dirname(saved_path) == os.path.normpath(target_dir)
    return _format_save_message(os.path.basename(saved_path), saved_path, in_target)


def create_from_text(raw_text, fmt="docx", filename="layout", target_dir=None):
    title, text, list_items, table_csv = parse_layout_from_text(raw_text)
    if not filename.endswith(f".{fmt}"):
        filename = f"{filename}.{fmt}"
    if fmt == "docx":
        return create_docx(title, text, list_items, table_csv, filename, target_dir=target_dir)
    if fmt == "pptx":
        return create_pptx(title, text, list_items, filename, target_dir=target_dir)
    if fmt == "xlsx":
        return create_xlsx(table_csv, filename, target_dir=target_dir)
    if fmt == "pdf":
        return create_pdf(title, text, list_items, filename, target_dir=target_dir)
    if fmt == "txt":
        return create_txt(title, text, list_items, table_csv, filename, target_dir=target_dir)
    if fmt == "md":
        return create_md(title, text, list_items, table_csv, filename, target_dir=target_dir)
    if fmt == "html":
        return create_html(title, text, list_items, table_csv, filename, target_dir=target_dir)
    raise ValueError(f"Неизвестный формат: {fmt}")


def _extract_python_blocks_from_markdown(raw_code):
    blocks = re.findall(r"```(?:python|py)?\s*\n([\s\S]*?)```", str(raw_code), re.IGNORECASE)
    return "\n\n".join(b.strip() for b in blocks if b.strip())


def _snapshot_dir(target_dir):
    snap = {}
    try:
        for name in os.listdir(target_dir):
            full = os.path.join(target_dir, name)
            if os.path.isfile(full):
                st = os.stat(full)
                snap[full] = (st.st_mtime_ns, st.st_size)
    except Exception:
        pass
    return snap


def _dir_is_writable(path_dir):
    try:
        ensure_dir(path_dir)
        probe = os.path.join(path_dir, ".pydocx_write_test")
        with open(probe, "wb") as f:
            f.write(b"1")
        os.remove(probe)
        return True
    except Exception:
        return False


def _looks_like_executable_script(raw_text, source_name=""):
    code = str(raw_text or "").strip()
    if not code:
        return False
    lower_name = str(source_name or "").lower()
    script_markers = (
        "from docx import",
        "import docx",
        "from pptx import",
        "import openpyxl",
        "from fpdf import",
        "create_docx(",
        "create_pptx(",
        "create_xlsx(",
        "create_pdf(",
        "create_from_text(",
        "doc.save(",
        'if __name__ == "__main__":',
        "if __name__ == '__main__':",
    )
    if lower_name.endswith(".py") or any(m in code for m in script_markers):
        try:
            compile(code, "<check>", "exec")
            return True
        except SyntaxError:
            extracted = _extract_python_blocks_from_markdown(code)
            return bool(extracted)
    return False


def execute_user_script(
    code,
    log_output,
    target_dir,
    source_path="",
    on_files_ready=None,
):
    def run():
        old_stdout = sys.stdout
        old_cwd = os.getcwd()
        _pop_recent_created_files()

        target_dir_norm = os.path.normpath(target_dir)
        staging_dir = None
        if _dir_is_writable(target_dir_norm):
            work_dir = target_dir_norm
        else:
            staging_dir = tempfile.mkdtemp(prefix="pydocx_run_")
            work_dir = staging_dir

        switched_cwd = False
        try:
            os.chdir(work_dir)
            switched_cwd = True
        except Exception:
            switched_cwd = False

        watch_dirs = list(dict.fromkeys([work_dir, os.getcwd(), old_cwd, APP_DIR, target_dir_norm]))
        before = {}
        for d in watch_dirs:
            before.update(_snapshot_dir(d))

        try:
            code_to_run = str(code or "")
            try:
                compile(code_to_run, "<script>", "exec")
            except SyntaxError:
                extracted = _extract_python_blocks_from_markdown(code_to_run)
                if extracted:
                    log_output("ℹ️ Обнаружены блоки кода в Markdown — выполняю код из блоков.")
                    code_to_run = extracted

            namespace = {
                "__name__": "__main__",
                "__file__": source_path or os.path.join(work_dir, "script.py"),
                "__builtins__": __builtins__,
                "create_docx": lambda t, txt, lst, tbl, fn: create_docx(t, txt, lst, tbl, fn, target_dir=target_dir_norm),
                "create_pptx": lambda t, txt, lst, fn: create_pptx(t, txt, lst, fn, target_dir=target_dir_norm),
                "create_xlsx": lambda tbl, fn: create_xlsx(tbl, fn, target_dir=target_dir_norm),
                "create_pdf": lambda t, txt, lst, fn: create_pdf(t, txt, lst, fn, target_dir=target_dir_norm),
                "create_txt": lambda t, txt, lst, tbl, fn: create_txt(t, txt, lst, tbl, fn, target_dir=target_dir_norm),
                "create_md": lambda t, txt, lst, tbl, fn: create_md(t, txt, lst, tbl, fn, target_dir=target_dir_norm),
                "create_html": lambda t, txt, lst, tbl, fn: create_html(t, txt, lst, tbl, fn, target_dir=target_dir_norm),
                "create_from_text": lambda raw, fmt="docx", fn="layout": create_from_text(raw, fmt, fn, target_dir=target_dir_norm),
                "parse_layout": parse_layout_from_text,
                "save_file_to_download": lambda fn, data, is_binary=True: save_file_to_target(fn, data, is_binary=is_binary, target_dir=target_dir_norm),
                "save_file": lambda fn, data, is_binary=True: save_file_to_target(fn, data, is_binary=is_binary, target_dir=target_dir_norm),
                "SAVE_DIR": target_dir_norm,
                "DOWNLOAD_DIR": target_dir_norm,
                "SOURCE_FILE": source_path or "",
                "SOURCE_DIR": os.path.dirname(source_path) if source_path else "",
                "SOURCE_TEXT": str(code or ""),
                "os": os,
                "json": json,
                "csv": csv,
                "re": re,
                "print": lambda *args, **kwargs: log_output(f"[PRINT] {' '.join(map(str, args))}"),
            }
            sys.stdout = io.StringIO()
            exec(code_to_run, namespace)
            output = sys.stdout.getvalue()
            if output:
                log_output(f"Вывод скрипта:\n{output}")

            recorded = _pop_recent_created_files()
            recorded_names = {item["filename"] for item in recorded}

            for d in watch_dirs:
                for full_path, stat_val in _snapshot_dir(d).items():
                    if full_path not in before or before[full_path] != stat_val:
                        fname = os.path.basename(full_path)
                        if fname.lower().endswith(GENERATED_EXTENSIONS) and fname not in recorded_names:
                            if os.path.normpath(d) == target_dir_norm and staging_dir is None:
                                with open(full_path, "rb") as fp:
                                    data = fp.read()
                                recorded.append(
                                    {
                                        "filename": fname,
                                        "bytes": data,
                                        "path": full_path,
                                        "in_target_dir": True,
                                    }
                                )
                            else:
                                save_file_to_target(
                                    fname,
                                    full_path,
                                    is_binary=False,
                                    target_dir=target_dir_norm,
                                )
                                recorded.extend(_pop_recent_created_files())
                            recorded_names.add(fname)

            for item in recorded:
                if item["in_target_dir"]:
                    log_output(f"📁 Сохранено: {item['path']}")
                else:
                    log_output(
                        f"⚠️ Сохранено во внутреннюю память ({item['path']}). "
                        f"Нажмите «Сохранить {item['filename']} в...»."
                    )

            if on_files_ready and recorded:
                on_files_ready(recorded)

            log_output("✅ Скрипт выполнен успешно.")
        except Exception as e:
            log_output(f"❌ Ошибка: {e}\n{traceback.format_exc()}")
        finally:
            sys.stdout = old_stdout
            if switched_cwd:
                try:
                    os.chdir(old_cwd)
                except Exception:
                    pass
            if staging_dir:
                shutil.rmtree(staging_dir, ignore_errors=True)

    threading.Thread(target=run, daemon=True).start()


def main(page: ft.Page):
    is_android = page.platform == ft.PagePlatform.ANDROID
    is_desktop = (
        page.platform in (ft.PagePlatform.WINDOWS, ft.PagePlatform.MACOS, ft.PagePlatform.LINUX)
        or (sys.platform == "win32" and not page.web)
    )

    page.title = "PyDocx"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.theme = ft.Theme(
        color_scheme_seed=PRIMARY_COLOR,
        font_family="Segoe UI" if is_desktop else None,
    )
    page.bgcolor = WIN_BG_COLOR if is_desktop else SURFACE_COLOR
    page.padding = 0

    if is_desktop:
        page.window.width = 1120
        page.window.height = 760
        page.window.min_width = 760
        page.window.min_height = 540
        if os.path.exists(ICON_PATH):
            page.window.icon = "icon.png"

    settings = load_app_settings()

    ph = fph.PermissionHandler() if (fph is not None and is_android) else None
    if ph is not None:
        page.services.append(ph)

    file_picker = ft.FilePicker()
    page.services.append(file_picker)

    # Источники файлов (чтобы понимать, открыт файл или код вставлен вручную)
    script_source = {"path": "", "dir": "", "name": "", "is_cache": False}
    layout_source = {"path": "", "dir": "", "name": "", "is_cache": False}

    def get_effective_target_dir(source_info):
        default_dir = resolve_save_dir(settings["default_dir"], is_android=is_android)
        if (
            settings["save_next_to_source"]
            and source_info.get("dir")
            and not source_info.get("is_cache")
            and os.path.isdir(source_info["dir"])
        ):
            return source_info["dir"], True
        return default_dir, False

    def make_field(label, code_font=False, **kwargs):
        return ft.TextField(
            label=label,
            bgcolor=SURFACE_COLOR,
            border_color=BORDER_COLOR,
            focused_border_color=PRIMARY_COLOR,
            border_radius=6,
            text_size=13 if (code_font and is_desktop) else 14,
            text_style=ft.TextStyle(font_family="Consolas") if (code_font and is_desktop) else None,
            content_padding=ft.Padding.symmetric(horizontal=12, vertical=10),
            **kwargs,
        )

    def make_elevated_btn(text, on_click, icon=None):
        return ft.Button(
            content=text,
            icon=icon,
            on_click=on_click,
            bgcolor=BTN_BG,
            color=PRIMARY_COLOR,
        )

    def make_outlined_btn(text, on_click, icon=None):
        return ft.OutlinedButton(
            content=text,
            icon=icon,
            on_click=on_click,
        )

    def wrap_desktop_card(controls, expand=False, padding=16):
        return ft.Container(
            bgcolor=SURFACE_COLOR,
            border=ft.Border.all(1, CARD_BORDER_COLOR),
            border_radius=8,
            padding=padding,
            expand=expand,
            content=ft.Column(
                controls=controls,
                horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
                spacing=12,
                expand=expand,
            ),
        )

    def show_snack(msg):
        page.show_dialog(ft.SnackBar(ft.Text(msg)))

    def make_export_button(item):
        async def on_save_click(e):
            try:
                saved = await file_picker.save_file(
                    dialog_title="Сохранить документ",
                    file_name=item["filename"],
                    src_bytes=item["bytes"],
                )
                if saved:
                    show_snack(f"Сохранено: {saved}")
                else:
                    show_snack(f"Выбор пути для {item['filename']} завершён.")
            except Exception as ex:
                show_snack(f"Ошибка сохранения: {ex}")

        return ft.Button(
            content=f"Сохранить {item['filename']} в...",
            icon=ft.Icons.SAVE_ALT if is_desktop else None,
            on_click=on_save_click,
            bgcolor="#eef4fc",
            color=PRIMARY_COLOR,
        )

    def populate_save_buttons(row_control, items):
        row_control.controls = [make_export_button(item) for item in items]
        page.update()

    script_log = ft.Text(
        "",
        selectable=True,
        size=12,
        color=TEXT_COLOR,
        font_family="Consolas" if is_desktop else None,
    )
    layout_log = ft.Text(
        "",
        selectable=True,
        size=12,
        color=TEXT_COLOR,
        font_family="Consolas" if is_desktop else None,
    )
    script_save_buttons = ft.Row(controls=[], wrap=True, spacing=8, run_spacing=8)
    layout_save_buttons = ft.Row(controls=[], wrap=True, spacing=8, run_spacing=8)

    def log_to_ui(message):
        script_log.value += f"{message}\n"
        page.update()

    def log_layout(message):
        layout_log.value += f"{message}\n"
        page.update()

    # =========================================================
    # РАЗДЕЛ 1 (ГЛАВНЫЙ): СВОЙ СКРИПТ
    # =========================================================
    script_code_field = make_field(
        "Код скрипта (Python / Markdown-блоки)",
        code_font=True,
        multiline=True,
        min_lines=12 if not is_desktop else 18,
        max_lines=20 if not is_desktop else 28,
        hint_text="Доступны: create_docx, create_pptx, create_xlsx, create_pdf, create_from_text, SAVE_DIR",
        value="# Пример:\n# create_docx('Заголовок', 'Текст', 'Пункт 1\\nПункт 2', 'A,B\\n1,2', 'my.docx')\nprint('Привет из PyDocx!')",
    )

    script_source_status = ft.Text("", size=12, color=MUTED_COLOR)

    def refresh_source_banners():
        default_resolved = resolve_save_dir(settings["default_dir"], is_android=is_android)
        if script_source["name"]:
            if settings["save_next_to_source"] and not script_source["is_cache"] and script_source["dir"]:
                script_source_status.value = (
                    f"Файл: {script_source['name']} • сохранение рядом: {script_source['dir']}"
                )
            elif script_source["is_cache"]:
                script_source_status.value = (
                    f"Файл: {script_source['name']} (из кэша) • сохранение в: {default_resolved}"
                )
            else:
                script_source_status.value = (
                    f"Файл: {script_source['name']} • сохранение в: {default_resolved}"
                )
        else:
            script_source_status.value = (
                f"Режим: ручной ввод / вставка • папка по умолчанию: {default_resolved}"
            )
        reset_script_source_btn.visible = bool(script_source["name"])

    def reset_script_source(e=None):
        script_source.update({"path": "", "dir": "", "name": "", "is_cache": False})
        refresh_source_banners()
        page.update()

    reset_script_source_btn = make_outlined_btn(
        "Сбросить привязку к файлу",
        reset_script_source,
    )
    reset_script_source_btn.visible = False

    async def pick_script_file(e):
        files = await file_picker.pick_files(
            allow_multiple=False,
            allowed_extensions=SCRIPT_ALLOWED_EXTENSIONS,
        )
        if files:
            f = files[0]
            if f.path and os.path.exists(f.path):
                with open(f.path, "r", encoding="utf-8", errors="replace") as fp:
                    script_code_field.value = fp.read()
                parent_dir = os.path.dirname(f.path)
                is_cache = _is_cache_or_temp_dir(parent_dir)
                script_source.update(
                    {
                        "path": f.path,
                        "dir": parent_dir,
                        "name": f.name,
                        "is_cache": is_cache,
                    }
                )
                refresh_source_banners()
                script_log.value = f"Загружен файл: {f.name}\n"
                page.update()

    def clear_script_editor(e):
        script_code_field.value = ""
        reset_script_source()

    def run_script(e):
        target_dir, is_src_dir = get_effective_target_dir(script_source)
        script_log.value = f"Выполнение... (целевая папка: {target_dir})\n"
        if script_source["name"] and script_source["is_cache"] and settings["save_next_to_source"]:
            script_log.value += (
                "ℹ️ Android передал временную копию файла; результат пойдёт в папку по умолчанию "
                "или через кнопку «Сохранить ... в...».\n"
            )
        elif is_src_dir:
            script_log.value += f"ℹ️ Активно сохранение рядом с исходным файлом: {target_dir}\n"
        script_save_buttons.controls = []
        page.update()
        execute_user_script(
            script_code_field.value,
            log_to_ui,
            target_dir=target_dir,
            source_path=script_source["path"],
            on_files_ready=lambda items: populate_save_buttons(script_save_buttons, items),
        )

    script_actions_row = ft.Row(
        controls=[
            make_elevated_btn("Выполнить", run_script, icon=ft.Icons.PLAY_ARROW if is_desktop else None),
            make_outlined_btn(
                "Загрузить файл (.py, .txt, .md, .json, .csv, .yaml, .js)",
                pick_script_file,
                icon=ft.Icons.UPLOAD_FILE if is_desktop else None,
            ),
            make_outlined_btn("Очистить", clear_script_editor),
        ],
        wrap=True,
        spacing=8,
        run_spacing=8,
    )

    script_source_bar = ft.Row(
        controls=[script_source_status, reset_script_source_btn],
        wrap=True,
        spacing=8,
        vertical_alignment=ft.CrossAxisAlignment.CENTER,
    )

    script_log_box = ft.Container(
        bgcolor=LOG_BG,
        border=ft.Border.all(1, DIVIDER_COLOR),
        border_radius=6,
        padding=10,
        expand=is_desktop,
        content=ft.Column(
            controls=[script_log],
            scroll=ft.ScrollMode.AUTO,
            expand=is_desktop,
        )
        if is_desktop
        else script_log,
    )

    if is_desktop:
        script_tab = ft.Container(
            padding=20,
            expand=True,
            content=ft.Row(
                expand=True,
                spacing=16,
                vertical_alignment=ft.CrossAxisAlignment.STRETCH,
                controls=[
                    wrap_desktop_card(
                        expand=3,
                        controls=[
                            ft.Text("Свой скрипт", size=20, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                            script_source_bar,
                            script_actions_row,
                            ft.Container(content=script_code_field, expand=True),
                        ],
                    ),
                    wrap_desktop_card(
                        expand=2,
                        controls=[
                            ft.Text("Лог выполнения и экспорт", size=16, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                            script_log_box,
                            script_save_buttons,
                        ],
                    ),
                ],
            ),
        )
    else:
        script_tab = ft.Container(
            padding=ft.Padding.only(left=16, right=16, top=16, bottom=28),
            expand=True,
            bgcolor=SURFACE_COLOR,
            content=ft.Column(
                controls=[
                    ft.Text("Свой скрипт", size=20, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                    script_source_bar,
                    script_code_field,
                    script_actions_row,
                    ft.Divider(color=DIVIDER_COLOR, height=1),
                    ft.Text("Лог выполнения:", size=14, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                    script_log_box,
                    script_save_buttons,
                ],
                horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
                scroll=ft.ScrollMode.AUTO,
                expand=True,
                spacing=12,
            ),
        )

    # =========================================================
    # РАЗДЕЛ 2: ВЕРСТКА ИЗ ФАЙЛА
    # =========================================================
    dropzone_title = ft.Text(
        "Нажмите, чтобы выбрать текстовый файл / скрипт",
        size=14,
        weight=ft.FontWeight.W_500,
        color=TEXT_COLOR,
        text_align=ft.TextAlign.CENTER,
    )
    layout_content_field = make_field(
        "Содержимое текстового файла / разметка",
        code_font=True,
        multiline=True,
        min_lines=8 if not is_desktop else 14,
        max_lines=14 if not is_desktop else 22,
        hint_text="# Заголовок\nТекст\n- Пункт 1\n| A | B |\n| --- | --- |\n| 1 | 2 |",
    )
    layout_filename_field = make_field(
        "Имя итогового файла (без расширения)",
        value="layout",
    )

    async def pick_layout_file(e):
        files = await file_picker.pick_files(
            allow_multiple=False,
            allowed_extensions=["txt", "md", "py", "js", "json", "csv", "yaml", "yml", "log"],
        )
        if files:
            f = files[0]
            if f.path and os.path.exists(f.path):
                with open(f.path, "r", encoding="utf-8", errors="replace") as fp:
                    layout_content_field.value = fp.read()
                parent_dir = os.path.dirname(f.path)
                layout_source.update(
                    {
                        "path": f.path,
                        "dir": parent_dir,
                        "name": f.name,
                        "is_cache": _is_cache_or_temp_dir(parent_dir),
                    }
                )
                dropzone_title.value = f"Загружен файл: {f.name}"
                layout_filename_field.value = os.path.splitext(f.name)[0] or "layout"
                layout_log.value = f"Загружен файл: {f.name}\n"
                page.update()

    def load_sample_layout(e):
        layout_source.update({"path": "", "dir": "", "name": "example.md", "is_cache": False})
        dropzone_title.value = "Загружен пример: example.md"
        layout_content_field.value = SAMPLE_LAYOUT_TEXT
        layout_filename_field.value = "report"
        layout_log.value = "Загружен пример текстовой разметки.\n"
        page.update()

    def make_typeset_handler(fmt):
        def handler(e):
            try:
                _pop_recent_created_files()
                raw = layout_content_field.value
                if not raw or not raw.strip():
                    raise ValueError("Текстовый файл пуст.")
                if _looks_like_executable_script(raw, layout_source["name"]):
                    target_dir, _ = get_effective_target_dir(layout_source)
                    layout_log.value = (
                        f"ℹ️ Обнаружен исполняемый Python-скрипт — выполняю генерацию документа...\n"
                        f"(целевая папка: {target_dir})\n"
                    )
                    layout_save_buttons.controls = []
                    page.update()
                    execute_user_script(
                        raw,
                        log_layout,
                        target_dir=target_dir,
                        source_path=layout_source["path"],
                        on_files_ready=lambda items: populate_save_buttons(layout_save_buttons, items),
                    )
                    return

                base = os.path.splitext(layout_filename_field.value.strip() or "layout")[0]
                t, txt, lst, tbl = parse_layout_from_text(raw, layout_source["name"])
                target_dir, _ = get_effective_target_dir(layout_source)
                fname = f"{base}.{fmt}"
                if fmt == "docx":
                    res = create_docx(t, txt, lst, tbl, fname, target_dir=target_dir)
                elif fmt == "pptx":
                    res = create_pptx(t, txt, lst, fname, target_dir=target_dir)
                elif fmt == "xlsx":
                    res = create_xlsx(tbl, fname, target_dir=target_dir)
                elif fmt == "pdf":
                    res = create_pdf(t, txt, lst, fname, target_dir=target_dir)
                elif fmt == "md":
                    res = create_md(t, txt, lst, tbl, fname, target_dir=target_dir)
                elif fmt == "html":
                    res = create_html(t, txt, lst, tbl, fname, target_dir=target_dir)
                else:
                    res = create_txt(t, txt, lst, tbl, fname, target_dir=target_dir)
                layout_log.value = f"{res}\n"
                populate_save_buttons(layout_save_buttons, _pop_recent_created_files())
                show_snack(res)
            except Exception as ex:
                layout_log.value = f"❌ Ошибка: {ex}\n"
                show_snack(f"Ошибка: {ex}")

        return handler

    def run_layout_as_script(e):
        target_dir, _ = get_effective_target_dir(layout_source)
        layout_log.value = f"Выполнение... (целевая папка: {target_dir})\n"
        layout_save_buttons.controls = []
        page.update()
        execute_user_script(
            layout_content_field.value,
            log_layout,
            target_dir=target_dir,
            source_path=layout_source["path"],
            on_files_ready=lambda items: populate_save_buttons(layout_save_buttons, items),
        )

    dropzone_card = ft.Container(
        bgcolor=DROPZONE_BG,
        border=ft.Border.all(1, BORDER_COLOR),
        border_radius=8,
        padding=16,
        on_click=pick_layout_file,
        content=ft.Column(
            controls=[
                dropzone_title,
                ft.Text(
                    "Поддерживаются: .txt, .md, .py, .js, .json, .csv, .yaml",
                    size=12,
                    color=MUTED_COLOR,
                    text_align=ft.TextAlign.CENTER,
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=6,
        ),
    )

    layout_pick_row = ft.Row(
        controls=[
            make_outlined_btn(
                "Выбрать файл (.txt, .md, .py, .json, .csv, .yaml)",
                pick_layout_file,
                icon=ft.Icons.UPLOAD_FILE if is_desktop else None,
            ),
            make_outlined_btn("Загрузить пример разметки", load_sample_layout),
        ],
        wrap=True,
        spacing=8,
        run_spacing=8,
    )

    layout_export_row = ft.Row(
        controls=[
            make_elevated_btn("Word (.docx)", make_typeset_handler("docx")),
            make_elevated_btn("PDF (.pdf)", make_typeset_handler("pdf")),
            make_elevated_btn("PowerPoint (.pptx)", make_typeset_handler("pptx")),
            make_elevated_btn("Excel (.xlsx)", make_typeset_handler("xlsx")),
            make_elevated_btn("Markdown (.md)", make_typeset_handler("md")),
            make_elevated_btn("HTML (.html)", make_typeset_handler("html")),
            make_elevated_btn("Текст (.txt)", make_typeset_handler("txt")),
            make_outlined_btn("Выполнить как скрипт", run_layout_as_script),
        ],
        wrap=True,
        spacing=8,
        run_spacing=8,
    )

    layout_log_box = ft.Container(
        bgcolor=LOG_BG,
        border=ft.Border.all(1, DIVIDER_COLOR),
        border_radius=6,
        padding=10,
        expand=is_desktop,
        content=ft.Column(
            controls=[layout_log],
            scroll=ft.ScrollMode.AUTO,
            expand=is_desktop,
        )
        if is_desktop
        else layout_log,
    )

    if is_desktop:
        layout_tab = ft.Container(
            padding=20,
            expand=True,
            content=ft.Row(
                expand=True,
                spacing=16,
                vertical_alignment=ft.CrossAxisAlignment.STRETCH,
                controls=[
                    wrap_desktop_card(
                        expand=3,
                        controls=[
                            ft.Text(
                                "Верстка из текстовых файлов и скриптов",
                                size=20,
                                weight=ft.FontWeight.BOLD,
                                color=TEXT_COLOR,
                            ),
                            dropzone_card,
                            layout_pick_row,
                            ft.Container(content=layout_content_field, expand=True),
                        ],
                    ),
                    wrap_desktop_card(
                        expand=2,
                        controls=[
                            ft.Text("Формат и экспорт", size=16, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                            layout_filename_field,
                            layout_export_row,
                            ft.Divider(color=DIVIDER_COLOR, height=1),
                            ft.Text("Результат верстки:", size=14, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                            layout_log_box,
                            layout_save_buttons,
                        ],
                    ),
                ],
            ),
        )
    else:
        layout_tab = ft.Container(
            padding=ft.Padding.only(left=16, right=16, top=16, bottom=28),
            expand=True,
            bgcolor=SURFACE_COLOR,
            content=ft.Column(
                controls=[
                    ft.Text(
                        "Верстка из текстовых файлов и скриптов",
                        size=20,
                        weight=ft.FontWeight.BOLD,
                        color=TEXT_COLOR,
                    ),
                    dropzone_card,
                    layout_pick_row,
                    layout_content_field,
                    layout_filename_field,
                    layout_export_row,
                    ft.Divider(color=DIVIDER_COLOR, height=1),
                    ft.Text("Результат верстки:", size=14, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                    layout_log_box,
                    layout_save_buttons,
                ],
                horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
                scroll=ft.ScrollMode.AUTO,
                expand=True,
                spacing=12,
            ),
        )

    # =========================================================
    # РАЗДЕЛ 3: НАСТРОЙКИ
    # =========================================================
    default_dir_field = make_field(
        "Папка сохранения по умолчанию",
        value=settings["default_dir"],
        hint_text="Documents/PyDocx",
    )
    resolved_path_hint = ft.Text("", size=12, color=MUTED_COLOR)
    settings_log = ft.Text("", size=12, color=TEXT_COLOR)

    def update_resolved_path_hint():
        resolved = resolve_save_dir(default_dir_field.value, is_android=is_android)
        resolved_path_hint.value = f"Полный путь: {resolved}"

    def on_default_dir_change(e):
        update_resolved_path_hint()
        page.update()

    default_dir_field.on_change = on_default_dir_change

    save_next_to_source_switch = ft.Switch(
        label="Сохранять рядом с исходным файлом (если открыт файл)",
        value=settings["save_next_to_source"],
    )

    def on_toggle_source_save(e):
        settings["save_next_to_source"] = bool(save_next_to_source_switch.value)
        save_app_settings(settings)
        refresh_source_banners()
        page.update()

    save_next_to_source_switch.on_change = on_toggle_source_save

    def apply_default_dir_setting(e):
        val = (default_dir_field.value or "").strip() or DEFAULT_REL_SAVE_DIR
        default_dir_field.value = val
        settings["default_dir"] = val
        save_app_settings(settings)
        update_resolved_path_hint()
        refresh_source_banners()
        resolved = resolve_save_dir(val, is_android=is_android)
        settings_log.value = f"Папка по умолчанию обновлена: {resolved}"
        show_snack("Настройки сохранены.")
        page.update()

    def reset_default_dir_setting(e):
        default_dir_field.value = DEFAULT_REL_SAVE_DIR
        settings["default_dir"] = DEFAULT_REL_SAVE_DIR
        save_app_settings(settings)
        update_resolved_path_hint()
        refresh_source_banners()
        settings_log.value = f"Сброшено на путь по умолчанию: {DEFAULT_REL_SAVE_DIR}"
        show_snack("Восстановлена папка Documents/PyDocx")
        page.update()

    async def pick_default_directory(e):
        try:
            picked = await file_picker.get_directory_path(
                dialog_title="Выберите папку сохранения по умолчанию"
            )
            if picked:
                default_dir_field.value = picked
                settings["default_dir"] = picked
                save_app_settings(settings)
                update_resolved_path_hint()
                refresh_source_banners()
                settings_log.value = f"Выбрана папка: {picked}"
                show_snack(f"Папка выбрана: {picked}")
                page.update()
        except Exception as ex:
            settings_log.value = f"Не удалось выбрать папку: {ex}"
            show_snack(f"Ошибка выбора папки: {ex}")
            page.update()

    async def request_storage_permission(e):
        if ph is None:
            msg = "На этой платформе отдельный запрос разрешений не требуется."
            settings_log.value = msg
            show_snack(msg)
            page.update()
            return
        try:
            status = await ph.request(fph.Permission.MANAGE_EXTERNAL_STORAGE)
            if status == fph.PermissionStatus.GRANTED:
                settings_log.value = "Разрешение на запись во внешнее хранилище получено."
                show_snack("Разрешение получено.")
            else:
                settings_log.value = (
                    f"Статус разрешения: {status}. Вы всегда можете выбрать папку через «Сохранить ... в...»."
                )
                show_snack("Разрешение не получено. Используйте «Сохранить ... в...».")
        except Exception as ex:
            settings_log.value = f"Ошибка при запросе разрешения: {ex}"
        page.update()

    update_resolved_path_hint()
    refresh_source_banners()

    settings_controls = [
        ft.Text("Настройки", size=20, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
        ft.Text(
            "Папка сохранения по умолчанию",
            size=15,
            weight=ft.FontWeight.BOLD,
            color=TEXT_COLOR,
        ),
        default_dir_field,
        resolved_path_hint,
        ft.Row(
            controls=[
                make_elevated_btn(
                    "Сохранить путь",
                    apply_default_dir_setting,
                    icon=ft.Icons.CHECK if is_desktop else None,
                ),
                make_outlined_btn(
                    "Выбрать папку...",
                    pick_default_directory,
                    icon=ft.Icons.FOLDER_OPEN if is_desktop else None,
                ),
                make_outlined_btn("Сбросить (Documents/PyDocx)", reset_default_dir_setting),
            ],
            wrap=True,
            spacing=8,
            run_spacing=8,
        ),
        ft.Divider(color=DIVIDER_COLOR, height=1),
        ft.Text(
            "Поведение при запуске скриптов и файлов",
            size=15,
            weight=ft.FontWeight.BOLD,
            color=TEXT_COLOR,
        ),
        save_next_to_source_switch,
        ft.Text(
            "Если открыт файл с диска — результат сохраняется в его папку. "
            "Если код скопирован/введён вручную или открыт из кэша — сохранение идёт в папку по умолчанию (Documents/PyDocx).",
            size=12,
            color=MUTED_COLOR,
        ),
        ft.Divider(color=DIVIDER_COLOR, height=1),
        ft.Text(
            "Разрешения хранилища (Android)",
            size=15,
            weight=ft.FontWeight.BOLD,
            color=TEXT_COLOR,
        ),
        ft.Text(
            "Для прямого создания файлов в Documents/PyDocx на Android требуется разрешение на доступ к файлам. "
            "Без него всегда доступна кнопка «Сохранить ... в...» через системный проводник.",
            size=12,
            color=MUTED_COLOR,
        ),
        ft.Row(
            controls=[
                make_elevated_btn("Запросить разрешение на запись", request_storage_permission),
            ],
            wrap=True,
        ),
        settings_log,
    ]

    if is_desktop:
        settings_tab = ft.Container(
            padding=20,
            expand=True,
            alignment=ft.Alignment.TOP_CENTER,
            content=ft.Container(
                width=720,
                bgcolor=SURFACE_COLOR,
                border=ft.Border.all(1, CARD_BORDER_COLOR),
                border_radius=8,
                padding=24,
                content=ft.Column(
                    controls=settings_controls,
                    horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
                    scroll=ft.ScrollMode.AUTO,
                    spacing=12,
                ),
            ),
        )
    else:
        settings_tab = ft.Container(
            padding=ft.Padding.only(left=16, right=16, top=16, bottom=28),
            expand=True,
            bgcolor=SURFACE_COLOR,
            content=ft.Column(
                controls=settings_controls,
                horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
                scroll=ft.ScrollMode.AUTO,
                expand=True,
                spacing=12,
            ),
        )

    # =========================================================
    # СБОРКА ИНТЕРФЕЙСА (WINDOWS DESKTOP vs ANDROID MOBILE)
    # =========================================================
    if is_desktop:
        desktop_views = [script_tab, layout_tab, settings_tab]
        desktop_content_host = ft.Container(content=desktop_views[0], expand=True)

        def on_rail_change(e):
            idx = e.control.selected_index
            desktop_content_host.content = desktop_views[idx]
            page.update()

        rail_leading = ft.Container(
            padding=ft.Padding.symmetric(vertical=12),
            content=ft.Column(
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=6,
                controls=[
                    ft.Image(src="icon.png", width=40, height=40)
                    if os.path.exists(ICON_PATH)
                    else ft.Icon(ft.Icons.CODE, color=PRIMARY_COLOR, size=32),
                    ft.Text("PyDocx", size=14, weight=ft.FontWeight.BOLD, color=TEXT_COLOR),
                ],
            ),
        )

        rail = ft.NavigationRail(
            selected_index=0,
            label_type=ft.NavigationRailLabelType.ALL,
            min_width=108,
            bgcolor=SURFACE_COLOR,
            leading=rail_leading,
            destinations=[
                ft.NavigationRailDestination(
                    icon=ft.Icons.CODE_OUTLINED,
                    selected_icon=ft.Icons.CODE,
                    label="Свой скрипт",
                ),
                ft.NavigationRailDestination(
                    icon=ft.Icons.DESCRIPTION_OUTLINED,
                    selected_icon=ft.Icons.DESCRIPTION,
                    label="Верстка",
                ),
                ft.NavigationRailDestination(
                    icon=ft.Icons.SETTINGS_OUTLINED,
                    selected_icon=ft.Icons.SETTINGS,
                    label="Настройки",
                ),
            ],
            on_change=on_rail_change,
        )

        page.add(
            ft.Row(
                expand=True,
                spacing=0,
                controls=[
                    rail,
                    ft.VerticalDivider(width=1, color=CARD_BORDER_COLOR),
                    desktop_content_host,
                ],
            )
        )
    else:
        tabs = ft.Tabs(
            length=3,
            selected_index=0,
            expand=True,
            content=ft.Column(
                expand=True,
                spacing=0,
                controls=[
                    ft.TabBar(
                        indicator_color=PRIMARY_COLOR,
                        label_color=PRIMARY_COLOR,
                        unselected_label_color=MUTED_COLOR,
                        tabs=[
                            ft.Tab(label="Свой скрипт"),
                            ft.Tab(label="Верстка из файла"),
                            ft.Tab(label="Настройки"),
                        ],
                    ),
                    ft.TabBarView(
                        expand=True,
                        controls=[script_tab, layout_tab, settings_tab],
                    ),
                ],
            ),
        )
        page.add(ft.SafeArea(content=tabs, expand=True))


if __name__ == "__main__":
    ft.run(main, assets_dir="assets")
