"""Office-Werkzeuge fuer nextcloud-mcp: Dokumente, Tabellen, Praesentationen und PDF in der Nextcloud anlegen,
bearbeiten und umwandeln.

Die Werkzeuge haengen als App "office" in AVAILABLE_APPS von nextcloud-mcp-server (bos_start.py) und benutzen
dessen Nextcloud-Client des angemeldeten Benutzers (get_client), dessen Scope-Pruefung und dessen Sperre fuer
Pfade mit ausgeschlossenen Tags. Die Dateien bleiben auf dem Server; das Modell schickt nur Anweisungen und Inhalte.

Bearbeitet werden Office Open XML (python-docx, openpyxl, python-pptx) und PDF (pypdf). Alte und offene Formate
(.doc, .odt, .rtf, .xls, .ods, .ppt, .odp) wandelt der Dienst dokumente (LibreOffice, DOKUMENTE_URL) vorher in
Office Open XML und danach zurueck; ebenso das Erzeugen von PDF. Lesen bleibt bei nc_webdav_read_file, das
ueber den Custom Processor (dokumente) Markdown liefert. Zellen einer xlsx setzt edit_xlsx_lossless direkt im XML,
ohne die uebrigen Teile anzufassen.

Anhaenge und Ergebnisse der Codeausfuehrung aus dem Chat kopiert nc_office_save_chat_file im Original in die
Nextcloud (siehe "Dateien aus dem Chat"); den umgekehrten Weg geht nc_office_copy_file_to_chat.
"""
import base64
import copy
import datetime
import io
import json
import logging
import mimetypes
import os
import pathlib
import re
import secrets
import zipfile
import zoneinfo
from typing import Any, Literal
from xml.sax.saxutils import escape as xml_escape

import anyio
import httpx2 as httpx
from mcp.server.mcpserver import Context, MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

from nextcloud_mcp_server.auth import require_scopes
from nextcloud_client.webdav import WEBDAV_SEARCH_MAX_RESULTS
from nextcloud_mcp_server.config import get_settings
from nextcloud_mcp_server.context import get_client
from nextcloud_mcp_server.links import with_links
from nextcloud_mcp_server.models.webdav import FileInfo, SearchFilesResponse
from nextcloud_mcp_server.observability.metrics import instrument_tool
from nextcloud_mcp_server.server.tag_exclusion import get_excluded_file_paths, is_path_excluded

logger = logging.getLogger(__name__)

DOKUMENTE_URL = os.environ.get("DOKUMENTE_URL", "http://dokumente:8080").rstrip("/")

WORD = {"docx", "doc", "odt", "rtf"}
SHEET = {"xlsx", "xls", "ods", "csv"}
SLIDES = {"pptx", "ppt", "odp"}
NATIVE = {"word": "docx", "sheet": "xlsx", "slides": "pptx"}
MIME = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "doc": "application/msword", "xls": "application/vnd.ms-excel", "ppt": "application/vnd.ms-powerpoint",
    "odt": "application/vnd.oasis.opendocument.text", "ods": "application/vnd.oasis.opendocument.spreadsheet",
    "odp": "application/vnd.oasis.opendocument.presentation", "rtf": "application/rtf",
    "pdf": "application/pdf", "csv": "text/csv",
}


def ext(path: str) -> str:
    return pathlib.PurePosixPath(path).suffix.lower().lstrip(".")


def kind(path: str) -> str:
    e = ext(path)
    if e in WORD:
        return "word"
    if e in SHEET:
        return "sheet"
    if e in SLIDES:
        return "slides"
    if e == "pdf":
        return "pdf"
    raise ToolError(f"Unsupported file type {e!r} for {path!r}")


# --- Listenabgleich ---------------------------------------------------------------------------------------------

CHECK_LIST_MAX_TERMS = 1000
CHECK_LIST_MAX_HITS = 20
CHECK_LIST_TERMS_PER_SEARCH = 50
FULLTEXT_MAX_LIMIT = 50
# LibreChat (tool-files) stores files from tool results up to this size
CHAT_FILE_MAX_BYTES = 50 * 1024 * 1024
FULLTEXT_EXCERPTS = 3
FULLTEXT_EXCERPT_CHARS = 300


def parse_terms(raw: Any) -> list[str]:
    """JSON-Liste oder durch Komma, Semikolon oder Zeilenumbruch getrennte Werte; ohne Leere und Doppelte."""
    if isinstance(raw, list):
        items = [str(item) for item in raw]
    else:
        text = str(raw or "").strip()
        items = []
        if text.startswith("["):
            try:
                parsed = json.loads(text)
                items = [str(item) for item in parsed] if isinstance(parsed, list) else []
            except ValueError:
                items = []
        if not items:
            items = re.split(r"[,;\n]+", text)
    seen: set[str] = set()
    out: list[str] = []
    for item in (i.strip() for i in items):
        if item and item.lower() not in seen:
            seen.add(item.lower())
            out.append(item)
    return out


def word_match(term: str, text: str) -> bool:
    """Ganzes Wort: "Plan 12" trifft nicht "Plan 123"."""
    return re.search(r"(?i)(^|[^\w])" + re.escape(term) + r"($|[^\w])", text or "") is not None


# --- Nextcloud und Umwandlung -----------------------------------------------------------------------------------

async def _check_path(client, *paths: str) -> None:
    excluded = await get_excluded_file_paths(client.webdav)
    for p in paths:
        if is_path_excluded(p, excluded):
            raise ToolError(f"Access denied: {p!r} is tagged with an excluded tag")


async def load(client, path: str) -> tuple[bytes, str | None]:
    await _check_path(client, path)
    data, _content_type, etag = await client.webdav.read_file(path)
    return data, etag


async def save(client, path: str, data: bytes, if_match: str | None) -> dict[str, Any]:
    """Schreiben wie nc_webdav_write_file: ohne if_match nur neu anlegen, mit etag nur ohne fremde Aenderung."""
    await _check_path(client, path)
    max_mb = get_settings().webdav_write_max_mb
    if max_mb and len(data) > max_mb * 1024 * 1024:
        raise ToolError(f"Refusing to write {path!r}: {len(data) / 1048576:.1f} MB exceeds WEBDAV_WRITE_MAX_MB")
    content_type = MIME.get(ext(path)) or mimetypes.guess_type(path)[0] or "application/octet-stream"
    result = await client.webdav.write_file(path, data, content_type, if_match=if_match)
    status = result.get("status_code")
    if status in (412, 423):
        hint = ("the file already exists; pass its etag as if_match to overwrite it" if if_match is None
                else "the file was changed by someone else since it was read; read it again")
        raise ToolError(f"{result['message']} ({path!r}): {hint}")
    return {"path": path, "created": status == 201, "size": len(data), "etag": result.get("etag")}


async def convert(data: bytes, src: str, to: str) -> bytes:
    if src == to:
        return data
    async with httpx.AsyncClient(timeout=180) as http:
        r = await http.post(f"{DOKUMENTE_URL}/convert", files={"file": (f"datei.{src}", data)}, data={"to": to})
    if r.status_code != 200:
        raise ToolError(f"Conversion {src} -> {to} failed: {r.text[:300]}")
    return r.content


async def recalculate(data: bytes) -> bytes:
    """xlsx einmal durch LibreOffice: Formeln bekommen berechnete Werte, die Leser ohne Excel brauchen."""
    async with httpx.AsyncClient(timeout=180) as http:
        r = await http.post(f"{DOKUMENTE_URL}/convert", files={"file": ("datei.xlsx", data)}, data={"to": "xlsx"})
    if r.status_code != 200:
        raise ToolError(f"Recalculation failed: {r.text[:300]}")
    return r.content


async def to_native(data: bytes, path: str) -> bytes:
    return await convert(data, ext(path), NATIVE[kind(path)])


async def from_native(data: bytes, path: str, native: str) -> bytes:
    """Aus dem bearbeiteten Office-Open-XML-Format (native) in das Format von path, auch PDF."""
    return await convert(data, native, ext(path))


# --- Text in Laeufen ersetzen (Word und PowerPoint) -------------------------------------------------------------

def replace_in_paragraph(paragraph, find: str, replace: str) -> int:
    """Ersetzt find in einem Absatz; die Formatierung des ersten betroffenen Laufs bleibt."""
    runs = list(paragraph.runs)
    text = "".join(r.text for r in runs)
    if not find or find not in text:
        return 0
    count = 0
    pos = text.rfind(find)
    while pos != -1:
        end = pos + len(find)
        start_off = 0
        first = last = None
        for i, r in enumerate(runs):
            r_end = start_off + len(r.text)
            if first is None and pos < r_end:
                first = (i, pos - start_off)
            if end <= r_end:
                last = (i, end - start_off)
                break
            start_off = r_end
        (fi, fo), (li, lo) = first, last
        if fi == li:
            t = runs[fi].text
            runs[fi].text = t[:fo] + replace + t[lo:]
        else:
            runs[fi].text = runs[fi].text[:fo] + replace
            for r in runs[fi + 1:li]:
                r.text = ""
            runs[li].text = runs[li].text[lo:]
        count += 1
        pos = text.rfind(find, 0, pos) if pos > 0 else -1
    return count


# --- Markdown -> Word -------------------------------------------------------------------------------------------

INLINE = re.compile(r"(\*\*[^*]+\*\*|__[^_]+__|\*[^*\s][^*]*\*|_[^_\s][^_]*_|`[^`]+`|\[[^\]]+\]\([^)]+\))")


def add_inline(paragraph, text: str) -> None:
    for token in INLINE.split(text):
        if not token:
            continue
        if token.startswith(("**", "__")) and len(token) > 4:
            paragraph.add_run(token[2:-2]).bold = True
        elif token.startswith("`"):
            paragraph.add_run(token[1:-1]).font.name = "Courier New"
        elif token.startswith("[") and "](" in token:
            label, url = token[1:-1].split("](", 1)
            paragraph.add_run(f"{label} ({url})")
        elif token[0] in "*_" and token[-1] == token[0] and len(token) > 2:
            paragraph.add_run(token[1:-1]).italic = True
        else:
            paragraph.add_run(token)


def _styled(doc, style: str, fallback: str | None = None):
    for name in (style, fallback):
        if name:
            try:
                return doc.add_paragraph(style=name)
            except KeyError:
                continue
    return doc.add_paragraph()


def _table_rows(lines: list[str]) -> list[list[str]]:
    rows = []
    for line in lines:
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if all(re.fullmatch(r":?-{3,}:?", c) for c in cells if c):
            continue
        rows.append(cells)
    return rows


def markdown_to_docx(doc, markdown: str) -> None:
    """Haengt Markdown an doc an: Ueberschriften, Absaetze, Listen, Tabellen, Zitate, Code, fett/kursiv."""
    lines = markdown.replace("\r\n", "\n").split("\n")
    i, para = 0, []

    def flush():
        if para:
            add_inline(doc.add_paragraph(), " ".join(para))
            para.clear()

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if stripped.startswith("```"):
            flush()
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                p = doc.add_paragraph()
                p.add_run(lines[i]).font.name = "Courier New"
                i += 1
        elif not stripped:
            flush()
        elif m := re.match(r"^(#{1,6})\s+(.*)$", stripped):
            flush()
            try:
                add_inline(doc.add_heading(level=min(len(m.group(1)), 9)), m.group(2))
            except KeyError:
                p = doc.add_paragraph()
                p.add_run(m.group(2)).bold = True
        elif stripped.startswith("|"):
            flush()
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i])
                i += 1
            rows = _table_rows(block)
            if rows:
                cols = max(len(r) for r in rows)
                table = doc.add_table(rows=len(rows), cols=cols)
                try:
                    table.style = "Table Grid"
                except (KeyError, ValueError):
                    pass
                for r, row in enumerate(rows):
                    for c, value in enumerate(row):
                        cell = table.cell(r, c)
                        add_inline(cell.paragraphs[0], value)
                        if r == 0:
                            for run in cell.paragraphs[0].runs:
                                run.bold = True
            continue
        elif m := re.match(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$", line):
            flush()
            level = len(m.group(1).replace("\t", "  ")) // 2
            base = "List Number" if m.group(2)[0].isdigit() else "List Bullet"
            style = base if level == 0 else f"{base} {min(level + 1, 3)}"
            add_inline(_styled(doc, style, base), m.group(3))
        elif stripped.startswith(">"):
            flush()
            add_inline(_styled(doc, "Quote"), stripped.lstrip("> "))
        elif re.fullmatch(r"(-{3,}|\*{3,}|_{3,})", stripped):
            flush()
        else:
            para.append(stripped)
        i += 1
    flush()
    # Wie Word: hinter einer Tabelle am Ende steht ein Absatz; ohne ihn laesst der .doc-Export von LibreOffice
    # die Tabelle weg
    children = [e for e in doc.element.body if not e.tag.endswith("}sectPr")]
    if children and children[-1].tag.endswith("}tbl"):
        doc.add_paragraph()


def _append_markdown(doc, markdown: str) -> list:
    """Haengt Markdown an und gibt die neuen Elemente des Rumpfs zurueck."""
    body = doc.element.body
    n_before = len(body)
    markdown_to_docx(doc, markdown)
    children = list(body)
    # python-docx fuegt vor dem abschliessenden sectPr ein
    if children and children[-1].tag.endswith("}sectPr"):
        return children[n_before - 1:-1]
    return children[n_before:]


def _all_paragraphs(doc):
    def walk(container):
        for p in container.paragraphs:
            yield p
        for t in container.tables:
            for row in t.rows:
                for cell in row.cells:
                    yield from walk(cell)
    yield from walk(doc)
    for section in doc.sections:
        for part in (section.header, section.footer, section.first_page_header, section.first_page_footer):
            if not part.is_linked_to_previous:
                yield from walk(part)


def _set_cell_text(cell, text: str) -> None:
    paragraphs = cell.paragraphs
    first = paragraphs[0]
    for p in paragraphs[1:]:
        p._element.getparent().remove(p._element)
    if first.runs:
        first.runs[0].text = str(text)
        for r in first.runs[1:]:
            r._element.getparent().remove(r._element)
    else:
        first.add_run(str(text))


def _table(doc, number: int):
    if not 1 <= number <= len(doc.tables):
        raise ToolError(f"Table {number} does not exist; the document has {len(doc.tables)} tables")
    return doc.tables[number - 1]


def edit_docx(data: bytes, operations: list[dict[str, Any]]) -> tuple[bytes, list[str]]:
    import docx

    doc = docx.Document(io.BytesIO(data))
    done = []
    for n, op in enumerate(operations, 1):
        name = op.get("op")
        if name == "replace_text":
            count = sum(replace_in_paragraph(p, op["find"], op.get("replace", "")) for p in _all_paragraphs(doc))
            if count == 0:
                raise ToolError(f"Operation {n}: text {op['find']!r} not found")
            done.append(f"replace_text: {count} replacements")
        elif name == "append_markdown":
            _append_markdown(doc, op["markdown"])
            done.append("append_markdown")
        elif name in ("insert_after", "insert_before"):
            anchor = next((p for p in doc.paragraphs if op["anchor"] in p.text), None)
            if anchor is None:
                raise ToolError(f"Operation {n}: no paragraph contains {op['anchor']!r}")
            new = _append_markdown(doc, op["markdown"])
            ref = anchor._element
            if name == "insert_after":
                for e in new:
                    ref.addnext(e)
                    ref = e
            else:
                for e in new:
                    ref.addprevious(e)
            done.append(f"{name}: {len(new)} elements")
        elif name == "delete_paragraphs":
            hits = [p for p in doc.paragraphs if op["contains"] in p.text]
            if not hits:
                raise ToolError(f"Operation {n}: no paragraph contains {op['contains']!r}")
            for p in hits:
                p._element.getparent().remove(p._element)
            done.append(f"delete_paragraphs: {len(hits)}")
        elif name == "set_cell":
            table = _table(doc, int(op["table"]))
            _set_cell_text(table.cell(int(op["row"]) - 1, int(op["column"]) - 1), op.get("text", ""))
            done.append("set_cell")
        elif name == "add_row":
            table = _table(doc, int(op["table"]))
            last = table.rows[-1]._tr
            new_tr = copy.deepcopy(last)
            last.addnext(new_tr)
            row = table.rows[-1]
            values = list(op.get("values", []))
            for c, cell in enumerate(row.cells):
                _set_cell_text(cell, values[c] if c < len(values) else "")
            done.append("add_row")
        else:
            raise ToolError(f"Operation {n}: unknown op {name!r} for documents")
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), done


def create_docx(markdown: str) -> bytes:
    import docx

    doc = docx.Document()
    markdown_to_docx(doc, markdown)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# --- Tabellen ---------------------------------------------------------------------------------------------------

def _cell_value(v):
    return v  # Zeichenketten mit "=" am Anfang sind Formeln; openpyxl schreibt sie als solche


def _autowidth(ws) -> None:
    from openpyxl.utils import get_column_letter

    widths: dict[int, int] = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                widths[cell.column] = max(widths.get(cell.column, 0), len(str(cell.value)))
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(max(width + 2, 8), 60)


def create_xlsx(sheets: list[dict[str, Any]]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font

    if not sheets:
        raise ToolError("At least one sheet is required")
    wb = Workbook()
    wb.remove(wb.active)
    for s in sheets:
        ws = wb.create_sheet(str(s.get("name") or f"Sheet{len(wb.worksheets) + 1}")[:31])
        for row in s.get("rows", []):
            ws.append([_cell_value(v) for v in row])
        if s.get("header", True) and ws.max_row >= 1 and s.get("rows"):
            for cell in ws[1]:
                cell.font = Font(bold=True)
            ws.freeze_panes = "A2"
        _autowidth(ws)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def lossy_parts(data: bytes) -> list[str]:
    """Teile einer xlsx, die openpyxl beim Speichern verliert."""
    names = zipfile.ZipFile(io.BytesIO(data)).namelist()
    found = []
    for prefix, label in (("xl/charts/", "charts"), ("xl/drawings/", "images or shapes"),
                          ("xl/pivotTables/", "pivot tables"), ("xl/vbaProject", "macros")):
        if any(n.startswith(prefix) for n in names):
            found.append(label)
    return found


# Zellen direkt im XML der xlsx setzen: alle anderen Teile (Bilder, Diagramme, Pivot, Makros, bedingte Formate)
# bleiben Byte fuer Byte erhalten. openpyxl liest und schreibt dagegen das ganze Paket neu und verliert sie.

LOSSLESS_OPS = {"set_cells", "append_rows"}
S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REF = re.compile(r"^([A-Z]{1,3})([1-9][0-9]*)$")


def _q(tag: str) -> str:
    return f"{{{S_NS}}}{tag}"


def _col_number(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _col_letters(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _parse_ref(ref: str) -> tuple[int, int]:
    m = REF.match(str(ref).strip().upper().replace("$", ""))
    if not m:
        raise ToolError(f"Invalid cell reference {ref!r}; use A1 notation")
    return _col_number(m.group(1)), int(m.group(2))


def _part_path(base: str, target: str) -> str:
    """Ziel einer Beziehung relativ zum Ordner von base (oder absolut mit /) als Pfad im Zip."""
    if target.startswith("/"):
        return target.lstrip("/")
    return os.path.normpath(os.path.join(os.path.dirname(base), target)).replace(os.sep, "/")


def _rels_path(part: str) -> str:
    d, f = os.path.split(part)
    return f"{d}/_rels/{f}.rels"


class XlsxPackage:
    def __init__(self, data: bytes):
        from lxml import etree

        self.etree = etree
        self.zin = zipfile.ZipFile(io.BytesIO(data))
        self.names = set(self.zin.namelist())
        self.trees: dict[str, Any] = {}
        self.dirty: set[str] = set()
        self.removed: set[str] = set()

    def xml(self, name: str, write: bool = False):
        """Geparster Teil; nur mit write=True wird er beim Speichern neu geschrieben."""
        if write:
            self.dirty.add(name)
        if name not in self.trees:
            parser = self.etree.XMLParser(resolve_entities=False, huge_tree=True)
            self.trees[name] = self.etree.fromstring(self.zin.read(name), parser)
        return self.trees[name]

    def save(self) -> bytes:
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for info in self.zin.infolist():
                if info.filename in self.removed:
                    continue
                if info.filename in self.dirty:
                    data = self.etree.tostring(self.trees[info.filename], xml_declaration=True,
                                               encoding="UTF-8", standalone=True)
                else:
                    data = self.zin.read(info.filename)
                z.writestr(info, data)
        return out.getvalue()

    def sheets(self) -> list[tuple[str, str]]:
        """(Name, Pfad im Zip) je Blatt in der Reihenfolge der Arbeitsmappe."""
        wb = self.xml("xl/workbook.xml")
        rels = {r.get("Id"): r.get("Target") for r in self.xml("xl/_rels/workbook.xml.rels")}
        return [(s.get("name"), _part_path("xl/workbook.xml", rels[s.get(f"{{{R_NS}}}id")]))
                for s in wb.find(_q("sheets"))]

    def active_sheet(self) -> int:
        view = self.xml("xl/workbook.xml").find(f"{_q('bookViews')}/{_q('workbookView')}")
        return int(view.get("activeTab", "0")) if view is not None else 0

    def drop_calc_chain(self) -> None:
        """calcChain verweist auf Formelzellen; nach Aenderungen meldet Excel sonst eine Reparatur."""
        name = "xl/calcChain.xml"
        if name not in self.names:
            return
        self.removed.add(name)
        rels = self.xml("xl/_rels/workbook.xml.rels", write=True)
        for r in list(rels):
            if r.get("Target", "").endswith("calcChain.xml"):
                rels.remove(r)
        ct = self.xml("[Content_Types].xml", write=True)
        for o in list(ct):
            if o.get("PartName") == "/" + name:
                ct.remove(o)

    def full_calc_on_load(self) -> None:
        wb = self.xml("xl/workbook.xml", write=True)
        calc = wb.find(_q("calcPr"))
        if calc is None:
            calc = self.etree.Element(_q("calcPr"))
            anchor = None
            for tag in ("sheets", "functionGroups", "externalReferences", "definedNames"):
                found = wb.find(_q(tag))
                if found is not None:
                    anchor = found
            anchor.addnext(calc)
        calc.set("fullCalcOnLoad", "1")


class SharedStrings:
    def __init__(self, pkg: XlsxPackage):
        self.pkg = pkg
        self.root = None
        self.index: dict[str, int] = {}
        name = "xl/sharedStrings.xml"
        if name in pkg.names:
            self.root = pkg.xml(name)
            for i, si in enumerate(self.root.findall(_q("si"))):
                t = si.find(_q("t"))
                if t is not None and len(si) == 1:
                    self.index.setdefault(t.text or "", i)

    def add(self, text: str) -> int | None:
        """Index des Eintrags, None ohne sharedStrings.xml (dann Inline-Text)."""
        if self.root is None:
            return None
        self.pkg.dirty.add("xl/sharedStrings.xml")
        self.root.set("count", str(int(self.root.get("count", "0")) + 1))
        if text in self.index:
            return self.index[text]
        si = self.pkg.etree.SubElement(self.root, _q("si"))
        t = self.pkg.etree.SubElement(si, _q("t"))
        t.text = text
        if text != text.strip():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        n = len(self.root.findall(_q("si"))) - 1
        self.root.set("uniqueCount", str(n + 1))
        self.index[text] = n
        return n


class SheetEditor:
    def __init__(self, pkg: XlsxPackage, strings: SharedStrings, path: str):
        self.pkg, self.strings, self.path = pkg, strings, path
        self.root = pkg.xml(path, write=True)
        self.data = self.root.find(_q("sheetData"))
        self.rows = {int(r.get("r")): r for r in self.data.findall(_q("row")) if r.get("r")}
        if len(self.rows) != len(self.data.findall(_q("row"))):
            raise ToolError("The sheet has rows without numbers; use allow_loss with a copy instead")
        self.col_styles = []
        cols = self.root.find(_q("cols"))
        for c in (cols if cols is not None else []):
            if c.get("style"):
                self.col_styles.append((int(c.get("min")), int(c.get("max")), c.get("style")))
        self.max_col = self.max_row = 0

    def _row(self, n: int):
        row = self.rows.get(n)
        if row is None:
            row = self.pkg.etree.Element(_q("row"))
            row.set("r", str(n))
            later = [r for r in self.rows if r > n]
            if later:
                self.rows[min(later)].addprevious(row)
            else:
                self.data.append(row)
            self.rows[n] = row
        row.attrib.pop("spans", None)
        return row

    def _cell(self, row, col: int, n: int):
        ref = f"{_col_letters(col)}{n}"
        for c in row.findall(_q("c")):
            r = c.get("r")
            if r == ref:
                return c
            if r and _parse_ref(r)[0] > col:
                new = self.pkg.etree.Element(_q("c"))
                new.set("r", ref)
                c.addprevious(new)
                return new
        new = self.pkg.etree.SubElement(row, _q("c"))
        new.set("r", ref)
        return new

    def _default_style(self, row, col: int) -> str | None:
        if row.get("customFormat") == "1" and row.get("s"):
            return row.get("s")
        for lo, hi, style in self.col_styles:
            if lo <= col <= hi:
                return style
        return None

    def set(self, ref_or_pos, value, style_from: str | None = None) -> None:
        col, n = _parse_ref(ref_or_pos) if isinstance(ref_or_pos, str) else ref_or_pos
        row = self._row(n)
        c = self._cell(row, col, n)
        f = c.find(_q("f"))
        if f is not None and (f.get("t") == "shared" and f.get("ref") or f.get("t") == "array"
                              and f.get("ref", c.get("r")) != c.get("r")):
            raise ToolError(f"Cell {c.get('r')} holds a formula shared with {f.get('ref')}; change that range "
                            "as a whole or use allow_loss with output_path")
        for child in list(c):
            c.remove(child)
        for attr in ("t", "cm", "vm"):
            c.attrib.pop(attr, None)
        if c.get("s") is None:
            style = style_from or self._default_style(row, col)
            if style:
                c.set("s", style)
        E = self.pkg.etree.SubElement
        if value is None or value == "":
            pass
        elif isinstance(value, bool):
            c.set("t", "b")
            E(c, _q("v")).text = "1" if value else "0"
        elif isinstance(value, (int, float)):
            E(c, _q("v")).text = repr(value)
        elif isinstance(value, str) and value.startswith("="):
            E(c, _q("f")).text = value[1:]
        else:
            text = str(value)
            idx = self.strings.add(text)
            if idx is None:
                c.set("t", "inlineStr")
                E(E(c, _q("is")), _q("t")).text = text
            else:
                c.set("t", "s")
                E(c, _q("v")).text = str(idx)
        self.max_col, self.max_row = max(self.max_col, col), max(self.max_row, n)

    def last_data_row(self) -> int:
        """Letzte Zeile mit Inhalt; formatierte leere Zeilen (oft bis 1000) zaehlen nicht."""
        for n in sorted(self.rows, reverse=True):
            for c in self.rows[n].findall(_q("c")):
                if c.find(_q("v")) is not None or c.find(_q("f")) is not None or c.find(_q("is")) is not None:
                    return n
        return 0

    def styles_of_row(self, n: int) -> dict[int, str]:
        row = self.rows.get(n)
        return {} if row is None else {_parse_ref(c.get("r"))[0]: c.get("s") for c in row.findall(_q("c"))
                                       if c.get("r") and c.get("s")}

    def extend_tables(self, last_row: int, new_last: int) -> list[str]:
        """Tabellen (Als Tabelle formatieren), die direkt ueber den angehaengten Zeilen enden, verlaengern."""
        grown = []
        rels_name = _rels_path(self.path)
        if rels_name not in self.pkg.names:
            return grown
        for rel in self.pkg.xml(rels_name):
            if not rel.get("Type", "").endswith("/table"):
                continue
            part = _part_path(self.path, rel.get("Target"))
            table = self.pkg.xml(part)
            m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", table.get("ref", ""))
            if not m or int(m.group(4)) != last_row:
                continue
            ref = f"{m.group(1)}{m.group(2)}:{m.group(3)}{new_last}"
            self.pkg.dirty.add(part)
            table.set("ref", ref)
            auto = table.find(_q("autoFilter"))
            if auto is not None:
                auto.set("ref", ref)
            grown.append(table.get("displayName") or table.get("name"))
        return grown

    def finish(self) -> None:
        dim = self.root.find(_q("dimension"))
        if dim is None or not self.max_row:
            return
        parts = dim.get("ref", "A1").split(":")
        c1, r1 = _parse_ref(parts[0])
        c2, r2 = _parse_ref(parts[-1])
        dim.set("ref", f"{_col_letters(c1)}{r1}:{_col_letters(max(c2, self.max_col))}{max(r2, self.max_row)}")


def edit_xlsx_lossless(data: bytes, operations: list[dict[str, Any]]) -> tuple[bytes, list[str]]:
    try:
        pkg = XlsxPackage(data)
        sheets = pkg.sheets()
    except (KeyError, zipfile.BadZipFile) as e:
        raise ToolError(f"Not a readable xlsx workbook: {e}") from e
    strings = SharedStrings(pkg)
    editors: dict[str, SheetEditor] = {}

    def editor(op) -> SheetEditor:
        name = op.get("sheet")
        if name is None:
            name, path = sheets[min(pkg.active_sheet(), len(sheets) - 1)]
        else:
            match = [p for n, p in sheets if n == name]
            if not match:
                raise ToolError(f"Sheet {name!r} not found; sheets: {[n for n, _ in sheets]}")
            path = match[0]
        if path not in editors:
            editors[path] = SheetEditor(pkg, strings, path)
        return editors[path]

    done = []
    for n, op in enumerate(operations, 1):
        name = op.get("op")
        if name == "set_cells":
            ed = editor(op)
            for ref, value in op["cells"].items():
                ed.set(ref, value)
            done.append(f"set_cells: {len(op['cells'])}")
        elif name == "append_rows":
            ed = editor(op)
            last = ed.last_data_row()
            styles = ed.styles_of_row(last)
            for i, row in enumerate(op["rows"], 1):
                for col, value in enumerate(row, 1):
                    ed.set((col, last + i), value, styles.get(col))
            note = f"append_rows: {len(op['rows'])} after row {last}"
            tables = ed.extend_tables(last, last + len(op["rows"])) if op["rows"] else []
            if tables:
                note += f", table {', '.join(tables)} extended"
            done.append(note)
        else:
            raise ToolError(f"Operation {n}: {name!r} cannot be applied losslessly")
    for ed in editors.values():
        ed.finish()
    pkg.drop_calc_chain()
    pkg.full_calc_on_load()
    return pkg.save(), done


# Formelergebnisse: LibreOffice rechnet beim Oeffnen nur Formeln ohne gespeichertes Ergebnis neu (fullCalcOnLoad
# beachtet es nicht). Eine Kopie ohne Ergebnisse geht deshalb durch LibreOffice; dessen Werte kommen in die
# Formelzellen der unveraenderten Mappe. Excel rechnet beim Oeffnen ohnehin selbst (fullCalcOnLoad).

XL_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"}


def strip_formula_results(data: bytes) -> bytes | None:
    """Kopie ohne gespeicherte Formelergebnisse, None ohne Formeln."""
    pkg = XlsxPackage(data)
    found = False
    for _name, path in pkg.sheets():
        root = pkg.xml(path)
        for c in root.iter(_q("c")):
            if c.find(_q("f")) is None:
                continue
            found = True
            pkg.dirty.add(path)
            v = c.find(_q("v"))
            if v is not None:
                c.remove(v)
            c.attrib.pop("t", None)
    return pkg.save() if found else None


def apply_formula_results(data: bytes, calculated: bytes) -> tuple[bytes, int]:
    import datetime

    from openpyxl import load_workbook
    from openpyxl.utils.datetime import to_excel

    wb = load_workbook(io.BytesIO(calculated), data_only=True)
    pkg = XlsxPackage(data)
    count = 0
    for name, path in pkg.sheets():
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        for c in pkg.xml(path).iter(_q("c")):
            if c.find(_q("f")) is None or not c.get("r"):
                continue
            value = ws[c.get("r")].value
            v = c.find(_q("v"))
            if value is None:
                continue
            if isinstance(value, str) and value in XL_ERRORS and v is not None and c.get("t") != "e":
                continue  # Funktion, die LibreOffice nicht kennt: Ergebnis von Excel behalten
            if v is None:
                v = pkg.etree.SubElement(c, _q("v"))
            if isinstance(value, bool):
                c.set("t", "b")
                v.text = "1" if value else "0"
            elif isinstance(value, (int, float)):
                c.attrib.pop("t", None)
                v.text = repr(value)
            elif isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
                c.attrib.pop("t", None)
                v.text = repr(to_excel(value))
            elif isinstance(value, str) and value in XL_ERRORS:
                c.set("t", "e")
                v.text = value
            else:
                c.set("t", "str")
                v.text = str(value)
            pkg.dirty.add(path)
            count += 1
    return pkg.save(), count


def edit_xlsx(data: bytes, operations: list[dict[str, Any]], allow_loss: bool) -> tuple[bytes, list[str]]:
    from openpyxl import load_workbook

    lost = lossy_parts(data)
    if lost and not allow_loss:
        raise ToolError(
            f"The workbook contains {', '.join(lost)}, which would be lost when saving. Ask the user; if they "
            "agree, repeat with allow_loss=true (ideally with output_path to keep the original).")
    wb = load_workbook(io.BytesIO(data))

    def sheet(op):
        name = op.get("sheet")
        if name is None:
            return wb.active
        if name not in wb.sheetnames:
            raise ToolError(f"Sheet {name!r} not found; sheets: {wb.sheetnames}")
        return wb[name]

    done = []
    for n, op in enumerate(operations, 1):
        name = op.get("op")
        if name == "set_cells":
            ws = sheet(op)
            for ref, value in op["cells"].items():
                ws[ref] = _cell_value(value)
            done.append(f"set_cells: {len(op['cells'])}")
        elif name == "append_rows":
            ws = sheet(op)
            for row in op["rows"]:
                ws.append([_cell_value(v) for v in row])
            done.append(f"append_rows: {len(op['rows'])}")
        elif name == "insert_rows":
            sheet(op).insert_rows(int(op["row"]), int(op.get("count", 1)))
            done.append("insert_rows")
        elif name == "delete_rows":
            sheet(op).delete_rows(int(op["row"]), int(op.get("count", 1)))
            done.append("delete_rows")
        elif name == "add_sheet":
            ws = wb.create_sheet(str(op["name"])[:31])
            for row in op.get("rows", []):
                ws.append([_cell_value(v) for v in row])
            done.append("add_sheet")
        elif name == "rename_sheet":
            sheet(op).title = str(op["new_name"])[:31]
            done.append("rename_sheet")
        elif name == "delete_sheet":
            wb.remove(sheet(op))
            done.append("delete_sheet")
        else:
            raise ToolError(f"Operation {n}: unknown op {name!r} for spreadsheets")
    wb.calculation.fullCalcOnLoad = True
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), done


# --- Praesentationen --------------------------------------------------------------------------------------------

def _layout(prs, with_body: bool):
    for layout in prs.slide_layouts:
        types = {ph.placeholder_format.type for ph in layout.placeholders}
        has_title = any(getattr(t, "name", "") in ("TITLE", "CENTER_TITLE") for t in types)
        has_body = any(getattr(t, "name", "") in ("BODY", "OBJECT") for t in types)
        if has_title and has_body == with_body:
            return layout
    return prs.slide_layouts[1 if with_body else 5] if len(prs.slide_layouts) > 5 else prs.slide_layouts[-1]


def _add_slide(prs, spec: dict[str, Any]):
    bullets = spec.get("bullets") or []
    slide = prs.slides.add_slide(_layout(prs, bool(bullets)))
    if slide.shapes.title is not None:
        slide.shapes.title.text = str(spec.get("title", ""))
    if bullets:
        body = next((ph for ph in slide.placeholders if ph.placeholder_format.idx != 0), None)
        if body is not None:
            tf = body.text_frame
            tf.clear()
            for k, item in enumerate(bullets):
                text = str(item)
                level = (len(text) - len(text.lstrip(" "))) // 2
                p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                p.text = text.strip()
                p.level = min(level, 4)
    if spec.get("notes"):
        slide.notes_slide.notes_text_frame.text = str(spec["notes"])
    return slide


def _slide_paragraphs(prs):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                yield from shape.text_frame.paragraphs
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        yield from cell.text_frame.paragraphs


def _move_slide(prs, old: int, new: int) -> None:
    ids = prs.slides._sldIdLst
    el = list(ids)[old]
    ids.remove(el)
    ids.insert(new, el)


def _slide_index(prs, number) -> int:
    number = int(number)
    if not 1 <= number <= len(prs.slides):
        raise ToolError(f"Slide {number} does not exist; the presentation has {len(prs.slides)} slides")
    return number - 1


def create_pptx(slides: list[dict[str, Any]]) -> bytes:
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for spec in slides:
        _add_slide(prs, spec)
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


def edit_pptx(data: bytes, operations: list[dict[str, Any]]) -> tuple[bytes, list[str]]:
    from pptx import Presentation

    prs = Presentation(io.BytesIO(data))
    done = []
    for n, op in enumerate(operations, 1):
        name = op.get("op")
        if name == "replace_text":
            count = sum(replace_in_paragraph(p, op["find"], op.get("replace", "")) for p in _slide_paragraphs(prs))
            if count == 0:
                raise ToolError(f"Operation {n}: text {op['find']!r} not found")
            done.append(f"replace_text: {count} replacements")
        elif name == "add_slide":
            _add_slide(prs, op)
            if op.get("position"):
                _move_slide(prs, len(prs.slides) - 1, _slide_index(prs, op["position"]))
            done.append("add_slide")
        elif name == "delete_slide":
            idx = _slide_index(prs, op["slide"])
            el = list(prs.slides._sldIdLst)[idx]
            prs.part.drop_rel(el.rId)
            prs.slides._sldIdLst.remove(el)
            done.append("delete_slide")
        elif name == "move_slide":
            _move_slide(prs, _slide_index(prs, op["slide"]), _slide_index(prs, op["to"]))
            done.append("move_slide")
        elif name == "set_notes":
            prs.slides[_slide_index(prs, op["slide"])].notes_slide.notes_text_frame.text = str(op.get("notes", ""))
            done.append("set_notes")
        elif name == "set_title":
            slide = prs.slides[_slide_index(prs, op["slide"])]
            if slide.shapes.title is None:
                raise ToolError(f"Operation {n}: slide {op['slide']} has no title placeholder")
            slide.shapes.title.text = str(op.get("title", ""))
            done.append("set_title")
        else:
            raise ToolError(f"Operation {n}: unknown op {name!r} for presentations")
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue(), done


# --- PDF --------------------------------------------------------------------------------------------------------

def _pages(spec, total: int) -> list[int]:
    pages = []
    for p in spec:
        if isinstance(p, str) and "-" in p:
            a, b = (int(x) for x in p.split("-", 1))
            pages.extend(range(a, b + 1))
        else:
            pages.append(int(p))
    for p in pages:
        if not 1 <= p <= total:
            raise ToolError(f"Page {p} does not exist; the PDF has {total} pages")
    return [p - 1 for p in pages]


def _pdf_select(data: bytes, order: list[int]) -> bytes:
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(io.BytesIO(data))
    writer = PdfWriter()
    for i in order:
        writer.add_page(reader.pages[i])
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def edit_pdf(data: bytes, operations: list[dict[str, Any]], extra: dict[str, bytes]) -> tuple[bytes, list[str]]:
    from pypdf import PdfReader, PdfWriter

    done = []
    for n, op in enumerate(operations, 1):
        name = op.get("op")
        total = len(PdfReader(io.BytesIO(data)).pages)
        if name == "delete_pages":
            drop = set(_pages(op["pages"], total))
            data = _pdf_select(data, [i for i in range(total) if i not in drop])
        elif name in ("keep_pages", "reorder_pages"):
            data = _pdf_select(data, _pages(op["pages"], total))
        elif name == "rotate_pages":
            writer = PdfWriter(clone_from=PdfReader(io.BytesIO(data)))
            for i in _pages(op.get("pages") or [f"1-{total}"], total):
                writer.pages[i].rotate(int(op.get("degrees", 90)))
            out = io.BytesIO()
            writer.write(out)
            data = out.getvalue()
        elif name == "append_pdf":
            writer = PdfWriter(clone_from=PdfReader(io.BytesIO(data)))
            writer.append(PdfReader(io.BytesIO(extra[op["path"]])))
            out = io.BytesIO()
            writer.write(out)
            data = out.getvalue()
        elif name == "fill_form":
            reader = PdfReader(io.BytesIO(data))
            fields = reader.get_fields() or {}
            unknown = [k for k in op["fields"] if k not in fields]
            if unknown:
                raise ToolError(f"Operation {n}: unknown form fields {unknown}; fields: {sorted(fields)}")
            writer = PdfWriter(clone_from=reader)
            for page in writer.pages:
                writer.update_page_form_field_values(page, {k: str(v) for k, v in op["fields"].items()},
                                                     auto_regenerate=False)
            writer.set_need_appearances_writer(True)
            out = io.BytesIO()
            writer.write(out)
            data = out.getvalue()
        else:
            raise ToolError(f"Operation {n}: unknown op {name!r} for PDF")
        done.append(name)
    return data, done


# --- Struktur ---------------------------------------------------------------------------------------------------

def inspect_native(data: bytes, k: str) -> dict[str, Any]:
    if k == "word":
        import docx

        doc = docx.Document(io.BytesIO(data))
        return {
            "paragraphs": len(doc.paragraphs),
            "headings": [p.text for p in doc.paragraphs if p.style is not None and p.style.name.lower().startswith(("heading", "überschrift", "title"))][:100],
            "tables": [{"table": i + 1, "rows": len(t.rows), "columns": len(t.columns),
                        "first_row": [c.text for c in t.rows[0].cells] if t.rows else []}
                       for i, t in enumerate(doc.tables)],
        }
    if k == "sheet":
        from openpyxl import load_workbook

        wb = load_workbook(io.BytesIO(data))
        return {"sheets": [{"name": ws.title, "dimensions": ws.dimensions, "rows": ws.max_row,
                            "columns": ws.max_column} for ws in wb.worksheets],
                "lost_on_edit": lossy_parts(data)}
    if k == "slides":
        from pptx import Presentation

        prs = Presentation(io.BytesIO(data))
        return {"slides": [{"slide": i + 1, "title": s.shapes.title.text if s.shapes.title is not None else None}
                           for i, s in enumerate(prs.slides)]}
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(data))
    fields = reader.get_fields() or {}
    return {"pages": len(reader.pages),
            "form_fields": {name: {"type": f.get("/FT"), "value": f.get("/V")} for name, f in fields.items()}}


# --- Dateien aus dem Chat ---------------------------------------------------------------------------------------
# LibreChat legt Anhaenge und Ergebnisse der Codeausfuehrung als uploads/<Benutzer-ID>/<file_id>__<Name> ab; das
# Verzeichnis ist hier nur lesbar eingehaengt (CHAT_FILES_DIR). Bilder aus der Codeausfuehrung wandelt LibreChat um
# und legt sie ohne Originalnamen als images/<Benutzer-ID>/<file_id>.<Format> ab (CHAT_IMAGES_DIR, ebenfalls nur
# lesbar). Bilder aus Werkzeugergebnissen liegen dort als <file_id>__<Name> (BOS-Patch tool-files von LibreChat);
# die Bilder, die LibreChat selbst aus Werkzeugergebnissen ablegt (<uuid>-<Werkzeug>_img_...), bleiben aussen vor. Die Benutzer-ID
# kommt als Header von LibreChat. Nur mit dem gemeinsamen Secret gilt sie: andere Container im Docker-Netz koennten
# sonst fremde IDs angeben.

CHAT_FILES_DIR = os.environ.get("CHAT_FILES_DIR", "")
CHAT_IMAGES_DIR = os.environ.get("CHAT_IMAGES_DIR", "")
CHAT_FILES_SECRET = os.environ.get("CHAT_FILES_SECRET", "")
CODE_IMAGE = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.(png|jpe?g|webp|gif)")


def chat_user(ctx: Context, client) -> str:
    if not CHAT_FILES_DIR or not CHAT_FILES_SECRET:
        raise ToolError("Chat files are not available on this server")
    request = ctx.request_context.request
    headers = request.headers if request is not None else {}
    if not secrets.compare_digest(headers.get("x-chat-files-secret", ""), CHAT_FILES_SECRET):
        raise ToolError("Chat files are only available when called from LibreChat")
    user_id = headers.get("x-librechat-user-id", "")
    if not re.fullmatch(r"[0-9a-f]{24}", user_id):
        raise ToolError("LibreChat did not send a valid user id")
    if client.username not in (headers.get("x-librechat-username"), headers.get("x-librechat-openid-id")):
        raise ToolError("The LibreChat user does not match the Nextcloud login")
    return user_id


def chat_files(user_id: str) -> list[dict[str, Any]]:
    """Dateien des Benutzers, neueste zuerst; Bilder aus der Codeausfuehrung mit "code_image": True."""
    import datetime

    def entry(p: pathlib.Path, name: str, file_id: str, **extra) -> dict[str, Any]:
        st = p.stat()
        return {"name": name, "file_id": file_id, **extra, "size": st.st_size, "mtime": st.st_mtime,
                "modified": datetime.datetime.fromtimestamp(st.st_mtime, datetime.timezone.utc)
                .isoformat(timespec="minutes"),
                "_path": p}

    found = []
    directory = pathlib.Path(CHAT_FILES_DIR) / user_id
    if directory.is_dir():
        for p in directory.iterdir():
            file_id, sep, name = p.name.partition("__")
            if sep and name and p.is_file() and not p.is_symlink():
                found.append(entry(p, name, file_id))
    images = pathlib.Path(CHAT_IMAGES_DIR) / user_id if CHAT_IMAGES_DIR else None
    if images and images.is_dir():
        for p in images.iterdir():
            if not p.is_file() or p.is_symlink():
                continue
            if CODE_IMAGE.fullmatch(p.name):
                found.append(entry(p, f"image-{p.stem[:8]}{p.suffix}", p.stem, code_image=True))
                continue
            # Bilder aus Werkzeugergebnissen speichert LibreChat (BOS-Patch tool-files) mit Namen
            file_id, sep, name = p.name.partition("__")
            if sep and name:
                found.append(entry(p, name, file_id))
    return sorted(found, key=lambda f: f["mtime"], reverse=True)


# --- Aktivitaeten ---------------------------------------------------------------------------------------------
# Die App activity begrenzt eine Seite auf 200 Eintraege (Data::get). 25 Seiten sind 5000 Eintraege; mehr hat ein
# Zeitraum, den ein Modell sinnvoll liest, nicht.
ACTIVITY_PAGE = 200
ACTIVITY_MAX_PAGES = 25
ACTIVITY_TEXT = 300


def _activity_bound(value: str, zone: zoneinfo.ZoneInfo, end: bool) -> datetime.datetime:
    """Tag (YYYY-MM-DD) oder Zeitpunkt (ISO) als Grenze; ein Tag zaehlt als Ganzes."""
    value = value.strip()
    try:
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
            day = datetime.date.fromisoformat(value)
            moment = datetime.datetime.combine(day, datetime.time.max if end else datetime.time.min, zone)
            return moment.replace(microsecond=0)
        moment = datetime.datetime.fromisoformat(value)
    except ValueError as e:
        raise ToolError(f"{value!r} is neither YYYY-MM-DD nor an ISO date-time") from e
    return moment if moment.tzinfo else moment.replace(tzinfo=zone)


async def _own_user_id(client) -> str:
    """Nextcloud-Benutzer-ID des Kontos; der Filter actor vergleicht mit ihr, nicht mit dem Anmeldenamen.

    Mit dem Anmeldenamen ("jdoe" statt der ID "jdoe@example.com") liefert der Aktivitaetsstrom 304, also nichts.
    """
    try:
        response = await client.webdav._make_request(
            "GET", "/ocs/v2.php/cloud/user", params={"format": "json"}, headers={"Accept": "application/json"})
        user_id = ((response.json().get("ocs") or {}).get("data") or {}).get("id")
    except (httpx.HTTPError, ValueError, AttributeError):
        user_id = None
    return user_id if isinstance(user_id, str) and user_id else client.username


def _activity_entry(item: dict[str, Any], excluded) -> dict[str, Any] | None:
    """Ein Eintrag der Aktivitaeten-API ohne Vorschauen, Symbole und Rich-Text; None fuer gesperrte Pfade."""
    paths = [p.strip("/") for p in (item.get("objects") or {}).values() if isinstance(p, str) and p]
    if not paths and item.get("object_type") == "files" and item.get("object_name"):
        paths = [str(item["object_name"]).strip("/")]
    if excluded and any(is_path_excluded(p, excluded) for p in paths):
        return None
    entry: dict[str, Any] = {"time": item.get("datetime"), "user": item.get("user"), "type": item.get("type"),
                             "app": item.get("app"), "subject": (item.get("subject") or "")[:ACTIVITY_TEXT]}
    if paths:
        entry["paths"] = paths
    if item.get("message"):
        entry["message"] = item["message"][:ACTIVITY_TEXT]
    if item.get("link"):
        entry["link"] = item["link"]
    return entry


# --- Werkzeuge --------------------------------------------------------------------------------------------------

# Sortierschluessel von nc_webdav_find_by_type und die WebDAV-Eigenschaft, nach der Nextcloud sortiert.
# Groesse ueber oc:size: d:getcontentlength ist in der Suche nur abrufbar, nicht sortierbar (HTTP 400
# "requested order by property is not a valid property for this scope").
_ORDER_PROPS = {"size": "{http://owncloud.org/ns}size", "modified": "getlastmodified", "name": "displayname"}

# Nextcloud nimmt fuer den MIME-Typ nur "typ/untertyp" oder "typ/%" an und lehnt alles andere mit HTTP 500 ab,
# das wie ein voruebergehender Fehler aussieht. "%" allein und aehnliche Schreibweisen meinen alle Dateien:
# Dafuer sucht das Werkzeug alles ausser Ordnern, statt dass das Modell jeden Typ einzeln abfragt.
_ALL_TYPES = {"", "%", "*", "*/*", "*/%", "%/%"}
_MIME_FORMAT = re.compile(r"[\w.+-]+/(%|[\w.+-]+)")
_NO_FOLDERS = ("<d:not><d:eq><d:prop><d:getcontenttype/></d:prop>"
               "<d:literal>httpd/unix-directory</d:literal></d:eq></d:not>")


def configure_office_tools(mcp: MCPServer) -> None:
    # Ersetzt das gleichnamige Werkzeug aus nextcloud_mcp_server.server.webdav (die App "office" wird zuletzt
    # eingerichtet). Ohne Sortierung war "die groessten/neuesten N Dateien eines Typs" nur ueber alle Treffer
    # zu beantworten, die das Modell Ordner fuer Ordner einsammelte, bis das Rekursionslimit griff.
    mcp.remove_tool("nc_webdav_find_by_type")

    @mcp.tool(title="Find Files by Type", annotations=ToolAnnotations(read_only_hint=True, open_world_hint=True))
    @require_scopes("files.read")
    @with_links
    @instrument_tool
    async def nc_webdav_find_by_type(mime_type: str, ctx: Context, scope: str = "", limit: int | None = None,
                                     order_by: Literal["size", "modified", "name"] | None = None,
                                     descending: bool = True) -> SearchFilesResponse:
        """Find files by MIME type in NextCloud, searching all subfolders of scope.

        Args:
            mime_type: "type/subtype" (application/pdf), "type/%" (image/%), "httpd/unix-directory" for
                folders, or "%" for files of every type (without folders)
            scope: Directory path to search in (empty string for user root)
            limit: Maximum number of results to return
            order_by: Sort on the server before limit applies: "size", "modified" or "name".
                For "the N largest / newest files" call once with order_by and limit=N, with "%" for
                every type; do not call once per type and do not collect files folder by folder.
            descending: Direction for order_by; true (default) puts the largest / newest first

        Returns:
            SearchFilesResponse with list of matching files
        """
        client = await get_client(ctx)
        excluded = await get_excluded_file_paths(client.webdav)
        if scope and is_path_excluded(scope, excluded):
            raise ToolError(f"Access denied: scope {scope!r} is tagged with an excluded tag")
        mime = mime_type.strip()
        if mime not in _ALL_TYPES and not _MIME_FORMAT.fullmatch(mime):
            raise ToolError(f'mime_type {mime_type!r} is not supported: use "type/subtype" (application/pdf), '
                            '"type/%" (image/%) or "%" for files of every type')
        where_conditions, properties = client.webdav._type_search_args(mime)
        if mime in _ALL_TYPES:
            where_conditions = _NO_FOLDERS
        results = await client.webdav.search_files(
            scope=scope, where_conditions=where_conditions, properties=properties,
            order_by=[(_ORDER_PROPS[order_by], "descending" if descending else "ascending")] if order_by else None,
            limit=limit)
        if excluded:
            results = [r for r in results if not is_path_excluded(r.get("path", ""), excluded)]
        filters: dict[str, Any] = {"mime_type": mime_type}
        if order_by:
            filters["order_by"] = f"{order_by} {'descending' if descending else 'ascending'}"
        file_infos = [FileInfo(**r) for r in results]
        return SearchFilesResponse(results=file_infos, total_found=len(file_infos), scope=scope,
                                   filters_applied=filters)

    @mcp.tool(title="Create Office Document", annotations=ToolAnnotations(idempotent_hint=False, open_world_hint=True))
    @require_scopes("files.write")
    @instrument_tool
    async def nc_office_create_document(path: str, markdown: str, ctx: Context, if_match: str | None = None):
        """Create a Word/Writer document or PDF in Nextcloud from Markdown.

        The format follows the extension of ``path``: .docx, .odt, .doc, .rtf or .pdf.
        Markdown supports headings (#), paragraphs, bullet and numbered lists (indent 2 spaces per level),
        tables (| a | b |), > quotes, ``` code blocks, **bold**, *italic*, `code` and [links](url).
        Plain structure only: text, tables and lists, no images or charts.
        For a designed file (charts, images, colors, layout), or to keep what the chat shows as an artifact,
        build the file with code execution (python-docx) and copy it with nc_office_save_chat_file;
        an artifact is a web page, not a file.

        Not for files attached in the chat: copy those with nc_office_save_chat_file.
        Fails if the file exists, unless ``if_match`` is its etag (or "*") to overwrite it.
        Returns path, size and the new etag.
        """
        if kind(path) not in ("word", "pdf"):
            raise ToolError("Use .docx, .odt, .doc, .rtf or .pdf")
        client = await get_client(ctx)
        data = await anyio.to_thread.run_sync(create_docx, markdown)
        return await save(client, path, await from_native(data, path, "docx"), if_match)

    @mcp.tool(title="Create Spreadsheet", annotations=ToolAnnotations(idempotent_hint=False, open_world_hint=True))
    @require_scopes("files.write")
    @instrument_tool
    async def nc_office_create_spreadsheet(path: str, sheets: list[dict[str, Any]], ctx: Context,
                                           if_match: str | None = None):
        """Create a spreadsheet in Nextcloud: .xlsx, .ods, .xls or .csv (first sheet only).

        ``sheets``: [{"name": "Checkliste", "rows": [["Punkt", "Status"], ["BitLocker", "erledigt"], ...],
        "header": true}]. Values may be strings, numbers, booleans or null; a string starting with "=" is a
        formula (e.g. "=SUM(B2:B10)"). With "header" (default true) the first row is bold and frozen.
        Plain structure only: values and formulas, no charts or cell formatting beyond the bold header.
        For a designed file (charts, images, colors, layout), or to keep what the chat shows as an artifact,
        build the file with code execution (openpyxl) and copy it with nc_office_save_chat_file;
        an artifact is a web page, not a file.
        Not for files attached in the chat: copy those with nc_office_save_chat_file.

        Fails if the file exists, unless ``if_match`` is its etag (or "*") to overwrite it.
        """
        if kind(path) != "sheet":
            raise ToolError("Use .xlsx, .ods, .xls or .csv")
        client = await get_client(ctx)
        data = await anyio.to_thread.run_sync(create_xlsx, sheets)
        if any(isinstance(v, str) and v.startswith("=") for s in sheets for row in s.get("rows", []) for v in row):
            data = await recalculate(data)
        return await save(client, path, await from_native(data, path, "xlsx"), if_match)

    @mcp.tool(title="Create Presentation", annotations=ToolAnnotations(idempotent_hint=False, open_world_hint=True))
    @require_scopes("files.write")
    @instrument_tool
    async def nc_office_create_presentation(path: str, slides: list[dict[str, Any]], ctx: Context,
                                            if_match: str | None = None):
        """Create a presentation in Nextcloud: .pptx, .odp, .ppt (16:9).

        ``slides``: [{"title": "Prototypenschutz", "bullets": ["Zutritt", "  Unterpunkt"], "notes": "..."}].
        Two leading spaces per level indent a bullet. A slide without bullets gets a title-only layout.
        Plain structure only: title, bullets and notes on a white slide, no charts, images or colors.
        For a designed file (charts, images, colors, layout), or to keep what the chat shows as an artifact,
        build the file with code execution (python-pptx) and copy it with nc_office_save_chat_file;
        an artifact is a web page, not a file.
        Not for files attached in the chat: copy those with nc_office_save_chat_file.

        Fails if the file exists, unless ``if_match`` is its etag (or "*") to overwrite it.
        """
        if kind(path) != "slides":
            raise ToolError("Use .pptx, .odp or .ppt")
        client = await get_client(ctx)
        data = await anyio.to_thread.run_sync(create_pptx, slides)
        return await save(client, path, await from_native(data, path, "pptx"), if_match)

    async def _edit(ctx, path, output_path, editor, expected: set[str]):
        k = kind(path)
        if k not in expected:
            raise ToolError(f"{path!r} is not a supported file for this tool")
        target = output_path or path
        if kind(target) not in (k, "pdf"):
            raise ToolError(f"output_path must be a {k} file or a PDF")
        client = await get_client(ctx)
        original, etag = await load(client, path)
        data, done, warnings = await editor(client, original if k == "pdf" else await to_native(original, path))
        if k != "pdf":
            data = await from_native(data, target, NATIVE[k])
        result = await save(client, target, data, None if output_path else (etag or "*"))
        result["applied"] = done
        if warnings:
            result["warnings"] = warnings
        return result

    @mcp.tool(title="Edit Document", annotations=ToolAnnotations(idempotent_hint=False, open_world_hint=True))
    @require_scopes("files.write")
    @instrument_tool
    async def nc_office_edit_document(path: str, operations: list[dict[str, Any]], ctx: Context,
                                      output_path: str | None = None):
        """Edit a Word/Writer document in Nextcloud (.docx, .odt, .doc, .rtf), keeping its formatting.

        Read it first (nc_webdav_read_file, or nc_office_inspect for table numbers). ``operations`` run in
        order:
        - {"op": "replace_text", "find": "alt", "replace": "neu"}  (all occurrences, also tables/headers)
        - {"op": "append_markdown", "markdown": "## Neuer Abschnitt\\n..."}
        - {"op": "insert_after" | "insert_before", "anchor": "text in a paragraph", "markdown": "..."}
        - {"op": "delete_paragraphs", "contains": "text"}
        - {"op": "set_cell", "table": 1, "row": 2, "column": 3, "text": "erledigt"}  (1-based)
        - {"op": "add_row", "table": 1, "values": ["a", "b"]}  (copies the last row's formatting)
        Saves in place (Nextcloud keeps the previous version) unless ``output_path`` names a new file.
        Fails if someone changed the file meanwhile. .odt/.doc/.rtf are edited via .docx and converted back.
        """
        async def editor(_client, data):
            out, done = await anyio.to_thread.run_sync(edit_docx, data, operations)
            return out, done, []
        return await _edit(ctx, path, output_path, editor, {"word"})

    @mcp.tool(title="Edit Spreadsheet", annotations=ToolAnnotations(idempotent_hint=False, open_world_hint=True))
    @require_scopes("files.write")
    @instrument_tool
    async def nc_office_edit_spreadsheet(path: str, operations: list[dict[str, Any]], ctx: Context,
                                         output_path: str | None = None, allow_loss: bool = False):
        """Edit a spreadsheet in Nextcloud (.xlsx, .ods, .xls, .csv), keeping cell formatting.

        ``operations`` run in order; "sheet" defaults to the active sheet:
        - {"op": "set_cells", "sheet": "Checkliste", "cells": {"E2": "erledigt", "F2": "=SUM(B2:D2)"}}
        - {"op": "append_rows", "sheet": "...", "rows": [["a", 1], ["b", 2]]}
        - {"op": "insert_rows" | "delete_rows", "sheet": "...", "row": 5, "count": 2}  (1-based)
        - {"op": "add_sheet", "name": "Neu", "rows": [...]}, {"op": "rename_sheet", "sheet": "...",
          "new_name": "..."}, {"op": "delete_sheet", "sheet": "..."}
        If all operations are set_cells or append_rows, only the touched cells change and everything else
        (formatting, images, charts, pivot tables, macros) stays exactly as it was; append_rows continues
        below the last row with content, copies that row's cell styles and extends a table ending there.
        The other operations rewrite the workbook: charts, images, pivot tables and macros would be lost, so
        the tool then refuses unless ``allow_loss`` is true (ask the user first, and prefer ``output_path``).
        Saves in place (Nextcloud keeps the previous version) unless ``output_path`` names a new file.
        """
        lossless = all(op.get("op") in LOSSLESS_OPS for op in operations)

        async def editor(_client, data):
            warnings = []
            if lossless:
                out, done = await anyio.to_thread.run_sync(edit_xlsx_lossless, data, operations)
                try:
                    stripped = await anyio.to_thread.run_sync(strip_formula_results, out)
                    if stripped is not None:
                        out, n = await anyio.to_thread.run_sync(apply_formula_results, out,
                                                                await recalculate(stripped))
                        done.append(f"formula results recalculated: {n}")
                except ToolError as e:
                    warnings.append(f"Formula results were not recalculated ({e}); they update when the file is "
                                    "opened in Excel or LibreOffice.")
                return out, done, warnings
            out, done = await anyio.to_thread.run_sync(edit_xlsx, data, operations, allow_loss)
            if any(isinstance(v, str) and v.startswith("=") for op in operations
                   for v in list((op.get("cells") or {}).values()) + [x for r in op.get("rows", []) for x in r]):
                warnings.append("New formulas have no computed value until the file is opened in Excel or "
                                "LibreOffice; nc_webdav_read_file shows them as empty (NaN) until then.")
            return out, done, warnings
        return await _edit(ctx, path, output_path, editor, {"sheet"})

    @mcp.tool(title="Edit Presentation", annotations=ToolAnnotations(idempotent_hint=False, open_world_hint=True))
    @require_scopes("files.write")
    @instrument_tool
    async def nc_office_edit_presentation(path: str, operations: list[dict[str, Any]], ctx: Context,
                                          output_path: str | None = None):
        """Edit a presentation in Nextcloud (.pptx, .odp, .ppt), keeping its design.

        ``operations`` run in order; slide numbers are 1-based:
        - {"op": "replace_text", "find": "alt", "replace": "neu"}  (all slides, also tables)
        - {"op": "add_slide", "title": "...", "bullets": ["..."], "notes": "...", "position": 3}
        - {"op": "delete_slide", "slide": 4}, {"op": "move_slide", "slide": 4, "to": 2}
        - {"op": "set_title", "slide": 2, "title": "..."}, {"op": "set_notes", "slide": 2, "notes": "..."}
        Saves in place (Nextcloud keeps the previous version) unless ``output_path`` names a new file.
        """
        async def editor(_client, data):
            out, done = await anyio.to_thread.run_sync(edit_pptx, data, operations)
            return out, done, []
        return await _edit(ctx, path, output_path, editor, {"slides"})

    @mcp.tool(title="Edit PDF", annotations=ToolAnnotations(idempotent_hint=False, open_world_hint=True))
    @require_scopes("files.write")
    @instrument_tool
    async def nc_office_edit_pdf(path: str, operations: list[dict[str, Any]], ctx: Context,
                                 output_path: str | None = None):
        """Edit a PDF in Nextcloud: pages and form fields. Its text cannot be rewritten; to change content,
        edit the source document and create the PDF again (nc_office_convert).

        ``operations`` run in order; pages are 1-based, ranges as "3-5":
        - {"op": "delete_pages", "pages": [2, "5-7"]}, {"op": "keep_pages", "pages": ["1-3"]}
        - {"op": "reorder_pages", "pages": [3, 1, 2]}, {"op": "rotate_pages", "pages": [1], "degrees": 90}
        - {"op": "append_pdf", "path": "Ordner/Anhang.pdf"}  (another PDF in Nextcloud)
        - {"op": "fill_form", "fields": {"Name": "Max"}}  (field names via nc_office_inspect)
        Saves in place (Nextcloud keeps the previous version) unless ``output_path`` names a new file.
        """
        async def editor(client, data):
            extra = {}
            for op in operations:
                if op.get("op") == "append_pdf":
                    if kind(op["path"]) != "pdf":
                        raise ToolError(f"{op['path']!r} is not a PDF")
                    extra[op["path"]], _ = await load(client, op["path"])
            out, done = await anyio.to_thread.run_sync(edit_pdf, data, operations, extra)
            return out, done, []
        return await _edit(ctx, path, output_path, editor, {"pdf"})

    @mcp.tool(title="Convert Office File", annotations=ToolAnnotations(idempotent_hint=False, open_world_hint=True))
    @require_scopes("files.write")
    @instrument_tool
    async def nc_office_convert(path: str, target_path: str, ctx: Context, if_match: str | None = None):
        """Convert a file in Nextcloud to another format with LibreOffice and save it as ``target_path``.

        Formats by extension: documents .docx .odt .doc .rtf, spreadsheets .xlsx .ods .xls .csv, presentations
        .pptx .odp .ppt; any of them to .pdf. Fails if the target exists, unless ``if_match`` is its etag
        (or "*").
        """
        src, dst = kind(path), kind(target_path)
        if src == "pdf" or (dst != "pdf" and dst != src):
            raise ToolError(f"Cannot convert {ext(path)} to {ext(target_path)}")
        client = await get_client(ctx)
        data, _ = await load(client, path)
        return await save(client, target_path, await convert(data, ext(path), ext(target_path)), if_match)

    @mcp.tool(title="Inspect Office File", annotations=ToolAnnotations(read_only_hint=True, open_world_hint=True))
    @require_scopes("files.read")
    @instrument_tool
    async def nc_office_inspect(path: str, ctx: Context):
        """Show the structure of an office file or PDF in Nextcloud, as a basis for the edit tools:
        documents: headings and numbered tables with their first row; spreadsheets: sheets, used ranges and
        parts that editing would lose; presentations: numbered slides with titles; PDF: page count and form
        fields. For the content itself use nc_webdav_read_file.
        """
        k = kind(path)
        client = await get_client(ctx)
        data, etag = await load(client, path)
        if k != "pdf":
            data = await to_native(data, path)
        info = await anyio.to_thread.run_sync(inspect_native, data, k)
        return {"path": path, "etag": etag, **info}

    @mcp.tool(title="List Chat Files", annotations=ToolAnnotations(read_only_hint=True, open_world_hint=False))
    @require_scopes("files.read")
    @instrument_tool
    async def nc_office_list_chat_files(ctx: Context, name_contains: str | None = None, limit: int = 30):
        """List the user's files in the chat: attachments they uploaded and files created by code execution,
        newest first, with name, file_id, size and time. Images created by code execution lose their name in the
        chat; they appear as image-<id>.<ext> with "code_image": true, recognizable by their time.
        Use nc_office_save_chat_file to put one into Nextcloud as the original file.
        """
        client = await get_client(ctx)
        files = await anyio.to_thread.run_sync(chat_files, chat_user(ctx, client))
        if name_contains:
            files = [f for f in files if name_contains.casefold() in f["name"].casefold()]
        return {"files": [{k: v for k, v in f.items() if k not in ("_path", "mtime")}
                          for f in files[:max(1, min(limit, 200))]]}

    @mcp.tool(title="Save Chat File to Nextcloud",
              annotations=ToolAnnotations(idempotent_hint=False, open_world_hint=True))
    @require_scopes("files.write")
    @instrument_tool
    async def nc_office_save_chat_file(name: str, target_path: str, ctx: Context, file_id: str | None = None,
                                       if_match: str | None = None):
        """Save a file from the chat into Nextcloud as the original, unchanged file: an attachment the user
        uploaded (Word, Excel, PowerPoint, PDF, images, ...) or a file created by code execution.

        Always use this to put a chat file into Nextcloud, never re-create it from its text with
        nc_office_create_* (that loses all formatting). To change an attached file, save it first, then use
        nc_office_edit_* on the saved copy.

        ``name`` is the file name as shown in the chat; with several files of that name the newest is taken
        unless ``file_id`` (from nc_office_list_chat_files) picks one. An image created by code execution loses
        its name in the chat: for an image name not found, the newest such image is taken (preferring the
        extension of ``target_path``); for an older one pass its file_id. LibreChat keeps these images reduced to
        at most 768 px on the short side. A different office extension in ``target_path`` converts it (e.g. .xlsx to .ods or .pdf).
        Fails if the target exists, unless ``if_match`` is its etag (or "*").
        Never write a chat file to Nextcloud as base64 text through another tool; this tool copies it directly.
        """
        client = await get_client(ctx)
        files = await anyio.to_thread.run_sync(chat_files, chat_user(ctx, client))
        match = [f for f in files if (file_id and f["file_id"] == file_id) or (not file_id and f["name"] == name)]
        if not match and not file_id:
            match = [f for f in files if f["name"].casefold() == name.casefold()]
        if not match and not file_id and ext(name) in ("png", "jpg", "jpeg", "webp", "gif"):
            images = [f for f in files if f.get("code_image")]
            match = ([f for f in images if ext(f["name"]) == ext(target_path)] or images)[:1]
        if not match:
            raise ToolError(f"No chat file {name!r}; recent files: {[f['name'] for f in files[:15]]}")
        source = match[0]
        data = await anyio.Path(source["_path"]).read_bytes()
        src, dst = ext(source["name"]), ext(target_path)
        if src != dst:
            try:
                k_src, k_dst = kind(source["name"]), kind(target_path)
            except ToolError:
                raise ToolError(f"Cannot convert {src!r} to {dst!r}; keep the extension .{src}") from None
            if k_src == "pdf" or (k_dst != "pdf" and k_dst != k_src):
                raise ToolError(f"Cannot convert {src} to {dst}")
            data = await convert(data, src, dst)
        result = await save(client, target_path, data, if_match)
        result["source"] = {"name": source["name"], "file_id": source["file_id"], "modified": source["modified"],
                            **({"code_image": True} if source.get("code_image") else {})}
        return result

    @mcp.tool(title="Copy Nextcloud File to Chat",
              annotations=ToolAnnotations(read_only_hint=True, open_world_hint=True))
    @require_scopes("files.read")
    @instrument_tool
    async def nc_office_copy_file_to_chat(path: str, ctx: Context):
        """Copy a Nextcloud file unchanged into the chat, of any type: also plain text, scripts, JSON and CSV.
        The result is a [FILE] line: code execution gets the file under /mnt/data, other tools get it through
        {{file:...}}, and nc_office_save_chat_file copies it back.

        Use this whenever code has to process, compare, hash or change a file. Never paste the content of a
        file read with nc_webdav_read_file into code: retyped content is slow and can differ from the original.
        To read a file yourself, use nc_webdav_read_file. At most 50 MiB.
        """
        client = await get_client(ctx)
        data, etag = await load(client, path)
        if len(data) > CHAT_FILE_MAX_BYTES:
            raise ToolError(f"{path!r} has {len(data) / 1048576:.1f} MB; at most 50 MiB can be copied into the chat")
        name = pathlib.PurePosixPath(path).name
        content_type = MIME.get(ext(path)) or mimetypes.guess_type(name)[0] or "application/octet-stream"
        return {"path": path, "filename": name, "content_type": content_type, "size": len(data), "etag": etag,
                "content_base64": base64.b64encode(data).decode("ascii"),
                # LibreChat (tool-files) then returns only the [FILE] line, also for text
                "as_file": True}


    @mcp.tool(title="Check List Against Files", annotations=ToolAnnotations(read_only_hint=True, open_world_hint=True))
    @require_scopes("files.read")
    @instrument_tool
    async def nc_files_check_list(values: str, ctx: Context, folder: str = "", match: str = "name",
                                  only_missing: bool = False):
        """Check a whole list of names or paths against the files of a folder and its subfolders in one call.

        ``values``: file names, name parts or paths, separated by commas or new lines (a JSON array is accepted
        too), at most 1000. ``match``: "name" (default) finds each value as a whole word in a file or folder
        name ("Plan 12" does not match "Plan 123"), "name_exact" compares the whole name, "path" compares the
        path relative to ``folder`` or to the root. A value with a "/" is always compared as a path.
        Returns which values have matches (path, size, modified; at most 20 per value) and which are missing.
        Use this instead of one search or directory listing per value.
        """
        terms = parse_terms(values)
        if not terms:
            raise ToolError("values is empty")
        if len(terms) > CHECK_LIST_MAX_TERMS:
            raise ToolError(f"At most {CHECK_LIST_MAX_TERMS} values per call, got {len(terms)}")
        if match not in ("name", "name_exact", "path"):
            raise ToolError('match must be "name", "name_exact" or "path"')
        client = await get_client(ctx)
        base = folder.strip("/")
        excluded = await get_excluded_file_paths(client.webdav)

        def wanted_of(term: str) -> str:
            return term.strip("/").lower()

        def name_of(term: str) -> str:
            return wanted_of(term).rsplit("/", 1)[-1]

        def by_path(term: str) -> bool:
            return match == "path" or "/" in wanted_of(term)

        # Nextcloud sucht die Namen in der Datenbank (iLike auf dem Namen), statt alle Dateien zu liefern: Die Wurzel
        # hat mehr als 50000 Dateien, ein Vollscan dauert laenger als das Zeitlimit von LibreChat (60 s).
        # Gefunden wird grob, der genaue Vergleich folgt unten.
        def condition(term: str) -> str:
            value = name_of(term).replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
            if not (by_path(term) or match == "name_exact"):
                value = f"%{value}%"
            return f"<d:like><d:prop><d:displayname/></d:prop><d:literal>{xml_escape(value)}</d:literal></d:like>"

        conditions = sorted({condition(term) for term in terms if name_of(term)})
        items: dict[Any, dict[str, Any]] = {}
        truncated = False
        for i in range(0, len(conditions), CHECK_LIST_TERMS_PER_SEARCH):
            chunk = conditions[i:i + CHECK_LIST_TERMS_PER_SEARCH]
            where = chunk[0] if len(chunk) == 1 else f"<d:or>{''.join(chunk)}</d:or>"
            rows = await client.webdav.search_files_all(scope=base, where_conditions=where)
            truncated = truncated or len(rows) >= WEBDAV_SEARCH_MAX_RESULTS
            for item in rows:
                if item.get("path") and not is_path_excluded(item["path"], excluded):
                    items[item.get("file_id") or item["path"]] = item
        found: dict[str, list[dict[str, Any]]] = {}
        for term in terms:
            wanted = wanted_of(term)
            for item in items.values():
                path = item["path"].strip("/")
                name = path.rsplit("/", 1)[-1]
                relative = path[len(base) + 1:] if base and path.lower().startswith(base.lower() + "/") else path
                if by_path(term):
                    hit = wanted in (path.lower(), relative.lower())
                elif match == "name_exact":
                    hit = name.lower() == wanted
                else:
                    hit = word_match(term, name)
                if hit and len(found.setdefault(term, [])) < CHECK_LIST_MAX_HITS:
                    found[term].append({"path": path, "size": item.get("size"),
                                        "modified": item.get("last_modified")})
        missing = [term for term in terms if term not in found]
        result: dict[str, Any] = {"values_checked": len(terms), "found_count": len(found),
                                  "missing_count": len(missing), "missing": missing}
        if truncated:
            result["truncated"] = (f"a search returned {WEBDAV_SEARCH_MAX_RESULTS} files, the limit: missing values "
                                   "may exist; check fewer or more specific values")
        if not only_missing:
            result["found"] = found
        return result

    @mcp.tool(title="Full Text Search", annotations=ToolAnnotations(read_only_hint=True, open_world_hint=True))
    @require_scopes("files.read")
    @instrument_tool
    async def nc_fulltext_search(query: str, ctx: Context, extension: str = "",
                                 match: Literal["all", "content", "filename"] = "all",
                                 limit: int = 20, page: int = 1):
        """Search the content of files with the full text index of Nextcloud, including OCR text of PDFs.

        Use this to find documents by what they say; nc_webdav_find_by_name only sees names. Finds only files the
        user may access, and files that are indexed: files with a password are not.

        ``query``: words, not case-sensitive. Several words find files with ANY of them, ranked by relevance;
        +word requires a word, -word excludes it, "two words" is a phrase, rech* a prefix. Use +word when all
        words must occur. ``extension``: only files with this extension, e.g. "pdf". ``match``: "content" or
        "filename" limit where to look. ``limit``: results per page, at most 50. ``page``: next results.
        Returns ``total`` (all hits, counted up to 10000), ``has_more`` and for each result its path, usable with
        nc_webdav_read_file, with mime type, size, modified time and short text excerpts around the hits. Excerpts
        are not the whole document: read the file before relying on its content.
        """
        if not query.strip():
            raise ToolError("query is empty")
        extension = extension.strip().lstrip(".")
        if extension and not re.fullmatch(r"[A-Za-z0-9]+", extension):
            raise ToolError("extension may only contain letters and digits")
        if not 1 <= limit <= FULLTEXT_MAX_LIMIT:
            raise ToolError(f"limit must be between 1 and {FULLTEXT_MAX_LIMIT}")
        if page < 1:
            raise ToolError("page must be at least 1")
        options: dict[str, Any] = {}
        if extension:
            # Die App setzt den Wert in einen regulaeren Ausdruck ein, der Gross- und Kleinschreibung unterscheidet
            options["files_extension"] = "".join(f"[{c.lower()}{c.upper()}]" if c.isalpha() else c for c in extension)
        if match != "all":
            options["in"] = [match]
        request = {"providers": "files", "search": query, "page": page, "size": limit, "options": options,
                   "tags": [], "subtags": [], "parts": []}
        client = await get_client(ctx)
        try:
            response = await client.webdav._make_request("GET", "/apps/fulltextsearch/v1/remote",
                                                         params={"request": json.dumps(request)},
                                                         headers={"Accept": "application/json"})
        except httpx.HTTPStatusError as e:
            raise ToolError(f"Full text search failed (HTTP {e.response.status_code})") from e
        data = response.json()
        if data.get("status") != 1 or not data.get("result"):
            raise ToolError(f"Full text search failed: {data.get('message') or data.get('status')}")
        found = data["result"][0]
        total = found["meta"]["total"]
        excluded = await get_excluded_file_paths(client.webdav)
        results = []
        for document in found["documents"]:
            info = document.get("info") or {}
            path = (info.get("path") or "").strip("/")
            if not path or is_path_excluded(path, excluded):
                continue
            mtime = info.get("mtime")
            excerpts = [e["excerpt"].strip()[:FULLTEXT_EXCERPT_CHARS] for e in document.get("excerpts") or []
                        if e.get("excerpt")]
            results.append({"path": path, "type": info.get("type"), "mime_type": info.get("mime"),
                            "size": info.get("size"),
                            "modified": datetime.datetime.fromtimestamp(mtime, datetime.UTC).isoformat() if mtime else None,
                            "excerpts": excerpts[:FULLTEXT_EXCERPTS]})
        return {"query": query, "total": total, "page": page, "limit": limit, "has_more": page * limit < total,
                "results": results}

    @mcp.tool(title="List Activities", annotations=ToolAnnotations(read_only_hint=True, open_world_hint=True))
    @require_scopes("files.read")
    @instrument_tool
    async def nc_activity_list(date_from: str, ctx: Context, date_to: str = "",
                               actor: Literal["me", "others", "all"] = "me", actor_name: str = "",
                               path_contains: str = "", time_zone: str = "Europe/Berlin"):
        """List who did what in Nextcloud in a period, from the user's activity stream: files and folders
        created, changed, renamed, moved, deleted, restored, shared, tagged and commented, and events of other
        apps. Each entry names the account that did it (``user``), so this is the tool for "what did I do on
        <day>" and "who changed this file". A file's modification time does not say who changed it; this does.

        ``date_from``/``date_to``: days as YYYY-MM-DD (or ISO date-times), both inclusive, in ``time_zone``;
        ``date_to`` empty means the day of ``date_from``. ``actor``: "me" (only what the user did), "others"
        (what other accounts did to files the user can see) or "all". ``actor_name``: only this account
        (e.g. "jdoe@example.com"); it overrides ``actor``. ``path_contains``: only entries whose file path
        contains this text. Reads every page of the period, oldest first; ``complete`` is false when the period
        holds more entries than the tool returns.
        """
        try:
            zone = zoneinfo.ZoneInfo(time_zone)
        except (zoneinfo.ZoneInfoNotFoundError, ValueError) as e:
            raise ToolError(f"unknown time_zone {time_zone!r}") from e
        start = _activity_bound(date_from, zone, end=False)
        end = _activity_bound(date_to or date_from, zone, end=True)
        if end < start:
            raise ToolError("date_to lies before date_from")
        client = await get_client(ctx)
        name = actor_name.strip() or (await _own_user_id(client) if actor == "me" else "")
        stream = "by" if actor == "others" and not actor_name.strip() else "all"
        params: dict[str, Any] = {"format": "json", "sort": "asc", "limit": ACTIVITY_PAGE,
                                  "from": int(start.timestamp()), "to": int(end.timestamp())}
        if name:
            params["actor"] = name
        if path_contains.strip():
            params["search"] = path_contains.strip()
        excluded = await get_excluded_file_paths(client.webdav)
        entries: list[dict[str, Any]] = []
        complete = True
        since = 0
        for _ in range(ACTIVITY_MAX_PAGES):
            try:
                response = await client.webdav._make_request(
                    "GET", f"/ocs/v2.php/apps/activity/api/v2/activity/{stream}",
                    params={**params, "since": since}, headers={"Accept": "application/json"})
            except httpx.HTTPStatusError as e:
                if e.response.status_code == 304:  # nichts (mehr) im Zeitraum
                    break
                raise ToolError(f"activity request failed (HTTP {e.response.status_code}): "
                                f"{e.response.text[:300]}") from e
            if response.status_code == 204:
                raise ToolError("the activity stream is switched off for this user")
            if response.status_code == 304 or not response.content:
                break
            batch = (response.json().get("ocs") or {}).get("data") or []
            for item in batch:
                entry = _activity_entry(item, excluded)
                if entry:
                    entries.append(entry)
            last = response.headers.get("X-Activity-Last-Given")
            if not batch or not last or 'rel="next"' not in response.headers.get("Link", ""):
                break
            since = int(last)
        else:
            complete = False
        counts: dict[str, int] = {}
        for entry in entries:
            counts[entry["type"]] = counts.get(entry["type"], 0) + 1
        return {"from": start.isoformat(), "to": end.isoformat(), "actor": name or actor, "total": len(entries),
                "complete": complete, "by_type": counts, "activities": entries}
