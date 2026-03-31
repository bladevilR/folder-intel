#!/usr/bin/env python3
"""
Archive and search Office-style document folders.

Supported inputs:
- .docx
- .xlsx
- .xlsm
- .pptx
- .pdf
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import textwrap
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from typing import Any, Iterable, Optional, Sequence

try:
    from office_legacy_win32 import (
        extract_doc_via_word,
        extract_ppt_via_powerpoint,
        extract_xls_via_excel,
    )
except Exception:
    extract_doc_via_word = None
    extract_xls_via_excel = None
    extract_ppt_via_powerpoint = None

try:
    import olefile  # type: ignore
except Exception:
    olefile = None

try:
    import xlrd  # type: ignore
except Exception:
    xlrd = None

WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
DRAWING_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
PKG_REL_NS = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
CORE_NS = {
    "dc": "http://purl.org/dc/elements/1.1/",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
}

SUPPORTED_EXTENSIONS = {
    ".doc": "doc",
    ".docx": "docx",
    ".xls": "xls",
    ".xlsx": "xlsx",
    ".xlsm": "xlsm",
    ".ppt": "ppt",
    ".pptx": "pptx",
    ".pdf": "pdf",
}
DEFAULT_SKIP_DIRS = {
    ".git",
    ".hg",
    ".idea",
    ".office-archive",
    ".svn",
    ".venv",
    "__pycache__",
    "build",
    "dist",
    "node_modules",
    "venv",
}
LEGACY_EXTRACT_TIMEOUT_SECONDS = 90


@dataclass
class ArchiveEntry:
    path: str
    file_type: str
    size_bytes: int
    modified_at: str
    mtime_ns: int
    title: str
    summary: str
    preview: str
    content: str
    content_hash: str
    status: str
    error: Optional[str]
    details: dict[str, Any]
    indexed_at: str


def eprint(message: str) -> None:
    print(message, file=sys.stderr)


def positive_int(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("must be an integer") from exc
    if parsed < 1:
        raise argparse.ArgumentTypeError("must be >= 1")
    return parsed


def current_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def iso_from_timestamp(timestamp: float) -> str:
    return datetime.fromtimestamp(timestamp, tz=timezone.utc).replace(microsecond=0).isoformat().replace(
        "+00:00", "Z"
    )


def normalize_whitespace(text: str) -> str:
    text = text.replace("\x00", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines: list[str] = []
    for raw_line in text.splitlines():
        line = re.sub(r"[ \t]+", " ", raw_line).strip()
        if line:
            lines.append(line)
    return "\n".join(lines)


def clip_text(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    if limit <= 3:
        return text[:limit]
    return text[: limit - 3].rstrip() + "..."


def excel_column_label(index: int) -> str:
    result = ""
    current = index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        result = chr(65 + remainder) + result
    return result


def build_summary(file_type: str, text: str, title: str, max_chars: int = 280) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return f"{title}: no extractable text found."
    line_limit = 5 if file_type in {"xls", "xlsx", "xlsm"} else 3
    summary = " | ".join(lines[:line_limit])
    return clip_text(summary, max_chars)


def build_preview(text: str, max_lines: int = 8, max_chars: int = 1200) -> str:
    preview_lines: list[str] = []
    char_count = 0
    for line in text.splitlines():
        clean = line.strip()
        if not clean:
            continue
        preview_lines.append(clean)
        char_count += len(clean)
        if len(preview_lines) >= max_lines or char_count >= max_chars:
            break
    return clip_text("\n".join(preview_lines), max_chars)


def sha1_text(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


def read_xml_from_zip(archive: zipfile.ZipFile, member: str) -> Optional[ET.Element]:
    try:
        raw = archive.read(member)
    except KeyError:
        return None
    return ET.fromstring(raw)


def read_core_properties(archive: zipfile.ZipFile) -> dict[str, str]:
    root = read_xml_from_zip(archive, "docProps/core.xml")
    if root is None:
        return {}
    result: dict[str, str] = {}
    title = root.findtext("dc:title", default="", namespaces=CORE_NS).strip()
    subject = root.findtext("dc:subject", default="", namespaces=CORE_NS).strip()
    if title:
        result["title"] = title
    if subject:
        result["subject"] = subject
    return result


def resolve_zip_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    resolved = PurePosixPath(source_part).parent.joinpath(target)
    normalized = str(resolved).replace("\\", "/")
    normalized = os.path.normpath(normalized).replace("\\", "/")
    return normalized.lstrip("./")


def extract_docx(path: Path) -> tuple[str, dict[str, Any]]:
    with zipfile.ZipFile(path) as archive:
        root = read_xml_from_zip(archive, "word/document.xml")
        if root is None:
            raise RuntimeError("word/document.xml not found")
        paragraphs: list[str] = []
        for paragraph in root.findall(".//w:p", WORD_NS):
            parts = [node.text or "" for node in paragraph.findall(".//w:t", WORD_NS)]
            text = "".join(parts).strip()
            if text:
                paragraphs.append(text)
        props = read_core_properties(archive)
        details: dict[str, Any] = {"paragraph_count": len(paragraphs)}
        details.update(props)
        return "\n".join(paragraphs), details


def extract_pptx(path: Path) -> tuple[str, dict[str, Any]]:
    with zipfile.ZipFile(path) as archive:
        slide_members = sorted(
            [name for name in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)],
            key=lambda name: int(re.search(r"(\d+)", name).group(1)),
        )
        slides: list[str] = []
        for member in slide_members:
            root = read_xml_from_zip(archive, member)
            if root is None:
                continue
            parts = [node.text or "" for node in root.findall(".//a:t", DRAWING_NS)]
            text = " | ".join(part.strip() for part in parts if part and part.strip())
            if text:
                number_match = re.search(r"slide(\d+)\.xml$", member)
                slide_number = number_match.group(1) if number_match else "?"
                slides.append(f"Slide {slide_number}: {text}")
        props = read_core_properties(archive)
        details: dict[str, Any] = {"slide_count": len(slide_members)}
        details.update(props)
        return "\n".join(slides), details


def extract_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    root = read_xml_from_zip(archive, "xl/sharedStrings.xml")
    if root is None:
        return []
    strings: list[str] = []
    for item in root.findall(".//{*}si"):
        parts = [node.text or "" for node in item.findall(".//{*}t")]
        strings.append("".join(parts).strip())
    return strings


def load_workbook_sheet_paths(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = read_xml_from_zip(archive, "xl/workbook.xml")
    rels = read_xml_from_zip(archive, "xl/_rels/workbook.xml.rels")
    if workbook is None or rels is None:
        raise RuntimeError("Workbook metadata missing")
    rel_map: dict[str, str] = {}
    for item in rels.findall(".//pr:Relationship", PKG_REL_NS):
        rel_id = item.attrib.get("Id")
        target = item.attrib.get("Target")
        if rel_id and target:
            rel_map[rel_id] = resolve_zip_target("xl/workbook.xml", target)

    sheets: list[tuple[str, str]] = []
    for sheet in workbook.findall(".//{*}sheet"):
        name = (sheet.attrib.get("name") or "").strip()
        rel_id = sheet.attrib.get(f"{{{REL_NS['r']}}}id")
        target = rel_map.get(rel_id or "")
        if name and target:
            sheets.append((name, target))
    return sheets


def extract_cell_value(cell: ET.Element, shared_strings: Sequence[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        parts = [node.text or "" for node in cell.findall(".//{*}t")]
        return "".join(parts).strip()

    value = cell.findtext("{*}v", default="").strip()
    if not value:
        formula = cell.findtext("{*}f", default="").strip()
        return formula

    if cell_type == "s":
        try:
            index = int(value)
        except ValueError:
            return value
        if 0 <= index < len(shared_strings):
            return shared_strings[index]
        return value

    if cell_type == "b":
        return "TRUE" if value == "1" else "FALSE"

    return value


def extract_xlsx(path: Path) -> tuple[str, dict[str, Any]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = extract_shared_strings(archive)
        sheet_specs = load_workbook_sheet_paths(archive)
        lines: list[str] = []
        nonempty_cells = 0
        for sheet_name, member in sheet_specs:
            root = read_xml_from_zip(archive, member)
            if root is None:
                continue
            lines.append(f"[Sheet] {sheet_name}")
            for row in root.findall(".//{*}row"):
                row_values: list[str] = []
                for cell in row.findall("{*}c"):
                    value = extract_cell_value(cell, shared_strings)
                    if not value:
                        continue
                    nonempty_cells += 1
                    reference = cell.attrib.get("r")
                    if reference:
                        row_values.append(f"{reference}={value}")
                    else:
                        row_values.append(value)
                if row_values:
                    lines.append(" | ".join(row_values))
        props = read_core_properties(archive)
        details: dict[str, Any] = {
            "sheet_count": len(sheet_specs),
            "nonempty_cell_count": nonempty_cells,
        }
        details.update(props)
        return "\n".join(lines), details


def extract_pdf_with_pdftotext(path: Path) -> tuple[str, dict[str, Any]]:
    executable = shutil.which("pdftotext")
    if not executable:
        raise RuntimeError("pdftotext not available")
    command = [executable, "-enc", "UTF-8", "-nopgbrk", str(path), "-"]
    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if result.returncode != 0:
        stderr = result.stderr.strip() or "pdftotext failed"
        raise RuntimeError(stderr)
    return result.stdout or "", {"extractor": "pdftotext"}


def extract_pdf_with_pypdf(path: Path) -> tuple[str, dict[str, Any]]:
    try:
        from pypdf import PdfReader  # type: ignore
    except ModuleNotFoundError as exc:
        raise RuntimeError("pypdf not available") from exc

    reader = PdfReader(str(path))
    parts: list[str] = []
    for page in reader.pages:
        parts.append(page.extract_text() or "")
    details: dict[str, Any] = {"extractor": "pypdf", "page_count": len(reader.pages)}
    metadata = getattr(reader, "metadata", None)
    title = getattr(metadata, "title", None) if metadata is not None else None
    if title:
        details["title"] = str(title)
    return "\n".join(parts), details


def extract_pdf(path: Path) -> tuple[str, dict[str, Any]]:
    try:
        return extract_pdf_with_pdftotext(path)
    except RuntimeError:
        return extract_pdf_with_pypdf(path)


def extract_doc_with_antiword(path: Path) -> tuple[str, dict[str, Any]]:
    executable = shutil.which("antiword")
    if not executable:
        raise RuntimeError("antiword not available")
    env = os.environ.copy()
    env.setdefault("HOME", str(Path.home()))
    result = subprocess.run(
        [executable, "-t", str(path)],
        capture_output=True,
        text=False,
        env=env,
        check=False,
    )
    if result.returncode != 0:
        stderr_bytes = result.stderr or b""
        stderr = stderr_bytes.decode("utf-8", errors="replace").strip()
        if not stderr:
            stderr = stderr_bytes.decode("gb18030", errors="replace").strip()
        stderr = stderr or "antiword failed"
        raise RuntimeError(stderr)
    stdout_bytes = result.stdout or b""
    try:
        text = stdout_bytes.decode("utf-8")
    except UnicodeDecodeError:
        text = stdout_bytes.decode("gb18030", errors="replace")
    return text, {"extractor": "antiword"}


def clean_heuristic_doc_line(text: str) -> str:
    text = " ".join(text.split())
    text = text.strip(" -_|:;,./\\")
    if len(text) < 2:
        return ""
    if text.lower().startswith("font") or text in {"Calibri", "Symbol", "Wingdings", "Times New Roman"}:
        return ""
    if text.count(" ") > 12:
        text = clip_text(text, 160)
    return text


def extract_doc_with_ole_heuristic(path: Path) -> tuple[str, dict[str, Any]]:
    if olefile is None:
        raise RuntimeError("olefile not available")
    if not olefile.isOleFile(str(path)):
        raise RuntimeError("Not an OLE compound document")

    lines: list[str] = []
    seen: set[str] = set()
    stream_names = [("WordDocument",), ("1Table",), ("0Table",), ("Data",)]
    with olefile.OleFileIO(str(path)) as ole:
        for stream_name in stream_names:
            if not ole.exists(stream_name):
                continue
            data = ole.openstream(stream_name).read()
            text = data.decode("utf-16le", errors="ignore")
            hits = re.findall(r"[A-Za-z0-9\u4e00-\u9fff][A-Za-z0-9\u4e00-\u9fff\s\-_/:：（）()，,。\.]{1,160}", text)
            for item in hits:
                cleaned = clean_heuristic_doc_line(item)
                if not cleaned or cleaned in seen:
                    continue
                seen.add(cleaned)
                lines.append(cleaned)
            if len(lines) >= 400:
                break

    if len(lines) < 5:
        raise RuntimeError("OLE heuristic did not recover enough text")

    return "\n".join(lines), {"extractor": "ole-heuristic", "line_count": len(lines)}


def extract_doc_filename_fallback(path: Path) -> tuple[str, dict[str, Any]]:
    text = path.stem
    return text, {"extractor": "filename-fallback", "partial": True}


def extract_xls_with_xlrd(path: Path) -> tuple[str, dict[str, Any]]:
    if xlrd is None:
        raise RuntimeError("xlrd not available")
    workbook = xlrd.open_workbook(str(path), on_demand=True)
    try:
        lines: list[str] = []
        nonempty_cells = 0
        for sheet in workbook.sheets():
            lines.append(f"[Sheet] {sheet.name}")
            for row_index in range(sheet.nrows):
                rendered: list[str] = []
                for col_index in range(sheet.ncols):
                    value = sheet.cell_value(row_index, col_index)
                    text = str(value).strip()
                    if not text:
                        continue
                    nonempty_cells += 1
                    cell_ref = f"{excel_column_label(col_index + 1)}{row_index + 1}"
                    rendered.append(f"{cell_ref}={text}")
                if rendered:
                    lines.append(" | ".join(rendered))
        details = {
            "extractor": "xlrd",
            "sheet_count": workbook.nsheets,
            "nonempty_cell_count": nonempty_cells,
        }
        return "\n".join(lines), details
    finally:
        workbook.release_resources()


def extract_legacy_with_subprocess(path: Path, kind: str) -> tuple[str, dict[str, Any]]:
    helper = Path(__file__).with_name("office_legacy_win32.py")
    command = [sys.executable, str(helper), kind, str(path)]
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=LEGACY_EXTRACT_TIMEOUT_SECONDS,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            f"Legacy .{kind} extraction timed out after {LEGACY_EXTRACT_TIMEOUT_SECONDS}s"
        ) from exc

    if result.returncode != 0:
        stderr = (result.stderr or "").strip()
        stdout = (result.stdout or "").strip()
        message = stderr or stdout or f"Legacy .{kind} extraction failed"
        raise RuntimeError(message)

    try:
        payload = json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Legacy .{kind} extraction returned invalid JSON") from exc

    text = str(payload.get("text") or "")
    details = payload.get("details")
    if not isinstance(details, dict):
        details = {}
    return text, details


def probe_legacy_support(kind: str, timeout_seconds: int = 20) -> tuple[bool, Optional[str]]:
    helper = Path(__file__).with_name("office_legacy_win32.py")
    command = [sys.executable, str(helper), kind, "--probe"]
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return False, f"startup timed out after {timeout_seconds}s"
    if result.returncode == 0:
        return True, None
    message = (result.stderr or "").strip() or (result.stdout or "").strip() or "startup failed"
    return False, message


def extract_doc(path: Path, allow_legacy_com: bool) -> tuple[str, dict[str, Any]]:
    try:
        return extract_doc_with_antiword(path)
    except Exception:
        pass
    try:
        return extract_doc_with_ole_heuristic(path)
    except Exception:
        pass
    if allow_legacy_com and extract_doc_via_word is not None:
        try:
            return extract_legacy_with_subprocess(path, "doc")
        except Exception:
            pass
    return extract_doc_filename_fallback(path)


def extract_xls(path: Path, allow_legacy_com: bool) -> tuple[str, dict[str, Any]]:
    if xlrd is not None:
        try:
            return extract_xls_with_xlrd(path)
        except Exception:
            pass
    if not allow_legacy_com:
        raise RuntimeError(
            "Legacy .xls extraction needs xlrd or explicit --allow-legacy-com. "
            "COM is disabled by default to avoid Microsoft Office repair popups. Re-save as .xlsx if possible."
        )
    if extract_xls_via_excel is None:
        raise RuntimeError("Legacy .xls extraction requires Microsoft Excel via COM on Windows.")
    return extract_legacy_with_subprocess(path, "xls")


def extract_ppt(path: Path, allow_legacy_com: bool) -> tuple[str, dict[str, Any]]:
    if not allow_legacy_com:
        raise RuntimeError(
            "Legacy .ppt extraction is disabled by default to avoid Microsoft Office repair popups. "
            "Re-save the file as .pptx or rerun with --allow-legacy-com."
        )
    if extract_ppt_via_powerpoint is None:
        raise RuntimeError("Legacy .ppt extraction requires Microsoft PowerPoint via COM on Windows.")
    return extract_legacy_with_subprocess(path, "ppt")


def extract_file(path: Path, root: Path, max_chars: int, allow_legacy_com: bool = False) -> ArchiveEntry:
    stat = path.stat()
    relative = path.relative_to(root).as_posix()
    extension = path.suffix.lower()
    file_type = SUPPORTED_EXTENSIONS[extension]
    indexed_at = current_iso()

    try:
        if extension == ".doc":
            raw_text, details = extract_doc(path, allow_legacy_com=allow_legacy_com)
        elif extension == ".docx":
            raw_text, details = extract_docx(path)
        elif extension == ".xls":
            raw_text, details = extract_xls(path, allow_legacy_com=allow_legacy_com)
        elif extension in {".xlsx", ".xlsm"}:
            raw_text, details = extract_xlsx(path)
        elif extension == ".ppt":
            raw_text, details = extract_ppt(path, allow_legacy_com=allow_legacy_com)
        elif extension == ".pptx":
            raw_text, details = extract_pptx(path)
        elif extension == ".pdf":
            raw_text, details = extract_pdf(path)
        else:
            raise RuntimeError(f"Unsupported extension: {extension}")

        text = clip_text(normalize_whitespace(raw_text), max_chars)
        title = str(details.get("title") or path.stem)
        summary = build_summary(file_type=file_type, text=text, title=title)
        preview = build_preview(text)
        return ArchiveEntry(
            path=relative,
            file_type=file_type,
            size_bytes=stat.st_size,
            modified_at=iso_from_timestamp(stat.st_mtime),
            mtime_ns=stat.st_mtime_ns,
            title=title,
            summary=summary,
            preview=preview,
            content=text,
            content_hash=sha1_text(text),
            status="ok",
            error=None,
            details=details,
            indexed_at=indexed_at,
        )
    except Exception as exc:
        error_message = str(exc)
        title = path.stem
        return ArchiveEntry(
            path=relative,
            file_type=file_type,
            size_bytes=stat.st_size,
            modified_at=iso_from_timestamp(stat.st_mtime),
            mtime_ns=stat.st_mtime_ns,
            title=title,
            summary=f"{title}: extraction failed.",
            preview="",
            content="",
            content_hash=sha1_text(error_message),
            status="error",
            error=error_message,
            details={},
            indexed_at=indexed_at,
        )


def should_skip(path: Path, root: Path, skip_dirs: set[str]) -> bool:
    try:
        relative = path.relative_to(root)
    except ValueError:
        return True
    for part in relative.parts[:-1]:
        if part in skip_dirs:
            return True
    return False


def iter_supported_files(root: Path, skip_dirs: set[str]) -> list[Path]:
    files: list[Path] = []
    for path in root.rglob("*"):
        if not path.is_file() or path.is_symlink():
            continue
        if path.name.startswith("~$"):
            continue
        if path.stat().st_size == 0:
            continue
        if should_skip(path, root, skip_dirs):
            continue
        if path.suffix.lower() in SUPPORTED_EXTENSIONS:
            files.append(path)
    files.sort(key=lambda item: item.relative_to(root).as_posix())
    return files


def connect_index(db_path: Path) -> tuple[sqlite3.Connection, bool]:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS docs (
            path TEXT PRIMARY KEY,
            file_type TEXT NOT NULL,
            size_bytes INTEGER NOT NULL,
            modified_at TEXT NOT NULL,
            mtime_ns INTEGER NOT NULL,
            title TEXT NOT NULL,
            summary TEXT NOT NULL,
            preview TEXT NOT NULL,
            content TEXT NOT NULL,
            content_hash TEXT NOT NULL,
            status TEXT NOT NULL,
            error TEXT,
            details_json TEXT NOT NULL,
            indexed_at TEXT NOT NULL
        )
        """
    )
    fts_enabled = True
    try:
        connection.execute(
            """
            CREATE VIRTUAL TABLE IF NOT EXISTS docs_fts USING fts5(
                path,
                title,
                summary,
                preview,
                content,
                tokenize = 'unicode61'
            )
            """
        )
    except sqlite3.OperationalError as exc:
        if "fts5" not in str(exc).lower():
            raise
        fts_enabled = False
    return connection, fts_enabled


def upsert_entry(connection: sqlite3.Connection, entry: ArchiveEntry, fts_enabled: bool) -> None:
    connection.execute(
        """
        INSERT INTO docs (
            path, file_type, size_bytes, modified_at, mtime_ns, title, summary, preview,
            content, content_hash, status, error, details_json, indexed_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(path) DO UPDATE SET
            file_type = excluded.file_type,
            size_bytes = excluded.size_bytes,
            modified_at = excluded.modified_at,
            mtime_ns = excluded.mtime_ns,
            title = excluded.title,
            summary = excluded.summary,
            preview = excluded.preview,
            content = excluded.content,
            content_hash = excluded.content_hash,
            status = excluded.status,
            error = excluded.error,
            details_json = excluded.details_json,
            indexed_at = excluded.indexed_at
        """,
        (
            entry.path,
            entry.file_type,
            entry.size_bytes,
            entry.modified_at,
            entry.mtime_ns,
            entry.title,
            entry.summary,
            entry.preview,
            entry.content,
            entry.content_hash,
            entry.status,
            entry.error,
            json.dumps(entry.details, ensure_ascii=False, sort_keys=True),
            entry.indexed_at,
        ),
    )
    if not fts_enabled:
        return
    connection.execute("DELETE FROM docs_fts WHERE path = ?", (entry.path,))
    if entry.status == "ok":
        connection.execute(
            "INSERT INTO docs_fts (path, title, summary, preview, content) VALUES (?, ?, ?, ?, ?)",
            (entry.path, entry.title, entry.summary, entry.preview, entry.content),
        )


def delete_paths(connection: sqlite3.Connection, paths: Iterable[str], fts_enabled: bool) -> None:
    for item in paths:
        connection.execute("DELETE FROM docs WHERE path = ?", (item,))
        if fts_enabled:
            connection.execute("DELETE FROM docs_fts WHERE path = ?", (item,))


def load_entries(connection: sqlite3.Connection) -> list[ArchiveEntry]:
    rows = connection.execute("SELECT * FROM docs ORDER BY path").fetchall()
    entries: list[ArchiveEntry] = []
    for row in rows:
        details_json = row["details_json"] if isinstance(row["details_json"], str) else "{}"
        details = json.loads(details_json)
        entries.append(
            ArchiveEntry(
                path=row["path"],
                file_type=row["file_type"],
                size_bytes=int(row["size_bytes"]),
                modified_at=row["modified_at"],
                mtime_ns=int(row["mtime_ns"]),
                title=row["title"],
                summary=row["summary"],
                preview=row["preview"],
                content=row["content"],
                content_hash=row["content_hash"],
                status=row["status"],
                error=row["error"],
                details=details,
                indexed_at=row["indexed_at"],
            )
        )
    return entries


def write_jsonl(entries: Sequence[ArchiveEntry], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        for entry in entries:
            handle.write(json.dumps(asdict(entry), ensure_ascii=False, sort_keys=True))
            handle.write("\n")


def write_markdown(entries: Sequence[ArchiveEntry], root: Path, path: Path) -> None:
    ok_count = sum(1 for entry in entries if entry.status == "ok")
    lines = [
        "# Office Archive",
        "",
        f"- Root: {root}",
        f"- Generated: {current_iso()}",
        f"- Files: {len(entries)}",
        f"- Extracted: {ok_count}",
        "",
    ]
    for entry in entries:
        lines.append(f"## {entry.path}")
        lines.append("")
        lines.append(f"- Type: {entry.file_type}")
        lines.append(f"- Modified: {entry.modified_at}")
        lines.append(f"- Size: {entry.size_bytes} bytes")
        lines.append(f"- Status: {entry.status}")
        lines.append(f"- Summary: {entry.summary}")
        if entry.error:
            lines.append(f"- Error: {entry.error}")
        if entry.preview:
            lines.append("- Preview:")
            for preview_line in entry.preview.splitlines():
                lines.append(f"  {preview_line}")
        lines.append("")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def build_match_query(query: str) -> str:
    tokens = re.findall(r"[0-9A-Za-z_\u0080-\uffff]+", query.lower())
    if not tokens:
        raise RuntimeError("Search query does not contain any searchable tokens.")
    return " ".join(f"{token}*" for token in tokens)


def build_capabilities_report(probe_legacy_com: bool = False) -> dict[str, Any]:
    antiword = shutil.which("antiword")
    pdftotext = shutil.which("pdftotext")
    if probe_legacy_com:
        word_ok, word_error = probe_legacy_support("doc") if extract_doc_via_word is not None else (False, None)
        excel_ok, excel_error = probe_legacy_support("xls") if extract_xls_via_excel is not None else (False, None)
        powerpoint_ok, powerpoint_error = (
            probe_legacy_support("ppt") if extract_ppt_via_powerpoint is not None else (False, None)
        )
    else:
        word_ok, word_error = False, None
        excel_ok, excel_error = False, None
        powerpoint_ok, powerpoint_error = False, None
    legacy_support = {
        "doc": bool(word_ok or antiword),
        "xls": bool(xlrd or excel_ok),
        "ppt": powerpoint_ok,
    }
    return {
        "python": sys.executable,
        "platform": sys.platform,
        "modernFormats": {
            "docx": True,
            "xlsx": True,
            "xlsm": True,
            "pptx": True,
            "pdf": True,
        },
        "legacyFormats": legacy_support,
        "helpers": {
            "pdftotext": pdftotext,
            "antiword": antiword,
            "olefile": olefile is not None,
            "xlrd": xlrd is not None,
            "legacyComDisabledByDefault": True,
            "legacyComProbed": probe_legacy_com,
            "wordCom": word_ok,
            "excelCom": excel_ok,
            "powerpointCom": powerpoint_ok,
            "wordComError": word_error,
            "excelComError": excel_error,
            "powerpointComError": powerpoint_error,
        },
    }


def print_capabilities(report: dict[str, Any]) -> None:
    print(f"Python: {report['python']}")
    print(f"Platform: {report['platform']}")
    print("Modern formats:")
    for extension, supported in report["modernFormats"].items():
        print(f"  - {extension}: {'yes' if supported else 'no'}")
    print("Legacy formats:")
    for extension, supported in report["legacyFormats"].items():
        print(f"  - {extension}: {'yes' if supported else 'no'}")
    print("Helpers:")
    helpers = report["helpers"]
    print(f"  - pdftotext: {helpers['pdftotext'] or 'missing'}")
    print(f"  - antiword: {helpers['antiword'] or 'missing'}")
    print(f"  - olefile: {'yes' if helpers['olefile'] else 'no'}")
    print(f"  - xlrd: {'yes' if helpers['xlrd'] else 'no'}")
    print(f"  - Legacy COM disabled by default: {'yes' if helpers['legacyComDisabledByDefault'] else 'no'}")
    print(f"  - Legacy COM probed: {'yes' if helpers['legacyComProbed'] else 'no'}")
    print(f"  - Word COM: {'yes' if helpers['wordCom'] else 'no'}")
    if helpers.get("wordComError"):
        print(f"    error: {helpers['wordComError']}")
    print(f"  - Excel COM: {'yes' if helpers['excelCom'] else 'no'}")
    if helpers.get("excelComError"):
        print(f"    error: {helpers['excelComError']}")
    print(f"  - PowerPoint COM: {'yes' if helpers['powerpointCom'] else 'no'}")
    if helpers.get("powerpointComError"):
        print(f"    error: {helpers['powerpointComError']}")


def archive_directory(
    root: Path,
    state_dir: Path,
    max_chars: int,
    skip_dirs: set[str],
    force: bool,
    write_outputs: bool,
    allow_legacy_com: bool,
) -> tuple[list[ArchiveEntry], dict[str, int]]:
    db_path = state_dir / "index.sqlite"
    connection, fts_enabled = connect_index(db_path)
    try:
        connection.execute("BEGIN")
        files = iter_supported_files(root, skip_dirs)
        existing_rows = {
            row["path"]: row
            for row in connection.execute("SELECT path, size_bytes, mtime_ns FROM docs").fetchall()
        }
        seen_paths: set[str] = set()
        stats = {"scanned": 0, "updated": 0, "reused": 0, "errors": 0}

        for path in files:
            relative = path.relative_to(root).as_posix()
            stat = path.stat()
            previous = existing_rows.get(relative)
            seen_paths.add(relative)
            stats["scanned"] += 1

            if (
                not force
                and previous is not None
                and int(previous["size_bytes"]) == stat.st_size
                and int(previous["mtime_ns"]) == stat.st_mtime_ns
            ):
                stats["reused"] += 1
                continue

            entry = extract_file(path=path, root=root, max_chars=max_chars, allow_legacy_com=allow_legacy_com)
            if entry.status != "ok":
                stats["errors"] += 1
            upsert_entry(connection, entry, fts_enabled=fts_enabled)
            stats["updated"] += 1

        stale_paths = set(existing_rows.keys()) - seen_paths
        delete_paths(connection, stale_paths, fts_enabled=fts_enabled)
        connection.commit()

        entries = load_entries(connection)
        if write_outputs:
            write_jsonl(entries, state_dir / "archive.jsonl")
            write_markdown(entries, root=root, path=state_dir / "archive.md")
        return entries, stats
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def search_like(connection: sqlite3.Connection, query: str, limit: int) -> list[dict[str, Any]]:
    terms = [token for token in re.findall(r"[0-9A-Za-z_\u0080-\uffff]+", query.lower()) if token]
    if not terms:
        raise RuntimeError("Search query does not contain any searchable tokens.")

    rows = connection.execute(
        "SELECT path, file_type, title, summary, preview, content FROM docs WHERE status = 'ok' ORDER BY path"
    ).fetchall()
    results: list[dict[str, Any]] = []
    for row in rows:
        haystack = "\n".join(
            [row["path"], row["title"], row["summary"], row["preview"], row["content"]]
        ).lower()
        matches = sum(1 for term in terms if term in haystack)
        if matches == 0:
            continue
        excerpt = row["preview"] or clip_text(row["content"], 240)
        results.append(
            {
                "path": row["path"],
                "file_type": row["file_type"],
                "title": row["title"],
                "summary": row["summary"],
                "excerpt": excerpt,
                "score": matches,
            }
        )
    results.sort(key=lambda item: (-int(item["score"]), item["path"]))
    return results[:limit]


def search_index(root: Path, state_dir: Path, query: str, limit: int) -> list[dict[str, Any]]:
    db_path = state_dir / "index.sqlite"
    if not db_path.exists():
        raise RuntimeError(f"Index not found: {db_path}")
    connection, fts_enabled = connect_index(db_path)
    try:
        if not fts_enabled:
            return search_like(connection, query=query, limit=limit)

        match_query = build_match_query(query)
        try:
            rows = connection.execute(
                """
                SELECT
                    docs.path AS path,
                    docs.file_type AS file_type,
                    docs.title AS title,
                    docs.summary AS summary,
                    snippet(docs_fts, 4, '[', ']', ' ... ', 16) AS excerpt,
                    bm25(docs_fts, 0.5, 4.0, 3.0, 1.5, 1.0) AS score
                FROM docs_fts
                JOIN docs ON docs.path = docs_fts.path
                WHERE docs_fts MATCH ? AND docs.status = 'ok'
                ORDER BY score ASC, docs.path ASC
                LIMIT ?
                """,
                (match_query, limit),
            ).fetchall()
        except sqlite3.OperationalError:
            return search_like(connection, query=query, limit=limit)

        if not rows:
            return search_like(connection, query=query, limit=limit)

        results: list[dict[str, Any]] = []
        for row in rows:
            excerpt = row["excerpt"] or row["summary"]
            results.append(
                {
                    "path": row["path"],
                    "file_type": row["file_type"],
                    "title": row["title"],
                    "summary": row["summary"],
                    "excerpt": excerpt,
                    "score": float(row["score"]),
                }
            )
        return results
    finally:
        connection.close()


def print_archive_result(root: Path, state_dir: Path, stats: dict[str, int], entries: Sequence[ArchiveEntry]) -> None:
    ok_count = sum(1 for entry in entries if entry.status == "ok")
    print(f"Root: {root}")
    print(f"State dir: {state_dir}")
    print(f"Files indexed: {len(entries)}")
    print(f"Extracted successfully: {ok_count}")
    print(f"Updated: {stats['updated']}")
    print(f"Reused: {stats['reused']}")
    print(f"Errors: {stats['errors']}")
    print(f"Archive JSONL: {state_dir / 'archive.jsonl'}")
    print(f"Archive Markdown: {state_dir / 'archive.md'}")
    print(f"Index SQLite: {state_dir / 'index.sqlite'}")


def print_search_results(results: Sequence[dict[str, Any]]) -> None:
    if not results:
        print("No matches found.")
        return
    for index, item in enumerate(results, start=1):
        print(f"{index}. {item['path']} [{item['file_type']}]")
        print(f"   Summary: {item['summary']}")
        print(f"   Match: {item['excerpt']}")


def print_inspect_result(entry: ArchiveEntry) -> None:
    print(f"Path: {entry.path}")
    print(f"Type: {entry.file_type}")
    print(f"Status: {entry.status}")
    print(f"Summary: {entry.summary}")
    if entry.error:
        print(f"Error: {entry.error}")
    if entry.preview:
        print("Preview:")
        print(textwrap.indent(entry.preview, "  "))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Archive and search Office-style document folders.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    archive_parser = subparsers.add_parser("archive", help="Index a folder and write archive outputs.")
    archive_parser.add_argument("root", help="Folder to scan.")
    archive_parser.add_argument(
        "--state-dir",
        help="Directory for index.sqlite, archive.jsonl, and archive.md. Defaults to <root>/.office-archive.",
    )
    archive_parser.add_argument("--max-chars", type=positive_int, default=120_000)
    archive_parser.add_argument("--skip-dir", action="append", default=[], help="Directory name to skip.")
    archive_parser.add_argument("--force", action="store_true", help="Re-extract all supported files.")
    archive_parser.add_argument(
        "--allow-legacy-com",
        action="store_true",
        help="Allow Microsoft Office COM automation for legacy .doc/.xls/.ppt files. Disabled by default to avoid repair popups.",
    )
    archive_parser.add_argument("--json", action="store_true", help="Print machine-readable summary.")

    search_parser = subparsers.add_parser("search", help="Search an indexed folder.")
    search_parser.add_argument("root", help="Folder that was archived.")
    search_parser.add_argument("query", help="Search text.")
    search_parser.add_argument("--state-dir", help="Directory containing index.sqlite.")
    search_parser.add_argument("--limit", type=positive_int, default=10)
    search_parser.add_argument("--max-chars", type=positive_int, default=120_000)
    search_parser.add_argument("--skip-dir", action="append", default=[], help="Directory name to skip.")
    search_parser.add_argument("--no-refresh", action="store_true", help="Search without updating the index first.")
    search_parser.add_argument(
        "--allow-legacy-com",
        action="store_true",
        help="Allow Microsoft Office COM automation during refresh for legacy .doc/.xls/.ppt files.",
    )
    search_parser.add_argument("--json", action="store_true", help="Print JSON results.")

    inspect_parser = subparsers.add_parser("inspect", help="Extract summary text from a single file.")
    inspect_parser.add_argument("file", help="Path to a supported file.")
    inspect_parser.add_argument("--root", help="Root folder for relative-path display.")
    inspect_parser.add_argument("--max-chars", type=positive_int, default=120_000)
    inspect_parser.add_argument(
        "--allow-legacy-com",
        action="store_true",
        help="Allow Microsoft Office COM automation for legacy .doc/.xls/.ppt files. Disabled by default.",
    )
    inspect_parser.add_argument("--json", action="store_true", help="Print JSON output.")

    check_parser = subparsers.add_parser("check", help="Show extractor support for this machine.")
    check_parser.add_argument(
        "--probe-legacy-com",
        action="store_true",
        help="Actually start Word/Excel/PowerPoint to test COM availability. Disabled by default to avoid popups.",
    )
    check_parser.add_argument("--json", action="store_true", help="Print JSON output.")

    return parser


def resolve_state_dir(root: Path, state_dir: Optional[str]) -> Path:
    if state_dir:
        return Path(state_dir).expanduser().resolve()
    return root / ".office-archive"


def ensure_root(path_text: str) -> Path:
    root = Path(path_text).expanduser().resolve()
    if not root.exists():
        raise RuntimeError(f"Path does not exist: {root}")
    if not root.is_dir():
        raise RuntimeError(f"Not a directory: {root}")
    return root


def ensure_file(path_text: str) -> Path:
    path = Path(path_text).expanduser().resolve()
    if not path.exists():
        raise RuntimeError(f"File does not exist: {path}")
    if not path.is_file():
        raise RuntimeError(f"Not a file: {path}")
    extension = path.suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        raise RuntimeError(f"Unsupported file type: {extension}. Supported: {supported}")
    return path


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        if args.command == "archive":
            root = ensure_root(args.root)
            skip_dirs = set(DEFAULT_SKIP_DIRS)
            skip_dirs.update(args.skip_dir)
            state_dir = resolve_state_dir(root, args.state_dir)
            entries, stats = archive_directory(
                root=root,
                state_dir=state_dir,
                max_chars=args.max_chars,
                skip_dirs=skip_dirs,
                force=args.force,
                write_outputs=True,
                allow_legacy_com=args.allow_legacy_com,
            )
            if args.json:
                payload = {
                    "root": str(root),
                    "stateDir": str(state_dir),
                    "stats": stats,
                    "entries": len(entries),
                    "ok": sum(1 for entry in entries if entry.status == "ok"),
                }
                print(json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True))
            else:
                print_archive_result(root=root, state_dir=state_dir, stats=stats, entries=entries)
            return 0

        if args.command == "search":
            root = ensure_root(args.root)
            skip_dirs = set(DEFAULT_SKIP_DIRS)
            skip_dirs.update(args.skip_dir)
            state_dir = resolve_state_dir(root, args.state_dir)
            if not args.no_refresh:
                archive_directory(
                    root=root,
                    state_dir=state_dir,
                    max_chars=args.max_chars,
                    skip_dirs=skip_dirs,
                    force=False,
                    write_outputs=False,
                    allow_legacy_com=args.allow_legacy_com,
                )
            results = search_index(root=root, state_dir=state_dir, query=args.query, limit=args.limit)
            if args.json:
                print(json.dumps(results, ensure_ascii=False, indent=2, sort_keys=True))
            else:
                print_search_results(results)
            return 0 if results else 2

        if args.command == "inspect":
            path = ensure_file(args.file)
            root = Path(args.root).expanduser().resolve() if args.root else path.parent
            entry = extract_file(
                path=path,
                root=root,
                max_chars=args.max_chars,
                allow_legacy_com=args.allow_legacy_com,
            )
            if args.json:
                print(json.dumps(asdict(entry), ensure_ascii=False, indent=2, sort_keys=True))
            else:
                print_inspect_result(entry)
            return 0 if entry.status == "ok" else 2

        if args.command == "check":
            report = build_capabilities_report(probe_legacy_com=args.probe_legacy_com)
            if args.json:
                print(json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True))
            else:
                print_capabilities(report)
            return 0

        parser.error(f"Unsupported command: {args.command}")
        return 1
    except Exception as exc:
        eprint(str(exc))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
