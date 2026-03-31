#!/usr/bin/env python3
"""
Windows-only helpers for legacy Office binary formats.
"""

from __future__ import annotations

import argparse
import json
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator, Optional

import pythoncom
import win32com.client


def _safe_builtin_property(container: Any, name: str) -> Optional[str]:
    try:
        value = container.BuiltInDocumentProperties(name).Value
    except Exception:
        return None
    if value is None:
        return None
    text = str(value).strip()
    return text or None


@contextmanager
def _dispatch(progid: str) -> Iterator[Any]:
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx(progid)
        yield app
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _excel_column_label(index: int) -> str:
    result = ""
    current = index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _stringify_excel_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _shape_text(shape: Any) -> str:
    try:
        if getattr(shape, "HasTextFrame", 0):
            text_frame = shape.TextFrame
            if getattr(text_frame, "HasText", 0):
                text = text_frame.TextRange.Text
                if text:
                    return str(text).strip()
    except Exception:
        pass
    try:
        if getattr(shape, "HasTextFrame", 0):
            text_frame2 = shape.TextFrame2
            text = text_frame2.TextRange.Text
            if text:
                return str(text).strip()
    except Exception:
        pass
    return ""


def extract_doc_via_word(path: Path) -> tuple[str, dict[str, Any]]:
    with _dispatch("Word.Application") as app:
        app.Visible = False
        app.DisplayAlerts = 0
        document = app.Documents.Open(
            str(path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        try:
            text = document.Content.Text or ""
            details: dict[str, Any] = {
                "extractor": "word-com",
                "paragraph_count": int(document.Paragraphs.Count),
            }
            title = _safe_builtin_property(document, "Title")
            subject = _safe_builtin_property(document, "Subject")
            if title:
                details["title"] = title
            if subject:
                details["subject"] = subject
            return text, details
        finally:
            document.Close(False)


def extract_xls_via_excel(path: Path) -> tuple[str, dict[str, Any]]:
    with _dispatch("Excel.Application") as app:
        app.Visible = False
        app.DisplayAlerts = False
        workbook = app.Workbooks.Open(
            str(path),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        try:
            lines: list[str] = []
            nonempty_cells = 0
            sheet_count = int(workbook.Worksheets.Count)
            for sheet in workbook.Worksheets:
                lines.append(f"[Sheet] {sheet.Name}")
                used_range = sheet.UsedRange
                start_row = int(used_range.Row)
                start_column = int(used_range.Column)
                rows_count = int(used_range.Rows.Count)
                columns_count = int(used_range.Columns.Count)
                values = used_range.Value
                if values is None:
                    continue
                if rows_count == 1 and columns_count == 1 and not isinstance(values, tuple):
                    values = ((values,),)
                elif rows_count == 1 and columns_count > 1 and not isinstance(values[0], tuple):
                    values = (values,)

                for row_offset, row_values in enumerate(values):
                    rendered: list[str] = []
                    for column_offset, cell_value in enumerate(row_values):
                        text = _stringify_excel_value(cell_value)
                        if not text:
                            continue
                        nonempty_cells += 1
                        row_number = start_row + row_offset
                        column_number = start_column + column_offset
                        cell_ref = f"{_excel_column_label(column_number)}{row_number}"
                        rendered.append(f"{cell_ref}={text}")
                    if rendered:
                        lines.append(" | ".join(rendered))

            details: dict[str, Any] = {
                "extractor": "excel-com",
                "sheet_count": sheet_count,
                "nonempty_cell_count": nonempty_cells,
            }
            title = _safe_builtin_property(workbook, "Title")
            subject = _safe_builtin_property(workbook, "Subject")
            if title:
                details["title"] = title
            if subject:
                details["subject"] = subject
            return "\n".join(lines), details
        finally:
            workbook.Close(False)


def extract_ppt_via_powerpoint(path: Path) -> tuple[str, dict[str, Any]]:
    with _dispatch("PowerPoint.Application") as app:
        presentation = app.Presentations.Open(str(path), ReadOnly=True, WithWindow=False)
        try:
            slides_text: list[str] = []
            slide_count = int(presentation.Slides.Count)
            for slide in presentation.Slides:
                parts: list[str] = []
                for shape in slide.Shapes:
                    text = _shape_text(shape)
                    if text:
                        parts.append(text)
                if parts:
                    slides_text.append(f"Slide {int(slide.SlideIndex)}: {' | '.join(parts)}")

            details: dict[str, Any] = {
                "extractor": "powerpoint-com",
                "slide_count": slide_count,
            }
            title = _safe_builtin_property(presentation, "Title")
            subject = _safe_builtin_property(presentation, "Subject")
            if title:
                details["title"] = title
            if subject:
                details["subject"] = subject
            return "\n".join(slides_text), details
        finally:
            presentation.Close()


def probe_app(kind: str) -> dict[str, Any]:
    progid = {
        "doc": "Word.Application",
        "xls": "Excel.Application",
        "ppt": "PowerPoint.Application",
    }[kind]
    with _dispatch(progid) as app:
        try:
            app.Visible = False
        except Exception:
            pass
    return {"ok": True, "kind": kind}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Extract legacy Office binary formats through COM.")
    parser.add_argument("kind", choices=["doc", "xls", "ppt"])
    parser.add_argument("path", nargs="?", help="Path to the legacy Office file.")
    parser.add_argument("--probe", action="store_true", help="Probe app startup only.")
    return parser


def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    args = build_parser().parse_args()
    if args.probe:
        print(json.dumps(probe_app(args.kind), ensure_ascii=False))
        return 0

    if not args.path:
        raise SystemExit("path is required unless --probe is used")

    path = Path(args.path).expanduser().resolve()
    if args.kind == "doc":
        text, details = extract_doc_via_word(path)
    elif args.kind == "xls":
        text, details = extract_xls_via_excel(path)
    else:
        text, details = extract_ppt_via_powerpoint(path)

    payload = {"text": text, "details": details}
    print(json.dumps(payload, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
