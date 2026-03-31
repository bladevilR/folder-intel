#!/usr/bin/env python3
"""
Tests for office_archive helpers.
"""

from __future__ import annotations

import json
import shutil
import tempfile
import zipfile
from contextlib import contextmanager
from pathlib import Path
from unittest import SkipTest, TestCase, main, skipUnless

import pywintypes

from office_archive import archive_directory, build_capabilities_report, extract_file, search_index

try:
    import pythoncom
    import win32com.client
except Exception:
    pythoncom = None
    win32com = None

HAS_OFFICE_COM = pythoncom is not None and win32com is not None


def write_zip(path: Path, members: dict[str, str]) -> None:
    with zipfile.ZipFile(path, "w") as archive:
        for name, content in members.items():
            archive.writestr(name, content)


@contextmanager
def dispatch_app(progid: str):
    if not HAS_OFFICE_COM:
        raise RuntimeError("Office COM automation unavailable")
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


def retry_com(operation, attempts: int = 3):
    last_error = None
    for _ in range(attempts):
        try:
            return operation()
        except pywintypes.com_error as exc:
            last_error = exc
    if last_error is not None:
        raise last_error
    raise RuntimeError("retry_com() called without attempts")


def make_docx(path: Path) -> None:
    write_zip(
        path,
        {
            "word/document.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Quarterly revenue summary</w:t></w:r></w:p>
    <w:p><w:r><w:t>Customer pipeline and renewals</w:t></w:r></w:p>
  </w:body>
</w:document>
""",
            "docProps/core.xml": """<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Revenue Memo</dc:title>
</cp:coreProperties>
""",
        },
    )


def make_docx_cjk(path: Path) -> None:
    write_zip(
        path,
        {
            "word/document.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>合同编号 A-2026-04</w:t></w:r></w:p>
    <w:p><w:r><w:t>供应商 测试公司</w:t></w:r></w:p>
  </w:body>
</w:document>
""",
        },
    )


def make_xlsx(path: Path) -> None:
    write_zip(
        path,
        {
            "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Deals" sheetId="1" r:id="rId1" />
  </sheets>
</workbook>
""",
            "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml" />
</Relationships>
""",
            "xl/sharedStrings.xml": """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Deal</t></si>
  <si><t>Amount</t></si>
  <si><t>Acme</t></si>
</sst>
""",
            "xl/worksheets/sheet1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2"><v>42000</v></c>
    </row>
  </sheetData>
</worksheet>
""",
        },
    )


def make_pptx(path: Path) -> None:
    write_zip(
        path,
        {
            "ppt/slides/slide1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Launch plan</a:t></a:r></a:p>
          <a:p><a:r><a:t>Migration timeline</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
""",
            "docProps/core.xml": """<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Launch Deck</dc:title>
</cp:coreProperties>
""",
        },
    )


def make_pdf(path: Path) -> None:
    pdf = """%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 144] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>
endobj
4 0 obj
<< /Length 58 >>
stream
BT
/F1 18 Tf
40 90 Td
(Invoice April 2026) Tj
ET
endstream
endobj
5 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj
xref
0 6
0000000000 65535 f 
0000000010 00000 n 
0000000063 00000 n 
0000000122 00000 n 
0000000248 00000 n 
0000000356 00000 n 
trailer
<< /Root 1 0 R /Size 6 >>
startxref
426
%%EOF
    """
    path.write_text(pdf, encoding="utf-8")


def make_doc(path: Path) -> None:
    with dispatch_app("Word.Application") as app:
        app.Visible = False
        app.DisplayAlerts = 0
        document = app.Documents.Add()
        try:
            document.Content.Text = "Legacy contract terms\nRenewal due in May"
            document.BuiltInDocumentProperties("Title").Value = "Legacy Contract"
            document.SaveAs(str(path), FileFormat=0)
        finally:
            document.Close(False)


def make_xls(path: Path) -> None:
    with dispatch_app("Excel.Application") as app:
        app.Visible = False
        app.DisplayAlerts = False
        workbook = retry_com(lambda: app.Workbooks.Add())
        try:
            sheet = workbook.Worksheets(1)
            sheet.Name = "LegacyDeals"
            sheet.Cells(1, 1).Value = "Customer"
            sheet.Cells(1, 2).Value = "Amount"
            sheet.Cells(2, 1).Value = "OldCo"
            sheet.Cells(2, 2).Value = 12500
            workbook.BuiltinDocumentProperties("Title").Value = "Legacy Pipeline"
            workbook.SaveAs(str(path), FileFormat=56)
        finally:
            workbook.Close(False)


def make_ppt(path: Path) -> None:
    with dispatch_app("PowerPoint.Application") as app:
        presentation = app.Presentations.Add()
        try:
            slide = presentation.Slides.Add(1, 12)
            title_box = slide.Shapes.AddTextbox(1, 40, 40, 500, 60)
            title_box.TextFrame.TextRange.Text = "Legacy Roadmap"
            body_box = slide.Shapes.AddTextbox(1, 40, 120, 500, 80)
            body_box.TextFrame.TextRange.Text = "Cutover weekend and rollback plan"
            presentation.BuiltInDocumentProperties("Title").Value = "Legacy Deck"
            presentation.SaveAs(str(path), 1)
        finally:
            presentation.Close()


class TestOfficeArchive(TestCase):
    def setUp(self) -> None:
        self.temp_dir = Path(tempfile.mkdtemp(prefix="office-archive-"))
        self.root = self.temp_dir / "docs"
        self.root.mkdir()
        make_docx(self.root / "memo.docx")
        make_xlsx(self.root / "pipeline.xlsx")
        make_pptx(self.root / "deck.pptx")
        make_pdf(self.root / "invoice.pdf")

    def tearDown(self) -> None:
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_extract_file_handles_docx(self) -> None:
        entry = extract_file(self.root / "memo.docx", root=self.root, max_chars=5000)
        self.assertEqual(entry.status, "ok")
        self.assertEqual(entry.title, "Revenue Memo")
        self.assertIn("Quarterly revenue summary", entry.content)

    def test_archive_directory_writes_outputs_and_reuses_rows(self) -> None:
        state_dir = self.root / ".office-archive"
        entries, stats = archive_directory(
            root=self.root,
            state_dir=state_dir,
            max_chars=5000,
            skip_dirs={".office-archive"},
            force=False,
            write_outputs=True,
            allow_legacy_com=False,
        )
        self.assertEqual(len(entries), 4)
        self.assertEqual(stats["updated"], 4)
        self.assertTrue((state_dir / "archive.jsonl").exists())
        self.assertTrue((state_dir / "archive.md").exists())

        second_entries, second_stats = archive_directory(
            root=self.root,
            state_dir=state_dir,
            max_chars=5000,
            skip_dirs={".office-archive"},
            force=False,
            write_outputs=False,
            allow_legacy_com=False,
        )
        self.assertEqual(len(second_entries), 4)
        self.assertEqual(second_stats["reused"], 4)

    def test_search_index_finds_xlsx_content(self) -> None:
        state_dir = self.root / ".office-archive"
        archive_directory(
            root=self.root,
            state_dir=state_dir,
            max_chars=5000,
            skip_dirs={".office-archive"},
            force=False,
            write_outputs=False,
            allow_legacy_com=False,
        )
        results = search_index(root=self.root, state_dir=state_dir, query="Acme", limit=5)
        self.assertTrue(results)
        self.assertEqual(results[0]["path"], "pipeline.xlsx")

    def test_search_index_finds_cjk_content(self) -> None:
        make_docx_cjk(self.root / "contract.docx")
        state_dir = self.root / ".office-archive"
        archive_directory(
            root=self.root,
            state_dir=state_dir,
            max_chars=5000,
            skip_dirs={".office-archive"},
            force=False,
            write_outputs=False,
            allow_legacy_com=False,
        )
        results = search_index(root=self.root, state_dir=state_dir, query="合同编号", limit=5)
        self.assertTrue(results)
        self.assertEqual(results[0]["path"], "contract.docx")

    def test_archive_jsonl_contains_serializable_entries(self) -> None:
        state_dir = self.root / ".office-archive"
        archive_directory(
            root=self.root,
            state_dir=state_dir,
            max_chars=5000,
            skip_dirs={".office-archive"},
            force=False,
            write_outputs=True,
            allow_legacy_com=False,
        )
        rows = (state_dir / "archive.jsonl").read_text(encoding="utf-8").strip().splitlines()
        payload = [json.loads(row) for row in rows]
        self.assertEqual(len(payload), 4)
        self.assertTrue(any(item["file_type"] == "pdf" for item in payload))

    def test_build_capabilities_report_has_expected_keys(self) -> None:
        report = build_capabilities_report()
        self.assertIn("python", report)
        self.assertIn("platform", report)
        self.assertEqual(report["modernFormats"]["docx"], True)
        self.assertIn("doc", report["legacyFormats"])
        self.assertIn("pdftotext", report["helpers"])


@skipUnless(HAS_OFFICE_COM, "Office COM automation unavailable")
class TestLegacyOfficeArchive(TestCase):
    def setUp(self) -> None:
        self.temp_dir = Path(tempfile.mkdtemp(prefix="office-archive-legacy-"))
        self.root = self.temp_dir / "docs"
        self.root.mkdir()
        try:
            make_doc(self.root / "legacy.doc")
        except pywintypes.com_error as exc:
            raise SkipTest(f"Word COM automation unavailable for .doc smoke test: {exc}") from exc
        try:
            make_xls(self.root / "legacy.xls")
        except pywintypes.com_error as exc:
            raise SkipTest(f"Excel COM automation unavailable for .xls smoke test: {exc}") from exc
        try:
            make_ppt(self.root / "legacy.ppt")
        except pywintypes.com_error as exc:
            raise SkipTest(f"PowerPoint COM automation unavailable for .ppt smoke test: {exc}") from exc

    def tearDown(self) -> None:
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_archive_directory_handles_legacy_binary_formats(self) -> None:
        state_dir = self.root / ".office-archive"
        entries, stats = archive_directory(
            root=self.root,
            state_dir=state_dir,
            max_chars=5000,
            skip_dirs={".office-archive"},
            force=False,
            write_outputs=True,
            allow_legacy_com=True,
        )
        self.assertEqual(len(entries), 3)
        self.assertEqual(stats["errors"], 0)

        by_path = {entry.path: entry for entry in entries}
        self.assertIn("legacy.doc", by_path)
        self.assertIn("legacy.xls", by_path)
        self.assertIn("legacy.ppt", by_path)
        self.assertIn("Legacy contract terms", by_path["legacy.doc"].content)
        self.assertIn("OldCo", by_path["legacy.xls"].content)
        self.assertIn("Legacy Roadmap", by_path["legacy.ppt"].content)


if __name__ == "__main__":
    main()
