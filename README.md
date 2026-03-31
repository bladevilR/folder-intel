# Folder Intel

**English:** A Windows-first OpenClaw skill for document archive, document search, Office file search, PDF search, and folder knowledge indexing.  
**中文：** 一个面向 Windows 的 OpenClaw 技能，用来做文档归档、文件内容搜索、Office 文档检索、PDF 检索和文件夹知识索引。

Turn messy Office and PDF folders into searchable, archived knowledge.

## What it does

**English:** Folder Intel indexes local folders of Word, Excel, PowerPoint, and PDF files, generates archive records for each file, and finds documents by extracted content instead of filename only.  
**中文：** Folder Intel 会索引本地文件夹中的 Word、Excel、PowerPoint 和 PDF 文件，为每个文件生成归档记录，并按提取出的内容而不是仅按文件名来查找文档。

- archive every supported file into a reusable inventory
- search by content, clause, invoice number, topic, or keyword
- keep incremental state in SQLite so repeat runs stay fast
- avoid risky Office automation by default on end-user machines

## Search keywords

If users search GitHub or the web for these terms, this repo should be relevant:

- OpenClaw skill
- document archive skill
- document search skill
- Office document search
- PDF folder search
- Windows document indexing
- local document archive
- searchable knowledge folder

## Supported formats

- `doc`
- `docx`
- `xls`
- `xlsx`
- `xlsm`
- `ppt`
- `pptx`
- `pdf`

## Output

Running `archive` creates a `.office-archive/` directory inside the target folder:

- `archive.md` for a human-readable inventory
- `archive.jsonl` for structured per-file records
- `index.sqlite` for incremental search

## Install

Recommended shared install location on Windows:

```text
%USERPROFILE%\.openclaw\skills\folder-intel\
```

OpenClaw also supports per-workspace installs:

```text
<workspace>\skills\folder-intel\
```

Keep these files together inside the skill directory:

```text
folder-intel/
  SKILL.md
  agents/openai.yaml
  scripts/
```

## Requirements

Required:

- Python 3

Optional but recommended:

- `antiword` for higher-quality legacy `.doc` extraction
- `pdftotext` for stronger PDF text extraction
- Python packages used by the scripts may include `olefile`, `xlrd`, and `pypdf`

## Verify install

After copying the skill into the target OpenClaw skill directory, run:

```bash
python scripts/office_archive.py check
```

This confirms which formats and helpers are available on the current machine.

## Quick start

Archive a folder:

```bash
python scripts/office_archive.py archive "D:\Docs\Contracts"
```

Search by content:

```bash
python scripts/office_archive.py search "D:\Docs\Contracts" "invoice april"
python scripts/office_archive.py search "D:\Docs\Contracts" "contract number"
```

Inspect one file:

```bash
python scripts/office_archive.py inspect "D:\Docs\Contracts\renewal.docx"
```

## Example workflow

Example target folder:

```text
D:\AI专班
```

Example archive output:

```text
D:\AI专班\.office-archive\
  archive.md
  archive.jsonl
  index.sqlite
```

Example usage pattern:

```bash
python scripts/office_archive.py archive "D:\AI专班"
python scripts/office_archive.py search "D:\AI专班" "工单"
python scripts/office_archive.py search "D:\AI专班" "AI大模型"
```

## Safety defaults

- Microsoft Office COM extraction is disabled by default
- Legacy extraction stays conservative to avoid Office repair popups and interactive prompts
- Re-runs reuse cached extraction results when files have not changed

## Known limits

- This is keyword and text extraction search, not vector semantic retrieval
- Legacy binary Office formats are handled conservatively
- Some damaged or non-standard files may only be partially indexed or skipped
- `ppt` legacy extraction is intentionally stricter than modern `pptx`

## Good fit

- contract folders
- report archives
- bid and compliance materials
- policy and process folders
- project delivery folders
- mixed Office and PDF knowledge folders

## Release

Current public version: `v0.1.0`
