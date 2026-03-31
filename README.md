# Folder Intel

**English:** A Windows-first OpenClaw skill for document archive, document search, Office file search, PDF search, and folder knowledge indexing.  
**中文：** 一个面向 Windows 的 OpenClaw 技能，用来做文档归档、文件内容搜索、Office 文档检索、PDF 检索和文件夹知识索引。

Turn messy Office and PDF folders into searchable, archived knowledge.

## What is Folder Intel

**English:** Folder Intel is an OpenClaw skill for indexing local folders of Word, Excel, PowerPoint, and PDF files, generating archive records for each file, and finding documents by extracted content instead of filename only.  
**中文：** Folder Intel 是一个 OpenClaw skill，用于索引本地文件夹中的 Word、Excel、PowerPoint 和 PDF 文件，为每个文件生成归档记录，并按提取出的内容而不是仅按文件名来查找文档。

## Why it exists

**English:** Teams often have folders full of contracts, reports, policies, slides, and spreadsheets with weak naming and no searchable archive. Folder Intel gives that folder a deterministic local index.  
**中文：** 很多团队都有一堆命名混乱、不可检索的合同、报告、制度、PPT 和表格。Folder Intel 给这类文件夹建立一个可重复、可本地运行的确定性索引。

- archive every supported file into a reusable inventory
- search by content, clause, invoice number, topic, or keyword
- keep incremental state in SQLite so repeat runs stay fast
- avoid risky Office automation by default on end-user machines

## Search keywords

If users search GitHub or the web for any of these terms, this repo should be relevant:

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

## What it writes

Running `archive` creates a `.office-archive/` directory inside the target folder:

- `archive.md` for a human-readable inventory
- `archive.jsonl` for structured per-file records
- `index.sqlite` for incremental search

## Install

Recommended shared install location on Windows:

```text
%USERPROFILE%\.openclaw\skills\folder-intel\
```

Keep these files together inside the skill directory:

```text
folder-intel/
  SKILL.md
  agents/openai.yaml
  scripts/
```

OpenClaw also supports per-workspace installs:

```text
<workspace>\skills\folder-intel\
```

## Requirements

- Python 3
- Optional but recommended: `antiword` for higher-quality legacy `.doc` extraction
- Optional but recommended: `pdftotext` for stronger PDF text extraction
- Python packages used by the scripts may include `olefile`, `xlrd`, and `pypdf`

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

Check machine support:

```bash
python scripts/office_archive.py check
```

## Safety defaults

- Microsoft Office COM extraction is disabled by default
- Legacy extraction stays conservative to avoid Office repair popups and interactive prompts
- Re-runs reuse cached extraction results when files have not changed

## Release

Current public version: `v0.1.0`
