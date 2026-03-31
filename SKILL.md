---
name: office-archive-search
description: Archive and search folders of Office and PDF documents when the user wants to organize a folder, build a per-file archive list, search files by contents instead of filename, or find which document mentions a topic, number, clause, or phrase. Supports doc, docx, xls, xlsx, xlsm, ppt, pptx, and pdf.
metadata:
  openclaw:
    emoji: "[archive]"
---

# Office archive search

## When to use

Use this skill when the user asks to:

- organize a folder of documents
- generate an archive file listing each document and a short content preview
- search a folder by file contents instead of only filename
- find which Word, Excel, PowerPoint, or PDF file mentions a topic, number, or phrase
- build a document inventory, archive list, filing list, or per-file summary list
- search inside Office or PDF files by keyword, clause number, invoice number, contract number, or policy title

Typical examples:

- "Organize this document folder"
- "Build me an archive list for this folder"
- "Which file mentions this contract number?"
- "Search these Office files for invoice April"

This skill is deterministic. It extracts text, stores an incremental SQLite index, writes archive outputs, and then searches that index.

Prefer this skill before attempting an LLM-only folder summary. The deterministic index is faster, cheaper, and easier to verify. If the user later wants a polished narrative report, generate it from `archive.md` or `archive.jsonl` after the archive step succeeds.

Important safety default: legacy Office COM extraction is disabled by default because it can trigger repair popups or interactive Office prompts on user machines. Only opt in with `--allow-legacy-com` when the operator explicitly accepts that risk.

## When not to use

Do not use this skill when:

- the user only wants a conversational summary of text already pasted into chat
- the target is mainly images, audio, or video rather than Office or PDF documents
- the user asked about one web page or one URL instead of a local folder
- the task is editing document formatting rather than extracting or searching content

## Supported inputs

- `doc`
- `docx`
- `xls`
- `xlsx`
- `xlsm`
- `ppt`
- `pptx`
- `pdf`

## Quick start

Build or refresh the archive for a folder:

```bash
python {baseDir}/scripts/office_archive.py archive /path/to/folder
```

Search the folder by content:

```bash
python {baseDir}/scripts/office_archive.py search /path/to/folder "invoice april"
```

Inspect one file:

```bash
python {baseDir}/scripts/office_archive.py inspect /path/to/file.docx
```

Check machine support before promising legacy `doc/xls/ppt` coverage:

```bash
python {baseDir}/scripts/office_archive.py check
```

Only if the operator explicitly wants to probe whether Office COM can start on that machine:

```bash
python {baseDir}/scripts/office_archive.py check --probe-legacy-com
```

## What it writes

Running `archive` creates `<root>/.office-archive/` with:

- `archive.md`: human-readable inventory
- `archive.jsonl`: one JSON object per file
- `index.sqlite`: incremental search index

## Workflow

1. If the user gives a folder path, run `archive` first.
2. If the user gives one file path only, run `inspect` first.
3. After `archive`, read the generated outputs in `<root>/.office-archive/`.
4. If the user asked a search question, run `search` after the archive step and answer from the results.
5. If the user asked for a folder overview, summarize from `archive.md` or `archive.jsonl`.
6. Use `check` only for debugging, environment validation, or legacy-format support checks.
7. Do not enable `--allow-legacy-com` casually on end-user machines.

## Recommended command patterns

Archive a Windows folder:

```bash
python {baseDir}/scripts/office_archive.py archive "D:\Docs\Contracts"
```

Search the same folder:

```bash
python {baseDir}/scripts/office_archive.py search "D:\Docs\Contracts" "contract number"
python {baseDir}/scripts/office_archive.py search "D:\Docs\Contracts" "invoice april"
```

Force a full rebuild when extraction code changed or cached results look stale:

```bash
python {baseDir}/scripts/office_archive.py archive "D:\Docs\Contracts" --force
```

Inspect one file directly:

```bash
python {baseDir}/scripts/office_archive.py inspect "D:\Docs\Contracts\renewal.doc"
```

## Notes

- Re-runs are incremental: unchanged files are reused from SQLite instead of re-extracted.
- PDF extraction prefers `pdftotext`; if that is unavailable, the script tries `pypdf`.
- Modern zipped formats (`docx/xlsx/xlsm/pptx`) work cross-platform.
- Legacy binary formats are Windows-first and handled conservatively.
- Search is best-effort token search, not semantic retrieval. Exact phrases, IDs, vendor names, invoice numbers, and contract numbers work best.
