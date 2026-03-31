# Folder Intel

Windows-first OpenClaw skill for archiving and searching Office and PDF folders by content.

## What it does

- Builds a per-file archive for document folders
- Searches files by extracted contents instead of filename only
- Supports `doc`, `docx`, `xls`, `xlsx`, `xlsm`, `ppt`, `pptx`, and `pdf`
- Writes reusable outputs to `.office-archive/`

## Output files

Running `archive` creates:

- `archive.md` for a human-readable inventory
- `archive.jsonl` for structured per-file records
- `index.sqlite` for incremental search

## Quick start

Install this skill into an OpenClaw workspace or shared skills directory, then use the commands below.

Recommended shared install location on Windows:

```text
%USERPROFILE%\.openclaw\skills\folder-intel\
```

Keep `SKILL.md` and `scripts/` together in the same skill directory.

### Commands

Archive a folder:

```bash
python scripts/office_archive.py archive "D:\Docs\Contracts"
```

Search by content:

```bash
python scripts/office_archive.py search "D:\Docs\Contracts" "invoice april"
```

Inspect one file:

```bash
python scripts/office_archive.py inspect "D:\Docs\Contracts\renewal.docx"
```

Check machine support:

```bash
python scripts/office_archive.py check
```

## Notes

- Default behavior avoids Microsoft Office COM for safety on end-user machines.
- Legacy Office extraction is conservative to avoid repair popups and interactive prompts.
- Re-runs are incremental and reuse cached extraction results when files have not changed.

## Version

Initial public release: `v0.1.0`
