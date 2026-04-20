# PST Analysis Toolkit

Local Python toolkit for extracting, cleaning, normalizing, threading, deduplicating, classifying, and exporting Outlook PST email archives.

## What is included

- PST extraction through a pluggable extractor interface
- EML parsing
- body cleaning and normalization
- thread grouping
- duplicate detection
- metadata-rich exports
- NotebookLM-oriented document packs
- topic and people search views

## What is not included

- No PST data files
- No generated outputs
- No caches
- No virtual environment

## Requirements

- Python 3.11+
- `readpst` on PATH for the default extraction path

## Install

```powershell
pip install -r requirements.txt
```

If you want the Word export path for NotebookLM packs:

```powershell
pip install python-docx
```

## Run

```powershell
python main.py --input .\pst_files --output .\processed_output --recursive
```

Useful flags:

- `--single-file path\to\file.pst`
- `--skip-attachments`
- `--export-sqlite`
- `--folder-path "Inbox\\Folder"`
- `--limit 1000`

## Main folders

- `pst_kb/extractors` - PST extraction adapters
- `pst_kb/parsers` - EML parsing
- `pst_kb/cleaners` - email text cleaning
- `pst_kb/normalizers` - normalization helpers
- `pst_kb/threading` - thread reconstruction
- `pst_kb/deduplication` - duplicate detection
- `pst_kb/exporters` - file and SQLite export
- `pst_kb/notebooklm` - downstream NotebookLM pack builders and search views

## GitHub upload

This folder is already trimmed for repository upload. Add it as a repo root or copy its contents into a new repository and commit normally.
