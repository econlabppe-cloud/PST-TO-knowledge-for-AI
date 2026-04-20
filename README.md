# PST Analysis Toolkit

Local Python toolkit for extracting, cleaning, normalizing, threading, deduplicating, classifying, and exporting Outlook PST email archives into a usable knowledge base.

## What is included

- PST extraction through a pluggable extractor interface
- EML parsing
- body cleaning and normalization
- thread grouping
- duplicate detection
- metadata-rich exports
- NotebookLM-oriented document packs
- topic and people search views
- a path to turn email archives into a searchable knowledge bot

## What is not included

- No PST data files
- No generated outputs
- No caches
- No virtual environment
- No Cygwin64 Terminal bundle

## Requirements

- Python 3.11+
- `readpst` on PATH for the default extraction path
- `python-docx` for Word-based NotebookLM packs

## Extraction note

The repository does not vendor Cygwin64 Terminal or any installer bundle.
The extraction flow uses a small adapter around the external `readpst` command-line tool.
If you used Cygwin locally during development, that was only an execution environment, not a shipped dependency.

## Install

```powershell
pip install -r requirements.txt
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

## NotebookLM workflow

1. Run extraction and normalization on the PST folder.
2. Build the NotebookLM pack in Word format.
3. Upload the generated `.docx` source files into NotebookLM.
4. Use the topic and people maps to navigate by subject or sender.
5. Ask questions against the archive as if it were an organized internal knowledge base.

Example:

```powershell
python build_notebooklm_pack.py --input .\data\intermediate\emails_clustered.csv --output-dir .\data\output\notebooklm_pack_word
```

Recommended upload order:

1. `00_INDEX.docx`
2. `01_TOPIC_MAP.docx`
3. `02_PEOPLE_MAP.docx`
4. `03_CLASSIFICATION_RULES.docx`
5. `sources\*.docx`

## Main folders

- `pst_kb/extractors` - PST extraction adapters
- `pst_kb/parsers` - EML parsing
- `pst_kb/cleaners` - email text cleaning
- `pst_kb/normalizers` - normalization helpers
- `pst_kb/threading` - thread reconstruction
- `pst_kb/deduplication` - duplicate detection
- `pst_kb/exporters` - file and SQLite export
- `pst_kb/notebooklm` - downstream NotebookLM pack builders and search views

## Why this becomes a knowledge bot

This project turns a raw mailbox into structured knowledge instead of a dump of messages.
That means:

- people can search by topic, sender, or question
- the archive becomes easier to review, audit, and explain
- repeated patterns become visible instead of buried in inbox noise
- NotebookLM can answer against curated source documents instead of an unstructured mailbox
- future embeddings, FAQ generation, and semantic search can sit on top of the same cleaned corpus

In short: it turns email from a storage problem into a knowledge asset.

## GitHub upload

This folder is already trimmed for repository upload. Add it as a repo root or copy its contents into a new repository and commit normally.
