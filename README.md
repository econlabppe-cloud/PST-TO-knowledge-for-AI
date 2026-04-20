# PST-TO-knowledge-for-AI

Turn a dusty Outlook PST archive into a living knowledge asset.
This toolkit takes thousands of emails that usually sit buried in a mailbox, extracts them cleanly, organizes them into meaningful topics, and turns the result into a searchable knowledge base that can power NotebookLM, semantic search, FAQ generation, and human review.

Instead of treating email as storage, it turns email into memory.
Instead of a mailbox full of noise, it gives you a structured corpus people can actually use.

## What this project does

- Extracts emails from PST files with a pluggable pipeline
- Cleans and normalizes bodies in Hebrew and English
- Reconstructs threads and keeps metadata
- Detects duplicates and system noise
- Classifies content into knowledge-oriented topics
- Exports structured data for analytics, search, and review
- Builds Word-based NotebookLM packs from the archive
- Creates a path from raw mail to a real knowledge bot

## What is not included

- No PST files
- No generated outputs
- No caches
- No virtual environment
- No Cygwin64 Terminal bundle

## Requirements

- Python 3.11+
- `readpst` available in the Cygwin64 Terminal environment or in PATH
- `python-docx` for Word-based NotebookLM packs

## Step-by-step workflow

### 1. Export the PST with Cygwin64 Terminal

The external extraction step uses `readpst`.
You can run it from Cygwin64 Terminal as the execution environment.
The repository does not include Cygwin itself, because it is only the launcher environment, not part of the codebase.

Typical flow:

```powershell
readpst -e -8 -o ./staging ./your_mailbox.pst
```

What happens here:

- `-e` exports messages into separate files
- `-8` keeps output in UTF-8
- `-o` chooses the staging folder

### 2. Parse the extracted mail in Python

The toolkit reads the exported message files, extracts metadata, and normalizes each message into a structured record.

### 3. Clean and organize the archive

The pipeline removes boilerplate, quoted chains, duplicate messages, and system noise.
Then it groups the archive by topic, sender, and thread so the material becomes readable instead of chaotic.

### 4. Build a NotebookLM pack

The project creates Word documents that NotebookLM can ingest well.
These documents are shaped for real exploration: topic maps, people maps, review queues, and source packs.

### 5. Use the archive like a knowledge product

At that point the mailbox is no longer just “emails”.
It becomes a knowledge layer you can search, summarize, browse, and build on.

## Install

```powershell
pip install -r requirements.txt
```

## Run the local pipeline

```powershell
python main.py --input .\pst_files --output .\processed_output --recursive
```

Useful flags:

- `--single-file path\to\file.pst`
- `--skip-attachments`
- `--export-sqlite`
- `--folder-path "Inbox\\Folder"`
- `--limit 1000`

## Build the NotebookLM pack

```powershell
python build_notebooklm_pack.py --input .\data\intermediate\emails_clustered.csv --output-dir .\data\output\notebooklm_pack_word
```

Recommended upload order to NotebookLM:

1. `00_INDEX.docx`
2. `01_TOPIC_MAP.docx`
3. `02_PEOPLE_MAP.docx`
4. `03_CLASSIFICATION_RULES.docx`
5. `sources\*.docx`

## Why this is valuable

This project helps teams move from inbox archaeology to actual knowledge work.

- Search by topic instead of hunting through folders
- Search by sender instead of guessing where a thread went
- Recover context from long, messy email chains
- Surface patterns across years of mail
- Turn recurring questions into reusable knowledge
- Create a base for an internal assistant or bot that speaks from the archive

In practice, this means the mailbox starts behaving like a living reference system instead of a dead archive.

## Main folders

- `pst_kb/extractors` - PST extraction adapters
- `pst_kb/parsers` - EML parsing
- `pst_kb/cleaners` - email text cleaning
- `pst_kb/normalizers` - normalization helpers
- `pst_kb/threading` - thread reconstruction
- `pst_kb/deduplication` - duplicate detection
- `pst_kb/exporters` - file and SQLite export
- `pst_kb/notebooklm` - NotebookLM pack builders and search views

## GitHub upload

This folder is already trimmed for repository upload.
Copy it into a GitHub repository, commit it, and push it as the starting point for the project.
