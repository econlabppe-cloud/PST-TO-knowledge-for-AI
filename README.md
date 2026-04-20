# PST-TO-knowledge-for-AI

Your organization already has a knowledge base. It is just trapped inside years of email.

PST-TO-knowledge-for-AI turns an Outlook archive into a clean, searchable, NotebookLM-ready knowledge system. It takes the conversations, decisions, answers, attachments, recurring questions, and hidden expertise buried in a mailbox, then reshapes them into organized Word documents and structured data that can become the foundation for an internal knowledge bot.

For non-technical teams, the idea is simple: export the mailbox, run the pipeline, upload the Word files to NotebookLM, and start asking questions against years of real organizational memory.

## The promise

Most important knowledge never reaches a formal knowledge base.
It lives in emails: answers from experienced employees, decisions made in long threads, explanations sent once and forgotten, supplier discussions, legal clarifications, pension calculations, user requests, and operational know-how.

This project helps recover that knowledge and make it usable again.

With it, a team can:

- turn an old Outlook archive into a searchable knowledge source
- ask NotebookLM questions based on real historical correspondence
- search by topic, person, thread, or recurring issue
- discover patterns across years of work
- prepare the ground for FAQ creation, semantic search, and an internal AI assistant
- preserve institutional memory before it disappears inside inboxes

## What this project does

- Extracts email from Outlook PST archives
- Converts extracted messages into structured records
- Cleans messy email bodies in Hebrew and English
- Removes boilerplate, repeated quotes, auto-replies, and system noise
- Preserves metadata such as sender, recipients, date, folder, subject, and thread signals
- Marks duplicate messages instead of deleting evidence
- Groups content into knowledge-oriented topics
- Builds NotebookLM-ready Word documents
- Creates topic maps and people maps so users know what to upload and where to look

## End-to-end workflow

The full path is:

1. Export a PST file from Outlook.
2. Use Cygwin64 Terminal with `readpst` to extract the mailbox into message files.
3. Run the Python pipeline on the extracted material.
4. Clean, classify, and organize the archive.
5. Build a Word-based NotebookLM pack.
6. Upload the Word files to NotebookLM.
7. Use NotebookLM as a knowledge bot over the email archive.

## Step 1: Export a PST from Outlook

In Outlook Desktop:

1. Open `File`.
2. Choose `Open & Export`.
3. Choose `Import/Export`.
4. Select `Export to a file`.
5. Choose `Outlook Data File (.pst)`.
6. Select the mailbox or folder you want to export.
7. Save the `.pst` file locally.

Example target folder:

```text
C:\pst_ai_project\pst_files\mailbox.pst
```

## Step 2: Extract with Cygwin64 Terminal

This project uses the external `readpst` tool for PST extraction.
Cygwin64 Terminal can be used as the Windows environment for running `readpst`.

Important: Cygwin64 Terminal itself is not included in this repository.
It is an external runtime environment, not project code.

Typical extraction command:

```bash
readpst -e -8 -o ./staging ./pst_files/mailbox.pst
```

What the flags mean:

- `-e` exports messages as individual email files
- `-8` keeps text in UTF-8
- `-o ./staging` writes extracted files into the staging folder

After this step, the PST archive is no longer a black box.
It becomes a folder of message files that Python can parse, clean, classify, and export.

## Step 3: Install Python dependencies

```powershell
pip install -r requirements.txt
```

Requirements:

- Python 3.11+
- `readpst` available through Cygwin64 Terminal or PATH
- `python-docx` for Word-based NotebookLM exports

## Step 4: Run the local PST pipeline

Run the full local processing flow:

```powershell
python main.py --input .\pst_files --output .\processed_output --recursive
```

Useful options:

- `--single-file path\to\file.pst`
- `--skip-attachments`
- `--export-sqlite`
- `--folder-path "Inbox\\Folder"`
- `--limit 1000`

The pipeline creates structured outputs such as:

- message records
- attachment metadata
- thread summaries
- duplicate markers
- processing reports
- optional SQLite tables

## Step 5: Clean, cluster, and prepare the knowledge layer

For the NotebookLM-oriented workflow:

```powershell
python clean_and_cluster.py --input-csv .\data\intermediate\emails_raw.csv --output-csv .\data\intermediate\emails_clustered.csv --embedding-backend tfidf
```

This step removes low-value noise, normalizes the text, and prepares the archive for topic-based browsing.

## Step 6: Build the NotebookLM Word pack

```powershell
python build_notebooklm_pack.py --input .\data\intermediate\emails_clustered.csv --output-dir .\data\output\notebooklm_pack_word
```

This creates Word documents designed for NotebookLM ingestion:

- `00_INDEX.docx`
- `01_TOPIC_MAP.docx`
- `02_PEOPLE_MAP.docx`
- `03_CLASSIFICATION_RULES.docx`
- `04_REVIEW_QUEUE.docx`
- `sources\*.docx`

## Step 7: Upload to NotebookLM

Recommended upload order:

1. `00_INDEX.docx`
2. `01_TOPIC_MAP.docx`
3. `02_PEOPLE_MAP.docx`
4. `03_CLASSIFICATION_RULES.docx`
5. The relevant files from `sources\`

After upload, NotebookLM can answer questions based on the email archive.
The Word format helps NotebookLM refer back to more specific places inside the source documents.

Example questions:

- What topics appear most often in this mailbox?
- Who handled pension continuity questions?
- What decisions were made about a specific supplier?
- What recurring issues appear in retiree requests?
- Which emails mention calculations, eligibility, or legal review?

## Why this becomes a knowledge bot

A normal mailbox answers nothing unless a human remembers what to search for.
This project changes that.

It turns raw correspondence into curated evidence:

- topics become visible
- people become navigable
- threads become readable
- repeated questions become reusable knowledge
- NotebookLM can speak from organized source documents

That is the difference between storing email and activating knowledge.

## What is included

- PST extraction adapter around `readpst`
- EML parsing
- Hebrew and English cleaning
- normalization utilities
- deduplication
- thread reconstruction
- classification heuristics
- structured exporters
- NotebookLM Word pack generation
- topic and people search views

## What is not included

- No PST files
- No private email data
- No generated output files
- No caches
- No virtual environment
- No Cygwin64 Terminal installer

## Main folders

- `pst_kb/extractors` - PST extraction adapters
- `pst_kb/parsers` - EML parsing
- `pst_kb/cleaners` - email text cleaning
- `pst_kb/normalizers` - normalization helpers
- `pst_kb/threading` - thread reconstruction
- `pst_kb/deduplication` - duplicate detection
- `pst_kb/exporters` - file and SQLite export
- `pst_kb/notebooklm` - NotebookLM pack builders and search views

## Who this is for

This is useful for teams with years of operational knowledge buried in Outlook:

- public sector teams
- legal and compliance teams
- pension and benefits teams
- finance and procurement teams
- customer service and case management teams
- any organization where the real knowledge lives in email

## Bottom line

If your mailbox contains years of answers, decisions, and expertise, this project helps turn it into something people can actually use: a searchable, explainable, AI-ready knowledge base.
