# PST-TO-knowledge-for-AI

Your organization already has a knowledge base. It is just trapped inside years of email.

**PST-TO-knowledge-for-AI turns an Outlook PST archive into a clean, searchable, NotebookLM-ready knowledge pack.** It takes years of conversations, decisions, answers, recurring questions, supplier discussions, legal clarifications, pension calculations, and operational know-how, then reshapes them into organized Word documents and structured data that can become the foundation for a real internal knowledge bot.

For non-technical teams, the promise is simple: export the mailbox, run the pipeline, upload the Word files to NotebookLM, and start asking questions against the actual memory of the organization.

## Why this matters

Most important knowledge never reaches a formal knowledge base. It lives in email:

- answers from experienced employees
- decisions buried inside long threads
- explanations sent once and forgotten
- recurring questions from citizens, retirees, customers, or employees
- supplier, contract, legal, finance, and operational correspondence
- professional details that disappear when people leave or roles change

This project helps recover that knowledge and make it usable again.

## What you get

The tool produces two layers of output:

1. **Structured dataset outputs** for analytics, audit, search, deduplication, and future AI ingestion.
2. **NotebookLM Word knowledge pack** with topic files, people index, topic map, review queue, and source documents ready for upload.

The Word pack is designed so NotebookLM can cite and navigate the uploaded source documents more usefully than one huge text dump.

## Main capabilities

- Extracts Outlook PST content through external local tools, especially `readpst`
- Parses EML messages without opening Outlook
- Cleans messy email bodies in Hebrew and English
- Removes repeated reply chains, signatures, auto-replies, delivery failures, and system noise
- Preserves sender, recipients, dates, folders, subjects, attachment metadata, and thread signals
- Marks duplicate messages instead of silently deleting evidence
- Builds topic and people views for two natural search paths: **by subject/topic** and **by person**
- Creates NotebookLM-ready `.docx` source files split by knowledge topic
- Creates CSV/JSONL/SQLite-friendly outputs for future search, dashboards, or vector databases

## End-to-end workflow

The full path is:

1. Export a PST file from Outlook.
2. Make sure `readpst` is available, usually through Cygwin64 Terminal on Windows.
3. Run the Python pipeline.
4. Build the NotebookLM Word pack.
5. Upload the Word files to NotebookLM.
6. Ask NotebookLM questions over years of real organizational email.

## Step 1: Export a PST from Outlook

In Outlook Desktop:

1. Open `File`.
2. Choose `Open & Export`.
3. Choose `Import/Export`.
4. Select `Export to a file`.
5. Choose `Outlook Data File (.pst)`.
6. Select the mailbox or folder you want to export.
7. Save the `.pst` file locally.

Example:

```text
C:\pst_ai_project\pst_files\mailbox.pst
```

## Step 2: Install readpst with Cygwin64 Terminal

This project does **not** open Outlook and does **not** require Outlook automation. The default extraction route is an external local extractor called `readpst`.

On Windows, one practical route is:

1. Install Cygwin64 from the official Cygwin installer.
2. During package selection, include the package that provides `readpst` / `libpst`.
3. Open **Cygwin64 Terminal**.
4. Confirm that `readpst` is available:

```bash
readpst --help
```

If you want to run Python from PowerShell and point directly to the Cygwin binary, use a path like:

```powershell
--readpst-command "C:\cygwin64\bin\readpst.exe"
```

Cygwin64 Terminal is an external runtime and is **not included** in this repository.

## Step 3: Install Python dependencies

Use Python 3.11+.

```powershell
pip install -r requirements.txt
```

For a lightweight first run, you can use the built-in TF-IDF fallback and avoid downloading the multilingual embedding model.

## Best quick start: one command to build the NotebookLM Word pack

```powershell
python run_notebooklm_pipeline.py `
  --pst-path "C:\pst_ai_project\pst_files\mailbox.pst" `
  --work-dir ".\data" `
  --readpst-command "C:\cygwin64\bin\readpst.exe" `
  --embedding-backend tfidf `
  --log-level INFO
```

For a small test run:

```powershell
python run_notebooklm_pipeline.py `
  --pst-path "C:\pst_ai_project\pst_files\mailbox.pst" `
  --work-dir ".\data" `
  --readpst-command "C:\cygwin64\bin\readpst.exe" `
  --embedding-backend tfidf `
  --max-emails 200
```

The output will be created under:

```text
data\output\notebooklm_pack_word
```

## Advanced workflow: run each stage separately

### 1. Extract PST to raw CSV

```powershell
python extract.py `
  --pst-path "C:\pst_ai_project\pst_files\mailbox.pst" `
  --output-csv ".\data\intermediate\emails_raw.csv" `
  --extractor readpst `
  --readpst-command "C:\cygwin64\bin\readpst.exe"
```

### 2. Clean and cluster

```powershell
python clean_and_cluster.py `
  --input ".\data\intermediate\emails_raw.csv" `
  --output ".\data\intermediate\emails_clustered.csv" `
  --embedding-backend tfidf
```

For higher-quality multilingual clustering, install `sentence-transformers` and run:

```powershell
python clean_and_cluster.py `
  --input ".\data\intermediate\emails_raw.csv" `
  --output ".\data\intermediate\emails_clustered.csv" `
  --embedding-backend sentence-transformers
```

### 3. Build the NotebookLM Word pack

```powershell
python build_notebooklm_pack.py `
  --input ".\data\intermediate\emails_clustered.csv" `
  --output-dir ".\data\output\notebooklm_pack_word"
```

## NotebookLM upload order

Upload these files first:

1. `00_GUIDE.docx`
2. `00_INDEX.docx`
3. `01_TOPIC_MAP.docx`
4. `02_PEOPLE_MAP.docx`
5. `03_CLASSIFICATION_RULES.docx`

Then upload the relevant `.docx` files from:

```text
notebooklm_pack_word\sources
```

This lets NotebookLM understand the structure before it reads the detailed email evidence.

## Example questions for NotebookLM

- What topics appear most often in this mailbox?
- Who handled pension continuity questions?
- What recurring issues appear in retiree requests?
- Which emails mention calculations, eligibility, appeals, or legal review?
- What did a specific sender ask about over time?
- Which suppliers or contracts appear repeatedly?
- What can be turned into an FAQ?

## Structured dataset outputs

The full dataset pipeline can also produce:

- `messages.jsonl` - one structured message per line
- `messages.csv` - flattened tabular export
- `attachments.csv` - attachment metadata
- `threads.csv` - grouped thread summaries
- `processing_report.json` - processing counts, errors, and stats
- optional SQLite database

Run it with:

```powershell
python main.py --input .\pst_files --output .\processed_output --recursive --export-sqlite
```

## What is included

- PST extraction adapter around `readpst`
- EML parsing
- Hebrew and English cleaning
- subject and email normalization
- deduplication
- thread reconstruction
- transparent topic rules
- optional LLM tagging hooks
- NotebookLM Word pack generation
- topic and people search utilities
- lightweight tests and GitHub Actions CI

## What is not included

- No PST files
- No private email data
- No generated output files
- No Cygwin64 installer
- No cloud API dependency
- No automatic NotebookLM upload, because NotebookLM upload is currently handled manually through the NotebookLM interface

## How this dataset can later be used for a knowledge base and semantic search

The outputs are designed to support several future paths:

- load `.docx` source files into NotebookLM for immediate question answering
- use `messages.jsonl` or `messages.csv` for analytics and human review
- generate FAQs from recurring topics and high-signal threads
- create embeddings from `body_text_clean` or `clean_body`
- ingest records into a vector database
- build a domain-specific internal assistant that cites message id, sender, date, subject, and source file

## Known limitations

- PST extraction depends on external tools such as `readpst` or optional `pypff`.
- Very large PST files can require significant disk space during extraction.
- Topic classification is transparent and useful, but still heuristic; sensitive archives should be reviewed by humans.
- Attachments are linked and can be exported, but deep attachment text extraction and OCR are future improvements.
- NotebookLM upload itself is manual.

## Project structure

- `main.py` - structured dataset CLI
- `run_notebooklm_pipeline.py` - one-command NotebookLM workflow
- `extract.py` - PST to `emails_raw.csv`
- `clean_and_cluster.py` - cleaning, filtering, clustering
- `build_notebooklm_pack.py` - Word pack builder
- `pst_kb/extractors` - PST extraction adapters
- `pst_kb/parsers` - EML parsing
- `pst_kb/cleaners` - email text cleaning
- `pst_kb/normalizers` - normalization helpers
- `pst_kb/threading` - thread reconstruction
- `pst_kb/deduplication` - duplicate detection
- `pst_kb/exporters` - CSV/JSONL/SQLite export
- `pst_kb/notebooklm` - NotebookLM workflow, topic rules, search views, Word pack builders
- `tests` - lightweight unit and workflow tests

## Bottom line

If your mailbox contains years of answers, decisions, and expertise, this project helps turn it into something people can actually use: a searchable, explainable, AI-ready knowledge base that can become a real knowledge bot.
