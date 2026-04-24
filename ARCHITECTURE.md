# Architecture

This project has two related workflows.

## 1. Structured Dataset Workflow

Use this when the goal is analytics, audit, future vector search, or a normalized database.

```text
PST folder
  -> ReadpstExtractor
  -> EmailParser
  -> MessageProcessor
  -> Deduplicator
  -> ThreadBuilder
  -> DatasetExporter
  -> messages.jsonl / messages.csv / attachments.csv / threads.csv / report / sqlite
```

Entry point:

```powershell
python main.py --input .\pst_files --output .\processed_output --recursive
```

## 2. NotebookLM Workflow

Use this when the goal is to turn one PST archive into a set of Word files that can be uploaded to NotebookLM.

```text
PST file
  -> extract.py
  -> emails_raw.csv
  -> clean_and_cluster.py
  -> emails_clustered.csv
  -> build_notebooklm_pack.py
  -> notebooklm_pack_word/*.docx
```

One-command entry point:

```powershell
python run_notebooklm_pipeline.py --pst-path .\pst_files\mailbox.pst --work-dir .\data
```

## Design Principles

- Outlook is not automated or opened by default.
- PST extraction is isolated behind adapter code.
- Raw email data is preserved before cleaning.
- Cleaning is intentionally conservative, especially for Hebrew professional terminology.
- Duplicate messages are marked, not silently removed.
- Topic classification is transparent and reviewable.
- NotebookLM sources are split into Word documents by knowledge topic and size.
- CSV exports use `utf-8-sig` for better Windows compatibility.

## Important Modules

- `pst_kb/extractors/readpst.py` - adapter around the external `readpst` utility.
- `pst_kb/parsers/eml_parser.py` - parses EML files into raw message objects.
- `pst_kb/processor.py` - converts raw messages into normalized records.
- `pst_kb/cleaners/email_cleaner.py` - removes reply history, signatures, boilerplate, and noise.
- `pst_kb/threading/thread_builder.py` - builds native or synthetic thread keys.
- `pst_kb/deduplication/deduplicator.py` - marks duplicates using identifiers and content hashes.
- `pst_kb/exporters/files.py` - writes JSONL, CSV, reports, and optional SQLite exports.
- `pst_kb/notebooklm/workflow.py` - orchestrates extraction, cleaning, clustering, and Word pack generation.
- `pst_kb/notebooklm/topic_taxonomy.py` - curated knowledge-topic rules.
- `pst_kb/notebooklm/notebook_pack_text.py` - builds the Word-based NotebookLM pack.

## Extension Points

The system is intentionally adapter-friendly:

- Add a new PST extractor by implementing the extractor interface.
- Add richer attachment text extraction after message parsing.
- Add a stronger topic taxonomy without changing extraction code.
- Add local embeddings or vector database ingestion after `emails_clustered.csv`.
- Add a thread-aware clustering stage before NotebookLM export.

## Operational Notes

- `readpst` is a system dependency and should be installed once outside the repository.
- For Windows pilots, Cygwin64 Terminal is a practical way to obtain `readpst`.
- For large PST files, start with `--max-emails 200` to validate the workflow.
- Use `--embedding-backend tfidf` for a fast local run.
- Use `--embedding-backend sentence-transformers` for better multilingual clustering when the model is installed.
