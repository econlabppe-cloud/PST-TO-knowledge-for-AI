from __future__ import annotations

from pathlib import Path

import pandas as pd

from pst_kb.notebooklm.workflow import run_notebooklm_workflow


def test_workflow_can_reuse_existing_intermediate_files(tmp_path: Path) -> None:
    raw_csv = tmp_path / "intermediate" / "emails_raw.csv"
    clustered_csv = tmp_path / "intermediate" / "emails_clustered.csv"
    pack_dir = tmp_path / "output" / "notebooklm_pack_word"
    raw_csv.parent.mkdir(parents=True)

    row = {
        "message_id": "m1",
        "thread_id": "",
        "subject": "Pension eligibility calculation",
        "subject_normalized": "pension eligibility calculation",
        "body": "Please review retirement eligibility and the pension calculation simulation.",
        "clean_body": "Please review retirement eligibility and the pension calculation simulation.",
        "from_email": "advisor@example.org",
        "to_emails": "team@example.org",
        "cc_emails": "",
        "bcc_emails": "",
        "date": "2024-01-01T10:00:00Z",
        "folder_path": "Inbox/Pension",
        "has_attachments": "False",
        "attachment_count": "0",
        "word_count": "9",
        "source_pst": "sample.pst",
        "extractor": "test",
        "extraction_errors": "",
        "cluster_name": "Pension",
        "llm_topic": "",
        "llm_subtopic": "",
        "llm_tags": "",
        "llm_confidence": "",
        "is_filtered": "False",
        "filter_reason": "",
    }
    pd.DataFrame([row]).to_csv(raw_csv, index=False, encoding="utf-8-sig")
    pd.DataFrame([row]).to_csv(clustered_csv, index=False, encoding="utf-8-sig")

    outputs = run_notebooklm_workflow(
        pst_path=None,
        work_dir=tmp_path,
        raw_csv=raw_csv,
        clustered_csv=clustered_csv,
        pack_dir=pack_dir,
        skip_extract=True,
        skip_clean=True,
    )

    assert outputs["raw_csv"] == raw_csv
    assert outputs["clustered_csv"] == clustered_csv
    assert outputs["pack_dir"] == pack_dir
    assert (pack_dir / "00_GUIDE.docx").exists()
    assert (pack_dir / "records_manifest.csv").exists()
