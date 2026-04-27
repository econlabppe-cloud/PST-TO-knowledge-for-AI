from __future__ import annotations

from pathlib import Path

import pandas as pd
from docx import Document

from pst_kb.notebooklm.notebook_pack_text import build_notebooklm_pack
from pst_kb.notebooklm.topic_taxonomy import classify_email_record_corpus


def test_curated_rules_classify_business_email() -> None:
    row = {
        "subject": "Pension eligibility calculation for retirement case",
        "body": "Please review the pension eligibility, retirement date, and calculation simulation.",
        "clean_body": "Please review the pension eligibility, retirement date, and calculation simulation.",
        "folder_path": "Inbox/Pension",
        "from_email": "advisor@example.org",
    }

    match = classify_email_record_corpus(row, topic_source="rules")

    assert match.topic
    assert match.score >= 4
    assert match.review_required is False
    assert match.is_system_noise is False


def test_curated_rules_classify_retiree_documents_request() -> None:
    row = {
        "subject": "מסמכים ואישור שנתי",
        "body": "אבקש לשלוח אישור שנתי וטופס 106 עבור גמלאי רשות השידור.",
        "clean_body": "אבקש לשלוח אישור שנתי וטופס 106 עבור גמלאי רשות השידור.",
        "folder_path": "Inbox",
        "from_email": "hr2@iba.org.il",
    }

    match = classify_email_record_corpus(row, topic_source="rules")

    assert match.topic == "מסמכים, אישורים וטפסים"
    assert match.review_required is False


def test_curated_rules_detect_system_noise() -> None:
    row = {
        "subject": "Delivery Status Notification",
        "body": "Message could not be delivered to the following recipient.",
        "clean_body": "Message could not be delivered to the following recipient.",
        "folder_path": "Inbox",
        "from_email": "mailer-daemon@example.org",
    }

    match = classify_email_record_corpus(row, topic_source="rules")

    assert match.is_system_noise is True
    assert match.score == 0


def test_word_pack_outputs_docx_sources_and_manifest(tmp_path: Path) -> None:
    input_csv = tmp_path / "emails_clustered.csv"
    output_dir = tmp_path / "pack"
    rows = [
        {
            "message_id": "m1",
            "date": "2024-01-01T10:00:00Z",
            "subject": "Pension eligibility calculation",
            "subject_normalized": "pension eligibility calculation",
            "body": "Please review retirement eligibility and the pension calculation simulation.",
            "clean_body": "Please review retirement eligibility and the pension calculation simulation.",
            "from_email": "advisor@example.org",
            "to_emails": "team@example.org",
            "cc_emails": "",
            "bcc_emails": "",
            "folder_path": "Inbox/Pension",
            "cluster_name": "Pension",
            "llm_topic": "",
            "llm_subtopic": "",
            "llm_tags": "",
            "llm_confidence": "",
            "is_filtered": "False",
            "filter_reason": "",
        },
        {
            "message_id": "m2",
            "date": "2024-01-02T10:00:00Z",
            "subject": "Supplier invoice contract",
            "subject_normalized": "supplier invoice contract",
            "body": "The supplier contract and invoice require approval before payment.",
            "clean_body": "The supplier contract and invoice require approval before payment.",
            "from_email": "procurement@example.org",
            "to_emails": "team@example.org",
            "cc_emails": "",
            "bcc_emails": "",
            "folder_path": "Inbox/Procurement",
            "cluster_name": "Contracts",
            "llm_topic": "",
            "llm_subtopic": "",
            "llm_tags": "",
            "llm_confidence": "",
            "is_filtered": "False",
            "filter_reason": "",
        },
    ]
    pd.DataFrame(rows).to_csv(input_csv, index=False, encoding="utf-8-sig")

    outputs = build_notebooklm_pack(input_csv=input_csv, output_dir=output_dir, max_emails_per_file=1)

    assert outputs["guide"].exists()
    assert outputs["manifest"].exists()
    assert (output_dir / "sources").exists()
    assert list((output_dir / "sources").glob("*.docx"))

    guide_text = "\n".join(paragraph.text for paragraph in Document(str(outputs["guide"])).paragraphs)
    assert "00_GUIDE.docx" in guide_text

    manifest = pd.read_csv(outputs["manifest"], dtype=str, keep_default_na=False, encoding="utf-8-sig")
    assert not manifest.empty
    assert manifest["source_file"].str.endswith(".docx").all()
    assert manifest["source_locator"].str.contains("EMAIL").all()
