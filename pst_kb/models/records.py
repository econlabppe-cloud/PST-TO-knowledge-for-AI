from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field


class Recipient(BaseModel):
    name: str | None = None
    email: str | None = None


class AttachmentRecord(BaseModel):
    attachment_id: str
    record_id: str
    source_pst: str
    original_filename: str
    saved_filename: str | None = None
    saved_path: str | None = None
    extension: str | None = None
    size: int | None = None
    content_type: str | None = None
    content_id: str | None = None
    is_inline: bool = False
    export_error: str | None = None


class ExtractedMessageRef(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)

    pst_path: Path
    eml_path: Path
    source_folder: str


class RawMessage(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)

    source_pst: str
    source_folder: str
    eml_path: Path
    message_id: str | None = None
    conversation_id: str | None = None
    in_reply_to: str | None = None
    references: list[str] = Field(default_factory=list)
    subject_raw: str | None = None
    sender_name: str | None = None
    sender_email: str | None = None
    to: list[Recipient] = Field(default_factory=list)
    cc: list[Recipient] = Field(default_factory=list)
    bcc: list[Recipient] = Field(default_factory=list)
    sent_at: datetime | None = None
    received_at: datetime | None = None
    body_text_raw: str = ""
    body_html_raw: str = ""
    internet_headers: dict[str, str] = Field(default_factory=dict)
    importance: str | None = None
    flags: list[str] = Field(default_factory=list)
    categories: list[str] = Field(default_factory=list)
    attachment_payloads: list[dict[str, object]] = Field(default_factory=list)
    extraction_errors: list[str] = Field(default_factory=list)
    parent_record_message_id: str | None = None
    nested_depth: int = 0
    container_attachment_filename: str | None = None


class MessageRecord(BaseModel):
    record_id: str
    source_pst: str
    source_folder: str
    message_id: str | None = None
    conversation_id: str | None = None
    in_reply_to: str | None = None
    references: list[str] = Field(default_factory=list)
    thread_key: str | None = None
    thread_position: int | None = None
    subject_raw: str | None = None
    subject_normalized: str = ""
    sender_name: str | None = None
    sender_email: str | None = None
    to_emails: list[str] = Field(default_factory=list)
    cc_emails: list[str] = Field(default_factory=list)
    bcc_emails: list[str] = Field(default_factory=list)
    sent_at: datetime | None = None
    received_at: datetime | None = None
    body_text_raw: str = ""
    body_html_raw: str = ""
    body_text_clean: str = ""
    language: str | None = None
    importance: str | None = None
    flags: list[str] = Field(default_factory=list)
    categories: list[str] = Field(default_factory=list)
    has_attachments: bool = False
    attachment_count: int = 0
    attachment_ids: list[str] = Field(default_factory=list)
    internet_headers: dict[str, str] = Field(default_factory=dict)
    reply_forward_indicator: str | None = None
    parent_message_id: str | None = None
    parent_record_message_id: str | None = None
    nested_depth: int = 0
    container_attachment_filename: str | None = None
    synthetic_thread_key: str | None = None
    mostly_quoted: bool = False
    content_hash: str
    is_duplicate: bool = False
    duplicate_of: str | None = None
    possible_intent: str = "unknown"
    possible_topic: str = "unknown"
    llm_topic: str | None = None
    llm_subtopic: str | None = None
    llm_tags: list[str] = Field(default_factory=list)
    llm_confidence: float | None = None
    llm_summary: str | None = None
    complaint_indicator: bool = False
    request_indicator: bool = False
    urgency_indicator: bool = False
    has_question: bool = False
    likely_sender_type: str = "unknown"
    extraction_errors: list[str] = Field(default_factory=list)
    processing_notes: list[str] = Field(default_factory=list)


class ThreadRecord(BaseModel):
    thread_key: str
    source_psts: list[str] = Field(default_factory=list)
    subject_normalized: str = ""
    message_count: int = 0
    first_message_at: datetime | None = None
    last_message_at: datetime | None = None
    participant_emails: list[str] = Field(default_factory=list)
    record_ids: list[str] = Field(default_factory=list)
    duplicate_count: int = 0
    has_attachments: bool = False


class ProcessingError(BaseModel):
    source: str
    message: str
    detail: str | None = None


class ProcessingReport(BaseModel):
    started_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))
    finished_at: datetime | None = None
    input_files_found: int = 0
    pst_files_processed: int = 0
    messages_extracted: int = 0
    messages_processed: int = 0
    attachments_exported: int = 0
    attachments_skipped: int = 0
    threads_created: int = 0
    duplicates_found: int = 0
    skipped_files: list[str] = Field(default_factory=list)
    errors: list[ProcessingError] = Field(default_factory=list)
    fatal_error: str | None = None
    stats: dict[str, object] = Field(default_factory=dict)
