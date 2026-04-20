from __future__ import annotations

import logging
from pathlib import Path

from pst_kb.classifiers import HeuristicClassifier
from pst_kb.cleaners import EmailCleaner
from pst_kb.config import AppConfig
from pst_kb.language import detect_language
from pst_kb.models import AttachmentRecord, MessageRecord, RawMessage, Recipient
from pst_kb.normalizers import detect_reply_forward_indicator, normalize_email, normalize_subject
from pst_kb.utils.files import sanitize_filename
from pst_kb.utils.hashing import sha256_bytes, sha256_text, stable_hash

logger = logging.getLogger(__name__)


class MessageProcessor:
    """
    Coordinates the processing of a raw message into a structured record.
    Handles cleaning, normalization, language detection, and classification.
    """

    def __init__(self, config: AppConfig, output_dir: Path) -> None:
        self.config = config
        self.output_dir = output_dir
        self.cleaner = EmailCleaner()
        self.classifier = HeuristicClassifier(
            internal_domains=config.internal_domains,
            topic_keywords=config.topic_keywords,
            sender_type_keywords=config.sender_type_keywords,
        )

    def process(self, raw: RawMessage) -> tuple[MessageRecord, list[AttachmentRecord]]:
        """
        Process a RawMessage into a MessageRecord and a list of AttachmentRecords.
        """
        clean_result = self.cleaner.clean(raw.body_text_raw, raw.body_html_raw)
        subject_normalized = normalize_subject(raw.subject_raw)
        body_for_hash = clean_result.text or raw.body_text_raw or self.cleaner.html_to_text(raw.body_html_raw)
        content_hash = sha256_text("\n".join([subject_normalized, body_for_hash.strip()]).lower())

        record_id = stable_hash(
            [
                raw.source_pst,
                raw.source_folder,
                raw.message_id,
                raw.sent_at,
                raw.sender_email,
                subject_normalized,
                content_hash,
            ]
        )[:24]

        attachment_records = self._build_attachment_records(record_id, raw)
        classification = self.classifier.classify(subject_normalized, clean_result.text, raw.sender_email)

        message = MessageRecord(
            record_id=record_id,
            source_pst=raw.source_pst,
            source_folder=raw.source_folder,
            message_id=raw.message_id,
            conversation_id=raw.conversation_id,
            in_reply_to=raw.in_reply_to,
            references=raw.references,
            subject_raw=raw.subject_raw,
            subject_normalized=subject_normalized,
            sender_name=raw.sender_name,
            sender_email=normalize_email(raw.sender_email),
            to_emails=_recipient_emails(raw.to),
            cc_emails=_recipient_emails(raw.cc),
            bcc_emails=_recipient_emails(raw.bcc),
            sent_at=raw.sent_at,
            received_at=raw.received_at,
            body_text_raw=raw.body_text_raw,
            body_html_raw=raw.body_html_raw,
            body_text_clean=clean_result.text,
            language=detect_language(clean_result.text or body_for_hash),
            importance=raw.importance,
            flags=raw.flags,
            categories=raw.categories,
            has_attachments=bool(attachment_records),
            attachment_count=len(attachment_records),
            attachment_ids=[record.attachment_id for record in attachment_records],
            internet_headers=raw.internet_headers,
            reply_forward_indicator=detect_reply_forward_indicator(raw.subject_raw),
            parent_message_id=raw.in_reply_to or (raw.references[-1] if raw.references else None),
            mostly_quoted=clean_result.mostly_quoted,
            content_hash=content_hash,
            possible_intent=classification.possible_intent,
            possible_topic=classification.possible_topic,
            complaint_indicator=classification.complaint_indicator,
            request_indicator=classification.request_indicator,
            urgency_indicator=classification.urgency_indicator,
            has_question=classification.has_question,
            likely_sender_type=classification.likely_sender_type,
            extraction_errors=raw.extraction_errors,
            processing_notes=clean_result.notes,
        )
        return message, attachment_records

    def _build_attachment_records(self, record_id: str, raw: RawMessage) -> list[AttachmentRecord]:
        """
        Extract and save attachments from a raw message.
        """
        records: list[AttachmentRecord] = []
        for index, payload_info in enumerate(raw.attachment_payloads, start=1):
            original_filename = str(payload_info.get("filename") or f"attachment_{index}")
            payload = payload_info.get("payload")
            payload_bytes = payload if isinstance(payload, bytes) else b""
            payload_hash = sha256_bytes(payload_bytes)[:12]
            safe_original = sanitize_filename(original_filename, fallback=f"attachment_{index}")
            saved_filename = f"{index:03d}_{payload_hash}_{safe_original}"
            attachment_id = stable_hash([record_id, index, original_filename, payload_hash])[:24]

            saved_path: str | None = None
            export_error: str | None = None
            if self.config.skip_attachments:
                export_error = "binary_export_skipped"
            else:
                try:
                    target_dir = (
                        self.output_dir
                        / "attachments"
                        / sanitize_filename(Path(raw.source_pst).stem, fallback="pst")
                        / record_id
                    )
                    target_dir.mkdir(parents=True, exist_ok=True)
                    target_path = target_dir / saved_filename
                    target_path.write_bytes(payload_bytes)
                    saved_path = str(target_path)
                except Exception as exc:
                    logger.warning("Failed to export attachment %s: %s", original_filename, exc)
                    export_error = str(exc)

            records.append(
                AttachmentRecord(
                    attachment_id=attachment_id,
                    record_id=record_id,
                    source_pst=raw.source_pst,
                    original_filename=original_filename,
                    saved_filename=None if self.config.skip_attachments else saved_filename,
                    saved_path=saved_path,
                    extension=Path(original_filename).suffix.lower() or None,
                    size=len(payload_bytes),
                    content_type=str(payload_info.get("content_type") or "") or None,
                    content_id=str(payload_info.get("content_id") or "") or None,
                    is_inline=bool(payload_info.get("is_inline")),
                    export_error=export_error,
                )
            )
        return records


def _recipient_emails(recipients: list[Recipient]) -> list[str]:
    """
    Extract normalized email addresses from a list of Recipient objects.
    """
    emails: list[str] = []
    for recipient in recipients:
        email = normalize_email(getattr(recipient, "email", None))
        if email:
            emails.append(email)
    return emails
