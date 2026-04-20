from __future__ import annotations

import argparse
import csv
import logging
import shutil
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from tqdm import tqdm

from pst_kb.cleaners import EmailCleaner
from pst_kb.extractors import ExtractionOptions, ExtractorError, ReadpstExtractor
from pst_kb.models import ExtractedMessageRef, RawMessage, Recipient
from pst_kb.notebooklm.common import configure_script_logging, read_text_resilient, word_count
from pst_kb.normalizers import normalize_whitespace
from pst_kb.parsers import EmailParser
from pst_kb.utils.hashing import stable_hash

logger = logging.getLogger(__name__)

RAW_COLUMNS = [
    "message_id",
    "thread_id",
    "subject",
    "body",
    "from_email",
    "to_emails",
    "cc_emails",
    "bcc_emails",
    "date",
    "folder_path",
    "has_attachments",
    "attachment_count",
    "attachment_metadata",
    "flags",
    "word_count",
    "source_pst",
    "extractor",
    "extraction_errors",
]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Extract PST email data to emails_raw.csv.")
    parser.add_argument("--pst-path", type=Path, required=True)
    parser.add_argument("--output-csv", type=Path, required=True)
    parser.add_argument("--log-file", type=Path)
    parser.add_argument("--extractor", choices=["auto", "pypff", "readpst"], default="auto")
    parser.add_argument("--temp-dir", type=Path)
    parser.add_argument("--max-emails", type=int)
    parser.add_argument("--readpst-command", default="readpst")
    parser.add_argument(
        "--keep-attachments-in-eml",
        action="store_true",
        help="Keep binary attachments inside exported EML files. Uses much more disk space.",
    )
    parser.add_argument("--log-level", default="INFO")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    configure_script_logging(args.log_level, args.log_file)
    rows = extract_to_rows(
        pst_path=args.pst_path,
        extractor=args.extractor,
        temp_dir=args.temp_dir,
        max_emails=args.max_emails,
        readpst_command=args.readpst_command,
        discard_attachments=not args.keep_attachments_in_eml,
    )
    write_raw_csv(rows, args.output_csv)
    log_extract_stats(rows)
    return 0


def extract_to_rows(
    pst_path: Path,
    extractor: str = "auto",
    temp_dir: Path | None = None,
    max_emails: int | None = None,
    readpst_command: str = "readpst",
    discard_attachments: bool = True,
) -> list[dict[str, object]]:
    if not pst_path.exists():
        raise FileNotFoundError(f"PST file not found: {pst_path}")

    if extractor in ("auto", "pypff"):
        try:
            return list(_extract_with_pypff(pst_path, max_emails=max_emails))
        except Exception as exc:
            if extractor == "pypff":
                raise
            logger.warning("pypff extraction unavailable or failed, falling back to readpst: %s", exc)

    return _extract_with_readpst(
        pst_path=pst_path,
        temp_dir=temp_dir,
        max_emails=max_emails,
        readpst_command=readpst_command,
        discard_attachments=discard_attachments,
    )


def write_raw_csv(rows: list[dict[str, object]], output_csv: Path) -> None:
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=RAW_COLUMNS, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        writer.writerows(rows)


def log_extract_stats(rows: list[dict[str, object]]) -> None:
    dates = pd.to_datetime([row.get("date") for row in rows], errors="coerce", utc=True).dropna()
    folders = {str(row.get("folder_path") or "") for row in rows}
    senders = Counter(str(row.get("from_email") or "") for row in rows)
    non_empty = sum(1 for row in rows if str(row.get("body") or "").strip())

    logger.info("Total emails: %s", len(rows))
    logger.info("Date range: %s to %s", dates.min() if not dates.empty else "", dates.max() if not dates.empty else "")
    logger.info("Folders scanned: %s", len(folders))
    logger.info("Top senders: %s", senders.most_common(10))
    logger.info("Empty or broken bodies: %s", len(rows) - non_empty)
    logger.info("Emails with non-empty body: %.1f%%", (non_empty / max(len(rows), 1)) * 100)


def _extract_with_readpst(
    pst_path: Path,
    temp_dir: Path | None,
    max_emails: int | None,
    readpst_command: str,
    discard_attachments: bool,
) -> list[dict[str, object]]:
    staging = temp_dir or (pst_path.parent / "_readpst_temp")
    staging.mkdir(parents=True, exist_ok=True)
    parser = EmailParser()
    extra_args = ["-t", "e", "-b"]
    if discard_attachments:
        # NotebookLM needs text first; keeping all MIME attachments can multiply disk usage.
        extra_args.extend(["-a", ".__pst_kb_keep_no_attachments__"])
    extractor = ReadpstExtractor(readpst_command, extra_args=extra_args)
    cleaner = EmailCleaner()

    try:
        refs = extractor.extract(pst_path, staging, ExtractionOptions(limit=max_emails))
        rows: list[dict[str, object]] = []
        for ref in tqdm(refs, desc="Parsing EML", unit="email"):
            try:
                raw = parser.parse_eml(ref)
                rows.append(_raw_message_to_row(raw, extractor_name="readpst", cleaner=cleaner))
            except Exception as exc:
                logger.warning("Skipping broken EML %s: %s", ref.eml_path, exc)
        return rows
    except ExtractorError:
        raise
    finally:
        if temp_dir is None:
            shutil.rmtree(staging, ignore_errors=True)


def _extract_with_pypff(pst_path: Path, max_emails: int | None = None) -> Iterable[dict[str, object]]:
    try:
        import pypff  # type: ignore[import-not-found]
    except Exception as exc:
        raise RuntimeError("pypff is not installed") from exc

    pst_file = pypff.file()
    pst_file.open(str(pst_path))
    count = 0
    cleaner = EmailCleaner()
    try:
        root = pst_file.get_root_folder()
        for raw in _walk_pypff_folder(root, pst_path.name, folder_path=""):
            yield _raw_message_to_row(raw, extractor_name="pypff", cleaner=cleaner)
            count += 1
            if max_emails is not None and count >= max_emails:
                return
    finally:
        pst_file.close()


def _walk_pypff_folder(folder: object, source_pst: str, folder_path: str) -> Iterable[RawMessage]:
    folder_name = str(_call_or_attr(folder, "name") or "")
    current_path = "/".join(part for part in [folder_path, folder_name] if part)

    message_count = int(_call_or_attr(folder, "number_of_sub_messages") or 0)
    for index in tqdm(range(message_count), desc=f"Reading {current_path or 'root'}", unit="email"):
        try:
            item = folder.get_sub_message(index)
            yield _pypff_message_to_raw(item, source_pst, current_path)
        except Exception as exc:
            logger.warning("Skipping pypff message in %s: %s", current_path, exc)

    sub_folder_count = int(_call_or_attr(folder, "number_of_sub_folders") or 0)
    for index in range(sub_folder_count):
        try:
            yield from _walk_pypff_folder(folder.get_sub_folder(index), source_pst, current_path)
        except Exception as exc:
            logger.warning("Skipping pypff folder under %s: %s", current_path, exc)


def _pypff_message_to_raw(item: object, source_pst: str, folder_path: str) -> RawMessage:
    subject = str(_call_or_attr(item, "subject") or "")
    plain_body = read_text_resilient(_call_or_attr(item, "plain_text_body"))
    html_body = read_text_resilient(_call_or_attr(item, "html_body"))
    sent = _call_or_attr(item, "client_submit_time") or _call_or_attr(item, "delivery_time")
    sender = str(_call_or_attr(item, "sender_email_address") or "")
    message_id = str(_call_or_attr(item, "identifier") or stable_hash([source_pst, folder_path, subject, sent]))[:48]
    attachment_count = int(_call_or_attr(item, "number_of_attachments") or 0)

    return RawMessage(
        source_pst=source_pst,
        source_folder=folder_path or "root",
        eml_path=Path("pypff"),
        message_id=message_id,
        conversation_id=str(_call_or_attr(item, "conversation_identifier") or "") or None,
        subject_raw=subject,
        sender_email=sender or None,
        sent_at=sent if isinstance(sent, datetime) else None,
        received_at=_call_or_attr(item, "delivery_time") if isinstance(_call_or_attr(item, "delivery_time"), datetime) else None,
        body_text_raw=plain_body,
        body_html_raw=html_body,
        attachment_payloads=[{"filename": f"attachment_{i + 1}", "payload": b""} for i in range(attachment_count)],
    )


def _raw_message_to_row(raw: RawMessage, extractor_name: str, cleaner: EmailCleaner) -> dict[str, object]:
    body = raw.body_text_raw or cleaner.html_to_text(raw.body_html_raw)
    body = normalize_whitespace(_strip_control_chars(body))
    msg_id = raw.message_id or stable_hash([raw.source_pst, raw.source_folder, raw.subject_raw, raw.sent_at, body])[:24]
    attachments = [
        {
            "filename": item.get("filename"),
            "content_type": item.get("content_type"),
            "size": len(item.get("payload") or b"") if isinstance(item.get("payload"), bytes) else None,
        }
        for item in raw.attachment_payloads
    ]
    return {
        "message_id": msg_id,
        "thread_id": raw.conversation_id or raw.in_reply_to or "",
        "subject": raw.subject_raw or "",
        "body": body,
        "from_email": raw.sender_email or "",
        "to_emails": ";".join(_emails(raw.to)),
        "cc_emails": ";".join(_emails(raw.cc)),
        "bcc_emails": ";".join(_emails(raw.bcc)),
        "date": (raw.sent_at or raw.received_at).isoformat() if (raw.sent_at or raw.received_at) else "",
        "folder_path": raw.source_folder,
        "has_attachments": bool(raw.attachment_payloads),
        "attachment_count": len(raw.attachment_payloads),
        "attachment_metadata": attachments,
        "flags": ";".join(raw.flags),
        "word_count": word_count(body),
        "source_pst": raw.source_pst,
        "extractor": extractor_name,
        "extraction_errors": ";".join(raw.extraction_errors),
    }


def _emails(recipients: list[Recipient]) -> list[str]:
    return [recipient.email for recipient in recipients if recipient.email]


def _strip_control_chars(text: str) -> str:
    return "".join(ch for ch in text if ch == "\n" or ch == "\t" or ord(ch) >= 32)


def _call_or_attr(obj: object, name: str) -> object | None:
    value = getattr(obj, name, None)
    if callable(value):
        try:
            return value()
        except TypeError:
            return None
    return value
