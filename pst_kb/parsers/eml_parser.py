from __future__ import annotations

import logging
import re
from datetime import datetime
from email import policy
from email.header import decode_header, make_header
from email.message import EmailMessage, Message
from email.parser import BytesParser
from email.utils import getaddresses, parseaddr, parsedate_to_datetime
from pathlib import Path
from typing import Iterable

from pst_kb.models import ExtractedMessageRef, RawMessage, Recipient
from pst_kb.normalizers.text import normalize_email

logger = logging.getLogger(__name__)


class EmailParser:
    def parse_eml(self, ref: ExtractedMessageRef) -> RawMessage:
        messages = self.parse_eml_with_nested(ref)
        if not messages:
            raise ValueError(f"Unable to parse EML file {ref.eml_path}: no message objects were produced")
        return messages[0]

    def parse_eml_with_nested(self, ref: ExtractedMessageRef) -> list[RawMessage]:
        try:
            with ref.eml_path.open("rb") as handle:
                message = BytesParser(policy=policy.default).parse(handle)
        except Exception as exc:
            raise ValueError(f"Unable to parse EML file {ref.eml_path}: {exc}") from exc

        raw_messages: list[RawMessage] = []
        self._collect_messages(
            message=message,
            ref=ref,
            output=raw_messages,
            nested_depth=0,
            parent_record_message_id=None,
            container_attachment_filename=None,
        )
        return raw_messages

    def _collect_messages(
        self,
        *,
        message: Message,
        ref: ExtractedMessageRef,
        output: list[RawMessage],
        nested_depth: int,
        parent_record_message_id: str | None,
        container_attachment_filename: str | None,
    ) -> None:
        errors: list[str] = []
        text_parts: list[str] = []
        html_parts: list[str] = []
        attachment_payloads: list[dict[str, object]] = []
        nested_parts: list[tuple[Message, str | None]] = []

        for part in _iter_leaf_parts(message):
            content_type = part.get_content_type()
            disposition = (part.get_content_disposition() or "").lower()
            filename = part.get_filename()

            if content_type == "message/rfc822":
                nested_message = _extract_nested_message(part, errors)
                if nested_message is not None:
                    nested_parts.append((nested_message, filename))
                    continue

            is_attachment = disposition == "attachment" or bool(filename)
            if is_attachment:
                try:
                    payload = part.get_payload(decode=True) or b""
                except Exception as exc:
                    payload = b""
                    errors.append(f"attachment_decode_error:{filename or content_type}:{exc}")
                attachment_payloads.append(
                    {
                        "filename": filename or "attachment",
                        "content_type": content_type,
                        "content_id": part.get("Content-ID"),
                        "is_inline": disposition == "inline",
                        "payload": payload,
                    }
                )
                continue

            if content_type == "text/plain":
                text_parts.append(_safe_get_content(part, errors))
            elif content_type == "text/html":
                html_parts.append(_safe_get_content(part, errors))

        sender_name, sender_email = parseaddr(message.get("From", ""))
        headers = _headers_to_dict(message.items())
        raw = RawMessage(
            source_pst=ref.pst_path.name,
            source_folder=ref.source_folder,
            eml_path=ref.eml_path,
            message_id=_clean_header_id(message.get("Message-ID")),
            conversation_id=_first_header(
                message,
                "X-MS-Conversation-ID",
                "Thread-Index",
                "X-MS-Conversation-Index",
                "Thread-Topic",
            ),
            in_reply_to=_clean_header_id(message.get("In-Reply-To")),
            references=[_clean_header_id(value) or value for value in message.get("References", "").split() if value],
            subject_raw=message.get("Subject"),
            sender_name=sender_name or None,
            sender_email=normalize_email(sender_email),
            to=_parse_recipients(message.get_all("To", [])),
            cc=_parse_recipients(message.get_all("Cc", [])),
            bcc=_parse_recipients(message.get_all("Bcc", [])),
            sent_at=_parse_date(message.get("Date"), errors),
            received_at=_parse_received_date(message.get_all("Received", []), errors),
            body_text_raw="\n\n".join(part for part in text_parts if part).strip(),
            body_html_raw="\n\n".join(part for part in html_parts if part).strip(),
            internet_headers=headers,
            importance=_first_header(message, "Importance", "X-Priority", "Priority"),
            flags=_split_header_values(_first_header(message, "X-Message-Flag", "Keywords")),
            categories=_split_header_values(_first_header(message, "Categories", "X-Categories")),
            attachment_payloads=attachment_payloads,
            extraction_errors=errors,
            parent_record_message_id=parent_record_message_id,
            nested_depth=nested_depth,
            container_attachment_filename=container_attachment_filename,
        )
        output.append(raw)

        current_parent_id = raw.message_id or parent_record_message_id
        embedded_text_messages = _extract_embedded_messages_from_text(
            raw=raw,
            parent_record_message_id=current_parent_id,
            nested_depth=nested_depth + 1,
        )
        if embedded_text_messages:
            output.extend(embedded_text_messages)
        for nested_message, nested_filename in nested_parts:
            self._collect_messages(
                message=nested_message,
                ref=ref,
                output=output,
                nested_depth=nested_depth + 1,
                parent_record_message_id=current_parent_id,
                container_attachment_filename=nested_filename,
            )


def _iter_leaf_parts(message: Message) -> Iterable[Message]:
    if message.get_content_type() == "message/rfc822":
        yield message
        return
    if message.is_multipart():
        parts = message.iter_parts() if isinstance(message, EmailMessage) else message.get_payload()
        for part in parts:
            yield from _iter_leaf_parts(part)
    else:
        yield message


def _extract_nested_message(part: Message, errors: list[str]) -> Message | None:
    payload = part.get_payload()
    if isinstance(payload, list) and payload:
        first = payload[0]
        if isinstance(first, Message):
            return first

    decoded = part.get_payload(decode=True)
    if decoded:
        try:
            return BytesParser(policy=policy.default).parsebytes(decoded)
        except Exception as exc:
            errors.append(f"nested_message_parse_error:{exc}")
    return None


def _safe_get_content(part: Message, errors: list[str]) -> str:
    payload = _get_text_payload_bytes(part)
    if payload:
        decoded = _decode_text_payload(payload, part.get_content_charset())
        if decoded:
            return decoded

    try:
        content = part.get_content()
        text = content if isinstance(content, str) else str(content)
        return _repair_text(text)
    except Exception as exc:
        errors.append(f"body_decode_error:{part.get_content_type()}:{exc}")
        try:
            charset = part.get_content_charset() or "utf-8"
            return _repair_text(payload.decode(charset, errors="replace")) if payload else ""
        except Exception as nested_exc:
            errors.append(f"body_decode_fallback_error:{nested_exc}")
            return ""


def _get_text_payload_bytes(part: Message) -> bytes:
    try:
        decoded = part.get_payload(decode=True)
        if isinstance(decoded, bytes):
            return decoded
    except Exception:
        pass

    try:
        payload = part.get_payload(decode=False)
    except Exception:
        return b""

    if isinstance(payload, str):
        charset = part.get_content_charset() or "utf-8"
        for encoding in (charset, "utf-8", "windows-1255", "latin-1"):
            try:
                return payload.encode(encoding, errors="surrogateescape")
            except Exception:
                continue
        return payload.encode("utf-8", errors="replace")
    return b""


def _decode_text_payload(payload: bytes, declared_charset: str | None) -> str:
    # readpst/Outlook exports often carry a wrong MIME charset. Decode several
    # plausible encodings and choose the least damaged Hebrew-friendly text.
    encodings: list[str] = []
    seen: set[str] = set()
    for encoding in (
        declared_charset,
        "utf-8",
        "utf-8-sig",
        "windows-1255",
        "cp1255",
        "iso-8859-8",
        "cp1252",
        "latin-1",
    ):
        if not encoding:
            continue
        normalized = encoding.lower()
        if normalized not in seen:
            seen.add(normalized)
            encodings.append(encoding)

    candidates: list[str] = []
    for encoding in encodings:
        try:
            candidates.append(payload.decode(encoding, errors="strict"))
        except Exception:
            try:
                candidates.append(payload.decode(encoding, errors="replace"))
            except Exception:
                continue

    repaired_candidates = [_repair_text(candidate) for candidate in candidates if candidate]
    return max(repaired_candidates, key=_text_quality_score, default="").strip()


def _text_quality_score(value: str) -> float:
    if not value:
        return -1_000_000

    length = max(len(value), 1)
    replacement_count = value.count("\ufffd") + value.count("?")
    hebrew_count = sum("\u05d0" <= char <= "\u05ea" for char in value)
    latin_count = sum("a" <= char.lower() <= "z" for char in value)
    printable_count = sum(char.isprintable() or char in "\r\n\t" for char in value)
    mojibake_markers = (
        value.count("×")
        + value.count("Ã")
        + value.count("Â")
        + value.count("׳")
        + value.count("ֲ")
        + value.count("ג€")
        + value.count("ן¿½")
    )
    control_count = sum((ord(char) < 32 and char not in "\r\n\t") or 127 <= ord(char) <= 159 for char in value)

    return (
        printable_count / length * 20
        + hebrew_count * 3
        + latin_count * 0.25
        - replacement_count * 50
        - mojibake_markers * 30
        - control_count * 15
    )


def _parse_recipients(values: list[str]) -> list[Recipient]:
    return [
        Recipient(name=name or None, email=normalize_email(email))
        for name, email in getaddresses(values)
        if name or email
    ]


def _parse_date(value: str | None, errors: list[str]) -> datetime | None:
    if not value:
        return None
    try:
        return parsedate_to_datetime(value)
    except Exception as exc:
        errors.append(f"date_parse_error:{value}:{exc}")
        return None


def _parse_received_date(values: list[str], errors: list[str]) -> datetime | None:
    for value in values:
        if ";" not in value:
            continue
        candidate = value.rsplit(";", 1)[-1].strip()
        parsed = _parse_date(candidate, errors)
        if parsed is not None:
            return parsed
    return None


def _headers_to_dict(items: Iterable[tuple[str, str]]) -> dict[str, str]:
    headers: dict[str, str] = {}
    for key, value in items:
        if key in headers:
            headers[key] = f"{headers[key]}\n{value}"
        else:
            headers[key] = value
    return headers


def _first_header(message: Message, *names: str) -> str | None:
    for name in names:
        value = message.get(name)
        if value:
            return str(value)
    return None


def _split_header_values(value: str | None) -> list[str]:
    if not value:
        return []
    return [part.strip() for part in value.replace(";", ",").split(",") if part.strip()]


def _clean_header_id(value: str | None) -> str | None:
    if not value:
        return None
    return value.strip().strip("<>")


EMBEDDED_START_RE = re.compile(r"(?m)^(?:>?\s*)From\s+\"?.+?\"\s+\w{3}\s+\w{3}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}\s+\d{4}\s*$")
HEADER_PREFIXES = ("from:", "to:", "cc:", "bcc:", "subject:", "date:", "sent:")


def _extract_embedded_messages_from_text(
    *,
    raw: RawMessage,
    parent_record_message_id: str | None,
    nested_depth: int,
) -> list[RawMessage]:
    text = raw.body_text_raw or ""
    matches = list(EMBEDDED_START_RE.finditer(text))
    if not matches:
        return []

    embedded_messages: list[RawMessage] = []
    for index, match in enumerate(matches):
        start = match.start()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        block = text[start:end].strip()
        parsed = _parse_embedded_text_block(
            raw=raw,
            block=block,
            parent_record_message_id=parent_record_message_id,
            nested_depth=nested_depth,
            block_index=index + 1,
        )
        if parsed is not None:
            embedded_messages.append(parsed)

    prefix = text[: matches[0].start()].strip()
    if embedded_messages:
        raw.body_text_raw = prefix
    return embedded_messages


def _parse_embedded_text_block(
    *,
    raw: RawMessage,
    block: str,
    parent_record_message_id: str | None,
    nested_depth: int,
    block_index: int,
) -> RawMessage | None:
    lines = [line.rstrip() for line in block.splitlines()]
    if len(lines) < 3:
        return None

    parsed_message = _parse_embedded_block_as_message(block)
    if parsed_message is not None:
        raw_from_message = _embedded_message_to_raw(
            raw=raw,
            message=parsed_message,
            parent_record_message_id=parent_record_message_id,
            nested_depth=nested_depth,
            block_index=block_index,
        )
        if raw_from_message is not None:
            return raw_from_message

    header_values: dict[str, str] = {}
    body_lines: list[str] = []
    in_headers = False
    body_started = False

    for raw_line in lines[1:]:
        line = raw_line.lstrip("> ").rstrip()
        lowered = line.lower()
        if not body_started and any(lowered.startswith(prefix) for prefix in HEADER_PREFIXES):
            in_headers = True
            key, _, value = line.partition(":")
            header_values[key.strip().lower()] = _decode_embedded_header(value.strip())
            continue
        if in_headers and not line.strip():
            body_started = True
            continue
        if in_headers and any(lowered.startswith(prefix) for prefix in HEADER_PREFIXES):
            key, _, value = line.partition(":")
            header_values[key.strip().lower()] = _decode_embedded_header(value.strip())
            continue
        body_started = True
        body_lines.append(raw_line.lstrip("> ").rstrip())

    if not header_values and not body_lines:
        return None

    mbox_sender = _sender_from_mbox_line(lines[0])
    sender_name, sender_email = parseaddr(header_values.get("from", "") or mbox_sender)
    sent_at = _parse_date(header_values.get("date") or header_values.get("sent"), [])
    subject = _repair_text(header_values.get("subject", ""))
    body_text = _repair_text("\n".join(body_lines).strip())

    if not any((subject.strip(), sender_email.strip(), body_text.strip())):
        return None

    embedded_message_id = _clean_header_id(header_values.get("message-id")) or (
        f"{raw.message_id or raw.eml_path.name}-embedded-{block_index}"
    )
    return RawMessage(
        source_pst=raw.source_pst,
        source_folder=raw.source_folder,
        eml_path=raw.eml_path,
        message_id=embedded_message_id,
        conversation_id=raw.conversation_id,
        in_reply_to=raw.message_id,
        references=list(raw.references),
        subject_raw=subject or None,
        sender_name=sender_name or None,
        sender_email=normalize_email(sender_email),
        to=_parse_recipients([header_values.get("to", "")] if header_values.get("to") else []),
        cc=_parse_recipients([header_values.get("cc", "")] if header_values.get("cc") else []),
        bcc=_parse_recipients([header_values.get("bcc", "")] if header_values.get("bcc") else []),
        sent_at=sent_at,
        received_at=sent_at,
        body_text_raw=body_text,
        body_html_raw="",
        internet_headers={key.title(): value for key, value in header_values.items()},
        importance=None,
        flags=[],
        categories=[],
        attachment_payloads=[],
        extraction_errors=["embedded_forwarded_text"],
        parent_record_message_id=parent_record_message_id,
        nested_depth=nested_depth,
        container_attachment_filename="embedded_forwarded_text",
    )


def _decode_embedded_header(value: str) -> str:
    try:
        return _repair_text(str(make_header(decode_header(value))))
    except Exception:
        return _repair_text(value)


def _parse_embedded_block_as_message(block: str) -> Message | None:
    lines = block.splitlines()
    if len(lines) < 2:
        return None
    candidate = "\n".join(lines[1:]).strip()
    if not _looks_like_rfc822_message(candidate):
        return None
    try:
        return BytesParser(policy=policy.default).parsebytes(candidate.encode("utf-8", errors="replace"))
    except Exception:
        return None


def _embedded_message_to_raw(
    *,
    raw: RawMessage,
    message: Message,
    parent_record_message_id: str | None,
    nested_depth: int,
    block_index: int,
) -> RawMessage | None:
    errors: list[str] = ["embedded_forwarded_text"]
    text_parts: list[str] = []
    html_parts: list[str] = []

    for part in _iter_leaf_parts(message):
        content_type = part.get_content_type()
        disposition = (part.get_content_disposition() or "").lower()
        filename = part.get_filename()
        if disposition == "attachment" or filename or content_type == "message/rfc822":
            continue
        if content_type == "text/plain":
            text_parts.append(_safe_get_content(part, errors))
        elif content_type == "text/html":
            html_parts.append(_safe_get_content(part, errors))

    sender_name, sender_email = parseaddr(message.get("From", ""))
    subject = _repair_text(str(message.get("Subject", "") or ""))
    body_text = _repair_text("\n\n".join(part for part in text_parts if part).strip())
    body_html = _repair_text("\n\n".join(part for part in html_parts if part).strip())
    sent_at = _parse_date(message.get("Date"), errors)

    if not any((subject.strip(), sender_email.strip(), body_text.strip(), body_html.strip())):
        return None

    message_id = _clean_header_id(message.get("Message-ID")) or f"{raw.message_id or raw.eml_path.name}-embedded-{block_index}"
    return RawMessage(
        source_pst=raw.source_pst,
        source_folder=raw.source_folder,
        eml_path=raw.eml_path,
        message_id=message_id,
        conversation_id=_first_header(message, "X-MS-Conversation-ID", "Thread-Index", "Thread-Topic") or raw.conversation_id,
        in_reply_to=_clean_header_id(message.get("In-Reply-To")) or raw.message_id,
        references=[_clean_header_id(value) or value for value in message.get("References", "").split() if value],
        subject_raw=subject or None,
        sender_name=sender_name or None,
        sender_email=normalize_email(sender_email),
        to=_parse_recipients(message.get_all("To", [])),
        cc=_parse_recipients(message.get_all("Cc", [])),
        bcc=_parse_recipients(message.get_all("Bcc", [])),
        sent_at=sent_at,
        received_at=sent_at or _parse_received_date(message.get_all("Received", []), errors),
        body_text_raw=body_text,
        body_html_raw=body_html,
        internet_headers=_headers_to_dict(message.items()),
        importance=_first_header(message, "Importance", "X-Priority", "Priority"),
        flags=_split_header_values(_first_header(message, "X-Message-Flag", "Keywords")),
        categories=_split_header_values(_first_header(message, "Categories", "X-Categories")),
        attachment_payloads=[],
        extraction_errors=errors,
        parent_record_message_id=parent_record_message_id,
        nested_depth=nested_depth,
        container_attachment_filename="embedded_forwarded_text",
    )


def _looks_like_rfc822_message(text: str) -> bool:
    lowered = text[:2000].lower()
    return "from:" in lowered and ("subject:" in lowered or "date:" in lowered or "message-id:" in lowered)


def _sender_from_mbox_line(line: str) -> str:
    match = re.match(r'^>?\s*From\s+"?([^"\s]+@[^"\s]+)', line.strip())
    return match.group(1) if match else ""


def _repair_text(value: str) -> str:
    if not value:
        return ""
    candidates = [value]
    queue = [value]
    for _ in range(2):
        next_queue: list[str] = []
        for candidate in queue:
            for encoding in ("latin-1", "cp1252", "cp1255"):
                try:
                    repaired = candidate.encode(encoding, errors="strict").decode("utf-8", errors="strict")
                except Exception:
                    try:
                        repaired = candidate.encode(encoding, errors="replace").decode("utf-8", errors="replace")
                    except Exception:
                        continue
                if repaired and repaired not in candidates:
                    candidates.append(repaired)
                    next_queue.append(repaired)
        queue = next_queue
    return max(candidates, key=_text_quality_score)
