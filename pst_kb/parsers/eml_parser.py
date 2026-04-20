from __future__ import annotations

import logging
from datetime import datetime
from email import policy
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
        errors: list[str] = []
        try:
            with ref.eml_path.open("rb") as handle:
                message = BytesParser(policy=policy.default).parse(handle)
        except Exception as exc:
            raise ValueError(f"Unable to parse EML file {ref.eml_path}: {exc}") from exc

        text_parts: list[str] = []
        html_parts: list[str] = []
        attachment_payloads: list[dict[str, object]] = []

        for part in _iter_leaf_parts(message):
            disposition = (part.get_content_disposition() or "").lower()
            content_type = part.get_content_type()
            filename = part.get_filename()
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
        return RawMessage(
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
            references=[_clean_header_id(value) or value for value in (message.get("References", "").split()) if value],
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
        )


def _iter_leaf_parts(message: Message) -> Iterable[Message]:
    if message.is_multipart():
        parts = message.iter_parts() if isinstance(message, EmailMessage) else message.get_payload()
        for part in parts:
            yield from _iter_leaf_parts(part)
    else:
        yield message


def _safe_get_content(part: Message, errors: list[str]) -> str:
    try:
        content = part.get_content()
        return content if isinstance(content, str) else str(content)
    except Exception as exc:
        errors.append(f"body_decode_error:{part.get_content_type()}:{exc}")
        try:
            payload = part.get_payload(decode=True) or b""
            charset = part.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace")
        except Exception as nested_exc:
            errors.append(f"body_decode_fallback_error:{nested_exc}")
            return ""


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
