from __future__ import annotations

from pathlib import Path

from pst_kb.models import ExtractedMessageRef
from pst_kb.parsers import EmailParser


def test_parser_extracts_nested_rfc822_messages(tmp_path: Path) -> None:
    eml_path = tmp_path / "nested.eml"
    eml_path.write_text(
        "\n".join(
            [
                "From: outer@example.org",
                "To: inbox@example.org",
                "Subject: Outer subject",
                "Message-ID: <outer-1@example.org>",
                "MIME-Version: 1.0",
                'Content-Type: multipart/mixed; boundary="outer-boundary"',
                "",
                "--outer-boundary",
                'Content-Type: text/plain; charset="utf-8"',
                "",
                "Outer body",
                "--outer-boundary",
                'Content-Type: message/rfc822; name="forwarded.eml"',
                "Content-Disposition: attachment; filename=\"forwarded.eml\"",
                "",
                "From: nested@example.org",
                "To: inbox@example.org",
                "Subject: Nested subject",
                "Message-ID: <nested-1@example.org>",
                "MIME-Version: 1.0",
                'Content-Type: text/plain; charset="utf-8"',
                "",
                "Nested body",
                "--outer-boundary--",
                "",
            ]
        ),
        encoding="utf-8",
    )

    ref = ExtractedMessageRef(
        pst_path=tmp_path / "sample.pst",
        eml_path=eml_path,
        source_folder="Inbox",
    )

    messages = EmailParser().parse_eml_with_nested(ref)

    assert len(messages) == 2
    assert messages[0].subject_raw == "Outer subject"
    assert messages[1].subject_raw == "Nested subject"
    assert messages[1].parent_record_message_id == "outer-1@example.org"
    assert messages[1].nested_depth == 1
    assert messages[1].container_attachment_filename == "forwarded.eml"


def test_parser_extracts_embedded_forwarded_text_messages(tmp_path: Path) -> None:
    eml_path = tmp_path / "forwarded_text.eml"
    eml_path.write_text(
        "\n".join(
            [
                "From: outer@example.org",
                "To: inbox@example.org",
                "Subject: Package",
                "Message-ID: <outer-2@example.org>",
                "MIME-Version: 1.0",
                'Content-Type: text/plain; charset="utf-8"',
                "",
                "Here are the forwarded emails",
                "",
                'From "sender@example.org" Mon Apr 22 12:36:41 2024',
                "From: Arik FN <yasissman@gmail.com>",
                "To: <hr2@iba.org.il>",
                "Subject: מסמכים",
                "Date: Mon, 22 Apr 2024 12:36:41 +0300",
                "",
                "אבקש מסמכים ואישור שנתי.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    ref = ExtractedMessageRef(
        pst_path=tmp_path / "sample.pst",
        eml_path=eml_path,
        source_folder="Inbox",
    )

    messages = EmailParser().parse_eml_with_nested(ref)

    assert len(messages) == 2
    assert messages[0].subject_raw == "Package"
    assert messages[0].body_text_raw == "Here are the forwarded emails"
    assert messages[1].subject_raw == "מסמכים"
    assert messages[1].sender_email == "yasissman@gmail.com"
    assert messages[1].parent_record_message_id == "outer-2@example.org"
    assert messages[1].container_attachment_filename == "embedded_forwarded_text"


def test_parser_decodes_embedded_rfc822_headers(tmp_path: Path) -> None:
    eml_path = tmp_path / "encoded_forwarded_text.eml"
    eml_path.write_text(
        "\n".join(
            [
                "From: outer@example.org",
                "To: inbox@example.org",
                "Subject: Package",
                "Message-ID: <outer-3@example.org>",
                "MIME-Version: 1.0",
                'Content-Type: text/plain; charset="utf-8"',
                "",
                'From "sender@example.org" Mon Apr 22 12:36:41 2024',
                "From: Sender <sender@example.org>",
                "To: <hr2@iba.org.il>",
                "Subject: =?UTF-8?B?15jXldek16EgMTA2?=",
                "Date: Mon, 22 Apr 2024 12:36:41 +0300",
                "Message-ID: <inner-encoded@example.org>",
                "MIME-Version: 1.0",
                'Content-Type: text/plain; charset="utf-8"',
                "",
                "אבקש לקבל טופס 106.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    ref = ExtractedMessageRef(
        pst_path=tmp_path / "sample.pst",
        eml_path=eml_path,
        source_folder="Inbox",
    )

    messages = EmailParser().parse_eml_with_nested(ref)

    assert len(messages) == 2
    assert messages[1].message_id == "inner-encoded@example.org"
    assert messages[1].subject_raw == "טופס 106"
    assert messages[1].sender_email == "sender@example.org"
