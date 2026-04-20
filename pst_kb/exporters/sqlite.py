from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any

from pydantic import BaseModel

from pst_kb.models import AttachmentRecord, MessageRecord, ThreadRecord


def export_sqlite(
    path: Path,
    messages: list[MessageRecord],
    attachments: list[AttachmentRecord],
    threads: list[ThreadRecord],
) -> None:
    if path.exists():
        path.unlink()

    with sqlite3.connect(path) as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        _create_table(conn, "messages", MessageRecord.model_fields)
        _create_table(conn, "attachments", AttachmentRecord.model_fields)
        _create_table(conn, "threads", ThreadRecord.model_fields)
        conn.execute(
            """
            CREATE TABLE recipients (
                record_id TEXT NOT NULL,
                recipient_type TEXT NOT NULL,
                email TEXT NOT NULL
            )
            """
        )

        _insert_models(conn, "messages", messages)
        _insert_models(conn, "attachments", attachments)
        _insert_models(conn, "threads", threads)
        _insert_recipients(conn, messages)
        _create_indexes(conn)


def _create_table(conn: sqlite3.Connection, table: str, fields: object) -> None:
    columns = ", ".join(f'"{name}" TEXT' for name in fields)
    conn.execute(f'CREATE TABLE "{table}" ({columns})')


def _insert_models(conn: sqlite3.Connection, table: str, rows: list[BaseModel]) -> None:
    if not rows:
        return
    fieldnames = list(rows[0].model_fields)
    placeholders = ", ".join("?" for _ in fieldnames)
    columns = ", ".join(f'"{field}"' for field in fieldnames)
    sql = f'INSERT INTO "{table}" ({columns}) VALUES ({placeholders})'
    conn.executemany(
        sql,
        [[_sqlite_value(row.model_dump(mode="json").get(field)) for field in fieldnames] for row in rows],
    )


def _insert_recipients(conn: sqlite3.Connection, messages: list[MessageRecord]) -> None:
    values: list[tuple[str, str, str]] = []
    for message in messages:
        for kind, emails in (
            ("to", message.to_emails),
            ("cc", message.cc_emails),
            ("bcc", message.bcc_emails),
        ):
            values.extend((message.record_id, kind, email) for email in emails)
    conn.executemany(
        "INSERT INTO recipients (record_id, recipient_type, email) VALUES (?, ?, ?)",
        values,
    )


def _create_indexes(conn: sqlite3.Connection) -> None:
    for statement in (
        "CREATE INDEX idx_messages_thread_key ON messages(thread_key)",
        "CREATE INDEX idx_messages_message_id ON messages(message_id)",
        "CREATE INDEX idx_messages_content_hash ON messages(content_hash)",
        "CREATE INDEX idx_messages_sender_email ON messages(sender_email)",
        "CREATE INDEX idx_messages_sent_at ON messages(sent_at)",
        "CREATE INDEX idx_recipients_record_id ON recipients(record_id)",
        "CREATE INDEX idx_attachments_record_id ON attachments(record_id)",
    ):
        conn.execute(statement)


def _sqlite_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, dict)):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    return str(value)
