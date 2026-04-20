from __future__ import annotations

from collections import defaultdict
from datetime import datetime

from pst_kb.models import MessageRecord, ThreadRecord
from pst_kb.utils.hashing import stable_hash


class ThreadBuilder:
    def assign_threads(self, records: list[MessageRecord]) -> tuple[list[MessageRecord], list[ThreadRecord]]:
        for record in records:
            native_key = self._native_thread_key(record)
            synthetic_key = self._synthetic_thread_key(record)
            record.synthetic_thread_key = synthetic_key
            record.thread_key = native_key or synthetic_key

        groups: dict[str, list[MessageRecord]] = defaultdict(list)
        for record in records:
            groups[record.thread_key or record.record_id].append(record)

        thread_records: list[ThreadRecord] = []
        for thread_key, group in groups.items():
            ordered = sorted(group, key=_message_order_key)
            for index, record in enumerate(ordered, start=1):
                record.thread_position = index
            thread_records.append(self._build_thread_record(thread_key, ordered))

        return records, sorted(thread_records, key=lambda item: item.thread_key)

    def _native_thread_key(self, record: MessageRecord) -> str | None:
        if record.conversation_id:
            return f"native:{stable_hash([record.conversation_id])[:24]}"
        if record.references:
            return f"refs:{stable_hash([record.references[0]])[:24]}"
        if record.in_reply_to:
            return f"reply:{stable_hash([record.in_reply_to])[:24]}"
        return None

    def _synthetic_thread_key(self, record: MessageRecord) -> str:
        participants = sorted(
            email for email in [record.sender_email, *record.to_emails, *record.cc_emails] if email
        )
        date_bucket = _date_bucket(record.sent_at or record.received_at)
        return f"synthetic:{stable_hash([record.subject_normalized, participants, date_bucket])[:24]}"

    def _build_thread_record(self, thread_key: str, records: list[MessageRecord]) -> ThreadRecord:
        participants = sorted(
            {
                email
                for record in records
                for email in [record.sender_email, *record.to_emails, *record.cc_emails, *record.bcc_emails]
                if email
            }
        )
        dates = [date for record in records for date in [record.sent_at or record.received_at] if date is not None]
        return ThreadRecord(
            thread_key=thread_key,
            source_psts=sorted({record.source_pst for record in records}),
            subject_normalized=records[0].subject_normalized if records else "",
            message_count=len(records),
            first_message_at=min(dates) if dates else None,
            last_message_at=max(dates) if dates else None,
            participant_emails=participants,
            record_ids=[record.record_id for record in records],
            duplicate_count=sum(1 for record in records if record.is_duplicate),
            has_attachments=any(record.has_attachments for record in records),
        )


def _date_bucket(value: datetime | None) -> str | None:
    if value is None:
        return None
    year, week, _ = value.isocalendar()
    return f"{year}-W{week:02d}"


def _message_order_key(record: MessageRecord) -> tuple[str, str]:
    value = record.sent_at or record.received_at
    return (value.isoformat() if value else "", record.record_id)
