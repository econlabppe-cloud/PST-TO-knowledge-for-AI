from __future__ import annotations

from collections import defaultdict

from pst_kb.models import MessageRecord
from pst_kb.utils.hashing import stable_hash


class Deduplicator:
    def mark_duplicates(self, records: list[MessageRecord]) -> list[MessageRecord]:
        canonical_by_key: dict[str, str] = {}
        for record in sorted(records, key=_sort_key):
            key = self._dedup_key(record)
            canonical = canonical_by_key.get(key)
            if canonical:
                record.is_duplicate = True
                record.duplicate_of = canonical
            else:
                canonical_by_key[key] = record.record_id
        return records

    def _dedup_key(self, record: MessageRecord) -> str:
        if record.message_id:
            return f"message-id:{record.message_id.lower()}"

        participant_key = sorted(
            email for email in [record.sender_email, *record.to_emails, *record.cc_emails] if email
        )
        metadata_fingerprint = stable_hash(
            [
                record.subject_normalized,
                participant_key,
                record.sent_at.date().isoformat() if record.sent_at else None,
                record.content_hash,
            ]
        )
        return f"metadata:{metadata_fingerprint}"


def group_duplicates(records: list[MessageRecord]) -> dict[str, list[str]]:
    groups: dict[str, list[str]] = defaultdict(list)
    for record in records:
        groups[record.duplicate_of or record.record_id].append(record.record_id)
    return dict(groups)


def _sort_key(record: MessageRecord) -> tuple[str, str]:
    date_value = record.sent_at or record.received_at
    return (date_value.isoformat() if date_value else "", record.record_id)
