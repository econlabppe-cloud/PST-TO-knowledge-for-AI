from __future__ import annotations

import hashlib
import json
from collections.abc import Iterable


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8", errors="replace")).hexdigest()


def stable_hash(parts: Iterable[object]) -> str:
    payload = json.dumps(list(parts), ensure_ascii=False, sort_keys=True, default=str)
    return sha256_text(payload)
