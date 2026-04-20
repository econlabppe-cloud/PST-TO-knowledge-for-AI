from __future__ import annotations

import re

PREFIX_PATTERN = re.compile(
    r"^\s*((re|fw|fwd|sv|aw|wg|תגובה|השב|הועבר|העברה)\s*[:：]\s*)+",
    re.IGNORECASE,
)


def normalize_whitespace(value: str) -> str:
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in value.split("\n")]
    compacted: list[str] = []
    blank_seen = False
    for line in lines:
        if not line:
            if not blank_seen:
                compacted.append("")
            blank_seen = True
        else:
            compacted.append(line)
            blank_seen = False
    return "\n".join(compacted).strip()


def normalize_subject(subject: str | None) -> str:
    if not subject:
        return ""
    normalized = subject.replace("\u200f", "").replace("\u200e", "")
    previous = None
    while previous != normalized:
        previous = normalized
        normalized = PREFIX_PATTERN.sub("", normalized)
    normalized = re.sub(r"\[[^\]]*(external|חיצוני)[^\]]*\]", "", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"\s+", " ", normalized).strip().lower()
    return normalized


def normalize_email(email: str | None) -> str | None:
    if not email:
        return None
    email = email.strip().strip("<>").lower()
    return email or None


def detect_reply_forward_indicator(subject: str | None) -> str | None:
    if not subject:
        return None
    lowered = subject.lower().strip()
    if re.match(r"^(fw|fwd|הועבר|העברה)\s*[:：]", lowered):
        return "forward"
    if re.match(r"^(re|תגובה|השב)\s*[:：]", lowered):
        return "reply"
    return None
