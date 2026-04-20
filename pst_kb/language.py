from __future__ import annotations

import re


def detect_language(text: str) -> str | None:
    if not text.strip():
        return None
    hebrew_chars = len(re.findall(r"[\u0590-\u05ff]", text))
    latin_chars = len(re.findall(r"[A-Za-z]", text))

    try:
        from langdetect import DetectorFactory, detect

        DetectorFactory.seed = 0
        return detect(text[:4000])
    except Exception:
        if hebrew_chars > latin_chars and hebrew_chars > 5:
            return "he"
        if latin_chars > 5:
            return "en"
        return None
