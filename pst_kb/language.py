from __future__ import annotations

import re

try:
    from langdetect import DetectorFactory, detect

    DetectorFactory.seed = 0
    HAS_LANGDETECT = True
except ImportError:
    HAS_LANGDETECT = False


def detect_language(text: str) -> str | None:
    """
    Detect the language of the text, preferring Hebrew or English.
    Uses langdetect if available, with a character-based fallback.
    """
    if not text or not text.strip():
        return None

    if HAS_LANGDETECT:
        try:
            return detect(text[:4000])
        except Exception:
            # Fallback to heuristics if langdetect fails (e.g., no features)
            pass

    hebrew_chars = len(re.findall(r"[\u0590-\u05ff]", text))
    latin_chars = len(re.findall(r"[A-Za-z]", text))

    if hebrew_chars > latin_chars and hebrew_chars > 5:
        return "he"
    if latin_chars > 5:
        return "en"

    return None
