from __future__ import annotations

import re
from dataclasses import dataclass

from bs4 import BeautifulSoup

from pst_kb.normalizers.text import normalize_whitespace


@dataclass(frozen=True)
class CleanResult:
    text: str
    mostly_quoted: bool
    notes: list[str]


class EmailCleaner:
    """Heuristic cleaner tuned for mixed Hebrew and English corporate email."""

    REPLY_SEPARATOR_PATTERNS = [
        re.compile(r"^-{2,}\s*(Original Message|הודעה מקורית)\s*-{2,}$", re.IGNORECASE),
        re.compile(r"^_{5,}\s*$"),
        re.compile(r"^\s*From:\s+.+", re.IGNORECASE),
        re.compile(r"^\s*Sent:\s+.+", re.IGNORECASE),
        re.compile(r"^\s*To:\s+.+", re.IGNORECASE),
        re.compile(r"^\s*Subject:\s+.+", re.IGNORECASE),
        re.compile(r"^\s*מאת:\s+.+"),
        re.compile(r"^\s*נשלח:\s+.+"),
        re.compile(r"^\s*אל:\s+.+"),
        re.compile(r"^\s*נושא:\s+.+"),
        re.compile(r"^\s*בתאריך\s+.+כתב/?ה?:\s*$"),
        re.compile(r"^\s*On\s+.+wrote:\s*$", re.IGNORECASE),
        re.compile(r"^\s*Begin forwarded message:", re.IGNORECASE),
    ]
    SIGNATURE_PATTERNS = [
        re.compile(r"^--\s*$"),
        re.compile(r"^\s*(best regards|kind regards|regards|thanks|thank you|sincerely|yours|cheers),?\s*$", re.IGNORECASE),
        re.compile(r"^\s*(בברכה|תודה|תודה רבה|בכבוד רב|בברכת חברים|יום טוב|שבת שלום),?\s*$"),
    ]
    QUOTED_LINE_PATTERN = re.compile(r"^\s*[>|]\s?")

    def html_to_text(self, html: str) -> str:
        if not html:
            return ""
        soup = BeautifulSoup(html, "html.parser")
        for tag in soup(["script", "style", "meta", "head"]):
            tag.decompose()
        for br in soup.find_all("br"):
            br.replace_with("\n")
        for block in soup.find_all(["p", "div", "tr", "li"]):
            block.append("\n")
        return normalize_whitespace(soup.get_text("\n"))

    def clean(self, text: str, html: str = "") -> CleanResult:
        source = text or self.html_to_text(html)
        source = normalize_whitespace(source)
        if not source:
            return CleanResult(text="", mostly_quoted=False, notes=["empty_body"])

        lines = source.split("\n")
        quoted_lines = sum(1 for line in lines if self.QUOTED_LINE_PATTERN.match(line))
        mostly_quoted = bool(lines) and quoted_lines / max(len(lines), 1) >= 0.6

        reply_cut = self._find_reply_cut(lines)
        notes: list[str] = []
        if reply_cut is not None:
            lines = lines[:reply_cut]
            notes.append("reply_history_removed")

        lines = self._strip_quoted_prefix_lines(lines)
        lines, signature_removed = self._strip_signature(lines)
        if signature_removed:
            notes.append("signature_removed")

        cleaned = normalize_whitespace("\n".join(lines))
        if source and len(cleaned) < max(80, len(source) * 0.25):
            mostly_quoted = True
        return CleanResult(text=cleaned, mostly_quoted=mostly_quoted, notes=notes)

    def _find_reply_cut(self, lines: list[str]) -> int | None:
        for index, line in enumerate(lines):
            if any(pattern.match(line.strip()) for pattern in self.REPLY_SEPARATOR_PATTERNS):
                # Avoid cutting a top-level message that naturally starts with a header-like word.
                if index > 0:
                    return index
        return None

    def _strip_quoted_prefix_lines(self, lines: list[str]) -> list[str]:
        stripped: list[str] = []
        for line in lines:
            if self.QUOTED_LINE_PATTERN.match(line):
                continue
            stripped.append(line)
        return stripped

    def _strip_signature(self, lines: list[str]) -> tuple[list[str], bool]:
        search_start = max(0, len(lines) - 12)
        for index in range(len(lines) - 1, search_start - 1, -1):
            if any(pattern.match(lines[index].strip()) for pattern in self.SIGNATURE_PATTERNS):
                return lines[:index], True
        return lines, False
