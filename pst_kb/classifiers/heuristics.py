from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass(frozen=True)
class ClassificationResult:
    possible_intent: str
    possible_topic: str
    complaint_indicator: bool
    request_indicator: bool
    urgency_indicator: bool
    has_question: bool
    likely_sender_type: str


class HeuristicClassifier:
    def __init__(
        self,
        internal_domains: list[str] | None = None,
        topic_keywords: dict[str, list[str]] | None = None,
        sender_type_keywords: dict[str, list[str]] | None = None,
    ) -> None:
        self.internal_domains = [domain.lower().lstrip("@") for domain in (internal_domains or [])]
        self.topic_keywords = topic_keywords or {}
        self.sender_type_keywords = sender_type_keywords or {}

    def classify(self, subject: str, body: str, sender_email: str | None) -> ClassificationResult:
        text = f"{subject}\n{body}".lower()
        complaint = _contains_any(
            text,
            [
                "complaint",
                "unhappy",
                "not acceptable",
                "dissatisfied",
                "disappointing",
                "בעיה",
                "תלונה",
                "לא תקין",
                "לא מרוצה",
                "מאוכזב",
                "מאוכזבת",
                "גרוע",
            ],
        )
        request = _contains_any(
            text,
            [
                "please",
                "request",
                "can you",
                "could you",
                "kindly",
                "אשמח",
                "מבקש",
                "מבקשת",
                "נא",
                "אפשר",
                "לתשומת לבך",
                "טיפולך",
            ],
        )
        urgent = _contains_any(
            text,
            [
                "urgent",
                "asap",
                "immediately",
                "priority",
                "critical",
                "דחוף",
                "בהקדם",
                "מיידי",
                "מיידית",
                "בהקדם האפשרי",
                "קריטי",
            ],
        )
        has_question = "?" in text or "؟" in text or _contains_any(text, ["האם", "מדוע", "למה", "איך", "מתי"])

        return ClassificationResult(
            possible_intent=self._infer_intent(complaint, request, urgent, has_question),
            possible_topic=self._infer_topic(text),
            complaint_indicator=complaint,
            request_indicator=request,
            urgency_indicator=urgent,
            has_question=has_question,
            likely_sender_type=self._infer_sender_type(text, sender_email),
        )

    def _infer_intent(self, complaint: bool, request: bool, urgent: bool, has_question: bool) -> str:
        if complaint:
            return "complaint"
        if urgent and request:
            return "urgent_request"
        if request:
            return "request"
        if has_question:
            return "question"
        return "unknown"

    def _infer_topic(self, text: str) -> str:
        for topic, keywords in self.topic_keywords.items():
            if _contains_any(text, keywords):
                return topic
        return "unknown"

    def _infer_sender_type(self, text: str, sender_email: str | None) -> str:
        domain = sender_email.split("@", 1)[-1].lower() if sender_email and "@" in sender_email else ""
        if domain and domain in self.internal_domains:
            return "internal_staff"
        for sender_type, keywords in self.sender_type_keywords.items():
            if _contains_any(text, keywords):
                return sender_type
        if _contains_any(text, ["retiree", "retired", "גמלאי", "פנסיונר"]):
            return "retiree"
        if _contains_any(text, ["daughter", "son", "spouse", "בת של", "בן של", "אשתו", "בעלה"]):
            return "family_member"
        if _contains_any(text, ["former employee", "ex employee", "עובד לשעבר", "עובדת לשעבר"]):
            return "former_employee"
        return "unknown"


def _contains_any(text: str, keywords: list[str]) -> bool:
    for keyword in keywords:
        keyword = keyword.lower().strip()
        if not keyword:
            continue
        if re.search(rf"(?<!\w){re.escape(keyword)}(?!\w)", text, flags=re.IGNORECASE):
            return True
    return False
