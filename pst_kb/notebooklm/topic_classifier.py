from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from pst_kb.normalizers import normalize_subject, normalize_whitespace


@dataclass(frozen=True)
class SubtopicRule:
    name: str
    keywords: tuple[str, ...]


@dataclass(frozen=True)
class TopicRule:
    name: str
    keywords: tuple[str, ...]
    subtopics: tuple[SubtopicRule, ...] = ()
    exclude_keywords: tuple[str, ...] = ()
    priority: int = 0


@dataclass(frozen=True)
class TopicMatch:
    topic: str
    subtopic: str
    score: float
    matched_terms: tuple[str, ...]
    is_system_noise: bool = False
    review_required: bool = False
    source_mode: str = "rules"


SYSTEM_SENDER_KEYWORDS = (
    "mailer-daemon",
    "postmaster",
    "noreply",
    "no-reply",
    "do-not-reply",
    "system",
    "bot",
    "notification",
    "notifications",
    "mrsysmail",
)

SYSTEM_SUBJECT_KEYWORDS = (
    "delivery status",
    "undeliverable",
    "failure notice",
    "automatic reply",
    "out of office",
    "sync issues",
    "returned to sender",
    "mail delivery failed",
    "mail could not be delivered",
    "message could not be delivered",
    "undelivered mail",
    "non-delivery",
    "delivery failure",
    "delivery to the following recipient failed",
    "returned mail",
    "יומן סינכרון",
    "מחוץ למשרד",
    "תשובה אוטומטית",
    "אישור מסירה",
    "הודעה לא נמסרה",
    "לא ניתן למסור",
    "חזר לשולח",
)

DEFAULT_TOPIC_RULES: tuple[TopicRule, ...] = (
    TopicRule(
        name="קרנות פנסיה ורציפות זכויות",
        keywords=(
            "קרן",
            "קרנות",
            "פנסיה",
            "פנסיוני",
            "ותיקה",
            "צוברת",
            "רציפות",
            "רכישת זכויות",
            "העברת זכויות",
            "רציפות זכויות",
            "פנסיה צוברת",
            "קרן ותיקה",
            "pension",
        ),
        subtopics=(
            SubtopicRule("רכישת זכויות", ("רכישת זכויות", "purchase rights")),
            SubtopicRule("העברת זכויות", ("העברת זכויות", "transfer rights", "רציפות זכויות")),
            SubtopicRule("קרן ותיקה", ("קרן ותיקה", "ותיקה")),
            SubtopicRule("פנסיה צוברת", ("פנסיה צוברת", "צוברת")),
        ),
        priority=6,
    ),
    TopicRule(
        name="זכאות לגמלה וקצבה",
        keywords=(
            "זכאות",
            "גמלה",
            "גמלאות",
            "קצבה",
            "קצבאות",
            "פרישה",
            "פנסיה תקציבית",
            "קצבת",
            "מועד זכאות",
            "החלטת זכאות",
            "eligibility",
            "retirement",
        ),
        subtopics=(
            SubtopicRule("תנאי זכאות", ("תנאי זכאות", "condition eligibility")),
            SubtopicRule("מועד זכאות", ("מועד זכאות", "date eligibility")),
            SubtopicRule("פרישה", ("פרישה", "retirement")),
            SubtopicRule("החלטת זכאות", ("החלטת זכאות", "decision eligibility")),
        ),
        priority=5,
    ),
    TopicRule(
        name="חישובי גמלה ושכר קובע",
        keywords=(
            "חישוב",
            "חישובי",
            "תחשיב",
            "משכורת קובעת",
            "שכר קובע",
            "נוסחת",
            "אחוז",
            "אחוזי",
            "סימולציה",
            "סימולציות",
            "חישוב גמלה",
            "calc",
        ),
        subtopics=(
            SubtopicRule("שכר קובע", ("שכר קובע", "משכורת קובעת")),
            SubtopicRule("תחשיב", ("תחשיב", "חישוב")),
            SubtopicRule("אחוזי משרה", ("אחוזי משרה", "percent")),
            SubtopicRule("סימולציה", ("סימולציה", "simulation")),
        ),
        priority=7,
    ),
    TopicRule(
        name="שאירים ויתומים",
        keywords=(
            "שאירים",
            "שאיר",
            "יתומים",
            "יתום",
            "אלמנה",
            "אלמן",
            "בן זוג",
            "בת זוג",
            "survivor",
        ),
        subtopics=(
            SubtopicRule("אלמנה ובן זוג", ("אלמנה", "אלמן", "בן זוג", "בת זוג")),
            SubtopicRule("יתומים", ("יתומים", "יתום")),
        ),
        priority=6,
    ),
    TopicRule(
        name="תביעות ערעורים ובירורים משפטיים",
        keywords=(
            "תביעה",
            "תביעות",
            "ערעור",
            "ערעורים",
            "פסק דין",
            "בית משפט",
            "יועץ משפטי",
            "חוות דעת משפטית",
            "הליך",
            "legal",
            "appeal",
        ),
        subtopics=(
            SubtopicRule("ערעור", ("ערעור", "ערעורים", "appeal")),
            SubtopicRule("תביעה", ("תביעה", "תביעות")),
            SubtopicRule("בירור משפטי", ("בית משפט", "יועץ משפטי", "חוות דעת משפטית", "legal")),
        ),
        priority=6,
    ),
    TopicRule(
        name="חוזים התקשרויות וספקים",
        keywords=(
            "חוזה",
            "חוזים",
            "הסכם",
            "התקשרות",
            "התקשרויות",
            "ספק",
            "ספקים",
            "מכרז",
            "הזמנה",
            "חשבונית",
            "contract",
            "supplier",
            "tender",
        ),
        subtopics=(
            SubtopicRule("מכרזים", ("מכרז", "tender")),
            SubtopicRule("הסכמים", ("חוזה", "חוזים", "הסכם", "contract")),
            SubtopicRule("ספקים וחשבוניות", ("ספק", "ספקים", "חשבונית", "supplier", "invoice")),
        ),
        priority=5,
    ),
    TopicRule(
        name="תשלומים גבייה והתחשבנות",
        keywords=(
            "תשלום",
            "תשלומים",
            "גבייה",
            "ניכוי",
            "החזר",
            "שיפוי",
            "התחשבנות",
            "חוב",
            "חיוב",
            "payment",
            "refund",
            "debit",
            "invoice",
        ),
        subtopics=(
            SubtopicRule("החזר ושיפוי", ("החזר", "שיפוי", "refund")),
            SubtopicRule("גבייה וחיוב", ("גבייה", "חיוב", "debit")),
            SubtopicRule("ניכויים", ("ניכוי", "ניכויים")),
            SubtopicRule("תשלומים", ("תשלום", "תשלומים", "payment")),
        ),
        priority=4,
    ),
    TopicRule(
        name="פניות עובדים ומבוטחים",
        keywords=(
            "פנייה",
            "פניות",
            "מבוטח",
            "מבוטחים",
            "עובד",
            "עובדים",
            "גמלאי",
            "גמלאים",
            "בירור",
            "בקשה",
            "שאלה",
            "request",
            "question",
            "inquiry",
        ),
        subtopics=(
            SubtopicRule("בירור סטטוס", ("בירור", "סטטוס", "status")),
            SubtopicRule("בקשת מסמכים", ("בקשה", "מסמכים", "request")),
            SubtopicRule("פנייה כללית", ("פנייה", "פניות", "question", "inquiry")),
            SubtopicRule("עדכון פרטים", ("עדכון", "פרטים", "update")),
        ),
        priority=2,
    ),
    TopicRule(
        name="מערכות דיווח ומרכבה",
        keywords=(
            "דיווח",
            "דיווחים",
            "מרכבה",
            "מערכת",
            "מערכות",
            "טופס",
            "טפסים",
            "ממשק",
            "הרשאה",
            "הרשאות",
            "קליטה",
            "שגיאה",
            "error",
            "system",
            "interface",
            "report",
        ),
        subtopics=(
            SubtopicRule("מרכבה", ("מרכבה", "merkava")),
            SubtopicRule("דיווחים", ("דיווח", "דיווחים", "report")),
            SubtopicRule("טפסים וממשקים", ("טופס", "טפסים", "ממשק", "interface")),
            SubtopicRule("שגיאות וקליטה", ("שגיאה", "קליטה", "error")),
        ),
        priority=4,
    ),
    TopicRule(
        name="נהלים הדרכות ועבודה פנימית",
        keywords=(
            "נוהל",
            "נהלים",
            "הדרכה",
            "הדרכות",
            "מצגת",
            "ישיבה",
            "פגישה",
            "סיכום דיון",
            "הנחיה",
            "עדכון פנימי",
            "manual",
            "training",
            "meeting",
        ),
        subtopics=(
            SubtopicRule("הדרכה", ("הדרכה", "הדרכות", "training")),
            SubtopicRule("פגישה וישיבה", ("ישיבה", "פגישה", "meeting")),
            SubtopicRule("נוהל", ("נוהל", "נהלים", "manual")),
        ),
        priority=1,
    ),
)


def classify_email_record(
    row: object,
    *,
    topic_source: str = "rules",
    rules: tuple[TopicRule, ...] = DEFAULT_TOPIC_RULES,
    min_score: float = 4.0,
) -> TopicMatch:
    getter = row.get if hasattr(row, "get") else lambda key, default=None: getattr(row, key, default)

    subject_raw = str(getter("subject", "") or "")
    subject = _normalize_search_text(normalize_subject(subject_raw))
    body_raw = str(getter("clean_body", "") or getter("body", "") or "")
    body = _normalize_search_text(body_raw)
    folder = _normalize_search_text(str(getter("folder_path", "") or ""))
    sender = _normalize_search_text(str(getter("from_email", "") or ""))
    llm_topic = str(getter("llm_topic", "") or "").strip()
    llm_subtopic = str(getter("llm_subtopic", "") or "").strip()
    llm_confidence = _as_float(getter("llm_confidence", None))
    cluster_name = str(getter("cluster_name", "") or "").strip()
    combined = " ".join(part for part in (subject, body[:6000], folder, sender) if part)

    if _is_system_noise(subject=subject, body=body, folder=folder, sender=sender):
        return TopicMatch(
            topic="מערכת / רעש",
            subtopic="הודעות מערכת",
            score=0.0,
            matched_terms=tuple(),
            is_system_noise=True,
            review_required=False,
            source_mode="rules",
        )

    if topic_source == "llm":
        topic = _normalize_topic_label(llm_topic) or _normalize_topic_label(cluster_name) or "לבדיקה ידנית"
        subtopic = llm_subtopic or _subject_snippet(subject_raw)
        return TopicMatch(topic=topic, subtopic=subtopic, score=1.0, matched_terms=tuple(), source_mode="llm")

    if topic_source == "cluster":
        topic = _normalize_topic_label(cluster_name) or _normalize_topic_label(llm_topic) or "לבדיקה ידנית"
        subtopic = llm_subtopic or _subject_snippet(subject_raw)
        return TopicMatch(topic=topic, subtopic=subtopic, score=1.0, matched_terms=tuple(), source_mode="cluster")

    if topic_source == "hybrid" and llm_topic and llm_confidence is not None and llm_confidence >= 0.75:
        topic = _normalize_topic_label(llm_topic)
        if topic:
            subtopic = llm_subtopic or _subject_snippet(subject_raw)
            return TopicMatch(topic=topic, subtopic=subtopic, score=llm_confidence, matched_terms=tuple(), source_mode="llm")

    best_rule: TopicRule | None = None
    best_score = float("-inf")
    best_terms: list[str] = []
    best_subtopic = ""

    for rule in rules:
        score, matched_terms, subtopic = _score_rule(rule, subject=subject, body=body, folder=folder, combined=combined)
        if score > best_score or (score == best_score and rule.priority > (best_rule.priority if best_rule else -999)):
            best_rule = rule
            best_score = score
            best_terms = matched_terms
            best_subtopic = subtopic

    if best_rule is None or best_score < min_score:
        return TopicMatch(
            topic="לבדיקה ידנית",
            subtopic=llm_subtopic or _fallback_subtopic(subject_raw, best_terms),
            score=max(best_score, 0.0),
            matched_terms=tuple(best_terms),
            review_required=True,
            source_mode="rules",
        )

    return TopicMatch(
        topic=best_rule.name,
        subtopic=best_subtopic or llm_subtopic or _fallback_subtopic(subject_raw, best_terms),
        score=best_score,
        matched_terms=tuple(best_terms),
        review_required=False,
        source_mode="rules",
    )


def render_rules_text(rules: tuple[TopicRule, ...] = DEFAULT_TOPIC_RULES) -> str:
    lines = [
        "CLASSIFICATION RULES",
        "",
        "The pack uses transparent keyword rules first.",
        "The rules prefer subject matches, then body matches, then folder matches.",
        "",
    ]
    for rule in rules:
        lines.append(f"TOPIC: {rule.name}")
        lines.append(f"Priority: {rule.priority}")
        lines.append("Keywords:")
        for keyword in rule.keywords:
            lines.append(f"  - {keyword}")
        if rule.subtopics:
            lines.append("Subtopics:")
            for subtopic in rule.subtopics:
                joined = "; ".join(subtopic.keywords)
                lines.append(f"  - {subtopic.name}: {joined}")
        if rule.exclude_keywords:
            lines.append("Exclusions:")
            for keyword in rule.exclude_keywords:
                lines.append(f"  - {keyword}")
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def _score_rule(
    rule: TopicRule,
    *,
    subject: str,
    body: str,
    folder: str,
    combined: str,
) -> tuple[float, list[str], str]:
    score = float(rule.priority)
    matched_terms: list[str] = []
    best_subtopic = ""
    best_subtopic_score = float("-inf")

    for keyword in rule.keywords:
        keyword_norm = _normalize_search_text(keyword)
        if not keyword_norm:
            continue
        keyword_score = 0.0
        if keyword_norm in subject:
            keyword_score += 7.0 + (1.0 if " " in keyword_norm else 0.0)
        if keyword_norm in folder:
            keyword_score += 4.5
        if keyword_norm in body:
            keyword_score += 2.5
        if keyword_norm in combined:
            keyword_score += 0.5
        if keyword_score > 0:
            score += keyword_score
            matched_terms.append(keyword)

    for exclusion in rule.exclude_keywords:
        exclusion_norm = _normalize_search_text(exclusion)
        if exclusion_norm and exclusion_norm in combined:
            score -= 4.0

    for subtopic in rule.subtopics:
        subtopic_score = 0.0
        for keyword in subtopic.keywords:
            keyword_norm = _normalize_search_text(keyword)
            if not keyword_norm:
                continue
            if keyword_norm in subject:
                subtopic_score += 5.0
            if keyword_norm in folder:
                subtopic_score += 2.0
            if keyword_norm in body:
                subtopic_score += 1.5
        if subtopic_score > best_subtopic_score:
            best_subtopic_score = subtopic_score
            best_subtopic = subtopic.name if subtopic_score > 0 else ""

    return score, matched_terms, best_subtopic


def _normalize_topic_label(value: str) -> str:
    return normalize_whitespace(value).strip()


def _normalize_search_text(value: str) -> str:
    text = normalize_whitespace(str(value or "")).lower()
    if not text:
        return ""
    import re

    text = re.sub(r"[^\w\s\u0590-\u05FF]+", " ", text, flags=re.UNICODE)
    text = re.sub(r"\s+", " ", text, flags=re.UNICODE).strip()
    return text


def _subject_snippet(subject: str, max_length: int = 120) -> str:
    cleaned = normalize_subject(subject)
    cleaned = normalize_whitespace(cleaned)
    return cleaned[:max_length] or "ללא כותרת"


def _fallback_subtopic(subject: str, matched_terms: list[str]) -> str:
    if matched_terms:
        candidates = sorted((term.strip() for term in matched_terms if term.strip()), key=len, reverse=True)
        if candidates:
            return candidates[0][:80]
    snippet = _subject_snippet(subject, max_length=80)
    words = snippet.split()
    if len(words) > 6:
        snippet = " ".join(words[:6])
    return snippet or "ללא תת-נושא"


def _is_system_noise(*, subject: str, body: str, folder: str, sender: str) -> bool:
    blob = " ".join((subject, body[:400], folder, sender))
    return any(keyword in blob for keyword in SYSTEM_SENDER_KEYWORDS) or any(
        keyword.lower() in blob for keyword in SYSTEM_SUBJECT_KEYWORDS
    )


def _as_float(value: object) -> float | None:
    try:
        if value is None or str(value).strip() == "":
            return None
        return float(value)
    except (TypeError, ValueError):
        return None
