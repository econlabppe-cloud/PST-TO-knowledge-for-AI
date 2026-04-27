from __future__ import annotations

from pst_kb.notebooklm import topic_classifier as base


CURATED_TOPIC_RULES: tuple[base.TopicRule, ...] = (
    base.TopicRule(
        name="פניות גמלאים וזכאות",
        keywords=(
            "גמלאי",
            "גמלאים",
            "גימלה",
            "גמלה",
            "קצבה",
            "קצבאות",
            "זכאות",
            "פרישה",
            "פנסיה",
            "רשות השידור",
            "retiree",
            "retirement",
            "pension",
        ),
        subtopics=(
            base.SubtopicRule("בירור זכאות", ("זכאות", "זכאי", "פרישה", "retirement")),
            base.SubtopicRule("קצבה וגמלה", ("גמלה", "גימלה", "קצבה", "קצבאות")),
            base.SubtopicRule("פניות גמלאים כלליות", ("גמלאי", "גמלאים", "רשות השידור")),
        ),
        priority=9,
    ),
    base.TopicRule(
        name="מסמכים, אישורים וטפסים",
        keywords=(
            "מסמכים",
            "מסמך",
            "אישור",
            "אישורים",
            "טופס",
            "טפסים",
            "אישור מס",
            "טופס 106",
            "טופס 161",
            "אישור שנתי",
            "תלוש",
            "תלושים",
            "צילום תעודה",
            "צילום תז",
            "מסמכי",
            "documents",
            "certificate",
            "form",
        ),
        subtopics=(
            base.SubtopicRule("בקשת מסמכים", ("מסמכים", "מסמך", "documents")),
            base.SubtopicRule("אישורים שנתיים", ("אישור", "אישורים", "אישור שנתי", "certificate")),
            base.SubtopicRule("טפסי מס", ("106", "161", "אישור מס", "טופס")),
            base.SubtopicRule("תלושי גמלה", ("תלוש", "תלושים")),
        ),
        priority=8,
    ),
    base.TopicRule(
        name="תשלומים, בנק וניכויים",
        keywords=(
            "תשלום",
            "תשלומים",
            "בנק",
            "חשבון",
            "חשבון בנק",
            "העברה בנקאית",
            "ניכוי",
            "ניכויים",
            "מס הכנסה",
            "ביטוח לאומי",
            "החזר",
            "שיק",
            "יתרה",
            "payment",
            "bank",
            "refund",
        ),
        subtopics=(
            base.SubtopicRule("עדכון חשבון בנק", ("בנק", "חשבון בנק", "העברה בנקאית")),
            base.SubtopicRule("ניכויים ומסים", ("ניכוי", "ניכויים", "מס הכנסה", "ביטוח לאומי")),
            base.SubtopicRule("בירור תשלום", ("תשלום", "תשלומים", "payment")),
            base.SubtopicRule("החזרים", ("החזר", "refund")),
        ),
        priority=8,
    ),
    base.TopicRule(
        name="שכר קובע, חישובים ורצף שירות",
        keywords=(
            "שכר קובע",
            "חישוב",
            "חישובים",
            "תחשיב",
            "וותק",
            "ותק",
            "רצף",
            "רציפות",
            "תקופת עבודה",
            "אחוז",
            "אחוזים",
            "מקדמה",
            "calc",
            "simulation",
        ),
        subtopics=(
            base.SubtopicRule("שכר קובע", ("שכר קובע",)),
            base.SubtopicRule("חישובי גמלה", ("חישוב", "חישובים", "תחשיב", "simulation")),
            base.SubtopicRule("וותק ורצף שירות", ("וותק", "ותק", "רצף", "רציפות", "תקופת עבודה")),
            base.SubtopicRule("אחוזים ומקדמות", ("אחוז", "אחוזים", "מקדמה")),
        ),
        priority=7,
    ),
    base.TopicRule(
        name="עדכון פרטים אישיים",
        keywords=(
            "עדכון פרטים",
            "שינוי כתובת",
            "כתובת",
            "טלפון",
            "נייד",
            "מייל",
            "דואר אלקטרוני",
            "שם משפחה",
            "מספר זהות",
            "פרטים אישיים",
            "update details",
        ),
        subtopics=(
            base.SubtopicRule("כתובת וטלפון", ("כתובת", "טלפון", "נייד")),
            base.SubtopicRule("מייל ותקשורת", ("מייל", "דואר אלקטרוני")),
            base.SubtopicRule("פרטי זיהוי", ("מספר זהות", "שם משפחה", "פרטים אישיים")),
        ),
        priority=7,
    ),
    base.TopicRule(
        name="שאירים, פטירה והעברת זכויות",
        keywords=(
            "שאירים",
            "שאירים",
            "אלמנה",
            "אלמן",
            "בן זוג",
            "בת זוג",
            "פטירה",
            "נפטר",
            "יורשים",
            "עזבון",
            "survivor",
        ),
        subtopics=(
            base.SubtopicRule("זכויות שאירים", ("שאירים", "שארים", "survivor")),
            base.SubtopicRule("אלמן ואלמנה", ("אלמנה", "אלמן", "בן זוג", "בת זוג")),
            base.SubtopicRule("פטירה ויורשים", ("פטירה", "נפטר", "יורשים", "עזבון")),
        ),
        priority=7,
    ),
    base.TopicRule(
        name="ערעורים, תלונות וטיפול חריג",
        keywords=(
            "ערעור",
            "ערעורים",
            "תלונה",
            "תלונות",
            "בירור",
            "בקשה חריגה",
            "דחוף",
            "בעיה",
            "תקלה",
            "טעות",
            "complaint",
            "appeal",
        ),
        subtopics=(
            base.SubtopicRule("ערעורים", ("ערעור", "ערעורים", "appeal")),
            base.SubtopicRule("תלונות", ("תלונה", "תלונות", "complaint")),
            base.SubtopicRule("טיפול בבעיה", ("בעיה", "תקלה", "טעות", "דחוף")),
        ),
        priority=6,
    ),
    base.TopicRule(
        name="פניות כלליות ושירות",
        keywords=(
            "שלום",
            "מבקש",
            "מבקשת",
            "אבקש",
            "אשמח",
            "שאלה",
            "בירור",
            "פניה",
            "פנייה",
            "מידע",
            "עזרה",
            "request",
            "inquiry",
        ),
        subtopics=(
            base.SubtopicRule("בקשה כללית", ("מבקש", "מבקשת", "אבקש", "אשמח")),
            base.SubtopicRule("שאלה ובירור", ("שאלה", "בירור", "inquiry")),
            base.SubtopicRule("בקשת עזרה", ("עזרה", "מידע", "request")),
        ),
        priority=4,
    ),
)


CURATED_SYSTEM_TERMS: tuple[str, ...] = (
    "יומן סינכרון",
    "בעיות סינכרון",
    "סינכרון:",
    "mail delivery failed",
    "delivery status notification",
    "undelivered mail returned to sender",
    "returned to sender",
    "message could not be delivered",
    "delivery to the following recipient failed",
    "returned mail",
    "mailbox full",
    "automatic reply",
    "out of office",
    "security risk scan",
    "security alert",
    "sync issue",
    "sync issues",
    "calendar sync",
    "mail sync",
    "outlook for android",
)


def classify_email_record_corpus(
    row: object,
    *,
    topic_source: str = "rules",
    min_score: float = 4.0,
) -> base.TopicMatch:
    subject = str(_get_value(row, "subject", "") or "")
    body = str(_get_value(row, "clean_body", "") or _get_value(row, "body", "") or "")
    folder = str(_get_value(row, "folder_path", "") or "")
    sender = str(_get_value(row, "from_email", "") or "")
    container_name = str(_get_value(row, "container_attachment_filename", "") or "")
    blob = " ".join((subject, body[:1800], folder, sender, container_name)).lower()

    if any(term.lower() in blob for term in CURATED_SYSTEM_TERMS):
        return base.TopicMatch(
            topic="מערכת / רעש",
            subtopic="דואר מערכת וסינכרון",
            score=0.0,
            matched_terms=tuple(),
            is_system_noise=True,
            review_required=False,
            source_mode="rules",
        )

    match = base.classify_email_record(
        row,
        topic_source=topic_source,
        rules=CURATED_TOPIC_RULES,
        min_score=min_score,
    )

    if match.topic == "לבדיקה ידנית" and _looks_like_retiree_request(blob):
        return base.TopicMatch(
            topic="פניות כלליות ושירות",
            subtopic="פניית גמלאי כללית",
            score=4.5,
            matched_terms=("hr2", "iba.org.il"),
            review_required=False,
            source_mode="rules",
        )

    return match


def render_curated_rules_text() -> str:
    return base.render_rules_text(CURATED_TOPIC_RULES)


def _looks_like_retiree_request(blob: str) -> bool:
    domain_hints = ("hr2@iba.org.il", "iba.org.il", "רשות השידור", "גמלאי")
    request_hints = ("gmail.com", "walla.co.il", "yahoo.com", "מבקש", "מבקשת", "אבקש", "מסמכים")
    return any(term in blob for term in domain_hints) and any(term in blob for term in request_hints)


def _get_value(row: object, key: str, default: object = None) -> object:
    if hasattr(row, "get"):
        return row.get(key, default)
    return getattr(row, key, default)
