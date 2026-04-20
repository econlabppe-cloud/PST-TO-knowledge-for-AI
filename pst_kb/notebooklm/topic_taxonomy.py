from __future__ import annotations

from pst_kb.notebooklm import topic_classifier as base


CURATED_TOPIC_RULES: tuple[base.TopicRule, ...] = (
    base.TopicRule(
        name="גמלאות, פרישה וזכאות",
        keywords=("גמלה", "גמלאות", "קצבה", "קצבאות", "פרישה", "פורש", "פורשים", "זכאות", "retirement", "eligibility"),
        subtopics=(
            base.SubtopicRule("זכאות לפרישה", ("זכאות", "פרישה", "retirement")),
            base.SubtopicRule("קצבה וגמלה", ("גמלה", "קצבה", "קצבאות")),
            base.SubtopicRule("תנאים והחלטות", ("תנאים", "החלטה", "החלטות", "קביעה")),
            base.SubtopicRule("מועדי תחולה", ("מועד", "מועדים", "תחולה", "effective")),
        ),
        priority=7,
    ),
    base.TopicRule(
        name="פנסיה צוברת, תקציבית ורציפות",
        keywords=("פנסיה", "פנסיוני", "קרן", "קרנות", "צוברת", "תקציבית", "רציפות", "העברה", "משיכת כספים", "pension"),
        subtopics=(
            base.SubtopicRule("פנסיה צוברת", ("פנסיה צוברת", "צוברת")),
            base.SubtopicRule("פנסיה תקציבית", ("פנסיה תקציבית", "תקציבית")),
            base.SubtopicRule("רציפות זכויות", ("רציפות", "רציפות זכויות")),
            base.SubtopicRule("מעבר בין מסלולים", ("העברה", "מעבר", "מסלול")),
            base.SubtopicRule("משיכת כספים", ("משיכת כספים", "משיכה")),
        ),
        priority=6,
    ),
    base.TopicRule(
        name="תחשיבים, שכר קובע ומקדמות",
        keywords=("תחשיב", "תחשיבים", "חישוב", "חישובי", "שכר קובע", "מקדמה", "מקדמות", "אחוז", "אחוזים", "סימולציה", "simulation", "calc", "יוקר"),
        subtopics=(
            base.SubtopicRule("שכר קובע", ("שכר קובע", "שכר")),
            base.SubtopicRule("אחוזים ויחסים", ("אחוז", "אחוזים", "percent")),
            base.SubtopicRule("מקדמות", ("מקדמה", "מקדמות")),
            base.SubtopicRule("סימולציות וחישובים", ("סימולציה", "simulation", "חישוב")),
        ),
        priority=8,
    ),
    base.TopicRule(
        name="פניות, בקשות ומסמכים",
        keywords=("פנייה", "פניות", "בקשה", "בקשות", "מסמכים", "בירור סטטוס", "עדכון פרטים", "מענה", "request", "inquiry", "status"),
        subtopics=(
            base.SubtopicRule("בקשות מסמכים", ("מסמכים", "בקשה", "בקשות")),
            base.SubtopicRule("בירור סטטוס", ("בירור", "סטטוס", "status")),
            base.SubtopicRule("עדכון פרטים", ("עדכון", "פרטים")),
            base.SubtopicRule("פניות גמלאים", ("גמלאים", "גמלאי")),
            base.SubtopicRule("פניות עובדים", ("עובדים", "עובד")),
        ),
        priority=6,
    ),
    base.TopicRule(
        name="מידע, הרשאות והעברת נתונים",
        keywords=("מידע", "נתונים", "הרשאה", "הרשאות", "גישה", "מאגר", "מאגרי מידע", "העברת מידע", "access", "permissions"),
        subtopics=(
            base.SubtopicRule("בקשות מידע", ("מידע", "נתונים")),
            base.SubtopicRule("הרשאות וגישה", ("הרשאה", "הרשאות", "גישה")),
            base.SubtopicRule("העברת נתונים", ("העברה", "העברת מידע", "נתונים")),
            base.SubtopicRule("מאגרי מידע", ("מאגר", "מאגרי מידע")),
        ),
        priority=5,
    ),
    base.TopicRule(
        name="ממשקים, דיווח ומרכבה",
        keywords=("ממשק", "ממשקים", "מרכבה", "דיווח", "דיווחים", "טפסים", "קליטה", "ייצוא", "bi", "sap", "ess", "interface", "report"),
        subtopics=(
            base.SubtopicRule("מרכבה", ("מרכבה", "merkava")),
            base.SubtopicRule("דיווחים", ("דיווח", "דיווחים", "report")),
            base.SubtopicRule("ממשקים", ("ממשק", "ממשקים", "interface")),
            base.SubtopicRule("טפסים וקליטה", ("טפסים", "קליטה", "ייצוא")),
            base.SubtopicRule("מערכות ותמיכה טכנית", ("bi", "sap", "ess", "system")),
        ),
        priority=6,
    ),
    base.TopicRule(
        name="הסכמים, מכרזים וספקים",
        keywords=("חוזה", "חוזים", "הסכם", "הסכמים", "התקשרות", "התקשרויות", "מכרז", "מכרזים", "ספק", "ספקים", "חשבונית", "invoice", "הזמנה"),
        subtopics=(
            base.SubtopicRule("מכרזים", ("מכרז", "מכרזים")),
            base.SubtopicRule("הסכמים", ("הסכם", "הסכמים")),
            base.SubtopicRule("ספקים וחשבוניות", ("ספק", "ספקים", "חשבונית", "invoice")),
            base.SubtopicRule("התקשרויות", ("התקשרות", "התקשרויות")),
            base.SubtopicRule("הזמנות", ("הזמנה", "הזמנות")),
        ),
        priority=5,
    ),
    base.TopicRule(
        name="תשלומים, חוב וניכויים",
        keywords=("תשלום", "תשלומים", "חוב", "חובות", "גבייה", "ניכוי", "ניכויים", "החזר", "החזרים", "שיפוי", "refund", "debit"),
        subtopics=(
            base.SubtopicRule("תשלומים", ("תשלום", "תשלומים")),
            base.SubtopicRule("חוב וגבייה", ("חוב", "חובות", "גבייה")),
            base.SubtopicRule("החזר ושיפוי", ("החזר", "החזרים", "שיפוי", "refund")),
            base.SubtopicRule("ניכויים", ("ניכוי", "ניכויים", "debit")),
        ),
        priority=5,
    ),
    base.TopicRule(
        name="תביעות, ערעורים והליכים משפטיים",
        keywords=("תביעה", "תביעות", "ערעור", "ערעורים", "הליך", "הליכים", "משפטי", "legal", "פסק דין", "court", "appeal"),
        subtopics=(
            base.SubtopicRule("ערעורים", ("ערעור", "ערעורים", "appeal")),
            base.SubtopicRule("תביעות", ("תביעה", "תביעות")),
            base.SubtopicRule("הליכים משפטיים", ("הליך", "הליכים", "משפטי", "legal")),
            base.SubtopicRule("פסקי דין", ("פסק דין", "court")),
        ),
        priority=6,
    ),
    base.TopicRule(
        name="שאירים, אלמנות ויתומים",
        keywords=("שאיר", "שאירים", "אלמנה", "אלמנת", "יתום", "יתומים", "בן זוג", "בת זוג", "survivor"),
        subtopics=(
            base.SubtopicRule("שאירים", ("שאיר", "שאירים")),
            base.SubtopicRule("אלמנות", ("אלמנה", "אלמנת")),
            base.SubtopicRule("יתומים", ("יתום", "יתומים")),
            base.SubtopicRule("בני זוג", ("בן זוג", "בת זוג")),
        ),
        priority=7,
    ),
    base.TopicRule(
        name="נהלים, הדרכות וישיבות",
        keywords=("נוהל", "נהלים", "הדרכה", "הדרכות", "פגישה", "פגישות", "ישיבה", "ישיבות", "מצגת", "מצגות", "סיכום דיון", "meeting", "training", "manual"),
        subtopics=(
            base.SubtopicRule("נהלים", ("נוהל", "נהלים")),
            base.SubtopicRule("הדרכות", ("הדרכה", "הדרכות", "training")),
            base.SubtopicRule("ישיבות ופגישות", ("פגישה", "פגישות", "ישיבה", "ישיבות", "meeting")),
            base.SubtopicRule("מצגות וסיכומים", ("מצגת", "מצגות", "סיכום דיון")),
        ),
        priority=2,
    ),
    base.TopicRule(
        name="בקרות, דוחות וביקורות",
        keywords=("בקרה", "בקרות", "ביקורת", "ביקורות", "דוח", "דוחות", "מעקב", "סטטוס", "audit", "report"),
        subtopics=(
            base.SubtopicRule("ביקורות", ("ביקורת", "ביקורות")),
            base.SubtopicRule("דוחות", ("דוח", "דוחות")),
            base.SubtopicRule("בקרות ומעקב", ("בקרה", "בקרות", "מעקב")),
            base.SubtopicRule("סטטוס ניהולי", ("סטטוס", "ניהולי")),
        ),
        priority=3,
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
    "אמת את זהותך",
    "תיבת הדואר שלך כמעט מלאה",
    "security alert",
    "sync issue",
    "sync issues",
    "calendar sync",
    "mail sync",
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
    blob = " ".join((subject, body[:800], folder, sender)).lower()
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

    return base.classify_email_record(
        row,
        topic_source=topic_source,
        rules=CURATED_TOPIC_RULES,
        min_score=min_score,
    )


def render_curated_rules_text() -> str:
    return base.render_rules_text(CURATED_TOPIC_RULES)


def _get_value(row: object, key: str, default: object = None) -> object:
    if hasattr(row, "get"):
        return row.get(key, default)
    return getattr(row, key, default)
