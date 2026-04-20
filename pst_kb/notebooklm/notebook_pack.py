from __future__ import annotations

import argparse
import logging
import re
from collections import Counter
from pathlib import Path

import pandas as pd
from tqdm import tqdm

from pst_kb.notebooklm.build_notebooks import select_representatives
from pst_kb.notebooklm.common import configure_script_logging
from pst_kb.utils.files import sanitize_filename

logger = logging.getLogger(__name__)

REQUIRED_COLUMNS = [
    "message_id",
    "date",
    "subject",
    "clean_body",
    "body",
    "from_email",
    "to_emails",
    "cc_emails",
    "bcc_emails",
    "folder_path",
    "cluster_name",
    "llm_topic",
    "llm_subtopic",
    "llm_tags",
    "llm_confidence",
    "is_filtered",
    "filter_reason",
]

SYSTEM_SENDER_RE = re.compile(
    r"(mailer-daemon|postmaster|mrsysmail|no-?reply|noreply|do-?not-?reply|system|bot|notification)",
    re.IGNORECASE,
)
SYSTEM_SUBJECT_RE = re.compile(
    r"(delivery status|undeliverable|failure notice|automatic reply|out of office|sync issues|יומן סינכרון|מחוץ למשרד|תשובה אוטומטית)",
    re.IGNORECASE,
)

KNOWLEDGE_TAXONOMY: dict[str, list[str]] = {
    "קרנות פנסיה ורציפות זכויות": [
        "קרן",
        "קרנות",
        "פנסיה",
        "ותיקה",
        "ותיקות",
        "צוברת",
        "רציפות",
        "רכישת זכויות",
        "העברת זכויות",
    ],
    "זכאות לגמלה וקצבה": [
        "זכאות",
        "גמלה",
        "גמלאות",
        "קצבה",
        "קצבאות",
        "פרישה",
        "פנסיה תקציבית",
        "מועד זכאות",
    ],
    "חישובי גמלה ושכר קובע": [
        "חישוב",
        "חישובי",
        "שכר קובע",
        "משכורת קובעת",
        "אחוז",
        "אחוזי",
        "סימולציה",
        "תחשיב",
    ],
    "שאירים ויתומים": [
        "שאירים",
        "שאיר",
        "יתומים",
        "יתום",
        "אלמנה",
        "אלמן",
        "בן זוג",
        "בת זוג",
    ],
    "תביעות ערעורים ובירורים משפטיים": [
        "תביעה",
        "תביעות",
        "ערעור",
        "ערעורים",
        "פסק דין",
        "בית משפט",
        "יועץ משפטי",
        "חוות דעת",
        "הליך",
    ],
    "חוזים התקשרויות וספקים": [
        "חוזה",
        "חוזים",
        "התקשרות",
        "התקשרויות",
        "ספק",
        "ספקים",
        "מכרז",
        "הזמנה",
        "חשבונית",
    ],
    "תשלומים גבייה והתחשבנות": [
        "תשלום",
        "תשלומים",
        "גבייה",
        "ניכוי",
        "החזר",
        "שיפוי",
        "התחשבנות",
        "חוב",
        "חיוב",
    ],
    "פניות עובדים ומבוטחים": [
        "פנייה",
        "פניות",
        "מבוטח",
        "מבוטחים",
        "עובד",
        "עובדים",
        "גמלאי",
        "גמלאים",
        "בירור",
    ],
    "מערכות דיווח ומרכבה": [
        "דיווח",
        "דיווחים",
        "מרכבה",
        "מערכת",
        "טופס",
        "קובץ",
        "ממשק",
        "קליטה",
        "הרשאה",
    ],
    "נהלים הדרכות ועבודה פנימית": [
        "נוהל",
        "נהלים",
        "הדרכה",
        "מצגת",
        "ישיבה",
        "פגישה",
        "סיכום דיון",
        "הנחיה",
    ],
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build an AI-ready NotebookLM knowledge pack from clustered/tagged emails.")
    parser.add_argument("--input-csv", type=Path, default=Path("data/intermediate/emails_clustered.csv"))
    parser.add_argument("--output-dir", type=Path, default=Path("data/output/notebooklm_pack"))
    parser.add_argument("--topic-source", choices=["auto", "llm", "taxonomy", "cluster"], default="auto")
    parser.add_argument("--max-source-chars", type=int, default=180_000)
    parser.add_argument("--max-emails-per-file", type=int, default=80)
    parser.add_argument("--max-body-chars", type=int, default=9_000)
    parser.add_argument("--top-people", type=int, default=60)
    parser.add_argument("--include-filtered", action="store_true")
    parser.add_argument("--include-system", action="store_true")
    parser.add_argument("--log-file", type=Path)
    parser.add_argument("--log-level", default="INFO")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    configure_script_logging(args.log_level, args.log_file)
    build_notebooklm_pack(
        input_csv=args.input_csv,
        output_dir=args.output_dir,
        topic_source=args.topic_source,
        max_source_chars=args.max_source_chars,
        max_emails_per_file=args.max_emails_per_file,
        max_body_chars=args.max_body_chars,
        top_people=args.top_people,
        include_filtered=args.include_filtered,
        include_system=args.include_system,
    )
    return 0


def build_notebooklm_pack(
    *,
    input_csv: Path,
    output_dir: Path,
    topic_source: str = "auto",
    max_source_chars: int = 180_000,
    max_emails_per_file: int = 80,
    max_body_chars: int = 9_000,
    top_people: int = 60,
    include_filtered: bool = False,
    include_system: bool = False,
) -> dict[str, Path]:
    df = _load(input_csv)
    df = _prepare(df, topic_source=topic_source)

    excluded_mask = pd.Series(False, index=df.index)
    if not include_filtered:
        excluded_mask |= df["is_filtered"].astype(str).str.lower() == "true"
    if not include_system:
        excluded_mask |= df["is_system_noise"]
        excluded_mask |= df["knowledge_topic"].astype(str).str.lower() == "system"

    included = df[~excluded_mask].copy()
    excluded = df[excluded_mask].copy()
    included = included[included["knowledge_topic"].astype(str).str.strip() != ""].copy()
    included = included.sort_values(["knowledge_topic", "date_sort", "message_id"], na_position="last")

    output_dir.mkdir(parents=True, exist_ok=True)
    sources_dir = output_dir / "sources"
    audit_dir = output_dir / "audit"
    sources_dir.mkdir(parents=True, exist_ok=True)
    audit_dir.mkdir(parents=True, exist_ok=True)

    manifest_rows: list[dict[str, object]] = []
    topic_rows: list[dict[str, object]] = []
    source_paths: list[Path] = []

    for topic, group in tqdm(list(included.groupby("knowledge_topic")), desc="Writing NotebookLM pack", unit="topic"):
        topic_paths = _write_topic_sources(
            topic=str(topic),
            group=group,
            sources_dir=sources_dir,
            max_source_chars=max_source_chars,
            max_emails_per_file=max_emails_per_file,
            max_body_chars=max_body_chars,
            manifest_rows=manifest_rows,
        )
        source_paths.extend(topic_paths)
        topic_rows.append(_topic_summary_row(str(topic), group, topic_paths, output_dir))

    manifest = pd.DataFrame(manifest_rows)
    topics = pd.DataFrame(topic_rows).sort_values("email_count", ascending=False) if topic_rows else pd.DataFrame()
    people = _people_summary(included, top_people=top_people)
    review = _review_queue(included)

    manifest_path = output_dir / "records_manifest.csv"
    topics_path = output_dir / "topics_index.csv"
    people_path = output_dir / "people_index.csv"
    review_path = output_dir / "review_queue.csv"
    excluded_path = audit_dir / "excluded_records.csv"

    manifest.to_csv(manifest_path, index=False, encoding="utf-8-sig")
    topics.to_csv(topics_path, index=False, encoding="utf-8-sig")
    people.to_csv(people_path, index=False, encoding="utf-8-sig")
    review.to_csv(review_path, index=False, encoding="utf-8-sig")
    excluded.drop(columns=["date_sort"], errors="ignore").to_csv(excluded_path, index=False, encoding="utf-8-sig")

    index_path = output_dir / "00_INDEX.md"
    guide_path = output_dir / "00_NOTEBOOKLM_GUIDE.md"
    topic_map_path = output_dir / "01_TOPIC_MAP.md"
    people_map_path = output_dir / "02_PEOPLE_MAP.md"
    review_md_path = output_dir / "03_REVIEW_QUEUE.md"

    index_path.write_text(_render_index(topics, manifest), encoding="utf-8-sig")
    guide_path.write_text(_render_guide(input_csv, included, excluded, source_paths, topic_source), encoding="utf-8-sig")
    topic_map_path.write_text(_render_topic_map(topics), encoding="utf-8-sig")
    people_map_path.write_text(_render_people_map(people), encoding="utf-8-sig")
    review_md_path.write_text(_render_review_markdown(review), encoding="utf-8-sig")

    logger.info("NotebookLM pack written to %s", output_dir)
    return {
        "index": index_path,
        "guide": guide_path,
        "topic_map": topic_map_path,
        "people_map": people_map_path,
        "review_queue": review_path,
        "manifest": manifest_path,
        "topics": topics_path,
        "people": people_path,
        "excluded": excluded_path,
    }


def _load(input_csv: Path) -> pd.DataFrame:
    df = pd.read_csv(input_csv, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    for column in REQUIRED_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    return df


def _prepare(df: pd.DataFrame, topic_source: str) -> pd.DataFrame:
    df = df.copy()
    df["date_sort"] = pd.to_datetime(df.get("date", ""), errors="coerce", utc=True)
    df["body_for_pack"] = df.apply(lambda row: str(row.get("clean_body") or row.get("body") or ""), axis=1)
    df["is_system_noise"] = df.apply(_is_system_noise, axis=1)
    df["knowledge_topic"] = df.apply(lambda row: _knowledge_topic(row, topic_source=topic_source), axis=1)
    df["knowledge_subtopic"] = df.apply(lambda row: _knowledge_subtopic(row), axis=1)
    df["knowledge_tags"] = df.apply(lambda row: _knowledge_tags(row), axis=1)
    return df


def _write_topic_sources(
    *,
    topic: str,
    group: pd.DataFrame,
    sources_dir: Path,
    max_source_chars: int,
    max_emails_per_file: int,
    max_body_chars: int,
    manifest_rows: list[dict[str, object]],
) -> list[Path]:
    topic_slug = sanitize_filename(topic.replace(" ", "_"), fallback="topic", max_length=90)
    chunks: list[list[tuple[pd.Series, str]]] = []
    current: list[tuple[pd.Series, str]] = []
    current_chars = 0

    ordered = group.sort_values(["date_sort", "message_id"], na_position="last")
    for _, row in ordered.iterrows():
        rendered = _render_email(row, max_body_chars=max_body_chars)
        would_exceed_chars = current and current_chars + len(rendered) > max_source_chars
        would_exceed_count = current and len(current) >= max_emails_per_file
        if would_exceed_chars or would_exceed_count:
            chunks.append(current)
            current = []
            current_chars = 0
        current.append((row, rendered))
        current_chars += len(rendered)
    if current:
        chunks.append(current)

    paths: list[Path] = []
    for part_number, chunk in enumerate(chunks, start=1):
        filename = f"{topic_slug}__part_{part_number:03d}.txt"
        path = sources_dir / filename
        chunk_df = pd.DataFrame([row.to_dict() for row, _ in chunk])
        path.write_text(
            _render_source_file(topic=topic, part_number=part_number, total_parts=len(chunks), group=chunk_df, emails=chunk),
            encoding="utf-8-sig",
        )
        rel_path = f"sources/{filename}"
        for row, _ in chunk:
            manifest_rows.append(
                {
                    "source_file": rel_path,
                    "knowledge_topic": topic,
                    "knowledge_subtopic": row.get("knowledge_subtopic", ""),
                    "message_id": row.get("message_id", ""),
                    "date": row.get("date", ""),
                    "from_email": row.get("from_email", ""),
                    "subject": row.get("subject", ""),
                    "folder_path": row.get("folder_path", ""),
                    "cluster_name": row.get("cluster_name", ""),
                    "llm_topic": row.get("llm_topic", ""),
                    "llm_confidence": row.get("llm_confidence", ""),
                }
            )
        paths.append(path)
    return paths


def _render_source_file(
    *,
    topic: str,
    part_number: int,
    total_parts: int,
    group: pd.DataFrame,
    emails: list[tuple[pd.Series, str]],
) -> str:
    start, end = _date_range(group)
    top_senders = "; ".join(group["from_email"].astype(str).value_counts().head(8).index.tolist())
    lines = [
        f"# NotebookLM Source: {topic}",
        "",
        f"Topic: {topic}",
        f"Part: {part_number} of {total_parts}",
        f"Email count in this file: {len(group)}",
        f"Date range: {start} to {end}",
        f"Top senders: {top_senders}",
        "",
        "Use this source as an evidence archive. Each email has metadata followed by the cleaned body.",
        "When answering, prefer citing the subject, date, sender, and message id.",
        "",
        "=====================================",
        "",
    ]
    for _, rendered in emails:
        lines.append(rendered)
    return "\n".join(lines).strip() + "\n"


def _render_email(row: pd.Series, max_body_chars: int) -> str:
    body = str(row.get("body_for_pack") or "")
    if len(body) > max_body_chars:
        body = body[:max_body_chars].rstrip() + "\n[body truncated for NotebookLM source sizing]"
    lines = [
        "## EMAIL",
        f"Message ID: {row.get('message_id', '')}",
        f"Date: {_format_date(row.get('date_sort'))}",
        f"From: {row.get('from_email', '')}",
        f"To: {row.get('to_emails', '')}",
        f"CC: {row.get('cc_emails', '')}",
        f"Subject: {row.get('subject', '')}",
        f"Folder: {row.get('folder_path', '')}",
        f"Knowledge topic: {row.get('knowledge_topic', '')}",
        f"Knowledge subtopic: {row.get('knowledge_subtopic', '')}",
        f"Tags: {row.get('knowledge_tags', '')}",
        f"Original cluster: {row.get('cluster_name', '')}",
        "",
        "Body:",
        body,
        "",
        "=====================================",
        "",
    ]
    return "\n".join(lines)


def _topic_summary_row(topic: str, group: pd.DataFrame, paths: list[Path], output_dir: Path) -> dict[str, object]:
    representatives = select_representatives(group, limit=3)
    subjects = " | ".join(str(row.get("subject", "")) for row in representatives)
    top_people = "; ".join(group["from_email"].astype(str).value_counts().head(8).index.tolist())
    start, end = _date_range(group)
    return {
        "knowledge_topic": topic,
        "email_count": len(group),
        "source_file_count": len(paths),
        "first_date": start,
        "last_date": end,
        "top_people": top_people,
        "representative_subjects": subjects,
        "source_files": "; ".join(path.relative_to(output_dir).as_posix() for path in paths),
    }


def _people_summary(df: pd.DataFrame, top_people: int) -> pd.DataFrame:
    rows = []
    for person, group in df.groupby("from_email", dropna=False):
        person = str(person).strip()
        if not person:
            continue
        rows.append(
            {
                "person": person,
                "email_count": len(group),
                "first_date": _format_date(group["date_sort"].min()),
                "last_date": _format_date(group["date_sort"].max()),
                "top_topics": "; ".join(group["knowledge_topic"].astype(str).value_counts().head(6).index.tolist()),
            }
        )
    return pd.DataFrame(rows).sort_values("email_count", ascending=False).head(top_people)


def _review_queue(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df.iterrows():
        reasons = []
        confidence = _as_float(row.get("llm_confidence"))
        if row.get("llm_topic") and confidence is not None and confidence < 0.65:
            reasons.append("low_llm_confidence")
        if not str(row.get("subject", "")).strip():
            reasons.append("missing_subject")
        if len(str(row.get("body_for_pack", ""))) < 200:
            reasons.append("short_body")
        if reasons:
            rows.append(
                {
                    "review_reason": "; ".join(reasons),
                    "knowledge_topic": row.get("knowledge_topic", ""),
                    "message_id": row.get("message_id", ""),
                    "date": row.get("date", ""),
                    "from_email": row.get("from_email", ""),
                    "subject": row.get("subject", ""),
                }
            )
    return pd.DataFrame(rows)


def _render_index(topics: pd.DataFrame, manifest: pd.DataFrame) -> str:
    lines = [
        "# NotebookLM Knowledge Pack Index",
        "",
        "This index organizes the email archive into topic-sized source files.",
        "Upload the guide, topic map, people map, and the source files that match the question area.",
        "",
        f"Total source files: {manifest['source_file'].nunique() if not manifest.empty else 0}",
        f"Total included emails: {len(manifest)}",
        "",
        "## Topics",
    ]
    for _, row in topics.iterrows():
        lines.extend(
            [
                f"### {row['knowledge_topic']}",
                f"- Emails: {row['email_count']}",
                f"- Source files: {row['source_file_count']}",
                f"- Date range: {row['first_date']} to {row['last_date']}",
                f"- Top people: {row['top_people']}",
                f"- Representative subjects: {row['representative_subjects']}",
                f"- Files: {row['source_files']}",
                "",
            ]
        )
    return "\n".join(lines).strip() + "\n"


def _render_guide(input_csv: Path, included: pd.DataFrame, excluded: pd.DataFrame, source_paths: list[Path], topic_source: str) -> str:
    return f"""# NotebookLM Guide

Purpose: this folder is a curated knowledge pack from the PST email archive.

Recommended upload order:
1. `00_NOTEBOOKLM_GUIDE.md`
2. `00_INDEX.md`
3. `01_TOPIC_MAP.md`
4. `02_PEOPLE_MAP.md`
5. Relevant files from `sources/`

How to ask questions:
- Ask by topic, person, date range, or message subject.
- Ask NotebookLM to cite email date, sender, subject, and message id.
- For broad questions, load the topic map first and then add only the relevant source files.

Processing summary:
- Input CSV: `{input_csv}`
- Included emails: {len(included)}
- Excluded or noisy records: {len(excluded)}
- Source files generated: {len(source_paths)}
- Topic source mode: `{topic_source}`

Notes:
- `knowledge_topic` uses `llm_topic` when present, otherwise the local pension-domain taxonomy and original cluster names are used.
- System mail, delivery failures, and sync logs are excluded from source files by default and kept in `audit/excluded_records.csv`.
- Original email metadata is preserved in `records_manifest.csv`.
"""


def _render_topic_map(topics: pd.DataFrame) -> str:
    lines = ["# Topic Map", ""]
    for _, row in topics.iterrows():
        lines.append(f"## {row['knowledge_topic']}")
        lines.append(f"- Emails: {row['email_count']}")
        lines.append(f"- Dates: {row['first_date']} to {row['last_date']}")
        lines.append(f"- People: {row['top_people']}")
        lines.append(f"- Representative subjects: {row['representative_subjects']}")
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def _render_people_map(people: pd.DataFrame) -> str:
    lines = ["# People Map", ""]
    for _, row in people.iterrows():
        lines.append(f"## {row['person']}")
        lines.append(f"- Emails: {row['email_count']}")
        lines.append(f"- Dates: {row['first_date']} to {row['last_date']}")
        lines.append(f"- Main topics: {row['top_topics']}")
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def _render_review_markdown(review: pd.DataFrame) -> str:
    lines = ["# Human Review Queue", ""]
    if review.empty:
        lines.append("No high-priority review records were detected.")
    else:
        for _, row in review.head(200).iterrows():
            lines.append(f"- {row['review_reason']} | {row['date']} | {row['from_email']} | {row['subject']}")
    return "\n".join(lines).strip() + "\n"


def _knowledge_topic(row: pd.Series, topic_source: str) -> str:
    llm_topic = str(row.get("llm_topic", "")).strip()
    cluster = str(row.get("cluster_name", "")).strip()
    if topic_source in ("auto", "llm") and llm_topic and llm_topic.lower() not in {"unknown", "nan"}:
        return llm_topic
    if topic_source == "llm":
        return cluster or "לא מסווג"
    if topic_source in ("auto", "taxonomy"):
        taxonomy_topic = _taxonomy_topic(row)
        if taxonomy_topic:
            return taxonomy_topic
    return cluster or "לא מסווג"


def _taxonomy_topic(row: pd.Series) -> str:
    text = " ".join(
        str(row.get(column, ""))
        for column in ("subject", "body_for_pack", "folder_path", "cluster_name")
    ).lower()
    scores: Counter[str] = Counter()
    for topic, terms in KNOWLEDGE_TAXONOMY.items():
        for term in terms:
            term_lower = term.lower()
            if term_lower in text:
                scores[topic] += 3 if term_lower in str(row.get("subject", "")).lower() else 1
    if not scores:
        return ""
    topic, score = scores.most_common(1)[0]
    return topic if score > 0 else ""


def _knowledge_subtopic(row: pd.Series) -> str:
    llm_subtopic = str(row.get("llm_subtopic", "")).strip()
    if llm_subtopic:
        return llm_subtopic
    subject = str(row.get("subject", "")).strip()
    return subject[:120]


def _knowledge_tags(row: pd.Series) -> str:
    llm_tags = str(row.get("llm_tags", "")).strip()
    if llm_tags:
        return llm_tags
    text = f"{row.get('subject', '')} {row.get('body_for_pack', '')}".lower()
    tags = []
    for topic, terms in KNOWLEDGE_TAXONOMY.items():
        if any(term.lower() in text for term in terms):
            tags.append(topic)
    return "; ".join(tags[:5])


def _is_system_noise(row: pd.Series) -> bool:
    sender = str(row.get("from_email", ""))
    subject = str(row.get("subject", ""))
    folder = str(row.get("folder_path", ""))
    if SYSTEM_SENDER_RE.search(sender):
        return True
    if SYSTEM_SUBJECT_RE.search(subject):
        return True
    return "sync issues" in folder.lower() or "יומן סינכרון" in folder


def _date_range(group: pd.DataFrame) -> tuple[str, str]:
    dates = pd.to_datetime(group.get("date_sort", ""), errors="coerce", utc=True).dropna()
    if dates.empty:
        return "", ""
    return _format_date(dates.min()), _format_date(dates.max())


def _format_date(value: object) -> str:
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%d/%m/%Y")


def _as_float(value: object) -> float | None:
    try:
        if value is None or str(value).strip() == "":
            return None
        return float(value)
    except (TypeError, ValueError):
        return None
