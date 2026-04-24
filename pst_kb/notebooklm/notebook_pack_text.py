from __future__ import annotations

import argparse
import logging
from collections import Counter
from pathlib import Path

import pandas as pd
from tqdm import tqdm

from pst_kb.notebooklm.docx_writer import write_text_docx
from pst_kb.notebooklm.build_notebooks import select_representatives
from pst_kb.notebooklm.common import configure_script_logging
from pst_kb.notebooklm.topic_classifier import TopicMatch
from pst_kb.notebooklm.topic_taxonomy import classify_email_record_corpus, render_curated_rules_text
from pst_kb.utils.files import sanitize_filename

logger = logging.getLogger(__name__)

REQUIRED_COLUMNS = [
    "message_id",
    "date",
    "subject",
    "body",
    "clean_body",
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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build a Word-based NotebookLM knowledge pack from clustered/tagged emails.")
    parser.add_argument(
        "--input-csv",
        "--input",
        dest="input_csv",
        type=Path,
        default=Path("data/intermediate/emails_clustered.csv"),
    )
    parser.add_argument(
        "--output-dir",
        "--output",
        dest="output_dir",
        type=Path,
        default=Path("data/output/notebooklm_pack_word"),
    )
    parser.add_argument("--topic-source", choices=["rules", "hybrid", "llm", "cluster"], default="rules")
    parser.add_argument("--max-source-chars", type=int, default=160_000)
    parser.add_argument("--max-emails-per-file", type=int, default=80)
    parser.add_argument("--max-body-chars", type=int, default=8_500)
    parser.add_argument("--top-people", type=int, default=60)
    parser.add_argument("--min-score", type=float, default=4.0)
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
        min_score=args.min_score,
        include_filtered=args.include_filtered,
        include_system=args.include_system,
    )
    return 0


def build_notebooklm_pack(
    *,
    input_csv: Path,
    output_dir: Path,
    topic_source: str = "rules",
    max_source_chars: int = 160_000,
    max_emails_per_file: int = 80,
    max_body_chars: int = 8_500,
    top_people: int = 60,
    min_score: float = 4.0,
    include_filtered: bool = False,
    include_system: bool = False,
) -> dict[str, Path]:
    df = _load_input(input_csv)
    df = _classify(df, topic_source=topic_source, min_score=min_score)

    excluded_mask = pd.Series(False, index=df.index)
    if not include_filtered and "is_filtered" in df.columns:
        excluded_mask |= df["is_filtered"].astype(str).str.lower() == "true"
    if not include_system:
        excluded_mask |= df["is_system_noise"]

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

    grouped_topics = list(included.groupby("knowledge_topic", dropna=False))
    for topic, group in tqdm(grouped_topics, total=len(grouped_topics), desc="Writing NotebookLM pack", unit="topic"):
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

    guide_path = output_dir / "00_GUIDE.docx"
    index_path = output_dir / "00_INDEX.docx"
    topic_map_path = output_dir / "01_TOPIC_MAP.docx"
    people_map_path = output_dir / "02_PEOPLE_MAP.docx"
    rules_path = output_dir / "03_CLASSIFICATION_RULES.docx"
    review_txt_path = output_dir / "04_REVIEW_QUEUE.docx"

    write_text_docx(guide_path, _render_guide(input_csv, included, excluded, source_paths, topic_source), title="NotebookLM Guide")
    write_text_docx(index_path, _render_index(topics, manifest), title="NotebookLM Knowledge Pack Index")
    write_text_docx(topic_map_path, _render_topic_map(topics), title="Topic Map")
    write_text_docx(people_map_path, _render_people_map(people), title="People Map")
    write_text_docx(rules_path, render_curated_rules_text(), title="Classification Rules")
    write_text_docx(review_txt_path, _render_review_text(review), title="Review Queue")

    logger.info("NotebookLM pack written to %s", output_dir)
    return {
        "guide": guide_path,
        "index": index_path,
        "topic_map": topic_map_path,
        "people_map": people_map_path,
        "rules": rules_path,
        "review": review_txt_path,
        "manifest": manifest_path,
        "topics": topics_path,
        "people": people_path,
        "review_csv": review_path,
        "excluded": excluded_path,
    }


def _load_input(input_csv: Path) -> pd.DataFrame:
    df = pd.read_csv(input_csv, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    for column in REQUIRED_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    df["date_sort"] = pd.to_datetime(df.get("date", ""), errors="coerce", utc=True)
    return df


def _classify(df: pd.DataFrame, *, topic_source: str, min_score: float) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for _, row in tqdm(df.iterrows(), total=len(df), desc="Classifying emails", unit="email"):
        match = classify_email_record_corpus(row, topic_source=topic_source, min_score=min_score)
        out = row.to_dict()
        out.update(
            {
                "knowledge_topic": match.topic,
                "knowledge_subtopic": match.subtopic,
                "knowledge_score": match.score,
                "knowledge_matched_terms": "; ".join(match.matched_terms),
                "knowledge_mode": match.source_mode,
                "knowledge_review_required": match.review_required,
                "is_system_noise": match.is_system_noise,
                "body_for_pack": str(row.get("clean_body") or row.get("body") or ""),
            }
        )
        rows.append(out)
    return pd.DataFrame(rows)


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
    ordered = group.sort_values(["date_sort", "message_id"], na_position="last")
    chunks: list[pd.DataFrame] = []
    current_rows: list[pd.Series] = []
    current_chars = 0

    for _, row in ordered.iterrows():
        rendered = _render_email(row, max_body_chars=max_body_chars)
        if current_rows and (len(current_rows) >= max_emails_per_file or current_chars + len(rendered) > max_source_chars):
            chunks.append(pd.DataFrame([item.to_dict() for item in current_rows]))
            current_rows = []
            current_chars = 0
        current_rows.append(row)
        current_chars += len(rendered)

    if current_rows:
        chunks.append(pd.DataFrame([item.to_dict() for item in current_rows]))

    paths: list[Path] = []
    total_parts = len(chunks)
    for index, chunk in enumerate(chunks, start=1):
        filename = f"{topic_slug}__part_{index:03d}.docx"
        path = sources_dir / filename
        write_text_docx(
            path,
            _render_source_file(topic=topic, part_number=index, total_parts=total_parts, group=chunk, max_body_chars=max_body_chars),
            title=topic,
        )
        rel_path = f"sources/{filename}"
        for _, row in chunk.iterrows():
            manifest_rows.append(
                {
                    "source_file": rel_path,
                    "knowledge_topic": row.get("knowledge_topic", ""),
                    "knowledge_subtopic": row.get("knowledge_subtopic", ""),
                    "knowledge_score": row.get("knowledge_score", ""),
                    "knowledge_mode": row.get("knowledge_mode", ""),
                    "knowledge_review_required": row.get("knowledge_review_required", ""),
                    "message_id": row.get("message_id", ""),
                    "date": row.get("date", ""),
                    "from_email": row.get("from_email", ""),
                    "subject": row.get("subject", ""),
                    "folder_path": row.get("folder_path", ""),
                    "cluster_name": row.get("cluster_name", ""),
                    "llm_topic": row.get("llm_topic", ""),
                    "llm_confidence": row.get("llm_confidence", ""),
                    "knowledge_matched_terms": row.get("knowledge_matched_terms", ""),
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
    max_body_chars: int,
) -> str:
    start, end = _date_range(group)
    top_senders = "; ".join(group["from_email"].astype(str).value_counts().head(8).index.tolist())
    top_subtopics = "; ".join(group["knowledge_subtopic"].astype(str).value_counts().head(6).index.tolist())
    lines = [
        f"TOPIC: {topic}",
        f"PART: {part_number} of {total_parts}",
        f"EMAIL COUNT: {len(group)}",
        f"DATE RANGE: {start} to {end}",
        f"TOP SENDERS: {top_senders}",
        f"TOP SUBTOPICS: {top_subtopics}",
        "",
        "Use this file as evidence for questions about this topic.",
        "Each email block keeps metadata plus the cleaned body.",
        "Prefer date, sender, subject, and message id when citing answers.",
        "",
        "------------------------------------------------------------",
        "",
    ]
    for _, row in group.iterrows():
        lines.append(_render_email(row, max_body_chars=max_body_chars))
    return "\n".join(lines).strip() + "\n"


def _render_email(row: pd.Series, max_body_chars: int) -> str:
    body = str(row.get("body_for_pack") or "")
    if len(body) > max_body_chars:
        body = body[:max_body_chars].rstrip() + "\n[body truncated for NotebookLM sizing]"

    lines = [
        "EMAIL",
        f"Message ID: {row.get('message_id', '')}",
        f"Date: {_format_date(row.get('date_sort'))}",
        f"From: {row.get('from_email', '')}",
        f"To: {row.get('to_emails', '')}",
        f"CC: {row.get('cc_emails', '')}",
        f"BCC: {row.get('bcc_emails', '')}",
        f"Subject: {row.get('subject', '')}",
        f"Subject Normalized: {row.get('subject_normalized', '')}",
        f"Folder: {row.get('folder_path', '')}",
        f"Knowledge Topic: {row.get('knowledge_topic', '')}",
        f"Knowledge Subtopic: {row.get('knowledge_subtopic', '')}",
        f"Knowledge Score: {row.get('knowledge_score', '')}",
        f"Matched Terms: {row.get('knowledge_matched_terms', '')}",
        f"Classification Mode: {row.get('knowledge_mode', '')}",
        f"Review Required: {row.get('knowledge_review_required', '')}",
        "",
        "Body:",
        body or "[empty body]",
        "",
        "------------------------------------------------------------",
        "",
    ]
    return "\n".join(lines)


def _topic_summary_row(topic: str, group: pd.DataFrame, paths: list[Path], output_dir: Path) -> dict[str, object]:
    representatives = select_representatives(group, limit=3)
    top_people = "; ".join(group["from_email"].astype(str).value_counts().head(8).index.tolist())
    start, end = _date_range(group)
    subtopic_counts = group["knowledge_subtopic"].astype(str).value_counts().head(5)
    representative_subjects = " | ".join(str(row.get("subject", "")) for row in representatives)
    return {
        "knowledge_topic": topic,
        "email_count": len(group),
        "source_file_count": len(paths),
        "first_date": start,
        "last_date": end,
        "top_people": top_people,
        "top_subtopics": "; ".join(f"{name} ({count})" for name, count in subtopic_counts.items()),
        "representative_subjects": representative_subjects,
        "source_files": "; ".join(path.relative_to(output_dir).as_posix() for path in paths),
    }


def _people_summary(df: pd.DataFrame, top_people: int) -> pd.DataFrame:
    rows = []
    topic_column = "knowledge_topic" if "knowledge_topic" in df.columns else "cluster_name"
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
                "top_topics": "; ".join(group[topic_column].astype(str).value_counts().head(6).index.tolist()),
            }
        )
    return pd.DataFrame(rows).sort_values("email_count", ascending=False).head(top_people)


def _review_queue(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df.iterrows():
        reasons = []
        score = _as_float(row.get("knowledge_score"))
        if row.get("knowledge_review_required"):
            reasons.append("manual_review")
        if score is not None and score < 4.0:
            reasons.append("low_score")
        if not str(row.get("subject", "")).strip():
            reasons.append("missing_subject")
        if len(str(row.get("body_for_pack", ""))) < 180:
            reasons.append("short_body")
        if reasons:
            rows.append(
                {
                    "review_reason": "; ".join(reasons),
                    "knowledge_topic": row.get("knowledge_topic", ""),
                    "knowledge_subtopic": row.get("knowledge_subtopic", ""),
                    "knowledge_score": row.get("knowledge_score", ""),
                    "message_id": row.get("message_id", ""),
                    "date": row.get("date", ""),
                    "from_email": row.get("from_email", ""),
                    "subject": row.get("subject", ""),
                }
            )
    return pd.DataFrame(rows)


def _render_index(topics: pd.DataFrame, manifest: pd.DataFrame) -> str:
    lines = [
        "NOTEBOOKLM KNOWLEDGE PACK INDEX",
        "",
        f"Total source files: {manifest['source_file'].nunique() if not manifest.empty else 0}",
        f"Total included emails: {len(manifest)}",
        "",
        "TOPICS",
        "",
    ]
    for _, row in topics.iterrows():
        lines.extend(
            [
                f"- {row['knowledge_topic']}",
                f"  emails: {row['email_count']}",
                f"  source files: {row['source_file_count']}",
                f"  date range: {row['first_date']} to {row['last_date']}",
                f"  top people: {row['top_people']}",
                f"  top subtopics: {row['top_subtopics']}",
                f"  representative subjects: {row['representative_subjects']}",
                f"  files: {row['source_files']}",
                "",
            ]
        )
    return "\n".join(lines).strip() + "\n"


def _render_topic_map(topics: pd.DataFrame) -> str:
    lines = [
        "TOPIC MAP",
        "",
        "This map helps NotebookLM and humans load the right files for a question.",
        "",
    ]
    for _, row in topics.iterrows():
        lines.extend(
            [
                f"TOPIC: {row['knowledge_topic']}",
                f"Emails: {row['email_count']}",
                f"Date range: {row['first_date']} to {row['last_date']}",
                f"Top people: {row['top_people']}",
                f"Top subtopics: {row['top_subtopics']}",
                f"Representative subjects: {row['representative_subjects']}",
                f"Source files: {row['source_files']}",
                "",
            ]
        )
    return "\n".join(lines).strip() + "\n"


def _render_people_map(people: pd.DataFrame) -> str:
    lines = [
        "PEOPLE MAP",
        "",
        "This map highlights who writes most often and which topics they touch.",
        "",
    ]
    for _, row in people.iterrows():
        lines.extend(
            [
                f"PERSON: {row['person']}",
                f"Emails: {row['email_count']}",
                f"Date range: {row['first_date']} to {row['last_date']}",
                f"Top topics: {row['top_topics']}",
                "",
            ]
        )
    return "\n".join(lines).strip() + "\n"


def _render_guide(
    input_csv: Path,
    included: pd.DataFrame,
    excluded: pd.DataFrame,
    source_paths: list[Path],
    topic_source: str,
) -> str:
    return "\n".join(
        [
            "NOTEBOOKLM GUIDE",
            "",
            "Recommended upload order:",
            "1. 00_GUIDE.docx",
            "2. 00_INDEX.docx",
            "3. 01_TOPIC_MAP.docx",
            "4. 02_PEOPLE_MAP.docx",
            "5. 03_CLASSIFICATION_RULES.docx",
            "6. Relevant files from sources/",
            "",
            "How to use it:",
            "- Ask by topic, person, date range, or subject.",
            "- Ask NotebookLM to cite date, sender, subject, and message id.",
            "- For broad questions, load only the relevant topic files.",
            "",
            f"Input CSV: {input_csv}",
            f"Included emails: {len(included)}",
            f"Excluded or noisy records: {len(excluded)}",
            f"Source files generated: {len(source_paths)}",
            f"Topic source mode: {topic_source}",
            "",
            "Classification approach:",
            "- The default mode is transparent keyword rules.",
            "- High-confidence LLM tags can be used only when explicitly selected.",
            "- System mail and delivery failures are excluded from NotebookLM sources by default.",
            "",
            "The manifest CSV preserves per-message provenance and classification evidence.",
            "",
        ]
    ).strip() + "\n"


def _render_review_text(review: pd.DataFrame) -> str:
    lines = [
        "REVIEW QUEUE",
        "",
        "These items are worth a human glance before using them for downstream decisions.",
        "",
    ]
    if review.empty:
        lines.append("No review items were detected.")
    else:
        for _, row in review.head(500).iterrows():
            lines.append(
                f"- {row['review_reason']} | {row['date']} | {row['from_email']} | {row['subject']} | {row['knowledge_topic']}"
            )
    return "\n".join(lines).strip() + "\n"


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
