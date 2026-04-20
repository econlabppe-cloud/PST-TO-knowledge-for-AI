from __future__ import annotations

import argparse
import logging
from pathlib import Path

import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from tqdm import tqdm

from pst_kb.notebooklm.common import configure_script_logging
from pst_kb.utils.files import sanitize_filename

logger = logging.getLogger(__name__)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build NotebookLM TXT files and index.txt from clustered emails.")
    parser.add_argument("--input-csv", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--max-body-chars", type=int, default=12000)
    parser.add_argument("--representatives", type=int, default=3)
    parser.add_argument("--include-filtered", action="store_true")
    parser.add_argument("--log-file", type=Path)
    parser.add_argument("--log-level", default="INFO")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    configure_script_logging(args.log_level, args.log_file)
    build_notebook_export(
        input_csv=args.input_csv,
        output_dir=args.output_dir,
        max_body_chars=args.max_body_chars,
        representatives=args.representatives,
        include_filtered=args.include_filtered,
    )
    return 0


def build_notebook_export(
    input_csv: Path,
    output_dir: Path,
    max_body_chars: int = 12000,
    representatives: int = 3,
    include_filtered: bool = False,
) -> dict[str, Path]:
    df = pd.read_csv(input_csv, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    if not include_filtered and "is_filtered" in df.columns:
        df = df[df["is_filtered"].astype(str).str.lower() != "true"].copy()

    topic_column = _preferred_topic_column(df)
    df = df[df[topic_column].astype(str).str.strip() != ""].copy()
    df["date_sort"] = pd.to_datetime(df.get("date", ""), errors="coerce", utc=True)

    output_dir.mkdir(parents=True, exist_ok=True)
    written: dict[str, Path] = {}
    index_sections: list[str] = ["NOTEBOOKLM TOPIC INDEX", ""]

    grouped = list(df.groupby(topic_column, dropna=False))
    for topic_name, group in tqdm(grouped, total=len(grouped), desc="Writing notebooks", unit="topic"):
        ordered = group.sort_values(["date_sort", "message_id"], na_position="last")
        filename = _cluster_filename(str(topic_name), used=set(path.name for path in written.values()))
        cluster_path = output_dir / filename
        cluster_path.write_text(
            render_cluster_file(str(topic_name), ordered, max_body_chars=max_body_chars),
            encoding="utf-8-sig",
        )
        written[str(topic_name)] = cluster_path
        index_sections.append(render_index_section(str(topic_name), ordered, filename, representatives=representatives))
        index_sections.append("")

    index_path = output_dir / "index.txt"
    index_path.write_text("\n".join(index_sections).strip() + "\n", encoding="utf-8-sig")
    written["index"] = index_path
    logger.info("Notebook export written to %s", output_dir)
    return written


def render_cluster_file(cluster_name: str, group: pd.DataFrame, max_body_chars: int) -> str:
    start, end = _date_range(group)
    parts = [
        f"TOPIC: {cluster_name}",
        f"EMAIL COUNT: {len(group)}",
        f"DATE RANGE: {start} to {end}",
        "====================================================",
        "",
    ]
    for _, row in group.iterrows():
        body = str(row.get("clean_body") or row.get("body") or "")
        if len(body) > max_body_chars:
            body = body[:max_body_chars].rstrip() + "\n[body truncated for NotebookLM source sizing]"
        parts.extend(
            [
                "EMAIL",
                f"Date: {_format_date(row.get('date', ''))}",
                f"From: {row.get('from_email', '')}",
                f"Subject: {row.get('subject', '')}",
                f"To: {row.get('to_emails', '')}",
                "",
                body or "[empty body]",
                "====================================================",
                "",
            ]
        )
    return "\n".join(parts).strip() + "\n"


def render_index_section(cluster_name: str, group: pd.DataFrame, filename: str, representatives: int) -> str:
    start, end = _date_range(group)
    representative_rows = select_representatives(group, limit=representatives)
    lines = [
        f"TOPIC: {cluster_name}",
        f"Email count: {len(group)}",
        f"Date range: {start} to {end}",
        f"File: {filename}",
        "Representative emails:",
    ]
    for index, row in enumerate(representative_rows, start=1):
        lines.append(
            f"{index}. {_format_date(row.get('date', ''))} - {row.get('subject', '')} - {row.get('from_email', '')}"
        )
    return "\n".join(lines)


def select_representatives(group: pd.DataFrame, limit: int = 3) -> list[pd.Series]:
    texts = (group.get("embedding_text", group.get("clean_body", ""))).astype(str).tolist()
    if len(group) <= limit:
        return [row for _, row in group.iterrows()]
    try:
        matrix = TfidfVectorizer(max_features=3000, ngram_range=(1, 2)).fit_transform(texts)
        scores = matrix.mean(axis=0)
        row_scores = matrix @ scores.T
        ranked_indices = list(reversed(row_scores.A.ravel().argsort().tolist()))
    except Exception:
        ranked_indices = list(range(len(group)))

    selected: list[pd.Series] = []
    seen_subjects: set[str] = set()
    rows = list(group.reset_index(drop=True).iterrows())
    for idx in ranked_indices:
        row = rows[idx][1]
        subject = str(row.get("subject", "")).strip().lower()
        if subject in seen_subjects and len(selected) < limit:
            continue
        selected.append(row)
        seen_subjects.add(subject)
        if len(selected) >= limit:
            break
    return selected


def _cluster_filename(cluster_name: str, used: set[str]) -> str:
    base = sanitize_filename(cluster_name.replace(" ", "_"), fallback="cluster", max_length=80)
    if not base.lower().endswith(".txt"):
        base = f"{base}.txt"
    candidate = base
    counter = 2
    while candidate in used:
        candidate = f"{Path(base).stem}_{counter}.txt"
        counter += 1
    return candidate


def _preferred_topic_column(df: pd.DataFrame) -> str:
    for column in ("knowledge_topic", "llm_topic", "cluster_name"):
        if column in df.columns and df[column].astype(str).str.strip().any():
            return column
    return "cluster_name"


def _date_range(group: pd.DataFrame) -> tuple[str, str]:
    dates = pd.to_datetime(group.get("date", ""), errors="coerce", utc=True).dropna()
    if dates.empty:
        return "", ""
    return _format_date(dates.min()), _format_date(dates.max())


def _format_date(value: object) -> str:
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%d/%m/%Y")
