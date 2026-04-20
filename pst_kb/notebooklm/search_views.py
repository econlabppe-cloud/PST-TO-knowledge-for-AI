from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
from tqdm import tqdm

from pst_kb.notebooklm.search_core import load_clustered_csv, render_search_results, search_dataframe, summarize_people, summarize_topics
from pst_kb.utils.files import sanitize_filename


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build browsing/search views by topic and by person.")
    parser.add_argument("--input-csv", type=Path, default=Path("data/intermediate/emails_clustered.csv"))
    parser.add_argument("--output-dir", type=Path, default=Path("data/output/search_views_txt"))
    parser.add_argument("--top-people", type=int, default=80)
    parser.add_argument("--max-results-per-file", type=int, default=300)
    parser.add_argument("--include-filtered", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    build_search_views(
        input_csv=args.input_csv,
        output_dir=args.output_dir,
        top_people=args.top_people,
        max_results_per_file=args.max_results_per_file,
        include_filtered=args.include_filtered,
    )
    return 0


def build_search_views(
    input_csv: Path,
    output_dir: Path,
    top_people: int = 80,
    max_results_per_file: int = 300,
    include_filtered: bool = False,
) -> dict[str, Path]:
    df = load_clustered_csv(input_csv, include_filtered=include_filtered)
    output_dir.mkdir(parents=True, exist_ok=True)
    by_topic_dir = output_dir / "by_topic"
    by_person_dir = output_dir / "by_person"
    by_topic_dir.mkdir(parents=True, exist_ok=True)
    by_person_dir.mkdir(parents=True, exist_ok=True)

    topic_summary = summarize_topics(df)
    people_summary = summarize_people(df)
    topic_summary.to_csv(output_dir / "topics.csv", index=False, encoding="utf-8-sig")
    people_summary.to_csv(output_dir / "people.csv", index=False, encoding="utf-8-sig")

    topic_index_lines = ["TOPIC SEARCH INDEX", ""]
    for _, row in tqdm(topic_summary.iterrows(), total=len(topic_summary), desc="Writing topic views", unit="topic"):
        topic = str(row["topic"])
        results = search_dataframe(df, topic=topic, limit=max_results_per_file)
        filename = _safe_txt_name(topic)
        path = by_topic_dir / filename
        path.write_text(render_search_results(results, max_body_chars=900), encoding="utf-8-sig")
        topic_index_lines.append(f"- {topic} | {row['email_count']} emails | {filename}")

    person_index_lines = ["PERSON SEARCH INDEX", ""]
    for _, row in tqdm(
        people_summary.head(top_people).iterrows(),
        total=min(top_people, len(people_summary)),
        desc="Writing person views",
        unit="person",
    ):
        person = str(row["person"])
        results = search_dataframe(df, person=person, limit=max_results_per_file)
        filename = _safe_txt_name(person)
        path = by_person_dir / filename
        path.write_text(render_search_results(results, max_body_chars=900), encoding="utf-8-sig")
        person_index_lines.append(f"- {person} | {row['email_count']} emails | {filename}")

    (output_dir / "index_by_topic.txt").write_text("\n".join(topic_index_lines) + "\n", encoding="utf-8-sig")
    (output_dir / "index_by_person.txt").write_text("\n".join(person_index_lines) + "\n", encoding="utf-8-sig")
    (output_dir / "README.txt").write_text(_readme_text(), encoding="utf-8-sig")

    return {
        "topics_csv": output_dir / "topics.csv",
        "people_csv": output_dir / "people.csv",
        "index_by_topic": output_dir / "index_by_topic.txt",
        "index_by_person": output_dir / "index_by_person.txt",
    }


def _safe_txt_name(value: str) -> str:
    name = sanitize_filename(value.replace("@", "_at_").replace(" ", "_"), fallback="unknown", max_length=110)
    return f"{Path(name).stem}.txt"


def _readme_text() -> str:
    return """SEARCH VIEWS

This folder provides lightweight browsing views for the archive.

Files:
- index_by_topic.txt: topic browsing index
- index_by_person.txt: person browsing index
- topics.csv: topic summary table
- people.csv: people summary table

Example:
python search_emails.py --topic "קרנות פנסיה" --person "yanivr@mof.gov.il" --limit 20
"""
