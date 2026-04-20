from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pst_kb.notebooklm.search_core import (
    load_clustered_csv,
    render_search_results,
    search_dataframe,
    summarize_people,
    summarize_topics,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Search clustered PST emails by topic, person, and text.")
    parser.add_argument("--input-csv", type=Path, default=Path("data/intermediate/emails_clustered.csv"))
    parser.add_argument("--topic", help="Topic or cluster name to search.")
    parser.add_argument("--person", help="Sender or recipient email/name fragment to search.")
    parser.add_argument("--query", help="Free-text query matched against subject and clean body.")
    parser.add_argument("--limit", type=int, default=25)
    parser.add_argument("--include-filtered", action="store_true")
    parser.add_argument("--list-topics", action="store_true")
    parser.add_argument("--list-people", action="store_true")
    parser.add_argument("--output-txt", type=Path)
    parser.add_argument("--output-csv", type=Path)
    return parser


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    args = build_parser().parse_args(argv)
    df = load_clustered_csv(args.input_csv, include_filtered=args.include_filtered)

    if args.list_topics:
        topics = summarize_topics(df).head(args.limit)
        print(topics.to_string(index=False))
        return 0

    if args.list_people:
        people = summarize_people(df).head(args.limit)
        print(people.to_string(index=False))
        return 0

    results = search_dataframe(
        df,
        topic=args.topic,
        person=args.person,
        query=args.query,
        limit=args.limit,
    )
    text = render_search_results(results)
    print(_console_summary(results))

    if args.output_txt:
        args.output_txt.parent.mkdir(parents=True, exist_ok=True)
        args.output_txt.write_text(text, encoding="utf-8-sig")

    if args.output_csv:
        args.output_csv.parent.mkdir(parents=True, exist_ok=True)
        results.drop(columns=["date_sort"], errors="ignore").to_csv(args.output_csv, index=False, encoding="utf-8-sig")

    if not args.output_txt:
        print()
        print(text)
    return 0


def _console_summary(results) -> str:
    lines = [f"Found {len(results)} results"]
    for index, row in results.head(10).iterrows():
        date = row.get("date", "")
        topic = row.get("knowledge_topic") or row.get("llm_topic") or row.get("cluster_name", "")
        lines.append(
            f"{index + 1}. {date} | {topic} | {row.get('from_email', '')} | {row.get('subject', '')}"
        )
    return "\n".join(lines)
