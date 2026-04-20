from __future__ import annotations

import argparse
from pathlib import Path

from pst_kb.notebooklm.llm_tagging import tag_with_llm


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Apply LLM topic tagging to clustered PST emails.")
    parser.add_argument("--input-csv", type=Path, default=Path("data/intermediate/emails_clustered.csv"))
    parser.add_argument("--output-csv", type=Path, default=Path("data/intermediate/emails_llm_tagged.csv"))
    parser.add_argument("--provider", choices=["openai", "mock"], default="openai")
    parser.add_argument("--model", default="gpt-4o-mini")
    parser.add_argument("--base-url", default="https://api.openai.com/v1")
    parser.add_argument("--api-key", default=None)
    parser.add_argument("--temperature", type=float, default=0.1)
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--include-filtered", action="store_true")
    parser.add_argument("--batch-size", type=int, default=1)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    tag_with_llm(
        input_csv=args.input_csv,
        output_csv=args.output_csv,
        provider=args.provider,
        model=args.model,
        base_url=args.base_url,
        api_key=args.api_key,
        temperature=args.temperature,
        limit=args.limit,
        include_filtered=args.include_filtered,
        batch_size=args.batch_size,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
