from __future__ import annotations

import argparse
import logging
from pathlib import Path
from time import perf_counter

from pst_kb.extractors import ExtractorError
from pst_kb.notebooklm.clean_cluster import DEFAULT_MODEL_NAME, run_clean_and_cluster
from pst_kb.notebooklm.common import configure_script_logging
from pst_kb.notebooklm.extract_csv import extract_to_rows, log_extract_stats, write_raw_csv
from pst_kb.notebooklm.notebook_pack_text import build_notebooklm_pack

logger = logging.getLogger(__name__)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the complete PST to NotebookLM Word knowledge-pack workflow."
    )
    parser.add_argument("--pst-path", type=Path, help="Path to the Outlook PST file.")
    parser.add_argument("--work-dir", type=Path, default=Path("data"), help="Base folder for intermediate and output files.")
    parser.add_argument("--raw-csv", type=Path, help="Override path for emails_raw.csv.")
    parser.add_argument("--clustered-csv", type=Path, help="Override path for emails_clustered.csv.")
    parser.add_argument("--pack-dir", type=Path, help="Override output folder for the NotebookLM Word pack.")
    parser.add_argument("--extractor", choices=["auto", "pypff", "readpst"], default="auto")
    parser.add_argument("--readpst-command", default="readpst")
    parser.add_argument("--temp-dir", type=Path)
    parser.add_argument("--max-emails", type=int, help="Limit messages for a small test run.")
    parser.add_argument(
        "--keep-attachments-in-eml",
        action="store_true",
        help="Keep MIME attachments in extracted EML files. This can use much more disk space.",
    )
    parser.add_argument("--skip-extract", action="store_true", help="Reuse an existing raw CSV.")
    parser.add_argument("--skip-clean", action="store_true", help="Reuse an existing clustered CSV.")
    parser.add_argument("--model-name", default=DEFAULT_MODEL_NAME)
    parser.add_argument("--embedding-backend", choices=["auto", "sentence-transformers", "tfidf"], default="tfidf")
    parser.add_argument("--min-words", type=int, default=50)
    parser.add_argument("--k-min", type=int, default=5)
    parser.add_argument("--k-max", type=int)
    parser.add_argument("--batch-size", type=int, default=64)
    parser.add_argument("--body-chars", type=int, default=3500)
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

    try:
        outputs = run_notebooklm_workflow(
            pst_path=args.pst_path,
            work_dir=args.work_dir,
            raw_csv=args.raw_csv,
            clustered_csv=args.clustered_csv,
            pack_dir=args.pack_dir,
            extractor=args.extractor,
            readpst_command=args.readpst_command,
            temp_dir=args.temp_dir,
            max_emails=args.max_emails,
            keep_attachments_in_eml=args.keep_attachments_in_eml,
            skip_extract=args.skip_extract,
            skip_clean=args.skip_clean,
            model_name=args.model_name,
            embedding_backend=args.embedding_backend,
            min_words=args.min_words,
            k_min=args.k_min,
            k_max=args.k_max,
            batch_size=args.batch_size,
            body_chars=args.body_chars,
            topic_source=args.topic_source,
            max_source_chars=args.max_source_chars,
            max_emails_per_file=args.max_emails_per_file,
            max_body_chars=args.max_body_chars,
            top_people=args.top_people,
            min_score=args.min_score,
            include_filtered=args.include_filtered,
            include_system=args.include_system,
        )
    except (ExtractorError, FileNotFoundError, RuntimeError) as exc:
        logger.error("%s", exc)
        return 2

    logger.info("Workflow completed.")
    logger.info("Raw CSV: %s", outputs["raw_csv"])
    logger.info("Clustered CSV: %s", outputs["clustered_csv"])
    logger.info("NotebookLM Word pack: %s", outputs["pack_dir"])
    return 0


def run_notebooklm_workflow(
    *,
    pst_path: Path | None,
    work_dir: Path = Path("data"),
    raw_csv: Path | None = None,
    clustered_csv: Path | None = None,
    pack_dir: Path | None = None,
    extractor: str = "auto",
    readpst_command: str = "readpst",
    temp_dir: Path | None = None,
    max_emails: int | None = None,
    keep_attachments_in_eml: bool = False,
    skip_extract: bool = False,
    skip_clean: bool = False,
    model_name: str = DEFAULT_MODEL_NAME,
    embedding_backend: str = "tfidf",
    min_words: int = 50,
    k_min: int = 5,
    k_max: int | None = None,
    batch_size: int = 64,
    body_chars: int = 3500,
    topic_source: str = "rules",
    max_source_chars: int = 160_000,
    max_emails_per_file: int = 80,
    max_body_chars: int = 8_500,
    top_people: int = 60,
    min_score: float = 4.0,
    include_filtered: bool = False,
    include_system: bool = False,
) -> dict[str, Path]:
    intermediate_dir = work_dir / "intermediate"
    output_dir = work_dir / "output"
    raw_csv = raw_csv or intermediate_dir / "emails_raw.csv"
    clustered_csv = clustered_csv or intermediate_dir / "emails_clustered.csv"
    pack_dir = pack_dir or output_dir / "notebooklm_pack_word"

    _run_extract_stage(
        pst_path=pst_path,
        raw_csv=raw_csv,
        extractor=extractor,
        temp_dir=temp_dir,
        max_emails=max_emails,
        readpst_command=readpst_command,
        keep_attachments_in_eml=keep_attachments_in_eml,
        skip_extract=skip_extract,
    )
    _run_clean_stage(
        raw_csv=raw_csv,
        clustered_csv=clustered_csv,
        model_name=model_name,
        embedding_backend=embedding_backend,
        min_words=min_words,
        k_min=k_min,
        k_max=k_max,
        batch_size=batch_size,
        body_chars=body_chars,
        skip_clean=skip_clean,
    )
    _run_pack_stage(
        clustered_csv=clustered_csv,
        pack_dir=pack_dir,
        topic_source=topic_source,
        max_source_chars=max_source_chars,
        max_emails_per_file=max_emails_per_file,
        max_body_chars=max_body_chars,
        top_people=top_people,
        min_score=min_score,
        include_filtered=include_filtered,
        include_system=include_system,
    )

    return {
        "raw_csv": raw_csv,
        "clustered_csv": clustered_csv,
        "pack_dir": pack_dir,
    }


def _run_extract_stage(
    *,
    pst_path: Path | None,
    raw_csv: Path,
    extractor: str,
    temp_dir: Path | None,
    max_emails: int | None,
    readpst_command: str,
    keep_attachments_in_eml: bool,
    skip_extract: bool,
) -> None:
    if skip_extract:
        if not raw_csv.exists():
            raise FileNotFoundError(f"--skip-extract was used but raw CSV does not exist: {raw_csv}")
        logger.info("Skipping extraction and reusing %s", raw_csv)
        return
    if pst_path is None:
        raise FileNotFoundError("--pst-path is required unless --skip-extract is used")

    start = perf_counter()
    logger.info("Stage 1/3: extracting PST to raw CSV")
    rows = extract_to_rows(
        pst_path=pst_path,
        extractor=extractor,
        temp_dir=temp_dir,
        max_emails=max_emails,
        readpst_command=readpst_command,
        discard_attachments=not keep_attachments_in_eml,
    )
    write_raw_csv(rows, raw_csv)
    log_extract_stats(rows)
    logger.info("Extraction stage finished in %.1f seconds", perf_counter() - start)


def _run_clean_stage(
    *,
    raw_csv: Path,
    clustered_csv: Path,
    model_name: str,
    embedding_backend: str,
    min_words: int,
    k_min: int,
    k_max: int | None,
    batch_size: int,
    body_chars: int,
    skip_clean: bool,
) -> None:
    if skip_clean:
        if not clustered_csv.exists():
            raise FileNotFoundError(f"--skip-clean was used but clustered CSV does not exist: {clustered_csv}")
        logger.info("Skipping clean/cluster and reusing %s", clustered_csv)
        return

    start = perf_counter()
    logger.info("Stage 2/3: cleaning and clustering raw email data")
    run_clean_and_cluster(
        input_csv=raw_csv,
        output_csv=clustered_csv,
        model_name=model_name,
        embedding_backend=embedding_backend,
        min_words=min_words,
        k_min=k_min,
        k_max=k_max,
        batch_size=batch_size,
        body_chars=body_chars,
    )
    logger.info("Cleaning/clustering stage finished in %.1f seconds", perf_counter() - start)


def _run_pack_stage(
    *,
    clustered_csv: Path,
    pack_dir: Path,
    topic_source: str,
    max_source_chars: int,
    max_emails_per_file: int,
    max_body_chars: int,
    top_people: int,
    min_score: float,
    include_filtered: bool,
    include_system: bool,
) -> None:
    start = perf_counter()
    logger.info("Stage 3/3: building NotebookLM Word pack")
    build_notebooklm_pack(
        input_csv=clustered_csv,
        output_dir=pack_dir,
        topic_source=topic_source,
        max_source_chars=max_source_chars,
        max_emails_per_file=max_emails_per_file,
        max_body_chars=max_body_chars,
        top_people=top_people,
        min_score=min_score,
        include_filtered=include_filtered,
        include_system=include_system,
    )
    logger.info("NotebookLM pack stage finished in %.1f seconds", perf_counter() - start)
