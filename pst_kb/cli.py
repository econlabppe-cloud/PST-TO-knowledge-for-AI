from __future__ import annotations

import argparse
import logging
from pathlib import Path

from pst_kb.config import AppConfig, load_config
from pst_kb.pipeline import PstKbPipeline
from pst_kb.utils.logging import configure_logging


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Process Microsoft Outlook PST files into a structured dataset."
    )
    parser.add_argument("--input", dest="input_dir", type=Path, help="Input folder containing PST files.")
    parser.add_argument("--output", dest="output_dir", type=Path, help="Output folder for exported dataset files.")
    parser.add_argument("--recursive", action="store_true", help="Recursively scan the input folder.")
    parser.add_argument("--skip-attachments", action="store_true", help="Do not export attachment binaries.")
    parser.add_argument("--export-sqlite", action="store_true", help="Also export a SQLite database.")
    parser.add_argument("--log-level", default=None, help="Logging level, e.g. INFO or DEBUG.")
    parser.add_argument("--limit", type=int, default=None, help="Maximum number of messages to process.")
    parser.add_argument("--single-file", type=Path, default=None, help="Process one PST file instead of scanning.")
    parser.add_argument("--folder-path", default=None, help="Best-effort Outlook folder path filter.")
    parser.add_argument("--config", type=Path, default=None, help="YAML config path.")
    parser.add_argument("--readpst-command", default=None, help="Path or command name for readpst.")
    parser.add_argument("--keep-staging", action="store_true", help="Keep intermediate EML files.")
    return parser


def merge_cli_config(args: argparse.Namespace, config: AppConfig) -> AppConfig:
    updates: dict[str, object] = {}
    for name in (
        "input_dir",
        "output_dir",
        "limit",
        "single_file",
        "folder_path",
        "readpst_command",
        "log_level",
    ):
        value = getattr(args, name, None)
        if value is not None:
            updates[name] = value

    if args.recursive:
        updates["recursive"] = True
    if args.skip_attachments:
        updates["skip_attachments"] = True
    if args.export_sqlite:
        updates["export_sqlite"] = True
    if args.keep_staging:
        updates["keep_staging"] = True

    return config.model_copy(update=updates)


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    config = merge_cli_config(args, load_config(args.config))

    configure_logging(config.log_level)
    logger = logging.getLogger(__name__)

    if not config.input_dir and not config.single_file:
        parser.error("Provide --input or --single-file, or set input_dir in the config.")
    if not config.output_dir:
        parser.error("Provide --output or set output_dir in the config.")

    logger.info("Starting PST dataset processing")
    pipeline = PstKbPipeline(config)
    report = pipeline.run()

    if report.fatal_error:
        logger.error("Processing finished with fatal error: %s", report.fatal_error)
        return 2

    logger.info(
        "Processing complete: %s messages, %s attachments, %s threads",
        report.messages_processed,
        report.attachments_exported,
        report.threads_created,
    )
    return 0
