from __future__ import annotations

import logging
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

from pst_kb.config import AppConfig
from pst_kb.deduplication import Deduplicator
from pst_kb.exporters import DatasetExporter
from pst_kb.extractors import ExtractionOptions, ExtractorError, ExtractorUnavailable, ReadpstExtractor
from pst_kb.models import AttachmentRecord, MessageRecord, ProcessingError, ProcessingReport, ThreadRecord
from pst_kb.parsers import EmailParser
from pst_kb.processor import MessageProcessor
from pst_kb.threading import ThreadBuilder

logger = logging.getLogger(__name__)


class PstKbPipeline:
    """
    Main orchestration class for the PST processing pipeline.
    """

    def __init__(self, config: AppConfig) -> None:
        self.config = config
        if config.output_dir is None:
            raise ValueError("output_dir is required")
        self.output_dir = config.output_dir
        self.exporter = DatasetExporter(self.output_dir, include_sqlite=config.export_sqlite)

    def run(self) -> ProcessingReport:
        """
        Executes the pipeline: find PST files, extract, parse, process, deduplicate, thread, and export.
        """
        report = ProcessingReport()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        staging_dir = self.output_dir / "_staging"
        staging_dir.mkdir(parents=True, exist_ok=True)

        messages: list[MessageRecord] = []
        attachments: list[AttachmentRecord] = []
        threads = []

        try:
            pst_files = self._find_pst_files()
            report.input_files_found = len(pst_files)
            if not pst_files:
                report.fatal_error = "No PST files found."
                return self._finish(report, messages, attachments, threads)

            extractor = ReadpstExtractor(self.config.readpst_command)
            parser = EmailParser()
            processor = MessageProcessor(self.config, self.output_dir)
            extraction_limit_remaining = self.config.limit

            for pst_path in pst_files:
                if extraction_limit_remaining is not None and extraction_limit_remaining <= 0:
                    break

                try:
                    refs = extractor.extract(
                        pst_path,
                        staging_dir,
                        ExtractionOptions(
                            folder_path=self.config.folder_path,
                            limit=extraction_limit_remaining,
                        ),
                    )
                    report.pst_files_processed += 1
                except ExtractorUnavailable as exc:
                    report.fatal_error = str(exc)
                    report.errors.append(ProcessingError(source=str(pst_path), message=str(exc)))
                    return self._finish(report, messages, attachments, threads)
                except ExtractorError as exc:
                    logger.exception("Extraction failed for %s", pst_path)
                    report.errors.append(ProcessingError(source=str(pst_path), message="extraction_failed", detail=str(exc)))
                    report.skipped_files.append(str(pst_path))
                    continue

                report.messages_extracted += len(refs)
                for ref in refs:
                    if extraction_limit_remaining is not None and extraction_limit_remaining <= 0:
                        break
                    try:
                        for raw in parser.parse_eml_with_nested(ref):
                            if extraction_limit_remaining is not None and extraction_limit_remaining <= 0:
                                break
                            message, message_attachments = processor.process(raw)
                            messages.append(message)
                            attachments.extend(message_attachments)
                            report.messages_processed += 1
                            if self.config.skip_attachments:
                                report.attachments_skipped += len(message_attachments)
                            else:
                                report.attachments_exported += sum(
                                    1 for item in message_attachments if item.saved_path and not item.export_error
                                )
                            if extraction_limit_remaining is not None:
                                extraction_limit_remaining -= 1
                    except Exception as exc:
                        logger.exception("Failed to process %s", ref.eml_path)
                        report.errors.append(
                            ProcessingError(source=str(ref.eml_path), message="message_processing_failed", detail=str(exc))
                        )

            messages = Deduplicator().mark_duplicates(messages)
            messages, threads = ThreadBuilder().assign_threads(messages)

            report.duplicates_found = sum(1 for message in messages if message.is_duplicate)
            report.threads_created = len(threads)
            report.stats = {
                "messages_with_attachments": sum(1 for message in messages if message.has_attachments),
                "messages_with_clean_body": sum(1 for message in messages if bool(message.body_text_clean)),
                "languages": _count_values(message.language for message in messages),
                "likely_sender_types": _count_values(message.likely_sender_type for message in messages),
                "possible_intents": _count_values(message.possible_intent for message in messages),
            }
            return self._finish(report, messages, attachments, threads)
        except Exception as exc:
            logger.exception("Fatal pipeline error")
            report.fatal_error = str(exc)
            report.errors.append(ProcessingError(source="pipeline", message="fatal_error", detail=str(exc)))
            report.finished_at = datetime.now(timezone.utc)
            self.exporter.write_report_only(report)
            return report
        finally:
            if not self.config.keep_staging:
                shutil.rmtree(staging_dir, ignore_errors=True)

    def _finish(
        self,
        report: ProcessingReport,
        messages: list[MessageRecord],
        attachments: list[AttachmentRecord],
        threads: Iterable[ThreadRecord],
    ) -> ProcessingReport:
        """
        Finalizes the report and exports the results.
        """
        report.finished_at = datetime.now(timezone.utc)
        if messages or attachments or threads:
            self.exporter.export(messages, attachments, list(threads), report)
        else:
            self.exporter.write_report_only(report)
        return report

    def _find_pst_files(self) -> list[Path]:
        """
        Scans the input directory (or single file) for PST files.
        """
        if self.config.single_file:
            single = self.config.single_file
            if not single.exists():
                raise FileNotFoundError(f"Single PST file not found: {single}")
            if single.suffix.lower() != ".pst":
                raise ValueError(f"--single-file must point to a .pst file: {single}")
            return [single]

        if self.config.input_dir is None:
            return []
        if not self.config.input_dir.exists():
            raise FileNotFoundError(f"Input directory not found: {self.config.input_dir}")

        pattern = "**/*.pst" if self.config.recursive else "*.pst"
        return sorted(path for path in self.config.input_dir.glob(pattern) if path.is_file())


def _count_values(values: Iterable[object]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for value in values:
        key = str(value or "unknown")
        counts[key] = counts.get(key, 0) + 1
    return counts
