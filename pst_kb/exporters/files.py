from __future__ import annotations

import csv
import json
import logging
from pathlib import Path
from typing import Any

from pydantic import BaseModel

from pst_kb.exporters.sqlite import export_sqlite
from pst_kb.models import AttachmentRecord, MessageRecord, ProcessingReport, ThreadRecord

logger = logging.getLogger(__name__)


class DatasetExporter:
    def __init__(self, output_dir: Path, include_sqlite: bool = False) -> None:
        self.output_dir = output_dir
        self.include_sqlite = include_sqlite

    def export(
        self,
        messages: list[MessageRecord],
        attachments: list[AttachmentRecord],
        threads: list[ThreadRecord],
        report: ProcessingReport,
    ) -> None:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._write_jsonl(self.output_dir / "messages.jsonl", messages)
        self._write_csv(self.output_dir / "messages.csv", messages, list(MessageRecord.model_fields))
        self._write_csv(self.output_dir / "attachments.csv", attachments, list(AttachmentRecord.model_fields))
        self._write_csv(self.output_dir / "threads.csv", threads, list(ThreadRecord.model_fields))
        (self.output_dir / "processing_report.json").write_text(
            report.model_dump_json(indent=2),
            encoding="utf-8",
        )
        if self.include_sqlite:
            export_sqlite(self.output_dir / "dataset.sqlite", messages, attachments, threads)

    def write_report_only(self, report: ProcessingReport) -> None:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        (self.output_dir / "processing_report.json").write_text(
            report.model_dump_json(indent=2),
            encoding="utf-8",
        )

    def _write_jsonl(self, path: Path, rows: list[BaseModel]) -> None:
        with path.open("w", encoding="utf-8", newline="\n") as handle:
            for row in rows:
                handle.write(row.model_dump_json())
                handle.write("\n")

    def _write_csv(self, path: Path, rows: list[BaseModel], fieldnames: list[str]) -> None:
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(
                handle,
                fieldnames=fieldnames,
                extrasaction="ignore",
                quoting=csv.QUOTE_ALL,
                lineterminator="\n",
            )
            writer.writeheader()
            for row in rows:
                writer.writerow(_flatten(row.model_dump(mode="json")))


def _flatten(row: dict[str, Any]) -> dict[str, Any]:
    flattened: dict[str, Any] = {}
    for key, value in row.items():
        if isinstance(value, (list, dict)):
            flattened[key] = json.dumps(value, ensure_ascii=False, sort_keys=True)
        elif value is None:
            flattened[key] = ""
        else:
            flattened[key] = value
    return flattened
