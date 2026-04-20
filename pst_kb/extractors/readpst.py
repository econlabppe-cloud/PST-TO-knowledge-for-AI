from __future__ import annotations

import logging
import shutil
import subprocess
from pathlib import Path

from pst_kb.extractors.base import ExtractionOptions, Extractor, ExtractorError, ExtractorUnavailable
from pst_kb.models import ExtractedMessageRef
from pst_kb.utils.files import normalize_folder_path, sanitize_filename
from pst_kb.utils.hashing import stable_hash

logger = logging.getLogger(__name__)


class ReadpstExtractor(Extractor):
    """Extract PST files with the external `readpst` utility."""

    def __init__(self, command: str = "readpst", extra_args: list[str] | None = None) -> None:
        self.command = command
        self.extra_args = extra_args or []

    def extract(
        self,
        pst_path: Path,
        staging_dir: Path,
        options: ExtractionOptions,
    ) -> list[ExtractedMessageRef]:
        executable = shutil.which(self.command)
        if executable is None:
            raise ExtractorUnavailable(
                "readpst was not found. Install pst-utils/libpst and ensure readpst is available in PATH, "
                "or pass --readpst-command with the full executable path."
            )

        safe_stem = sanitize_filename(pst_path.stem, fallback="pst")
        pst_stage = staging_dir / f"{safe_stem}_{stable_hash([str(pst_path.resolve())])[:10]}"
        pst_stage.mkdir(parents=True, exist_ok=True)

        cmd = [executable, "-e", "-8", *self.extra_args, "-o", str(pst_stage), str(pst_path)]
        logger.info("Extracting %s with readpst", pst_path)
        completed = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if completed.returncode != 0:
            raise ExtractorError(
                f"readpst failed for {pst_path} with exit code {completed.returncode}: "
                f"{completed.stderr.strip() or completed.stdout.strip()}"
            )

        refs: list[ExtractedMessageRef] = []
        folder_filter = _normalize_filter(options.folder_path)
        for eml_path in sorted(pst_stage.rglob("*.eml")):
            relative_parent = eml_path.parent.relative_to(pst_stage)
            source_folder = normalize_folder_path(relative_parent)
            if folder_filter and folder_filter not in source_folder.lower():
                continue
            refs.append(
                ExtractedMessageRef(
                    pst_path=pst_path,
                    eml_path=eml_path,
                    source_folder=source_folder or "root",
                )
            )
            if options.limit is not None and len(refs) >= options.limit:
                break

        logger.info("readpst produced %s EML messages for %s", len(refs), pst_path.name)
        return refs


def _normalize_filter(folder_path: str | None) -> str | None:
    if not folder_path:
        return None
    return folder_path.replace("\\", "/").strip("/").lower()
