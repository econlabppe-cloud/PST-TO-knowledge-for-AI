from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path

from pst_kb.models import ExtractedMessageRef


class ExtractorError(RuntimeError):
    """Raised when extraction fails for a specific PST file."""


class ExtractorUnavailable(ExtractorError):
    """Raised when the configured extraction backend is not installed."""


@dataclass(frozen=True)
class ExtractionOptions:
    folder_path: str | None = None
    limit: int | None = None


class Extractor(ABC):
    @abstractmethod
    def extract(
        self,
        pst_path: Path,
        staging_dir: Path,
        options: ExtractionOptions,
    ) -> list[ExtractedMessageRef]:
        """Extract a PST into message references."""
