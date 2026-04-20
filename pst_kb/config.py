from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml
from pydantic import BaseModel, ConfigDict, Field


class AppConfig(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)

    input_dir: Path | None = None
    output_dir: Path | None = None
    recursive: bool = False
    skip_attachments: bool = False
    export_sqlite: bool = False
    limit: int | None = None
    single_file: Path | None = None
    folder_path: str | None = None
    readpst_command: str = "readpst"
    keep_staging: bool = False
    log_level: str = "INFO"
    internal_domains: list[str] = Field(default_factory=list)
    topic_keywords: dict[str, list[str]] = Field(default_factory=dict)
    sender_type_keywords: dict[str, list[str]] = Field(default_factory=dict)
    llm_provider: str | None = None
    llm_model: str | None = None
    llm_base_url: str | None = None
    llm_api_key: str | None = None
    llm_temperature: float = 0.1
    llm_batch_size: int = 8


def load_config(path: Path | None) -> AppConfig:
    if path is None:
        return AppConfig()
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")

    with path.open("r", encoding="utf-8") as handle:
        data: dict[str, Any] = yaml.safe_load(handle) or {}
    return AppConfig.model_validate(data)
