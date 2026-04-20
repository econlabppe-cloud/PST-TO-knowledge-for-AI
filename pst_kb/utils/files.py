from __future__ import annotations

import re
from pathlib import Path

WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def sanitize_filename(value: str | None, fallback: str = "file", max_length: int = 160) -> str:
    name = (value or fallback).strip()
    name = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", name)
    name = re.sub(r"_+", "_", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    if not name:
        name = fallback
    stem = Path(name).stem
    suffix = Path(name).suffix
    if stem.upper() in WINDOWS_RESERVED_NAMES:
        stem = f"{stem}_"
    trimmed_stem = stem[: max(1, max_length - len(suffix))]
    return f"{trimmed_stem}{suffix}"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 2
    while True:
        candidate = parent / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def normalize_folder_path(path: Path) -> str:
    return "/".join(part for part in path.parts if part not in ("", "."))
