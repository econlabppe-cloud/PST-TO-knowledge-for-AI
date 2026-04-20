from __future__ import annotations

import argparse
import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol
from urllib import request

import pandas as pd
from tqdm import tqdm

from pst_kb.notebooklm.search_core import load_clustered_csv


@dataclass(slots=True)
class LLMTagResult:
    topic: str
    subtopic: str | None = None
    tags: list[str] | None = None
    confidence: float | None = None
    summary: str | None = None


class Tagger(Protocol):
    def tag(self, *, subject: str, body: str, sender: str, recipients: str, cluster: str) -> LLMTagResult: ...


class OpenAICompatibleTagger:
    def __init__(self, api_key: str, model: str, base_url: str = "https://api.openai.com/v1", temperature: float = 0.1):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url.rstrip("/")
        self.temperature = temperature

    def tag(self, *, subject: str, body: str, sender: str, recipients: str, cluster: str) -> LLMTagResult:
        prompt = _build_prompt(subject=subject, body=body, sender=sender, recipients=recipients, cluster=cluster)
        payload = {
            "model": self.model,
            "temperature": self.temperature,
            "messages": [
                {
                    "role": "system",
                    "content": "You are a precise classifier for Hebrew and English email archives. Return only valid JSON.",
                },
                {"role": "user", "content": prompt},
            ],
            "response_format": {"type": "json_object"},
        }
        req = request.Request(
            f"{self.base_url}/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        with request.urlopen(req, timeout=120) as resp:
            raw = json.loads(resp.read().decode("utf-8"))
        content = raw["choices"][0]["message"]["content"]
        data = json.loads(content)
        return _parse_llm_result(data)


class MockTagger:
    def __init__(self, mapping: dict[str, LLMTagResult] | None = None):
        self.mapping = mapping or {}

    def tag(self, *, subject: str, body: str, sender: str, recipients: str, cluster: str) -> LLMTagResult:
        key = _topic_key(subject=subject, body=body, sender=sender, cluster=cluster)
        return self.mapping.get(key) or LLMTagResult(topic=cluster or "unknown", tags=[cluster or "unknown"])


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="LLM-based topic tagging for clustered PST emails.")
    parser.add_argument("--input-csv", type=Path, default=Path("data/intermediate/emails_clustered.csv"))
    parser.add_argument("--output-csv", type=Path, default=Path("data/intermediate/emails_llm_tagged.csv"))
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--include-filtered", action="store_true")
    parser.add_argument("--provider", choices=["openai", "mock"], default="openai")
    parser.add_argument("--model", default=os.environ.get("OPENAI_MODEL", "gpt-4o-mini"))
    parser.add_argument("--base-url", default=os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1"))
    parser.add_argument("--api-key", default=os.environ.get("OPENAI_API_KEY"))
    parser.add_argument("--temperature", type=float, default=0.1)
    parser.add_argument("--batch-size", type=int, default=1)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.provider == "openai" and not args.api_key:
        raise SystemExit("OPENAI_API_KEY is required for --provider openai")
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


def tag_with_llm(
    *,
    input_csv: Path,
    output_csv: Path,
    provider: str = "openai",
    model: str = "gpt-4o-mini",
    base_url: str = "https://api.openai.com/v1",
    api_key: str | None = None,
    temperature: float = 0.1,
    limit: int | None = None,
    include_filtered: bool = False,
    batch_size: int = 1,
) -> Path:
    df = load_clustered_csv(input_csv, include_filtered=include_filtered)
    if limit is not None:
        df = df.head(limit).copy()

    tagger = _make_tagger(provider=provider, api_key=api_key, model=model, base_url=base_url, temperature=temperature)
    rows: list[dict[str, object]] = []
    iterator = df.iterrows()
    for _, row in tqdm(iterator, total=len(df), desc="LLM tagging", unit="email"):
        result = tagger.tag(
            subject=str(row.get("subject", "")),
            body=str(row.get("clean_body") or row.get("body") or ""),
            sender=str(row.get("from_email", "")),
            recipients=" ".join(str(row.get(col, "")) for col in ("to_emails", "cc_emails", "bcc_emails")),
            cluster=str(row.get("cluster_name", "")),
        )
        updated = dict(row)
        updated["llm_topic"] = result.topic
        updated["llm_subtopic"] = result.subtopic
        updated["llm_tags"] = "; ".join(result.tags or [])
        updated["llm_confidence"] = result.confidence
        updated["llm_summary"] = result.summary
        rows.append(updated)

    out_df = pd.DataFrame(rows)
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_csv(output_csv, index=False, encoding="utf-8-sig")
    return output_csv


def _make_tagger(*, provider: str, api_key: str | None, model: str, base_url: str, temperature: float) -> Tagger:
    if provider == "mock":
        return MockTagger()
    if not api_key:
        raise ValueError("api_key is required for openai provider")
    return OpenAICompatibleTagger(api_key=api_key, model=model, base_url=base_url, temperature=temperature)


def _build_prompt(*, subject: str, body: str, sender: str, recipients: str, cluster: str) -> str:
    body = (body or "").strip()
    if len(body) > 6000:
        body = body[:6000]
    return f"""
Classify this email archive item into a single business topic and optional subtopic.
Return JSON with keys: topic, subtopic, tags, confidence, summary.

Context:
- cluster_hint: {cluster}
- sender: {sender}
- recipients: {recipients}
- subject: {subject}
- body:
{body}

Rules:
- Prefer concise Hebrew topic names.
- Preserve domain-specific terms like pensions, claims, contracts, benefits, survivors, calculations.
- If the item is a system message, set topic to "system".
- confidence must be a number from 0 to 1.
""".strip()


def _parse_llm_result(data: dict[str, object]) -> LLMTagResult:
    return LLMTagResult(
        topic=str(data.get("topic") or "unknown"),
        subtopic=(str(data["subtopic"]) if data.get("subtopic") else None),
        tags=[str(item) for item in data.get("tags", []) if str(item).strip()] if isinstance(data.get("tags"), list) else None,
        confidence=_maybe_float(data.get("confidence")),
        summary=(str(data["summary"]) if data.get("summary") else None),
    )


def _maybe_float(value: object) -> float | None:
    try:
        return float(value) if value is not None else None
    except (TypeError, ValueError):
        return None


def _topic_key(*, subject: str, body: str, sender: str, cluster: str) -> str:
    return "|".join(part.strip().lower() for part in (subject[:80], body[:80], sender, cluster))
