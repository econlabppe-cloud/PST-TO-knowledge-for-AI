from __future__ import annotations

import re
from pathlib import Path

import pandas as pd


SEARCH_COLUMNS = [
    "message_id",
    "date",
    "knowledge_topic",
    "knowledge_subtopic",
    "cluster_name",
    "llm_topic",
    "llm_subtopic",
    "from_email",
    "to_emails",
    "subject",
    "folder_path",
    "clean_body",
]


def load_clustered_csv(path: Path, include_filtered: bool = False) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    if not include_filtered and "is_filtered" in df.columns:
        df = df[df["is_filtered"].astype(str).str.lower() != "true"].copy()
    for column in SEARCH_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    df["date_sort"] = pd.to_datetime(df.get("date", ""), errors="coerce", utc=True)
    return df


def search_dataframe(
    df: pd.DataFrame,
    topic: str | None = None,
    person: str | None = None,
    query: str | None = None,
    limit: int = 50,
) -> pd.DataFrame:
    results = df.copy()
    topic_column = _preferred_topic_column(results)
    if topic:
        topic_pattern = _contains_pattern(topic)
        results = results[
            results[topic_column].str.contains(topic_pattern, case=False, na=False, regex=True)
            | results.get("llm_subtopic", "").astype(str).str.contains(topic_pattern, case=False, na=False, regex=True)
            | results.get("knowledge_subtopic", "").astype(str).str.contains(topic_pattern, case=False, na=False, regex=True)
            | results["subject"].str.contains(topic_pattern, case=False, na=False, regex=True)
        ]
    if person:
        person_pattern = _contains_pattern(person)
        people_blob = (
            results["from_email"].astype(str)
            + " "
            + results.get("to_emails", "").astype(str)
            + " "
            + results.get("cc_emails", "").astype(str)
            + " "
            + results.get("bcc_emails", "").astype(str)
        )
        results = results[people_blob.str.contains(person_pattern, case=False, na=False, regex=True)]
    if query:
        query_pattern = _contains_pattern(query)
        text_blob = (
            results["subject"].astype(str)
            + " "
            + results["clean_body"].astype(str)
            + " "
            + results.get("body", "").astype(str)
        )
        results = results[text_blob.str.contains(query_pattern, case=False, na=False, regex=True)]

    results = results.assign(search_score=_score_results(results, topic=topic, person=person, query=query))
    results = results.sort_values(["search_score", "date_sort"], ascending=[False, False], na_position="last")
    return results.head(limit).reset_index(drop=True)


def summarize_topics(df: pd.DataFrame) -> pd.DataFrame:
    topic_column = _preferred_topic_column(df)
    grouped = df.groupby(topic_column, dropna=False)
    rows = []
    for topic, group in grouped:
        if not str(topic).strip():
            continue
        rows.append(
            {
                "topic": str(topic),
                "email_count": len(group),
                "first_date": _format_date(group["date_sort"].min()),
                "last_date": _format_date(group["date_sort"].max()),
                "top_people": "; ".join(group["from_email"].value_counts().head(5).index.tolist()),
            }
        )
    return pd.DataFrame(rows).sort_values("email_count", ascending=False)


def summarize_people(df: pd.DataFrame) -> pd.DataFrame:
    topic_column = _preferred_topic_column(df)
    rows = []
    for person, group in df.groupby("from_email", dropna=False):
        if not str(person).strip():
            continue
        rows.append(
            {
                "person": person,
                "email_count": len(group),
                "first_date": _format_date(group["date_sort"].min()),
                "last_date": _format_date(group["date_sort"].max()),
                "top_topics": "; ".join(group[topic_column].value_counts().head(5).index.tolist()),
            }
        )
    return pd.DataFrame(rows).sort_values("email_count", ascending=False)


def render_search_results(results: pd.DataFrame, max_body_chars: int = 700) -> str:
    parts = [f"Search results: {len(results)}", ""]
    for index, row in results.iterrows():
        body = str(row.get("clean_body") or row.get("body") or "")
        if len(body) > max_body_chars:
            body = body[:max_body_chars].rstrip() + "\n[truncated]"
        topic_value = _row_topic_value(row)
        parts.extend(
            [
                f"=== Result {index + 1} ===",
                f"Date: {_format_date(row.get('date_sort'))}",
                f"Topic: {topic_value}",
                f"Subtopic: {row.get('knowledge_subtopic') or row.get('llm_subtopic') or ''}",
                f"From: {row.get('from_email', '')}",
                f"To: {row.get('to_emails', '')}",
                f"Subject: {row.get('subject', '')}",
                f"Message ID: {row.get('message_id', '')}",
                "",
                body,
                "",
            ]
        )
    return "\n".join(parts).strip() + "\n"


def _score_results(
    df: pd.DataFrame,
    topic: str | None,
    person: str | None,
    query: str | None,
) -> list[int]:
    scores: list[int] = []
    topic_column = _preferred_topic_column(df)
    for _, row in df.iterrows():
        score = 0
        subject = str(row.get("subject", "")).lower()
        body = str(row.get("clean_body", "")).lower()
        cluster = str(row.get(topic_column, "")).lower()
        subtopic = str(row.get("knowledge_subtopic", "") or row.get("llm_subtopic", "")).lower()
        from_email = str(row.get("from_email", "")).lower()
        recipients = " ".join(str(row.get(column, "")).lower() for column in ("to_emails", "cc_emails", "bcc_emails"))
        if topic and topic.lower() in cluster:
            score += 50
        if topic and topic.lower() in subject:
            score += 20
        if topic and topic.lower() in subtopic:
            score += 25
        if person and person.lower() in from_email:
            score += 40
        if person and person.lower() in recipients:
            score += 15
        if query and query.lower() in subject:
            score += 25
        if query and query.lower() in body:
            score += 10
        scores.append(score)
    return scores


def _contains_pattern(value: str) -> str:
    terms = [term for term in re.split(r"\s+", value.strip()) if term]
    if not terms:
        return r"$^"
    return ".*".join(re.escape(term) for term in terms)


def _format_date(value: object) -> str:
    parsed = pd.to_datetime(value, errors="coerce", utc=True)
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%d/%m/%Y")


def _preferred_topic_column(df: pd.DataFrame) -> str:
    for column in ("knowledge_topic", "llm_topic", "cluster_name"):
        if column in df.columns and df[column].astype(str).str.strip().any():
            return column
    return "cluster_name"


def _row_topic_value(row: pd.Series) -> str:
    for column in ("knowledge_topic", "llm_topic", "cluster_name"):
        value = str(row.get(column, "") or "").strip()
        if value:
            return value
    return ""
