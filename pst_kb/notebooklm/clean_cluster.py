from __future__ import annotations

import argparse
import logging
import math
import re
from collections import Counter
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics import silhouette_score
from tqdm import tqdm

from pst_kb.cleaners import EmailCleaner
from pst_kb.notebooklm.common import configure_script_logging, word_count
from pst_kb.normalizers import normalize_subject, normalize_whitespace

logger = logging.getLogger(__name__)

DEFAULT_MODEL_NAME = "paraphrase-multilingual-MiniLM-L12-v2"

AUTO_REPLY_PATTERNS = [
    r"automatic reply",
    r"out of office",
    r"auto reply",
    r"תשובה אוטומטית",
    r"מחוץ למשרד",
    r"היעדרות",
    r"אישור מסירה",
    r"delivery status notification",
]
NOISE_SENDER_PATTERNS = [
    r"no-?reply",
    r"do-?not-?reply",
    r"newsletter",
    r"digest",
    r"mailing-list",
    r"roundrobin",
    r"system",
    r"bot",
    r"notifications?",
]
DOMAIN_TERMS = [
    "גמלה",
    "גמלאות",
    "פנסיה",
    "שאירים",
    "קצבה",
    "קרן",
    "זכאות",
    "תביעה",
    "ערעור",
    "חוזה",
    "רציפות",
    "מבוטח",
    "תגמולים",
    "שיפוי",
]
HEBREW_STOPWORDS = {
    "של",
    "על",
    "עם",
    "אל",
    "או",
    "אם",
    "לא",
    "כן",
    "זה",
    "זו",
    "הוא",
    "היא",
    "את",
    "אני",
    "אנחנו",
    "שלום",
    "תודה",
    "בברכה",
    "לגבי",
    "בנושא",
    "הנדון",
    "מצורף",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Clean and cluster emails_raw.csv for NotebookLM.")
    parser.add_argument("--input-csv", "--input", dest="input_csv", type=Path, required=True)
    parser.add_argument("--output-csv", "--output", dest="output_csv", type=Path, required=True)
    parser.add_argument("--model-name", default=DEFAULT_MODEL_NAME)
    parser.add_argument("--embedding-backend", choices=["auto", "sentence-transformers", "tfidf"], default="auto")
    parser.add_argument("--min-words", type=int, default=50)
    parser.add_argument("--k-min", type=int, default=5)
    parser.add_argument("--k-max", type=int)
    parser.add_argument("--batch-size", type=int, default=64)
    parser.add_argument("--body-chars", type=int, default=3500)
    parser.add_argument("--log-file", type=Path)
    parser.add_argument("--log-level", default="INFO")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    configure_script_logging(args.log_level, args.log_file)
    run_clean_and_cluster(
        input_csv=args.input_csv,
        output_csv=args.output_csv,
        model_name=args.model_name,
        embedding_backend=args.embedding_backend,
        min_words=args.min_words,
        k_min=args.k_min,
        k_max=args.k_max,
        batch_size=args.batch_size,
        body_chars=args.body_chars,
    )
    return 0


def run_clean_and_cluster(
    input_csv: Path,
    output_csv: Path,
    model_name: str = DEFAULT_MODEL_NAME,
    embedding_backend: str = "auto",
    min_words: int = 50,
    k_min: int = 5,
    k_max: int | None = None,
    batch_size: int = 64,
    body_chars: int = 3500,
) -> pd.DataFrame:
    logger.info("Loading raw CSV: %s", input_csv)
    df = pd.read_csv(input_csv, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    logger.info("Loaded %s rows", len(df))

    processed = prepare_clean_dataframe(df, min_words=min_words, body_chars=body_chars)
    candidates = processed[~processed["is_filtered"]].copy()
    logger.info("Rows after filtering: %s", len(candidates))

    if candidates.empty:
        processed["cluster_id"] = ""
        processed["cluster_name"] = ""
        processed.to_csv(output_csv, index=False, encoding="utf-8-sig")
        return processed

    vectors, vectorizer = build_embeddings(
        candidates["embedding_text"].tolist(),
        model_name=model_name,
        backend=embedding_backend,
        batch_size=batch_size,
    )
    selected_k = choose_k(vectors, len(candidates), k_min=k_min, k_max=k_max)
    logger.info("Selected cluster count: %s", selected_k)

    labels = KMeans(n_clusters=selected_k, random_state=42, n_init=10).fit_predict(vectors)
    candidates["cluster_id"] = labels.astype(str)
    cluster_names = name_clusters(candidates, vectorizer)
    candidates["cluster_name"] = candidates["cluster_id"].map(cluster_names)

    processed["cluster_id"] = ""
    processed["cluster_name"] = ""
    processed.loc[candidates.index, "cluster_id"] = candidates["cluster_id"]
    processed.loc[candidates.index, "cluster_name"] = candidates["cluster_name"]

    output_csv.parent.mkdir(parents=True, exist_ok=True)
    processed.to_csv(output_csv, index=False, encoding="utf-8-sig")
    log_cluster_stats(processed)
    return processed


def prepare_clean_dataframe(df: pd.DataFrame, min_words: int, body_chars: int) -> pd.DataFrame:
    cleaner = EmailCleaner()
    rows = []
    for _, row in tqdm(df.iterrows(), total=len(df), desc="Cleaning emails", unit="email"):
        subject = str(row.get("subject", ""))
        body = normalize_whitespace(str(row.get("body", "")))
        cleaned = cleaner.clean(body).text
        clean_words = word_count(cleaned)
        is_filtered, reason = filter_reason(row, subject, cleaned, clean_words, min_words)
        normalized_subject = normalize_subject(subject)
        embedding_text = build_embedding_text(normalized_subject, cleaned, body_chars=body_chars)

        out = row.to_dict()
        out.update(
            {
                "subject_normalized": normalized_subject,
                "clean_body": cleaned,
                "word_count": clean_words,
                "embedding_text": embedding_text,
                "is_filtered": is_filtered,
                "filter_reason": reason,
            }
        )
        rows.append(out)
    return pd.DataFrame(rows)


def filter_reason(row: pd.Series, subject: str, clean_body: str, words: int, min_words: int) -> tuple[bool, str]:
    if not clean_body.strip():
        return True, "empty_body"

    subject_lower = subject.lower()
    sender_lower = str(row.get("from_email", "")).lower()
    if any(re.search(pattern, subject_lower, re.IGNORECASE) for pattern in AUTO_REPLY_PATTERNS):
        return True, "auto_reply"
    if any(re.search(pattern, sender_lower, re.IGNORECASE) for pattern in NOISE_SENDER_PATTERNS):
        return True, "noise_sender"
    if words < min_words and not contains_domain_term(clean_body + " " + subject):
        return True, "too_short"
    return False, ""


def contains_domain_term(text: str) -> bool:
    return any(term in text for term in DOMAIN_TERMS)


def build_embedding_text(subject: str, body: str, body_chars: int) -> str:
    # נותנים משקל כפול לנושא בלי לפגוע בטקסט המקצועי עצמו.
    return normalize_whitespace(f"{subject}\n{subject}\n{body[:body_chars]}")


def build_embeddings(
    texts: list[str],
    model_name: str,
    backend: str,
    batch_size: int,
) -> tuple[np.ndarray, TfidfVectorizer | None]:
    if backend in ("auto", "sentence-transformers"):
        try:
            from sentence_transformers import SentenceTransformer

            logger.info("Building embeddings with sentence-transformers model %s", model_name)
            model = SentenceTransformer(model_name)
            vectors = model.encode(
                texts,
                batch_size=batch_size,
                show_progress_bar=True,
                normalize_embeddings=True,
            )
            return np.asarray(vectors), None
        except Exception as exc:
            if backend == "sentence-transformers":
                raise
            logger.warning("sentence-transformers unavailable, using local TF-IDF fallback: %s", exc)

    logger.info("Building local TF-IDF vectors")
    vectorizer = TfidfVectorizer(max_features=5000, ngram_range=(1, 2), min_df=1)
    vectors = vectorizer.fit_transform(texts)
    return vectors.toarray(), vectorizer


def choose_k(vectors: np.ndarray, n_rows: int, k_min: int, k_max: int | None) -> int:
    if n_rows <= 2:
        return 1

    distinct_vectors = max(1, np.unique(np.round(vectors, decimals=8), axis=0).shape[0])
    if distinct_vectors <= 2:
        return min(distinct_vectors, n_rows)

    dynamic_max = int(max(2, min(50, max(10, math.sqrt(n_rows)))))
    upper = min(k_max or dynamic_max, n_rows - 1, distinct_vectors)
    lower = min(max(2, k_min), upper)
    if lower >= upper:
        return lower

    candidates = list(range(lower, upper + 1))
    inertias: list[float] = []
    silhouettes: list[float] = []
    for k in candidates:
        model = KMeans(n_clusters=k, random_state=42, n_init=10).fit(vectors)
        labels = model.labels_
        inertias.append(float(model.inertia_))
        if len(set(labels)) > 1:
            try:
                silhouettes.append(float(silhouette_score(vectors, labels)))
            except Exception:
                silhouettes.append(-1.0)
        else:
            silhouettes.append(-1.0)

    elbow_k = _elbow_k(candidates, inertias)
    best_silhouette_k = candidates[int(np.argmax(silhouettes))]
    if max(silhouettes) > 0:
        return best_silhouette_k
    return elbow_k


def name_clusters(df: pd.DataFrame, vectorizer: TfidfVectorizer | None) -> dict[str, str]:
    names: dict[str, str] = {}
    for cluster_id, group in df.groupby("cluster_id"):
        text = " ".join(group["subject_normalized"].tolist() + group["clean_body"].tolist())
        domain_name = domain_name_from_terms(text)
        if domain_name:
            names[str(cluster_id)] = domain_name
            continue
        top_terms = top_terms_for_group(group, vectorizer)
        names[str(cluster_id)] = top_terms[0] if top_terms else f"נושא_{int(cluster_id) + 1}"
    return names


def domain_name_from_terms(text: str) -> str | None:
    mapping = [
        ("קרנות פנסיה", ["קרן", "פנסיה", "קרנות"]),
        ("זכאות לגמלה", ["זכאות", "גמלה", "קצבה"]),
        ("שאירים ויתומים", ["שאירים", "יתומים", "אלמנה"]),
        ("תביעות וערעורים", ["תביעה", "ערעור", "ערעורים"]),
        ("חישובי גמלה", ["חישוב", "משכורת קובעת", "גמלה"]),
        ("חוזים והתקשרויות", ["חוזה", "התקשרות", "ספק"]),
        ("רציפות זכויות", ["רציפות", "זכויות"]),
        ("ביטוח וגבייה", ["ביטוח", "גבייה", "תגמולים"]),
    ]
    lowered = text.lower()
    scores = [(name, sum(1 for term in terms if term in lowered)) for name, terms in mapping]
    best_name, score = max(scores, key=lambda item: item[1])
    return best_name if score else None


def top_terms_for_group(group: pd.DataFrame, vectorizer: TfidfVectorizer | None) -> list[str]:
    tokens = []
    for text in group["embedding_text"].tolist():
        tokens.extend(_tokens(text))
    counts = Counter(token for token in tokens if token not in HEBREW_STOPWORDS and len(token) > 2)
    return [term for term, _ in counts.most_common(5)]


def log_cluster_stats(df: pd.DataFrame) -> None:
    filtered = df[df["is_filtered"]]
    active = df[~df["is_filtered"]]
    logger.info("Input rows: %s", len(df))
    logger.info("Filtered rows by reason: %s", filtered["filter_reason"].value_counts().to_dict())
    logger.info("Rows kept: %s", len(active))
    logger.info("Clusters created: %s", active["cluster_id"].nunique())
    logger.info("Cluster sizes: %s", active["cluster_name"].value_counts().to_dict())


def _tokens(text: str) -> list[str]:
    return re.findall(r"[\u0590-\u05ffA-Za-z0-9][\u0590-\u05ffA-Za-z0-9_-]+", text.lower())


def _elbow_k(candidates: list[int], inertias: list[float]) -> int:
    if len(candidates) <= 2:
        return candidates[0]
    drops = [inertias[index - 1] - inertias[index] for index in range(1, len(inertias))]
    if not drops:
        return candidates[0]
    ratios = [
        drops[index] / drops[index - 1]
        for index in range(1, len(drops))
        if drops[index - 1] > 0
    ]
    if not ratios:
        return candidates[0]
    elbow_index = int(np.argmin(ratios)) + 1
    return candidates[min(elbow_index, len(candidates) - 1)]
