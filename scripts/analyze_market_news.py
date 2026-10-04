#!/usr/bin/env python3

"""
VGrat FMS - Market News AI Analyzer
===================================

Sequential one-article-per-request Gemini analyzer.

PROCESSING RULES
================

1. Read data/market_news/current.json as READ-ONLY input.
2. Process newest articles first.
3. Exactly ONE article per Gemini request.
4. Save every successful analysis immediately.
5. Only start the next Gemini request after the previous result
   has been successfully validated and saved.
6. Stop immediately on:
      - HTTP 429
      - HTTP 503
      - any other HTTP/API error
      - invalid Gemini response
      - invalid JSON
      - failed validation
      - unexpected Python/code error
7. NO automatic retry.
8. Never continue after a failure.
9. Never re-analyze an article whose articleId already exists in
   current analysis or archived history.
10. Keep only the latest 14 publication-date days in analysis/current.json.
11. Older analyses are moved to:
      data/market_news/analysis/history/YYYY/MM/YYYY-MM-DD.json
12. Archive files are append-only.
13. data/market_news/current.json is never modified.

Gemini
======

Model:
    gemini-3.8-flash

Environment:
    GEMINI_API_KEY
    MARKET_NEWS_MODEL (optional; defaults to gemini-3.8-flash)

Output
======

Each Gemini response must contain exactly one analysis object:

{
    "articleId": "...",
    "relevant": true,
    "category": "MARKET",
    "sentiment": "positive",
    "importance": "high",
    "summary": "...",
    "assetClasses": [],
    "geographies": [],
    "sectors": [],
    "investorImpact": "...",
    "reasoning": "...",
    "fundMonitoringRelevant": true
}

The response is validated before it is written.

IMPORTANT DATA-PRESERVATION DESIGN
==================================

The successful record is written to disk immediately after validation.

Therefore:

    Gemini request
        ↓
    validate response
        ↓
    save successful article
        ↓
    next Gemini request

NOT:

    Gemini requests
        ↓
    hold everything in memory
        ↓
    save at the end

This ensures a later API failure cannot destroy earlier successful work.
"""

from __future__ import annotations

import json
import os
import sys
import time
from datetime import datetime, date, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


# ==============================================================
# PATHS
# ==============================================================

ROOT = Path(__file__).resolve().parent.parent

NEWS_FILE = ROOT / "data" / "market_news" / "current.json"

ANALYSIS_DIR = ROOT / "data" / "market_news" / "analysis"
ANALYSIS_CURRENT_FILE = ANALYSIS_DIR / "current.json"
HISTORY_DIR = ANALYSIS_DIR / "history"


# ==============================================================
# CONFIGURATION
# ==============================================================

DEFAULT_MODEL = "gemini-3.8-flash"

MODEL = os.environ.get(
    "MARKET_NEWS_MODEL",
    DEFAULT_MODEL,
).strip()

API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()

GEMINI_API_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    f"{MODEL}:generateContent"
)

REQUEST_TIMEOUT_SECONDS = 120

ROLLING_DAYS = 14


# ==============================================================
# ALLOWED VALUES
# ==============================================================

ALLOWED_CATEGORIES = {
    "MARKET",
    "ECONOMIC",
    "TECHNOLOGY",
    "GEOPOLITICAL",
    "REJECT",
}

ALLOWED_SENTIMENTS = {
    "positive",
    "negative",
    "mixed",
    "neutral",
}

ALLOWED_IMPORTANCE = {
    "high",
    "medium",
    "low",
}


# ==============================================================
# LOGGING
# ==============================================================

def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


# ==============================================================
# FAIL FAST
# ==============================================================

class AnalysisError(Exception):
    """Expected analyzer failure that must stop processing immediately."""


def fail(message: str) -> None:
    raise AnalysisError(message)


# ==============================================================
# JSON HELPERS
# ==============================================================

def read_json(path: Path) -> Any:
    if not path.exists():
        fail(f"Required file does not exist: {path}")

    try:
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as exc:
        fail(f"Invalid JSON in {path}: {exc}")
    except OSError as exc:
        fail(f"Unable to read {path}: {exc}")

    return None


def write_json_atomic(path: Path, data: Any) -> None:
    """
    Write JSON atomically.

    The temporary file is placed beside the destination and then
    replaced into position. This prevents a partial JSON file if
    the process is interrupted during a write.
    """

    path.parent.mkdir(parents=True, exist_ok=True)

    temporary = path.with_name(
        f".{path.name}.tmp"
    )

    try:
        with temporary.open("w", encoding="utf-8") as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=2,
            )
            f.write("\n")

        temporary.replace(path)

    except OSError as exc:
        try:
            if temporary.exists():
                temporary.unlink()
        except OSError:
            pass

        fail(f"Unable to write {path}: {exc}")


# ==============================================================
# SINGAPORE TIME
# ==============================================================

def singapore_now() -> datetime:
    """
    Singapore is UTC+8 year-round.
    """

    return datetime.now(
        timezone.utc
    ).astimezone(
        timezone(timedelta(hours=8))
    )


# ==============================================================
# PUBLICATION DATE PARSING
# ==============================================================

def parse_publication_datetime(
    value: Any,
) -> datetime | None:
    """
    Parse common ISO/RFC publication timestamps.

    Returns an aware datetime.

    If no valid publication timestamp exists, returns None.
    """

    if value is None:
        return None

    if isinstance(value, datetime):
        dt = value

    elif isinstance(value, (int, float)):
        try:
            dt = datetime.fromtimestamp(
                value,
                tz=timezone.utc,
            )
        except (OverflowError, OSError, ValueError):
            return None

    elif isinstance(value, str):
        text_value = value.strip()

        if not text_value:
            return None

        # ISO 8601 "Z"
        if text_value.endswith("Z"):
            text_value = text_value[:-1] + "+00:00"

        try:
            dt = datetime.fromisoformat(text_value)
        except ValueError:
            # Common RFC 2822 fallback
            try:
                from email.utils import parsedate_to_datetime

                dt = parsedate_to_datetime(text_value)

            except (TypeError, ValueError, OverflowError):
                return None

    else:
        return None

    if dt.tzinfo is None:
        dt = dt.replace(
            tzinfo=timezone.utc
        )

    return dt


def publication_date(
    article: dict[str, Any],
) -> date | None:
    """
    Return publication date in Singapore calendar time.

    The collector's `published` field is the primary source.
    """

    dt = parse_publication_datetime(
        article.get("published")
    )

    if dt is None:
        return None

    singapore_tz = timezone(timedelta(hours=8))

    return dt.astimezone(
        singapore_tz
    ).date()


def publication_sort_key(
    article: dict[str, Any],
) -> tuple[int, str]:
    """
    Newest publication first.

    Articles without a publication date are sorted after dated
    articles, while retaining deterministic ordering by articleId.
    """

    dt = parse_publication_datetime(
        article.get("published")
    )

    if dt is None:
        return (
            0,
            str(article.get("articleId", "")),
        )

    return (
        1,
        dt.astimezone(timezone.utc).isoformat(),
    )


# ==============================================================
# ARTICLE EXTRACTION
# ==============================================================

def extract_articles(
    raw: Any,
) -> list[dict[str, Any]]:
    """
    Accept the collector's common envelope forms.

    Supported:

        {
            "articles": [...]
        }

    or:

        [...]
    """

    if isinstance(raw, list):
        articles = raw

    elif isinstance(raw, dict):
        articles = raw.get("articles")

        if articles is None:
            fail(
                "data/market_news/current.json does not contain "
                "an 'articles' array."
            )

    else:
        fail(
            "data/market_news/current.json must contain "
            "a JSON object or array."
        )

    if not isinstance(articles, list):
        fail("'articles' must be an array.")

    result: list[dict[str, Any]] = []

    for index, article in enumerate(articles):
        if not isinstance(article, dict):
            fail(
                f"Article at index {index} is not an object."
            )

        article_id = article.get("articleId")

        if not isinstance(article_id, str) or not article_id.strip():
            fail(
                f"Article at index {index} has no valid articleId."
            )

        result.append(article)

    return result


# ==============================================================
# ANALYSIS STORAGE
# ==============================================================

def empty_current_analysis() -> dict[str, Any]:
    return {
        "generatedAt": None,
        "source": "Gemini",
        "configuration": {
            "model": MODEL,
            "articlesPerRequest": 1,
            "processingOrder": "newest_first",
            "retry": False,
            "rollingDays": ROLLING_DAYS,
        },
        "analyses": [],
    }


def load_current_analysis() -> dict[str, Any]:
    if not ANALYSIS_CURRENT_FILE.exists():
        return empty_current_analysis()

    data = read_json(
        ANALYSIS_CURRENT_FILE
    )

    if not isinstance(data, dict):
        fail(
            f"{ANALYSIS_CURRENT_FILE} must contain an object."
        )

    analyses = data.get("analyses")

    if analyses is None:
        data["analyses"] = []

    elif not isinstance(analyses, list):
        fail(
            f"'analyses' in {ANALYSIS_CURRENT_FILE} "
            "must be an array."
        )

    return data


def collect_archived_article_ids() -> set[str]:
    """
    Read every history JSON file and collect successfully archived
    article IDs.

    History is append-only, so these IDs represent permanent
    successful analyses.
    """

    ids: set[str] = set()

    if not HISTORY_DIR.exists():
        return ids

    history_files = sorted(
        HISTORY_DIR.rglob("*.json")
    )

    for path in history_files:
        data = read_json(path)

        if isinstance(data, dict):
            analyses = data.get("analyses", [])

        elif isinstance(data, list):
            analyses = data

        else:
            fail(
                f"History file must contain an object or array: {path}"
            )

        if not isinstance(analyses, list):
            fail(
                f"'analyses' must be an array in history file: {path}"
            )

        for record in analyses:
            if not isinstance(record, dict):
                fail(
                    f"Invalid analysis record in history file: {path}"
                )

            article_id = record.get("articleId")

            if isinstance(article_id, str) and article_id:
                ids.add(article_id)

    return ids


def current_article_ids(
    current_analysis: dict[str, Any],
) -> set[str]:
    ids: set[str] = set()

    for record in current_analysis.get(
        "analyses",
        [],
    ):
        if not isinstance(record, dict):
            fail(
                "Current analysis contains a non-object record."
            )

        article_id = record.get("articleId")

        if isinstance(article_id, str) and article_id:
            ids.add(article_id)

    return ids


# ==============================================================
# GEMINI PROMPT
# ==============================================================

SYSTEM_PROMPT = """
You are the market-news analysis engine for VGrat FMS.

Analyze exactly ONE financial/news article.

Your task is to determine whether the article is relevant to
financial-market and investment monitoring.

Use ONLY information supported by the supplied article.

Do not invent facts, companies, sectors, countries, asset classes,
events, impacts, or investment implications.

Return ONLY valid JSON.

The JSON must contain exactly these fields:

{
  "articleId": "string",
  "relevant": true,
  "category": "MARKET",
  "sentiment": "positive",
  "importance": "high",
  "summary": "string",
  "assetClasses": [],
  "geographies": [],
  "sectors": [],
  "investorImpact": "string",
  "reasoning": "string",
  "fundMonitoringRelevant": true
}

CATEGORY
========

Must be exactly one of:

MARKET
ECONOMIC
TECHNOLOGY
GEOPOLITICAL
REJECT

Use REJECT when the article is not meaningfully relevant to
investment, financial markets, economics, technology affecting
markets, or geopolitical developments affecting markets.

SENTIMENT
=========

Must be exactly one of:

positive
negative
mixed
neutral

IMPORTANCE
==========

Must be exactly one of:

high
medium
low

SUMMARY
=======

Give a concise factual summary.

Maximum 9 sentences.

ASSET CLASSES
=============

Return an array of relevant asset classes.

Use only asset-class names that are clearly supported by the article.

GEOGRAPHIES
===========

Return an array of relevant geographic areas.

Use actual geographic names supported by the article.

SECTORS
=======

Return an array of relevant sectors.

Use actual sector names supported by the article.

INVESTOR IMPACT
===============

Explain the practical investment significance of the article.

Do not provide personalized financial advice.

REASONING
=========

Explain why the article received its relevance, category,
sentiment, and importance classification.

FUND MONITORING RELEVANT
========================

Set true only when the article contains information that could
reasonably matter when monitoring investment funds.

FALSE should be used for clearly irrelevant articles.

ARTICLE ID
==========

Copy the supplied articleId exactly.
"""


def build_prompt(
    article: dict[str, Any],
) -> str:
    article_json = json.dumps(
        article,
        ensure_ascii=False,
        indent=2,
    )

    return (
        SYSTEM_PROMPT
        + "\n\nARTICLE TO ANALYZE:\n"
        + article_json
    )


# ==============================================================
# GEMINI RESPONSE SCHEMA
# ==============================================================

RESPONSE_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "articleId": {
            "type": "STRING",
        },
        "relevant": {
            "type": "BOOLEAN",
        },
        "category": {
            "type": "STRING",
            "enum": [
                "MARKET",
                "ECONOMIC",
                "TECHNOLOGY",
                "GEOPOLITICAL",
                "REJECT",
            ],
        },
        "sentiment": {
            "type": "STRING",
            "enum": [
                "positive",
                "negative",
                "mixed",
                "neutral",
            ],
        },
        "importance": {
            "type": "STRING",
            "enum": [
                "high",
                "medium",
                "low",
            ],
        },
        "summary": {
            "type": "STRING",
        },
        "assetClasses": {
            "type": "ARRAY",
            "items": {
                "type": "STRING",
            },
        },
        "geographies": {
            "type": "ARRAY",
            "items": {
                "type": "STRING",
            },
        },
        "sectors": {
            "type": "ARRAY",
            "items": {
                "type": "STRING",
            },
        },
        "investorImpact": {
            "type": "STRING",
        },
        "reasoning": {
            "type": "STRING",
        },
        "fundMonitoringRelevant": {
            "type": "BOOLEAN",
        },
    },
    "required": [
        "articleId",
        "relevant",
        "category",
        "sentiment",
        "importance",
        "summary",
        "assetClasses",
        "geographies",
        "sectors",
        "investorImpact",
        "reasoning",
        "fundMonitoringRelevant",
    ],
}


# ==============================================================
# GEMINI REQUEST
# ==============================================================

def call_gemini(
    article: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    """
    Make exactly one Gemini request.

    There is intentionally NO retry logic here.

    Any HTTP/API failure raises immediately.
    """

    if not API_KEY:
        fail(
            "GEMINI_API_KEY environment variable is missing."
        )

    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": build_prompt(article),
                    }
                ],
            }
        ],
        "generationConfig": {
            "temperature": 0,
            "responseMimeType": "application/json",
            "responseSchema": RESPONSE_SCHEMA,
        },
    }

    body = json.dumps(
        payload,
        ensure_ascii=False,
    ).encode("utf-8")

    url = (
        GEMINI_API_URL
        + "?key="
        + API_KEY
    )

    request = Request(
        url=url,
        data=body,
        method="POST",
        headers={
            "Content-Type": "application/json",
        },
    )

    log(
        f"Sending Gemini request: "
        f"articleId={article.get('articleId')}"
    )

    try:
        with urlopen(
            request,
            timeout=REQUEST_TIMEOUT_SECONDS,
        ) as response:

            status = response.status
            response_body = response.read().decode(
                "utf-8"
            )

    except HTTPError as exc:
        error_body = ""

        try:
            error_body = exc.read().decode(
                "utf-8",
                errors="replace",
            )
        except Exception:
            pass

        # IMPORTANT:
        # No retry is performed for ANY HTTP error.
        if exc.code == 429:
            fail(
                "Gemini API rate/quota limit reached "
                f"(HTTP 429). Stopping immediately.\n"
                f"{error_body}"
            )

        if exc.code == 503:
            fail(
                "Gemini service unavailable "
                f"(HTTP 503). Stopping immediately.\n"
                f"{error_body}"
            )

        fail(
            f"Gemini API HTTP error {exc.code}. "
            "Stopping immediately.\n"
            f"{error_body}"
        )

    except URLError as exc:
        fail(
            "Gemini network error. "
            "Stopping immediately: "
            f"{exc}"
        )

    except TimeoutError as exc:
        fail(
            "Gemini request timed out. "
            "Stopping immediately: "
            f"{exc}"
        )

    except Exception as exc:
        fail(
            "Unexpected Gemini request error. "
            "Stopping immediately: "
            f"{type(exc).__name__}: {exc}"
        )

    if status < 200 or status >= 300:
        fail(
            f"Unexpected Gemini HTTP status {status}. "
            "Stopping immediately."
        )

    try:
        envelope = json.loads(
            response_body
        )
    except json.JSONDecodeError as exc:
        fail(
            "Gemini returned invalid JSON envelope. "
            "Stopping immediately: "
            f"{exc}"
        )

    if not isinstance(envelope, dict):
        fail(
            "Gemini response envelope is not an object. "
            "Stopping immediately."
        )

    # Gemini API may return an explicit error object even with an
    # unexpected successful HTTP status.
    if "error" in envelope:
        fail(
            "Gemini returned an API error. "
            "Stopping immediately: "
            + json.dumps(
                envelope["error"],
                ensure_ascii=False,
            )
        )

    candidates = envelope.get("candidates")

    if not isinstance(candidates, list) or not candidates:
        fail(
            "Gemini response contains no candidates. "
            "Stopping immediately."
        )

    candidate = candidates[0]

    if not isinstance(candidate, dict):
        fail(
            "Gemini candidate is invalid. "
            "Stopping immediately."
        )

    content = candidate.get("content")

    if not isinstance(content, dict):
        fail(
            "Gemini response contains no content. "
            "Stopping immediately."
        )

    parts = content.get("parts")

    if not isinstance(parts, list) or not parts:
        fail(
            "Gemini response contains no parts. "
            "Stopping immediately."
        )

    text_parts: list[str] = []

    for part in parts:
        if not isinstance(part, dict):
            continue

        text_value = part.get("text")

        if isinstance(text_value, str):
            text_parts.append(text_value)

    if not text_parts:
        fail(
            "Gemini response contains no text. "
            "Stopping immediately."
        )

    generated_text = "\n".join(
        text_parts
    ).strip()

    try:
        analysis = json.loads(
            generated_text
        )
    except json.JSONDecodeError as exc:
        fail(
            "Gemini generated invalid analysis JSON. "
            "Stopping immediately: "
            f"{exc}\n"
            f"Response:\n{generated_text}"
        )

    usage = envelope.get(
        "usageMetadata",
        {},
    )

    if not isinstance(usage, dict):
        usage = {}

    metadata = {
        "httpStatus": status,
        "usage": usage,
        "model": MODEL,
    }

    return analysis, metadata


# ==============================================================
# VALIDATION
# ==============================================================

REQUIRED_ANALYSIS_FIELDS = {
    "articleId",
    "relevant",
    "category",
    "sentiment",
    "importance",
    "summary",
    "assetClasses",
    "geographies",
    "sectors",
    "investorImpact",
    "reasoning",
    "fundMonitoringRelevant",
}


def validate_string(
    record: dict[str, Any],
    field: str,
) -> None:
    value = record.get(field)

    if not isinstance(value, str):
        fail(
            f"Validation failed: '{field}' must be a string."
        )

    if not value.strip():
        fail(
            f"Validation failed: '{field}' cannot be empty."
        )


def validate_string_array(
    record: dict[str, Any],
    field: str,
) -> None:
    value = record.get(field)

    if not isinstance(value, list):
        fail(
            f"Validation failed: '{field}' must be an array."
        )

    for item in value:
        if not isinstance(item, str):
            fail(
                f"Validation failed: all values in "
                f"'{field}' must be strings."
            )

        if not item.strip():
            fail(
                f"Validation failed: '{field}' contains "
                "an empty string."
            )


def validate_analysis(
    analysis: Any,
    article: dict[str, Any],
) -> dict[str, Any]:
    if not isinstance(analysis, dict):
        fail(
            "Gemini analysis is not a JSON object. "
            "Stopping immediately."
        )

    actual_fields = set(
        analysis.keys()
    )

    missing = (
        REQUIRED_ANALYSIS_FIELDS
        - actual_fields
    )

    if missing:
        fail(
            "Gemini analysis is missing required fields: "
            + ", ".join(sorted(missing))
        )

    extra = (
        actual_fields
        - REQUIRED_ANALYSIS_FIELDS
    )

    if extra:
        fail(
            "Gemini analysis contains unexpected fields: "
            + ", ".join(sorted(extra))
        )

    # ----------------------------------------------------------
    # articleId
    # ----------------------------------------------------------

    validate_string(
        analysis,
        "articleId",
    )

    expected_id = article.get(
        "articleId"
    )

    if analysis["articleId"] != expected_id:
        fail(
            "Validation failed: Gemini articleId does not match "
            "the source articleId."
        )

    # ----------------------------------------------------------
    # booleans
    # ----------------------------------------------------------

    if not isinstance(
        analysis["relevant"],
        bool,
    ):
        fail(
            "Validation failed: 'relevant' must be boolean."
        )

    if not isinstance(
        analysis["fundMonitoringRelevant"],
        bool,
    ):
        fail(
            "Validation failed: "
            "'fundMonitoringRelevant' must be boolean."
        )

    # ----------------------------------------------------------
    # enums
    # ----------------------------------------------------------

    if analysis["category"] not in ALLOWED_CATEGORIES:
        fail(
            "Validation failed: invalid category "
            f"'{analysis['category']}'."
        )

    if analysis["sentiment"] not in ALLOWED_SENTIMENTS:
        fail(
            "Validation failed: invalid sentiment "
            f"'{analysis['sentiment']}'."
        )

    if analysis["importance"] not in ALLOWED_IMPORTANCE:
        fail(
            "Validation failed: invalid importance "
            f"'{analysis['importance']}'."
        )

    # ----------------------------------------------------------
    # text fields
    # ----------------------------------------------------------

    validate_string(
        analysis,
        "summary",
    )

    validate_string(
        analysis,
        "investorImpact",
    )

    validate_string(
        analysis,
        "reasoning",
    )

    # ----------------------------------------------------------
    # summary sentence limit
    # ----------------------------------------------------------

    summary = analysis["summary"].strip()

    sentence_count = sum(
        1
        for character in summary
        if character in ".!?"
    )

    if sentence_count > 9:
        fail(
            "Validation failed: summary exceeds the "
            "maximum of 9 sentences."
        )

    # ----------------------------------------------------------
    # arrays
    # ----------------------------------------------------------

    validate_string_array(
        analysis,
        "assetClasses",
    )

    validate_string_array(
        analysis,
        "geographies",
    )

    validate_string_array(
        analysis,
        "sectors",
    )

    return analysis


# ==============================================================
# RECORD CONSTRUCTION
# ==============================================================

def build_analysis_record(
    article: dict[str, Any],
    analysis: dict[str, Any],
    metadata: dict[str, Any],
) -> dict[str, Any]:
    """
    Preserve the source article identity alongside Gemini's
    analysis.

    The source fields are copied rather than modifying the
    collector's current.json.
    """

    record = dict(analysis)

    record["title"] = article.get(
        "title"
    )

    record["url"] = article.get(
        "url"
    )

    record["published"] = article.get(
        "published"
    )

    record["source"] = article.get(
        "source"
    )

    record["analysisGeneratedAt"] = (
        singapore_now().isoformat()
    )

    record["analysisModel"] = MODEL

    record["geminiUsage"] = metadata.get(
        "usage",
        {},
    )

    return record


# ==============================================================
# IMMEDIATE SAVE
# ==============================================================

def save_successful_analysis(
    current_analysis: dict[str, Any],
    record: dict[str, Any],
) -> None:
    """
    Append exactly one successful record and immediately persist
    the entire current analysis file.

    This function is called BEFORE the next Gemini request.
    """

    analyses = current_analysis.setdefault(
        "analyses",
        [],
    )

    article_id = record["articleId"]

    existing_ids = {
        item.get("articleId")
        for item in analyses
        if isinstance(item, dict)
    }

    if article_id in existing_ids:
        fail(
            "Attempted to save duplicate articleId: "
            f"{article_id}"
        )

    analyses.append(record)

    current_analysis["generatedAt"] = (
        singapore_now().isoformat()
    )

    write_json_atomic(
        ANALYSIS_CURRENT_FILE,
        current_analysis,
    )

    log(
        "SUCCESSFULLY SAVED analysis: "
        f"{article_id}"
    )


# ==============================================================
# HISTORY HELPERS
# ==============================================================

def history_path_for_date(
    archive_date: date,
) -> Path:
    return (
        HISTORY_DIR
        / f"{archive_date.year:04d}"
        / f"{archive_date.month:02d}"
        / f"{archive_date.isoformat()}.json"
    )


def load_history_file(
    path: Path,
) -> dict[str, Any]:
    if not path.exists():
        return {
            "date": path.stem,
            "source": "Gemini",
            "analyses": [],
        }

    data = read_json(path)

    if not isinstance(data, dict):
        fail(
            f"History file must contain an object: {path}"
        )

    analyses = data.get(
        "analyses",
        [],
    )

    if not isinstance(analyses, list):
        fail(
            f"History 'analyses' must be an array: {path}"
        )

    return data


def archive_record(
    record: dict[str, Any],
    archive_date: date,
) -> None:
    path = history_path_for_date(
        archive_date
    )

    history = load_history_file(
        path
    )

    analyses = history.setdefault(
        "analyses",
        [],
    )

    article_id = record.get(
        "articleId"
    )

    existing_ids = {
        item.get("articleId")
        for item in analyses
        if isinstance(item, dict)
    }

    if article_id in existing_ids:
        fail(
            "Duplicate article already exists in archive: "
            f"{article_id}"
        )

    analyses.append(record)

    history["date"] = archive_date.isoformat()
    history["source"] = "Gemini"

    write_json_atomic(
        path,
        history,
    )

    log(
        f"ARCHIVED {article_id} → "
        f"{path.relative_to(ROOT)}"
    )


# ==============================================================
# ROLLING 14-DAY MAINTENANCE
# ==============================================================

def archive_old_current_records(
    current_analysis: dict[str, Any],
) -> None:
    """
    Move current records outside the 14-day rolling publication
    window into daily history.

    The publication date determines the archive location.

    If a record has no valid publication date, it is retained in
    current.json rather than guessing an archive date.
    """

    today_sg = singapore_now().date()

    earliest_current_date = (
        today_sg
        - timedelta(days=ROLLING_DAYS - 1)
    )

    current_records = current_analysis.get(
        "analyses",
        [],
    )

    retained: list[dict[str, Any]] = []

    moved = 0

    for record in current_records:
        if not isinstance(record, dict):
            fail(
                "Current analysis contains an invalid record."
            )

        archive_date = publication_date(
            record
        )

        if archive_date is None:
            # Never guess publication date.
            retained.append(record)
            continue

        if archive_date < earliest_current_date:
            archive_record(
                record,
                archive_date,
            )
            moved += 1
        else:
            retained.append(record)

    current_analysis["analyses"] = retained

    if moved > 0:
        current_analysis["generatedAt"] = (
            singapore_now().isoformat()
        )

        write_json_atomic(
            ANALYSIS_CURRENT_FILE,
            current_analysis,
        )

        log(
            f"Archived {moved} article(s) outside "
            f"the {ROLLING_DAYS}-day rolling window."
        )


# ==============================================================
# DEDUPLICATION
# ==============================================================

def build_processed_ids(
    current_analysis: dict[str, Any],
) -> set[str]:
    log("Scanning existing analysis history...")

    ids = current_article_ids(
        current_analysis
    )

    archived_ids = collect_archived_article_ids()

    ids.update(
        archived_ids
    )

    log(
        f"Existing successful analyses: {len(ids)}"
    )

    return ids


# ==============================================================
# MAIN
# ==============================================================

def main() -> int:
    started_at = singapore_now()

    log("=" * 70)
    log("VGrat FMS - MARKET NEWS AI ANALYZER")
    log("=" * 70)

    log(
        f"Model: {MODEL}"
    )

    log(
        f"Input: {NEWS_FILE.relative_to(ROOT)}"
    )

    log(
        f"Output: {ANALYSIS_CURRENT_FILE.relative_to(ROOT)}"
    )

    log(
        "Mode: sequential / 1 article per request / newest first"
    )

    log(
        "Retry: DISABLED"
    )

    log(
        "Failure behaviour: STOP IMMEDIATELY"
    )

    # ----------------------------------------------------------
    # Validate environment
    # ----------------------------------------------------------

    if not API_KEY:
        fail(
            "GEMINI_API_KEY is not configured."
        )

    # ----------------------------------------------------------
    # Load collector input
    #
    # IMPORTANT:
    # This file is NEVER written by this script.
    # ----------------------------------------------------------

    log("Loading collector current.json...")

    raw_news = read_json(
        NEWS_FILE
    )

    articles = extract_articles(
        raw_news
    )

    log(
        f"Collector contains {len(articles)} article(s)."
    )

    # ----------------------------------------------------------
    # Load analysis state
    # ----------------------------------------------------------

    current_analysis = load_current_analysis()

    # ----------------------------------------------------------
    # Archive records that have already crossed the cutoff.
    #
    # This happens before new analysis so current.json remains
    # a clean rolling 14-day file.
    # ----------------------------------------------------------

    archive_old_current_records(
        current_analysis
    )

    # ----------------------------------------------------------
    # Build permanent deduplication index
    # ----------------------------------------------------------

    processed_ids = build_processed_ids(
        current_analysis
    )

    # ----------------------------------------------------------
    # Determine new articles
    # ----------------------------------------------------------

    new_articles: list[dict[str, Any]] = []

    skipped = 0

    for article in articles:
        article_id = article["articleId"]

        if article_id in processed_ids:
            skipped += 1
            continue

        new_articles.append(
            article
        )

    # ----------------------------------------------------------
    # Newest first
    # ----------------------------------------------------------

    new_articles.sort(
        key=publication_sort_key,
        reverse=True,
    )

    log(
        f"Already analyzed: {skipped}"
    )

    log(
        f"New articles requiring analysis: "
        f"{len(new_articles)}"
    )

    if not new_articles:
        log("No new articles to analyze.")

        # Ensure current file exists even on an empty first run.
        if not ANALYSIS_CURRENT_FILE.exists():
            current_analysis["generatedAt"] = (
                singapore_now().isoformat()
            )

            write_json_atomic(
                ANALYSIS_CURRENT_FILE,
                current_analysis,
            )

        log("Nothing else to do.")
        return 0

    # ----------------------------------------------------------
    # SEQUENTIAL ANALYSIS LOOP
    #
    # There is intentionally no try/except around individual
    # requests that continues to the next article.
    #
    # Any failure propagates immediately and terminates main().
    # ----------------------------------------------------------

    successful = 0

    for position, article in enumerate(
        new_articles,
        start=1,
    ):
        article_id = article["articleId"]

        published = article.get(
            "published"
        )

        title = article.get(
            "title",
            "",
        )

        log("")
        log("-" * 70)
        log(
            f"ARTICLE {position}/{len(new_articles)}"
        )
        log(
            f"articleId: {article_id}"
        )
        log(
            f"published: {published}"
        )
        log(
            f"title: {title}"
        )
        log("-" * 70)

        # ------------------------------------------------------
        # EXACTLY ONE GEMINI REQUEST
        # ------------------------------------------------------

        analysis, metadata = call_gemini(
            article
        )

        # ------------------------------------------------------
        # VALIDATE BEFORE SAVING
        # ------------------------------------------------------

        validated = validate_analysis(
            analysis,
            article,
        )

        # ------------------------------------------------------
        # BUILD COMPLETE RECORD
        # ------------------------------------------------------

        record = build_analysis_record(
            article,
            validated,
            metadata,
        )

        # ------------------------------------------------------
        # IMMEDIATE SAVE
        #
        # Nothing else is requested from Gemini until this
        # function completes successfully.
        # ------------------------------------------------------

        save_successful_analysis(
            current_analysis,
            record,
        )

        processed_ids.add(
            article_id
        )

        successful += 1

        log(
            f"Progress: {successful}/"
            f"{len(new_articles)} successful."
        )

        # ------------------------------------------------------
        # Tiny local pause only.
        #
        # This is NOT retry logic.
        # It does not repeat a request.
        #
        # It simply gives the filesystem a moment before the
        # next request. It can also reduce burstiness.
        # ------------------------------------------------------

        if position < len(new_articles):
            time.sleep(0.2)

    # ----------------------------------------------------------
    # Final rolling-window maintenance
    # ----------------------------------------------------------

    archive_old_current_records(
        current_analysis
    )

    # ----------------------------------------------------------
    # FINAL SUMMARY
    # ----------------------------------------------------------

    elapsed = (
        singapore_now()
        - started_at
    ).total_seconds()

    log("")
    log("=" * 70)
    log("ANALYSIS COMPLETE")
    log("=" * 70)
    log(
        f"Successful: {successful}"
    )
    log(
        f"Skipped already analyzed: {skipped}"
    )
    log(
        f"Elapsed seconds: {elapsed:.1f}"
    )
    log("=" * 70)

    return 0


# ==============================================================
# ENTRY POINT
# ==============================================================

if __name__ == "__main__":
    try:
        exit_code = main()

    except AnalysisError as exc:
        log("")
        log("=" * 70)
        log("ANALYZER STOPPED")
        log("=" * 70)
        log(str(exc))
        log("=" * 70)
        log(
            "No further Gemini requests will be made."
        )

        # IMPORTANT:
        # Return non-zero so GitHub Actions marks the analyzer job
        # as failed. The following workflow commit step uses
        # `if: always()` and therefore still commits any successful
        # records already written.
        sys.exit(1)

    except KeyboardInterrupt:
        log("")
        log(
            "Interrupted by user. "
            "No further requests will be made."
        )
        sys.exit(1)

    except Exception as exc:
        log("")
        log("=" * 70)
        log("UNEXPECTED CODE ERROR")
        log("=" * 70)
        log(
            f"{type(exc).__name__}: {exc}"
        )
        log("=" * 70)
        log(
            "No further Gemini requests will be made."
        )

        # Unexpected Python errors are also fatal.
        sys.exit(1)

    sys.exit(exit_code)
