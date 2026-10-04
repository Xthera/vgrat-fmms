#!/usr/bin/env python3

"""
VGrat FMS - MARKET NEWS AI ANALYZER
===================================

Sequential, one-article-per-request Gemini analyzer.

INPUT
=====

    data/market_news/current.json

The collector-owned current.json is READ-ONLY.

Actual collector article schema:

    {
        "id": "...",
        "source": "...",
        "title": "...",
        "publishedAtSgt": "...",
        "url": "..."
    }

OUTPUT
======

Rolling 14-day analysis:

    data/market_news/analysis/current.json

Permanent history:

    data/market_news/analysis/history/
        YYYY/
            MM/
                YYYY-MM-DD.json

PROCESSING RULES
================

1. Read current.json.
2. Process newest articles first.
3. Send exactly ONE article per Gemini request.
4. Never batch multiple articles into one request.
5. Validate every Gemini response.
6. Save every successful analysis IMMEDIATELY.
7. Only after the successful save, send the next request.
8. HTTP 429 -> stop immediately.
9. HTTP 503 -> stop immediately.
10. Any other HTTP/API error -> stop immediately.
11. Invalid Gemini JSON -> stop immediately.
12. Failed validation -> stop immediately.
13. Unexpected Python/code error -> stop immediately.
14. NO automatic retry.
15. Never continue after an error.
16. Never modify collector current.json.
17. Never analyze an article whose ID already exists in current
    analysis or archived history.
18. Keep the latest 14 publication-date days in analysis/current.json.
19. Archive older successful analyses using the article publication
    date.
20. History files are append-only and existing records are preserved.

IMPORTANT
=========

The save operation happens after EACH successful Gemini request.

Therefore:

    Article 1 -> Gemini -> validate -> SAVE
    Article 2 -> Gemini -> validate -> SAVE
    Article 3 -> Gemini -> validate -> SAVE
    ...

If Article 13 receives HTTP 429:

    Articles 1-12 remain safely written to disk.
    Article 13 is not saved.
    No Article 14 request is made.

The GitHub Actions workflow can then use `if: always()` on its
commit step to commit the successful work.
"""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime, date, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


# =============================================================================
# PATHS
# =============================================================================

ROOT = Path(__file__).resolve().parent.parent

NEWS_FILE = (
    ROOT
    / "data"
    / "market_news"
    / "current.json"
)

ANALYSIS_DIR = (
    ROOT
    / "data"
    / "market_news"
    / "analysis"
)

ANALYSIS_CURRENT_FILE = (
    ANALYSIS_DIR
    / "current.json"
)

HISTORY_DIR = (
    ANALYSIS_DIR
    / "history"
)


# =============================================================================
# CONFIGURATION
# =============================================================================

DEFAULT_MODEL = "gemini-3.5-flash-lite"

MODEL = os.environ.get(
    "MARKET_NEWS_MODEL",
    DEFAULT_MODEL,
).strip()

API_KEY = os.environ.get(
    "GEMINI_API_KEY",
    "",
).strip()

GEMINI_API_URL = (
    "https://generativelanguage.googleapis.com/"
    "v1beta/models/"
    f"{MODEL}:generateContent"
)

REQUEST_TIMEOUT_SECONDS = 120

ROLLING_DAYS = 14


# =============================================================================
# ALLOWED VALUES
# =============================================================================

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


# =============================================================================
# EXCEPTIONS
# =============================================================================

class AnalysisError(Exception):
    """
    Fatal analyzer error.

    Any AnalysisError stops the entire run.
    There is deliberately no retry or continue behaviour.
    """


# =============================================================================
# LOGGING
# =============================================================================

def log(message: str) -> None:
    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    print(
        f"[{timestamp}] {message}",
        flush=True,
    )


def stop(message: str) -> None:
    raise AnalysisError(message)


# =============================================================================
# JSON
# =============================================================================

def read_json(path: Path) -> Any:
    if not path.exists():
        stop(
            f"Required file does not exist: {path}"
        )

    try:
        with path.open(
            "r",
            encoding="utf-8",
        ) as file:
            return json.load(file)

    except json.JSONDecodeError as exc:
        stop(
            f"Invalid JSON in {path}: {exc}"
        )

    except OSError as exc:
        stop(
            f"Unable to read {path}: {exc}"
        )

    return None


def write_json_atomic(
    path: Path,
    data: Any,
) -> None:
    """
    Atomically replace the destination file.

    This prevents a partially-written JSON file if the process
    is interrupted during a save.
    """

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary_path = path.with_name(
        f".{path.name}.tmp"
    )

    try:
        with temporary_path.open(
            "w",
            encoding="utf-8",
        ) as file:

            json.dump(
                data,
                file,
                ensure_ascii=False,
                indent=2,
            )

            file.write("\n")

        temporary_path.replace(path)

    except OSError as exc:

        try:
            if temporary_path.exists():
                temporary_path.unlink()
        except OSError:
            pass

        stop(
            f"Unable to write {path}: {exc}"
        )


# =============================================================================
# SINGAPORE TIME
# =============================================================================

SINGAPORE_TZ = timezone(
    timedelta(hours=8)
)


def singapore_now() -> datetime:
    return datetime.now(
        timezone.utc
    ).astimezone(
        SINGAPORE_TZ
    )


# =============================================================================
# PUBLICATION DATE
# =============================================================================

def parse_publication_datetime(
    value: Any,
) -> datetime | None:
    """
    Parse the collector's publishedAtSgt value.

    Handles ISO 8601 timestamps including:
        2026-10-04T23:00:00+08:00
        2026-10-04T23:00:00Z
        2026-10-04 23:00:00+08:00
    """

    if value is None:
        return None

    if isinstance(
        value,
        datetime,
    ):
        result = value

    elif isinstance(
        value,
        str,
    ):
        text_value = value.strip()

        if not text_value:
            return None

        if text_value.endswith("Z"):
            text_value = (
                text_value[:-1]
                + "+00:00"
            )

        try:
            result = datetime.fromisoformat(
                text_value
            )

        except ValueError:

            # Fallback for RFC-style timestamps.
            try:
                from email.utils import (
                    parsedate_to_datetime,
                )

                result = parsedate_to_datetime(
                    text_value
                )

            except (
                TypeError,
                ValueError,
                OverflowError,
            ):
                return None

    else:
        return None

    if result.tzinfo is None:
        result = result.replace(
            tzinfo=SINGAPORE_TZ
        )

    return result


def publication_date(
    article: dict[str, Any],
) -> date | None:
    """
    Return the article publication date in Singapore calendar time.

    IMPORTANT:
    The collector field `publishedAtSgt` is used.
    """

    value = article.get(
        "publishedAtSgt"
    )

    parsed = parse_publication_datetime(
        value
    )

    if parsed is None:
        return None

    return parsed.astimezone(
        SINGAPORE_TZ
    ).date()


def publication_sort_key(
    article: dict[str, Any],
) -> tuple[int, str, str]:
    """
    Sort newest publication first.

    Articles with valid publication dates come first.

    The article ID is used as a deterministic tie-breaker.
    """

    parsed = parse_publication_datetime(
        article.get("publishedAtSgt")
    )

    article_id = str(
        article.get(
            "id",
            "",
        )
    )

    if parsed is None:
        return (
            0,
            "",
            article_id,
        )

    return (
        1,
        parsed.astimezone(
            timezone.utc
        ).isoformat(),
        article_id,
    )


# =============================================================================
# COLLECTOR INPUT
# =============================================================================

def load_collector_articles() -> list[dict[str, Any]]:
    """
    Load the collector's current.json.

    This function ONLY reads the file.
    It never writes to it.
    """

    log(
        "Loading collector current.json..."
    )

    data = read_json(
        NEWS_FILE
    )

    if not isinstance(
        data,
        dict,
    ):
        stop(
            "current.json root must be an object."
        )

    articles = data.get(
        "articles"
    )

    if not isinstance(
        articles,
        list,
    ):
        stop(
            "current.json 'articles' must be an array."
        )

    result: list[
        dict[str, Any]
    ] = []

    seen_ids: set[str] = set()

    for index, article in enumerate(
        articles
    ):

        if not isinstance(
            article,
            dict,
        ):
            stop(
                f"Article at index {index} "
                "is not an object."
            )

        article_id = article.get(
            "id"
        )

        if not isinstance(
            article_id,
            str,
        ) or not article_id.strip():

            stop(
                f"Article at index {index} "
                "has no valid id."
            )

        article_id = article_id.strip()

        if article_id in seen_ids:
            stop(
                f"Duplicate article ID in collector: "
                f"{article_id}"
            )

        seen_ids.add(
            article_id
        )

        # ----------------------------------------------------------
        # Match the collector's own validation requirements.
        # ----------------------------------------------------------

        for field in (
            "source",
            "title",
            "publishedAtSgt",
            "url",
        ):
            value = article.get(
                field
            )

            if not value:
                stop(
                    f"Article {article_id} "
                    f"missing required field: {field}"
                )

        # ----------------------------------------------------------
        # Publication timestamp must be parseable because it is
        # required for newest-first ordering and 14-day archiving.
        # ----------------------------------------------------------

        if publication_date(
            article
        ) is None:
            stop(
                f"Article {article_id} has an invalid "
                "publishedAtSgt value."
            )

        result.append(
            article
        )

    log(
        f"Loaded {len(result)} collector article(s)."
    )

    return result


# =============================================================================
# CURRENT ANALYSIS FILE
# =============================================================================

def empty_analysis_document() -> dict[str, Any]:
    return {
        "generatedAtSgt": None,
        "source": "Gemini",
        "model": MODEL,
        "rollingDays": ROLLING_DAYS,
        "articlesPerRequest": 1,
        "processingOrder": "newest_first",
        "retry": False,
        "analyses": [],
    }


def load_current_analysis() -> dict[str, Any]:
    if not ANALYSIS_CURRENT_FILE.exists():
        log(
            "analysis/current.json does not exist. "
            "Creating a new analysis store."
        )

        return empty_analysis_document()

    log(
        "Loading existing analysis/current.json..."
    )

    data = read_json(
        ANALYSIS_CURRENT_FILE
    )

    if not isinstance(
        data,
        dict,
    ):
        stop(
            "analysis/current.json root must be an object."
        )

    analyses = data.get(
        "analyses"
    )

    if analyses is None:
        data["analyses"] = []

    elif not isinstance(
        analyses,
        list,
    ):
        stop(
            "analysis/current.json 'analyses' "
            "must be an array."
        )

    return data


# =============================================================================
# HISTORY
# =============================================================================

def history_files() -> list[Path]:
    if not HISTORY_DIR.exists():
        return []

    return sorted(
        HISTORY_DIR.rglob(
            "*.json"
        )
    )


def read_history_article_ids() -> set[str]:
    """
    Build a permanent ID index from archived analyses.

    This prevents an article from ever being analyzed twice,
    even after it has left the rolling current analysis file.
    """

    ids: set[str] = set()

    files = history_files()

    log(
        f"Scanning {len(files)} history file(s)..."
    )

    for path in files:

        data = read_json(
            path
        )

        if isinstance(
            data,
            dict,
        ):
            analyses = data.get(
                "analyses",
                [],
            )

        elif isinstance(
            data,
            list,
        ):
            analyses = data

        else:
            stop(
                f"Invalid history JSON structure: {path}"
            )

        if not isinstance(
            analyses,
            list,
        ):
            stop(
                f"History analyses must be an array: {path}"
            )

        for record in analyses:

            if not isinstance(
                record,
                dict,
            ):
                stop(
                    f"Invalid analysis record in {path}"
                )

            article_id = record.get(
                "articleId"
            )

            if not isinstance(
                article_id,
                str,
            ) or not article_id.strip():

                stop(
                    f"History record missing articleId: {path}"
                )

            ids.add(
                article_id.strip()
            )

    return ids


def current_analysis_article_ids(
    analysis_document: dict[str, Any],
) -> set[str]:

    ids: set[str] = set()

    for record in analysis_document.get(
        "analyses",
        [],
    ):

        if not isinstance(
            record,
            dict,
        ):
            stop(
                "analysis/current.json contains "
                "a non-object analysis record."
            )

        article_id = record.get(
            "articleId"
        )

        if not isinstance(
            article_id,
            str,
        ) or not article_id.strip():

            stop(
                "analysis/current.json contains "
                "a record without articleId."
            )

        ids.add(
            article_id.strip()
        )

    return ids


def load_history_file(
    path: Path,
) -> dict[str, Any]:

    if not path.exists():
        return {
            "date": path.stem,
            "source": "Gemini",
            "analyses": [],
        }

    data = read_json(
        path
    )

    if not isinstance(
        data,
        dict,
    ):
        stop(
            f"History file must contain an object: {path}"
        )

    analyses = data.get(
        "analyses",
        [],
    )

    if not isinstance(
        analyses,
        list,
    ):
        stop(
            f"History 'analyses' must be an array: {path}"
        )

    return data


def history_path_for_date(
    archive_date: date,
) -> Path:

    return (
        HISTORY_DIR
        / f"{archive_date.year:04d}"
        / f"{archive_date.month:02d}"
        / f"{archive_date.isoformat()}.json"
    )


# =============================================================================
# GEMINI PROMPT
# =============================================================================

SYSTEM_PROMPT = """
You are the market-news analysis engine for VGrat FMS.

Analyze exactly ONE article.

Use ONLY information contained in the supplied article.

Do not invent facts.
Do not infer unsupported facts.
Do not fabricate companies, countries, sectors,
asset classes, market impacts, or events.

Return ONLY a single valid JSON object.

Required JSON structure:

{
  "articleId": "string",
  "relevant": true,
  "category": "MARKET",
  "sentiment": "neutral",
  "importance": "medium",
  "summary": "string",
  "assetClasses": [],
  "geographies": [],
  "sectors": [],
  "investorImpact": "string",
  "reasoning": "string",
  "fundMonitoringRelevant": false
}

FIELD RULES
===========

articleId
---------
Copy the supplied article ID exactly.

relevant
--------
Boolean.

Set true when the article has meaningful relevance to:
- financial markets
- investment markets
- economics
- technology affecting markets
- geopolitical developments affecting markets

Otherwise false.

category
--------
Must be exactly one of:

MARKET
ECONOMIC
TECHNOLOGY
GEOPOLITICAL
REJECT

Use REJECT when the article is not meaningfully relevant to
financial-market, investment, economic, technology-market, or
market-relevant geopolitical developments.

sentiment
---------
Must be exactly one of:

positive
negative
mixed
neutral

importance
----------
Must be exactly one of:

high
medium
low

summary
-------
Concise factual summary of the article.

Maximum 9 sentences.

assetClasses
------------
Array of strings.

Include only asset classes clearly relevant to the article.

Examples may include:
- Equities
- Bonds
- Commodities
- Currencies
- Real Estate
- Alternatives

Do not invent exposure.

geographies
-----------
Array of strings.

Include geographic markets or regions clearly supported by
the article.

sectors
-------
Array of strings.

Include sectors clearly supported by the article.

investorImpact
--------------
Explain the practical significance of the article for investors
and financial-market monitoring.

Do not provide personalized financial advice.

reasoning
---------
Briefly explain why the article received its relevance,
category, sentiment, and importance classifications.

fundMonitoringRelevant
----------------------
Boolean.

Set true only when the article could reasonably matter when
monitoring investment funds.

Set false for clearly irrelevant articles.
"""


# =============================================================================
# GEMINI RESPONSE SCHEMA
# =============================================================================

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


def build_gemini_prompt(
    article: dict[str, Any],
) -> str:

    article_payload = {
        "id": article.get("id"),
        "source": article.get("source"),
        "title": article.get("title"),
        "publishedAtSgt": article.get(
            "publishedAtSgt"
        ),
        "url": article.get("url"),
    }

    return (
        SYSTEM_PROMPT
        + "\n\nARTICLE:\n"
        + json.dumps(
            article_payload,
            ensure_ascii=False,
            indent=2,
        )
    )


# =============================================================================
# GEMINI REQUEST
# =============================================================================

def call_gemini(
    article: dict[str, Any],
) -> tuple[
    dict[str, Any],
    dict[str, Any],
]:
    """
    Make exactly ONE Gemini API request.

    There is deliberately NO retry logic.

    Any failure raises AnalysisError and terminates the entire
    analyzer.
    """

    if not API_KEY:
        stop(
            "GEMINI_API_KEY environment variable is missing."
        )

    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": build_gemini_prompt(
                            article
                        ),
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

    request_body = json.dumps(
        payload,
        ensure_ascii=False,
    ).encode(
        "utf-8"
    )

    request_url = (
        GEMINI_API_URL
        + "?key="
        + API_KEY
    )

    request = Request(
        request_url,
        data=request_body,
        method="POST",
        headers={
            "Content-Type": "application/json",
        },
    )

    article_id = article["id"]

    log(
        f"Gemini request START: {article_id}"
    )

    try:

        with urlopen(
            request,
            timeout=REQUEST_TIMEOUT_SECONDS,
        ) as response:

            status = response.status

            response_body = (
                response
                .read()
                .decode(
                    "utf-8"
                )
            )

    except HTTPError as exc:

        error_body = ""

        try:
            error_body = (
                exc
                .read()
                .decode(
                    "utf-8",
                    errors="replace",
                )
            )
        except Exception:
            pass

        # ----------------------------------------------------------
        # IMPORTANT:
        #
        # Every HTTP error is fatal.
        # No retry.
        # No sleep-and-retry.
        # No continuation.
        # ----------------------------------------------------------

        if exc.code == 429:

            stop(
                "Gemini API limit/quota reached "
                "(HTTP 429). "
                "Stopping immediately.\n"
                + error_body
            )

        if exc.code == 503:

            stop(
                "Gemini service unavailable "
                "(HTTP 503). "
                "Stopping immediately.\n"
                + error_body
            )

        stop(
            f"Gemini API HTTP error {exc.code}. "
            "Stopping immediately.\n"
            + error_body
        )

    except URLError as exc:

        stop(
            "Gemini network error. "
            "Stopping immediately: "
            f"{exc}"
        )

    except TimeoutError as exc:

        stop(
            "Gemini request timed out. "
            "Stopping immediately: "
            f"{exc}"
        )

    except Exception as exc:

        stop(
            "Unexpected Gemini request error. "
            "Stopping immediately: "
            f"{type(exc).__name__}: {exc}"
        )

    if status < 200 or status >= 300:
        stop(
            f"Unexpected Gemini HTTP status {status}. "
            "Stopping immediately."
        )

    # -----------------------------------------------------------------
    # Parse Gemini envelope
    # -----------------------------------------------------------------

    try:

        envelope = json.loads(
            response_body
        )

    except json.JSONDecodeError as exc:

        stop(
            "Gemini returned invalid JSON. "
            "Stopping immediately.\n"
            f"JSON error: {exc}\n"
            f"Response: {response_body}"
        )

    if not isinstance(
        envelope,
        dict,
    ):
        stop(
            "Gemini response envelope is not an object. "
            "Stopping immediately."
        )

    if "error" in envelope:
        stop(
            "Gemini returned an API error. "
            "Stopping immediately.\n"
            + json.dumps(
                envelope["error"],
                ensure_ascii=False,
            )
        )

    candidates = envelope.get(
        "candidates"
    )

    if not isinstance(
        candidates,
        list,
    ) or not candidates:

        stop(
            "Gemini returned no candidates. "
            "Stopping immediately."
        )

    candidate = candidates[0]

    if not isinstance(
        candidate,
        dict,
    ):
        stop(
            "Gemini candidate is invalid. "
            "Stopping immediately."
        )

    finish_reason = candidate.get(
        "finishReason"
    )

    # ----------------------------------------------------------
    # A truncated response is not considered successful.
    # ----------------------------------------------------------

    if finish_reason in {
        "MAX_TOKENS",
        "SAFETY",
        "RECITATION",
        "LANGUAGE",
        "BLOCKLIST",
        "PROHIBITED_CONTENT",
    }:

        stop(
            "Gemini response was not completed successfully. "
            f"finishReason={finish_reason}. "
            "Stopping immediately."
        )

    content = candidate.get(
        "content"
    )

    if not isinstance(
        content,
        dict,
    ):
        stop(
            "Gemini candidate has no valid content. "
            "Stopping immediately."
        )

    parts = content.get(
        "parts"
    )

    if not isinstance(
        parts,
        list,
    ) or not parts:

        stop(
            "Gemini candidate has no parts. "
            "Stopping immediately."
        )

    text_parts: list[str] = []

    for part in parts:

        if not isinstance(
            part,
            dict,
        ):
            continue

        text_value = part.get(
            "text"
        )

        if isinstance(
            text_value,
            str,
        ):
            text_parts.append(
                text_value
            )

    if not text_parts:
        stop(
            "Gemini returned no text. "
            "Stopping immediately."
        )

    generated_text = "\n".join(
        text_parts
    ).strip()

    # -----------------------------------------------------------------
    # Parse the actual analysis JSON.
    # -----------------------------------------------------------------

    try:

        analysis = json.loads(
            generated_text
        )

    except json.JSONDecodeError as exc:

        stop(
            "Gemini analysis is invalid JSON. "
            "Stopping immediately.\n"
            f"JSON error: {exc}\n"
            f"Generated response:\n{generated_text}"
        )

    usage = envelope.get(
        "usageMetadata",
        {},
    )

    if not isinstance(
        usage,
        dict,
    ):
        usage = {}

    metadata = {
        "model": MODEL,
        "httpStatus": status,
        "finishReason": finish_reason,
        "usage": usage,
    }

    log(
        f"Gemini request SUCCESS: {article_id}"
    )

    return (
        analysis,
        metadata,
    )


# =============================================================================
# VALIDATION
# =============================================================================

REQUIRED_FIELDS = {
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


def validate_nonempty_string(
    record: dict[str, Any],
    field: str,
) -> None:

    value = record.get(
        field
    )

    if not isinstance(
        value,
        str,
    ):
        stop(
            f"Validation failed: {field} "
            "must be a string."
        )

    if not value.strip():
        stop(
            f"Validation failed: {field} "
            "cannot be empty."
        )


def validate_string_array(
    record: dict[str, Any],
    field: str,
) -> None:

    value = record.get(
        field
    )

    if not isinstance(
        value,
        list,
    ):
        stop(
            f"Validation failed: {field} "
            "must be an array."
        )

    for item in value:

        if not isinstance(
            item,
            str,
        ):
            stop(
                f"Validation failed: all values in "
                f"{field} must be strings."
            )

        if not item.strip():
            stop(
                f"Validation failed: {field} "
                "contains an empty string."
            )


def validate_analysis(
    analysis: Any,
    article: dict[str, Any],
) -> dict[str, Any]:
    """
    Strictly validate the Gemini analysis before it can be saved.
    """

    if not isinstance(
        analysis,
        dict,
    ):
        stop(
            "Gemini analysis is not a JSON object. "
            "Stopping immediately."
        )

    actual_fields = set(
        analysis.keys()
    )

    missing = (
        REQUIRED_FIELDS
        - actual_fields
    )

    if missing:
        stop(
            "Gemini analysis missing required fields: "
            + ", ".join(
                sorted(missing)
            )
        )

    extra = (
        actual_fields
        - REQUIRED_FIELDS
    )

    if extra:
        stop(
            "Gemini analysis contains unexpected fields: "
            + ", ".join(
                sorted(extra)
            )
        )

    # -----------------------------------------------------------------
    # articleId
    # -----------------------------------------------------------------

    validate_nonempty_string(
        analysis,
        "articleId",
    )

    expected_id = article.get(
        "id"
    )

    if analysis["articleId"] != expected_id:

        stop(
            "Validation failed: Gemini articleId "
            "does not match collector article id."
        )

    # -----------------------------------------------------------------
    # Boolean fields
    # -----------------------------------------------------------------

    if not isinstance(
        analysis["relevant"],
        bool,
    ):
        stop(
            "Validation failed: relevant "
            "must be boolean."
        )

    if not isinstance(
        analysis["fundMonitoringRelevant"],
        bool,
    ):
        stop(
            "Validation failed: "
            "fundMonitoringRelevant must be boolean."
        )

    # -----------------------------------------------------------------
    # Enum fields
    # -----------------------------------------------------------------

    if analysis["category"] not in ALLOWED_CATEGORIES:
        stop(
            "Validation failed: invalid category: "
            f"{analysis['category']}"
        )

    if analysis["sentiment"] not in ALLOWED_SENTIMENTS:
        stop(
            "Validation failed: invalid sentiment: "
            f"{analysis['sentiment']}"
        )

    if analysis["importance"] not in ALLOWED_IMPORTANCE:
        stop(
            "Validation failed: invalid importance: "
            f"{analysis['importance']}"
        )

    # -----------------------------------------------------------------
    # Text fields
    # -----------------------------------------------------------------

    validate_nonempty_string(
        analysis,
        "summary",
    )

    validate_nonempty_string(
        analysis,
        "investorImpact",
    )

    validate_nonempty_string(
        analysis,
        "reasoning",
    )

    # -----------------------------------------------------------------
    # Summary length
    # -----------------------------------------------------------------

    summary = analysis["summary"].strip()

    sentence_count = sum(
        1
        for character in summary
        if character in ".!?"
    )

    if sentence_count > 9:
        stop(
            "Validation failed: summary contains "
            "more than 9 sentences."
        )

    # -----------------------------------------------------------------
    # Arrays
    # -----------------------------------------------------------------

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


# =============================================================================
# COMPLETE ANALYSIS RECORD
# =============================================================================

def build_analysis_record(
    article: dict[str, Any],
    analysis: dict[str, Any],
    metadata: dict[str, Any],
) -> dict[str, Any]:
    """
    Build the persisted analysis record.

    The collector article fields are copied into the analysis
    record so the historical analysis remains self-contained.
    """

    record = dict(analysis)

    record["source"] = article.get(
        "source"
    )

    record["title"] = article.get(
        "title"
    )

    record["publishedAtSgt"] = article.get(
        "publishedAtSgt"
    )

    record["url"] = article.get(
        "url"
    )

    record["analysisGeneratedAtSgt"] = (
        singapore_now().isoformat()
    )

    record["analysisModel"] = MODEL

    record["geminiUsage"] = metadata.get(
        "usage",
        {},
    )

    return record


# =============================================================================
# IMMEDIATE SUCCESS SAVE
# =============================================================================

def save_successful_analysis(
    analysis_document: dict[str, Any],
    record: dict[str, Any],
) -> None:
    """
    Save ONE successful analysis immediately.

    This is the critical data-preservation operation.

    The next Gemini request must not begin until this function
    successfully completes.
    """

    analyses = analysis_document.setdefault(
        "analyses",
        [],
    )

    article_id = record.get(
        "articleId"
    )

    if not isinstance(
        article_id,
        str,
    ) or not article_id:
        stop(
            "Cannot save analysis without articleId."
        )

    existing_ids = {
        item.get("articleId")
        for item in analyses
        if isinstance(item, dict)
    }

    if article_id in existing_ids:
        stop(
            "Duplicate articleId detected immediately before save: "
            f"{article_id}"
        )

    analyses.append(
        record
    )

    analysis_document["generatedAtSgt"] = (
        singapore_now().isoformat()
    )

    write_json_atomic(
        ANALYSIS_CURRENT_FILE,
        analysis_document,
    )

    log(
        f"IMMEDIATE SAVE COMPLETE: {article_id}"
    )


# =============================================================================
# ARCHIVING
# =============================================================================

def archive_record(
    record: dict[str, Any],
    archive_date: date,
) -> None:
    """
    Append one record to its YYYY/MM/YYYY-MM-DD.json history file.

    Existing records are never overwritten.
    """

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
        stop(
            "Article already exists in archive: "
            f"{article_id}"
        )

    analyses.append(
        record
    )

    history["date"] = (
        archive_date.isoformat()
    )

    history["source"] = "Gemini"

    write_json_atomic(
        path,
        history,
    )

    log(
        f"ARCHIVED: {article_id} -> "
        f"{path.relative_to(ROOT)}"
    )


def archive_old_current_records(
    analysis_document: dict[str, Any],
) -> int:
    """
    Move analyses older than the rolling 14-day window into
    permanent history.

    Publication date, NOT analysis date, determines the cutoff.

    Current date counts as day 1.

    Example on 2026-10-04:

        2026-10-04 through 2026-09-21
        remain in current.json.

        2026-09-20 and earlier
        are archived.
    """

    today = singapore_now().date()

    earliest_current_date = (
        today
        - timedelta(
            days=ROLLING_DAYS - 1
        )
    )

    analyses = analysis_document.get(
        "analyses",
        [],
    )

    retained: list[
        dict[str, Any]
    ] = []

    archived_count = 0

    for record in analyses:

        if not isinstance(
            record,
            dict,
        ):
            stop(
                "Invalid analysis record encountered "
                "during rolling-window maintenance."
            )

        article_id = record.get(
            "articleId"
        )

        if not isinstance(
            article_id,
            str,
        ) or not article_id:
            stop(
                "Analysis record missing articleId "
                "during rolling-window maintenance."
            )

        archive_date = publication_date(
            {
                "publishedAtSgt": record.get(
                    "publishedAtSgt"
                )
            }
        )

        # ----------------------------------------------------------
        # Never guess a date.
        # ----------------------------------------------------------

        if archive_date is None:
            stop(
                f"Analysis {article_id} has invalid "
                "publishedAtSgt; cannot determine archive date."
            )

        if archive_date < earliest_current_date:

            archive_record(
                record,
                archive_date,
            )

            archived_count += 1

        else:

            retained.append(
                record
            )

    if archived_count > 0:

        analysis_document["analyses"] = (
            retained
        )

        analysis_document["generatedAtSgt"] = (
            singapore_now().isoformat()
        )

        write_json_atomic(
            ANALYSIS_CURRENT_FILE,
            analysis_document,
        )

        log(
            f"Rolling archive moved "
            f"{archived_count} article(s)."
        )

    return archived_count


# =============================================================================
# DEDUPLICATION
# =============================================================================

def build_processed_ids(
    analysis_document: dict[str, Any],
) -> set[str]:
    """
    Combine current analysis IDs and permanent history IDs.
    """

    current_ids = (
        current_analysis_article_ids(
            analysis_document
        )
    )

    archived_ids = (
        read_history_article_ids()
    )

    combined = (
        current_ids
        | archived_ids
    )

    log(
        f"Current analysis IDs: {len(current_ids)}"
    )

    log(
        f"Archived analysis IDs: {len(archived_ids)}"
    )

    log(
        f"Total processed IDs: {len(combined)}"
    )

    return combined


# =============================================================================
# MAIN
# =============================================================================

def main() -> int:

    started_at = singapore_now()

    log("=" * 70)
    log("VGrat FMS - MARKET NEWS AI ANALYZER")
    log("=" * 70)

    log(
        f"Model: {MODEL}"
    )

    log(
        f"Input: "
        f"{NEWS_FILE.relative_to(ROOT)}"
    )

    log(
        f"Output: "
        f"{ANALYSIS_CURRENT_FILE.relative_to(ROOT)}"
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

    log(
        f"Rolling window: {ROLLING_DAYS} publication days"
    )

    # =========================================================================
    # ENVIRONMENT
    # =========================================================================

    if not API_KEY:
        stop(
            "GEMINI_API_KEY is not configured."
        )

    # =========================================================================
    # LOAD COLLECTOR INPUT
    # =========================================================================

    articles = load_collector_articles()

    # =========================================================================
    # LOAD ANALYSIS STORAGE
    # =========================================================================

    analysis_document = (
        load_current_analysis()
    )

    # =========================================================================
    # ROLLING WINDOW MAINTENANCE
    #
    # Do this before selecting articles to analyze.
    # =========================================================================

    archive_old_current_records(
        analysis_document
    )

    # =========================================================================
    # BUILD DEDUPLICATION INDEX
    # =========================================================================

    processed_ids = build_processed_ids(
        analysis_document
    )

    # =========================================================================
    # FIND NEW ARTICLES
    # =========================================================================

    new_articles: list[
        dict[str, Any]
    ] = []

    skipped = 0

    for article in articles:

        article_id = article["id"]

        if article_id in processed_ids:
            skipped += 1
            continue

        new_articles.append(
            article
        )

    # =========================================================================
    # NEWEST FIRST
    # =========================================================================

    new_articles.sort(
        key=publication_sort_key,
        reverse=True,
    )

    log(
        f"Collector articles: {len(articles)}"
    )

    log(
        f"Already analyzed: {skipped}"
    )

    log(
        f"New articles: {len(new_articles)}"
    )

    # =========================================================================
    # NOTHING TO ANALYZE
    # =========================================================================

    if not new_articles:

        log(
            "No new articles require analysis."
        )

        # Ensure the analysis file exists.
        if not ANALYSIS_CURRENT_FILE.exists():

            analysis_document[
                "generatedAtSgt"
            ] = singapore_now().isoformat()

            write_json_atomic(
                ANALYSIS_CURRENT_FILE,
                analysis_document,
            )

        elapsed = (
            singapore_now()
            - started_at
        ).total_seconds()

        log(
            f"Completed in {elapsed:.1f} seconds."
        )

        return 0

    # =========================================================================
    # SEQUENTIAL ANALYSIS
    #
    # CRITICAL:
    #
    # There is NO try/except here that continues to the next article.
    #
    # If anything fails:
    #
    #     call_gemini()
    #          ↓
    #     validate()
    #          ↓
    #     save()
    #
    # raises AnalysisError and main exits.
    # =========================================================================

    successful = 0

    for position, article in enumerate(
        new_articles,
        start=1,
    ):

        article_id = article["id"]

        title = article.get(
            "title",
            "",
        )

        published = article.get(
            "publishedAtSgt",
            "",
        )

        log("")
        log("=" * 70)
        log(
            f"ARTICLE {position}/{len(new_articles)}"
        )
        log(
            f"ID: {article_id}"
        )
        log(
            f"Published SGT: {published}"
        )
        log(
            f"Title: {title}"
        )
        log("=" * 70)

        # ---------------------------------------------------------------------
        # EXACTLY ONE GEMINI REQUEST
        # ---------------------------------------------------------------------

        analysis, metadata = call_gemini(
            article
        )

        # ---------------------------------------------------------------------
        # STRICT VALIDATION
        # ---------------------------------------------------------------------

        validated = validate_analysis(
            analysis,
            article,
        )

        log(
            f"Validation SUCCESS: {article_id}"
        )

        # ---------------------------------------------------------------------
        # BUILD PERSISTED RECORD
        # ---------------------------------------------------------------------

        record = build_analysis_record(
            article,
            validated,
            metadata,
        )

        # ---------------------------------------------------------------------
        # IMMEDIATE SAVE
        #
        # The next Gemini request cannot start until this succeeds.
        # ---------------------------------------------------------------------

        save_successful_analysis(
            analysis_document,
            record,
        )

        processed_ids.add(
            article_id
        )

        successful += 1

        log(
            f"SUCCESS: {successful}/"
            f"{len(new_articles)} articles analyzed and saved."
        )

    # =========================================================================
    # FINAL ROLLING-WINDOW MAINTENANCE
    # =========================================================================

    archive_old_current_records(
        analysis_document
    )

    # =========================================================================
    # COMPLETE
    # =========================================================================

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
        f"Skipped: {skipped}"
    )
    log(
        f"Elapsed: {elapsed:.1f} seconds"
    )
    log("=" * 70)

    return 0


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":

    try:

        exit_code = main()

    except AnalysisError as exc:

        log("")
        log("=" * 70)
        log("ANALYZER STOPPED")
        log("=" * 70)
        log(
            str(exc)
        )
        log("=" * 70)
        log(
            "No further Gemini requests will be made."
        )
        log(
            "Previously successful analyses remain saved on disk."
        )
        log("=" * 70)

        # Non-zero exit is intentional.
        #
        # GitHub Actions marks this analysis step as failed, but
        # the workflow's later `if: always()` commit step still
        # commits successful files already written.
        sys.exit(1)

    except KeyboardInterrupt:

        log("")
        log(
            "Analyzer interrupted."
        )
        log(
            "No further Gemini requests will be made."
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
        log(
            "Previously successful analyses remain saved on disk."
        )
        log("=" * 70)

        # Unexpected Python errors are fatal.
        sys.exit(1)
