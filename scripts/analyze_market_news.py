#!/usr/bin/env python3

"""
VGrat FMS - MARKET NEWS AI ANALYZER
===================================

FIRST REAL-DATA BATCH TEST

This version intentionally processes ONE batch of exactly 5 articles.

INPUTS - READ ONLY
------------------
    data/market_news/current.json
    Research Funds.xlsx

OUTPUT
------
    data/market_news/analysis/current.json

BEHAVIOUR
---------
1. Load the real production current.json.
2. Load Research Funds.xlsx as read-only research context.
3. Identify articles already successfully stored in analysis/current.json.
4. Select the newest 5 pending articles.
5. Send all 5 articles in ONE Gemini request.
6. Validate that Gemini returned exactly one analysis per article.
7. Only after complete validation, write the successful batch.
8. Preserve all previously successful batches.
9. If the Gemini request fails or validation fails, nothing is written.
10. Production current.json is never modified.

Each successful batch stores:
    - batch number
    - article IDs
    - model
    - prompt
    - request timestamp
    - response timestamp
    - response duration
    - input article count
    - output article count
    - validation status
    - Gemini response
    - individual article analyses
"""

from __future__ import annotations

import hashlib
import json
import os
import sys
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any, Dict, List, Tuple

import requests
from openpyxl import load_workbook


# ============================================================
# PATHS
# ============================================================

ROOT = Path(__file__).resolve().parent.parent

CURRENT_JSON = (
    ROOT
    / "data"
    / "market_news"
    / "current.json"
)

RESEARCH_FUNDS = ROOT / "Research Funds.xlsx"

ANALYSIS_DIR = (
    ROOT
    / "data"
    / "market_news"
    / "analysis"
)

ANALYSIS_CURRENT = (
    ANALYSIS_DIR
    / "current.json"
)


# ============================================================
# TEST CONFIGURATION
# ============================================================

# FIRST TEST IS FIXED TO ONE BATCH OF FIVE.
BATCH_SIZE = 5
MAX_BATCHES = 1

PROCESS_NEWEST_FIRST = True


# ============================================================
# GEMINI CONFIGURATION
# ============================================================

GEMINI_API_KEY = os.getenv(
    "GEMINI_API_KEY",
    "",
).strip()

GEMINI_MODEL = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.8-flash",
).strip()

GEMINI_URL = (
    "https://generativelanguage.googleapis.com/"
    "v1beta/models/"
    f"{GEMINI_MODEL}:generateContent"
)

GEMINI_TIMEOUT = int(
    os.getenv(
        "GEMINI_TIMEOUT_SECONDS",
        "180",
    )
)

MAX_OUTPUT_TOKENS = int(
    os.getenv(
        "GEMINI_MAX_OUTPUT_TOKENS",
        "16384",
    )
)


# ============================================================
# LOGGING
# ============================================================

PROGRAM_START = time.perf_counter()


def log(message: str) -> None:
    elapsed = time.perf_counter() - PROGRAM_START

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    print(
        f"[{now}] "
        f"[+{elapsed:.2f}s] "
        f"{message}",
        flush=True,
    )


# ============================================================
# TIME
# ============================================================

def utc_now() -> str:
    return datetime.now(
        timezone.utc
    ).isoformat()


def singapore_now() -> datetime:
    return datetime.now(
        timezone.utc
    ).astimezone(
        timezone(
            timedelta(hours=8)
        )
    )


# ============================================================
# JSON
# ============================================================

def load_json(path: Path) -> Any:
    with path.open(
        "r",
        encoding="utf-8",
    ) as f:
        return json.load(f)


def write_json_atomic(
    path: Path,
    data: Any,
) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temp_path = path.with_suffix(
        path.suffix + ".tmp"
    )

    with temp_path.open(
        "w",
        encoding="utf-8",
    ) as f:
        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2,
        )
        f.write("\n")

    temp_path.replace(path)


# ============================================================
# GENERAL HELPERS
# ============================================================

def text(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, str):
        return value.strip()

    return str(value).strip()


def first_value(
    obj: Dict[str, Any],
    keys: List[str],
) -> Any:
    for key in keys:
        if key in obj:
            value = obj[key]

            if value is None:
                continue

            if isinstance(value, str):
                if value.strip():
                    return value.strip()
            else:
                return value

    return ""


def sha256_text(value: str) -> str:
    return hashlib.sha256(
        value.encode("utf-8")
    ).hexdigest()


# ============================================================
# ARTICLE EXTRACTION
# ============================================================

def extract_articles(
    payload: Any,
) -> List[Dict[str, Any]]:
    if isinstance(payload, list):
        return [
            x for x in payload
            if isinstance(x, dict)
        ]

    if not isinstance(payload, dict):
        raise ValueError(
            "current.json must contain an object or array."
        )

    preferred_keys = [
        "articles",
        "news",
        "items",
        "data",
        "results",
    ]

    for key in preferred_keys:
        value = payload.get(key)

        if isinstance(value, list):
            return [
                x for x in value
                if isinstance(x, dict)
            ]

    raise ValueError(
        "Could not find an article array in current.json."
    )


def article_id(
    article: Dict[str, Any],
) -> str:
    value = first_value(
        article,
        [
            "id",
            "articleId",
            "article_id",
            "newsId",
            "news_id",
        ],
    )

    if value:
        return text(value)

    url = first_value(
        article,
        [
            "url",
            "link",
            "articleUrl",
            "article_url",
        ],
    )

    if url:
        return sha256_text(
            text(url)
        )[:24]

    # Deterministic fallback.
    canonical = json.dumps(
        article,
        ensure_ascii=False,
        sort_keys=True,
    )

    return sha256_text(
        canonical
    )[:24]


def article_title(
    article: Dict[str, Any],
) -> str:
    return text(
        first_value(
            article,
            [
                "title",
                "headline",
                "name",
            ],
        )
    )


def article_summary(
    article: Dict[str, Any],
) -> str:
    return text(
        first_value(
            article,
            [
                "summary",
                "description",
                "excerpt",
                "shortDescription",
            ],
        )
    )


def article_content(
    article: Dict[str, Any],
) -> str:
    return text(
        first_value(
            article,
            [
                "content",
                "body",
                "text",
                "articleText",
                "article_text",
            ],
        )
    )


def article_url(
    article: Dict[str, Any],
) -> str:
    return text(
        first_value(
            article,
            [
                "url",
                "link",
                "articleUrl",
                "article_url",
            ],
        )
    )


def article_published(
    article: Dict[str, Any],
) -> str:
    return text(
        first_value(
            article,
            [
                "publishedAt",
                "published_at",
                "publishedDate",
                "publishDate",
                "date",
                "timestamp",
            ],
        )
    )


def article_source(
    article: Dict[str, Any],
) -> str:
    return text(
        first_value(
            article,
            [
                "source",
                "publisher",
                "provider",
                "site",
            ],
        )
    )


# ============================================================
# PUBLISHED SORT
# ============================================================

def sort_key(
    article: Dict[str, Any],
) -> str:
    """
    CNBC collector timestamps are expected to be sortable.
    The original published value is preserved.
    """

    return article_published(article)


# ============================================================
# RESEARCH FUNDS
# ============================================================

def load_research_funds() -> Dict[str, Any]:
    """
    Read Research Funds.xlsx only.

    No research is performed here.
    The workbook is treated as already-researched input.
    """

    if not RESEARCH_FUNDS.exists():
        log(
            "Research Funds.xlsx not found."
        )

        return {
            "available": False,
            "rows": [],
            "masterCategories": {},
        }

    log(
        f"Loading {RESEARCH_FUNDS.name}..."
    )

    workbook = load_workbook(
        RESEARCH_FUNDS,
        read_only=True,
        data_only=True,
    )

    result: Dict[str, Any] = {
        "available": True,
        "rows": [],
        "masterCategories": {},
    }

    try:
        # ----------------------------------------------------
        # Fund Research
        # ----------------------------------------------------

        if "Fund Research" in workbook.sheetnames:
            sheet = workbook[
                "Fund Research"
            ]
        else:
            sheet = workbook[
                workbook.sheetnames[0]
            ]

        rows = sheet.iter_rows(
            values_only=True
        )

        try:
            headers = next(rows)
        except StopIteration:
            headers = []

        headers = [
            text(value)
            for value in headers
        ]

        for row in rows:
            record = {}

            for index, value in enumerate(row):
                if index >= len(headers):
                    continue

                header = headers[index]

                if not header:
                    continue

                if value is not None:
                    record[header] = value

            if record:
                result["rows"].append(
                    record
                )

        # ----------------------------------------------------
        # Master categories
        #
        # Existing project convention:
        # worksheet index 1 = Geography
        # worksheet index 2 = Sector
        # ----------------------------------------------------

        if len(workbook.worksheets) > 1:
            result[
                "masterCategories"
            ]["geography"] = (
                sheet_to_values(
                    workbook.worksheets[1]
                )
            )

        if len(workbook.worksheets) > 2:
            result[
                "masterCategories"
            ]["sector"] = (
                sheet_to_values(
                    workbook.worksheets[2]
                )
            )

    finally:
        workbook.close()

    log(
        "Research Funds rows loaded: "
        f"{len(result['rows'])}"
    )

    return result


def sheet_to_values(
    sheet,
) -> List[List[Any]]:
    output = []

    for row in sheet.iter_rows(
        values_only=True
    ):
        values = [
            value
            for value in row
            if value is not None
            and text(value)
        ]

        if values:
            output.append(values)

    return output


# ============================================================
# ANALYSIS CURRENT FILE
# ============================================================

def load_analysis_file() -> Dict[str, Any]:
    """
    Create an empty analysis/current.json structure if needed.

    Existing file is preserved.
    """

    if not ANALYSIS_CURRENT.exists():
        return {
            "version": 1,
            "createdAt": utc_now(),
            "updatedAt": utc_now(),
            "totalSuccessfulBatches": 0,
            "totalSuccessfulArticles": 0,
            "batches": [],
        }

    payload = load_json(
        ANALYSIS_CURRENT
    )

    if not isinstance(payload, dict):
        raise ValueError(
            "analysis/current.json must contain a JSON object."
        )

    payload.setdefault(
        "version",
        1,
    )

    payload.setdefault(
        "createdAt",
        utc_now(),
    )

    payload.setdefault(
        "updatedAt",
        utc_now(),
    )

    payload.setdefault(
        "totalSuccessfulBatches",
        len(
            payload.get(
                "batches",
                [],
            )
        ),
    )

    payload.setdefault(
        "totalSuccessfulArticles",
        0,
    )

    payload.setdefault(
        "batches",
        [],
    )

    return payload


def get_completed_article_ids(
    analysis: Dict[str, Any],
) -> set:
    completed = set()

    for batch in analysis.get(
        "batches",
        [],
    ):
        if not isinstance(batch, dict):
            continue

        # Preferred explicit article IDs.
        for article_id_value in batch.get(
            "articleIds",
            [],
        ):
            if article_id_value:
                completed.add(
                    text(article_id_value)
                )

        # Also inspect individual analyses.
        for item in batch.get(
            "analyses",
            [],
        ):
            if not isinstance(item, dict):
                continue

            value = item.get(
                "articleId"
            )

            if value:
                completed.add(
                    text(value)
                )

    return completed


def next_batch_number(
    analysis: Dict[str, Any],
) -> int:
    batches = analysis.get(
        "batches",
        [],
    )

    numbers = []

    for batch in batches:
        if isinstance(batch, dict):
            value = batch.get(
                "batchNumber"
            )

            if isinstance(value, int):
                numbers.append(value)

    if not numbers:
        return 1

    return max(numbers) + 1


# ============================================================
# GEMINI PROMPT
# ============================================================

def make_article_payload(
    article: Dict[str, Any],
) -> Dict[str, Any]:
    return {
        "articleId": article_id(article),
        "title": article_title(article),
        "publishedAt": article_published(article),
        "source": article_source(article),
        "url": article_url(article),
        "summary": article_summary(article),
        "content": article_content(article),
    }


def build_prompt(
    articles: List[Dict[str, Any]],
    research: Dict[str, Any],
) -> str:
    article_payloads = [
        make_article_payload(article)
        for article in articles
    ]

    # Research Funds is already researched production input.
    # We provide it as context only.
    fund_rows = research.get(
        "rows",
        [],
    )

    prompt = f"""
You are the VGrat FMS Market News Analyst.

Analyze every article supplied below.

IMPORTANT BATCH RULES
=====================
- There are exactly {len(articles)} articles.
- Return exactly {len(articles)} analysis results.
- Each result MUST use the exact articleId supplied.
- Do not omit an article.
- Do not combine articles.
- Do not invent article IDs.
- Do not return duplicate article IDs.
- Do not browse the internet.
- Use only the supplied article information and Research Funds context.
- If information is unavailable, say so rather than inventing it.

PURPOSE
=======
Determine the investment relevance of each CNBC market-news article
for VGrat FMS monitoring.

CLASSIFICATION
==============
Use one of:

MARKET
ECONOMIC
TECHNOLOGY
GEOPOLITICAL
REJECT

For every article provide:

- articleId
- relevant
- category
- sentiment
- importance
- assetClasses
- geographies
- sectors
- investorImpact
- reasoning
- fundMonitoringRelevant

SENTIMENT
=========
Use:
- positive
- negative
- neutral
- mixed

IMPORTANCE
==========
Use:
- high
- medium
- low

The final response MUST be valid JSON matching the requested schema.

RESEARCH FUNDS CONTEXT
======================
The following information comes from Research Funds.xlsx.
It is already researched input. Do not perform additional research.

{json.dumps(
    fund_rows,
    ensure_ascii=False,
)}

ARTICLES
========
{json.dumps(
    article_payloads,
    ensure_ascii=False,
)}
""".strip()

    return prompt


def response_schema() -> Dict[str, Any]:
    return {
        "type": "object",
        "properties": {
            "articles": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "articleId": {
                            "type": "string",
                        },
                        "relevant": {
                            "type": "boolean",
                        },
                        "category": {
                            "type": "string",
                            "enum": [
                                "MARKET",
                                "ECONOMIC",
                                "TECHNOLOGY",
                                "GEOPOLITICAL",
                                "REJECT",
                            ],
                        },
                        "sentiment": {
                            "type": "string",
                            "enum": [
                                "positive",
                                "negative",
                                "neutral",
                                "mixed",
                            ],
                        },
                        "importance": {
                            "type": "string",
                            "enum": [
                                "high",
                                "medium",
                                "low",
                            ],
                        },
                        "assetClasses": {
                            "type": "array",
                            "items": {
                                "type": "string",
                            },
                        },
                        "geographies": {
                            "type": "array",
                            "items": {
                                "type": "string",
                            },
                        },
                        "sectors": {
                            "type": "array",
                            "items": {
                                "type": "string",
                            },
                        },
                        "investorImpact": {
                            "type": "string",
                        },
                        "reasoning": {
                            "type": "string",
                        },
                        "fundMonitoringRelevant": {
                            "type": "boolean",
                        },
                    },
                    "required": [
                        "articleId",
                        "relevant",
                        "category",
                        "sentiment",
                        "importance",
                        "assetClasses",
                        "geographies",
                        "sectors",
                        "investorImpact",
                        "reasoning",
                        "fundMonitoringRelevant",
                    ],
                },
            },
        },
        "required": [
            "articles",
        ],
    }


# ============================================================
# GEMINI API
# ============================================================

def call_gemini(
    prompt: str,
) -> Tuple[
    Dict[str, Any],
    float,
    Dict[str, Any],
]:
    if not GEMINI_API_KEY:
        raise RuntimeError(
            "GEMINI_API_KEY is not configured."
        )

    body = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": prompt,
                    }
                ],
            }
        ],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": response_schema(),
            "maxOutputTokens": MAX_OUTPUT_TOKENS,
        },
    }

    started = time.perf_counter()

    response = requests.post(
        GEMINI_URL,
        headers={
            "Content-Type": "application/json",
            "x-goog-api-key": GEMINI_API_KEY,
        },
        json=body,
        timeout=GEMINI_TIMEOUT,
    )

    duration = (
        time.perf_counter()
        - started
    )

    if response.status_code >= 400:
        raise RuntimeError(
            "Gemini API error "
            f"{response.status_code}: "
            f"{response.text[:4000]}"
        )

    payload = response.json()

    generated_text = extract_response_text(
        payload
    )

    if not generated_text:
        raise RuntimeError(
            "Gemini returned no generated text."
        )

    try:
        parsed = json.loads(
            generated_text
        )
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            "Gemini returned invalid JSON: "
            f"{exc}"
        ) from exc

    return (
        parsed,
        duration,
        payload,
    )


def extract_response_text(
    payload: Dict[str, Any],
) -> str:
    candidates = payload.get(
        "candidates",
        [],
    )

    if not candidates:
        return ""

    candidate = candidates[0]

    content = candidate.get(
        "content",
        {},
    )

    parts = content.get(
        "parts",
        [],
    )

    result = []

    for part in parts:
        if not isinstance(part, dict):
            continue

        value = part.get(
            "text"
        )

        if isinstance(value, str):
            result.append(value)

    return "".join(result).strip()


# ============================================================
# VALIDATION
# ============================================================

def validate_response(
    articles: List[Dict[str, Any]],
    response: Dict[str, Any],
) -> Tuple[
    bool,
    List[Dict[str, Any]],
    List[str],
]:
    errors: List[str] = []

    expected_ids = [
        article_id(article)
        for article in articles
    ]

    expected_set = set(
        expected_ids
    )

    results = response.get(
        "articles"
    )

    if not isinstance(results, list):
        return (
            False,
            [],
            [
                "Response does not contain articles array."
            ],
        )

    if len(results) != len(articles):
        errors.append(
            "Expected "
            f"{len(articles)} results but received "
            f"{len(results)}."
        )

    seen = set()

    for index, result in enumerate(
        results
    ):
        if not isinstance(
            result,
            dict,
        ):
            errors.append(
                f"Result {index + 1} is not an object."
            )
            continue

        value = text(
            result.get(
                "articleId"
            )
        )

        if not value:
            errors.append(
                f"Result {index + 1} has no articleId."
            )
            continue

        if value in seen:
            errors.append(
                f"Duplicate articleId: {value}"
            )

        seen.add(value)

    actual_set = set(seen)

    missing = (
        expected_set
        - actual_set
    )

    unexpected = (
        actual_set
        - expected_set
    )

    if missing:
        errors.append(
            "Missing article IDs: "
            + ", ".join(
                sorted(missing)
            )
        )

    if unexpected:
        errors.append(
            "Unexpected article IDs: "
            + ", ".join(
                sorted(unexpected)
            )
        )

    required = [
        "articleId",
        "relevant",
        "category",
        "sentiment",
        "importance",
        "assetClasses",
        "geographies",
        "sectors",
        "investorImpact",
        "reasoning",
        "fundMonitoringRelevant",
    ]

    for result in results:
        if not isinstance(
            result,
            dict,
        ):
            continue

        result_id = text(
            result.get(
                "articleId"
            )
        )

        for field in required:
            if field not in result:
                errors.append(
                    f"{result_id}: missing {field}"
                )

    return (
        len(errors) == 0,
        results,
        errors,
    )


# ============================================================
# SUCCESSFUL BATCH STORAGE
# ============================================================

def append_successful_batch(
    analysis: Dict[str, Any],
    batch_number: int,
    articles: List[Dict[str, Any]],
    prompt: str,
    response: Dict[str, Any],
    raw_gemini_response: Dict[str, Any],
    response_seconds: float,
    request_started_at: str,
    request_finished_at: str,
    validation_errors: List[str],
) -> Dict[str, Any]:

    article_ids = [
        article_id(article)
        for article in articles
    ]

    batch_record = {
        "batchNumber": batch_number,

        "storedAt": utc_now(),

        "model": GEMINI_MODEL,

        "request": {
            "startedAt": request_started_at,
            "finishedAt": request_finished_at,
            "responseSeconds": round(
                response_seconds,
                4,
            ),
            "articleCount": len(articles),
            "articleIds": article_ids,
        },

        "validation": {
            "success": True,
            "errors": validation_errors,
        },

        # Store the exact prompt used.
        "prompt": prompt,

        # Store parsed response.
        "response": response,

        # Store individual results for easy future access.
        "analyses": response.get(
            "articles",
            [],
        ),

        # Keep raw API response for debugging/performance
        # investigation.
        "rawGeminiResponse": raw_gemini_response,

        # Keep original article metadata alongside the result.
        "inputArticles": [
            make_article_payload(article)
            for article in articles
        ],
    }

    analysis.setdefault(
        "batches",
        [],
    )

    analysis["batches"].append(
        batch_record
    )

    analysis[
        "totalSuccessfulBatches"
    ] = len(
        analysis["batches"]
    )

    analysis[
        "totalSuccessfulArticles"
    ] = sum(
        len(
            batch.get(
                "articleIds",
                [],
            )
        )
        for batch in analysis["batches"]
        if isinstance(batch, dict)
    )

    analysis[
        "updatedAt"
    ] = utc_now()

    return analysis


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    log("=" * 72)
    log("VGrat FMS - MARKET NEWS AI ANALYZER")
    log("=" * 72)

    log(
        "TEST MODE: exactly ONE batch of FIVE articles."
    )

    log(
        f"Gemini model: {GEMINI_MODEL}"
    )

    log(
        f"Input: {CURRENT_JSON}"
    )

    log(
        f"Output: {ANALYSIS_CURRENT}"
    )

    # --------------------------------------------------------
    # Validate inputs
    # --------------------------------------------------------

    if not CURRENT_JSON.exists():
        log(
            "ERROR: current.json does not exist."
        )
        return 1

    if not GEMINI_API_KEY:
        log(
            "ERROR: GEMINI_API_KEY is not configured."
        )
        return 1

    # --------------------------------------------------------
    # Load current.json
    # --------------------------------------------------------

    log(
        "Loading production current.json..."
    )

    try:
        payload = load_json(
            CURRENT_JSON
        )

        all_articles = extract_articles(
            payload
        )

    except Exception as exc:
        log(
            f"ERROR loading current.json: {exc}"
        )
        return 1

    log(
        f"Articles found: {len(all_articles)}"
    )

    # --------------------------------------------------------
    # Load Research Funds
    # --------------------------------------------------------

    try:
        research = load_research_funds()

    except Exception as exc:
        log(
            f"ERROR loading Research Funds.xlsx: {exc}"
        )
        return 1

    # --------------------------------------------------------
    # Load test analysis file
    # --------------------------------------------------------

    try:
        analysis = load_analysis_file()

    except Exception as exc:
        log(
            f"ERROR loading analysis/current.json: {exc}"
        )
        return 1

    completed_ids = (
        get_completed_article_ids(
            analysis
        )
    )

    log(
        "Previously successful articles in test analysis: "
        f"{len(completed_ids)}"
    )

    # --------------------------------------------------------
    # Determine pending articles
    # --------------------------------------------------------

    pending = []

    for article in all_articles:
        current_id = article_id(
            article
        )

        if current_id in completed_ids:
            continue

        pending.append(article)

    log(
        f"Pending articles: {len(pending)}"
    )

    if not pending:
        log(
            "No pending articles remain."
        )
        return 0

    # --------------------------------------------------------
    # Newest first
    # --------------------------------------------------------

    pending.sort(
        key=sort_key,
        reverse=PROCESS_NEWEST_FIRST,
    )

    batch = pending[
        :BATCH_SIZE
    ]

    if len(batch) < BATCH_SIZE:
        log(
            "WARNING: fewer than 5 pending articles "
            f"available. Found {len(batch)}."
        )

    log("")
    log(
        "Selected batch:"
    )

    for index, article in enumerate(
        batch,
        start=1,
    ):
        log(
            f"  {index}. "
            f"{article_id(article)} | "
            f"{article_published(article)} | "
            f"{article_title(article)}"
        )

    # --------------------------------------------------------
    # Batch number
    # --------------------------------------------------------

    batch_number = next_batch_number(
        analysis
    )

    # --------------------------------------------------------
    # Build prompt
    # --------------------------------------------------------

    prompt = build_prompt(
        batch,
        research,
    )

    log("")
    log(
        f"Preparing ONE Gemini request for "
        f"{len(batch)} articles..."
    )

    request_started_at = utc_now()

    # --------------------------------------------------------
    # Gemini
    # --------------------------------------------------------

    try:

        response, response_seconds, raw_response = (
            call_gemini(
                prompt
            )
        )

    except Exception as exc:

        log("")
        log(
            "GEMINI REQUEST FAILED"
        )

        log(
            str(exc)
        )

        log(
            "Nothing was written to analysis/current.json."
        )

        return 1

    request_finished_at = utc_now()

    log("")
    log(
        "Gemini request succeeded."
    )

    log(
        f"Response time: "
        f"{response_seconds:.4f} seconds"
    )

    # --------------------------------------------------------
    # Validate
    # --------------------------------------------------------

    log(
        "Validating batch response..."
    )

    valid, results, validation_errors = (
        validate_response(
            batch,
            response,
        )
    )

    if not valid:

        log(
            "BATCH VALIDATION FAILED."
        )

        for error in validation_errors:
            log(
                f"  - {error}"
            )

        log(
            "Nothing was written to analysis/current.json."
        )

        return 1

    log(
        "Batch validation successful."
    )

    log(
        f"Validated {len(results)} / "
        f"{len(batch)} articles."
    )

    # --------------------------------------------------------
    # ONLY NOW WRITE SUCCESSFUL BATCH
    # --------------------------------------------------------

    try:

        analysis = append_successful_batch(
            analysis=analysis,
            batch_number=batch_number,
            articles=batch,
            prompt=prompt,
            response=response,
            raw_gemini_response=raw_response,
            response_seconds=response_seconds,
            request_started_at=request_started_at,
            request_finished_at=request_finished_at,
            validation_errors=validation_errors,
        )

        # Atomic write.
        write_json_atomic(
            ANALYSIS_CURRENT,
            analysis,
        )

    except Exception as exc:

        log(
            "ERROR writing analysis/current.json:"
        )

        log(
            str(exc)
        )

        return 1

    # --------------------------------------------------------
    # Final statistics
    # --------------------------------------------------------

    total_execution = (
        time.perf_counter()
        - PROGRAM_START
    )

    log("")
    log("=" * 72)
    log("BATCH SUCCESS")
    log("=" * 72)

    log(
        f"Batch number:       {batch_number}"
    )

    log(
        f"Articles processed: {len(batch)}"
    )

    log(
        f"Gemini requests:    1"
    )

    log(
        f"API response time:  {response_seconds:.4f}s"
    )

    log(
        f"Avg/article:        "
        f"{response_seconds / len(batch):.4f}s"
    )

    log(
        f"Execution time:     {total_execution:.4f}s"
    )

    log(
        f"Total test batches: "
        f"{analysis['totalSuccessfulBatches']}"
    )

    log(
        f"Total test articles: "
        f"{analysis['totalSuccessfulArticles']}"
    )

    log(
        f"Saved to: {ANALYSIS_CURRENT}"
    )

    log("=" * 72)

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )
