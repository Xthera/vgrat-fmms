#!/usr/bin/env python3

"""
VGrat FMS - MARKET NEWS AI ANALYZER
===================================

PURPOSE
-------
Analyze pending market-news articles using Gemini.

AUTHORITATIVE RESEARCH INPUT
----------------------------
Research Funds.xlsx is the authoritative source for:

    Geography
    Sector

Gemini must use exact values from those Excel masters.

If Gemini initially returns an invalid Geography or Sector value,
the script enters correction mode and asks Gemini to remap it.

SUMMARY
-------
Every article must contain a "summary" field.

Rules:

    - Factual summary of the article
    - Less than 10 sentences
    - Maximum 9 sentences
    - No minimum sentence count
    - Should not replace investorImpact
    - Should not contain unnecessary investment interpretation

INVESTOR IMPACT
---------------
Investor impact must explicitly consider the Geography and Sector
classifications selected from Research Funds.xlsx.

It should explain potential investment implications across the
identified geographical markets and industry sectors.

It must NOT be a generic statement such as:

    "This could affect investors."

Instead it should connect the news to:

    Geography
    Sector
    Market conditions
    Earnings
    Valuations
    Demand
    Costs
    Regulation
    Competition
    Other relevant investment factors

INPUTS
------
    data/market_news/current.json
    Research Funds.xlsx

OUTPUT
------
    data/market_news/analysis/current.json

PRODUCTION INPUTS ARE READ ONLY
--------------------------------
The script never modifies:

    data/market_news/current.json
    Research Funds.xlsx

ONLY successful analysis is written to:

    data/market_news/analysis/current.json

BATCH
-----
BATCH_SIZE = 1

This is intentionally kept at 2 because previous 5-article requests
experienced Gemini HTTP 503 responses while 2-article requests
succeeded.

PROCESSING
----------
Newest pending articles first.

CORRECTION MODE
---------------
Initial Gemini analysis
        ↓
Python validation
        ↓
Invalid Geography / Sector / Summary > 9 sentences
        ↓
Gemini correction
        ↓
Python validation again
        ↓
Accept only if fully valid

If correction fails, nothing is written.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


# ============================================================
# PATHS
# ============================================================

ROOT = Path(__file__).resolve().parent.parent

CURRENT_JSON = ROOT / "data" / "market_news" / "current.json"
RESEARCH_FUNDS = ROOT / "Research Funds.xlsx"

ANALYSIS_DIR = ROOT / "data" / "market_news" / "analysis"
ANALYSIS_CURRENT = ANALYSIS_DIR / "current.json"


# ============================================================
# CONFIGURATION
# ============================================================

BATCH_SIZE = 1
MAX_BATCHES = 1
PROCESS_NEWEST_FIRST = True

DEFAULT_MODEL = "gemini-3.8-flash"
DEFAULT_TIMEOUT = 180
DEFAULT_MAX_OUTPUT_TOKENS = 16384

MAX_CORRECTION_ATTEMPTS = 2

MAX_SUMMARY_SENTENCES = 9

GEMINI_API_BASE = (
    "https://generativelanguage.googleapis.com/v1beta/models"
)


# ============================================================
# LOGGING
# ============================================================

def log(message: str) -> None:
    timestamp = datetime.now(timezone.utc).strftime(
        "%Y-%m-%d %H:%M:%S"
    )
    print(
        f"[{timestamp} UTC] {message}",
        flush=True,
    )


# ============================================================
# TIME
# ============================================================

def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


# ============================================================
# JSON
# ============================================================

def load_json(path: Path) -> Any:
    with path.open(
        "r",
        encoding="utf-8",
    ) as f:
        return json.load(f)


def save_json_atomic(
    path: Path,
    data: Any,
) -> None:
    """
    Atomically replace a JSON file.
    """

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
# ARTICLE EXTRACTION
# ============================================================

def extract_articles(
    data: Any,
) -> list[dict[str, Any]]:
    """
    Support the common wrappers used by the market-news collector.
    """

    if isinstance(data, list):
        return [
            item
            for item in data
            if isinstance(item, dict)
        ]

    if not isinstance(data, dict):
        raise ValueError(
            "current.json must contain a JSON object or array."
        )

    for key in (
        "articles",
        "news",
        "items",
        "data",
        "results",
    ):
        value = data.get(key)

        if isinstance(value, list):
            return [
                item
                for item in value
                if isinstance(item, dict)
            ]

    raise ValueError(
        "Could not find article list in current.json. "
        "Expected one of: articles, news, items, data, results."
    )


# ============================================================
# ARTICLE ID
# ============================================================

def article_id(
    article: dict[str, Any],
) -> str:
    """
    Prefer the ID already generated by the collector.

    Fallback:
        URL SHA256
        ↓
        canonical article SHA256
    """

    for key in (
        "id",
        "articleId",
        "article_id",
        "newsId",
        "news_id",
    ):
        value = article.get(key)

        if value is not None and str(value).strip():
            return str(value).strip()

    url = article.get("url")

    if url:
        return hashlib.sha256(
            str(url).encode("utf-8")
        ).hexdigest()

    canonical = json.dumps(
        article,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )

    return hashlib.sha256(
        canonical.encode("utf-8")
    ).hexdigest()


# ============================================================
# ARTICLE NORMALIZATION
# ============================================================

def normalize_article(
    article: dict[str, Any],
) -> dict[str, Any]:

    return {
        "articleId": article_id(article),
        "title": article.get("title"),
        "description": article.get("description"),
        "url": article.get("url"),
        "published": article.get("published"),
        "source": article.get("source"),
    }


# ============================================================
# PUBLICATION SORTING
# ============================================================

def publication_sort_key(
    article: dict[str, Any],
) -> tuple[int, str]:

    value = article.get("published")

    if value is None:
        return (
            0,
            "",
        )

    value = str(value).strip()

    if not value:
        return (
            0,
            "",
        )

    return (
        1,
        value,
    )


# ============================================================
# EXCEL MASTER LOADING
# ============================================================

def clean_excel_value(
    value: Any,
) -> str | None:

    if value is None:
        return None

    value = str(value).strip()

    if not value:
        return None

    return value


def load_master_values(
    worksheet: Any,
) -> list[str]:
    """
    Read all non-empty cells from a master worksheet.

    Duplicate values are removed while preserving order.
    """

    values: list[str] = []
    seen: set[str] = set()

    for row in worksheet.iter_rows():

        for cell in row:

            value = clean_excel_value(
                cell.value
            )

            if value is None:
                continue

            if value in seen:
                continue

            seen.add(value)
            values.append(value)

    return values


def load_research_masters() -> tuple[
    list[str],
    list[str],
]:
    """
    Research Funds.xlsx structure:

        Worksheet 0:
            Fund Research

        Worksheet 1:
            Geography master

        Worksheet 2:
            Sector master
    """

    if not RESEARCH_FUNDS.exists():
        raise FileNotFoundError(
            f"Research Funds.xlsx not found: "
            f"{RESEARCH_FUNDS}"
        )

    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError(
            "openpyxl is required."
        ) from exc

    log(
        f"Loading research masters from: "
        f"{RESEARCH_FUNDS}"
    )

    workbook = load_workbook(
        RESEARCH_FUNDS,
        read_only=True,
        data_only=True,
    )

    try:

        worksheets = workbook.worksheets

        if len(worksheets) < 3:
            raise ValueError(
                "Research Funds.xlsx must contain at least "
                "3 worksheets."
            )

        geography_master = load_master_values(
            worksheets[1]
        )

        sector_master = load_master_values(
            worksheets[2]
        )

    finally:
        workbook.close()

    if not geography_master:
        raise ValueError(
            "Geography master is empty."
        )

    if not sector_master:
        raise ValueError(
            "Sector master is empty."
        )

    log(
        f"Geography master: "
        f"{len(geography_master)} values"
    )

    log(
        f"Sector master: "
        f"{len(sector_master)} values"
    )

    return (
        geography_master,
        sector_master,
    )


# ============================================================
# ANALYSIS FILE
# ============================================================

def new_analysis_document() -> dict[str, Any]:

    now = utc_now_iso()

    return {
        "version": 1,
        "createdAt": now,
        "updatedAt": now,
        "totalSuccessfulBatches": 0,
        "totalSuccessfulArticles": 0,
        "batches": [],
    }


def load_existing_analysis() -> dict[str, Any]:

    if not ANALYSIS_CURRENT.exists():

        log(
            "analysis/current.json does not exist. "
            "Creating a new analysis file."
        )

        return new_analysis_document()

    data = load_json(
        ANALYSIS_CURRENT
    )

    if not isinstance(data, dict):
        raise ValueError(
            "analysis/current.json must contain an object."
        )

    data.setdefault(
        "version",
        1,
    )

    data.setdefault(
        "createdAt",
        utc_now_iso(),
    )

    data.setdefault(
        "updatedAt",
        utc_now_iso(),
    )

    data.setdefault(
        "totalSuccessfulBatches",
        0,
    )

    data.setdefault(
        "totalSuccessfulArticles",
        0,
    )

    if not isinstance(
        data.get("batches"),
        list,
    ):
        data["batches"] = []

    return data


# ============================================================
# COMPLETED IDS
# ============================================================

def completed_article_ids(
    analysis: dict[str, Any],
) -> set[str]:

    completed: set[str] = set()

    for batch in analysis.get(
        "batches",
        [],
    ):

        if not isinstance(
            batch,
            dict,
        ):
            continue

        for value in batch.get(
            "articleIds",
            [],
        ):

            if value is not None:
                completed.add(
                    str(value)
                )

        for item in batch.get(
            "analyses",
            [],
        ):

            if not isinstance(
                item,
                dict,
            ):
                continue

            value = item.get(
                "articleId"
            )

            if value is not None:
                completed.add(
                    str(value)
                )

    return completed


# ============================================================
# GEMINI CONFIG
# ============================================================

def gemini_model() -> str:

    return os.environ.get(
        "GEMINI_MODEL",
        DEFAULT_MODEL,
    )


def gemini_timeout() -> int:

    raw = os.environ.get(
        "GEMINI_TIMEOUT_SECONDS",
        str(DEFAULT_TIMEOUT),
    )

    try:
        return int(raw)
    except ValueError:
        return DEFAULT_TIMEOUT


def gemini_max_output_tokens() -> int:

    raw = os.environ.get(
        "GEMINI_MAX_OUTPUT_TOKENS",
        str(DEFAULT_MAX_OUTPUT_TOKENS),
    )

    try:
        return int(raw)
    except ValueError:
        return DEFAULT_MAX_OUTPUT_TOKENS


# ============================================================
# GEMINI API
# ============================================================

def call_gemini(
    *,
    model: str,
    api_key: str,
    prompt: str,
) -> dict[str, Any]:

    url = (
        f"{GEMINI_API_BASE}/"
        f"{model}:generateContent"
        f"?key={api_key}"
    )

    request_body = {
        "contents": [
            {
                "parts": [
                    {
                        "text": prompt,
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0,
            "maxOutputTokens": gemini_max_output_tokens(),
            "responseMimeType": "application/json",
        },
    }

    encoded = json.dumps(
        request_body,
        ensure_ascii=False,
    ).encode("utf-8")

    request = urllib.request.Request(
        url,
        data=encoded,
        headers={
            "Content-Type": "application/json",
        },
        method="POST",
    )

    started = time.monotonic()

    try:

        with urllib.request.urlopen(
            request,
            timeout=gemini_timeout(),
        ) as response:

            response_seconds = (
                time.monotonic()
                - started
            )

            raw_bytes = response.read()

            raw_text = raw_bytes.decode(
                "utf-8",
                errors="replace",
            )

            parsed = json.loads(
                raw_text
            )

            return {
                "ok": True,
                "httpStatus": response.status,
                "responseSeconds": round(
                    response_seconds,
                    3,
                ),
                "raw": parsed,
                "requestBody": request_body,
            }

    except urllib.error.HTTPError as exc:

        response_seconds = (
            time.monotonic()
            - started
        )

        try:
            error_body = (
                exc.read()
                .decode(
                    "utf-8",
                    errors="replace",
                )
            )
        except Exception:
            error_body = ""

        try:
            error_json = json.loads(
                error_body
            )
        except Exception:
            error_json = error_body

        return {
            "ok": False,
            "httpStatus": exc.code,
            "responseSeconds": round(
                response_seconds,
                3,
            ),
            "error": error_json,
            "requestBody": request_body,
        }

    except Exception as exc:

        response_seconds = (
            time.monotonic()
            - started
        )

        return {
            "ok": False,
            "httpStatus": None,
            "responseSeconds": round(
                response_seconds,
                3,
            ),
            "error": str(exc),
            "requestBody": request_body,
        }


# ============================================================
# GEMINI RESPONSE PARSING
# ============================================================

def extract_response_text(
    response: dict[str, Any],
) -> str:

    raw = response.get(
        "raw"
    )

    if not isinstance(
        raw,
        dict,
    ):
        raise ValueError(
            "Gemini response is not an object."
        )

    candidates = raw.get(
        "candidates"
    )

    if not isinstance(
        candidates,
        list,
    ) or not candidates:
        raise ValueError(
            "Gemini returned no candidates."
        )

    candidate = candidates[0]

    if not isinstance(
        candidate,
        dict,
    ):
        raise ValueError(
            "Gemini candidate is invalid."
        )

    content = candidate.get(
        "content"
    )

    if not isinstance(
        content,
        dict,
    ):
        raise ValueError(
            "Gemini candidate has no content."
        )

    parts = content.get(
        "parts"
    )

    if not isinstance(
        parts,
        list,
    ):
        raise ValueError(
            "Gemini content has no parts."
        )

    texts: list[str] = []

    for part in parts:

        if not isinstance(
            part,
            dict,
        ):
            continue

        value = part.get(
            "text"
        )

        if isinstance(
            value,
            str,
        ):
            texts.append(value)

    if not texts:
        raise ValueError(
            "Gemini returned no text."
        )

    return "".join(
        texts
    ).strip()


def parse_generated_json(
    text_value: str,
) -> Any:

    cleaned = text_value.strip()

    if cleaned.startswith("```"):

        lines = cleaned.splitlines()

        if (
            lines
            and lines[0].startswith("```")
        ):
            lines = lines[1:]

        if (
            lines
            and lines[-1].strip() == "```"
        ):
            lines = lines[:-1]

        cleaned = "\n".join(
            lines
        ).strip()

        if cleaned.lower().startswith(
            "json"
        ):
            cleaned = cleaned[4:].strip()

    return json.loads(
        cleaned
    )


def normalize_analysis_list(
    value: Any,
) -> list[dict[str, Any]]:

    if isinstance(
        value,
        list,
    ):
        return [
            item
            for item in value
            if isinstance(
                item,
                dict,
            )
        ]

    if isinstance(
        value,
        dict,
    ):

        for key in (
            "analyses",
            "results",
        ):

            nested = value.get(
                key
            )

            if isinstance(
                nested,
                list,
            ):
                return [
                    item
                    for item in nested
                    if isinstance(
                        item,
                        dict,
                    )
                ]

        if value.get(
            "articleId"
        ) is not None:
            return [value]

    raise ValueError(
        "Generated JSON does not contain "
        "a valid analysis list."
    )


# ============================================================
# SENTENCE COUNTING
# ============================================================

def count_sentences(
    text_value: Any,
) -> int:

    if not isinstance(
        text_value,
        str,
    ):
        return 0

    text_value = text_value.strip()

    if not text_value:
        return 0

    """
    This deliberately uses a practical sentence-boundary heuristic
    rather than requiring a specific punctuation style.

    It recognizes:
        .
        !
        ?

    It also handles common abbreviations reasonably by requiring
    whitespace or end-of-string after punctuation.
    """

    matches = re.findall(
        r"[.!?](?=\s|$)",
        text_value,
    )

    return len(matches)


# ============================================================
# VALIDATION
# ============================================================

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


def validate_analysis(
    analyses: list[dict[str, Any]],
    articles: list[dict[str, Any]],
    geography_master: set[str],
    sector_master: set[str],
) -> dict[str, Any]:

    expected_ids = [
        str(article["articleId"])
        for article in articles
    ]

    expected_set = set(
        expected_ids
    )

    received_ids = [
        str(
            item.get(
                "articleId"
            )
        )
        for item in analyses
    ]

    received_set = set(
        received_ids
    )

    errors: list[str] = []

    invalid_geographies: dict[
        str,
        list[str],
    ] = {}

    invalid_sectors: dict[
        str,
        list[str],
    ] = {}

    summary_sentence_counts: dict[
        str,
        int,
    ] = {}

    # --------------------------------------------------------
    # Count
    # --------------------------------------------------------

    if len(analyses) != len(articles):

        errors.append(
            f"Expected {len(articles)} analyses, "
            f"received {len(analyses)}."
        )

    # --------------------------------------------------------
    # Duplicate IDs
    # --------------------------------------------------------

    if len(received_ids) != len(
        received_set
    ):

        errors.append(
            "Duplicate article IDs detected."
        )

    # --------------------------------------------------------
    # Exact article coverage
    # --------------------------------------------------------

    missing_ids = sorted(
        expected_set - received_set
    )

    unexpected_ids = sorted(
        received_set - expected_set
    )

    if missing_ids:

        errors.append(
            "Missing article IDs: "
            + ", ".join(
                missing_ids
            )
        )

    if unexpected_ids:

        errors.append(
            "Unexpected article IDs: "
            + ", ".join(
                unexpected_ids
            )
        )

    # --------------------------------------------------------
    # Per-analysis validation
    # --------------------------------------------------------

    for item in analyses:

        current_id = str(
            item.get(
                "articleId"
            )
        )

        # ----------------------------------------------------
        # Required fields
        # ----------------------------------------------------

        missing_fields = sorted(
            REQUIRED_ANALYSIS_FIELDS
            - set(item.keys())
        )

        if missing_fields:

            errors.append(
                f"{current_id}: missing fields: "
                + ", ".join(
                    missing_fields
                )
            )

        # ----------------------------------------------------
        # Boolean fields
        # ----------------------------------------------------

        if not isinstance(
            item.get("relevant"),
            bool,
        ):

            errors.append(
                f"{current_id}: relevant must be boolean."
            )

        if not isinstance(
            item.get(
                "fundMonitoringRelevant"
            ),
            bool,
        ):

            errors.append(
                f"{current_id}: "
                "fundMonitoringRelevant must be boolean."
            )

        # ----------------------------------------------------
        # Summary
        # ----------------------------------------------------

        summary = item.get(
            "summary"
        )

        if not isinstance(
            summary,
            str,
        ):

            errors.append(
                f"{current_id}: summary must be a string."
            )

        else:

            summary = summary.strip()

            if not summary:

                errors.append(
                    f"{current_id}: summary cannot be empty."
                )

            sentence_count = count_sentences(
                summary
            )

            summary_sentence_counts[
                current_id
            ] = sentence_count

            if sentence_count > MAX_SUMMARY_SENTENCES:

                errors.append(
                    f"{current_id}: summary contains "
                    f"{sentence_count} sentences; "
                    f"maximum is "
                    f"{MAX_SUMMARY_SENTENCES}."
                )

        # ----------------------------------------------------
        # Arrays
        # ----------------------------------------------------

        for field in (
            "assetClasses",
            "geographies",
            "sectors",
        ):

            value = item.get(
                field
            )

            if not isinstance(
                value,
                list,
            ):

                errors.append(
                    f"{current_id}: {field} "
                    "must be an array."
                )

                continue

            if not all(
                isinstance(
                    x,
                    str,
                )
                for x in value
            ):

                errors.append(
                    f"{current_id}: {field} "
                    "must contain only strings."
                )

        # ----------------------------------------------------
        # Geography master
        # ----------------------------------------------------

        geographies = item.get(
            "geographies"
        )

        if isinstance(
            geographies,
            list,
        ):

            invalid = [
                str(value)
                for value in geographies
                if str(value)
                not in geography_master
            ]

            if invalid:

                invalid_geographies[
                    current_id
                ] = invalid

                errors.append(
                    f"{current_id}: invalid Geography "
                    "values: "
                    + ", ".join(
                        invalid
                    )
                )

        # ----------------------------------------------------
        # Sector master
        # ----------------------------------------------------

        sectors = item.get(
            "sectors"
        )

        if isinstance(
            sectors,
            list,
        ):

            invalid = [
                str(value)
                for value in sectors
                if str(value)
                not in sector_master
            ]

            if invalid:

                invalid_sectors[
                    current_id
                ] = invalid

                errors.append(
                    f"{current_id}: invalid Sector "
                    "values: "
                    + ", ".join(
                        invalid
                    )
                )

    return {
        "passed": not errors,
        "errors": errors,
        "expectedArticleCount": len(
            articles
        ),
        "receivedArticleCount": len(
            analyses
        ),
        "expectedArticleIds": expected_ids,
        "receivedArticleIds": received_ids,
        "invalidGeographies": invalid_geographies,
        "invalidSectors": invalid_sectors,
        "summarySentenceCounts": (
            summary_sentence_counts
        ),
    }


# ============================================================
# INITIAL ANALYSIS PROMPT
# ============================================================

def numbered_values(
    values: list[str],
) -> str:

    return "\n".join(
        f"{index}. {value}"
        for index, value in enumerate(
            values,
            start=1,
        )
    )


def article_prompt_json(
    articles: list[dict[str, Any]],
) -> str:

    return json.dumps(
        articles,
        ensure_ascii=False,
        indent=2,
    )


def build_analysis_prompt(
    *,
    articles: list[dict[str, Any]],
    geography_master: list[str],
    sector_master: list[str],
) -> str:

    geography_text = numbered_values(
        geography_master
    )

    sector_text = numbered_values(
        sector_master
    )

    articles_text = article_prompt_json(
        articles
    )

    return f"""
You are the market-news intelligence analyst for VGrat FMS.

Analyze EVERY article supplied below.

============================================================
ARTICLE COVERAGE
============================================================

Return exactly ONE analysis object for EVERY article.

Do not omit articles.

Do not merge articles.

Do not duplicate articles.

Do not create additional article IDs.

The articleId supplied in the input is authoritative.

Return the exact same articleId.

============================================================
SUMMARY
============================================================

Every article MUST have a "summary".

The summary must:

- Explain what happened in the article.
- Be factual and concise.
- Be based only on the supplied article.
- Contain LESS THAN 10 sentences.
- Therefore contain a maximum of 9 sentences.
- Have no minimum sentence count.
- Avoid unnecessary investment interpretation.
- Do not use the summary as a replacement for investorImpact.

============================================================
GEOGRAPHY MASTER
============================================================

The following values come directly from
Research Funds.xlsx.

These are the ONLY valid Geography values.

Every Geography value MUST be copied EXACTLY from this list.

Do not:

- invent values
- abbreviate values
- create synonyms
- create alternative names
- change capitalization
- change spelling

For example:

If "United States" is in the master, do not return "US"
unless "US" is also explicitly present in the master.

VALID GEOGRAPHY VALUES:

{geography_text}

============================================================
SECTOR MASTER
============================================================

The following values come directly from
Research Funds.xlsx.

These are the ONLY valid Sector values.

Every Sector value MUST be copied EXACTLY from this list.

Do not:

- invent values
- abbreviate values
- create synonyms
- create alternative names
- change capitalization
- change spelling

For example:

If "Information Technology" is in the master,
do not return "Technology" unless "Technology" is also
explicitly present in the master.

VALID SECTOR VALUES:

{sector_text}

============================================================
ASSET CLASSES
============================================================

Asset Classes are not restricted by the Geography or Sector
masters.

Use appropriate investment asset-class descriptions.

============================================================
INVESTOR IMPACT
============================================================

This is an important requirement.

The "investorImpact" field MUST specifically consider the
Geography and Sector classifications selected from the
Research Funds.xlsx masters.

Do NOT provide a generic investment statement.

You must connect the article to the relevant:

- geographical markets
- industry sectors
- market conditions
- demand
- earnings
- valuations
- costs
- regulation
- competition
- capital expenditure
- supply chains
- other relevant investment factors

Where appropriate, explain how the news could affect investors
with exposure to the identified Geography + Sector combinations.

For example, if:

geographies:
["United States", "North America"]

sectors:
["Information Technology"]

then investorImpact should discuss implications for the US /
North American Information Technology sector.

If multiple geographies and sectors are relevant, consider the
relationships between them.

Do not simply repeat the geography and sector names.

Do not assume that every identified geography or sector will
experience the same impact.

If the effect is uncertain, explain the uncertainty.

============================================================
FUND MONITORING
============================================================

"fundMonitoringRelevant" should indicate whether the article
could reasonably matter when monitoring funds with exposure to
the identified asset classes, geographies or sectors.

============================================================
OUTPUT SCHEMA
============================================================

Return JSON only.

Return a JSON array.

Each object MUST contain:

{{
  "articleId": "string",
  "relevant": true,
  "category": "MARKET",
  "sentiment": "positive",
  "importance": "medium",
  "summary": "Concise factual article summary.",
  "assetClasses": ["Equities"],
  "geographies": ["United States"],
  "sectors": ["Information Technology"],
  "investorImpact": "Geography- and sector-specific investment impact.",
  "reasoning": "Why the article was classified this way.",
  "fundMonitoringRelevant": true
}}

============================================================
FIELD RULES
============================================================

category must be one of:

MARKET
ECONOMIC
TECHNOLOGY
GEOPOLITICAL
REJECT

sentiment must be one of:

positive
negative
mixed
neutral

importance must be one of:

low
medium
high

assetClasses:
    JSON array of strings.

geographies:
    JSON array of strings.
    EVERY value must exactly match the Geography Master.

sectors:
    JSON array of strings.
    EVERY value must exactly match the Sector Master.

summary:
    Required.
    Maximum 9 sentences.

investorImpact:
    Required.
    Must explicitly consider the selected Geography and Sector.

reasoning:
    Required.

fundMonitoringRelevant:
    Required boolean.

============================================================
FINAL SELF-CHECK
============================================================

Before returning the JSON:

1. Confirm there is exactly one result per input article.
2. Confirm every articleId exactly matches the input.
3. Confirm every Geography exactly matches the Geography Master.
4. Confirm every Sector exactly matches the Sector Master.
5. Confirm every summary has fewer than 10 sentences.
6. Confirm investorImpact specifically discusses the relevant
   Geography and Sector implications.
7. Return JSON only.

============================================================
ARTICLES
============================================================

{articles_text}
""".strip()


# ============================================================
# CORRECTION PROMPT
# ============================================================

def build_correction_prompt(
    *,
    articles: list[dict[str, Any]],
    analyses: list[dict[str, Any]],
    validation: dict[str, Any],
    geography_master: list[str],
    sector_master: list[str],
) -> str:

    geography_text = numbered_values(
        geography_master
    )

    sector_text = numbered_values(
        sector_master
    )

    problems = {
        "invalidGeographies": validation.get(
            "invalidGeographies",
            {},
        ),
        "invalidSectors": validation.get(
            "invalidSectors",
            {},
        ),
        "summarySentenceCounts": validation.get(
            "summarySentenceCounts",
            {},
        ),
        "validationErrors": validation.get(
            "errors",
            [],
        ),
    }

    return f"""
You are correcting a market-news analysis produced by another
AI analyst for VGrat FMS.

The original analysis contains one or more validation problems.

Your job is to produce a fully corrected analysis.

============================================================
AUTHORITATIVE SOURCE
============================================================

Research Funds.xlsx is the authoritative source for Geography
and Sector.

A Geography is valid ONLY when it exactly matches a value in the
Geography Master.

A Sector is valid ONLY when it exactly matches a value in the
Sector Master.

Do not invent values.

Do not abbreviate values.

Do not use synonyms.

Do not use alternative spellings.

============================================================
GEOGRAPHY MASTER
============================================================

{geography_text}

============================================================
SECTOR MASTER
============================================================

{sector_text}

============================================================
VALIDATION PROBLEMS
============================================================

{json.dumps(
    problems,
    ensure_ascii=False,
    indent=2,
)}

============================================================
SUMMARY CORRECTION
============================================================

Every summary must contain fewer than 10 sentences.

Maximum:

9 sentences.

If a summary contains more than 9 sentences, rewrite it into
9 or fewer sentences while preserving the important factual
information.

There is no minimum sentence count.

============================================================
INVESTOR IMPACT
============================================================

The investorImpact field must explicitly consider the Geography
and Sector values selected from the Excel masters.

It must explain the potential investment implications for the
identified geographical markets and industry sectors.

Do not replace this with a generic statement.

If you correct a Geography or Sector, make sure investorImpact
remains consistent with the corrected classification.

============================================================
PRESERVATION RULE
============================================================

Preserve the original analysis wherever it is already valid.

Do not unnecessarily rewrite valid fields.

Do not change the articleId.

Do not change the article coverage.

Do not create additional articles.

Do not omit articles.

============================================================
ARTICLES
============================================================

{json.dumps(
    articles,
    ensure_ascii=False,
    indent=2,
)}

============================================================
ORIGINAL ANALYSES
============================================================

{json.dumps(
    analyses,
    ensure_ascii=False,
    indent=2,
)}

============================================================
OUTPUT
============================================================

Return JSON only.

Return exactly one analysis object per article.

Each object MUST contain:

{{
  "articleId": "string",
  "relevant": true,
  "category": "MARKET",
  "sentiment": "positive",
  "importance": "medium",
  "summary": "Factual summary of fewer than 10 sentences.",
  "assetClasses": ["Equities"],
  "geographies": ["United States"],
  "sectors": ["Information Technology"],
  "investorImpact": "Specific impact across the selected Geography and Sector.",
  "reasoning": "Classification reasoning.",
  "fundMonitoringRelevant": true
}}

FINAL CHECK:

- Exactly one result per article.
- Exact articleId.
- Every Geography exists exactly in the master.
- Every Sector exists exactly in the master.
- Every summary contains no more than 9 sentences.
- InvestorImpact is geographically and sector-specific.
- JSON only.
""".strip()


# ============================================================
# BATCH ANALYSIS
# ============================================================

def analyze_batch(
    *,
    articles: list[dict[str, Any]],
    geography_master: list[str],
    sector_master: list[str],
    api_key: str,
    model: str,
) -> dict[str, Any]:

    # --------------------------------------------------------
    # Initial analysis
    # --------------------------------------------------------

    prompt = build_analysis_prompt(
        articles=articles,
        geography_master=geography_master,
        sector_master=sector_master,
    )

    log(
        f"Initial prompt characters: "
        f"{len(prompt)}"
    )

    initial_response = call_gemini(
        model=model,
        api_key=api_key,
        prompt=prompt,
    )

    if not initial_response["ok"]:

        return {
            "success": False,
            "stage": "initial_api",
            "prompt": prompt,
            "response": initial_response,
            "error": (
                "Initial Gemini request failed."
            ),
        }

    try:

        raw_text = extract_response_text(
            initial_response
        )

        parsed = parse_generated_json(
            raw_text
        )

        analyses = normalize_analysis_list(
            parsed
        )

    except Exception as exc:

        return {
            "success": False,
            "stage": "initial_parse",
            "prompt": prompt,
            "response": initial_response,
            "error": str(exc),
        }

    validation = validate_analysis(
        analyses=analyses,
        articles=articles,
        geography_master=set(
            geography_master
        ),
        sector_master=set(
            sector_master
        ),
    )

    # --------------------------------------------------------
    # Initial success
    # --------------------------------------------------------

    if validation["passed"]:

        log(
            "Initial analysis passed validation."
        )

        return {
            "success": True,
            "mode": "initial",
            "prompt": prompt,
            "response": initial_response,
            "rawText": raw_text,
            "parsedResponse": parsed,
            "analyses": analyses,
            "validation": validation,
            "correctionAttempts": [],
        }

    # --------------------------------------------------------
    # Correction required
    # --------------------------------------------------------

    log(
        "Initial analysis requires correction."
    )

    correction_attempts: list[
        dict[str, Any]
    ] = []

    current_analyses = analyses
    current_validation = validation

    for attempt in range(
        1,
        MAX_CORRECTION_ATTEMPTS + 1,
    ):

        log(
            f"Correction attempt "
            f"{attempt}/{MAX_CORRECTION_ATTEMPTS}"
        )

        correction_prompt = (
            build_correction_prompt(
                articles=articles,
                analyses=current_analyses,
                validation=current_validation,
                geography_master=geography_master,
                sector_master=sector_master,
            )
        )

        correction_response = call_gemini(
            model=model,
            api_key=api_key,
            prompt=correction_prompt,
        )

        correction_record: dict[
            str,
            Any,
        ] = {
            "attempt": attempt,
            "prompt": correction_prompt,
            "response": correction_response,
        }

        if not correction_response["ok"]:

            correction_record[
                "success"
            ] = False

            correction_record[
                "error"
            ] = (
                "Correction Gemini request failed."
            )

            correction_attempts.append(
                correction_record
            )

            continue

        try:

            correction_raw_text = (
                extract_response_text(
                    correction_response
                )
            )

            correction_parsed = (
                parse_generated_json(
                    correction_raw_text
                )
            )

            corrected_analyses = (
                normalize_analysis_list(
                    correction_parsed
                )
            )

        except Exception as exc:

            correction_record[
                "success"
            ] = False

            correction_record[
                "error"
            ] = str(exc)

            correction_attempts.append(
                correction_record
            )

            continue

        corrected_validation = (
            validate_analysis(
                analyses=corrected_analyses,
                articles=articles,
                geography_master=set(
                    geography_master
                ),
                sector_master=set(
                    sector_master
                ),
            )
        )

        correction_record.update(
            {
                "rawText": correction_raw_text,
                "parsedResponse": correction_parsed,
                "analyses": corrected_analyses,
                "validation": corrected_validation,
                "success": corrected_validation[
                    "passed"
                ],
            }
        )

        correction_attempts.append(
            correction_record
        )

        if corrected_validation[
            "passed"
        ]:

            log(
                f"Correction attempt {attempt} "
                "passed validation."
            )

            return {
                "success": True,
                "mode": "corrected",
                "prompt": prompt,
                "response": initial_response,
                "rawText": raw_text,
                "parsedResponse": parsed,
                "analyses": corrected_analyses,
                "validation": corrected_validation,
                "initialValidation": validation,
                "initialAnalyses": analyses,
                "correctionAttempts": correction_attempts,
            }

        current_analyses = (
            corrected_analyses
        )

        current_validation = (
            corrected_validation
        )

    # --------------------------------------------------------
    # Correction failed
    # --------------------------------------------------------

    return {
        "success": False,
        "stage": "correction_validation",
        "prompt": prompt,
        "response": initial_response,
        "rawText": raw_text,
        "parsedResponse": parsed,
        "analyses": analyses,
        "validation": validation,
        "correctionAttempts": correction_attempts,
        "error": (
            "Correction failed to produce a fully "
            "valid analysis."
        ),
    }


# ============================================================
# BATCH RECORD
# ============================================================

def create_batch_record(
    *,
    batch_number: int,
    articles: list[dict[str, Any]],
    result: dict[str, Any],
    model: str,
) -> dict[str, Any]:

    initial_response = result.get(
        "response",
        {},
    )

    usage = {}

    if isinstance(
        initial_response,
        dict,
    ):

        raw = initial_response.get(
            "raw"
        )

        if isinstance(
            raw,
            dict,
        ):

            usage_metadata = raw.get(
                "usageMetadata"
            )

            if isinstance(
                usage_metadata,
                dict,
            ):
                usage = usage_metadata

    record = {
        "batchNumber": batch_number,
        "storedAt": utc_now_iso(),
        "model": model,
        "articleIds": [
            article["articleId"]
            for article in articles
        ],
        "request": {
            "articleCount": len(
                articles
            ),
            "promptCharacters": len(
                result.get(
                    "prompt",
                    "",
                )
            ),
            "responseSeconds": (
                initial_response.get(
                    "responseSeconds"
                )
            ),
            "maxOutputTokens": (
                gemini_max_output_tokens()
            ),
            "responseMimeType": (
                "application/json"
            ),
            "correctionMode": True,
            "correctionAttempts": len(
                result.get(
                    "correctionAttempts",
                    [],
                )
            ),
        },
        "analysisMode": result.get(
            "mode"
        ),
        "validation": result.get(
            "validation",
            {},
        ),
        "prompt": result.get(
            "prompt"
        ),
        "requestBody": initial_response.get(
            "requestBody"
        ),
        "parsedResponse": result.get(
            "parsedResponse"
        ),
        "analyses": result.get(
            "analyses",
            [],
        ),
        "rawGeminiResponse": initial_response.get(
            "raw"
        ),
        "inputArticles": articles,
        "geminiUsage": usage,
    }

    # --------------------------------------------------------
    # Preserve correction history
    # --------------------------------------------------------

    if result.get(
        "mode"
    ) == "corrected":

        record[
            "initialValidation"
        ] = result.get(
            "initialValidation"
        )

        record[
            "initialAnalyses"
        ] = result.get(
            "initialAnalyses"
        )

        record[
            "correctionAttempts"
        ] = result.get(
            "correctionAttempts"
        )

    return record


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    log("=" * 70)
    log(
        "VGrat FMS - MARKET NEWS AI ANALYZER"
    )
    log("=" * 70)

    # --------------------------------------------------------
    # API key
    # --------------------------------------------------------

    api_key = os.environ.get(
        "GEMINI_API_KEY"
    )

    if not api_key:

        log(
            "ERROR: GEMINI_API_KEY is missing."
        )

        return 1

    model = gemini_model()

    log(
        f"Gemini model: {model}"
    )

    log(
        f"Batch size: {BATCH_SIZE}"
    )

    log(
        f"Max batches: {MAX_BATCHES}"
    )

    log(
        f"Summary maximum: "
        f"{MAX_SUMMARY_SENTENCES} sentences"
    )

    # --------------------------------------------------------
    # Verify inputs
    # --------------------------------------------------------

    if not CURRENT_JSON.exists():

        log(
            f"ERROR: Missing "
            f"{CURRENT_JSON}"
        )

        return 1

    if not RESEARCH_FUNDS.exists():

        log(
            f"ERROR: Missing "
            f"{RESEARCH_FUNDS}"
        )

        return 1

    # --------------------------------------------------------
    # Load production news
    # --------------------------------------------------------

    log(
        f"Loading market news: "
        f"{CURRENT_JSON}"
    )

    current_data = load_json(
        CURRENT_JSON
    )

    raw_articles = extract_articles(
        current_data
    )

    log(
        f"Articles found: "
        f"{len(raw_articles)}"
    )

    normalized_articles = [
        normalize_article(
            article
        )
        for article in raw_articles
    ]

    # --------------------------------------------------------
    # Load Excel masters
    # --------------------------------------------------------

    (
        geography_master,
        sector_master,
    ) = load_research_masters()

    # --------------------------------------------------------
    # Load existing analysis
    # --------------------------------------------------------

    analysis = (
        load_existing_analysis()
    )

    completed_ids = (
        completed_article_ids(
            analysis
        )
    )

    log(
        f"Previously analyzed articles: "
        f"{len(completed_ids)}"
    )

    # --------------------------------------------------------
    # Find pending
    # --------------------------------------------------------

    pending = [
        article
        for article in normalized_articles
        if article["articleId"]
        not in completed_ids
    ]

    if PROCESS_NEWEST_FIRST:

        pending.sort(
            key=publication_sort_key,
            reverse=True,
        )

    log(
        f"Pending articles: "
        f"{len(pending)}"
    )

    if not pending:

        log(
            "No pending articles."
        )

        return 0

    # --------------------------------------------------------
    # Process batches
    # --------------------------------------------------------

    successful_batches = 0

    for batch_index in range(
        MAX_BATCHES
    ):

        start = (
            batch_index
            * BATCH_SIZE
        )

        end = (
            start
            + BATCH_SIZE
        )

        batch_articles = pending[
            start:end
        ]

        if not batch_articles:
            break

        batch_number = (
            len(
                analysis.get(
                    "batches",
                    [],
                )
            )
            + 1
        )

        log("")
        log("=" * 70)
        log(
            f"BATCH {batch_number}"
        )
        log("=" * 70)

        for index, article in enumerate(
            batch_articles,
            start=1,
        ):

            log(
                f"{index}. "
                f"{article['articleId']} | "
                f"{article.get('title')}"
            )

        # ----------------------------------------------------
        # Gemini
        # ----------------------------------------------------

        result = analyze_batch(
            articles=batch_articles,
            geography_master=geography_master,
            sector_master=sector_master,
            api_key=api_key,
            model=model,
        )

        # ----------------------------------------------------
        # Failure
        # ----------------------------------------------------

        if not result[
            "success"
        ]:

            log("")
            log("=" * 70)
            log(
                "BATCH FAILED"
            )
            log("=" * 70)

            log(
                str(
                    result.get(
                        "error",
                        "Unknown error.",
                    )
                )
            )

            validation = result.get(
                "validation"
            )

            if validation:

                log(
                    "Validation errors:"
                )

                for error in validation.get(
                    "errors",
                    [],
                ):

                    log(
                        f"  - {error}"
                    )

            log(
                "NO CHANGES WRITTEN."
            )

            return 1

        # ----------------------------------------------------
        # Create batch record
        # ----------------------------------------------------

        batch_record = (
            create_batch_record(
                batch_number=batch_number,
                articles=batch_articles,
                result=result,
                model=model,
            )
        )

        # ----------------------------------------------------
        # Update in memory
        # ----------------------------------------------------

        analysis[
            "batches"
        ].append(
            batch_record
        )

        analysis[
            "totalSuccessfulBatches"
        ] = len(
            analysis[
                "batches"
            ]
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
            for batch in analysis[
                "batches"
            ]
        )

        analysis[
            "updatedAt"
        ] = utc_now_iso()

        successful_batches += 1

        # ----------------------------------------------------
        # Write ONLY after successful validation
        # ----------------------------------------------------

        save_json_atomic(
            ANALYSIS_CURRENT,
            analysis,
        )

        log("")
        log(
            f"BATCH {batch_number} SUCCESS"
        )

        log(
            f"Analysis mode: "
            f"{result.get('mode')}"
        )

        log(
            f"Articles accepted: "
            f"{len(batch_articles)}"
        )

        log(
            f"Correction attempts: "
            f"{len(result.get('correctionAttempts', []))}"
        )

        log(
            f"Written: "
            f"{ANALYSIS_CURRENT}"
        )

    # --------------------------------------------------------
    # Final summary
    # --------------------------------------------------------

    log("")
    log("=" * 70)
    log(
        "ANALYSIS COMPLETE"
    )
    log("=" * 70)

    log(
        f"Successful batches this run: "
        f"{successful_batches}"
    )

    log(
        f"Total successful batches: "
        f"{analysis['totalSuccessfulBatches']}"
    )

    log(
        f"Total successful articles: "
        f"{analysis['totalSuccessfulArticles']}"
    )

    log(
        f"Output: "
        f"{ANALYSIS_CURRENT}"
    )

    return 0


if __name__ == "__main__":
    sys.exit(main())
