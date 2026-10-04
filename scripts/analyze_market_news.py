#!/usr/bin/env python3

"""
VGrat FMS - MARKET NEWS AI ANALYZER
===================================

TEST MODE
=========

Processes exactly ONE batch of FIVE newest pending market-news articles.

Production inputs are READ-ONLY:
    data/market_news/current.json
    Research Funds.xlsx

Test output:
    data/market_news/analysis/current.json

IMPORTANT
=========

The production market-news files are never modified.

A batch is written to analysis/current.json ONLY when:

1. Gemini request succeeds.
2. Gemini response contains valid JSON.
3. Exactly one analysis is returned for every article.
4. Every returned articleId matches an article in the batch.

Transient Gemini errors such as HTTP 503 are retried with bounded
exponential backoff:

    Initial request
        ↓
    503 / 429 / 500 / 502 / 504
        ↓
    wait 10 seconds
        ↓
    retry
        ↓
    wait 30 seconds
        ↓
    retry
        ↓
    wait 60 seconds
        ↓
    retry
        ↓
    success OR fail safely

No partial result is written.
"""

from __future__ import annotations

import hashlib
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests
from openpyxl import load_workbook


# ============================================================================
# PATHS
# ============================================================================

ROOT = Path(__file__).resolve().parent.parent

CURRENT_JSON = ROOT / "data" / "market_news" / "current.json"
RESEARCH_FUNDS = ROOT / "Research Funds.xlsx"

ANALYSIS_DIR = ROOT / "data" / "market_news" / "analysis"
ANALYSIS_CURRENT = ANALYSIS_DIR / "current.json"


# ============================================================================
# TEST CONFIGURATION
# ============================================================================

BATCH_SIZE = 5
MAX_BATCHES = 1
PROCESS_NEWEST_FIRST = True


# ============================================================================
# GEMINI CONFIGURATION
# ============================================================================

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()

GEMINI_MODEL = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.8-flash",
).strip()

GEMINI_TIMEOUT_SECONDS = int(
    os.getenv("GEMINI_TIMEOUT_SECONDS", "180")
)

GEMINI_MAX_OUTPUT_TOKENS = int(
    os.getenv("GEMINI_MAX_OUTPUT_TOKENS", "16384")
)


# Number of retries AFTER the initial request.
MAX_RETRIES = int(
    os.getenv("GEMINI_MAX_RETRIES", "3")
)

# Retry delays in seconds.
RETRY_DELAYS = [10, 30, 60]

# HTTP errors which may be temporary.
RETRYABLE_STATUS_CODES = {
    429,
    500,
    502,
    503,
    504,
}


# ============================================================================
# LOGGING
# ============================================================================

SCRIPT_START = time.monotonic()


def log(message: str = "") -> None:
    elapsed = time.monotonic() - SCRIPT_START
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(
        f"[{timestamp}] [+{elapsed:.2f}s] {message}",
        flush=True,
    )


def fail(message: str) -> None:
    log(message)
    sys.exit(1)


# ============================================================================
# GENERAL HELPERS
# ============================================================================

def utc_now() -> str:
    return (
        datetime.now(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )


def atomic_write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

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


def load_json(path: Path) -> Any:
    with path.open(
        "r",
        encoding="utf-8",
    ) as f:
        return json.load(f)


def first_nonempty(
    obj: dict[str, Any],
    keys: list[str],
    default: Any = "",
) -> Any:
    for key in keys:
        value = obj.get(key)

        if value is not None and value != "":
            return value

    return default


# ============================================================================
# ARTICLE NORMALISATION
# ============================================================================

def extract_articles(payload: Any) -> list[dict[str, Any]]:
    """
    Handles common market-news JSON envelopes.
    """

    if isinstance(payload, list):
        return [
            item
            for item in payload
            if isinstance(item, dict)
        ]

    if not isinstance(payload, dict):
        raise ValueError(
            "current.json must contain either an object or array."
        )

    for key in (
        "articles",
        "news",
        "items",
        "data",
        "results",
    ):
        value = payload.get(key)

        if isinstance(value, list):
            return [
                item
                for item in value
                if isinstance(item, dict)
            ]

    raise ValueError(
        "Could not find article list in current.json."
    )


def make_article_id(article: dict[str, Any]) -> str:
    """
    Prefer an existing source ID.

    Otherwise hash the article URL.

    As a final fallback, hash the canonical JSON representation.
    """

    existing_id = first_nonempty(
        article,
        [
            "id",
            "articleId",
            "article_id",
            "newsId",
            "news_id",
        ],
        "",
    )

    if existing_id:
        return str(existing_id)

    url = first_nonempty(
        article,
        [
            "url",
            "link",
            "articleUrl",
            "article_url",
        ],
        "",
    )

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


def normalise_article(
    article: dict[str, Any],
) -> dict[str, Any]:

    article_id = make_article_id(article)

    title = first_nonempty(
        article,
        [
            "title",
            "headline",
            "name",
        ],
        "",
    )

    summary = first_nonempty(
        article,
        [
            "summary",
            "description",
            "dek",
            "excerpt",
        ],
        "",
    )

    content = first_nonempty(
        article,
        [
            "content",
            "body",
            "text",
            "article",
        ],
        "",
    )

    url = first_nonempty(
        article,
        [
            "url",
            "link",
            "articleUrl",
            "article_url",
        ],
        "",
    )

    published_at = first_nonempty(
        article,
        [
            "publishedAt",
            "published_at",
            "published",
            "pubDate",
            "pub_date",
            "date",
            "timestamp",
        ],
        "",
    )

    source = first_nonempty(
        article,
        [
            "source",
            "publisher",
            "provider",
        ],
        "",
    )

    return {
        "articleId": article_id,
        "title": str(title),
        "summary": str(summary),
        "content": str(content),
        "url": str(url),
        "publishedAt": str(published_at),
        "source": str(source),
        "original": article,
    }


# ============================================================================
# DATE SORTING
# ============================================================================

def parse_date_for_sort(value: Any) -> float:
    """
    Convert common ISO timestamps to a sortable timestamp.

    Unknown / missing dates return 0.

    This preserves the existing behaviour where an article without a
    recognised publication field does not crash the pipeline.
    """

    if value is None:
        return 0.0

    value = str(value).strip()

    if not value:
        return 0.0

    candidates = [
        value,
        value.replace("Z", "+00:00"),
    ]

    for candidate in candidates:
        try:
            dt = datetime.fromisoformat(candidate)

            if dt.tzinfo is None:
                dt = dt.replace(
                    tzinfo=timezone.utc
                )

            return dt.timestamp()

        except ValueError:
            pass

    return 0.0


# ============================================================================
# RESEARCH FUNDS
# ============================================================================

def load_research_funds() -> dict[str, Any]:
    """
    Loads Research Funds.xlsx as READ-ONLY reference data.

    Worksheet preference:
        - Fund Research
        - worksheet index 1 = Geography master
        - worksheet index 2 = Sector master

    No AI or web research is performed here.
    """

    log("Loading Research Funds.xlsx...")

    if not RESEARCH_FUNDS.exists():
        raise FileNotFoundError(
            f"Research Funds.xlsx not found: {RESEARCH_FUNDS}"
        )

    workbook = load_workbook(
        RESEARCH_FUNDS,
        read_only=True,
        data_only=True,
    )

    try:
        sheet_names = workbook.sheetnames

        # ------------------------------------------------------------
        # Fund Research
        # ------------------------------------------------------------

        if "Fund Research" in sheet_names:
            ws_funds = workbook["Fund Research"]
        elif len(sheet_names) >= 1:
            ws_funds = workbook[sheet_names[0]]
        else:
            raise ValueError(
                "Research Funds.xlsx contains no worksheets."
            )

        rows = list(
            ws_funds.iter_rows(
                values_only=True
            )
        )

        if not rows:
            fund_rows = []
        else:
            headers = [
                str(value).strip()
                if value is not None
                else ""
                for value in rows[0]
            ]

            fund_rows = []

            for row in rows[1:]:
                record = {}

                for index, header in enumerate(headers):
                    if not header:
                        continue

                    value = (
                        row[index]
                        if index < len(row)
                        else None
                    )

                    record[header] = value

                # Ignore completely empty rows.
                if any(
                    value not in (None, "")
                    for value in record.values()
                ):
                    fund_rows.append(record)

        # ------------------------------------------------------------
        # Geography master
        # ------------------------------------------------------------

        geography_master = []

        if len(sheet_names) > 1:
            ws_geo = workbook.worksheets[1]

            for row in ws_geo.iter_rows(
                values_only=True
            ):
                values = [
                    value
                    for value in row
                    if value not in (None, "")
                ]

                if values:
                    geography_master.append(values)

        # ------------------------------------------------------------
        # Sector master
        # ------------------------------------------------------------

        sector_master = []

        if len(sheet_names) > 2:
            ws_sector = workbook.worksheets[2]

            for row in ws_sector.iter_rows(
                values_only=True
            ):
                values = [
                    value
                    for value in row
                    if value not in (None, "")
                ]

                if values:
                    sector_master.append(values)

    finally:
        workbook.close()

    log(
        f"Research Funds rows loaded: {len(fund_rows)}"
    )

    return {
        "fundResearch": fund_rows,
        "masterCategories": {
            "geography": geography_master,
            "sector": sector_master,
        },
    }


# ============================================================================
# ANALYSIS FILE
# ============================================================================

def empty_analysis_file() -> dict[str, Any]:
    now = utc_now()

    return {
        "version": 1,
        "createdAt": now,
        "updatedAt": now,
        "totalSuccessfulBatches": 0,
        "totalSuccessfulArticles": 0,
        "batches": [],
    }


def load_analysis_file() -> dict[str, Any]:
    if not ANALYSIS_CURRENT.exists():
        return empty_analysis_file()

    log(
        "Loading existing analysis/current.json..."
    )

    data = load_json(ANALYSIS_CURRENT)

    if not isinstance(data, dict):
        raise ValueError(
            "analysis/current.json must contain a JSON object."
        )

    data.setdefault("version", 1)
    data.setdefault("createdAt", utc_now())
    data.setdefault("updatedAt", utc_now())
    data.setdefault(
        "totalSuccessfulBatches",
        0,
    )
    data.setdefault(
        "totalSuccessfulArticles",
        0,
    )
    data.setdefault("batches", [])

    if not isinstance(
        data["batches"],
        list,
    ):
        raise ValueError(
            "analysis/current.json 'batches' must be a list."
        )

    return data


def get_completed_article_ids(
    analysis_data: dict[str, Any],
) -> set[str]:

    completed: set[str] = set()

    for batch in analysis_data.get(
        "batches",
        [],
    ):

        if not isinstance(batch, dict):
            continue

        article_ids = batch.get(
            "articleIds",
            [],
        )

        if isinstance(article_ids, list):
            for article_id in article_ids:
                if article_id:
                    completed.add(
                        str(article_id)
                    )

        analyses = batch.get(
            "analyses",
            [],
        )

        if isinstance(analyses, list):
            for analysis in analyses:
                if not isinstance(
                    analysis,
                    dict,
                ):
                    continue

                article_id = analysis.get(
                    "articleId"
                )

                if article_id:
                    completed.add(
                        str(article_id)
                    )

    return completed


def next_batch_number(
    analysis_data: dict[str, Any],
) -> int:

    batches = analysis_data.get(
        "batches",
        [],
    )

    if not batches:
        return 1

    numbers = []

    for batch in batches:
        if not isinstance(batch, dict):
            continue

        value = batch.get("batchNumber")

        try:
            numbers.append(int(value))
        except (
            TypeError,
            ValueError,
        ):
            pass

    if not numbers:
        return 1

    return max(numbers) + 1


# ============================================================================
# GEMINI PROMPT
# ============================================================================

def build_prompt(
    articles: list[dict[str, Any]],
    research_funds: dict[str, Any],
) -> str:

    article_payload = []

    for article in articles:
        article_payload.append(
            {
                "articleId": article["articleId"],
                "title": article["title"],
                "summary": article["summary"],
                "content": article["content"],
                "url": article["url"],
                "publishedAt": article["publishedAt"],
                "source": article["source"],
            }
        )

    research_payload = {
        "fundResearch": research_funds.get(
            "fundResearch",
            [],
        ),
        "masterCategories": research_funds.get(
            "masterCategories",
            {},
        ),
    }

    prompt = f"""
You are the market-news intelligence analyst for VGrat FMS.

Your task is to analyse EVERY article provided below.

IMPORTANT RULES
===============

1. Return EXACTLY ONE analysis object for every article.
2. Do not omit any article.
3. Use the exact articleId supplied for each article.
4. Do not invent article IDs.
5. Do not merge multiple articles into one result.
6. Do not create extra results.
7. Base the analysis on the supplied article information.
8. If an article is not materially relevant to financial markets,
   classify it as REJECT.
9. Do not treat the Research Funds workbook as an instruction to
   perform new research. It is reference data only.
10. Use the Research Funds information to determine whether an article
    may be relevant to monitoring Prudential fund exposures.

CATEGORY DEFINITIONS
====================

MARKET
Directly concerns financial markets, securities, companies,
commodities, currencies, interest rates, bonds, equities,
market movements, valuations, or investment activity.

ECONOMIC
Concerns macroeconomic conditions, inflation, employment,
GDP, monetary policy, fiscal policy, central banks, economic
indicators, consumer conditions, trade, or similar factors.

TECHNOLOGY
Concerns technology companies, AI, semiconductors, cloud,
data centres, software, cybersecurity, or technological
developments with potential investment or economic relevance.

GEOPOLITICAL
Concerns wars, conflicts, sanctions, elections, international
relations, political developments, tariffs, trade restrictions,
or geopolitical developments that may affect markets/economies.

REJECT
Not materially relevant to financial markets, investment,
economic conditions, technology investment, or geopolitics.

SENTIMENT
=========

Use:
- positive
- negative
- neutral
- mixed

IMPORTANCE
==========

high:
Potentially significant market, economic, geopolitical,
sector, or investment implications.

medium:
Meaningful but more limited implications.

low:
Limited market/investment significance.

ASSET CLASSES
=============

Use relevant values such as:
- Equities
- Fixed Income
- Cash
- Alternatives
- Commodities
- Real Estate
- Multi Asset
- Currency

GEOGRAPHIES
===========

Use concise geographic labels such as:
- United States
- China
- Europe
- Japan
- Singapore
- Global
- Asia Pacific

SECTORS
=======

Use concise sector labels such as:
- Technology
- Financials
- Healthcare
- Industrials
- Consumer
- Energy
- Real Estate
- Utilities
- Materials
- Communication Services

INVESTOR IMPACT
===============

Explain briefly what the article could mean for investors.

REASONING
=========

Explain the classification and importance in concise terms.
Do not repeat the entire article.

FUND MONITORING RELEVANT
========================

Set true if the article could reasonably warrant monitoring
fund exposure based on the supplied Research Funds reference data.

Set false when it has no meaningful connection.

RESEARCH FUNDS REFERENCE
========================

{json.dumps(
    research_payload,
    ensure_ascii=False,
    indent=2,
)}

ARTICLES TO ANALYSE
===================

{json.dumps(
    article_payload,
    ensure_ascii=False,
    indent=2,
)}

Return JSON only.
""".strip()

    return prompt


# ============================================================================
# GEMINI REQUEST
# ============================================================================

def build_response_schema() -> dict[str, Any]:

    return {
        "type": "ARRAY",
        "items": {
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
                        "neutral",
                        "mixed",
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
                "assetClasses",
                "geographies",
                "sectors",
                "investorImpact",
                "reasoning",
                "fundMonitoringRelevant",
            ],
        },
    }


def make_gemini_request(
    prompt: str,
) -> tuple[Any, dict[str, Any], float, int]:

    if not GEMINI_API_KEY:
        raise RuntimeError(
            "GEMINI_API_KEY environment variable is not set."
        )

    endpoint = (
        "https://generativelanguage.googleapis.com/v1beta/"
        f"models/{GEMINI_MODEL}:generateContent"
    )

    params = {
        "key": GEMINI_API_KEY,
    }

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
            "temperature": 0.1,
            "maxOutputTokens": GEMINI_MAX_OUTPUT_TOKENS,
            "responseMimeType": "application/json",
            "responseSchema": build_response_schema(),
        },
    }

    total_attempts = MAX_RETRIES + 1

    last_response = None
    last_exception = None

    for attempt in range(1, total_attempts + 1):

        request_started = time.monotonic()

        log(
            f"Sending Gemini request "
            f"(attempt {attempt}/{total_attempts})..."
        )

        try:

            response = requests.post(
                endpoint,
                params=params,
                json=body,
                timeout=GEMINI_TIMEOUT_SECONDS,
            )

            request_seconds = (
                time.monotonic()
                - request_started
            )

            last_response = response

            # --------------------------------------------------------
            # SUCCESS
            # --------------------------------------------------------

            if response.ok:

                log(
                    f"Gemini HTTP {response.status_code} "
                    f"after {request_seconds:.2f}s."
                )

                try:
                    response_json = response.json()
                except ValueError as exc:
                    raise RuntimeError(
                        "Gemini returned a non-JSON HTTP response."
                    ) from exc

                return (
                    response_json,
                    body,
                    request_seconds,
                    attempt,
                )

            # --------------------------------------------------------
            # RETRYABLE ERROR
            # --------------------------------------------------------

            if (
                response.status_code
                in RETRYABLE_STATUS_CODES
            ):

                log(
                    f"Gemini HTTP {response.status_code} "
                    f"on attempt {attempt}/{total_attempts}."
                )

                if attempt < total_attempts:

                    delay_index = min(
                        attempt - 1,
                        len(RETRY_DELAYS) - 1,
                    )

                    delay = RETRY_DELAYS[
                        delay_index
                    ]

                    log(
                        f"Transient Gemini error. "
                        f"Waiting {delay}s before retry..."
                    )

                    time.sleep(delay)

                    continue

                # All retries exhausted.
                log(
                    "Gemini transient-error retries exhausted."
                )

                raise RuntimeError(
                    f"Gemini API error "
                    f"{response.status_code}: "
                    f"{response.text}"
                )

            # --------------------------------------------------------
            # NON-RETRYABLE ERROR
            # --------------------------------------------------------

            raise RuntimeError(
                f"Gemini API error "
                f"{response.status_code}: "
                f"{response.text}"
            )

        except requests.RequestException as exc:

            last_exception = exc

            request_seconds = (
                time.monotonic()
                - request_started
            )

            log(
                f"Gemini network error on attempt "
                f"{attempt}/{total_attempts}: "
                f"{exc}"
            )

            if attempt < total_attempts:

                delay_index = min(
                    attempt - 1,
                    len(RETRY_DELAYS) - 1,
                )

                delay = RETRY_DELAYS[
                    delay_index
                ]

                log(
                    f"Waiting {delay}s before retry..."
                )

                time.sleep(delay)

                continue

            raise RuntimeError(
                "Gemini request failed after "
                f"{total_attempts} attempts: "
                f"{exc}"
            ) from exc

    if last_exception:
        raise RuntimeError(
            f"Gemini request failed: {last_exception}"
        )

    if last_response is not None:
        raise RuntimeError(
            f"Gemini request failed with "
            f"HTTP {last_response.status_code}: "
            f"{last_response.text}"
        )

    raise RuntimeError(
        "Gemini request failed for an unknown reason."
    )


# ============================================================================
# GEMINI RESPONSE EXTRACTION
# ============================================================================

def extract_gemini_text(
    response_json: dict[str, Any],
) -> str:

    candidates = []

    candidates.append(
        response_json.get("candidates")
    )

    if not isinstance(
        response_json.get("candidates"),
        list,
    ):
        raise ValueError(
            "Gemini response contains no candidates."
        )

    candidates = response_json["candidates"]

    if not candidates:
        raise ValueError(
            "Gemini response contains an empty candidates array."
        )

    candidate = candidates[0]

    if not isinstance(candidate, dict):
        raise ValueError(
            "Gemini candidate is not an object."
        )

    content = candidate.get("content")

    if not isinstance(content, dict):
        raise ValueError(
            "Gemini candidate contains no content object."
        )

    parts = content.get("parts")

    if not isinstance(parts, list):
        raise ValueError(
            "Gemini candidate contains no parts array."
        )

    texts = []

    for part in parts:

        if not isinstance(part, dict):
            continue

        text_value = part.get("text")

        if isinstance(
            text_value,
            str,
        ):
            texts.append(text_value)

    if not texts:
        raise ValueError(
            "Gemini response contained no text output."
        )

    return "".join(texts).strip()


def parse_gemini_json(
    response_json: dict[str, Any],
) -> Any:

    text_output = extract_gemini_text(
        response_json
    )

    try:
        return json.loads(text_output)

    except json.JSONDecodeError as exc:

        # Some models may occasionally wrap JSON in markdown
        # despite responseMimeType.
        cleaned = text_output.strip()

        if cleaned.startswith("```"):
            lines = cleaned.splitlines()

            if lines:
                lines = lines[1:]

            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]

            cleaned = "\n".join(lines).strip()

        try:
            return json.loads(cleaned)

        except json.JSONDecodeError:
            raise ValueError(
                "Gemini returned invalid JSON.\n"
                f"Raw output:\n{text_output}"
            ) from exc


# ============================================================================
# RESPONSE VALIDATION
# ============================================================================

def validate_analysis(
    parsed: Any,
    articles: list[dict[str, Any]],
) -> list[dict[str, Any]]:

    if not isinstance(parsed, list):
        raise ValueError(
            "Gemini output must be a JSON array."
        )

    expected_ids = [
        article["articleId"]
        for article in articles
    ]

    expected_id_set = set(expected_ids)

    if len(parsed) != len(articles):
        raise ValueError(
            "Gemini returned "
            f"{len(parsed)} analyses for "
            f"{len(articles)} articles."
        )

    analyses = []

    for index, item in enumerate(parsed):

        if not isinstance(item, dict):
            raise ValueError(
                f"Analysis #{index + 1} is not an object."
            )

        article_id = item.get("articleId")

        if not article_id:
            raise ValueError(
                f"Analysis #{index + 1} has no articleId."
            )

        article_id = str(article_id)

        if article_id not in expected_id_set:
            raise ValueError(
                "Gemini returned an unknown articleId: "
                f"{article_id}"
            )

        analyses.append(item)

    returned_ids = [
        str(item["articleId"])
        for item in analyses
    ]

    if len(set(returned_ids)) != len(returned_ids):
        raise ValueError(
            "Gemini returned duplicate articleIds."
        )

    returned_id_set = set(returned_ids)

    if returned_id_set != expected_id_set:

        missing = sorted(
            expected_id_set - returned_id_set
        )

        extra = sorted(
            returned_id_set - expected_id_set
        )

        raise ValueError(
            "Article ID mismatch. "
            f"Missing={missing}, Extra={extra}"
        )

    # Return analyses in the SAME order as the input batch.
    by_id = {
        str(item["articleId"]): item
        for item in analyses
    }

    ordered = [
        by_id[article_id]
        for article_id in expected_ids
    ]

    return ordered


# ============================================================================
# MAIN
# ============================================================================

def main() -> None:

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
        f"Gemini max retries after initial request: "
        f"{MAX_RETRIES}"
    )
    log(
        f"Retry delays: {RETRY_DELAYS[:MAX_RETRIES]}"
    )
    log(
        f"Input: {CURRENT_JSON}"
    )
    log(
        f"Output: {ANALYSIS_CURRENT}"
    )
    log()

    # ========================================================================
    # INPUT CHECKS
    # ========================================================================

    if not CURRENT_JSON.exists():
        fail(
            f"Production current.json not found: "
            f"{CURRENT_JSON}"
        )

    if not RESEARCH_FUNDS.exists():
        fail(
            f"Research Funds.xlsx not found: "
            f"{RESEARCH_FUNDS}"
        )

    if not GEMINI_API_KEY:
        fail(
            "GEMINI_API_KEY is not configured."
        )

    # ========================================================================
    # LOAD CURRENT NEWS
    # ========================================================================

    log(
        "Loading production current.json..."
    )

    current_payload = load_json(
        CURRENT_JSON
    )

    raw_articles = extract_articles(
        current_payload
    )

    log(
        f"Articles found: {len(raw_articles)}"
    )

    articles = [
        normalise_article(article)
        for article in raw_articles
    ]

    # ========================================================================
    # LOAD RESEARCH FUNDS
    # ========================================================================

    research_funds = load_research_funds()

    # ========================================================================
    # LOAD EXISTING TEST ANALYSIS
    # ========================================================================

    analysis_data = load_analysis_file()

    completed_ids = get_completed_article_ids(
        analysis_data
    )

    log(
        "Previously successful articles in test analysis: "
        f"{len(completed_ids)}"
    )

    # ========================================================================
    # FIND PENDING ARTICLES
    # ========================================================================

    pending_articles = [
        article
        for article in articles
        if article["articleId"]
        not in completed_ids
    ]

    if PROCESS_NEWEST_FIRST:
        pending_articles.sort(
            key=lambda article: parse_date_for_sort(
                article["publishedAt"]
            ),
            reverse=True,
        )

    log(
        f"Pending articles: {len(pending_articles)}"
    )
    log()

    if not pending_articles:
        log(
            "No pending articles remain."
        )
        log(
            "Nothing to analyse."
        )
        return

    # ========================================================================
    # SELECT EXACTLY ONE BATCH
    # ========================================================================

    batch_articles = pending_articles[
        :BATCH_SIZE
    ]

    if len(batch_articles) < BATCH_SIZE:
        log(
            f"Only {len(batch_articles)} pending "
            f"article(s) available."
        )

        # In test mode we want exactly five.
        # Do not make a smaller test batch.
        if len(batch_articles) != BATCH_SIZE:
            fail(
                "TEST MODE requires exactly "
                f"{BATCH_SIZE} pending articles."
            )

    log(
        "Selected batch:"
    )

    for index, article in enumerate(
        batch_articles,
        start=1,
    ):

        log(
            f"  {index}. "
            f"{article['articleId']} | "
            f"{article['publishedAt']} | "
            f"{article['title']}"
        )

    log()

    # ========================================================================
    # BUILD PROMPT
    # ========================================================================

    prompt = build_prompt(
        batch_articles,
        research_funds,
    )

    log(
        "Preparing ONE Gemini request for "
        f"{len(batch_articles)} articles..."
    )

    request_started_at = utc_now()
    request_timer = time.monotonic()

    # ========================================================================
    # GEMINI REQUEST WITH RETRIES
    # ========================================================================

    try:

        (
            gemini_response,
            gemini_request_body,
            request_seconds,
            attempts_used,
        ) = make_gemini_request(prompt)

    except Exception as exc:

        log()
        log("GEMINI REQUEST FAILED")
        log(str(exc))
        log()
        log(
            "Nothing was written to "
            "analysis/current.json."
        )

        sys.exit(1)

    request_finished_at = utc_now()

    # Ensure timing includes all retry attempts.
    total_elapsed_seconds = (
        time.monotonic()
        - request_timer
    )

    log(
        f"Gemini request completed in "
        f"{total_elapsed_seconds:.2f}s."
    )

    log(
        f"Gemini attempts used: {attempts_used}"
    )

    # ========================================================================
    # PARSE RESPONSE
    # ========================================================================

    log(
        "Parsing Gemini response..."
    )

    try:

        parsed_response = parse_gemini_json(
            gemini_response
        )

    except Exception as exc:

        log()
        log("GEMINI RESPONSE PARSE FAILED")
        log(str(exc))
        log()
        log(
            "Nothing was written to "
            "analysis/current.json."
        )

        sys.exit(1)

    # ========================================================================
    # VALIDATE RESPONSE
    # ========================================================================

    log(
        "Validating Gemini response..."
    )

    try:

        validated_analyses = validate_analysis(
            parsed_response,
            batch_articles,
        )

    except Exception as exc:

        log()
        log("GEMINI RESPONSE VALIDATION FAILED")
        log(str(exc))
        log()
        log(
            "Nothing was written to "
            "analysis/current.json."
        )

        sys.exit(1)

    log(
        "Validation successful."
    )

    log(
        f"Exactly {len(validated_analyses)} "
        "article analyses received."
    )

    # ========================================================================
    # BUILD SUCCESSFUL BATCH RECORD
    # ========================================================================

    batch_number = next_batch_number(
        analysis_data
    )

    batch_stored_at = utc_now()

    article_ids = [
        article["articleId"]
        for article in batch_articles
    ]

    batch_record = {
        "batchNumber": batch_number,
        "storedAt": batch_stored_at,
        "model": GEMINI_MODEL,
        "articleIds": article_ids,

        "request": {
            "requestStartedAt": request_started_at,
            "requestFinishedAt": request_finished_at,
            "responseSeconds": round(
                total_elapsed_seconds,
                3,
            ),
            "geminiAttempts": attempts_used,
            "articleCount": len(batch_articles),
        },

        "validation": {
            "success": True,
            "expectedArticleCount": len(
                batch_articles
            ),
            "returnedArticleCount": len(
                validated_analyses
            ),
            "articleIdsMatch": True,
            "duplicates": False,
        },

        "prompt": prompt,

        "requestBody": gemini_request_body,

        "parsedResponse": parsed_response,

        "analyses": validated_analyses,

        "rawGeminiResponse": gemini_response,

        "inputArticles": batch_articles,
    }

    # ========================================================================
    # UPDATE ANALYSIS FILE
    # ========================================================================

    analysis_data["batches"].append(
        batch_record
    )

    analysis_data[
        "totalSuccessfulBatches"
    ] = len(
        analysis_data["batches"]
    )

    analysis_data[
        "totalSuccessfulArticles"
    ] = sum(
        len(batch.get("articleIds", []))
        for batch in analysis_data["batches"]
        if isinstance(batch, dict)
    )

    analysis_data["updatedAt"] = (
        batch_stored_at
    )

    # ========================================================================
    # ONLY NOW WRITE THE FILE
    # ========================================================================

    log()
    log(
        "Writing successful batch to "
        "analysis/current.json..."
    )

    atomic_write_json(
        ANALYSIS_CURRENT,
        analysis_data,
    )

    log(
        "Analysis successfully written."
    )

    log(
        f"File: {ANALYSIS_CURRENT}"
    )

    log(
        f"Successful batches: "
        f"{analysis_data['totalSuccessfulBatches']}"
    )

    log(
        f"Successful articles: "
        f"{analysis_data['totalSuccessfulArticles']}"
    )

    log(
        f"Latest batch: {batch_number}"
    )

    log(
        f"Batch articles: {len(batch_articles)}"
    )

    log(
        f"Gemini attempts: {attempts_used}"
    )

    log(
        f"Response time: "
        f"{total_elapsed_seconds:.2f}s"
    )

    log()
    log(
        "TEST BATCH COMPLETED SUCCESSFULLY."
    )


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    main()
