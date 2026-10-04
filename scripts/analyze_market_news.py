#!/usr/bin/env python3

"""
VGrat FMS - MARKET NEWS AI ANALYZER
===================================

First production-data batch test:
- Reads data/market_news/current.json as READ-ONLY
- Reads Research Funds.xlsx as READ-ONLY
- Selects newest 5 pending articles
- Sends all 5 articles in ONE Gemini request
- Validates exactly 5 analysis results
- Writes only after the entire batch succeeds
- Creates data/market_news/analysis/current.json if needed
- Stores every successful prompt/request/response
- No retries in this first batch test
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
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

try:
    from openpyxl import load_workbook
except ImportError:
    print("ERROR: openpyxl is required.")
    sys.exit(1)


# ============================================================
# PATHS
# ============================================================

ROOT = Path(__file__).resolve().parent.parent

CURRENT_JSON = ROOT / "data" / "market_news" / "current.json"
RESEARCH_FUNDS = ROOT / "Research Funds.xlsx"

ANALYSIS_DIR = ROOT / "data" / "market_news" / "analysis"
ANALYSIS_CURRENT = ANALYSIS_DIR / "current.json"


# ============================================================
# TEST CONFIGURATION
# ============================================================

BATCH_SIZE = 5
MAX_BATCHES = 1
PROCESS_NEWEST_FIRST = True

DEFAULT_MODEL = "gemini-3.8-flash"
DEFAULT_TIMEOUT = 180
DEFAULT_MAX_OUTPUT_TOKENS = 16384


# ============================================================
# LOGGING
# ============================================================

START_TIME = time.monotonic()


def log(message: str = "") -> None:
    elapsed = time.monotonic() - START_TIME
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(
        f"[{timestamp}] [+{elapsed:.2f}s] {message}",
        flush=True,
    )


# ============================================================
# HELPERS
# ============================================================

def now_utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def canonical_json(value: Any) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def make_hash(value: Any) -> str:
    return hashlib.sha256(
        canonical_json(value).encode("utf-8")
    ).hexdigest()


def first_value(obj: dict[str, Any], keys: list[str]) -> Any:
    for key in keys:
        value = obj.get(key)

        if value is not None and value != "":
            return value

    return None


# ============================================================
# ARTICLE NORMALIZATION
# ============================================================

def article_id(article: dict[str, Any]) -> str:
    existing = first_value(
        article,
        [
            "id",
            "articleId",
            "article_id",
            "newsId",
            "news_id",
        ],
    )

    if existing is not None:
        return str(existing)

    url = first_value(
        article,
        [
            "url",
            "link",
            "articleUrl",
        ],
    )

    if url:
        return hashlib.sha256(
            str(url).encode("utf-8")
        ).hexdigest()

    return make_hash(article)


def normalize_article(article: dict[str, Any]) -> dict[str, Any]:
    return {
        "articleId": article_id(article),
        "title": first_value(
            article,
            ["title", "headline", "name"],
        ),
        "description": first_value(
            article,
            ["description", "summary", "snippet", "text"],
        ),
        "url": first_value(
            article,
            ["url", "link", "articleUrl"],
        ),
        "published": first_value(
            article,
            [
                "published",
                "publishedAt",
                "published_at",
                "date",
                "datetime",
                "timestamp",
            ],
        ),
        "source": first_value(
            article,
            ["source", "publisher", "site"],
        ),
    }


# ============================================================
# LOAD CURRENT NEWS
# ============================================================

def find_article_list(data: Any) -> list[dict[str, Any]]:
    if isinstance(data, list):
        return [
            x for x in data
            if isinstance(x, dict)
        ]

    if not isinstance(data, dict):
        return []

    for key in [
        "articles",
        "news",
        "items",
        "data",
        "results",
    ]:
        value = data.get(key)

        if isinstance(value, list):
            return [
                x for x in value
                if isinstance(x, dict)
            ]

    return []


def load_current_articles() -> list[dict[str, Any]]:
    log("=" * 72)
    log("LOADING PRODUCTION CURRENT NEWS")
    log("=" * 72)

    if not CURRENT_JSON.exists():
        raise FileNotFoundError(
            f"Missing production input: {CURRENT_JSON}"
        )

    log(f"Input: {CURRENT_JSON}")

    with CURRENT_JSON.open(
        "r",
        encoding="utf-8",
    ) as f:
        data = json.load(f)

    raw_articles = find_article_list(data)

    articles = [
        normalize_article(article)
        for article in raw_articles
    ]

    log(f"Articles discovered: {len(articles)}")

    return articles


# ============================================================
# DATE SORTING
# ============================================================

def sort_key(article: dict[str, Any]) -> str:
    value = article.get("published")

    if value is None:
        return ""

    return str(value)


# ============================================================
# LOAD RESEARCH FUNDS
# ============================================================

def load_research_context() -> dict[str, Any]:
    """
    Research Funds.xlsx is treated as READ-ONLY.

    It is already post-research. No AI/web research is performed here.
    """

    log("=" * 72)
    log("LOADING RESEARCH FUNDS CONTEXT")
    log("=" * 72)

    if not RESEARCH_FUNDS.exists():
        raise FileNotFoundError(
            f"Missing production input: {RESEARCH_FUNDS}"
        )

    log(f"Input: {RESEARCH_FUNDS}")

    workbook = load_workbook(
        RESEARCH_FUNDS,
        read_only=True,
        data_only=True,
    )

    context: dict[str, Any] = {
        "fundResearch": [],
        "geographyMaster": [],
        "sectorMaster": [],
    }

    try:
        # ----------------------------------------------------
        # Fund Research
        # ----------------------------------------------------

        if "Fund Research" in workbook.sheetnames:
            ws = workbook["Fund Research"]

            rows = list(
                ws.iter_rows(
                    values_only=True
                )
            )

            if rows:
                headers = [
                    str(x).strip()
                    if x is not None
                    else ""
                    for x in rows[0]
                ]

                for row in rows[1:]:
                    item = {}

                    for index, header in enumerate(headers):
                        if not header:
                            continue

                        value = (
                            row[index]
                            if index < len(row)
                            else None
                        )

                        if value is not None:
                            item[header] = value

                    if item:
                        context["fundResearch"].append(item)

        # ----------------------------------------------------
        # Geography master = worksheet index 1
        # ----------------------------------------------------

        if len(workbook.worksheets) > 1:
            ws = workbook.worksheets[1]

            for row in ws.iter_rows(values_only=True):
                values = [
                    value
                    for value in row
                    if value is not None
                    and str(value).strip()
                ]

                if values:
                    context["geographyMaster"].append(
                        values
                    )

        # ----------------------------------------------------
        # Sector master = worksheet index 2
        # ----------------------------------------------------

        if len(workbook.worksheets) > 2:
            ws = workbook.worksheets[2]

            for row in ws.iter_rows(values_only=True):
                values = [
                    value
                    for value in row
                    if value is not None
                    and str(value).strip()
                ]

                if values:
                    context["sectorMaster"].append(
                        values
                    )

    finally:
        workbook.close()

    log(
        "Fund Research rows: "
        f"{len(context['fundResearch'])}"
    )

    log(
        "Geography master rows: "
        f"{len(context['geographyMaster'])}"
    )

    log(
        "Sector master rows: "
        f"{len(context['sectorMaster'])}"
    )

    return context


# ============================================================
# LOAD EXISTING ANALYSIS
# ============================================================

def load_existing_analysis() -> dict[str, Any]:
    if not ANALYSIS_CURRENT.exists():
        return {
            "version": 1,
            "createdAt": now_utc(),
            "updatedAt": now_utc(),
            "totalSuccessfulBatches": 0,
            "totalSuccessfulArticles": 0,
            "batches": [],
        }

    with ANALYSIS_CURRENT.open(
        "r",
        encoding="utf-8",
    ) as f:
        data = json.load(f)

    if not isinstance(data, dict):
        raise ValueError(
            "analysis/current.json must contain a JSON object."
        )

    data.setdefault("version", 1)
    data.setdefault("createdAt", now_utc())
    data.setdefault("updatedAt", now_utc())
    data.setdefault("totalSuccessfulBatches", 0)
    data.setdefault("totalSuccessfulArticles", 0)
    data.setdefault("batches", [])

    return data


def completed_article_ids(
    analysis: dict[str, Any],
) -> set[str]:

    completed: set[str] = set()

    for batch in analysis.get("batches", []):
        if not isinstance(batch, dict):
            continue

        for value in batch.get("articleIds", []):
            completed.add(str(value))

        for item in batch.get("analyses", []):
            if not isinstance(item, dict):
                continue

            value = item.get("articleId")

            if value is not None:
                completed.add(str(value))

    return completed


# ============================================================
# PROMPT
# ============================================================

def build_prompt(
    articles: list[dict[str, Any]],
    research_context: dict[str, Any],
) -> str:

    geography_master = research_context.get(
        "geographyMaster",
        [],
    )

    sector_master = research_context.get(
        "sectorMaster",
        [],
    )

    article_payload = json.dumps(
        articles,
        ensure_ascii=False,
        indent=2,
    )

    geography_payload = json.dumps(
        geography_master,
        ensure_ascii=False,
        indent=2,
    )

    sector_payload = json.dumps(
        sector_master,
        ensure_ascii=False,
        indent=2,
    )

    return f"""
You are the market-news analysis engine for VGrat FMS.

Analyze every supplied article independently.

IMPORTANT:
- Return EXACTLY ONE analysis object for EVERY supplied article.
- Do not omit any article.
- Preserve every articleId exactly.
- Do not invent article IDs.
- Return valid JSON only.
- Do not return markdown.
- Do not add commentary outside the JSON.
- The result must be a JSON array.

For each article return:

{{
  "articleId": "exact supplied articleId",
  "relevant": true,
  "category": "MARKET | ECONOMIC | TECHNOLOGY | GEOPOLITICAL | REJECT",
  "sentiment": "positive | neutral | negative | mixed",
  "importance": "low | medium | high | critical",
  "assetClasses": [],
  "geographies": [],
  "sectors": [],
  "investorImpact": "",
  "reasoning": "",
  "fundMonitoringRelevant": true
}}

Rules:

1. category must be one of:
   MARKET
   ECONOMIC
   TECHNOLOGY
   GEOPOLITICAL
   REJECT

2. assetClasses should identify affected investment asset classes.

3. geographies should use the supplied research geography master where appropriate.

4. sectors should use the supplied research sector master where appropriate.

5. relevant means materially useful for investment/fund monitoring.

6. REJECT should be used when the article is not meaningfully relevant
   to investment, economic, market, technology, or geopolitical monitoring.

7. investorImpact should explain the practical investment implication,
   not merely repeat the headline.

8. reasoning should briefly explain why the article received its
   classification.

9. fundMonitoringRelevant should indicate whether this article could
   reasonably matter when monitoring investment funds.

RESEARCH GEOGRAPHY MASTER:
{geography_payload}

RESEARCH SECTOR MASTER:
{sector_payload}

ARTICLES:
{article_payload}
""".strip()


# ============================================================
# GEMINI REQUEST
# ============================================================

def call_gemini(
    prompt: str,
    model: str,
    api_key: str,
    timeout: int,
    max_output_tokens: int,
) -> tuple[int, float, str]:

    endpoint = (
        "https://generativelanguage.googleapis.com"
        f"/v1beta/models/{model}:generateContent"
        f"?key={api_key}"
    )

    body = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": prompt
                    }
                ],
            }
        ],
        "generationConfig": {
            "temperature": 0,
            "maxOutputTokens": max_output_tokens,
            "responseMimeType": "application/json",
        },
    }

    payload = json.dumps(
        body,
        ensure_ascii=False,
    ).encode("utf-8")

    request = Request(
        endpoint,
        data=payload,
        headers={
            "Content-Type": "application/json",
        },
        method="POST",
    )

    started = time.monotonic()

    try:
        with urlopen(
            request,
            timeout=timeout,
        ) as response:

            response_body = response.read().decode(
                "utf-8",
                errors="replace",
            )

            elapsed = time.monotonic() - started

            return (
                response.status,
                elapsed,
                response_body,
            )

    except HTTPError as exc:

        elapsed = time.monotonic() - started

        try:
            response_body = exc.read().decode(
                "utf-8",
                errors="replace",
            )
        except Exception:
            response_body = str(exc)

        return (
            exc.code,
            elapsed,
            response_body,
        )

    except URLError as exc:

        elapsed = time.monotonic() - started

        return (
            0,
            elapsed,
            f"URL ERROR: {exc}",
        )

    except Exception as exc:

        elapsed = time.monotonic() - started

        return (
            0,
            elapsed,
            f"REQUEST ERROR: {type(exc).__name__}: {exc}",
        )


# ============================================================
# PARSE GEMINI RESPONSE
# ============================================================

def extract_text(response_data: dict[str, Any]) -> str:

    candidates = response_data.get(
        "candidates",
        [],
    )

    if not candidates:
        raise ValueError(
            "Gemini response contains no candidates."
        )

    candidate = candidates[0]

    content = candidate.get(
        "content",
        {},
    )

    parts = content.get(
        "parts",
        [],
    )

    texts = []

    for part in parts:
        if isinstance(part, dict):
            text_value = part.get("text")

            if text_value:
                texts.append(str(text_value))

    if not texts:
        raise ValueError(
            "Gemini response contains no text output."
        )

    return "".join(texts)


def parse_analysis(
    response_body: str,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:

    try:
        response_data = json.loads(response_body)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"Gemini HTTP response is not valid JSON: {exc}"
        )

    generated_text = extract_text(response_data)

    try:
        parsed = json.loads(generated_text)
    except json.JSONDecodeError as exc:
        raise ValueError(
            "Gemini generated text is not valid JSON: "
            f"{exc}\nGenerated text:\n{generated_text}"
        )

    if isinstance(parsed, list):
        analyses = parsed

    elif isinstance(parsed, dict):
        if isinstance(parsed.get("analyses"), list):
            analyses = parsed["analyses"]
        elif isinstance(parsed.get("results"), list):
            analyses = parsed["results"]
        else:
            analyses = [parsed]

    else:
        raise ValueError(
            "Gemini output must be a JSON array or object."
        )

    if not all(
        isinstance(item, dict)
        for item in analyses
    ):
        raise ValueError(
            "Every analysis result must be a JSON object."
        )

    return response_data, analyses


# ============================================================
# VALIDATION
# ============================================================

def validate_results(
    articles: list[dict[str, Any]],
    analyses: list[dict[str, Any]],
) -> None:

    expected_ids = [
        str(article["articleId"])
        for article in articles
    ]

    actual_ids = [
        str(item.get("articleId"))
        for item in analyses
    ]

    if len(analyses) != len(articles):
        raise ValueError(
            "Result count mismatch: "
            f"expected {len(articles)}, "
            f"received {len(analyses)}."
        )

    if len(set(actual_ids)) != len(actual_ids):
        raise ValueError(
            "Duplicate articleId detected in Gemini response."
        )

    if set(actual_ids) != set(expected_ids):
        missing = sorted(
            set(expected_ids) - set(actual_ids)
        )

        unexpected = sorted(
            set(actual_ids) - set(expected_ids)
        )

        raise ValueError(
            "Article ID validation failed.\n"
            f"Missing: {missing}\n"
            f"Unexpected: {unexpected}"
        )

    required_fields = {
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
    }

    for index, item in enumerate(analyses, start=1):

        missing_fields = sorted(
            required_fields - set(item.keys())
        )

        if missing_fields:
            raise ValueError(
                f"Article {index} is missing fields: "
                f"{missing_fields}"
            )

        if not isinstance(
            item["assetClasses"],
            list,
        ):
            raise ValueError(
                f"Article {index}: assetClasses must be a list."
            )

        if not isinstance(
            item["geographies"],
            list,
        ):
            raise ValueError(
                f"Article {index}: geographies must be a list."
            )

        if not isinstance(
            item["sectors"],
            list,
        ):
            raise ValueError(
                f"Article {index}: sectors must be a list."
            )

    log("Validation passed.")
    log(f"Expected article results: {len(expected_ids)}")
    log(f"Received article results: {len(actual_ids)}")


# ============================================================
# ATOMIC WRITE
# ============================================================

def atomic_write_json(
    path: Path,
    data: dict[str, Any],
) -> None:

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary = path.with_suffix(
        path.suffix + ".tmp"
    )

    with temporary.open(
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

    temporary.replace(path)


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    log("=" * 72)
    log("VGrat FMS - MARKET NEWS AI ANALYZER")
    log("=" * 72)

    api_key = os.environ.get(
        "GEMINI_API_KEY"
    )

    if not api_key:
        log("ERROR: GEMINI_API_KEY is not set.")
        return 1

    model = os.environ.get(
        "GEMINI_MODEL",
        DEFAULT_MODEL,
    )

    timeout = int(
        os.environ.get(
            "GEMINI_TIMEOUT_SECONDS",
            str(DEFAULT_TIMEOUT),
        )
    )

    max_output_tokens = int(
        os.environ.get(
            "GEMINI_MAX_OUTPUT_TOKENS",
            str(DEFAULT_MAX_OUTPUT_TOKENS),
        )
    )

    log(f"Model: {model}")
    log(f"Batch size: {BATCH_SIZE}")
    log(f"Maximum batches: {MAX_BATCHES}")
    log("Processing newest first.")
    log("No retries.")
    log("Production current.json: READ-ONLY")
    log("Research Funds.xlsx: READ-ONLY")

    # --------------------------------------------------------
    # Load inputs
    # --------------------------------------------------------

    articles = load_current_articles()

    research_context = load_research_context()

    existing_analysis = load_existing_analysis()

    completed = completed_article_ids(
        existing_analysis
    )

    log("=" * 72)
    log("SELECTING PENDING ARTICLES")
    log("=" * 72)

    pending = [
        article
        for article in articles
        if str(article["articleId"]) not in completed
    ]

    pending.sort(
        key=sort_key,
        reverse=PROCESS_NEWEST_FIRST,
    )

    log(
        f"Completed article IDs: {len(completed)}"
    )

    log(
        f"Pending articles: {len(pending)}"
    )

    if not pending:
        log("No pending articles.")
        return 0

    batch = pending[:BATCH_SIZE]

    log(
        f"Selected batch size: {len(batch)}"
    )

    for index, article in enumerate(
        batch,
        start=1,
    ):

        log(
            f"{index}. "
            f"{article['articleId']} | "
            f"{article.get('title')}"
        )

    if len(batch) != BATCH_SIZE:
        log(
            "WARNING: Fewer than 5 pending articles are available."
        )

    # --------------------------------------------------------
    # Build prompt
    # --------------------------------------------------------

    prompt = build_prompt(
        batch,
        research_context,
    )

    log("=" * 72)
    log("PROMPT PREPARED")
    log("=" * 72)

    log(
        f"Prompt characters: {len(prompt)}"
    )

    # --------------------------------------------------------
    # ONE Gemini request
    # --------------------------------------------------------

    log("=" * 72)
    log("SENDING ONE GEMINI REQUEST")
    log("=" * 72)

    request_started = now_utc()

    status, response_seconds, response_body = call_gemini(
        prompt=prompt,
        model=model,
        api_key=api_key,
        timeout=timeout,
        max_output_tokens=max_output_tokens,
    )

    request_finished = now_utc()

    log(
        f"Gemini HTTP status: {status}"
    )

    log(
        f"Gemini response time: "
        f"{response_seconds:.2f}s"
    )

    if status != 200:

        log("=" * 72)
        log("GEMINI REQUEST FAILED")
        log("=" * 72)

        log(response_body)

        log("NO FILES WERE WRITTEN.")

        return 1

    # --------------------------------------------------------
    # Parse
    # --------------------------------------------------------

    log("=" * 72)
    log("PARSING GEMINI RESPONSE")
    log("=" * 72)

    try:

        response_data, analyses = parse_analysis(
            response_body
        )

        log(
            f"Parsed analysis results: "
            f"{len(analyses)}"
        )

    except Exception as exc:

        log(
            f"ERROR parsing Gemini response: {exc}"
        )

        log("NO FILES WERE WRITTEN.")

        return 1

    # --------------------------------------------------------
    # Validate
    # --------------------------------------------------------

    log("=" * 72)
    log("VALIDATING BATCH")
    log("=" * 72)

    try:

        validate_results(
            batch,
            analyses,
        )

    except Exception as exc:

        log(
            f"VALIDATION FAILED: {exc}"
        )

        log("NO FILES WERE WRITTEN.")

        return 1

    # --------------------------------------------------------
    # Store successful batch
    # --------------------------------------------------------

    log("=" * 72)
    log("STORING SUCCESSFUL BATCH")
    log("=" * 72)

    existing_analysis = load_existing_analysis()

    existing_batches = existing_analysis.get(
        "batches",
        [],
    )

    next_batch_number = (
        len(existing_batches) + 1
    )

    batch_record = {
        "batchNumber": next_batch_number,
        "storedAt": now_utc(),
        "model": model,

        "articleIds": [
            article["articleId"]
            for article in batch
        ],

        "request": {
            "startedAt": request_started,
            "finishedAt": request_finished,
            "responseSeconds": round(
                response_seconds,
                3,
            ),
            "articleCount": len(batch),
            "promptCharacters": len(prompt),
            "maxOutputTokens": max_output_tokens,
            "responseMimeType": "application/json",
        },

        "validation": {
            "passed": True,
            "expectedArticleCount": len(batch),
            "receivedArticleCount": len(analyses),
        },

        "prompt": prompt,

        "requestBody": {
            "generationConfig": {
                "temperature": 0,
                "maxOutputTokens": max_output_tokens,
                "responseMimeType": "application/json",
            },
        },

        "parsedResponse": response_data,

        "analyses": analyses,

        "rawGeminiResponse": response_body,

        "inputArticles": batch,
    }

    existing_batches.append(
        batch_record
    )

    existing_analysis["batches"] = existing_batches

    existing_analysis[
        "totalSuccessfulBatches"
    ] = len(existing_batches)

    existing_analysis[
        "totalSuccessfulArticles"
    ] = sum(
        len(
            batch_item.get(
                "articleIds",
                [],
            )
        )
        for batch_item in existing_batches
        if isinstance(batch_item, dict)
    )

    existing_analysis[
        "updatedAt"
    ] = now_utc()

    atomic_write_json(
        ANALYSIS_CURRENT,
        existing_analysis,
    )

    log(
        f"Analysis file written: "
        f"{ANALYSIS_CURRENT}"
    )

    log(
        f"Successful batch: "
        f"{next_batch_number}"
    )

    log(
        f"Successful articles: "
        f"{len(analyses)}"
    )

    log("=" * 72)
    log("BATCH TEST SUCCESSFUL")
    log("=" * 72)

    return 0


if __name__ == "__main__":
    sys.exit(main())
