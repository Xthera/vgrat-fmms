#!/usr/bin/env python3

"""
VGrat FMS - Gemini Request Configuration Diagnostic
====================================================

Purpose
-------

Determine which part of the market-news Gemini request causes
HTTP 503 responses.

This diagnostic uses:

    data/market_news/current.json

as a READ-ONLY production input.

It does NOT read or modify:

    Research Funds.xlsx
    data/market_news/analysis/current.json

It does NOT write any files.

The diagnostic performs four independent Gemini requests:

    TEST 1
        Actual article
        Plain text response
        Small output limit

    TEST 2
        Actual article
        JSON MIME type
        Small output limit

    TEST 3
        Actual article
        JSON MIME type
        2048 output tokens

    TEST 4
        Full market-news analysis prompt
        JSON MIME type
        2048 output tokens

All tests use:

    gemini-3.8-flash

unless GEMINI_MODEL is explicitly supplied.

No retries are performed.

This is intentional: we want each request to expose the
actual API behavior rather than hiding a failure behind retries.
"""

from __future__ import annotations

import hashlib
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any


# ============================================================
# CONFIGURATION
# ============================================================

ROOT = Path(__file__).resolve().parent.parent

CURRENT_JSON = ROOT / "data" / "market_news" / "current.json"

API_BASE = (
    "https://generativelanguage.googleapis.com/v1beta"
)

API_KEY = os.getenv("GEMINI_API_KEY", "").strip()

GEMINI_MODEL = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.8-flash",
).strip()

TIMEOUT_SECONDS = int(
    os.getenv(
        "GEMINI_DIAGNOSTIC_TIMEOUT_SECONDS",
        "120",
    )
)


# ============================================================
# TIMING
# ============================================================

START_TIME = time.monotonic()


def elapsed() -> str:
    return f"{time.monotonic() - START_TIME:.2f}s"


def log(message: str = "") -> None:
    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    print(
        f"[{timestamp}] [+{elapsed()}] {message}",
        flush=True,
    )


def separator(
    char: str = "=",
    width: int = 72,
) -> None:

    print(
        char * width,
        flush=True,
    )


# ============================================================
# JSON HELPERS
# ============================================================

def pretty_json(value: Any) -> str:
    return json.dumps(
        value,
        indent=2,
        ensure_ascii=False,
    )


# ============================================================
# ARTICLE EXTRACTION
# ============================================================

def extract_articles(payload: Any) -> list[Any]:

    if isinstance(payload, list):
        return payload

    if not isinstance(payload, dict):
        return []

    envelope_keys = [
        "articles",
        "news",
        "items",
        "data",
        "results",
    ]

    for key in envelope_keys:

        value = payload.get(key)

        if isinstance(value, list):
            return value

        if isinstance(value, dict):

            nested = extract_articles(value)

            if nested:
                return nested

    return []


def first_nonempty(
    article: dict[str, Any],
    keys: list[str],
) -> Any:

    for key in keys:

        value = article.get(key)

        if value is not None:

            if isinstance(value, str):

                if value.strip():
                    return value.strip()

            else:
                return value

    return None


def make_article_id(
    article: dict[str, Any],
) -> str:

    existing = first_nonempty(
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

    url = first_nonempty(
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
        ).hexdigest()[:16]

    canonical = json.dumps(
        article,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )

    return hashlib.sha256(
        canonical.encode("utf-8")
    ).hexdigest()[:16]


def normalize_article(
    article: Any,
) -> dict[str, Any]:

    if not isinstance(article, dict):
        article = {
            "value": article
        }

    article_id = make_article_id(article)

    title = first_nonempty(
        article,
        [
            "title",
            "headline",
            "name",
        ],
    )

    description = first_nonempty(
        article,
        [
            "description",
            "summary",
            "snippet",
            "abstract",
        ],
    )

    url = first_nonempty(
        article,
        [
            "url",
            "link",
            "articleUrl",
        ],
    )

    published = first_nonempty(
        article,
        [
            "publishedAt",
            "published",
            "publishDate",
            "date",
            "datetime",
            "timestamp",
        ],
    )

    source = first_nonempty(
        article,
        [
            "source",
            "publisher",
            "provider",
        ],
    )

    return {
        "articleId": article_id,
        "title": title,
        "description": description,
        "url": url,
        "publishedAt": published,
        "source": source,
        "rawArticle": article,
    }


# ============================================================
# LOAD CURRENT NEWS
# ============================================================

def load_first_article() -> dict[str, Any]:

    separator()

    log("LOADING PRODUCTION CURRENT NEWS")

    separator()

    log(
        f"Input: {CURRENT_JSON}"
    )

    if not CURRENT_JSON.exists():

        raise FileNotFoundError(
            f"Missing input file: {CURRENT_JSON}"
        )

    with CURRENT_JSON.open(
        "r",
        encoding="utf-8",
    ) as handle:

        payload = json.load(handle)

    articles = extract_articles(payload)

    log(
        f"Articles discovered: {len(articles)}"
    )

    if not articles:

        raise RuntimeError(
            "No articles found in current.json."
        )

    article = normalize_article(
        articles[0]
    )

    log(
        f"Selected articleId: "
        f"{article['articleId']}"
    )

    log(
        f"Title: {article.get('title')}"
    )

    log(
        f"Published: {article.get('publishedAt')}"
    )

    log(
        f"Source: {article.get('source')}"
    )

    return article


# ============================================================
# PROMPT
# ============================================================

def build_full_prompt(
    article: dict[str, Any],
) -> str:

    article_json = json.dumps(
        {
            "articleId": article["articleId"],
            "title": article.get("title"),
            "description": article.get("description"),
            "url": article.get("url"),
            "publishedAt": article.get("publishedAt"),
            "source": article.get("source"),
        },
        ensure_ascii=False,
        indent=2,
    )

    return f"""
You are the market-news analysis engine for VGrat FMS.

Analyze the following single financial market news article.

Return exactly ONE JSON object.

Do not return markdown.
Do not return commentary outside the JSON object.

Required fields:

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

Rules:

category must be exactly one of:

MARKET
ECONOMIC
TECHNOLOGY
GEOPOLITICAL
REJECT

sentiment must be exactly one of:

positive
negative
neutral
mixed

importance must be exactly one of:

high
medium
low

relevant must be a boolean.

fundMonitoringRelevant must be a boolean.

assetClasses, geographies, and sectors must be arrays of strings.

investorImpact must be a concise explanation of how the
article could affect investors or financial markets.

reasoning must explain why the article received its
classification.

The articleId must exactly match the supplied articleId.

ARTICLE:

{article_json}
""".strip()


# ============================================================
# GEMINI REQUEST
# ============================================================

def call_gemini(
    test_name: str,
    prompt: str,
    response_mime_type: str | None,
    max_output_tokens: int,
) -> tuple[bool, int | None, float, str]:

    separator("-")

    log(test_name)

    separator("-")

    log(
        f"Model: {GEMINI_MODEL}"
    )

    log(
        f"Prompt characters: {len(prompt)}"
    )

    log(
        f"maxOutputTokens: {max_output_tokens}"
    )

    log(
        "responseMimeType: "
        f"{response_mime_type if response_mime_type else 'NOT SET'}"
    )

    request_body: dict[str, Any] = {
        "contents": [
            {
                "parts": [
                    {
                        "text": prompt
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0,
            "maxOutputTokens": max_output_tokens,
        },
    }

    if response_mime_type:
        request_body["generationConfig"][
            "responseMimeType"
        ] = response_mime_type

    encoded_model = urllib.parse.quote(
        GEMINI_MODEL,
        safe="",
    )

    url = (
        f"{API_BASE}/models/"
        f"{encoded_model}:generateContent"
        f"?key={urllib.parse.quote(API_KEY)}"
    )

    request_data = json.dumps(
        request_body
    ).encode("utf-8")

    request = urllib.request.Request(
        url,
        data=request_data,
        method="POST",
        headers={
            "Content-Type": "application/json",
        },
    )

    log("Sending request...")

    started = time.monotonic()

    try:

        with urllib.request.urlopen(
            request,
            timeout=TIMEOUT_SECONDS,
        ) as response:

            status = response.status

            raw_response = (
                response.read()
                .decode(
                    "utf-8",
                    errors="replace",
                )
            )

    except urllib.error.HTTPError as exc:

        duration = time.monotonic() - started

        try:
            raw_response = (
                exc.read()
                .decode(
                    "utf-8",
                    errors="replace",
                )
            )
        except Exception:
            raw_response = str(exc)

        log(
            f"HTTP status: {exc.code}"
        )

        log(
            f"Response time: {duration:.2f}s"
        )

        print(
            "Response:",
            flush=True,
        )

        try:
            parsed = json.loads(
                raw_response
            )

            print(
                pretty_json(parsed),
                flush=True,
            )

        except Exception:
            print(
                raw_response,
                flush=True,
            )

        return (
            False,
            exc.code,
            duration,
            raw_response,
        )

    except urllib.error.URLError as exc:

        duration = time.monotonic() - started

        log(
            f"NETWORK ERROR: {exc}"
        )

        log(
            f"Response time: {duration:.2f}s"
        )

        return (
            False,
            None,
            duration,
            str(exc),
        )

    except TimeoutError as exc:

        duration = time.monotonic() - started

        log(
            "REQUEST TIMEOUT"
        )

        log(
            f"Response time: {duration:.2f}s"
        )

        return (
            False,
            None,
            duration,
            str(exc),
        )

    duration = time.monotonic() - started

    log(
        f"HTTP status: {status}"
    )

    log(
        f"Response time: {duration:.2f}s"
    )

    print(
        "Response:",
        flush=True,
    )

    try:

        parsed = json.loads(
            raw_response
        )

        print(
            pretty_json(parsed),
            flush=True,
        )

    except Exception:

        print(
            raw_response,
            flush=True,
        )

    return (
        status == 200,
        status,
        duration,
        raw_response,
    )


# ============================================================
# TEST RUNNER
# ============================================================

def run_test(
    test_number: int,
    name: str,
    prompt: str,
    response_mime_type: str | None,
    max_output_tokens: int,
) -> dict[str, Any]:

    test_name = (
        f"TEST {test_number} - {name}"
    )

    success, status, duration, raw = call_gemini(
        test_name=test_name,
        prompt=prompt,
        response_mime_type=response_mime_type,
        max_output_tokens=max_output_tokens,
    )

    if success:

        log(
            f"{test_name}: SUCCESS"
        )

    else:

        log(
            f"{test_name}: FAILED"
        )

        if status is not None:

            log(
                f"Failure HTTP status: {status}"
            )

    return {
        "test": test_number,
        "name": name,
        "success": success,
        "status": status,
        "durationSeconds": round(
            duration,
            2,
        ),
    }


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    separator()

    log(
        "VGrat FMS - GEMINI REQUEST DIAGNOSTIC"
    )

    separator()

    log(
        f"Model: {GEMINI_MODEL}"
    )

    log(
        "No retries."
    )

    log(
        "No files will be modified."
    )

    log(
        f"Production input: {CURRENT_JSON}"
    )

    print("", flush=True)

    if not API_KEY:

        log(
            "ERROR: GEMINI_API_KEY is not set."
        )

        return 1

    # --------------------------------------------------------
    # Load actual production article.
    # --------------------------------------------------------

    try:

        article = load_first_article()

    except Exception as exc:

        log(
            f"ERROR loading current.json: {exc}"
        )

        return 1

    # --------------------------------------------------------
    # Build prompts.
    # --------------------------------------------------------

    title = article.get("title")

    description = article.get(
        "description"
    )

    article_text = (
        f"Title: {title}\n\n"
        f"Description: {description}"
    )

    full_prompt = build_full_prompt(
        article
    )

    log(
        f"Full market-news prompt characters: "
        f"{len(full_prompt)}"
    )

    print("", flush=True)

    results: list[dict[str, Any]] = []

    # ========================================================
    # TEST 1
    # ========================================================

    results.append(
        run_test(
            test_number=1,
            name=(
                "Actual article + plain text + "
                "small output"
            ),
            prompt=(
                "Read this financial news article "
                "and reply with one short sentence "
                "describing its main topic.\n\n"
                + article_text
            ),
            response_mime_type=None,
            max_output_tokens=32,
        )
    )

    # ========================================================
    # TEST 2
    # ========================================================

    results.append(
        run_test(
            test_number=2,
            name=(
                "Actual article + JSON MIME + "
                "small output"
            ),
            prompt=(
                "Return a JSON object with exactly "
                "one field called \"topic\" containing "
                "a short description of the main topic "
                "of this article.\n\n"
                + article_text
            ),
            response_mime_type="application/json",
            max_output_tokens=32,
        )
    )

    # ========================================================
    # TEST 3
    # ========================================================

    results.append(
        run_test(
            test_number=3,
            name=(
                "Actual article + JSON MIME + "
                "2048 output tokens"
            ),
            prompt=(
                "Return a JSON object with exactly "
                "one field called \"topic\" containing "
                "a short description of the main topic "
                "of this article.\n\n"
                + article_text
            ),
            response_mime_type="application/json",
            max_output_tokens=2048,
        )
    )

    # ========================================================
    # TEST 4
    # ========================================================

    results.append(
        run_test(
            test_number=4,
            name=(
                "FULL market-news analysis request"
            ),
            prompt=full_prompt,
            response_mime_type="application/json",
            max_output_tokens=2048,
        )
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    separator()

    log(
        "DIAGNOSTIC SUMMARY"
    )

    separator()

    print("", flush=True)

    for result in results:

        status = result["status"]

        status_text = (
            str(status)
            if status is not None
            else "N/A"
        )

        result_text = (
            "SUCCESS"
            if result["success"]
            else "FAILED"
        )

        print(
            f"TEST {result['test']}: "
            f"{result['name']}",
            flush=True,
        )

        print(
            f"    Result: {result_text}",
            flush=True,
        )

        print(
            f"    HTTP: {status_text}",
            flush=True,
        )

        print(
            f"    Time: "
            f"{result['durationSeconds']:.2f}s",
            flush=True,
        )

        print("", flush=True)

    # --------------------------------------------------------
    # Find first failure.
    # --------------------------------------------------------

    first_failure = next(
        (
            result
            for result in results
            if not result["success"]
        ),
        None,
    )

    separator()

    if first_failure is None:

        log(
            "DIAGNOSTIC RESULT: ALL TESTS PASSED"
        )

        separator()

        log(
            "Gemini successfully handled the "
            "actual market-news request configuration."
        )

        log(
            "The earlier HTTP 503 was therefore "
            "most likely transient."
        )

        log(
            "No files were modified."
        )

        return 0

    failed_number = first_failure["test"]

    failed_status = first_failure["status"]

    log(
        f"FIRST FAILURE: TEST {failed_number}"
    )

    log(
        f"HTTP status: {failed_status}"
    )

    # --------------------------------------------------------
    # Interpret failure boundary.
    # --------------------------------------------------------

    if failed_number == 1:

        log(
            "The actual article itself triggers "
            "a generation failure."
        )

        log(
            "This is NOT related to JSON MIME type "
            "or the larger output limit."
        )

    elif failed_number == 2:

        log(
            "The failure begins when "
            "responseMimeType=application/json "
            "is added."
        )

        log(
            "The article can generate plain text."
        )

    elif failed_number == 3:

        log(
            "The failure begins when maxOutputTokens "
            "is increased from 32 to 2048."
        )

        log(
            "The JSON response mode itself works."
        )

    elif failed_number == 4:

        log(
            "The first three request configurations "
            "work."
        )

        log(
            "The failure occurs only with the "
            "full market-news analysis prompt."
        )

        log(
            "This points toward the full prompt/request "
            "configuration rather than basic Gemini access."
        )

    separator()

    log(
        "No files were modified."
    )

    # --------------------------------------------------------
    # Deliberately fail workflow if any test fails.
    # --------------------------------------------------------

    return 1


if __name__ == "__main__":
    sys.exit(main())
