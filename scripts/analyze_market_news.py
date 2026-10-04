#!/usr/bin/env python3

"""
VGrat FMS - GEMINI ONE-ARTICLE DIAGNOSTIC
=========================================

PURPOSE
=======

This is a temporary diagnostic script.

It selects exactly ONE newest article from:

    data/market_news/current.json

and sends exactly ONE Gemini request.

It does NOT:

- modify current.json
- modify Research Funds.xlsx
- create analysis/current.json
- create analysis history
- commit anything
- process multiple articles
- retry the request

The purpose is to determine whether the Gemini API/model itself
can successfully process a minimal request.

SUCCESS
=======

If Gemini succeeds, the script prints:

- HTTP status
- response time
- model
- article ID
- parsed JSON response

FAILURE
=======

If Gemini returns 503, 429, 404, 400, etc., the script prints
the exact API response and exits with code 1.

No files are written by this script.
"""

from __future__ import annotations

import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests


# ============================================================================
# PATHS
# ============================================================================

ROOT = Path(__file__).resolve().parent.parent

CURRENT_JSON = (
    ROOT
    / "data"
    / "market_news"
    / "current.json"
)


# ============================================================================
# CONFIGURATION
# ============================================================================

# Keep the current model for the diagnostic.
GEMINI_MODEL = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.8-flash",
).strip()

GEMINI_API_KEY = os.getenv(
    "GEMINI_API_KEY",
    "",
).strip()

GEMINI_TIMEOUT_SECONDS = int(
    os.getenv(
        "GEMINI_TIMEOUT_SECONDS",
        "120",
    )
)


# ============================================================================
# LOGGING
# ============================================================================

SCRIPT_START = time.monotonic()


def log(message: str = "") -> None:
    elapsed = (
        time.monotonic()
        - SCRIPT_START
    )

    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    print(
        f"[{timestamp}] [+{elapsed:.2f}s] {message}",
        flush=True,
    )


# ============================================================================
# HELPERS
# ============================================================================

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


def extract_articles(
    payload: Any,
) -> list[dict[str, Any]]:

    if isinstance(payload, list):
        return [
            item
            for item in payload
            if isinstance(item, dict)
        ]

    if not isinstance(payload, dict):
        raise ValueError(
            "current.json must contain an object or array."
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
        "Could not find an article list in current.json."
    )


def make_article_id(
    article: dict[str, Any],
) -> str:

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
        import hashlib

        return hashlib.sha256(
            str(url).encode("utf-8")
        ).hexdigest()

    canonical = json.dumps(
        article,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )

    import hashlib

    return hashlib.sha256(
        canonical.encode("utf-8")
    ).hexdigest()


def normalise_article(
    article: dict[str, Any],
) -> dict[str, Any]:

    return {
        "articleId": make_article_id(article),

        "title": str(
            first_nonempty(
                article,
                [
                    "title",
                    "headline",
                    "name",
                ],
                "",
            )
        ),

        "summary": str(
            first_nonempty(
                article,
                [
                    "summary",
                    "description",
                    "dek",
                    "excerpt",
                ],
                "",
            )
        ),

        "content": str(
            first_nonempty(
                article,
                [
                    "content",
                    "body",
                    "text",
                    "article",
                ],
                "",
            )
        ),

        "url": str(
            first_nonempty(
                article,
                [
                    "url",
                    "link",
                    "articleUrl",
                    "article_url",
                ],
                "",
            )
        ),

        "publishedAt": str(
            first_nonempty(
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
        ),

        "source": str(
            first_nonempty(
                article,
                [
                    "source",
                    "publisher",
                    "provider",
                ],
                "",
            )
        ),
    }


# ============================================================================
# GEMINI SCHEMA
# ============================================================================

def build_response_schema() -> dict[str, Any]:

    return {
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
            "reasoning": {
                "type": "STRING",
            },
        },
        "required": [
            "articleId",
            "relevant",
            "category",
            "sentiment",
            "importance",
            "reasoning",
        ],
    }


# ============================================================================
# PROMPT
# ============================================================================

def build_prompt(
    article: dict[str, Any],
) -> str:

    article_json = json.dumps(
        article,
        ensure_ascii=False,
        indent=2,
    )

    return f"""
You are a financial market-news classifier.

Analyse the ONE article below.

Return exactly ONE JSON object.

Do not return markdown.
Do not return explanations outside JSON.
Do not invent information.

Classification:

MARKET:
Financial markets, securities, companies, stocks, bonds,
commodities, currencies, interest rates, valuations or
investment activity.

ECONOMIC:
Inflation, employment, GDP, central banks, monetary policy,
fiscal policy, economic indicators, consumer conditions,
trade or other macroeconomic developments.

TECHNOLOGY:
Technology companies, AI, semiconductors, data centres,
cloud, software, cybersecurity or technology developments
with meaningful investment relevance.

GEOPOLITICAL:
Wars, conflicts, sanctions, elections, international
relations, tariffs, trade restrictions or geopolitical
developments affecting markets or economies.

REJECT:
Not materially relevant to financial markets, investment,
economics, technology investment or geopolitics.

Sentiment must be:
positive
negative
neutral
mixed

Importance must be:
high
medium
low

Article:

{article_json}
""".strip()


# ============================================================================
# GEMINI REQUEST
# ============================================================================

def send_gemini_request(
    prompt: str,
) -> tuple[dict[str, Any], float]:

    if not GEMINI_API_KEY:
        raise RuntimeError(
            "GEMINI_API_KEY is not set."
        )

    endpoint = (
        "https://generativelanguage.googleapis.com/"
        f"v1beta/models/{GEMINI_MODEL}:generateContent"
    )

    params = {
        "key": GEMINI_API_KEY,
    }

    request_body = {
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
            "maxOutputTokens": 2048,
            "responseMimeType": "application/json",
            "responseSchema": build_response_schema(),
        },
    }

    log(
        "Sending ONE Gemini request..."
    )

    log(
        f"Endpoint model: {GEMINI_MODEL}"
    )

    started = time.monotonic()

    try:

        response = requests.post(
            endpoint,
            params=params,
            json=request_body,
            timeout=GEMINI_TIMEOUT_SECONDS,
        )

    except requests.RequestException as exc:

        elapsed = (
            time.monotonic()
            - started
        )

        log(
            f"Network error after {elapsed:.2f}s."
        )

        raise RuntimeError(
            f"Gemini network request failed: {exc}"
        ) from exc

    elapsed = (
        time.monotonic()
        - started
    )

    log(
        f"Gemini HTTP status: "
        f"{response.status_code}"
    )

    log(
        f"Gemini response time: "
        f"{elapsed:.2f}s"
    )

    if not response.ok:

        log()
        log(
            "GEMINI REQUEST FAILED"
        )

        log(
            f"HTTP {response.status_code}"
        )

        log(
            "Response:"
        )

        log(response.text)

        raise RuntimeError(
            f"Gemini API returned HTTP "
            f"{response.status_code}."
        )

    try:
        response_json = response.json()

    except ValueError as exc:

        log(
            "Gemini returned non-JSON HTTP response."
        )

        log(response.text)

        raise RuntimeError(
            "Gemini returned invalid HTTP JSON."
        ) from exc

    return response_json, elapsed


# ============================================================================
# RESPONSE EXTRACTION
# ============================================================================

def extract_text(
    response_json: dict[str, Any],
) -> str:

    candidates = response_json.get(
        "candidates"
    )

    if not isinstance(
        candidates,
        list,
    ) or not candidates:

        raise ValueError(
            "Gemini response has no candidates."
        )

    candidate = candidates[0]

    if not isinstance(
        candidate,
        dict,
    ):

        raise ValueError(
            "Gemini candidate is not an object."
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

    text_parts = []

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

        raise ValueError(
            "Gemini returned no text."
        )

    return "".join(
        text_parts
    ).strip()


# ============================================================================
# MAIN
# ============================================================================

def main() -> None:

    log("=" * 72)
    log(
        "VGrat FMS - GEMINI ONE-ARTICLE DIAGNOSTIC"
    )
    log("=" * 72)

    log(
        "TEST: exactly ONE article."
    )

    log(
        f"Gemini model: {GEMINI_MODEL}"
    )

    log(
        f"Input: {CURRENT_JSON}"
    )

    log(
        "Output: NONE - diagnostic only."
    )

    log()

    # ========================================================================
    # CHECK API KEY
    # ========================================================================

    if not GEMINI_API_KEY:

        log(
            "ERROR: GEMINI_API_KEY is not configured."
        )

        sys.exit(1)

    # ========================================================================
    # LOAD CURRENT.JSON
    # ========================================================================

    if not CURRENT_JSON.exists():

        log(
            f"ERROR: current.json not found: "
            f"{CURRENT_JSON}"
        )

        sys.exit(1)

    log(
        "Loading production current.json..."
    )

    try:

        with CURRENT_JSON.open(
            "r",
            encoding="utf-8",
        ) as f:

            payload = json.load(f)

    except Exception as exc:

        log(
            f"ERROR loading current.json: {exc}"
        )

        sys.exit(1)

    # ========================================================================
    # EXTRACT ARTICLES
    # ========================================================================

    try:

        raw_articles = extract_articles(
            payload
        )

    except Exception as exc:

        log(
            f"ERROR extracting articles: {exc}"
        )

        sys.exit(1)

    log(
        f"Articles found: {len(raw_articles)}"
    )

    if not raw_articles:

        log(
            "ERROR: No articles found."
        )

        sys.exit(1)

    # ========================================================================
    # NORMALISE ARTICLES
    # ========================================================================

    articles = [
        normalise_article(article)
        for article in raw_articles
    ]

    # ========================================================================
    # NEWEST FIRST
    #
    # current.json is already produced by the collector in current order.
    # For this diagnostic, use the FIRST article exactly as it appears.
    #
    # This avoids introducing any additional date parsing variable into
    # the Gemini availability test.
    # ========================================================================

    article = articles[0]

    # ========================================================================
    # DISPLAY ARTICLE
    # ========================================================================

    log()

    log(
        "Selected ONE article:"
    )

    log(
        f"Article ID: {article['articleId']}"
    )

    log(
        f"Published: {article['publishedAt']}"
    )

    log(
        f"Source: {article['source']}"
    )

    log(
        f"Title: {article['title']}"
    )

    log(
        f"URL: {article['url']}"
    )

    log()

    # ========================================================================
    # BUILD PROMPT
    # ========================================================================

    prompt = build_prompt(
        article
    )

    log(
        "Prompt prepared."
    )

    log(
        f"Prompt characters: {len(prompt)}"
    )

    log()

    # ========================================================================
    # SEND EXACTLY ONE REQUEST
    # ========================================================================

    try:

        response_json, elapsed = (
            send_gemini_request(
                prompt
            )
        )

    except Exception as exc:

        log()
        log(
            "DIAGNOSTIC RESULT: FAILED"
        )

        log(
            str(exc)
        )

        log()
        log(
            "No files were modified."
        )

        sys.exit(1)

    # ========================================================================
    # EXTRACT TEXT
    # ========================================================================

    log()
    log(
        "Gemini request succeeded."
    )

    log(
        "Extracting response..."
    )

    try:

        text_output = extract_text(
            response_json
        )

    except Exception as exc:

        log()
        log(
            "DIAGNOSTIC RESULT: HTTP SUCCESS "
            "BUT RESPONSE PARSING FAILED"
        )

        log(
            str(exc)
        )

        log()
        log(
            "Raw Gemini response:"
        )

        print(
            json.dumps(
                response_json,
                ensure_ascii=False,
                indent=2,
            ),
            flush=True,
        )

        sys.exit(1)

    # ========================================================================
    # PARSE JSON
    # ========================================================================

    log(
        "Parsing structured JSON..."
    )

    try:

        parsed = json.loads(
            text_output
        )

    except json.JSONDecodeError as exc:

        log()
        log(
            "DIAGNOSTIC RESULT: GEMINI WORKED "
            "BUT RETURNED INVALID JSON"
        )

        log(
            str(exc)
        )

        log()
        log(
            "Raw model output:"
        )

        print(
            text_output,
            flush=True,
        )

        sys.exit(1)

    # ========================================================================
    # VALIDATE ARTICLE ID
    # ========================================================================

    if not isinstance(
        parsed,
        dict,
    ):

        log(
            "DIAGNOSTIC RESULT: INVALID JSON TYPE"
        )

        log(
            "Expected one JSON object."
        )

        log(
            f"Received: {type(parsed).__name__}"
        )

        sys.exit(1)

    returned_article_id = str(
        parsed.get(
            "articleId",
            ""
        )
    )

    if returned_article_id != article[
        "articleId"
    ]:

        log()
        log(
            "DIAGNOSTIC RESULT: ARTICLE ID MISMATCH"
        )

        log(
            f"Expected: {article['articleId']}"
        )

        log(
            f"Received: {returned_article_id}"
        )

        log()
        log(
            "Model response:"
        )

        print(
            json.dumps(
                parsed,
                ensure_ascii=False,
                indent=2,
            ),
            flush=True,
        )

        sys.exit(1)

    # ========================================================================
    # SUCCESS
    # ========================================================================

    log()
    log("=" * 72)
    log(
        "DIAGNOSTIC RESULT: SUCCESS"
    )
    log("=" * 72)

    log(
        f"Model: {GEMINI_MODEL}"
    )

    log(
        f"HTTP request time: {elapsed:.2f}s"
    )

    log(
        f"Article ID: {article['articleId']}"
    )

    log(
        "Gemini returned valid structured JSON."
    )

    log()
    log(
        "MODEL RESPONSE:"
    )

    print(
        json.dumps(
            parsed,
            ensure_ascii=False,
            indent=2,
        ),
        flush=True,
    )

    log()
    log(
        "No files were modified."
    )

    log(
        "ONE-ARTICLE DIAGNOSTIC COMPLETED."
    )


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    main()
