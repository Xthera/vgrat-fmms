#!/usr/bin/env python3

"""
VGrat FMS - MARKET NEWS AI ANALYZER
===================================

FIRST-STAGE ANALYSIS ONLY

Purpose:
    Analyze the 2 newest articles from:

        data/market_news/current.json

    using Gemini.

Output:

        data/market_news/analysis/current.json

Rules:
    - No Research Funds.xlsx
    - No geography/sector master matching
    - No correction mode
    - No existing-analysis/history logic
    - No automatic batch continuation
    - Only the newest 2 articles are analyzed
    - Output is written ONLY after successful validation
    - Failed requests write nothing
    - Successful prompts and Gemini responses are stored
    - Summary must contain no more than 9 sentences

Environment:
    GEMINI_API_KEY

Python:
    3.11+

Dependencies:
    None beyond Python standard library.
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
import hashlib
from datetime import datetime, timezone
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


# ============================================================
# CONFIGURATION
# ============================================================

ROOT = Path(__file__).resolve().parent.parent

CURRENT_JSON = ROOT / "data" / "market_news" / "current.json"

ANALYSIS_DIR = ROOT / "data" / "market_news" / "analysis"
ANALYSIS_CURRENT = ANALYSIS_DIR / "current.json"

DEFAULT_MODEL = "gemini-3.8-flash"

MAX_ARTICLES = 2

REQUEST_TIMEOUT = 180

MAX_OUTPUT_TOKENS = 8192

# Keep this at 1 for the first-stage test.
MAX_BATCHES = 1


# ============================================================
# LOGGING
# ============================================================

def log(message: str = "") -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    print(f"[{now}] {message}", flush=True)


# ============================================================
# JSON HELPERS
# ============================================================

def load_json(path: Path) -> object:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_json_atomic(path: Path, data: object) -> None:
    """
    Write atomically so a failed write cannot leave a partially
    written current.json.
    """

    path.parent.mkdir(parents=True, exist_ok=True)

    temp_path = path.with_suffix(path.suffix + ".tmp")

    with temp_path.open("w", encoding="utf-8") as f:
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

def get_article_list(data: object) -> list[dict]:
    """
    Accept common market-news envelope layouts.
    """

    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]

    if not isinstance(data, dict):
        raise ValueError("current.json must contain an object or array.")

    possible_keys = [
        "articles",
        "news",
        "items",
        "data",
        "results",
    ]

    for key in possible_keys:
        value = data.get(key)

        if isinstance(value, list):
            return [x for x in value if isinstance(x, dict)]

        if isinstance(value, dict):
            for nested_key in possible_keys:
                nested = value.get(nested_key)
                if isinstance(nested, list):
                    return [
                        x for x in nested
                        if isinstance(x, dict)
                    ]

    raise ValueError(
        "Could not find an article list in current.json. "
        "Expected one of: articles, news, items, data, results."
    )


# ============================================================
# ARTICLE ID
# ============================================================

def article_id(article: dict) -> str:
    """
    Prefer an existing source ID.

    If unavailable, generate a deterministic SHA-256 ID.
    """

    for key in [
        "id",
        "articleId",
        "article_id",
        "newsId",
        "news_id",
    ]:
        value = article.get(key)

        if value is not None and str(value).strip():
            return str(value).strip()

    url = article.get("url") or article.get("link")

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
# DATE SORTING
# ============================================================

def article_sort_timestamp(article: dict) -> float:
    """
    Try to determine article publication time.

    Articles without a usable date are placed after dated
    articles while preserving their original order.
    """

    date_keys = [
        "publishedAt",
        "published",
        "publishDate",
        "publishedDate",
        "date",
        "datetime",
        "timestamp",
        "createdAt",
    ]

    for key in date_keys:
        value = article.get(key)

        if value is None:
            continue

        text_value = str(value).strip()

        if not text_value:
            continue

        # Unix timestamp.
        try:
            number = float(text_value)

            # Handle milliseconds.
            if number > 10_000_000_000:
                number /= 1000

            return number
        except ValueError:
            pass

        # ISO datetime.
        try:
            normalized = text_value.replace("Z", "+00:00")

            dt = datetime.fromisoformat(normalized)

            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)

            return dt.timestamp()

        except ValueError:
            pass

    return float("-inf")


# ============================================================
# ARTICLE NORMALIZATION
# ============================================================

def prepare_article(article: dict) -> dict:
    """
    Keep the source article intact while ensuring articleId is
    available to Gemini.
    """

    result = dict(article)

    result["articleId"] = article_id(article)

    return result


# ============================================================
# GEMINI PROMPT
# ============================================================

def build_prompt(articles: list[dict]) -> str:
    article_json = json.dumps(
        articles,
        ensure_ascii=False,
        indent=2,
    )

    return f"""
You are the market-news analysis engine for VGrat FMS.

Analyze the supplied CNBC market-news articles.

Your task is ONLY to analyze the articles.
Do not perform external research.
Do not invent information that is not supported by the article.

Return EXACTLY ONE analysis object for EACH supplied article.

The number of analysis objects MUST equal the number of supplied articles.

Required output format:

{{
  "analyses": [
    {{
      "articleId": "string",
      "relevant": true,
      "category": "MARKET",
      "sentiment": "positive",
      "importance": "medium",
      "summary": "string",
      "assetClasses": ["Equities"],
      "geographies": ["United States"],
      "sectors": ["Technology"],
      "investorImpact": "string",
      "reasoning": "string",
      "fundMonitoringRelevant": true
    }}
  ]
}}

FIELD RULES
===========

articleId
---------
Must exactly match the articleId supplied in the input.

relevant
--------
Boolean.

Set true when the article has meaningful relevance to financial
markets, economies, investments, companies, sectors, asset prices,
interest rates, currencies, commodities, policy or other matters
that could reasonably matter to investors.

category
--------
Use exactly one of:

MARKET
ECONOMIC
TECHNOLOGY
GEOPOLITICAL
REJECT

sentiment
---------
Use exactly one of:

positive
negative
mixed
neutral

importance
----------
Use exactly one of:

high
medium
low

summary
-------
Provide a concise factual summary of the article.

Maximum 9 sentences.

There is NO minimum sentence requirement.

Do not add information that is not supported by the article.

assetClasses
------------
Use an array.

Identify the financial asset classes meaningfully affected by the
article.

Examples include:

Equities
Fixed Income
Commodities
Currencies
Real Estate
Infrastructure
Alternatives
Cash

geographies
-----------
Use an array.

Identify the geographic markets or regions directly relevant to
the article.

sectors
-------
Use an array.

Identify the industries or economic sectors directly relevant to
the article.

investorImpact
--------------
Explain why the article matters to investors.

Be specific to the article.

Discuss relevant implications such as:

- market conditions
- earnings
- demand
- valuations
- costs
- regulation
- competition
- capital expenditure
- supply chains
- interest rates
- institutional positioning
- investment flows
- economic growth
- risk

Do not simply repeat the summary.

reasoning
---------
Briefly explain why the article received its relevance, category,
sentiment and importance classifications.

fundMonitoringRelevant
----------------------
Boolean.

Set true if the article is sufficiently relevant that an investor
or fund-monitoring system should retain it for monitoring.

Otherwise false.

IMPORTANT
=========

1. Return valid JSON only.
2. Do not use Markdown.
3. Do not wrap the JSON in ``` fences.
4. Return exactly one analysis per article.
5. Do not omit any required field.
6. Do not create additional top-level fields.
7. Do not change articleId values.
8. Summary must be 9 sentences or fewer.
9. Do not invent facts.

ARTICLES TO ANALYZE
===================

{article_json}
""".strip()


# ============================================================
# GEMINI API
# ============================================================

def call_gemini(
    api_key: str,
    model: str,
    prompt: str,
) -> tuple[dict, dict]:
    """
    Call Gemini generateContent.

    Returns:
        parsed_response, request_metadata

    Raises:
        RuntimeError on any API failure.
    """

    url = (
        "https://generativelanguage.googleapis.com/v1beta/"
        f"models/{model}:generateContent?key={api_key}"
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
            "temperature": 0.1,
            "maxOutputTokens": MAX_OUTPUT_TOKENS,
            "responseMimeType": "application/json",
        },
    }

    encoded_body = json.dumps(
        body,
        ensure_ascii=False,
    ).encode("utf-8")

    request = Request(
        url,
        data=encoded_body,
        headers={
            "Content-Type": "application/json",
        },
        method="POST",
    )

    started = time.monotonic()

    try:
        with urlopen(
            request,
            timeout=REQUEST_TIMEOUT,
        ) as response:

            raw_response = response.read().decode(
                "utf-8",
                errors="replace",
            )

            elapsed = time.monotonic() - started

            try:
                response_json = json.loads(raw_response)
            except json.JSONDecodeError as exc:
                raise RuntimeError(
                    "Gemini returned invalid JSON at API level: "
                    f"{exc}"
                ) from exc

            candidates = response_json.get("candidates", [])

            if not candidates:
                raise RuntimeError(
                    "Gemini response contains no candidates."
                )

            candidate = candidates[0]

            parts = (
                candidate
                .get("content", {})
                .get("parts", [])
            )

            text_parts = []

            for part in parts:
                if isinstance(part, dict):
                    text_value = part.get("text")

                    if text_value:
                        text_parts.append(str(text_value))

            generated_text = "".join(text_parts).strip()

            if not generated_text:
                raise RuntimeError(
                    "Gemini returned an empty generated response."
                )

            # Gemini can occasionally return JSON fenced in Markdown
            # despite responseMimeType=json.
            generated_text = clean_json_text(
                generated_text
            )

            try:
                parsed = json.loads(generated_text)
            except json.JSONDecodeError as exc:
                log("Gemini generated response:")
                log(generated_text)

                raise RuntimeError(
                    f"Gemini generated invalid JSON: {exc}"
                ) from exc

            usage = response_json.get(
                "usageMetadata",
                {},
            )

            metadata = {
                "model": model,
                "responseSeconds": round(elapsed, 3),
                "promptCharacters": len(prompt),
                "maxOutputTokens": MAX_OUTPUT_TOKENS,
                "responseMimeType": "application/json",
                "usage": usage,
                "rawResponse": raw_response,
                "generatedText": generated_text,
                "requestBody": body,
            }

            return parsed, metadata

    except HTTPError as exc:
        elapsed = time.monotonic() - started

        try:
            error_body = exc.read().decode(
                "utf-8",
                errors="replace",
            )
        except Exception:
            error_body = ""

        log(
            f"Gemini HTTP error: {exc.code} "
            f"after {elapsed:.2f}s"
        )

        if error_body:
            log(f"Gemini error body: {error_body}")

        raise RuntimeError(
            f"Gemini API HTTP {exc.code}: {error_body}"
        ) from exc

    except URLError as exc:
        elapsed = time.monotonic() - started

        raise RuntimeError(
            f"Gemini network error after "
            f"{elapsed:.2f}s: {exc}"
        ) from exc

    except TimeoutError as exc:
        elapsed = time.monotonic() - started

        raise RuntimeError(
            f"Gemini request timed out after "
            f"{elapsed:.2f}s."
        ) from exc


# ============================================================
# CLEAN GEMINI JSON
# ============================================================

def clean_json_text(text: str) -> str:
    """
    Remove accidental Markdown fences.
    """

    value = text.strip()

    if value.startswith("```"):
        value = re.sub(
            r"^```(?:json)?\s*",
            "",
            value,
            flags=re.IGNORECASE,
        )

        value = re.sub(
            r"\s*```$",
            "",
            value,
        )

    return value.strip()


# ============================================================
# VALIDATION
# ============================================================

VALID_CATEGORIES = {
    "MARKET",
    "ECONOMIC",
    "TECHNOLOGY",
    "GEOPOLITICAL",
    "REJECT",
}

VALID_SENTIMENTS = {
    "positive",
    "negative",
    "mixed",
    "neutral",
}

VALID_IMPORTANCE = {
    "high",
    "medium",
    "low",
}


def count_sentences(text: str) -> int:
    """
    Simple sentence counter suitable for enforcing the
    maximum-summary rule.

    Counts sentence-ending punctuation followed by whitespace
    or end-of-string.
    """

    text = text.strip()

    if not text:
        return 0

    matches = re.findall(
        r"[.!?]+(?:\s+|$)",
        text,
    )

    return len(matches)


def validate_analysis(
    response: object,
    articles: list[dict],
) -> tuple[bool, list[str]]:
    """
    Validate the entire Gemini response before writing anything.
    """

    errors: list[str] = []

    if not isinstance(response, dict):
        return False, [
            "Gemini response must be a JSON object."
        ]

    allowed_top_level = {"analyses"}

    unexpected = set(response.keys()) - allowed_top_level

    if unexpected:
        errors.append(
            "Unexpected top-level fields: "
            + ", ".join(sorted(unexpected))
        )

    analyses = response.get("analyses")

    if not isinstance(analyses, list):
        errors.append(
            "'analyses' must be an array."
        )
        return False, errors

    expected_ids = [
        article_id(article)
        for article in articles
    ]

    actual_ids = []

    for index, analysis in enumerate(analyses):

        prefix = f"analysis[{index}]"

        if not isinstance(analysis, dict):
            errors.append(
                f"{prefix} must be an object."
            )
            continue

        required_fields = {
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

        missing = required_fields - set(analysis.keys())

        if missing:
            errors.append(
                f"{prefix} missing fields: "
                + ", ".join(sorted(missing))
            )

        article_value = analysis.get("articleId")

        if article_value is not None:
            actual_ids.append(str(article_value))

        if not isinstance(
            analysis.get("relevant"),
            bool,
        ):
            errors.append(
                f"{prefix}.relevant must be boolean."
            )

        if analysis.get("category") not in VALID_CATEGORIES:
            errors.append(
                f"{prefix}.category has invalid value: "
                f"{analysis.get('category')}"
            )

        if analysis.get("sentiment") not in VALID_SENTIMENTS:
            errors.append(
                f"{prefix}.sentiment has invalid value: "
                f"{analysis.get('sentiment')}"
            )

        if analysis.get("importance") not in VALID_IMPORTANCE:
            errors.append(
                f"{prefix}.importance has invalid value: "
                f"{analysis.get('importance')}"
            )

        summary = analysis.get("summary")

        if not isinstance(summary, str):
            errors.append(
                f"{prefix}.summary must be a string."
            )
        else:
            sentence_count = count_sentences(summary)

            if sentence_count > 9:
                errors.append(
                    f"{prefix}.summary contains "
                    f"{sentence_count} sentences; maximum is 9."
                )

        for field in [
            "assetClasses",
            "geographies",
            "sectors",
        ]:
            value = analysis.get(field)

            if not isinstance(value, list):
                errors.append(
                    f"{prefix}.{field} must be an array."
                )
            elif not all(
                isinstance(item, str)
                for item in value
            ):
                errors.append(
                    f"{prefix}.{field} must contain strings only."
                )

        for field in [
            "investorImpact",
            "reasoning",
        ]:
            value = analysis.get(field)

            if not isinstance(value, str):
                errors.append(
                    f"{prefix}.{field} must be a string."
                )

        if not isinstance(
            analysis.get("fundMonitoringRelevant"),
            bool,
        ):
            errors.append(
                f"{prefix}.fundMonitoringRelevant "
                "must be boolean."
            )

    if len(analyses) != len(articles):
        errors.append(
            "Analysis count mismatch: expected "
            f"{len(articles)}, got {len(analyses)}."
        )

    if len(actual_ids) != len(set(actual_ids)):
        errors.append(
            "Duplicate articleId values returned."
        )

    if set(actual_ids) != set(expected_ids):
        missing_ids = set(expected_ids) - set(actual_ids)
        extra_ids = set(actual_ids) - set(expected_ids)

        if missing_ids:
            errors.append(
                "Missing article IDs: "
                + ", ".join(sorted(missing_ids))
            )

        if extra_ids:
            errors.append(
                "Unexpected article IDs: "
                + ", ".join(sorted(extra_ids))
            )

    return len(errors) == 0, errors


# ============================================================
# OUTPUT ENVELOPE
# ============================================================

def build_output(
    articles: list[dict],
    prompt: str,
    parsed_response: dict,
    metadata: dict,
    validation_errors: list[str],
) -> dict:

    analyses = parsed_response["analyses"]

    return {
        "generatedAt": datetime.now(
            timezone.utc
        ).isoformat(),

        "source": {
            "file": str(
                CURRENT_JSON.relative_to(ROOT)
            ),
            "articleCount": len(articles),
        },

        "configuration": {
            "model": DEFAULT_MODEL,
            "maxArticles": MAX_ARTICLES,
            "maxBatches": MAX_BATCHES,
            "maxOutputTokens": MAX_OUTPUT_TOKENS,
            "summaryMaximumSentences": 9,
        },

        "batch": {
            "batchNumber": 3,
            "articleCount": len(articles),
            "articleIds": [
                article_id(article)
                for article in articles
            ],
        },

        "prompt": prompt,

        "request": {
            "model": metadata["model"],
            "requestBody": metadata["requestBody"],
            "promptCharacters": metadata[
                "promptCharacters"
            ],
            "responseSeconds": metadata[
                "responseSeconds"
            ],
            "maxOutputTokens": metadata[
                "maxOutputTokens"
            ],
            "responseMimeType": metadata[
                "responseMimeType"
            ],
            "usage": metadata.get(
                "usage",
                {},
            ),
        },

        "inputArticles": articles,

        "analyses": analyses,

        "geminiResponse": {
            "rawResponse": metadata[
                "rawResponse"
            ],
            "generatedText": metadata[
                "generatedText"
            ],
        },

        "validation": {
            "passed": True,
            "errors": validation_errors,
        },
    }


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    log("=" * 70)
    log("VGrat FMS - MARKET NEWS AI ANALYZER")
    log("=" * 70)

    model = os.getenv(
        "GEMINI_MODEL",
        DEFAULT_MODEL,
    )

    api_key = os.getenv("GEMINI_API_KEY")

    log(f"Gemini model: {model}")
    log(f"Articles per test: {MAX_ARTICLES}")
    log("Mode: analysis only")
    log("Correction mode: disabled")
    log("Research Funds.xlsx: not used")

    if not api_key:
        log("ERROR: GEMINI_API_KEY is not set.")
        return 1

    # --------------------------------------------------------
    # Load current market news.
    # --------------------------------------------------------

    log("")
    log(f"Loading market news: {CURRENT_JSON}")

    if not CURRENT_JSON.exists():
        log("ERROR: current.json does not exist.")
        return 1

    try:
        current_data = load_json(CURRENT_JSON)
        all_articles = get_article_list(current_data)
    except Exception as exc:
        log(f"ERROR loading current.json: {exc}")
        return 1

    log(f"Articles found: {len(all_articles)}")

    if not all_articles:
        log("ERROR: No articles found.")
        return 1

    # --------------------------------------------------------
    # Normalize and sort newest first.
    # --------------------------------------------------------

    prepared_articles = [
        prepare_article(article)
        for article in all_articles
    ]

    prepared_articles = sorted(
        enumerate(prepared_articles),
        key=lambda item: (
            article_sort_timestamp(item[1]),
            item[0],
        ),
        reverse=True,
    )

    newest_articles = [
        item[1]
        for item in prepared_articles[:MAX_ARTICLES]
    ]

    log("")
    log("=" * 70)
    log("SELECTED ARTICLES")
    log("=" * 70)

    for index, article in enumerate(
        newest_articles,
        start=1,
    ):
        title = (
            article.get("title")
            or article.get("headline")
            or "Untitled"
        )

        log(
            f"{index}. "
            f"{article['articleId']} | {title}"
        )

    # --------------------------------------------------------
    # Build prompt.
    # --------------------------------------------------------

    prompt = build_prompt(
        newest_articles
    )

    log("")
    log(
        f"Prompt characters: {len(prompt)}"
    )

    # --------------------------------------------------------
    # Gemini request.
    # --------------------------------------------------------

    log("")
    log("=" * 70)
    log("GEMINI ANALYSIS")
    log("=" * 70)

    try:
        parsed_response, metadata = call_gemini(
            api_key=api_key,
            model=model,
            prompt=prompt,
        )

    except Exception as exc:
        log("")
        log("=" * 70)
        log("ANALYSIS FAILED")
        log("=" * 70)
        log(str(exc))
        log("")
        log("NO CHANGES WRITTEN.")
        return 1

    # --------------------------------------------------------
    # Validate.
    # --------------------------------------------------------

    log("")
    log("=" * 70)
    log("VALIDATING RESPONSE")
    log("=" * 70)

    valid, errors = validate_analysis(
        parsed_response,
        newest_articles,
    )

    if not valid:
        log("Validation failed.")

        for error in errors:
            log(f"- {error}")

        log("")
        log("NO CHANGES WRITTEN.")
        return 1

    log("Validation passed.")
    log(
        f"Analyses returned: "
        f"{len(parsed_response['analyses'])}"
    )

    log(
        f"Gemini response time: "
        f"{metadata['responseSeconds']}s"
    )

    usage = metadata.get("usage", {})

    if usage:
        log(
            "Token usage: "
            f"prompt={usage.get('promptTokenCount', '-')}, "
            f"candidates={usage.get('candidatesTokenCount', '-')}, "
            f"thoughts={usage.get('thoughtsTokenCount', '-')}, "
            f"total={usage.get('totalTokenCount', '-')}"
        )

    # --------------------------------------------------------
    # Build final output.
    # --------------------------------------------------------

    output = build_output(
        articles=newest_articles,
        prompt=prompt,
        parsed_response=parsed_response,
        metadata=metadata,
        validation_errors=[],
    )

    # --------------------------------------------------------
    # ONLY NOW write the file.
    # --------------------------------------------------------

    log("")
    log("=" * 70)
    log("WRITING ANALYSIS")
    log("=" * 70)

    try:
        save_json_atomic(
            ANALYSIS_CURRENT,
            output,
        )
    except Exception as exc:
        log(f"ERROR writing analysis: {exc}")
        return 1

    log(
        f"Successfully wrote: "
        f"{ANALYSIS_CURRENT}"
    )

    log("")
    log("=" * 70)
    log("SUCCESS")
    log("=" * 70)

    return 0


if __name__ == "__main__":
    sys.exit(main())
