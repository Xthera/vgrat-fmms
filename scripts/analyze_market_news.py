#!/usr/bin/env python3

"""
VGrat FMS - GEMINI MARKET NEWS ANALYSIS
=======================================

Reads:
    data/market_news/current.json
    Research Funds.xlsx

Writes:
    data/market_news/analysis/current.json
    data/market_news/analysis/history/YYYY/MM/YYYY-MM-DD.json

Purpose:
    Analyze every NEW article found in market_news/current.json using
    Google's Gemini API.

AI MODEL:
    Configurable through MARKET_NEWS_MODEL.
    Default:
        gemini-3.8-flash

IMPORTANT:
    - Original article identity is preserved.
    - Original title and URL are never rewritten.
    - Article publication time is treated as Singapore time.
    - Previously analyzed articles are not analyzed again.
    - Historical analysis is stored by Singapore publication date.
    - The current analysis file contains a rolling 14-day window.
    - Geography and sector taxonomies are read dynamically from
      Research Funds.xlsx.
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests
import trafilatura
from openpyxl import load_workbook
from google import genai


# ============================================================
# PATHS
# ============================================================

ROOT = Path(__file__).resolve().parent.parent

CURRENT_NEWS = ROOT / "data" / "market_news" / "current.json"

ANALYSIS_ROOT = ROOT / "data" / "market_news" / "analysis"
ANALYSIS_CURRENT = ANALYSIS_ROOT / "current.json"
ANALYSIS_HISTORY = ANALYSIS_ROOT / "history"

RESEARCH_FUNDS = ROOT / "Research Funds.xlsx"


# ============================================================
# CONFIGURATION
# ============================================================

MODEL_NAME = os.getenv(
    "MARKET_NEWS_MODEL",
    "gemini-3.8-flash",
)

AI_RETRIES = int(
    os.getenv(
        "MARKET_NEWS_AI_RETRIES",
        "3",
    )
)

ARTICLE_RETRIEVAL_RETRIES = int(
    os.getenv(
        "MARKET_NEWS_RETRIEVAL_RETRIES",
        "2",
    )
)

ARTICLE_RETRIEVAL_TIMEOUT = int(
    os.getenv(
        "MARKET_NEWS_RETRIEVAL_TIMEOUT",
        "20",
    )
)

MAX_ARTICLE_CHARS = int(
    os.getenv(
        "MARKET_NEWS_MAX_ARTICLE_CHARS",
        "30000",
    )
)

MAX_SUMMARY_CHARS = int(
    os.getenv(
        "MARKET_NEWS_MAX_SUMMARY_CHARS",
        "8000",
    )
)

WINDOW_DAYS = 14

REQUEST_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 "
    "(KHTML, like Gecko) "
    "Chrome/154.0.0.0 Safari/537.36"
)

SGT = timezone(timedelta(hours=8))


# ============================================================
# ALLOWED ANALYSIS VALUES
# ============================================================

ALLOWED_IMPACT_DIRECTION = {
    "Positive",
    "Negative",
    "Mixed",
    "Neutral",
}

ALLOWED_IMPACT_SEVERITY = {
    "Low",
    "Moderate",
    "High",
    "Critical",
}

ALLOWED_TIME_HORIZON = {
    "Immediate",
    "Short-term",
    "Medium-term",
    "Long-term",
}

ALLOWED_AI_CONFIDENCE = {
    "High",
    "Medium",
    "Low",
}


# ============================================================
# LOGGING
# ============================================================

PREFIX = "[MARKET-NEWS-ANALYSIS]"


def log(message: str) -> None:
    print(f"{PREFIX} {message}", flush=True)


def fail(message: str) -> None:
    log(f"ERROR: {message}")
    sys.exit(1)


# ============================================================
# TIME
# ============================================================

def now_sgt() -> datetime:
    return datetime.now(SGT)


def now_sgt_iso() -> str:
    return now_sgt().isoformat()


def parse_sgt_datetime(value: str) -> datetime:
    """
    Parse an ISO datetime and normalize it to Singapore time.

    current.json already stores publishedAtSgt, so this function
    preserves the instant and normalizes it to SGT.
    """

    if not value:
        raise ValueError("Empty datetime value.")

    dt = datetime.fromisoformat(value)

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=SGT)

    return dt.astimezone(SGT)


def publication_date(article: Dict[str, Any]) -> str:
    """
    Return YYYY-MM-DD using the article's Singapore publication time.
    """

    published = article.get("publishedAtSgt")

    if not published:
        raise ValueError(
            f"Article {article.get('id')} has no publishedAtSgt."
        )

    return parse_sgt_datetime(published).strftime("%Y-%m-%d")


# ============================================================
# JSON HELPERS
# ============================================================

def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

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
# TAXONOMY
# ============================================================

def clean_taxonomy_values(
    values: List[Any],
) -> List[str]:

    cleaned: List[str] = []

    for value in values:
        text = str(value).strip()

        if not text:
            continue

        cleaned.append(text)

    # Preserve order while removing duplicates.
    result: List[str] = []
    seen = set()

    for value in cleaned:
        if value not in seen:
            seen.add(value)
            result.append(value)

    return result


def read_taxonomy_sheet(
    workbook,
    sheet_index: int,
) -> List[str]:

    if sheet_index >= len(workbook.worksheets):
        raise ValueError(
            f"Research Funds.xlsx does not contain worksheet "
            f"index {sheet_index}."
        )

    ws = workbook.worksheets[sheet_index]

    values: List[Any] = []

    # Find the first column containing taxonomy values.
    for row in ws.iter_rows(values_only=True):
        for value in row:
            if value is not None and str(value).strip():
                values.append(value)
                break

    categories = clean_taxonomy_values(values)

    # Remove an obvious header if present.
    if categories:
        first_lower = categories[0].lower()

        headers = {
            "geography",
            "geographies",
            "geographic",
            "geographic category",
            "geography category",
            "sector",
            "sectors",
            "sector category",
        }

        if first_lower in headers:
            categories = categories[1:]

    return categories


def load_taxonomies() -> Tuple[List[str], List[str]]:

    if not RESEARCH_FUNDS.exists():
        raise FileNotFoundError(
            f"Missing taxonomy workbook: {RESEARCH_FUNDS}"
        )

    wb = load_workbook(
        RESEARCH_FUNDS,
        read_only=True,
        data_only=True,
    )

    try:
        # Worksheet index 1 = Geography master
        # Worksheet index 2 = Sector master
        geography = read_taxonomy_sheet(wb, 1)
        sector = read_taxonomy_sheet(wb, 2)
    finally:
        wb.close()

    if not geography:
        raise ValueError(
            "Geography taxonomy is empty."
        )

    if not sector:
        raise ValueError(
            "Sector taxonomy is empty."
        )

    return geography, sector


def taxonomy_prompt(
    geography: List[str],
    sector: List[str],
) -> str:

    geography_text = "\n".join(
        f"- {item}"
        for item in geography
    )

    sector_text = "\n".join(
        f"- {item}"
        for item in sector
    )

    return f"""
GEOGRAPHY MASTER TAXONOMY
The following are the ONLY valid geography categories:

{geography_text}

SECTOR MASTER TAXONOMY
The following are the ONLY valid sector categories:

{sector_text}

Rules for taxonomy selection:

1. Select between 1 and 3 geography categories.
2. Select between 1 and 3 sector categories.
3. Do not force three categories if fewer are genuinely relevant.
4. Use the category names EXACTLY as written above.
5. Do not invent new categories.
6. Do not rename categories.
7. Do not combine categories.
8. Do not use broader or narrower alternatives that are not listed.
9. Only select categories materially affected by the article.
10. The labels themselves must contain only the exact category names.
"""


# ============================================================
# ORIGINAL ARTICLE RETRIEVAL
# ============================================================

def retrieve_original_article(
    url: str,
) -> Optional[str]:

    if not url:
        return None

    headers = {
        "User-Agent": REQUEST_USER_AGENT,
        "Accept": (
            "text/html,application/xhtml+xml,"
            "application/xml;q=0.9,*/*;q=0.8"
        ),
        "Accept-Language": "en-SG,en;q=0.9",
    }

    for attempt in range(
        1,
        ARTICLE_RETRIEVAL_RETRIES + 1,
    ):
        try:
            response = requests.get(
                url,
                headers=headers,
                timeout=ARTICLE_RETRIEVAL_TIMEOUT,
                allow_redirects=True,
            )

            response.raise_for_status()

            extracted = trafilatura.extract(
                response.text,
                include_comments=False,
                include_tables=True,
                favor_precision=True,
            )

            if extracted:
                extracted = extracted.strip()

                if extracted:
                    return extracted[:MAX_ARTICLE_CHARS]

            log(
                f"Original article extraction returned no usable "
                f"text (attempt {attempt})."
            )

        except Exception as exc:
            log(
                f"Original article retrieval failed "
                f"(attempt {attempt}): {exc}"
            )

        if attempt < ARTICLE_RETRIEVAL_RETRIES:
            time.sleep(2)

    return None


# ============================================================
# ANALYSIS STORAGE
# ============================================================

def load_existing_analysis() -> Dict[str, Dict[str, Any]]:
    """
    Load all existing historical analysis records.

    This prevents previously analyzed articles from being
    sent to Gemini again.
    """

    existing: Dict[str, Dict[str, Any]] = {}

    if not ANALYSIS_HISTORY.exists():
        return existing

    for path in ANALYSIS_HISTORY.rglob("*.json"):

        try:
            data = load_json(path)

        except Exception as exc:
            log(
                f"Skipping unreadable analysis history "
                f"{path}: {exc}"
            )
            continue

        if not isinstance(data, dict):
            continue

        articles = data.get("articles", [])

        if not isinstance(articles, list):
            continue

        for article in articles:

            if not isinstance(article, dict):
                continue

            article_id = article.get("id")

            if article_id:
                existing[str(article_id)] = article

    # Also consider current.json.
    if ANALYSIS_CURRENT.exists():

        try:
            data = load_json(ANALYSIS_CURRENT)

            if isinstance(data, dict):

                for article in data.get(
                    "articles",
                    [],
                ):

                    if not isinstance(article, dict):
                        continue

                    article_id = article.get("id")

                    if article_id:
                        existing[str(article_id)] = article

        except Exception as exc:
            log(
                f"Unable to read existing analysis/current.json: "
                f"{exc}"
            )

    return existing


# ============================================================
# GEMINI RESPONSE PARSING
# ============================================================

def extract_json_object(
    text: str,
) -> Dict[str, Any]:

    text = text.strip()

    # Direct JSON.
    try:
        value = json.loads(text)

        if isinstance(value, dict):
            return value

    except json.JSONDecodeError:
        pass

    # Markdown fenced JSON.
    fenced = re.findall(
        r"```(?:json)?\s*(\{.*?\})\s*```",
        text,
        flags=re.DOTALL | re.IGNORECASE,
    )

    for block in fenced:

        try:
            value = json.loads(block)

            if isinstance(value, dict):
                return value

        except json.JSONDecodeError:
            continue

    # Find first JSON object.
    start = text.find("{")

    if start >= 0:

        depth = 0
        in_string = False
        escaped = False

        for index in range(
            start,
            len(text),
        ):

            char = text[index]

            if escaped:
                escaped = False
                continue

            if char == "\\":
                escaped = True
                continue

            if char == '"':
                in_string = not in_string
                continue

            if in_string:
                continue

            if char == "{":
                depth += 1

            elif char == "}":
                depth -= 1

                if depth == 0:

                    candidate = text[
                        start:index + 1
                    ]

                    try:
                        value = json.loads(
                            candidate
                        )

                        if isinstance(value, dict):
                            return value

                    except json.JSONDecodeError:
                        break

    raise ValueError(
        "Gemini response did not contain valid JSON."
    )


# ============================================================
# VALIDATION
# ============================================================

def require_text(
    value: Any,
    field: str,
) -> str:

    if not isinstance(value, str):
        raise ValueError(
            f"{field} must be a string."
        )

    value = value.strip()

    if not value:
        raise ValueError(
            f"{field} cannot be empty."
        )

    return value


def validate_categories(
    value: Any,
    allowed: List[str],
    field: str,
) -> List[str]:

    if not isinstance(value, list):
        raise ValueError(
            f"{field} must be a list."
        )

    if not 1 <= len(value) <= 3:
        raise ValueError(
            f"{field} must contain between 1 and 3 categories."
        )

    result: List[str] = []

    for item in value:

        if not isinstance(item, str):
            raise ValueError(
                f"{field} contains a non-string value."
            )

        item = item.strip()

        if item not in allowed:
            raise ValueError(
                f"Invalid {field} category: '{item}'"
            )

        if item in result:
            raise ValueError(
                f"Duplicate {field} category: '{item}'"
            )

        result.append(item)

    return result


def validate_analysis(
    analysis: Dict[str, Any],
    geography: List[str],
    sector: List[str],
) -> Dict[str, Any]:

    if not isinstance(analysis, dict):
        raise ValueError(
            "Analysis result must be an object."
        )

    quick = analysis.get("quickRead")

    if not isinstance(quick, dict):
        raise ValueError(
            "Missing quickRead object."
        )

    detailed = analysis.get("detailedAnalysis")

    if not isinstance(detailed, dict):
        raise ValueError(
            "Missing detailedAnalysis object."
        )

    quick_result = {
        "whatHappened": require_text(
            quick.get("whatHappened"),
            "quickRead.whatHappened",
        ),
        "whyItMatters": require_text(
            quick.get("whyItMatters"),
            "quickRead.whyItMatters",
        ),
        "keyImpact": require_text(
            quick.get("keyImpact"),
            "quickRead.keyImpact",
        ),
    }

    detailed_result = {
        "background": require_text(
            detailed.get("background"),
            "detailedAnalysis.background",
        ),
        "keyDevelopments": require_text(
            detailed.get("keyDevelopments"),
            "detailedAnalysis.keyDevelopments",
        ),
        "marketImplications": require_text(
            detailed.get("marketImplications"),
            "detailedAnalysis.marketImplications",
        ),
        "geographicalImpactExplanation": require_text(
            detailed.get(
                "geographicalImpactExplanation"
            ),
            "detailedAnalysis.geographicalImpactExplanation",
        ),
        "sectorImpactExplanation": require_text(
            detailed.get(
                "sectorImpactExplanation"
            ),
            "detailedAnalysis.sectorImpactExplanation",
        ),
        "timeHorizonExplanation": require_text(
            detailed.get(
                "timeHorizonExplanation"
            ),
            "detailedAnalysis.timeHorizonExplanation",
        ),
        "overallAssessment": require_text(
            detailed.get("overallAssessment"),
            "detailedAnalysis.overallAssessment",
        ),
    }

    return {
        "quickRead": quick_result,

        "geographicalImpact": validate_categories(
            analysis.get("geographicalImpact"),
            geography,
            "geographicalImpact",
        ),

        "sectorImpact": validate_categories(
            analysis.get("sectorImpact"),
            sector,
            "sectorImpact",
        ),

        "overallImpactDirection": validate_enum(
            analysis.get("overallImpactDirection"),
            ALLOWED_IMPACT_DIRECTION,
            "overallImpactDirection",
        ),

        "impactSeverity": validate_enum(
            analysis.get("impactSeverity"),
            ALLOWED_IMPACT_SEVERITY,
            "impactSeverity",
        ),

        "timeHorizon": validate_enum(
            analysis.get("timeHorizon"),
            ALLOWED_TIME_HORIZON,
            "timeHorizon",
        ),

        "aiConfidence": validate_enum(
            analysis.get("aiConfidence"),
            ALLOWED_AI_CONFIDENCE,
            "aiConfidence",
        ),

        "detailedAnalysis": detailed_result,
    }


def validate_enum(
    value: Any,
    allowed: set,
    field: str,
) -> str:

    if not isinstance(value, str):
        raise ValueError(
            f"{field} must be a string."
        )

    value = value.strip()

    if value not in allowed:
        raise ValueError(
            f"Invalid {field}: '{value}'"
        )

    return value


# ============================================================
# GEMINI ANALYSIS
# ============================================================

def build_prompt(
    article: Dict[str, Any],
    original_text: Optional[str],
    geography: List[str],
    sector: List[str],
) -> str:

    title = str(
        article.get(
            "title",
            "",
        )
    ).strip()

    summary = str(
        article.get(
            "summary",
            "",
        )
    ).strip()

    if len(summary) > MAX_SUMMARY_CHARS:
        summary = summary[:MAX_SUMMARY_CHARS]

    source = str(
        article.get(
            "source",
            "",
        )
    ).strip()

    published = str(
        article.get(
            "publishedAtSgt",
            "",
        )
    ).strip()

    category = str(
        article.get(
            "category",
            "",
        )
    ).strip()

    relevance_score = article.get(
        "relevanceScore"
    )

    relevance_reason = str(
        article.get(
            "relevanceReason",
            "",
        )
    ).strip()

    original_section = (
        original_text
        if original_text
        else
        "Original article text could not be retrieved. "
        "Use the supplied article information only."
    )

    taxonomy = taxonomy_prompt(
        geography,
        sector,
    )

    return f"""
You are the market-news analysis engine for VGrat FMS.

Analyze the following article objectively, factually, and in detail.

The purpose is market and economic intelligence.

Do NOT:
- give investment advice;
- recommend investments;
- recommend individual funds;
- predict individual fund performance;
- fabricate facts;
- invent information not supported by the article;
- exaggerate uncertain consequences.

IMPORTANT:
The collector's category, relevance score, and relevance reason are
context only. They must NOT determine the analysis automatically.

============================================================
ARTICLE IDENTITY
============================================================

Article ID:
{article.get("id", "")}

Source:
{source}

Original Title:
{title}

Original URL:
{article.get("url", "")}

Published Singapore Time:
{published}

Collector Category:
{category}

Collector Relevance Score:
{relevance_score}

Collector Relevance Reason:
{relevance_reason}

============================================================
COLLECTOR SUMMARY
============================================================

{summary}

============================================================
ORIGINAL ARTICLE CONTENT
============================================================

{original_section}

============================================================
TAXONOMY
============================================================

{taxonomy}

============================================================
ANALYSIS REQUIREMENTS
============================================================

Every article must receive:

1. Quick Read Summary

What happened:
A concise factual description of the event.

Why it matters:
Why the event matters economically, financially, commercially,
or for markets.

Key impact:
The most important immediate or potential consequence.

2. Geography Impact

Select 1 to 3 genuinely relevant geography categories.

Use ONLY exact category names from the supplied master taxonomy.

3. Sector Impact

Select 1 to 3 genuinely relevant sector categories.

Use ONLY exact category names from the supplied master taxonomy.

4. Overall Impact Direction

Choose exactly one:

Positive
Negative
Mixed
Neutral

This represents the combined economic and financial-market
implication of the event, NOT an investment recommendation.

5. Impact Severity

Choose exactly one:

Low
Moderate
High
Critical

This describes the potential significance of the event.

6. Time Horizon

Choose exactly one:

Immediate
Short-term
Medium-term
Long-term

7. AI Confidence

Choose exactly one:

High
Medium
Low

This represents confidence in the analysis and classification,
NOT certainty that the predicted consequence will occur.

============================================================
DETAILED ANALYSIS
============================================================

Provide:

Background
Key Developments
Market Implications
Geographical Impact Explanation
Sector Impact Explanation
Time Horizon Explanation
Overall Assessment

Do not create a separate Evidence section.

============================================================
OUTPUT
============================================================

Return ONLY valid JSON.

Use exactly this structure:

{{
  "quickRead": {{
    "whatHappened": "...",
    "whyItMatters": "...",
    "keyImpact": "..."
  }},
  "geographicalImpact": [
    "exact taxonomy category"
  ],
  "sectorImpact": [
    "exact taxonomy category"
  ],
  "overallImpactDirection": "Positive",
  "impactSeverity": "Moderate",
  "timeHorizon": "Short-term",
  "aiConfidence": "High",
  "detailedAnalysis": {{
    "background": "...",
    "keyDevelopments": "...",
    "marketImplications": "...",
    "geographicalImpactExplanation": "...",
    "sectorImpactExplanation": "...",
    "timeHorizonExplanation": "...",
    "overallAssessment": "..."
  }}
}}

Remember:

- Geography: 1 to 3 categories.
- Sector: 1 to 3 categories.
- Use exact taxonomy names.
- Do not invent taxonomy names.
- Return JSON only.
"""


def analyze_with_gemini(
    client: genai.Client,
    article: Dict[str, Any],
    original_text: Optional[str],
    geography: List[str],
    sector: List[str],
) -> Dict[str, Any]:

    prompt = build_prompt(
        article,
        original_text,
        geography,
        sector,
    )

    last_error: Optional[Exception] = None

    for attempt in range(
        1,
        AI_RETRIES + 1,
    ):

        log(
            f"AI analysis for {article.get('id')} "
            f"(attempt {attempt}/{AI_RETRIES})"
        )

        try:

            interaction = client.interactions.create(
                model=MODEL_NAME,
                input=prompt,
            )

            output_text = (
                interaction.output_text
                if interaction
                else ""
            )

            if not output_text:
                raise ValueError(
                    "Gemini returned empty output."
                )

            raw = extract_json_object(
                output_text
            )

            return validate_analysis(
                raw,
                geography,
                sector,
            )

        except Exception as exc:

            last_error = exc

            log(
                f"AI validation failed for "
                f"{article.get('id')}: {exc}"
            )

            if attempt < AI_RETRIES:
                time.sleep(2)

    raise RuntimeError(
        f"Gemini analysis failed after "
        f"{AI_RETRIES} attempts: {last_error}"
    )


# ============================================================
# ARTICLE RECORD
# ============================================================

def build_analyzed_article(
    article: Dict[str, Any],
    analysis: Dict[str, Any],
    original_retrieved: bool,
) -> Dict[str, Any]:

    result = dict(article)

    result["analysis"] = analysis

    content_sources = [
        "market_news/current.json"
    ]

    if original_retrieved:
        content_sources.append(
            "original_article_url"
        )

    result["analysisMetadata"] = {
        "model": MODEL_NAME,
        "analysedAtSgt": now_sgt_iso(),
        "contentSources": content_sources,
    }

    return result


# ============================================================
# HISTORY
# ============================================================

def history_path(
    article: Dict[str, Any],
) -> Path:

    date = publication_date(article)

    year, month, _ = date.split("-")

    return (
        ANALYSIS_HISTORY
        / year
        / month
        / f"{date}.json"
    )


def load_history_file(
    path: Path,
) -> Dict[str, Any]:

    if not path.exists():
        return {
            "timezone": "Asia/Singapore",
            "timezoneLabel": "SGT",
            "articles": [],
        }

    try:
        data = load_json(path)

        if isinstance(data, dict):
            return data

    except Exception as exc:
        log(
            f"Unable to load history file {path}: {exc}"
        )

    return {
        "timezone": "Asia/Singapore",
        "timezoneLabel": "SGT",
        "articles": [],
    }


def upsert_history_article(
    article: Dict[str, Any],
) -> None:

    path = history_path(article)

    data = load_history_file(path)

    existing = data.get(
        "articles",
        [],
    )

    if not isinstance(existing, list):
        existing = []

    article_id = article.get("id")

    replaced = False

    for index, item in enumerate(existing):

        if (
            isinstance(item, dict)
            and item.get("id") == article_id
        ):
            existing[index] = article
            replaced = True
            break

    if not replaced:
        existing.append(article)

    existing.sort(
        key=lambda item: (
            item.get(
                "publishedAtSgt",
                "",
            ),
            item.get(
                "id",
                "",
            ),
        )
    )

    data["timezone"] = "Asia/Singapore"
    data["timezoneLabel"] = "SGT"
    data["publicationDateSgt"] = publication_date(
        article
    )
    data["articleCount"] = len(existing)
    data["articles"] = existing

    save_json(
        path,
        data,
    )


# ============================================================
# CURRENT 14-DAY FILE
# ============================================================

def rebuild_analysis_current(
    existing_articles: Dict[str, Dict[str, Any]],
    news_data: Dict[str, Any],
) -> None:

    articles: List[Dict[str, Any]] = []

    now = now_sgt()

    cutoff = now - timedelta(
        days=WINDOW_DAYS
    )

    for article in existing_articles.values():

        if not isinstance(article, dict):
            continue

        published = article.get(
            "publishedAtSgt"
        )

        if not published:
            continue

        try:
            published_dt = parse_sgt_datetime(
                published
            )

        except Exception:
            continue

        if published_dt >= cutoff:
            articles.append(article)

    articles.sort(
        key=lambda item: (
            item.get(
                "publishedAtSgt",
                "",
            ),
            item.get(
                "id",
                "",
            ),
        ),
        reverse=True,
    )

    # Keep the current article window aligned with the
    # collector's 14-day window.
    max_articles = news_data.get(
        "maxArticles"
    )

    if isinstance(max_articles, int) and max_articles > 0:
        articles = articles[:max_articles]

    output = {
        "generatedAtSgt": now_sgt_iso(),
        "timezone": "Asia/Singapore",
        "timezoneLabel": "SGT",
        "windowDays": WINDOW_DAYS,
        "articleCount": len(articles),
        "source": news_data.get(
            "source",
            "CNBC",
        ),
        "articles": articles,
    }

    save_json(
        ANALYSIS_CURRENT,
        output,
    )


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    log("Starting Gemini Market News Analysis.")

    # --------------------------------------------------------
    # API KEY
    # --------------------------------------------------------

    api_key = os.getenv(
        "GEMINI_API_KEY"
    )

    if not api_key:
        fail(
            "GEMINI_API_KEY environment variable is missing."
        )

    # --------------------------------------------------------
    # INPUT FILE
    # --------------------------------------------------------

    if not CURRENT_NEWS.exists():
        fail(
            f"Missing input file: {CURRENT_NEWS}"
        )

    news_data = load_json(
        CURRENT_NEWS
    )

    if not isinstance(news_data, dict):
        fail(
            "current.json must contain a JSON object."
        )

    news_articles = news_data.get(
        "articles",
        [],
    )

    if not isinstance(news_articles, list):
        fail(
            "current.json articles must be a list."
        )

    log(
        f"Loaded {len(news_articles)} articles "
        f"from current.json."
    )

    # --------------------------------------------------------
    # TAXONOMY
    # --------------------------------------------------------

    geography, sector = load_taxonomies()

    log(
        f"Loaded {len(geography)} geography categories "
        f"and {len(sector)} sector categories."
    )

    # --------------------------------------------------------
    # EXISTING ANALYSIS
    # --------------------------------------------------------

    existing = load_existing_analysis()

    log(
        f"Loaded {len(existing)} previously analyzed articles."
    )

    # --------------------------------------------------------
    # GEMINI CLIENT
    # --------------------------------------------------------

    client = genai.Client(
        api_key=api_key
    )

    # --------------------------------------------------------
    # DETERMINE NEW ARTICLES
    # --------------------------------------------------------

    pending: List[Dict[str, Any]] = []

    for article in news_articles:

        if not isinstance(article, dict):
            continue

        article_id = article.get("id")

        if not article_id:
            log(
                "Skipping article without ID."
            )
            continue

        if str(article_id) not in existing:
            pending.append(article)

    log(
        f"New articles requiring Gemini analysis: "
        f"{len(pending)}"
    )

    # --------------------------------------------------------
    # ANALYZE NEW ARTICLES
    # --------------------------------------------------------

    success_count = 0
    failure_count = 0

    for article in pending:

        article_id = article.get(
            "id",
            "",
        )

        title = article.get(
            "title",
            "",
        )

        log(
            f"Processing article {article_id}: {title}"
        )

        try:

            original_text = retrieve_original_article(
                article.get(
                    "url",
                    "",
                )
            )

            original_retrieved = bool(
                original_text
            )

            if original_retrieved:
                log(
                    f"Original article retrieved for "
                    f"{article_id}."
                )
            else:
                log(
                    f"Using current.json content only for "
                    f"{article_id}."
                )

            analysis = analyze_with_gemini(
                client,
                article,
                original_text,
                geography,
                sector,
            )

            analyzed_article = build_analyzed_article(
                article,
                analysis,
                original_retrieved,
            )

            existing[str(article_id)] = analyzed_article

            upsert_history_article(
                analyzed_article
            )

            success_count += 1

            log(
                f"Successfully analyzed {article_id}."
            )

        except Exception as exc:

            failure_count += 1

            log(
                f"FAILED article {article_id}: {exc}"
            )

    # --------------------------------------------------------
    # REBUILD CURRENT ANALYSIS WINDOW
    # --------------------------------------------------------

    rebuild_analysis_current(
        existing,
        news_data,
    )

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    log("=" * 60)
    log("ANALYSIS COMPLETE")
    log("=" * 60)
    log(
        f"Articles in current.json: {len(news_articles)}"
    )
    log(
        f"Previously analyzed: "
        f"{len(existing) - success_count}"
    )
    log(
        f"New articles attempted: {len(pending)}"
    )
    log(
        f"New articles successfully analyzed: "
        f"{success_count}"
    )
    log(
        f"New article failures: {failure_count}"
    )
    log(
        f"Analysis current file: {ANALYSIS_CURRENT}"
    )
    log(
        f"Analysis history: {ANALYSIS_HISTORY}"
    )

    # --------------------------------------------------------
    # IMPORTANT:
    #
    # Do NOT fail the workflow just because some articles failed.
    # Failed articles remain absent from history and will be
    # retried on the next workflow run.
    # --------------------------------------------------------

    return 0


if __name__ == "__main__":
    raise SystemExit(
        main()
    )
