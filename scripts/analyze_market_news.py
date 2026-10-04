#!/usr/bin/env python3

"""
VGrat FMS - MARKET NEWS ANALYZER
================================

Purpose
-------
Analyze articles collected by the CNBC Market News Collector.

PRIMARY INPUT
-------------
data/market_news/current.json

TAXONOMY INPUT
--------------
Research Funds.xlsx

    Sheet 1: Fund Research
    Sheet 2: Geography master categories
    Sheet 3: Sector master categories

OUTPUT
------
data/market_news/analysis/
    current.json

    history/
        YYYY/
            MM/
                YYYY-MM-DD.json

IMPORTANT PROCESSING RULE
-------------------------
ONE ARTICLE = ONE GEMINI REQUEST

An article is considered successfully analyzed ONLY after its
analysis record has been successfully written to the historical
publication-date file.

Once an article exists successfully in history, it is permanently
considered analyzed and will NEVER be sent to Gemini again.

The rolling current.json is NOT the completion baseline.

PROCESSING ORDER
----------------
Pending articles are processed oldest publication date first.

SUCCESS
-------
1. Gemini returns valid analysis.
2. Analysis passes validation.
3. Article is written successfully to history.
4. Only then is the article marked as completed in memory.

RATE LIMIT
----------
If Gemini returns HTTP 429 / quota exhaustion:
    - stop immediately
    - do not process further articles
    - leave remaining articles pending for the next run

OTHER FAILURES
--------------
For non-429 article failures:
    - do not retry Gemini in the same run
    - leave article pending
    - continue with the next pending article

TAXONOMY
--------
Geography and sector labels MUST come from Research Funds.xlsx.

The AI is instructed to:
    - use only exact taxonomy labels
    - select 1-3 geography labels
    - select 1-3 sector labels
    - explain why each selected geography/sector is impacted

NO FUND RECOMMENDATIONS
-----------------------
This module does not:
    - recommend Prudential funds
    - predict individual fund returns
    - analyse individual Prudential funds
    - provide investment advice

TIME
----
All publication-date grouping uses Singapore time.

MODEL
-----
Default:
    gemini-3.8-flash

API
---
Google Gemini Interactions API.

Environment variable:
    GEMINI_API_KEY
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
from google import genai
from openpyxl import load_workbook


# ============================================================
# PATHS
# ============================================================

ROOT = Path(__file__).resolve().parent.parent

NEWS_CURRENT_PATH = (
    ROOT / "data" / "market_news" / "current.json"
)

ANALYSIS_ROOT = (
    ROOT / "data" / "market_news" / "analysis"
)

ANALYSIS_CURRENT_PATH = (
    ANALYSIS_ROOT / "current.json"
)

ANALYSIS_HISTORY_ROOT = (
    ANALYSIS_ROOT / "history"
)

RESEARCH_FUNDS_PATH = (
    ROOT / "Research Funds.xlsx"
)


# ============================================================
# CONFIGURATION
# ============================================================

MODEL_NAME = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.8-flash",
)

MAX_ARTICLES = int(
    os.getenv("MAX_ARTICLES", "100")
)

CURRENT_DAYS = int(
    os.getenv("CURRENT_DAYS", "14")
)

ARTICLE_RETRIEVAL_RETRIES = int(
    os.getenv("ARTICLE_RETRIEVAL_RETRIES", "2")
)

ARTICLE_RETRIEVAL_TIMEOUT = int(
    os.getenv("ARTICLE_RETRIEVAL_TIMEOUT", "20")
)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/154.0.0.0 Safari/537.36"
)

SINGAPORE_TZ = timezone(
    timedelta(hours=8),
    name="Asia/Singapore",
)


# ============================================================
# LOGGING
# ============================================================

def log(message: str) -> None:
    print(
        f"[{datetime.now(SINGAPORE_TZ).strftime('%Y-%m-%d %H:%M:%S')}] "
        f"{message}",
        flush=True,
    )


# ============================================================
# BASIC HELPERS
# ============================================================

def now_sgt() -> datetime:
    return datetime.now(SINGAPORE_TZ)


def now_sgt_iso() -> str:
    return now_sgt().isoformat()


def safe_string(value: Any) -> str:
    if value is None:
        return ""

    return str(value).strip()


def normalise_label(value: Any) -> str:
    """
    Normalise whitespace only.

    Taxonomy labels otherwise remain unchanged.
    """
    return re.sub(r"\s+", " ", safe_string(value)).strip()


# ============================================================
# DATE HANDLING
# ============================================================

def parse_sgt_datetime(value: Any) -> Optional[datetime]:
    """
    Parse publishedAtSgt into Singapore-aware datetime.
    """

    if not value:
        return None

    text = safe_string(value)

    try:
        dt = datetime.fromisoformat(
            text.replace("Z", "+00:00")
        )

        if dt.tzinfo is None:
            dt = dt.replace(
                tzinfo=SINGAPORE_TZ
            )

        return dt.astimezone(SINGAPORE_TZ)

    except Exception:
        pass

    formats = [
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%d %H:%M:%S%z",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(text, fmt)

            if dt.tzinfo is None:
                dt = dt.replace(
                    tzinfo=SINGAPORE_TZ
                )

            return dt.astimezone(SINGAPORE_TZ)

        except Exception:
            continue

    return None


def publication_date_sgt(article: Dict[str, Any]) -> Optional[str]:
    dt = parse_sgt_datetime(
        article.get("publishedAtSgt")
    )

    if dt is None:
        return None

    return dt.strftime("%Y-%m-%d")


# ============================================================
# JSON LOADING
# ============================================================

def load_json(path: Path, default: Any = None) -> Any:
    if not path.exists():
        return default

    with path.open(
        "r",
        encoding="utf-8",
    ) as f:
        return json.load(f)


def save_json(path: Path, data: Any) -> None:
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
# NEWS INPUT
# ============================================================

def load_current_news() -> List[Dict[str, Any]]:
    log(
        f"Loading news: {NEWS_CURRENT_PATH}"
    )

    data = load_json(
        NEWS_CURRENT_PATH,
        default={},
    )

    if isinstance(data, list):
        articles = data

    elif isinstance(data, dict):
        articles = data.get(
            "articles",
            [],
        )

    else:
        raise ValueError(
            "current.json must contain either "
            "a list or an object containing 'articles'."
        )

    if not isinstance(articles, list):
        raise ValueError(
            "current.json 'articles' must be a list."
        )

    valid_articles: List[Dict[str, Any]] = []

    for article in articles:
        if not isinstance(article, dict):
            continue

        article_id = safe_string(
            article.get("id")
        )

        if not article_id:
            continue

        valid_articles.append(article)

    log(
        f"Loaded {len(valid_articles)} valid news articles."
    )

    return valid_articles


# ============================================================
# RESEARCH FUNDS TAXONOMY
# ============================================================

def extract_values_from_sheet(
    ws,
) -> List[str]:
    """
    Extract non-empty cell values from a worksheet.

    The function intentionally does not assume a specific
    column because the master workbook may have headers or
    formatting variations.
    """

    values: List[str] = []

    for row in ws.iter_rows():
        for cell in row:
            value = normalise_label(
                cell.value
            )

            if not value:
                continue

            values.append(value)

    return values


def deduplicate_preserve_order(
    values: List[str],
) -> List[str]:

    seen = set()
    output = []

    for value in values:
        key = value.casefold()

        if key in seen:
            continue

        seen.add(key)
        output.append(value)

    return output


def load_taxonomy() -> Tuple[
    List[str],
    List[str],
]:
    """
    Load geography and sector master categories.

    Expected workbook:
        Sheet 2 = Geography master categories
        Sheet 3 = Sector master categories
    """

    if not RESEARCH_FUNDS_PATH.exists():
        raise FileNotFoundError(
            f"Research workbook not found: "
            f"{RESEARCH_FUNDS_PATH}"
        )

    log(
        f"Loading taxonomy from: "
        f"{RESEARCH_FUNDS_PATH.name}"
    )

    wb = load_workbook(
        RESEARCH_FUNDS_PATH,
        read_only=True,
        data_only=True,
    )

    try:
        sheet_names = wb.sheetnames

        if len(sheet_names) < 3:
            raise ValueError(
                "Research Funds.xlsx must contain at least "
                "three worksheets."
            )

        geography_ws = wb.worksheets[1]
        sector_ws = wb.worksheets[2]

        geography = deduplicate_preserve_order(
            extract_values_from_sheet(
                geography_ws
            )
        )

        sectors = deduplicate_preserve_order(
            extract_values_from_sheet(
                sector_ws
            )
        )

    finally:
        wb.close()

    if not geography:
        raise ValueError(
            "No geography categories found in "
            "Research Funds.xlsx."
        )

    if not sectors:
        raise ValueError(
            "No sector categories found in "
            "Research Funds.xlsx."
        )

    log(
        f"Geography taxonomy: {len(geography)} categories."
    )

    log(
        f"Sector taxonomy: {len(sectors)} categories."
    )

    return geography, sectors


# ============================================================
# HISTORY
# ============================================================

def iter_history_files():
    if not ANALYSIS_HISTORY_ROOT.exists():
        return

    yield from ANALYSIS_HISTORY_ROOT.rglob(
        "*.json"
    )


def extract_history_records(
    data: Any,
) -> List[Dict[str, Any]]:
    """
    Extract article records from a historical JSON file.

    Supports:
        {"articles": [...]}
        {"results": [...]}
        [...]
        {"id": ..., "analysis": {...}}
    """

    if isinstance(data, list):
        return [
            item
            for item in data
            if isinstance(item, dict)
        ]

    if not isinstance(data, dict):
        return []

    if (
        isinstance(data.get("articles"), list)
    ):
        return [
            item
            for item in data["articles"]
            if isinstance(item, dict)
        ]

    if (
        isinstance(data.get("results"), list)
    ):
        return [
            item
            for item in data["results"]
            if isinstance(item, dict)
        ]

    if data.get("id") is not None:
        return [data]

    return []


def is_successful_analysis_record(
    record: Dict[str, Any],
) -> bool:
    """
    History is the permanent completion baseline.

    A record counts as successfully analysed when:
        - it has an article ID
        - it contains a non-empty analysis object
    """

    article_id = safe_string(
        record.get("id")
    )

    analysis = record.get(
        "analysis"
    )

    if not article_id:
        return False

    if not isinstance(
        analysis,
        dict,
    ):
        return False

    return bool(analysis)


def load_existing_analysis() -> Dict[
    str,
    Dict[str, Any],
]:
    """
    Scan ALL historical files.

    This is deliberately not limited to the current
    14-day window.

    Once an article has a successful history record,
    it stays completed permanently.
    """

    existing: Dict[
        str,
        Dict[str, Any],
    ] = {}

    files = list(
        iter_history_files()
        or []
    )

    log(
        f"Scanning {len(files)} historical analysis files."
    )

    for path in files:

        try:
            data = load_json(
                path,
                default=None,
            )

        except Exception as exc:
            log(
                f"WARNING: Could not read history "
                f"{path}: {exc}"
            )
            continue

        records = extract_history_records(
            data
        )

        for record in records:

            if not is_successful_analysis_record(
                record
            ):
                continue

            article_id = safe_string(
                record.get("id")
            )

            existing[article_id] = record

    log(
        f"Found {len(existing)} successfully analysed "
        f"articles in history."
    )

    return existing


# ============================================================
# ORIGINAL ARTICLE RETRIEVAL
# ============================================================

def retrieve_original_article(
    article: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Retrieve original article text where possible.

    If retrieval fails, current.json content is used as
    the fallback source.
    """

    url = safe_string(
        article.get("url")
    )

    fallback = {
        "source": "collector_current_json",
        "title": safe_string(
            article.get("title")
        ),
        "text": safe_string(
            article.get("summary")
        ),
    }

    if not url:
        return fallback

    headers = {
        "User-Agent": USER_AGENT,
        "Accept": (
            "text/html,application/xhtml+xml,"
            "application/xml;q=0.9,*/*;q=0.8"
        ),
    }

    for attempt in range(
        1,
        ARTICLE_RETRIEVAL_RETRIES + 2,
    ):

        try:
            log(
                f"Retrieving original article "
                f"(attempt {attempt}): {url}"
            )

            response = requests.get(
                url,
                headers=headers,
                timeout=ARTICLE_RETRIEVAL_TIMEOUT,
            )

            response.raise_for_status()

            downloaded = trafilatura.extract(
                response.text,
                include_comments=False,
                include_tables=True,
                include_links=True,
            )

            if downloaded:
                return {
                    "source": "original_article",
                    "title": safe_string(
                        article.get("title")
                    ),
                    "text": downloaded,
                }

            log(
                "Original page retrieved but "
                "no usable article text was extracted."
            )

        except Exception as exc:
            log(
                f"Original article retrieval failed: {exc}"
            )

        if attempt <= ARTICLE_RETRIEVAL_RETRIES:
            time.sleep(1)

    log(
        "Using current.json article data as fallback."
    )

    return fallback


# ============================================================
# GEMINI PROMPT
# ============================================================

def build_prompt(
    article: Dict[str, Any],
    article_content: Dict[str, Any],
    geography_categories: List[str],
    sector_categories: List[str],
) -> str:

    article_id = safe_string(
        article.get("id")
    )

    title = safe_string(
        article.get("title")
    )

    source = safe_string(
        article.get("source")
    )

    url = safe_string(
        article.get("url")
    )

    published_at = safe_string(
        article.get("publishedAtSgt")
    )

    collector_category = safe_string(
        article.get("category")
    )

    relevance_score = article.get(
        "relevanceScore"
    )

    relevance_reason = safe_string(
        article.get("relevanceReason")
    )

    original_text = safe_string(
        article_content.get("text")
    )

    # Keep prompts reasonably sized while preserving a large
    # amount of article context.
    max_text_chars = 30000

    if len(original_text) > max_text_chars:
        original_text = (
            original_text[:max_text_chars]
            + "\n\n[Article text truncated for analysis.]"
        )

    geography_text = "\n".join(
        f"- {label}"
        for label in geography_categories
    )

    sector_text = "\n".join(
        f"- {label}"
        for label in sector_categories
    )

    return f"""
You are the market-news analysis engine for VGrat FMS.

Analyze ONE financial/economic/business news article.

Your job is to explain the article objectively and identify
the geographical and industry-sector areas that are materially
affected.

This is NOT investment advice.

Do NOT recommend securities, funds, stocks, ETFs, insurance
products, or individual Prudential funds.

Do NOT predict the future return of any individual investment
fund.

============================================================
ARTICLE IDENTITY
============================================================

Article ID:
{article_id}

Title:
{title}

Source:
{source}

URL:
{url}

Published Singapore Time:
{published_at}

Collector Category:
{collector_category}

Collector Relevance Score:
{relevance_score}

Collector Relevance Reason:
{relevance_reason}

============================================================
ARTICLE CONTENT
============================================================

{original_text}

============================================================
GEOGRAPHY MASTER TAXONOMY
============================================================

You MUST select geography labels ONLY from this exact list.

{geography_text}

Rules:
- Select 1 to 3 labels.
- Do not invent labels.
- Do not rename labels.
- Do not combine labels.
- Do not output synonyms.
- Use "Global" only when the article genuinely has broad
  global implications.
- Select the most materially affected geographic areas.

============================================================
SECTOR MASTER TAXONOMY
============================================================

You MUST select sector labels ONLY from this exact list.

{sector_text}

Rules:
- Select 1 to 3 labels.
- Do not invent labels.
- Do not rename labels.
- Do not combine labels.
- Do not output synonyms.
- Select sectors that are materially affected by the
  developments in the article.

============================================================
REQUIRED ANALYSIS
============================================================

Return EXACTLY ONE JSON OBJECT.

The JSON structure MUST be:

{{
  "id": "{article_id}",
  "analysis": {{
    "quickRead": {{
      "whatHappened": "...",
      "whyItMatters": "...",
      "keyImpact": "..."
    }},

    "geographyImpact": {{
      "labels": [
        "EXACT TAXONOMY LABEL"
      ],
      "explanation": "Explain why the selected geographies are affected."
    }},

    "industrySectorImpact": {{
      "labels": [
        "EXACT TAXONOMY LABEL"
      ],
      "explanation": "Explain why the selected industry sectors are affected."
    }},

    "overallImpactDirection": "Positive | Negative | Mixed | Neutral",

    "impactSeverity": "Low | Moderate | High | Critical",

    "timeHorizon": "Immediate | Short-term | Medium-term | Long-term",

    "aiConfidence": "High | Medium | Low",

    "detailedAnalysis": {{
      "background": "...",
      "keyDevelopments": "...",
      "marketImplications": "...",
      "geographicalImpactExplanation": "...",
      "industrySectorImpactExplanation": "...",
      "timeHorizonExplanation": "...",
      "overallAssessment": "..."
    }}
  }}
}}

============================================================
ANALYSIS RULES
============================================================

1. whatHappened
   State the central factual development.

2. whyItMatters
   Explain why the development matters for markets, economies,
   businesses, investors or policy.

3. keyImpact
   Summarise the most important market/economic consequence.

4. Geography
   Select 1-3 exact geography taxonomy labels.

5. Geography explanation
   Explain the actual transmission mechanism:
   trade, supply chains, monetary policy, fiscal policy,
   currency, commodities, regulation, investment flows,
   growth, inflation, employment, etc., where relevant.

6. Industry sectors
   Select 1-3 exact sector taxonomy labels.

7. Industry-sector explanation
   Explain how the article could affect the selected sectors.
   Discuss demand, costs, revenues, margins, regulation,
   investment, financing, supply chains or other relevant
   mechanisms where appropriate.

8. Overall Impact Direction
   This is the combined market/economic direction:
       Positive
       Negative
       Mixed
       Neutral

   It is NOT a prediction of asset prices.

9. Impact Severity
   This measures significance, NOT certainty:
       Low
       Moderate
       High
       Critical

10. Time Horizon
       Immediate
       Short-term
       Medium-term
       Long-term

11. AI Confidence
   This measures confidence in the quality of the analysis
   and classification, NOT confidence that the predicted
   outcome will happen.

12. Detailed Analysis
   Provide meaningful analysis in every required section.

13. Political or geopolitical stories
   Remain factual and neutral.
   Do not advocate, campaign, rank political actors, or make
   unsupported predictions about political outcomes.

14. Do not fabricate facts that are not supported by the
   article or reliable context contained in it.

15. If the article contains uncertainty, clearly state the
   uncertainty.

16. Return valid JSON only.
   No markdown.
   No code fence.
   No commentary before or after the JSON.

17. The returned "id" MUST exactly equal:
   {article_id}
"""


# ============================================================
# GEMINI RESPONSE PARSING
# ============================================================

def extract_json_object(
    text: str,
) -> Optional[Dict[str, Any]]:
    """
    Robustly extract a JSON object from Gemini output.

    Handles:
        - pure JSON
        - ```json ... ```
        - surrounding commentary
        - multiple candidate '{' positions
    """

    if not text:
        return None

    cleaned = text.strip()

    # --------------------------------------------------------
    # Direct JSON
    # --------------------------------------------------------

    try:
        value = json.loads(cleaned)

        if isinstance(value, dict):
            return value

    except Exception:
        pass

    # --------------------------------------------------------
    # Markdown fenced JSON
    # --------------------------------------------------------

    fenced = re.search(
        r"```(?:json)?\s*(.*?)\s*```",
        cleaned,
        flags=re.IGNORECASE | re.DOTALL,
    )

    if fenced:
        candidate = fenced.group(1).strip()

        try:
            value = json.loads(candidate)

            if isinstance(value, dict):
                return value

        except Exception:
            pass

    # --------------------------------------------------------
    # Raw decoder search
    # --------------------------------------------------------

    decoder = json.JSONDecoder()

    for index, character in enumerate(
        cleaned
    ):
        if character != "{":
            continue

        candidate = cleaned[index:]

        try:
            value, _ = decoder.raw_decode(
                candidate
            )

            if isinstance(value, dict):
                return value

        except Exception:
            continue

    return None


# ============================================================
# GEMINI ERROR CLASSIFICATION
# ============================================================

def is_rate_limit_error(
    exc: Exception,
) -> bool:

    text = str(exc).lower()

    indicators = [
        "429",
        "too_many_requests",
        "resource_exhausted",
        "rate limit",
        "quota",
        "quota exceeded",
        "daily limit",
    ]

    return any(
        indicator in text
        for indicator in indicators
    )


# ============================================================
# GEMINI ANALYSIS
# ============================================================

def analyse_one_article(
    client: genai.Client,
    article: Dict[str, Any],
    geography_categories: List[str],
    sector_categories: List[str],
) -> Tuple[
    Optional[Dict[str, Any]],
    Optional[str],
]:
    """
    Perform exactly ONE Gemini request for this article.

    Returns:
        (analysis_record, error_type)

    error_type:
        None
        "rate_limit"
        "request_error"
        "invalid_json"
        "validation_error"
    """

    article_id = safe_string(
        article.get("id")
    )

    log(
        f"Sending article to Gemini: "
        f"{article_id}"
    )

    article_content = retrieve_original_article(
        article
    )

    prompt = build_prompt(
        article,
        article_content,
        geography_categories,
        sector_categories,
    )

    try:
        interaction = client.interactions.create(
            model=MODEL_NAME,
            input=prompt,
        )

    except Exception as exc:

        if is_rate_limit_error(exc):
            log(
                f"RATE LIMIT / QUOTA STOP: {article_id}: {exc}"
            )

            return None, "rate_limit"

        log(
            f"Gemini request failed for {article_id}: "
            f"{exc}"
        )

        return None, "request_error"

    output_text = safe_string(
        getattr(
            interaction,
            "output_text",
            "",
        )
    )

    if not output_text:
        log(
            f"Gemini returned empty output: {article_id}"
        )

        return None, "invalid_json"

    parsed = extract_json_object(
        output_text
    )

    if parsed is None:
        log(
            f"Could not parse Gemini JSON for {article_id}."
        )

        log(
            "Gemini response preview:"
        )

        log(
            output_text[:4000]
        )

        return None, "invalid_json"

    valid, error = validate_analysis(
        parsed,
        article,
        geography_categories,
        sector_categories,
    )

    if not valid:
        log(
            f"Analysis validation failed for "
            f"{article_id}: {error}"
        )

        return None, "validation_error"

    return parsed, None


# ============================================================
# VALIDATION
# ============================================================

ALLOWED_DIRECTION = {
    "Positive",
    "Negative",
    "Mixed",
    "Neutral",
}

ALLOWED_SEVERITY = {
    "Low",
    "Moderate",
    "High",
    "Critical",
}

ALLOWED_HORIZON = {
    "Immediate",
    "Short-term",
    "Medium-term",
    "Long-term",
}

ALLOWED_CONFIDENCE = {
    "High",
    "Medium",
    "Low",
}


def validate_string(
    value: Any,
) -> bool:

    return (
        isinstance(value, str)
        and bool(value.strip())
    )


def validate_label_list(
    value: Any,
    allowed: List[str],
) -> Tuple[
    bool,
    str,
]:

    if not isinstance(value, list):
        return (
            False,
            "labels must be a list",
        )

    if not 1 <= len(value) <= 3:
        return (
            False,
            "labels must contain 1-3 values",
        )

    allowed_set = set(allowed)

    for label in value:

        if not isinstance(
            label,
            str,
        ):
            return (
                False,
                "taxonomy label must be a string",
            )

        if label not in allowed_set:
            return (
                False,
                f"invalid taxonomy label: {label}",
            )

    if len(set(value)) != len(value):
        return (
            False,
            "duplicate taxonomy labels",
        )

    return True, ""


def validate_analysis(
    result: Dict[str, Any],
    article: Dict[str, Any],
    geography_categories: List[str],
    sector_categories: List[str],
) -> Tuple[
    bool,
    str,
]:

    expected_id = safe_string(
        article.get("id")
    )

    returned_id = safe_string(
        result.get("id")
    )

    if returned_id != expected_id:
        return (
            False,
            (
                "returned id does not match article id "
                f"(expected={expected_id}, "
                f"returned={returned_id})"
            ),
        )

    analysis = result.get(
        "analysis"
    )

    if not isinstance(
        analysis,
        dict,
    ):
        return (
            False,
            "analysis must be an object",
        )

    # --------------------------------------------------------
    # Quick Read
    # --------------------------------------------------------

    quick_read = analysis.get(
        "quickRead"
    )

    if not isinstance(
        quick_read,
        dict,
    ):
        return (
            False,
            "quickRead must be an object",
        )

    for field in (
        "whatHappened",
        "whyItMatters",
        "keyImpact",
    ):
        if not validate_string(
            quick_read.get(field)
        ):
            return (
                False,
                f"quickRead.{field} missing or invalid",
            )

    # --------------------------------------------------------
    # Geography
    # --------------------------------------------------------

    geography = analysis.get(
        "geographyImpact"
    )

    if not isinstance(
        geography,
        dict,
    ):
        return (
            False,
            "geographyImpact must be an object",
        )

    valid, error = validate_label_list(
        geography.get("labels"),
        geography_categories,
    )

    if not valid:
        return (
            False,
            f"geographyImpact: {error}",
        )

    if not validate_string(
        geography.get("explanation")
    ):
        return (
            False,
            "geographyImpact.explanation missing",
        )

    # --------------------------------------------------------
    # Industry Sector
    # --------------------------------------------------------

    sector = analysis.get(
        "industrySectorImpact"
    )

    if not isinstance(
        sector,
        dict,
    ):
        return (
            False,
            "industrySectorImpact must be an object",
        )

    valid, error = validate_label_list(
        sector.get("labels"),
        sector_categories,
    )

    if not valid:
        return (
            False,
            f"industrySectorImpact: {error}",
        )

    if not validate_string(
        sector.get("explanation")
    ):
        return (
            False,
            "industrySectorImpact.explanation missing",
        )

    # --------------------------------------------------------
    # Classification fields
    # --------------------------------------------------------

    direction = analysis.get(
        "overallImpactDirection"
    )

    if direction not in ALLOWED_DIRECTION:
        return (
            False,
            "invalid overallImpactDirection",
        )

    severity = analysis.get(
        "impactSeverity"
    )

    if severity not in ALLOWED_SEVERITY:
        return (
            False,
            "invalid impactSeverity",
        )

    horizon = analysis.get(
        "timeHorizon"
    )

    if horizon not in ALLOWED_HORIZON:
        return (
            False,
            "invalid timeHorizon",
        )

    confidence = analysis.get(
        "aiConfidence"
    )

    if confidence not in ALLOWED_CONFIDENCE:
        return (
            False,
            "invalid aiConfidence",
        )

    # --------------------------------------------------------
    # Detailed Analysis
    # --------------------------------------------------------

    detailed = analysis.get(
        "detailedAnalysis"
    )

    if not isinstance(
        detailed,
        dict,
    ):
        return (
            False,
            "detailedAnalysis must be an object",
        )

    required_detailed_fields = [
        "background",
        "keyDevelopments",
        "marketImplications",
        "geographicalImpactExplanation",
        "industrySectorImpactExplanation",
        "timeHorizonExplanation",
        "overallAssessment",
    ]

    for field in required_detailed_fields:

        if not validate_string(
            detailed.get(field)
        ):
            return (
                False,
                f"detailedAnalysis.{field} missing or invalid",
            )

    return True, ""


# ============================================================
# HISTORY FILE OPERATIONS
# ============================================================

def history_path_for_article(
    article: Dict[str, Any],
) -> Optional[Path]:

    publication_date = publication_date_sgt(
        article
    )

    if not publication_date:
        return None

    year, month, day = publication_date.split(
        "-"
    )

    return (
        ANALYSIS_HISTORY_ROOT
        / year
        / month
        / f"{publication_date}.json"
    )


def load_history_file(
    path: Path,
) -> Dict[str, Any]:

    if not path.exists():
        return {
            "date": path.stem,
            "timezone": "Asia/Singapore",
            "articles": [],
        }

    data = load_json(
        path,
        default={},
    )

    if isinstance(data, list):
        return {
            "date": path.stem,
            "timezone": "Asia/Singapore",
            "articles": data,
        }

    if not isinstance(data, dict):
        return {
            "date": path.stem,
            "timezone": "Asia/Singapore",
            "articles": [],
        }

    if not isinstance(
        data.get("articles"),
        list,
    ):
        data["articles"] = []

    return data


def upsert_history_article(
    article: Dict[str, Any],
    analysis_result: Dict[str, Any],
) -> Path:

    path = history_path_for_article(
        article
    )

    if path is None:
        raise ValueError(
            "Cannot save article to history because "
            "publishedAtSgt is missing or invalid."
        )

    history = load_history_file(
        path
    )

    articles = history.setdefault(
        "articles",
        [],
    )

    article_id = safe_string(
        article.get("id")
    )

    # --------------------------------------------------------
    # Preserve original article identity.
    # --------------------------------------------------------

    record = dict(article)

    record["analysis"] = (
        analysis_result.get("analysis")
    )

    record["analysisMetadata"] = {
        "model": MODEL_NAME,
        "analysedAtSgt": now_sgt_iso(),
        "contentSource": (
            "original_article_or_current_json_fallback"
        ),
    }

    replaced = False

    for index, existing_record in enumerate(
        articles
    ):

        if not isinstance(
            existing_record,
            dict,
        ):
            continue

        if safe_string(
            existing_record.get("id")
        ) != article_id:
            continue

        articles[index] = record
        replaced = True
        break

    if not replaced:
        articles.append(
            record
        )

    # Keep historical file deterministic by publication date.
    articles.sort(
        key=lambda item: (
            safe_string(
                item.get("publishedAtSgt")
            ),
            safe_string(
                item.get("id")
            ),
        )
    )

    # --------------------------------------------------------
    # CRITICAL:
    # The function does not return until the file has been
    # successfully written.
    #
    # The caller only treats the article as completed after
    # this function returns successfully.
    # --------------------------------------------------------

    save_json(
        path,
        history,
    )

    log(
        f"HISTORY WRITE SUCCESS: {article_id} -> {path}"
    )

    return path


# ============================================================
# CURRENT ROLLING ANALYSIS
# ============================================================

def rebuild_analysis_current(
    news_articles: List[Dict[str, Any]],
    existing_analysis: Dict[str, Dict[str, Any]],
) -> None:

    cutoff = (
        now_sgt()
        - timedelta(
            days=CURRENT_DAYS
        )
    )

    output_articles: List[
        Dict[str, Any]
    ] = []

    for article in news_articles:

        article_id = safe_string(
            article.get("id")
        )

        if not article_id:
            continue

        dt = parse_sgt_datetime(
            article.get("publishedAtSgt")
        )

        if dt is None:
            continue

        if dt < cutoff:
            continue

        analysed = existing_analysis.get(
            article_id
        )

        if not analysed:
            continue

        output_articles.append(
            analysed
        )

    output_articles.sort(
        key=lambda item: (
            safe_string(
                item.get("publishedAtSgt")
            ),
            safe_string(
                item.get("id")
            ),
        ),
        reverse=True,
    )

    output_articles = output_articles[
        :MAX_ARTICLES
    ]

    output = {
        "generatedAtSgt": now_sgt_iso(),
        "windowDays": CURRENT_DAYS,
        "articleCount": len(
            output_articles
        ),
        "articles": output_articles,
    }

    save_json(
        ANALYSIS_CURRENT_PATH,
        output,
    )

    log(
        f"Rebuilt analysis/current.json "
        f"with {len(output_articles)} articles."
    )


# ============================================================
# PENDING QUEUE
# ============================================================

def build_pending_queue(
    news_articles: List[Dict[str, Any]],
    existing_analysis: Dict[str, Dict[str, Any]],
) -> List[Dict[str, Any]]:

    pending: List[
        Tuple[datetime, Dict[str, Any]]
    ] = []

    for article in news_articles:

        article_id = safe_string(
            article.get("id")
        )

        if not article_id:
            continue

        # ----------------------------------------------------
        # Permanent success check.
        #
        # This checks ALL history, not current.json.
        # ----------------------------------------------------

        if article_id in existing_analysis:
            continue

        dt = parse_sgt_datetime(
            article.get("publishedAtSgt")
        )

        if dt is None:
            log(
                f"WARNING: Cannot parse publishedAtSgt "
                f"for pending article {article_id}; "
                f"skipping this run."
            )
            continue

        pending.append(
            (
                dt,
                article,
            )
        )

    # Oldest article first.
    pending.sort(
        key=lambda item: (
            item[0],
            safe_string(
                item[1].get("id")
            ),
        )
    )

    return [
        article
        for _, article in pending
    ]


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    log(
        "============================================================"
    )
    log(
        "VGrat FMS - MARKET NEWS ANALYZER"
    )
    log(
        "============================================================"
    )

    # --------------------------------------------------------
    # Environment
    # --------------------------------------------------------

    api_key = os.getenv(
        "GEMINI_API_KEY"
    )

    if not api_key:
        log(
            "ERROR: GEMINI_API_KEY is not set."
        )
        return 1

    # --------------------------------------------------------
    # Inputs
    # --------------------------------------------------------

    try:
        news_articles = load_current_news()

        geography_categories, sector_categories = (
            load_taxonomy()
        )

        existing_analysis = (
            load_existing_analysis()
        )

    except Exception as exc:

        log(
            f"FATAL INPUT ERROR: {exc}"
        )

        return 1

    # --------------------------------------------------------
    # Queue
    # --------------------------------------------------------

    pending = build_pending_queue(
        news_articles,
        existing_analysis,
    )

    already_analysed = (
        len(news_articles)
        - len(pending)
    )

    log(
        f"Total news articles: {len(news_articles)}"
    )

    log(
        f"Already successfully analysed: "
        f"{already_analysed}"
    )

    log(
        f"Pending articles: {len(pending)}"
    )

    if pending:

        oldest = pending[0]

        log(
            "Oldest pending article: "
            f"{safe_string(oldest.get('publishedAtSgt'))} | "
            f"{safe_string(oldest.get('title'))}"
        )

    else:

        log(
            "No pending articles."
        )

    # --------------------------------------------------------
    # Gemini client
    # --------------------------------------------------------

    client = genai.Client(
        api_key=api_key
    )

    successful_this_run = 0
    failed_this_run = 0
    requests_attempted = 0
    quota_stopped = False

    # --------------------------------------------------------
    # Process ONE article at a time.
    # --------------------------------------------------------

    for article in pending:

        article_id = safe_string(
            article.get("id")
        )

        title = safe_string(
            article.get("title")
        )

        log(
            "------------------------------------------------------------"
        )

        log(
            f"Processing article: {article_id}"
        )

        log(
            f"Published: "
            f"{safe_string(article.get('publishedAtSgt'))}"
        )

        log(
            f"Title: {title}"
        )

        requests_attempted += 1

        result, error_type = analyse_one_article(
            client,
            article,
            geography_categories,
            sector_categories,
        )

        # ----------------------------------------------------
        # RATE LIMIT
        # ----------------------------------------------------

        if error_type == "rate_limit":

            quota_stopped = True

            log(
                "Gemini quota/rate limit reached."
            )

            log(
                "Stopping immediately."
            )

            log(
                "Remaining articles remain pending "
                "for the next scheduled run."
            )

            break

        # ----------------------------------------------------
        # ARTICLE FAILURE
        # ----------------------------------------------------

        if result is None:

            failed_this_run += 1

            log(
                f"Article NOT completed: {article_id}"
            )

            log(
                "No same-run Gemini retry will be attempted."
            )

            log(
                "Article remains pending for a future run."
            )

            continue

        # ----------------------------------------------------
        # HISTORY WRITE
        # ----------------------------------------------------

        try:

            history_path = upsert_history_article(
                article,
                result,
            )

        except Exception as exc:

            log(
                f"FATAL HISTORY WRITE ERROR for "
                f"{article_id}: {exc}"
            )

            log(
                "Article has NOT been marked as completed."
            )

            log(
                "Stopping to avoid wasting further "
                "Gemini requests."
            )

            # Do not add to existing_analysis.
            # Therefore it remains pending next run.

            failed_this_run += 1

            break

        # ----------------------------------------------------
        # CRITICAL SUCCESS POINT
        #
        # Only now does the article become permanently
        # completed in this process.
        # ----------------------------------------------------

        existing_analysis[
            article_id
        ] = dict(article)

        existing_analysis[
            article_id
        ]["analysis"] = result.get(
            "analysis"
        )

        existing_analysis[
            article_id
        ]["analysisMetadata"] = {
            "model": MODEL_NAME,
            "analysedAtSgt": now_sgt_iso(),
            "contentSource": (
                "original_article_or_current_json_fallback"
            ),
            "historyPath": str(
                history_path.relative_to(ROOT)
            ),
        }

        successful_this_run += 1

        log(
            f"ARTICLE SUCCESSFULLY COMPLETED: "
            f"{article_id}"
        )

        log(
            "This article will not be analysed again."
        )

    # --------------------------------------------------------
    # Rebuild rolling current analysis.
    # --------------------------------------------------------

    try:

        rebuild_analysis_current(
            news_articles,
            existing_analysis,
        )

    except Exception as exc:

        log(
            f"ERROR rebuilding analysis/current.json: "
            f"{exc}"
        )

        # History records already saved successfully should
        # remain valid even if current.json rebuilding fails.
        return 1

    # --------------------------------------------------------
    # Final statistics
    # --------------------------------------------------------

    remaining_pending = build_pending_queue(
        news_articles,
        existing_analysis,
    )

    log(
        "============================================================"
    )

    log(
        "RUN SUMMARY"
    )

    log(
        f"Total news articles: {len(news_articles)}"
    )

    log(
        f"Already analysed before run: "
        f"{already_analysed}"
    )

    log(
        f"Pending at run start: "
        f"{len(pending)}"
    )

    log(
        f"Gemini requests attempted: "
        f"{requests_attempted}"
    )

    log(
        f"Successfully analysed this run: "
        f"{successful_this_run}"
    )

    log(
        f"Failed this run: "
        f"{failed_this_run}"
    )

    log(
        f"Quota stopped: "
        f"{str(quota_stopped).lower()}"
    )

    log(
        f"Still pending: "
        f"{len(remaining_pending)}"
    )

    if remaining_pending:

        next_article = remaining_pending[0]

        log(
            "Next article in queue: "
            f"{safe_string(next_article.get('publishedAtSgt'))} | "
            f"{safe_string(next_article.get('title'))}"
        )

    else:

        log(
            "No remaining pending articles."
        )

    log(
        "============================================================"
    )

    # --------------------------------------------------------
    # Important:
    #
    # Partial processing is still a successful workflow run.
    # GitHub Actions must be allowed to commit the successfully
    # saved history records.
    # --------------------------------------------------------

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )
