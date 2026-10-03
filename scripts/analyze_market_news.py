
#!/usr/bin/env python3
"""
VGrat FMS - GEMINI MARKET NEWS BATCH ANALYSIS
=============================================

Input:
    data/market_news/current.json
    Research Funds.xlsx

Output:
    data/market_news/analysis/current.json
    data/market_news/analysis/history/YYYY/MM/YYYY-MM-DD.json

Batch behaviour:
    - Default batch size: 3 articles per Gemini request.
    - Each article receives an independent analysis object.
    - Each article is validated independently.
    - Successful articles are saved immediately.
    - Failed articles remain pending.
    - HTTP 429 stops the entire run immediately.
    - Existing successful articles are never reanalysed.

Test configuration:
    MARKET_NEWS_BATCH_SIZE=3
    MARKET_NEWS_MAX_BATCHES=1

The test configuration processes only the first batch of 3 articles.
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

BATCH_SIZE = max(
    1,
    int(os.getenv("MARKET_NEWS_BATCH_SIZE", "3")),
)

MAX_BATCHES = max(
    0,
    int(os.getenv("MARKET_NEWS_MAX_BATCHES", "1")),
)

AI_RETRIES = max(
    1,
    int(os.getenv("MARKET_NEWS_AI_RETRIES", "3")),
)

ARTICLE_RETRIEVAL_RETRIES = max(
    1,
    int(os.getenv("MARKET_NEWS_RETRIEVAL_RETRIES", "2")),
)

ARTICLE_RETRIEVAL_TIMEOUT = int(
    os.getenv("MARKET_NEWS_RETRIEVAL_TIMEOUT", "20")
)

MAX_ARTICLE_CHARS = int(
    os.getenv("MARKET_NEWS_MAX_ARTICLE_CHARS", "30000")
)

MAX_SUMMARY_CHARS = int(
    os.getenv("MARKET_NEWS_MAX_SUMMARY_CHARS", "8000")
)

WINDOW_DAYS = 14

REQUEST_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 "
    "(KHTML, like Gecko) "
    "Chrome/154.0.0.0 Safari/537.36"
)

SGT = timezone(timedelta(hours=8))

PREFIX = "[MARKET-NEWS-ANALYSIS]"


# ============================================================
# ENUMERATIONS
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

    if not value:
        raise ValueError("Empty datetime value.")

    dt = datetime.fromisoformat(value)

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=SGT)

    return dt.astimezone(SGT)


def publication_date(article: Dict[str, Any]) -> str:

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

    temporary = path.with_suffix(path.suffix + ".tmp")

    with temporary.open("w", encoding="utf-8") as f:
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

def clean_taxonomy_values(values: List[Any]) -> List[str]:

    cleaned = [
        str(value).strip()
        for value in values
        if value is not None and str(value).strip()
    ]

    result = []
    seen = set()

    for value in cleaned:
        if value not in seen:
            seen.add(value)
            result.append(value)

    return result


def read_taxonomy_sheet(workbook, sheet_index: int) -> List[str]:

    if sheet_index >= len(workbook.worksheets):
        raise ValueError(
            f"Research Funds.xlsx does not contain worksheet "
            f"index {sheet_index}."
        )

    ws = workbook.worksheets[sheet_index]
    values = []

    for row in ws.iter_rows(values_only=True):
        for value in row:
            if value is not None and str(value).strip():
                values.append(value)
                break

    categories = clean_taxonomy_values(values)

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

    if categories and categories[0].lower() in headers:
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
        geography = read_taxonomy_sheet(wb, 1)
        sector = read_taxonomy_sheet(wb, 2)
    finally:
        wb.close()

    if not geography:
        raise ValueError("Geography taxonomy is empty.")

    if not sector:
        raise ValueError("Sector taxonomy is empty.")

    return geography, sector


def taxonomy_prompt(
    geography: List[str],
    sector: List[str],
) -> str:

    geography_text = "\n".join(
        f"- {item}" for item in geography
    )

    sector_text = "\n".join(
        f"- {item}" for item in sector
    )

    return f"""
GEOGRAPHY MASTER TAXONOMY
ONLY valid geography categories:

{geography_text}

SECTOR MASTER TAXONOMY
ONLY valid sector categories:

{sector_text}

Taxonomy rules:
1. Select 1 to 3 geography categories.
2. Select 1 to 3 sector categories.
3. Do not force three categories.
4. Use exact category names.
5. Do not invent, rename, combine or reinterpret category names.
6. Select only categories materially affected by the article.
7. Labels must contain only exact category names.
"""


# ============================================================
# ORIGINAL ARTICLE RETRIEVAL
# ============================================================

def retrieve_original_article(url: str) -> Optional[str]:

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

    for attempt in range(1, ARTICLE_RETRIEVAL_RETRIES + 1):

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

            if extracted and extracted.strip():
                return extracted.strip()[:MAX_ARTICLE_CHARS]

            log(
                f"Article extraction returned no usable text "
                f"(attempt {attempt})."
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
# EXISTING ANALYSIS
# ============================================================

def load_existing_analysis() -> Dict[str, Dict[str, Any]]:

    existing = {}

    if ANALYSIS_HISTORY.exists():

        for path in ANALYSIS_HISTORY.rglob("*.json"):

            try:
                data = load_json(path)
            except Exception as exc:
                log(f"Skipping unreadable history {path}: {exc}")
                continue

            if not isinstance(data, dict):
                continue

            for article in data.get("articles", []):

                if not isinstance(article, dict):
                    continue

                article_id = article.get("id")

                if article_id and isinstance(
                    article.get("analysis"),
                    dict,
                ):
                    existing[str(article_id)] = article

    if ANALYSIS_CURRENT.exists():

        try:
            data = load_json(ANALYSIS_CURRENT)

            if isinstance(data, dict):

                for article in data.get("articles", []):

                    if not isinstance(article, dict):
                        continue

                    article_id = article.get("id")

                    if article_id and isinstance(
                        article.get("analysis"),
                        dict,
                    ):
                        existing[str(article_id)] = article

        except Exception as exc:
            log(f"Unable to read analysis/current.json: {exc}")

    return existing


# ============================================================
# JSON RESPONSE PARSING
# ============================================================

def extract_json_object(text: str) -> Dict[str, Any]:

    text = text.strip()

    try:
        value = json.loads(text)

        if isinstance(value, dict):
            return value

    except json.JSONDecodeError:
        pass

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

    start = text.find("{")

    if start >= 0:

        depth = 0
        in_string = False
        escaped = False

        for index in range(start, len(text)):

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

                    candidate = text[start:index + 1]

                    try:
                        value = json.loads(candidate)

                        if isinstance(value, dict):
                            return value

                    except json.JSONDecodeError:
                        break

    raise ValueError("Gemini response did not contain valid JSON.")


# ============================================================
# VALIDATION
# ============================================================

def require_text(value: Any, field: str) -> str:

    if not isinstance(value, str):
        raise ValueError(f"{field} must be a string.")

    value = value.strip()

    if not value:
        raise ValueError(f"{field} cannot be empty.")

    return value


def validate_categories(
    value: Any,
    allowed: List[str],
    field: str,
) -> List[str]:

    if not isinstance(value, list):
        raise ValueError(f"{field} must be a list.")

    if not 1 <= len(value) <= 3:
        raise ValueError(
            f"{field} must contain between 1 and 3 categories."
        )

    result = []

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


def validate_enum(value: Any, allowed: set, field: str) -> str:

    if not isinstance(value, str):
        raise ValueError(f"{field} must be a string.")

    value = value.strip()

    if value not in allowed:
        raise ValueError(f"Invalid {field}: '{value}'")

    return value


def validate_analysis(
    analysis: Dict[str, Any],
    geography: List[str],
    sector: List[str],
) -> Dict[str, Any]:

    if not isinstance(analysis, dict):
        raise ValueError("Analysis result must be an object.")

    quick = analysis.get("quickRead")
    detailed = analysis.get("detailedAnalysis")

    if not isinstance(quick, dict):
        raise ValueError("Missing quickRead object.")

    if not isinstance(detailed, dict):
        raise ValueError("Missing detailedAnalysis object.")

    return {
        "quickRead": {
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
        },

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

        "detailedAnalysis": {
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
                detailed.get("geographicalImpactExplanation"),
                "detailedAnalysis.geographicalImpactExplanation",
            ),
            "sectorImpactExplanation": require_text(
                detailed.get("sectorImpactExplanation"),
                "detailedAnalysis.sectorImpactExplanation",
            ),
            "timeHorizonExplanation": require_text(
                detailed.get("timeHorizonExplanation"),
                "detailedAnalysis.timeHorizonExplanation",
            ),
            "overallAssessment": require_text(
                detailed.get("overallAssessment"),
                "detailedAnalysis.overallAssessment",
            ),
        },
    }


# ============================================================
# BATCH PROMPT
# ============================================================

def build_batch_prompt(
    articles: List[Dict[str, Any]],
    original_texts: Dict[str, Optional[str]],
    geography: List[str],
    sector: List[str],
) -> str:

    taxonomy = taxonomy_prompt(geography, sector)

    article_sections = []

    for index, article in enumerate(articles, start=1):

        article_id = str(article.get("id", ""))

        summary = str(article.get("summary", "")).strip()

        if len(summary) > MAX_SUMMARY_CHARS:
            summary = summary[:MAX_SUMMARY_CHARS]

        original_text = original_texts.get(article_id)

        if not original_text:
            original_text = (
                "Original article could not be retrieved. "
                "Use the supplied article information only."
            )

        article_sections.append(f"""
---------------- ARTICLE {index} ----------------

ARTICLE ID:
{article_id}

SOURCE:
{article.get("source", "")}

ORIGINAL TITLE:
{article.get("title", "")}

ORIGINAL URL:
{article.get("url", "")}

PUBLISHED SINGAPORE TIME:
{article.get("publishedAtSgt", "")}

COLLECTOR CATEGORY:
{article.get("category", "")}

COLLECTOR RELEVANCE SCORE:
{article.get("relevanceScore", "")}

COLLECTOR RELEVANCE REASON:
{article.get("relevanceReason", "")}

COLLECTOR SUMMARY:
{summary}

ORIGINAL ARTICLE CONTENT:
{original_text}
""")

    articles_text = "\n".join(article_sections)

    return f"""
You are the VGrat FMS market-news analysis engine.

You must independently analyse every supplied article.

This is a batch of separate articles.
Do NOT merge the articles into one combined analysis.
Do NOT allow one article to determine another article's classification.
Each article must receive its own complete analysis.

OBJECTIVITY:
- Be factual, balanced and evidence-based.
- Do not invent facts or unsupported causal relationships.
- Distinguish confirmed developments from possible implications.
- Do not give investment advice.
- Do not recommend individual funds.
- Do not predict individual fund performance.
- Do not exaggerate uncertain consequences.

The collector's category, relevance score and relevance reason
are context only. They must not determine the analysis automatically.

============================================================
TAXONOMY
============================================================

{taxonomy}

============================================================
ANALYSIS REQUIREMENTS FOR EACH ARTICLE
============================================================

Quick Read Summary:
- whatHappened
- whyItMatters
- keyImpact

Geography:
- Select 1 to 3 materially relevant categories.
- Use exact names from the geography taxonomy.

Sector:
- Select 1 to 3 materially relevant categories.
- Use exact names from the sector taxonomy.

Overall Impact Direction:
- Positive
- Negative
- Mixed
- Neutral

This describes the combined economic and financial-market
implication, not an investment recommendation.

Impact Severity:
- Low
- Moderate
- High
- Critical

This describes significance, not certainty.

Time Horizon:
- Immediate
- Short-term
- Medium-term
- Long-term

AI Confidence:
- High
- Medium
- Low

This describes confidence in the analysis and classification,
not certainty about future outcomes.

Detailed Analysis:
- background
- keyDevelopments
- marketImplications
- geographicalImpactExplanation
- sectorImpactExplanation
- timeHorizonExplanation
- overallAssessment

Do not create a separate Evidence section.

============================================================
OUTPUT REQUIREMENTS
============================================================

Return ONLY one valid JSON object.

The root object must contain a "results" array.

There must be exactly one result for every supplied article.

Every result must contain the original article ID exactly.

Use this structure:

{{
  "results": [
    {{
      "id": "ORIGINAL ARTICLE ID",
      "analysis": {{
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
    }}
  ]
}}

CRITICAL:
- Do not omit any article.
- Do not duplicate article IDs.
- Do not change article IDs.
- Do not combine article analyses.
- Do not return commentary outside JSON.
- Every analysis must be complete.
- Return valid JSON only.

============================================================
ARTICLES TO ANALYSE
============================================================

{articles_text}
"""


# ============================================================
# RATE LIMIT DETECTION
# ============================================================

def is_rate_limit_error(exc: Exception) -> bool:

    message = str(exc).lower()

    return (
        "429" in message
        or "too_many_requests" in message
        or "rate limit exceeded" in message
        or "resource_exhausted" in message
    )


# ============================================================
# GEMINI BATCH REQUEST
# ============================================================

def analyze_batch_with_gemini(
    client: genai.Client,
    articles: List[Dict[str, Any]],
    original_texts: Dict[str, Optional[str]],
    geography: List[str],
    sector: List[str],
) -> Dict[str, Dict[str, Any]]:

    prompt = build_batch_prompt(
        articles,
        original_texts,
        geography,
        sector,
    )

    expected_ids = {
        str(article.get("id"))
        for article in articles
    }

    last_error = None

    for attempt in range(1, AI_RETRIES + 1):

        log(
            f"Gemini batch request "
            f"(attempt {attempt}/{AI_RETRIES}; "
            f"{len(articles)} articles)"
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
                raise ValueError("Gemini returned empty output.")

            raw = extract_json_object(output_text)

            results = raw.get("results")

            if not isinstance(results, list):
                raise ValueError(
                    "Gemini batch response has no results array."
                )

            returned = {}

            for item in results:

                if not isinstance(item, dict):
                    log("Ignoring malformed batch result.")
                    continue

                article_id = item.get("id")

                if not isinstance(article_id, str):
                    continue

                article_id = article_id.strip()

                if article_id not in expected_ids:
                    log(
                        f"Ignoring unexpected article ID: "
                        f"{article_id}"
                    )
                    continue

                if article_id in returned:
                    log(
                        f"Duplicate article ID in response: "
                        f"{article_id}"
                    )
                    continue

                returned[article_id] = item.get("analysis")

            if not returned:
                raise ValueError(
                    "Gemini returned no matching article IDs."
                )

            return returned

        except Exception as exc:

            last_error = exc

            if is_rate_limit_error(exc):
                raise

            log(
                f"Batch request failed: {exc}"
            )

            if attempt < AI_RETRIES:
                time.sleep(3)

    raise RuntimeError(
        f"Gemini batch failed after {AI_RETRIES} attempts: "
        f"{last_error}"
    )


# ============================================================
# ANALYSED ARTICLE RECORD
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
        content_sources.append("original_article_url")

    result["analysisMetadata"] = {
        "model": MODEL_NAME,
        "analysedAtSgt": now_sgt_iso(),
        "contentSources": content_sources,
        "batchSize": BATCH_SIZE,
    }

    return result


# ============================================================
# HISTORY STORAGE
# ============================================================

def history_path(article: Dict[str, Any]) -> Path:

    date = publication_date(article)
    year, month, _ = date.split("-")

    return (
        ANALYSIS_HISTORY
        / year
        / month
        / f"{date}.json"
    )


def load_history_file(path: Path) -> Dict[str, Any]:

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
        log(f"Unable to load history file {path}: {exc}")

    return {
        "timezone": "Asia/Singapore",
        "timezoneLabel": "SGT",
        "articles": [],
    }


def upsert_history_article(article: Dict[str, Any]) -> None:

    path = history_path(article)
    data = load_history_file(path)

    existing = data.get("articles", [])

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
            item.get("publishedAtSgt", ""),
            item.get("id", ""),
        )
    )

    data["timezone"] = "Asia/Singapore"
    data["timezoneLabel"] = "SGT"
    data["publicationDateSgt"] = publication_date(article)
    data["articleCount"] = len(existing)
    data["articles"] = existing

    save_json(path, data)


# ============================================================
# CURRENT 14-DAY FILE
# ============================================================

def rebuild_analysis_current(
    existing_articles: Dict[str, Dict[str, Any]],
    news_data: Dict[str, Any],
) -> None:

    articles = []

    cutoff = now_sgt() - timedelta(days=WINDOW_DAYS)

    for article in existing_articles.values():

        if not isinstance(article, dict):
            continue

        published = article.get("publishedAtSgt")

        if not published:
            continue

        try:
            published_dt = parse_sgt_datetime(published)
        except Exception:
            continue

        if published_dt >= cutoff:
            articles.append(article)

    articles.sort(
        key=lambda item: (
            item.get("publishedAtSgt", ""),
            item.get("id", ""),
        ),
        reverse=True,
    )

    max_articles = news_data.get("maxArticles")

    if isinstance(max_articles, int) and max_articles > 0:
        articles = articles[:max_articles]

    output = {
        "generatedAtSgt": now_sgt_iso(),
        "timezone": "Asia/Singapore",
        "timezoneLabel": "SGT",
        "windowDays": WINDOW_DAYS,
        "articleCount": len(articles),
        "source": news_data.get("source", "CNBC"),
        "articles": articles,
    }

    save_json(ANALYSIS_CURRENT, output)


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    log("Starting Gemini Market News Batch Analysis.")

    api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:
        fail("GEMINI_API_KEY environment variable is missing.")

    if not CURRENT_NEWS.exists():
        fail(f"Missing input file: {CURRENT_NEWS}")

    news_data = load_json(CURRENT_NEWS)

    if not isinstance(news_data, dict):
        fail("current.json must contain a JSON object.")

    news_articles = news_data.get("articles", [])

    if not isinstance(news_articles, list):
        fail("current.json articles must be a list.")

    log(f"Loaded {len(news_articles)} articles from current.json.")

    geography, sector = load_taxonomies()

    log(
        f"Loaded {len(geography)} geography categories "
        f"and {len(sector)} sector categories."
    )

    existing = load_existing_analysis()

    log(
        f"Loaded {len(existing)} previously analyzed articles."
    )

    pending = []

    for article in news_articles:

        if not isinstance(article, dict):
            continue

        article_id = article.get("id")

        if not article_id:
            log("Skipping article without ID.")
            continue

        if str(article_id) not in existing:
            pending.append(article)

    log(
        f"New articles requiring Gemini analysis: "
        f"{len(pending)}"
    )

    if not pending:

        log("No new articles require analysis.")

        rebuild_analysis_current(existing, news_data)

        return 0

    batches = [
        pending[index:index + BATCH_SIZE]
        for index in range(0, len(pending), BATCH_SIZE)
    ]

    log(
        f"Batch size: {BATCH_SIZE}; "
        f"total pending batches: {len(batches)}"
    )

    if MAX_BATCHES:
        batches_to_process = batches[:MAX_BATCHES]
        log(
            f"Run limit: processing at most "
            f"{MAX_BATCHES} batches."
        )
    else:
        batches_to_process = batches

    client = genai.Client(api_key=api_key)

    success_count = 0
    failure_count = 0
    request_count = 0
    quota_limited = False
    processed_ids = set()

    for batch_number, batch in enumerate(
        batches_to_process,
        start=1,
    ):

        log("=" * 60)
        log(
            f"Processing batch {batch_number}/"
            f"{len(batches_to_process)}"
        )

        original_texts = {}
        original_retrieved = {}

        for article in batch:

            article_id = str(article.get("id"))

            log(
                f"Retrieving original article: "
                f"{article_id}"
            )

            original_text = retrieve_original_article(
                article.get("url", "")
            )

            original_texts[article_id] = original_text
            original_retrieved[article_id] = bool(original_text)

            if original_text:
                log(f"Original article retrieved: {article_id}")
            else:
                log(
                    f"Using current.json content only: "
                    f"{article_id}"
                )

        try:

            request_count += 1

            returned = analyze_batch_with_gemini(
                client,
                batch,
                original_texts,
                geography,
                sector,
            )

        except Exception as exc:

            if is_rate_limit_error(exc):

                quota_limited = True

                log(
                    "GEMINI RATE LIMIT / QUOTA REACHED. "
                    "Stopping the run immediately. "
                    "No further Gemini requests will be sent."
                )

                log(str(exc))

                failure_count += len(batch)

                break

            failure_count += len(batch)

            log(
                f"Batch {batch_number} failed: {exc}"
            )

            continue

        for article in batch:

            article_id = str(article.get("id"))

            raw_analysis = returned.get(article_id)

            if raw_analysis is None:

                failure_count += 1

                log(
                    f"FAILED article {article_id}: "
                    f"Missing from Gemini batch response."
                )

                continue

            try:

                analysis = validate_analysis(
                    raw_analysis,
                    geography,
                    sector,
                )

                analyzed_article = build_analyzed_article(
                    article,
                    analysis,
                    original_retrieved.get(article_id, False),
                )

                # Save each successful article immediately.
                upsert_history_article(analyzed_article)

                existing[article_id] = analyzed_article
                processed_ids.add(article_id)

                success_count += 1

                log(
                    f"Successfully analyzed and saved "
                    f"{article_id}."
                )

            except Exception as exc:

                failure_count += 1

                log(
                    f"FAILED article {article_id}: {exc}"
                )

        log(
            f"Batch {batch_number} completed. "
            f"Successful so far: {success_count}; "
            f"failed so far: {failure_count}."
        )

    # --------------------------------------------------------
    # REBUILD CURRENT WINDOW
    # --------------------------------------------------------

    rebuild_analysis_current(existing, news_data)

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    unprocessed_count = sum(
        1
        for article in pending
        if str(article.get("id")) not in processed_ids
    )

    log("=" * 60)
    log("BATCH ANALYSIS SUMMARY")
    log("=" * 60)
    log(f"Articles in current.json: {len(news_articles)}")
    log(f"Previously analyzed: {len(existing) - success_count}")
    log(f"New articles pending at start: {len(pending)}")
    log(f"Gemini batch requests attempted: {request_count}")
    log(f"Batch size: {BATCH_SIZE}")
    log(f"Successfully analyzed: {success_count}")
    log(f"Failed articles: {failure_count}")
    log(f"Still pending: {unprocessed_count}")
    log(f"Quota limited: {quota_limited}")
    log(f"Analysis current: {ANALYSIS_CURRENT}")
    log(f"Analysis history: {ANALYSIS_HISTORY}")

    if quota_limited:
        log(
            "Run stopped because of Gemini rate limiting. "
            "Unprocessed articles remain pending."
        )

    log("Batch analysis run finished.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
