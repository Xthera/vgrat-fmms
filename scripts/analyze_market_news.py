#!/usr/bin/env python3

"""
VGrat FMS - MARKET NEWS ANALYSIS
================================

Reads:
    data/market_news/current.json
    Research Funds.xlsx

Produces:
    data/market_news/analysis/current.json
    data/market_news/analysis/history/YYYY/MM/YYYY-MM-DD.json

AI MODEL
========
Qwen/Qwen3-4B-Instruct-2507

DESIGN
======
- Only articles that have not previously been analyzed are sent to the AI.
- Every article in current.json is eventually analyzed.
- Original article identity is preserved.
- Original URL may be retrieved for additional context.
- Geography and sector labels MUST come from Research Funds.xlsx.
- 1-3 geography labels.
- 1-3 sector labels.
- No individual Prudential fund analysis.
- No investment recommendations.
- Historical analysis is stored by Singapore publication date.
- analysis/current.json is a rolling 14-day page.
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from copy import deepcopy
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import requests
import trafilatura
from openpyxl import load_workbook
import torch
from transformers import AutoModelForCausalLM, AutoTokenizer


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
# MODEL
# ============================================================

MODEL_NAME = os.getenv(
    "MARKET_NEWS_MODEL",
    "Qwen/Qwen3-4B-Instruct-2507",
)

MAX_NEW_TOKENS = int(
    os.getenv("MARKET_NEWS_MAX_NEW_TOKENS", "1800")
)

MAX_ARTICLE_CHARS = int(
    os.getenv("MARKET_NEWS_MAX_ARTICLE_CHARS", "24000")
)

REQUEST_TIMEOUT = int(
    os.getenv("MARKET_NEWS_REQUEST_TIMEOUT", "20")
)

AI_RETRIES = int(
    os.getenv("MARKET_NEWS_AI_RETRIES", "3")
)

ARTICLE_RETRIEVAL_RETRIES = int(
    os.getenv("MARKET_NEWS_RETRIEVAL_RETRIES", "2")
)


# ============================================================
# TIMEZONE
# ============================================================

SGT = timezone(timedelta(hours=8))


# ============================================================
# ALLOWED VALUES
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


# ============================================================
# HTTP
# ============================================================

SESSION = requests.Session()

SESSION.headers.update(
    {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/131.0 Safari/537.36 "
            "VGrat-FMS-News-Analysis/1.0"
        ),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml",
    }
)


# ============================================================
# LOGGING
# ============================================================

def log(message: str) -> None:
    print(f"[MARKET-NEWS-ANALYSIS] {message}", flush=True)


# ============================================================
# JSON HELPERS
# ============================================================

def load_json(path: Path, default: Any = None) -> Any:
    if not path.exists():
        return default

    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    temp = path.with_suffix(path.suffix + ".tmp")

    with temp.open("w", encoding="utf-8") as f:
        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2,
        )
        f.write("\n")

    temp.replace(path)


# ============================================================
# SINGAPORE TIME
# ============================================================

def now_sgt() -> datetime:
    return datetime.now(SGT)


def now_sgt_iso() -> str:
    return now_sgt().isoformat(timespec="seconds")


def parse_sgt_datetime(value: str) -> datetime:
    """
    Parses the collector's publishedAtSgt.

    The collector already provides Singapore time, so this function
    does NOT perform another timezone conversion.
    """

    if not value:
        raise ValueError("Empty datetime")

    dt = datetime.fromisoformat(value)

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=SGT)

    return dt.astimezone(SGT)


# ============================================================
# RESEARCH FUNDS TAXONOMY
# ============================================================

def clean_category(value: Any) -> Optional[str]:
    if value is None:
        return None

    text = str(value).strip()

    if not text:
        return None

    return text


def load_master_categories() -> Tuple[List[str], List[str]]:
    """
    Reads:
        worksheet index 1 = Geography master
        worksheet index 2 = Sector master

    The first non-empty row is treated as a header.

    Categories are collected from the first non-empty column.
    """

    if not RESEARCH_FUNDS.exists():
        raise FileNotFoundError(
            f"Required taxonomy workbook not found: {RESEARCH_FUNDS}"
        )

    wb = load_workbook(
        RESEARCH_FUNDS,
        read_only=True,
        data_only=True,
    )

    if len(wb.worksheets) < 3:
        raise RuntimeError(
            "Research Funds.xlsx must contain at least three worksheets: "
            "Fund Research, Geography master and Sector master."
        )

    geography_ws = wb.worksheets[1]
    sector_ws = wb.worksheets[2]

    def read_categories(ws) -> List[str]:
        values: List[str] = []

        for row in ws.iter_rows():
            first_value = None

            for cell in row:
                candidate = clean_category(cell.value)

                if candidate:
                    first_value = candidate
                    break

            if first_value:
                values.append(first_value)

        if not values:
            raise RuntimeError(
                f"No categories found in worksheet '{ws.title}'."
            )

        # Treat the first row as a possible header.
        # If it is clearly a header, remove it.
        header_words = {
            "geography",
            "geographies",
            "region",
            "regions",
            "sector",
            "sectors",
            "industry",
            "industry sector",
            "category",
            "categories",
        }

        if values and values[0].strip().lower() in header_words:
            values = values[1:]

        # Preserve order while removing duplicates.
        result = []
        seen = set()

        for value in values:
            if value not in seen:
                seen.add(value)
                result.append(value)

        return result

    geography = read_categories(geography_ws)
    sector = read_categories(sector_ws)

    wb.close()

    if not geography:
        raise RuntimeError("Geography master list is empty.")

    if not sector:
        raise RuntimeError("Sector master list is empty.")

    log(
        f"Loaded taxonomy: "
        f"{len(geography)} geography categories, "
        f"{len(sector)} sector categories."
    )

    return geography, sector


# ============================================================
# ARTICLE CONTENT RETRIEVAL
# ============================================================

def retrieve_original_article(url: str) -> Optional[str]:
    """
    Retrieves article text from the original URL.

    This is supplemental context only.
    Failure never prevents analysis from using current.json.
    """

    if not url:
        return None

    last_error = None

    for attempt in range(1, ARTICLE_RETRIEVAL_RETRIES + 1):
        try:
            log(
                f"Retrieving original article "
                f"(attempt {attempt}/{ARTICLE_RETRIEVAL_RETRIES}): {url}"
            )

            response = SESSION.get(
                url,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
            )

            response.raise_for_status()

            text = trafilatura.extract(
                response.text,
                include_comments=False,
                include_tables=True,
                favor_precision=True,
                favor_recall=True,
            )

            if text:
                text = text.strip()

                if len(text) > MAX_ARTICLE_CHARS:
                    text = text[:MAX_ARTICLE_CHARS]

                if len(text) >= 300:
                    log(
                        f"Retrieved {len(text):,} characters "
                        "of original article context."
                    )
                    return text

            log("Original page returned insufficient extractable text.")

        except Exception as exc:
            last_error = exc
            log(f"Article retrieval failed: {exc}")

            if attempt < ARTICLE_RETRIEVAL_RETRIES:
                time.sleep(2)

    if last_error:
        log("Falling back to current.json article content.")

    return None


# ============================================================
# ANALYSIS HISTORY
# ============================================================

def iter_history_files() -> List[Path]:
    if not ANALYSIS_HISTORY.exists():
        return []

    return sorted(
        ANALYSIS_HISTORY.glob("*/*/*.json")
    )


def load_existing_analysis() -> Dict[str, Dict[str, Any]]:
    """
    Loads all historical analysis files and indexes them by article ID.
    """

    existing: Dict[str, Dict[str, Any]] = {}

    for path in iter_history_files():
        try:
            data = load_json(path, {})
        except Exception as exc:
            log(f"WARNING: Could not read {path}: {exc}")
            continue

        for article in data.get("articles", []):
            article_id = article.get("id")

            if article_id:
                existing[str(article_id)] = article

    # Also include current.json.
    current = load_json(ANALYSIS_CURRENT, {})

    for article in current.get("articles", []):
        article_id = article.get("id")

        if article_id:
            existing[str(article_id)] = article

    log(
        f"Previously analyzed articles available: {len(existing)}"
    )

    return existing


# ============================================================
# PROMPT
# ============================================================

def build_prompt(
    article: Dict[str, Any],
    article_text: Optional[str],
    geography_categories: List[str],
    sector_categories: List[str],
) -> str:

    title = article.get("title", "")
    summary = article.get("summary", "")
    category = article.get("category", "")
    relevance_score = article.get("relevanceScore", "")
    relevance_reason = article.get("relevanceReason", "")
    published_at = article.get("publishedAtSgt", "")
    source = article.get("source", "")
    url = article.get("url", "")

    supplemental = article_text or summary

    if len(supplemental) > MAX_ARTICLE_CHARS:
        supplemental = supplemental[:MAX_ARTICLE_CHARS]

    geography_json = json.dumps(
        geography_categories,
        ensure_ascii=False,
    )

    sector_json = json.dumps(
        sector_categories,
        ensure_ascii=False,
    )

    return f"""
You are the VGrat FMS Market News Analysis engine.

Analyze the supplied financial/economic/market news article.

IMPORTANT:
- Analyze the article itself.
- Do not make investment recommendations.
- Do not analyze or recommend individual Prudential funds.
- Do not predict election outcomes.
- For political or geopolitical matters, remain factual and neutral.
- Do not invent facts.
- Distinguish what happened from possible implications.
- "Impact" describes potential economic and financial-market implications,
  not an investment signal.
- Severity means potential significance, not certainty.
- Confidence means confidence in your classification and analysis,
  not confidence that the future outcome will occur.

TAXONOMY RULES
==============
Geography labels MUST be selected exactly from this list:

{geography_json}

Sector labels MUST be selected exactly from this list:

{sector_json}

Rules:
- Select 1 to 3 geography labels.
- Select 1 to 3 sector labels.
- Use only exact category strings from the supplied lists.
- Never invent, rename, combine, abbreviate, or paraphrase categories.
- Do not select a category merely because it is indirectly imaginable.
- Select categories with a meaningful connection to the article.
- "Global" may be used when the impact is genuinely global.
- Geography and sector labels are UI tags only.
- Explain the reasoning for the labels separately in the detailed analysis.

OVERALL IMPACT DIRECTION
========================
Allowed values:
- Positive
- Negative
- Mixed
- Neutral

This is the combined potential economic and financial-market direction.

IMPACT SEVERITY
===============
Allowed:
- Low
- Moderate
- High
- Critical

TIME HORIZON
============
Allowed:
- Immediate
- Short-term
- Medium-term
- Long-term

AI CONFIDENCE
=============
Allowed:
- High
- Medium
- Low

QUICK READ
==========
Provide:
- What happened
- Why it matters
- Key impact

DETAILED ANALYSIS
=================
Provide exactly these sections:
- Background
- Key Developments
- Market Implications
- Geographical Impact Explanation
- Sector Impact Explanation
- Time Horizon Explanation
- Overall Assessment

OUTPUT
======
Return ONLY valid JSON.
Do not use Markdown.
Do not use ```json.
Do not include commentary outside the JSON.

Required JSON structure:

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
  "overallImpactDirection": "Positive|Negative|Mixed|Neutral",
  "impactSeverity": "Low|Moderate|High|Critical",
  "timeHorizon": "Immediate|Short-term|Medium-term|Long-term",
  "aiConfidence": "High|Medium|Low",
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

ARTICLE INFORMATION
===================
Source:
{source}

Title:
{title}

Published Singapore time:
{published_at}

URL:
{url}

Collector category:
{category}

Collector relevance score:
{relevance_score}

Collector relevance reason:
{relevance_reason}

Collector summary:
{summary}

ARTICLE CONTENT
===============
{supplemental}
""".strip()


# ============================================================
# JSON EXTRACTION
# ============================================================

def extract_json(text: str) -> Dict[str, Any]:
    """
    Handles:
    - clean JSON
    - accidental Markdown fences
    - surrounding prose
    """

    text = text.strip()

    if text.startswith("```"):
        text = re.sub(
            r"^```(?:json)?\s*",
            "",
            text,
            flags=re.IGNORECASE,
        )

        text = re.sub(
            r"\s*```$",
            "",
            text,
        )

        text = text.strip()

    try:
        result = json.loads(text)

        if not isinstance(result, dict):
            raise ValueError("AI response JSON is not an object.")

        return result

    except json.JSONDecodeError:
        pass

    start = text.find("{")
    end = text.rfind("}")

    if start >= 0 and end > start:
        candidate = text[start:end + 1]

        result = json.loads(candidate)

        if not isinstance(result, dict):
            raise ValueError("Extracted JSON is not an object.")

        return result

    raise ValueError("No valid JSON object found in AI response.")


# ============================================================
# VALIDATION
# ============================================================

def validate_string(value: Any, field: str) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{field} must be a string.")

    value = value.strip()

    if not value:
        raise ValueError(f"{field} cannot be empty.")

    return value


def validate_label_list(
    value: Any,
    field: str,
    allowed: Set[str],
) -> List[str]:

    if not isinstance(value, list):
        raise ValueError(f"{field} must be an array.")

    if not (1 <= len(value) <= 3):
        raise ValueError(
            f"{field} must contain between 1 and 3 labels."
        )

    result = []

    for item in value:
        item = validate_string(item, field)

        if item not in allowed:
            raise ValueError(
                f"Invalid {field} category: '{item}'"
            )

        if item not in result:
            result.append(item)

    if not (1 <= len(result) <= 3):
        raise ValueError(
            f"{field} must contain between 1 and 3 unique labels."
        )

    return result


def validate_analysis(
    analysis: Dict[str, Any],
    geography_categories: List[str],
    sector_categories: List[str],
) -> Dict[str, Any]:

    if not isinstance(analysis, dict):
        raise ValueError("Analysis is not an object.")

    quick = analysis.get("quickRead")

    if not isinstance(quick, dict):
        raise ValueError("quickRead must be an object.")

    quick_read = {
        "whatHappened": validate_string(
            quick.get("whatHappened"),
            "quickRead.whatHappened",
        ),
        "whyItMatters": validate_string(
            quick.get("whyItMatters"),
            "quickRead.whyItMatters",
        ),
        "keyImpact": validate_string(
            quick.get("keyImpact"),
            "quickRead.keyImpact",
        ),
    }

    geography = validate_label_list(
        analysis.get("geographicalImpact"),
        "geographicalImpact",
        set(geography_categories),
    )

    sector = validate_label_list(
        analysis.get("sectorImpact"),
        "sectorImpact",
        set(sector_categories),
    )

    direction = validate_string(
        analysis.get("overallImpactDirection"),
        "overallImpactDirection",
    )

    severity = validate_string(
        analysis.get("impactSeverity"),
        "impactSeverity",
    )

    horizon = validate_string(
        analysis.get("timeHorizon"),
        "timeHorizon",
    )

    confidence = validate_string(
        analysis.get("aiConfidence"),
        "aiConfidence",
    )

    if direction not in ALLOWED_DIRECTION:
        raise ValueError(
            f"Invalid overallImpactDirection: {direction}"
        )

    if severity not in ALLOWED_SEVERITY:
        raise ValueError(
            f"Invalid impactSeverity: {severity}"
        )

    if horizon not in ALLOWED_HORIZON:
        raise ValueError(
            f"Invalid timeHorizon: {horizon}"
        )

    if confidence not in ALLOWED_CONFIDENCE:
        raise ValueError(
            f"Invalid aiConfidence: {confidence}"
        )

    detailed = analysis.get("detailedAnalysis")

    if not isinstance(detailed, dict):
        raise ValueError(
            "detailedAnalysis must be an object."
        )

    detailed_analysis = {
        "background": validate_string(
            detailed.get("background"),
            "detailedAnalysis.background",
        ),
        "keyDevelopments": validate_string(
            detailed.get("keyDevelopments"),
            "detailedAnalysis.keyDevelopments",
        ),
        "marketImplications": validate_string(
            detailed.get("marketImplications"),
            "detailedAnalysis.marketImplications",
        ),
        "geographicalImpactExplanation": validate_string(
            detailed.get("geographicalImpactExplanation"),
            "detailedAnalysis.geographicalImpactExplanation",
        ),
        "sectorImpactExplanation": validate_string(
            detailed.get("sectorImpactExplanation"),
            "detailedAnalysis.sectorImpactExplanation",
        ),
        "timeHorizonExplanation": validate_string(
            detailed.get("timeHorizonExplanation"),
            "detailedAnalysis.timeHorizonExplanation",
        ),
        "overallAssessment": validate_string(
            detailed.get("overallAssessment"),
            "detailedAnalysis.overallAssessment",
        ),
    }

    return {
        "quickRead": quick_read,
        "geographicalImpact": geography,
        "sectorImpact": sector,
        "overallImpactDirection": direction,
        "impactSeverity": severity,
        "timeHorizon": horizon,
        "aiConfidence": confidence,
        "detailedAnalysis": detailed_analysis,
    }


# ============================================================
# MODEL
# ============================================================

class QwenAnalyzer:

    def __init__(self) -> None:
        log(f"Loading AI model: {MODEL_NAME}")

        self.tokenizer = AutoTokenizer.from_pretrained(
            MODEL_NAME,
            trust_remote_code=True,
        )

        # CPU GitHub runner.
        #
        # bfloat16 significantly reduces model memory compared with
        # float32 while remaining supported by modern PyTorch CPUs.
        self.model = AutoModelForCausalLM.from_pretrained(
            MODEL_NAME,
            torch_dtype=torch.bfloat16,
            low_cpu_mem_usage=True,
            device_map="auto",
            trust_remote_code=True,
        )

        self.model.eval()

        log(
            f"Model loaded. Device: {self.model.device}"
        )

    def generate(
        self,
        prompt: str,
    ) -> str:

        messages = [
            {
                "role": "system",
                "content": (
                    "You are a precise financial news analysis engine. "
                    "Return only valid JSON when JSON is requested."
                ),
            },
            {
                "role": "user",
                "content": prompt,
            },
        ]

        text = self.tokenizer.apply_chat_template(
            messages,
            tokenize=False,
            add_generation_prompt=True,
        )

        inputs = self.tokenizer(
            [text],
            return_tensors="pt",
        )

        inputs = {
            key: value.to(self.model.device)
            for key, value in inputs.items()
        }

        with torch.inference_mode():
            output_ids = self.model.generate(
                **inputs,
                max_new_tokens=MAX_NEW_TOKENS,
                do_sample=False,
                temperature=None,
                top_p=None,
                use_cache=True,
            )

        generated_ids = output_ids[
            0,
            inputs["input_ids"].shape[-1]:,
        ]

        output = self.tokenizer.decode(
            generated_ids,
            skip_special_tokens=True,
        )

        return output.strip()


# ============================================================
# ARTICLE ANALYSIS
# ============================================================

def analyze_article(
    analyzer: QwenAnalyzer,
    article: Dict[str, Any],
    geography_categories: List[str],
    sector_categories: List[str],
) -> Dict[str, Any]:

    article_id = str(article.get("id", ""))

    if not article_id:
        raise ValueError("Article has no stable ID.")

    article_text = retrieve_original_article(
        article.get("url", "")
    )

    prompt = build_prompt(
        article,
        article_text,
        geography_categories,
        sector_categories,
    )

    last_error: Optional[Exception] = None

    for attempt in range(1, AI_RETRIES + 1):
        log(
            f"AI analysis for {article_id[:12]}... "
            f"(attempt {attempt}/{AI_RETRIES})"
        )

        try:
            raw = analyzer.generate(prompt)

            parsed = extract_json(raw)

            validated = validate_analysis(
                parsed,
                geography_categories,
                sector_categories,
            )

            log(
                f"Analysis validated successfully for "
                f"{article_id[:12]}..."
            )

            return validated

        except Exception as exc:
            last_error = exc

            log(
                f"AI validation failed for "
                f"{article_id[:12]}...: {exc}"
            )

            if attempt < AI_RETRIES:
                time.sleep(2)

    raise RuntimeError(
        f"AI analysis failed after {AI_RETRIES} attempts "
        f"for article {article_id}: {last_error}"
    )


# ============================================================
# ARTICLE OUTPUT
# ============================================================

def build_analyzed_article(
    original: Dict[str, Any],
    analysis: Dict[str, Any],
) -> Dict[str, Any]:

    result = deepcopy(original)

    # Preserve every original collector field exactly.

    result["analysis"] = analysis

    result["analysisMetadata"] = {
        "model": MODEL_NAME,
        "analysedAtSgt": now_sgt_iso(),
        "contentSources": [
            "market_news/current.json",
            "original_article_url",
        ],
    }

    return result


# ============================================================
# HISTORY HELPERS
# ============================================================

def history_path_for_article(article: Dict[str, Any]) -> Path:
    published = parse_sgt_datetime(
        article["publishedAtSgt"]
    )

    year = published.strftime("%Y")
    month = published.strftime("%m")
    date = published.strftime("%Y-%m-%d")

    return (
        ANALYSIS_HISTORY
        / year
        / month
        / f"{date}.json"
    )


def load_history_day(path: Path) -> Dict[str, Any]:
    if path.exists():
        data = load_json(path, {})

        if isinstance(data, dict):
            return data

    return {
        "generatedAtSgt": now_sgt_iso(),
        "timezone": "Asia/Singapore",
        "timezoneLabel": "SGT",
        "source": "CNBC",
        "publicationDateSgt": path.stem,
        "articles": [],
    }


def upsert_history_article(
    article: Dict[str, Any],
) -> None:

    path = history_path_for_article(article)

    data = load_history_day(path)

    articles = data.get("articles", [])

    article_id = article["id"]

    replaced = False

    for index, existing in enumerate(articles):
        if existing.get("id") == article_id:
            articles[index] = article
            replaced = True
            break

    if not replaced:
        articles.append(article)

    # Stable chronological ordering.
    articles.sort(
        key=lambda item: item.get(
            "publishedAtSgt",
            "",
        )
    )

    data["generatedAtSgt"] = now_sgt_iso()
    data["articleCount"] = len(articles)
    data["articles"] = articles

    save_json(path, data)

    log(
        f"{'Updated' if replaced else 'Created'} history: "
        f"{path}"
    )


# ============================================================
# CURRENT ANALYSIS
# ============================================================

def build_current_analysis(
    current_news: Dict[str, Any],
    analyzed_by_id: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:

    generated = current_news.get(
        "generatedAtSgt",
        now_sgt_iso(),
    )

    current_articles = []

    cutoff = now_sgt() - timedelta(days=14)

    for article in current_news.get("articles", []):

        article_id = article.get("id")

        if not article_id:
            continue

        analyzed = analyzed_by_id.get(str(article_id))

        if not analyzed:
            continue

        try:
            published = parse_sgt_datetime(
                analyzed["publishedAtSgt"]
            )
        except Exception:
            continue

        if published < cutoff:
            continue

        current_articles.append(analyzed)

    # Match the collector's article order where possible.
    current_order = {
        str(article.get("id")): index
        for index, article in enumerate(
            current_news.get("articles", [])
        )
    }

    current_articles.sort(
        key=lambda article: current_order.get(
            str(article.get("id")),
            999999,
        )
    )

    return {
        "generatedAtSgt": generated,
        "analysedAtSgt": now_sgt_iso(),
        "timezone": "Asia/Singapore",
        "timezoneLabel": "SGT",
        "windowDays": 14,
        "source": current_news.get(
            "source",
            "CNBC",
        ),
        "analysisModel": MODEL_NAME,
        "articleCount": len(current_articles),
        "articles": current_articles,
    }


# ============================================================
# VALIDATION OF ORIGINAL CURRENT.JSON
# ============================================================

def validate_current_news(data: Dict[str, Any]) -> None:
    if not isinstance(data, dict):
        raise ValueError(
            "market_news/current.json must contain an object."
        )

    articles = data.get("articles")

    if not isinstance(articles, list):
        raise ValueError(
            "market_news/current.json has no articles array."
        )

    seen = set()

    for article in articles:
        article_id = article.get("id")

        if not article_id:
            raise ValueError(
                "Article without ID found in current.json."
            )

        if article_id in seen:
            raise ValueError(
                f"Duplicate article ID in current.json: {article_id}"
            )

        seen.add(article_id)


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    log("Starting Market News Analysis.")

    if not CURRENT_NEWS.exists():
        log(
            f"ERROR: Missing input file: {CURRENT_NEWS}"
        )
        return 1

    current_news = load_json(
        CURRENT_NEWS,
        {},
    )

    validate_current_news(current_news)

    geography_categories, sector_categories = (
        load_master_categories()
    )

    existing = load_existing_analysis()

    articles = current_news.get("articles", [])

    pending = [
        article
        for article in articles
        if str(article.get("id")) not in existing
    ]

    log(
        f"Articles currently in collector: {len(articles)}"
    )

    log(
        f"Already analyzed: {len(existing)}"
    )

    log(
        f"New articles requiring analysis: {len(pending)}"
    )

    # Nothing new:
    # still rebuild current.json so its rolling 14-day window remains
    # synchronized with the collector.
    if not pending:
        current_analysis = build_current_analysis(
            current_news,
            existing,
        )

        save_json(
            ANALYSIS_CURRENT,
            current_analysis,
        )

        log(
            "No new articles. Analysis skipped."
        )

        log(
            f"Current analysis contains "
            f"{current_analysis['articleCount']} articles."
        )

        return 0

    analyzer = QwenAnalyzer()

    success_count = 0
    failure_count = 0

    for index, article in enumerate(
        pending,
        start=1,
    ):

        article_id = str(article.get("id"))

        log(
            f"=================================================="
        )

        log(
            f"Processing article {index}/{len(pending)}"
        )

        log(
            f"ID: {article_id}"
        )

        log(
            f"Title: {article.get('title', '')}"
        )

        try:
            analysis = analyze_article(
                analyzer,
                article,
                geography_categories,
                sector_categories,
            )

            analyzed_article = build_analyzed_article(
                article,
                analysis,
            )

            existing[article_id] = analyzed_article

            upsert_history_article(
                analyzed_article
            )

            success_count += 1

        except Exception as exc:
            failure_count += 1

            log(
                f"ERROR analyzing article "
                f"{article_id}: {exc}"
            )

            # Do NOT write an incomplete analysis record.
            #
            # This means the article remains pending and will be
            # retried on the next workflow run.
            continue

    # ========================================================
    # BUILD CURRENT ANALYSIS
    # ========================================================

    current_analysis = build_current_analysis(
        current_news,
        existing,
    )

    save_json(
        ANALYSIS_CURRENT,
        current_analysis,
    )

    # ========================================================
    # FINAL VALIDATION
    # ========================================================

    written = load_json(
        ANALYSIS_CURRENT,
        {},
    )

    if not isinstance(
        written,
        dict,
    ):
        raise RuntimeError(
            "analysis/current.json is not a JSON object."
        )

    if not isinstance(
        written.get("articles"),
        list,
    ):
        raise RuntimeError(
            "analysis/current.json has no articles array."
        )

    log(
        "=================================================="
    )

    log(
        f"Analysis complete."
    )

    log(
        f"Successful new analyses: {success_count}"
    )

    log(
        f"Failed new analyses: {failure_count}"
    )

    log(
        f"Current analysis article count: "
        f"{written.get('articleCount', 0)}"
    )

    if failure_count > 0:
        log(
            "WARNING: Some articles failed analysis and "
            "will be retried on the next run."
        )

    # Do not fail the entire workflow merely because one article
    # failed. Successful analyses and history are still persisted.
    return 0


if __name__ == "__main__":
    sys.exit(main())
