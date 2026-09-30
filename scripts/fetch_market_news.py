#!/usr/bin/env python3

"""
VGrat FMS - CNA Market News Collector

PURPOSE
=======

Collect relevant CNA market/economic/geopolitical news.

Pipeline:

    CNA RSS
        |
        v
    deduplicate
        |
        v
    relevance filter
        |
        v
    fetch article
        |
        v
    extract article text
        |
        v
    generate summary
        |
        v
    permanent historical archive
        |
        v
    rolling 14-day current.json

IMPORTANT
=========

The RSS feed is broad and contains many non-market stories.

Therefore:

    RSS title + description
            |
            v
      relevance filter
            |
       relevant only
            |
            v
      article extraction

This prevents unnecessary requests for irrelevant articles.

Permanent storage contains:

    - title
    - source
    - publication date
    - URL
    - category
    - relevance score
    - summary
    - extracted timestamp

It does NOT permanently store the full third-party article text.
"""

from __future__ import annotations

import hashlib
import json
import re
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit, urlunsplit

import feedparser
import requests
import trafilatura
from bs4 import BeautifulSoup


# ============================================================
# CONFIGURATION
# ============================================================

SOURCE_NAME = "CNA"

BASE_DIR = Path(__file__).resolve().parent.parent
NEWS_DIR = BASE_DIR / "data" / "market_news"
HISTORY_DIR = NEWS_DIR / "history"
CURRENT_FILE = NEWS_DIR / "current.json"

WINDOW_DAYS = 14

REQUEST_TIMEOUT = 25

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/154.0.0.0 Safari/537.36"
)

CNA_FEEDS = {
    "latest": "https://www.channelnewsasia.com/api/v1/rss-outbound-feed?_format=xml",
    "asia": "https://www.channelnewsasia.com/api/v1/rss-outbound-feed?_format=xml&category=6511",
    "business": "https://www.channelnewsasia.com/api/v1/rss-outbound-feed?_format=xml&category=6936",
    "singapore": "https://www.channelnewsasia.com/api/v1/rss-outbound-feed?_format=xml&category=10416",
    "world": "https://www.channelnewsasia.com/api/v1/rss-outbound-feed?_format=xml&category=6311",
}


# ============================================================
# RELEVANCE KEYWORDS
# ============================================================

# Strong market / financial signals.
MARKET_KEYWORDS = {
    "stocks": 5,
    "stock": 5,
    "shares": 5,
    "share price": 5,
    "equities": 5,
    "equity": 5,
    "markets": 5,
    "market": 4,
    "investors": 4,
    "investor": 4,
    "investment": 4,
    "fund": 3,
    "funds": 3,
    "asset management": 5,
    "portfolio": 4,
    "ipo": 5,
    "listing": 4,
    "listed company": 5,
    "earnings": 5,
    "profit": 4,
    "profits": 4,
    "revenue": 4,
    "dividend": 4,
    "dividends": 4,
    "merger": 5,
    "acquisition": 5,
    "takeover": 5,
    "bank": 4,
    "banks": 4,
    "banking": 5,
    "financial institution": 5,
    "financial institutions": 5,
    "insurer": 4,
    "insurance": 4,
    "credit": 4,
    "loan": 3,
    "loans": 3,
    "bond": 5,
    "bonds": 5,
    "yield": 5,
    "yields": 5,
    "treasury": 4,
    "interest rate": 6,
    "interest rates": 6,
    "central bank": 6,
    "central banks": 6,
    "monetary policy": 6,
    "rate cut": 6,
    "rate cuts": 6,
    "rate hike": 6,
    "rate hikes": 6,
    "inflation": 6,
    "cpi": 6,
    "ppi": 6,
    "gdp": 6,
    "economic growth": 6,
    "recession": 6,
    "employment": 4,
    "unemployment": 5,
    "jobs data": 5,
    "nonfarm payroll": 6,
    "payrolls": 5,
    "consumer spending": 5,
    "consumer prices": 5,
    "producer prices": 5,
    "retail sales": 5,
    "industrial production": 5,
    "factory output": 5,
    "manufacturing": 4,
    "manufacturing output": 5,
    "pmi": 6,
    "trade": 4,
    "exports": 5,
    "imports": 5,
    "tariff": 6,
    "tariffs": 6,
    "trade war": 7,
    "supply chain": 5,
    "currency": 5,
    "currencies": 5,
    "forex": 6,
    "fx": 5,
    "dollar": 4,
    "yuan": 4,
    "renminbi": 4,
    "yen": 4,
    "euro": 4,
    "sterling": 4,
    "commodity": 5,
    "commodities": 5,
    "oil": 5,
    "crude": 5,
    "gold": 5,
    "copper": 4,
    "natural gas": 5,
    "energy prices": 6,
    "property market": 5,
    "property prices": 5,
    "housing market": 5,
    "real estate": 5,
    "reit": 6,
    "reits": 6,
    "data centre": 4,
    "data center": 4,
    "technology": 3,
    "technology company": 4,
    "semiconductor": 5,
    "semiconductors": 5,
    "chip": 4,
    "chips": 4,
    "artificial intelligence": 5,
    "ai": 3,
    "deepseek": 5,
    "nvidia": 5,
    "apple": 4,
    "amazon": 4,
    "microsoft": 4,
    "google": 4,
    "meta": 4,
    "tesla": 4,
}

# Macro / policy terms.
ECONOMIC_KEYWORDS = {
    "economy": 5,
    "economic": 5,
    "economics": 5,
    "fiscal policy": 6,
    "government spending": 5,
    "budget": 5,
    "tax": 4,
    "taxes": 4,
    "tax policy": 5,
    "subsidy": 4,
    "subsidies": 4,
    "stimulus": 6,
    "economic policy": 6,
    "business confidence": 5,
    "consumer confidence": 5,
    "cost of living": 4,
    "electricity tariffs": 5,
    "gas tariffs": 5,
    "utility tariffs": 5,
    "household tariffs": 4,
    "interest": 3,
}

# Geopolitical terms that can have substantial market/economic impact.
GEOPOLITICAL_KEYWORDS = {
    "war": 4,
    "conflict": 4,
    "ceasefire": 5,
    "sanctions": 6,
    "sanction": 6,
    "geopolitical": 6,
    "military escalation": 6,
    "trade restrictions": 6,
    "export controls": 6,
    "import restrictions": 5,
    "strategic minerals": 5,
    "rare earth": 5,
    "rare earths": 5,
    "shipping disruption": 6,
    "shipping disruptions": 6,
    "strait": 3,
    "taiwan strait": 6,
    "south china sea": 5,
    "ukraine": 4,
    "russia": 4,
    "china": 3,
    "united states": 3,
    "us economy": 6,
    "european union": 4,
    "europe": 3,
}

# Strong negative signals.
# These are deliberately strong because these categories are
# generally outside a market-news dashboard.
EXCLUSION_KEYWORDS = {
    # Sports
    "football": -10,
    "soccer": -10,
    "rugby": -10,
    "cricket": -10,
    "tennis": -10,
    "golf": -10,
    "basketball": -10,
    "baseball": -10,
    "formula 1": -10,
    "f1": -10,
    "premier league": -10,
    "champions league": -10,
    "world cup": -10,
    "olympics": -10,
    "olympic": -10,
    "athlete": -8,
    "athletes": -8,
    "coach": -8,
    "match": -8,
    "matches": -8,
    "tournament": -8,
    "player": -8,
    "players": -8,
    "team": -7,
    "teams": -7,

    # Entertainment
    "movie": -8,
    "movies": -8,
    "film": -8,
    "films": -8,
    "actor": -8,
    "actress": -8,
    "singer": -8,
    "concert": -8,
    "celebrity": -9,
    "celebrity": -9,
    "music": -7,
    "album": -7,
    "television": -7,
    "tv show": -7,
    "reality show": -8,

    # Food / lifestyle
    "restaurant": -8,
    "restaurants": -8,
    "bakery": -8,
    "cafe": -8,
    "food": -5,
    "recipe": -8,
    "recipes": -8,
    "travel": -6,
    "tourism": -5,
    "holiday": -5,
    "fashion": -7,
    "lifestyle": -7,
    "beauty": -7,
    "wellness": -7,

    # Crime / accidents / general incidents
    "murder": -10,
    "murdered": -10,
    "homicide": -10,
    "robbery": -10,
    "burglary": -10,
    "assault": -9,
    "arrested": -8,
    "arrest": -8,
    "charged": -7,
    "jail": -8,
    "prison": -8,
    "court": -5,
    "crime": -9,
    "criminal": -8,
    "police": -6,
    "accident": -9,
    "crash": -8,
    "collision": -8,
    "fire": -7,
    "died": -8,
    "dies": -8,
    "death": -8,
    "killed": -9,
    "injured": -8,
    "injury": -8,
    "missing person": -9,
    "rescue": -7,

    # General human-interest
    "viral": -8,
    "social media star": -8,
    "influencer": -7,
    "viral video": -8,
    "pets": -8,
    "pet": -8,
    "birthday": -8,
}


# ============================================================
# HELPERS
# ============================================================

def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def iso_utc(dt: datetime | None = None) -> str:
    if dt is None:
        dt = utc_now()

    return (
        dt.astimezone(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )


def clean_text(value: Any) -> str:
    if value is None:
        return ""

    text = str(value)

    text = BeautifulSoup(text, "html.parser").get_text(" ", strip=True)

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def canonicalize_url(url: str) -> str:
    """
    Normalize URLs for deduplication.

    Tracking query parameters are removed.
    """
    if not url:
        return ""

    try:
        parts = urlsplit(url.strip())

        query_parts = []

        for item in parts.query.split("&"):
            if not item:
                continue

            key = item.split("=", 1)[0].lower()

            if key.startswith("utm_"):
                continue

            if key in {
                "fbclid",
                "gclid",
                "mc_cid",
                "mc_eid",
            }:
                continue

            query_parts.append(item)

        normalized_query = "&".join(query_parts)

        return urlunsplit(
            (
                parts.scheme.lower(),
                parts.netloc.lower(),
                parts.path.rstrip("/"),
                normalized_query,
                "",
            )
        )

    except Exception:
        return url.strip()


def article_id(source: str, url: str) -> str:
    raw = f"{source}|{canonicalize_url(url)}".encode("utf-8")

    return hashlib.sha256(raw).hexdigest()


def parse_entry_datetime(entry: Any) -> datetime:
    """
    Parse RSS publication date.

    feedparser exposes published_parsed / updated_parsed.
    """
    parsed = getattr(entry, "published_parsed", None)

    if parsed is None:
        parsed = getattr(entry, "updated_parsed", None)

    if parsed is not None:
        try:
            from calendar import timegm

            timestamp = timegm(parsed)

            return datetime.fromtimestamp(
                timestamp,
                tz=timezone.utc,
            )
        except Exception:
            pass

    return utc_now()


def shorten(text: str, max_chars: int = 320) -> str:
    text = clean_text(text)

    if len(text) <= max_chars:
        return text

    truncated = text[:max_chars].rsplit(" ", 1)[0]

    return truncated.rstrip(" .,;:") + "..."


# ============================================================
# RELEVANCE CLASSIFIER
# ============================================================

def score_keywords(
    text: str,
    keywords: dict[str, int],
) -> tuple[int, list[str]]:
    """
    Score keyword occurrences.

    Each keyword is counted at most once.
    """
    lower = text.lower()

    score = 0
    matches: list[str] = []

    for keyword, weight in keywords.items():
        if keyword.lower() in lower:
            score += weight
            matches.append(keyword)

    return score, matches


def classify_relevance(
    title: str,
    description: str,
) -> dict[str, Any]:
    """
    Stage-1 relevance filter.

    Only title + RSS description are used.

    No article request is made before this function passes
    the article.
    """

    text = f"{title}. {description}"

    market_score, market_matches = score_keywords(
        text,
        MARKET_KEYWORDS,
    )

    economic_score, economic_matches = score_keywords(
        text,
        ECONOMIC_KEYWORDS,
    )

    geopolitical_score, geopolitical_matches = score_keywords(
        text,
        GEOPOLITICAL_KEYWORDS,
    )

    exclusion_score, exclusion_matches = score_keywords(
        text,
        EXCLUSION_KEYWORDS,
    )

    total_score = (
        market_score
        + economic_score
        + geopolitical_score
        + exclusion_score
    )

    positive_score = (
        market_score
        + economic_score
        + geopolitical_score
    )

    # Strong exclusions normally reject an article even if
    # a broad word such as "market" appears in the description.
    strong_exclusion = exclusion_score <= -8

    if strong_exclusion and positive_score < 10:
        relevant = False

    else:
        relevant = total_score >= 5

    if market_score >= 5:
        category = "MARKET"

    elif economic_score >= 5:
        category = "ECONOMIC"

    elif geopolitical_score >= 5:
        category = "GEOPOLITICAL"

    else:
        category = "GENERAL"

    return {
        "relevant": relevant,
        "category": category,
        "score": total_score,
        "positiveScore": positive_score,
        "marketScore": market_score,
        "economicScore": economic_score,
        "geopoliticalScore": geopolitical_score,
        "exclusionScore": exclusion_score,
        "matchedKeywords": (
            market_matches
            + economic_matches
            + geopolitical_matches
        ),
        "excludedKeywords": exclusion_matches,
    }


# ============================================================
# RSS COLLECTION
# ============================================================

def fetch_feed(
    feed_name: str,
    feed_url: str,
) -> list[dict[str, Any]]:
    print(f"\nFetching CNA feed: {feed_name}")

    try:
        response = requests.get(
            feed_url,
            headers={
                "User-Agent": USER_AGENT,
                "Accept": "application/rss+xml, application/xml, text/xml",
            },
            timeout=REQUEST_TIMEOUT,
        )

        response.raise_for_status()

        parsed = feedparser.parse(response.content)

        records: list[dict[str, Any]] = []

        for entry in parsed.entries:
            title = clean_text(
                getattr(entry, "title", "")
            )

            description = clean_text(
                getattr(
                    entry,
                    "summary",
                    getattr(entry, "description", ""),
                )
            )

            url = clean_text(
                getattr(entry, "link", "")
            )

            if not title or not url:
                continue

            published_dt = parse_entry_datetime(entry)

            records.append(
                {
                    "feed": feed_name,
                    "title": title,
                    "description": description,
                    "url": canonicalize_url(url),
                    "publishedAtUtc": iso_utc(published_dt),
                }
            )

        print(f"  Entries: {len(records)}")

        return records

    except Exception as exc:
        print(
            f"  ERROR fetching {feed_name}: "
            f"{type(exc).__name__}: {exc}"
        )

        return []


def collect_rss() -> list[dict[str, Any]]:
    all_records: list[dict[str, Any]] = []

    for feed_name, feed_url in CNA_FEEDS.items():
        records = fetch_feed(
            feed_name,
            feed_url,
        )

        all_records.extend(records)

    return all_records


# ============================================================
# ARTICLE EXTRACTION
# ============================================================

def fetch_article_text(url: str) -> str:
    """
    Temporarily retrieve article text.

    The extracted full article is NOT stored permanently.
    """

    response = requests.get(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml",
        },
        timeout=REQUEST_TIMEOUT,
    )

    response.raise_for_status()

    html = response.text

    extracted = trafilatura.extract(
        html,
        include_comments=False,
        include_tables=False,
        favor_precision=True,
    )

    if extracted:
        return clean_text(extracted)

    # Fallback extraction.
    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    for tag in soup(
        [
            "script",
            "style",
            "noscript",
            "nav",
            "footer",
            "header",
        ]
    ):
        tag.decompose()

    paragraphs = []

    for paragraph in soup.find_all("p"):
        text = clean_text(
            paragraph.get_text(" ", strip=True)
        )

        if len(text) >= 40:
            paragraphs.append(text)

    return clean_text(" ".join(paragraphs))


# ============================================================
# SUMMARY
# ============================================================

def split_sentences(text: str) -> list[str]:
    text = clean_text(text)

    if not text:
        return []

    sentences = re.split(
        r"(?<=[.!?])\s+",
        text,
    )

    return [
        sentence.strip()
        for sentence in sentences
        if len(sentence.strip()) >= 45
    ]


def generate_summary(
    title: str,
    description: str,
    article_text: str,
) -> str:
    """
    Temporary extractive summarizer.

    This is intentionally local and deterministic.

    A local LLM can replace this later without changing
    the permanent archive structure.
    """

    candidates = split_sentences(article_text)

    if not candidates:
        description = clean_text(description)

        if description:
            return shorten(
                description,
                500,
            )

        return shorten(
            title,
            500,
        )

    # Prefer sentences that contain useful market terms.
    priority_terms = [
        "market",
        "stock",
        "shares",
        "investor",
        "economy",
        "economic",
        "growth",
        "inflation",
        "interest rate",
        "central bank",
        "bank",
        "bond",
        "yield",
        "currency",
        "trade",
        "tariff",
        "oil",
        "energy",
        "technology",
        "semiconductor",
        "chip",
        "ai",
        "gdp",
        "manufacturing",
        "exports",
        "imports",
    ]

    scored: list[tuple[int, int, str]] = []

    for index, sentence in enumerate(candidates):
        lower = sentence.lower()

        relevance = sum(
            1
            for term in priority_terms
            if term in lower
        )

        scored.append(
            (
                relevance,
                -index,
                sentence,
            )
        )

    scored.sort(
        reverse=True,
        key=lambda item: (
            item[0],
            item[1],
        ),
    )

    selected = []

    # Keep up to three useful sentences.
    for _, _, sentence in scored[:3]:
        selected.append(sentence)

    # Preserve article order.
    selected_set = set(selected)

    ordered = [
        sentence
        for sentence in candidates
        if sentence in selected_set
    ][:3]

    summary = " ".join(ordered)

    return shorten(
        summary,
        700,
    )


# ============================================================
# HISTORY LOADING
# ============================================================

def load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default

    try:
        with path.open(
            "r",
            encoding="utf-8",
        ) as handle:
            return json.load(handle)

    except Exception as exc:
        print(
            f"WARNING: Could not read {path}: "
            f"{type(exc).__name__}: {exc}"
        )

        return default


def load_existing_history_ids() -> set[str]:
    ids: set[str] = set()

    if not HISTORY_DIR.exists():
        return ids

    for path in HISTORY_DIR.rglob("*.json"):
        data = load_json(
            path,
            {},
        )

        articles = data.get(
            "articles",
            [],
        )

        for article in articles:
            value = article.get("id")

            if value:
                ids.add(value)

    return ids


def load_all_history_articles() -> list[dict[str, Any]]:
    articles: list[dict[str, Any]] = []

    if not HISTORY_DIR.exists():
        return articles

    for path in HISTORY_DIR.rglob("*.json"):
        data = load_json(
            path,
            {},
        )

        for article in data.get(
            "articles",
            [],
        ):
            if isinstance(article, dict):
                articles.append(article)

    return articles


# ============================================================
# HISTORY WRITING
# ============================================================

def history_file_for_date(
    published_at_utc: str,
) -> Path:
    try:
        dt = datetime.fromisoformat(
            published_at_utc.replace(
                "Z",
                "+00:00",
            )
        )
    except Exception:
        dt = utc_now()

    directory = (
        HISTORY_DIR
        / f"{dt.year:04d}"
        / f"{dt.month:02d}"
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    return directory / f"{dt.date().isoformat()}.json"


def save_history_article(
    article: dict[str, Any],
) -> bool:
    path = history_file_for_date(
        article["publishedAtUtc"]
    )

    data = load_json(
        path,
        {
            "date": article["publishedAtUtc"][:10],
            "source": SOURCE_NAME,
            "articles": [],
        },
    )

    articles = data.setdefault(
        "articles",
        [],
    )

    existing_ids = {
        item.get("id")
        for item in articles
        if isinstance(item, dict)
    }

    if article["id"] in existing_ids:
        return False

    articles.append(article)

    articles.sort(
        key=lambda item: item.get(
            "publishedAtUtc",
            "",
        ),
        reverse=True,
    )

    data["articleCount"] = len(articles)

    data["updatedAtUtc"] = iso_utc()

    with path.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            data,
            handle,
            indent=2,
            ensure_ascii=False,
        )

        handle.write("\n")

    return True


# ============================================================
# CURRENT 14-DAY INDEX
# ============================================================

def build_current_index(
    all_history_articles: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    cutoff = utc_now() - timedelta(
        days=WINDOW_DAYS
    )

    current: list[dict[str, Any]] = []

    for article in all_history_articles:
        published = article.get(
            "publishedAtUtc",
            "",
        )

        try:
            published_dt = datetime.fromisoformat(
                published.replace(
                    "Z",
                    "+00:00",
                )
            )
        except Exception:
            continue

        if published_dt >= cutoff:
            current.append(article)

    # Deduplicate one final time.
    unique: dict[str, dict[str, Any]] = {}

    for article in current:
        article_id_value = article.get("id")

        if not article_id_value:
            continue

        unique[article_id_value] = article

    current = list(unique.values())

    current.sort(
        key=lambda item: item.get(
            "publishedAtUtc",
            "",
        ),
        reverse=True,
    )

    return current


def write_current_index(
    articles: list[dict[str, Any]],
) -> None:
    NEWS_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    data = {
        "generatedAtUtc": iso_utc(),
        "windowDays": WINDOW_DAYS,
        "source": SOURCE_NAME,
        "articleCount": len(articles),
        "articles": articles,
    }

    with CURRENT_FILE.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            data,
            handle,
            indent=2,
            ensure_ascii=False,
        )

        handle.write("\n")


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    print("=" * 72)
    print("VGrat FMS - CNA Market News Collector")
    print("=" * 72)

    NEWS_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    HISTORY_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    run_started = utc_now()

    stats = {
        "rssEntries": 0,
        "duplicates": 0,
        "rejectedIrrelevant": 0,
        "relevantArticles": 0,
        "articleExtractionAttempts": 0,
        "articleExtractionSuccess": 0,
        "articleExtractionFailures": 0,
        "newHistoricalRecords": 0,
    }

    # --------------------------------------------------------
    # 1. Collect RSS
    # --------------------------------------------------------

    rss_records = collect_rss()

    stats["rssEntries"] = len(rss_records)

    print(
        f"\nRSS entries collected: "
        f"{stats['rssEntries']}"
    )

    # --------------------------------------------------------
    # 2. Deduplicate RSS records
    # --------------------------------------------------------

    unique_records: dict[str, dict[str, Any]] = {}

    for record in rss_records:
        url = record.get(
            "url",
            "",
        )

        if not url:
            continue

        record_id = article_id(
            SOURCE_NAME,
            url,
        )

        if record_id in unique_records:
            stats["duplicates"] += 1

            # Preserve all feed names.
            existing = unique_records[record_id]

            existing_feed = existing.get(
                "feed"
            )

            feeds = existing.setdefault(
                "feeds",
                [],
            )

            if existing_feed and existing_feed not in feeds:
                feeds.append(existing_feed)

            current_feed = record.get(
                "feed"
            )

            if current_feed and current_feed not in feeds:
                feeds.append(current_feed)

            continue

        record["id"] = record_id

        record["feeds"] = [
            record.get(
                "feed",
                "unknown",
            )
        ]

        unique_records[record_id] = record

    print(
        f"Unique RSS entries: "
        f"{len(unique_records)}"
    )

    print(
        f"RSS duplicates: "
        f"{stats['duplicates']}"
    )

    # --------------------------------------------------------
    # 3. Relevance filtering
    # --------------------------------------------------------

    relevant_records: list[dict[str, Any]] = []

    print("\nRelevance filtering:")

    for record in unique_records.values():
        relevance = classify_relevance(
            record["title"],
            record["description"],
        )

        record["relevance"] = relevance

        if not relevance["relevant"]:
            stats["rejectedIrrelevant"] += 1

            print(
                "  REJECT:",
                record["title"],
            )

            continue

        stats["relevantArticles"] += 1

        relevant_records.append(record)

        print(
            "  KEEP:",
            record["title"],
            f"[{relevance['category']}]",
            f"score={relevance['score']}",
        )

    print(
        f"\nRelevant articles: "
        f"{stats['relevantArticles']}"
    )

    print(
        f"Rejected as irrelevant: "
        f"{stats['rejectedIrrelevant']}"
    )

    # --------------------------------------------------------
    # 4. Load existing history IDs
    # --------------------------------------------------------

    existing_ids = load_existing_history_ids()

    # --------------------------------------------------------
    # 5. Process relevant articles
    # --------------------------------------------------------

    for record in relevant_records:
        record_id = record["id"]

        # Existing article:
        # no need to fetch/extract it again.
        if record_id in existing_ids:
            continue

        stats["articleExtractionAttempts"] += 1

        print(
            "\nProcessing:",
            record["title"],
        )

        try:
            article_text = fetch_article_text(
                record["url"]
            )

            if not article_text:
                raise RuntimeError(
                    "Article extraction returned no text"
                )

            stats["articleExtractionSuccess"] += 1

            summary = generate_summary(
                record["title"],
                record["description"],
                article_text,
            )

        except Exception as exc:
            stats["articleExtractionFailures"] += 1

            print(
                "  Extraction failed:",
                type(exc).__name__,
                str(exc),
            )

            # We can still preserve the RSS description
            # as a fallback summary.
            summary = shorten(
                record["description"],
                700,
            )

            if not summary:
                summary = record["title"]

        permanent_record = {
            "id": record["id"],
            "source": SOURCE_NAME,
            "title": record["title"],
            "publishedAtUtc": record["publishedAtUtc"],
            "url": record["url"],
            "category": record["relevance"]["category"],
            "relevanceScore": record["relevance"]["score"],
            "summary": summary,
            "feeds": record.get(
                "feeds",
                [],
            ),
            "collectedAtUtc": iso_utc(),
        }

        created = save_history_article(
            permanent_record
        )

        if created:
            stats["newHistoricalRecords"] += 1

            existing_ids.add(
                record_id
            )

            print(
                "  Saved historical record."
            )

    # --------------------------------------------------------
    # 6. Rebuild current 14-day index
    # --------------------------------------------------------

    all_history = load_all_history_articles()

    current_articles = build_current_index(
        all_history
    )

    write_current_index(
        current_articles
    )

    # --------------------------------------------------------
    # 7. Run summary
    # --------------------------------------------------------

    run_completed = utc_now()

    run_summary = {
        "status": "success",
        "source": SOURCE_NAME,
        "startedAtUtc": iso_utc(run_started),
        "completedAtUtc": iso_utc(run_completed),
        "windowDays": WINDOW_DAYS,
        "stats": stats,
        "currentArticleCount": len(
            current_articles
        ),
        "historicalArticleCount": len(
            all_history
        ),
    }

    summary_path = (
        NEWS_DIR / "run_summary.json"
    )

    with summary_path.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            run_summary,
            handle,
            indent=2,
            ensure_ascii=False,
        )

        handle.write("\n")

    print("\n" + "=" * 72)
    print("NEWS COLLECTION COMPLETE")
    print("=" * 72)

    print(
        f"RSS entries:              "
        f"{stats['rssEntries']}"
    )

    print(
        f"Duplicates:               "
        f"{stats['duplicates']}"
    )

    print(
        f"Rejected irrelevant:      "
        f"{stats['rejectedIrrelevant']}"
    )

    print(
        f"Relevant articles:        "
        f"{stats['relevantArticles']}"
    )

    print(
        f"Extraction attempts:      "
        f"{stats['articleExtractionAttempts']}"
    )

    print(
        f"Extraction successful:    "
        f"{stats['articleExtractionSuccess']}"
    )

    print(
        f"Extraction failures:     "
        f"{stats['articleExtractionFailures']}"
    )

    print(
        f"New historical records:   "
        f"{stats['newHistoricalRecords']}"
    )

    print(
        f"Historical records total: "
        f"{len(all_history)}"
    )

    print(
        f"Current 14-day records:   "
        f"{len(current_articles)}"
    )

    print("=" * 72)


if __name__ == "__main__":
    main()
