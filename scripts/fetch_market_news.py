#!/usr/bin/env python3

"""
VGrat FMS - CNA Market News Collector

PURPOSE
=======

Collect CNA news articles for the VGrat FMS market-intelligence system.

PIPELINE
========

CNA RSS
    |
    v
Discover articles
    |
    v
Deduplicate
    |
    v
Open CNA article
    |
    v
Temporarily extract article text
    |
    v
Generate local extractive summary
    |
    v
Discard full article text
    |
    v
Persist metadata + summary + URL
    |
    v
Historical archive
    |
    v
14-day current index


IMPORTANT
=========

The full CNA article text is extracted only temporarily during the run.

The permanent VGrat data files contain:

    - article ID
    - title
    - source
    - published time
    - URL
    - summary
    - first-seen time
    - category
    - extraction status

The original CNA URL is retained so the future VGrat HTML interface
can link directly to the full article on CNA.

HISTORICAL STORAGE
==================

Historical records are stored by publication date:

    data/market_news/history/YYYY/MM/YYYY-MM-DD.json

Historical records are append-only.

CURRENT STORAGE
===============

current.json contains only articles from the latest 14 days.

DUPLICATION
===========

Articles are deduplicated using a SHA-256 hash generated from:

    source + canonical URL

RSS feeds
=========

CNA currently provides feeds including:

    Latest News
    Asia
    Business
    Singapore
    World

The feed URLs are defined below.

DEPENDENCIES
============

    requests
    feedparser
    beautifulsoup4
    trafilatura

Install with:

    pip install requests feedparser beautifulsoup4 trafilatura
"""

from __future__ import annotations

import hashlib
import json
import logging
import re
import sys
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse, urlunparse

import feedparser
import requests
import trafilatura
from bs4 import BeautifulSoup


# ============================================================
# CONFIGURATION
# ============================================================

ROOT_DIR = Path(__file__).resolve().parents[1]

DATA_DIR = ROOT_DIR / "data" / "market_news"
HISTORY_DIR = DATA_DIR / "history"
CURRENT_FILE = DATA_DIR / "current.json"

CURRENT_DAYS = 14

REQUEST_TIMEOUT = 30

USER_AGENT = (
    "Mozilla/5.0 "
    "(Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 "
    "(KHTML, like Gecko) "
    "Chrome/154.0.0.0 Safari/537.36 "
    "VGrat-FMS-NewsCollector/1.0"
)


CNA_FEEDS = {
    "latest": (
        "https://www.channelnewsasia.com/"
        "api/v1/rss-outbound-feed?_format=xml"
    ),
    "asia": (
        "https://www.channelnewsasia.com/"
        "api/v1/rss-outbound-feed?_format=xml&category=6511"
    ),
    "business": (
        "https://www.channelnewsasia.com/"
        "api/v1/rss-outbound-feed?_format=xml&category=6936"
    ),
    "singapore": (
        "https://www.channelnewsasia.com/"
        "api/v1/rss-outbound-feed?_format=xml&category=10416"
    ),
    "world": (
        "https://www.channelnewsasia.com/"
        "api/v1/rss-outbound-feed?_format=xml&category=6311"
    ),
}


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

LOGGER = logging.getLogger("vgrat-market-news")


# ============================================================
# HTTP SESSION
# ============================================================

SESSION = requests.Session()

SESSION.headers.update(
    {
        "User-Agent": USER_AGENT,
        "Accept": (
            "text/html,application/xhtml+xml,"
            "application/xml;q=0.9,*/*;q=0.8"
        ),
        "Accept-Language": "en-SG,en;q=0.9",
    }
)


# ============================================================
# GENERAL HELPERS
# ============================================================

def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def iso_utc(value: datetime) -> str:
    return value.astimezone(timezone.utc).replace(
        microsecond=0
    ).isoformat().replace("+00:00", "Z")


def ensure_directories() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)


def load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default

    try:
        with path.open("r", encoding="utf-8") as handle:
            return json.load(handle)
    except Exception as exc:
        LOGGER.warning(
            "Unable to read JSON file %s: %s",
            path,
            exc,
        )
        return default


def save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    temporary = path.with_suffix(
        path.suffix + ".tmp"
    )

    with temporary.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            data,
            handle,
            ensure_ascii=False,
            indent=2,
        )
        handle.write("\n")

    temporary.replace(path)


# ============================================================
# URL NORMALISATION
# ============================================================

def canonicalise_url(url: str) -> str:
    """
    Remove tracking parameters and fragments while preserving
    the original article URL structure.
    """

    if not url:
        return ""

    url = url.strip()

    parsed = urlparse(url)

    clean_query = []

    if parsed.query:
        for item in parsed.query.split("&"):
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

            clean_query.append(item)

    query = "&".join(clean_query)

    cleaned = urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            query,
            "",
        )
    )

    return cleaned.rstrip("/")


def article_id(source: str, url: str) -> str:
    value = f"{source}|{canonicalise_url(url)}"

    return hashlib.sha256(
        value.encode("utf-8")
    ).hexdigest()


# ============================================================
# DATE PARSING
# ============================================================

def parse_feed_datetime(entry: Any) -> datetime:
    """
    Prefer parsed RSS timestamps.

    Falls back to current UTC time if the feed does not provide
    a usable publication timestamp.
    """

    for attribute in (
        "published_parsed",
        "updated_parsed",
        "created_parsed",
    ):
        parsed = getattr(entry, attribute, None)

        if parsed:
            try:
                from calendar import timegm

                timestamp = timegm(parsed)

                return datetime.fromtimestamp(
                    timestamp,
                    tz=timezone.utc,
                )
            except Exception:
                pass

    for attribute in (
        "published",
        "updated",
        "created",
    ):
        value = getattr(entry, attribute, None)

        if value:
            try:
                parsed = datetime.fromisoformat(
                    value.replace("Z", "+00:00")
                )

                if parsed.tzinfo is None:
                    parsed = parsed.replace(
                        tzinfo=timezone.utc
                    )

                return parsed.astimezone(timezone.utc)

            except Exception:
                pass

    return utc_now()


# ============================================================
# RSS EXTRACTION
# ============================================================

def clean_html(value: str) -> str:
    if not value:
        return ""

    soup = BeautifulSoup(
        value,
        "html.parser",
    )

    return " ".join(
        soup.get_text(" ", strip=True).split()
    )


def extract_rss_entry(
    entry: Any,
    category: str,
) -> dict[str, Any] | None:

    title = clean_html(
        getattr(entry, "title", "")
    )

    url = (
        getattr(entry, "link", "")
        or ""
    ).strip()

    if not title or not url:
        return None

    url = canonicalise_url(url)

    summary = clean_html(
        getattr(entry, "summary", "")
    )

    published = parse_feed_datetime(entry)

    return {
        "id": article_id(
            "CNA",
            url,
        ),
        "source": "CNA",
        "category": category,
        "title": title,
        "url": url,
        "publishedAt": iso_utc(published),
        "rssSummary": summary,
    }


# ============================================================
# RSS COLLECTION
# ============================================================

def fetch_feed(
    category: str,
    feed_url: str,
) -> list[dict[str, Any]]:

    LOGGER.info(
        "Fetching CNA %s RSS feed",
        category,
    )

    try:
        response = SESSION.get(
            feed_url,
            timeout=REQUEST_TIMEOUT,
        )

        response.raise_for_status()

        parsed = feedparser.parse(
            response.content
        )

        if getattr(
            parsed,
            "bozo",
            False,
        ):
            LOGGER.warning(
                "CNA %s RSS parser reported a malformed feed",
                category,
            )

        results: list[dict[str, Any]] = []

        for entry in parsed.entries:
            article = extract_rss_entry(
                entry,
                category,
            )

            if article:
                results.append(article)

        LOGGER.info(
            "CNA %s: %d RSS articles discovered",
            category,
            len(results),
        )

        return results

    except Exception as exc:
        LOGGER.error(
            "CNA %s RSS failed: %s",
            category,
            exc,
        )

        return []


# ============================================================
# ARTICLE EXTRACTION
# ============================================================

def fetch_article_text(url: str) -> str:
    """
    Temporarily download and extract the article text.

    The returned text is used only during processing.

    It is NOT written to the permanent historical JSON.
    """

    try:
        response = SESSION.get(
            url,
            timeout=REQUEST_TIMEOUT,
        )

        response.raise_for_status()

        html = response.text

        extracted = trafilatura.extract(
            html,
            include_comments=False,
            include_tables=False,
            include_links=False,
            favor_precision=True,
        )

        if extracted:
            return clean_text(extracted)

        # Fallback parser
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
                "aside",
            ]
        ):
            tag.decompose()

        text = soup.get_text(
            "\n",
            strip=True,
        )

        return clean_text(text)

    except Exception as exc:
        LOGGER.warning(
            "Article extraction failed for %s: %s",
            url,
            exc,
        )

        return ""


def clean_text(text: str) -> str:
    if not text:
        return ""

    lines = []

    for line in text.splitlines():
        line = " ".join(
            line.split()
        )

        if line:
            lines.append(line)

    return "\n".join(lines)


# ============================================================
# EXTRACTIVE SUMMARY
# ============================================================

STOPWORDS = {
    "about",
    "after",
    "again",
    "against",
    "being",
    "between",
    "could",
    "their",
    "there",
    "these",
    "those",
    "through",
    "where",
    "which",
    "while",
    "would",
    "from",
    "have",
    "with",
    "this",
    "that",
    "were",
    "they",
    "will",
    "into",
    "than",
    "then",
    "them",
    "been",
    "also",
    "said",
    "more",
    "such",
    "some",
    "what",
    "when",
    "over",
    "under",
    "because",
    "could",
    "should",
    "might",
    "other",
    "only",
    "many",
    "most",
    "very",
    "your",
    "ours",
    "ourselves",
    "the",
    "and",
    "for",
    "are",
    "but",
    "not",
    "you",
    "was",
    "were",
    "had",
    "has",
    "its",
    "his",
    "her",
    "our",
    "out",
    "who",
    "how",
    "why",
    "all",
    "any",
    "can",
    "may",
    "per",
    "via",
    "too",
    "just",
}


def split_sentences(text: str) -> list[str]:
    if not text:
        return []

    normalised = re.sub(
        r"\s+",
        " ",
        text,
    ).strip()

    sentences = re.split(
        r"(?<=[.!?])\s+",
        normalised,
    )

    return [
        sentence.strip()
        for sentence in sentences
        if len(sentence.strip()) >= 40
    ]


def word_tokens(text: str) -> list[str]:
    return re.findall(
        r"[A-Za-z0-9']+",
        text.lower(),
    )


def extractive_summary(
    text: str,
    max_sentences: int = 4,
) -> str:

    sentences = split_sentences(text)

    if not sentences:
        return ""

    if len(sentences) <= max_sentences:
        return " ".join(sentences)

    frequencies = Counter()

    for sentence in sentences:
        for word in word_tokens(sentence):
            if (
                len(word) >= 3
                and word not in STOPWORDS
            ):
                frequencies[word] += 1

    if not frequencies:
        return " ".join(
            sentences[:max_sentences]
        )

    scored: list[tuple[float, int, str]] = []

    for index, sentence in enumerate(sentences):
        words = word_tokens(sentence)

        if not words:
            continue

        score = sum(
            frequencies[word]
            for word in words
            if word not in STOPWORDS
        )

        # Slightly favour early sentences.
        position_bonus = max(
            0,
            1.0 - (index * 0.04),
        )

        score *= position_bonus

        scored.append(
            (
                score / max(len(words), 1),
                index,
                sentence,
            )
        )

    selected = sorted(
        scored,
        key=lambda item: item[0],
        reverse=True,
    )[:max_sentences]

    selected.sort(
        key=lambda item: item[1]
    )

    return " ".join(
        item[2]
        for item in selected
    )


# ============================================================
# ARTICLE PROCESSING
# ============================================================

def process_article(
    article: dict[str, Any],
    first_seen_at: str,
) -> dict[str, Any]:

    LOGGER.info(
        "Processing: %s",
        article["title"],
    )

    article_text = fetch_article_text(
        article["url"]
    )

    summary = ""

    extraction_status = "failed"

    if article_text:
        summary = extractive_summary(
            article_text
        )

        if summary:
            extraction_status = "success"

    if not summary:
        summary = article.get(
            "rssSummary",
            "",
        )

    result = {
        "id": article["id"],
        "source": article["source"],
        "category": article["category"],
        "title": article["title"],
        "publishedAt": article["publishedAt"],
        "url": article["url"],
        "summary": summary,
        "firstSeenAt": first_seen_at,
        "extractionStatus": extraction_status,
    }

    return result


# ============================================================
# HISTORY LOADING
# ============================================================

def history_file_for(
    published_at: str,
) -> Path:

    parsed = datetime.fromisoformat(
        published_at.replace("Z", "+00:00")
    )

    return (
        HISTORY_DIR
        / f"{parsed.year:04d}"
        / f"{parsed.month:02d}"
        / f"{parsed.date().isoformat()}.json"
    )


def load_all_history_records() -> dict[str, dict[str, Any]]:
    """
    Load historical records so duplicate articles can be detected
    across all previous runs.

    The number of articles is expected to remain manageable because
    only metadata + summaries are stored.
    """

    records: dict[str, dict[str, Any]] = {}

    if not HISTORY_DIR.exists():
        return records

    for path in HISTORY_DIR.rglob("*.json"):
        try:
            data = load_json(
                path,
                [],
            )

            if not isinstance(
                data,
                list,
            ):
                continue

            for record in data:
                record_id = record.get("id")

                if record_id:
                    records[record_id] = record

        except Exception as exc:
            LOGGER.warning(
                "Unable to process history file %s: %s",
                path,
                exc,
            )

    return records


# ============================================================
# HISTORY MERGING
# ============================================================

def save_historical_record(
    record: dict[str, Any],
) -> bool:

    path = history_file_for(
        record["publishedAt"]
    )

    existing = load_json(
        path,
        [],
    )

    if not isinstance(
        existing,
        list,
    ):
        existing = []

    existing_ids = {
        item.get("id")
        for item in existing
        if isinstance(item, dict)
    }

    if record["id"] in existing_ids:
        return False

    existing.append(record)

    existing.sort(
        key=lambda item: item.get(
            "publishedAt",
            "",
        ),
        reverse=True,
    )

    save_json(
        path,
        existing,
    )

    return True


# ============================================================
# CURRENT 14-DAY INDEX
# ============================================================

def build_current_index(
    all_records: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:

    cutoff = utc_now() - timedelta(
        days=CURRENT_DAYS
    )

    current: list[dict[str, Any]] = []

    for record in all_records.values():
        published = record.get(
            "publishedAt"
        )

        if not published:
            continue

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
            current.append(record)

    current.sort(
        key=lambda item: item.get(
            "publishedAt",
            "",
        ),
        reverse=True,
    )

    return current


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    LOGGER.info(
        "=================================================="
    )

    LOGGER.info(
        "VGrat FMS - CNA Market News Collector"
    )

    LOGGER.info(
        "=================================================="
    )

    ensure_directories()

    run_started = utc_now()

    first_seen_at = iso_utc(
        run_started
    )

    # --------------------------------------------------------
    # Load existing permanent history
    # --------------------------------------------------------

    historical_records = (
        load_all_history_records()
    )

    LOGGER.info(
        "Existing historical records: %d",
        len(historical_records),
    )

    # --------------------------------------------------------
    # Fetch RSS feeds
    # --------------------------------------------------------

    discovered: dict[
        str,
        dict[str, Any],
    ] = {}

    for category, feed_url in CNA_FEEDS.items():

        articles = fetch_feed(
            category,
            feed_url,
        )

        for article in articles:

            existing = discovered.get(
                article["id"]
            )

            if existing is None:
                discovered[
                    article["id"]
                ] = article
                continue

            # Prefer the more specific category if the same
            # article appeared in multiple CNA feeds.
            existing_category = existing.get(
                "category",
                "",
            )

            if (
                existing_category == "latest"
                and category != "latest"
            ):
                article["category"] = category

            discovered[
                article["id"]
            ] = article

    LOGGER.info(
        "Unique RSS articles discovered: %d",
        len(discovered),
    )

    # --------------------------------------------------------
    # Process only genuinely new articles
    # --------------------------------------------------------

    new_count = 0
    duplicate_count = 0
    failed_count = 0

    for article in discovered.values():

        if article["id"] in historical_records:
            duplicate_count += 1
            continue

        record = process_article(
            article,
            first_seen_at,
        )

        if record.get(
            "extractionStatus"
        ) == "failed":
            failed_count += 1

        # Permanent history
        inserted = save_historical_record(
            record
        )

        if inserted:
            historical_records[
                record["id"]
            ] = record

            new_count += 1

    # --------------------------------------------------------
    # Build current 14-day index
    # --------------------------------------------------------

    current_records = build_current_index(
        historical_records
    )

    current_payload = {
        "generatedAtUtc": iso_utc(
            utc_now()
        ),
        "windowDays": CURRENT_DAYS,
        "source": "CNA",
        "articleCount": len(
            current_records
        ),
        "articles": current_records,
    }

    save_json(
        CURRENT_FILE,
        current_payload,
    )

    # --------------------------------------------------------
    # Run summary
    # --------------------------------------------------------

    LOGGER.info(
        "=================================================="
    )

    LOGGER.info(
        "CNA NEWS COLLECTION COMPLETE"
    )

    LOGGER.info(
        "RSS articles discovered : %d",
        len(discovered),
    )

    LOGGER.info(
        "New articles            : %d",
        new_count,
    )

    LOGGER.info(
        "Duplicates skipped      : %d",
        duplicate_count,
    )

    LOGGER.info(
        "Extraction failures     : %d",
        failed_count,
    )

    LOGGER.info(
        "Historical records      : %d",
        len(historical_records),
    )

    LOGGER.info(
        "Current 14-day articles : %d",
        len(current_records),
    )

    LOGGER.info(
        "Current index            : %s",
        CURRENT_FILE,
    )

    LOGGER.info(
        "=================================================="
    )

    return 0


if __name__ == "__main__":
    sys.exit(main())
