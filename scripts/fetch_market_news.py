#!/usr/bin/env python3

"""
VGrat FMS - CNA Market News Collector

PURPOSE
=======

Collect relevant CNA market, economic, financial, technology and
economically significant geopolitical news.

PIPELINE
========

    CNA RSS feeds
          |
          v
    normalize + deduplicate
          |
          v
    HARD EXCLUSION FILTER
          |
          +---- sports
          +---- entertainment
          +---- crime
          +---- accidents
          +---- lifestyle
          +---- human interest
          +---- weather
          |
          v
    MARKET / ECONOMIC / TECHNOLOGY / GEOPOLITICAL CLASSIFICATION
          |
          v
    relevant only
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

The RSS feeds are intentionally broad.

The collector therefore does NOT fetch the full article before
classification.

Only articles that pass the relevance filter are fetched.

Permanent records contain metadata and summary only.

Full third-party article text is NOT stored permanently.

CURRENT NEWS
============

The Current News page is designed to expose up to:

    MAX_CURRENT_ARTICLES = 100

relevant articles from the rolling 14-day window.

Historical records are unlimited.

The 100-record value is a maximum, not a fabrication target.
If only 72 relevant articles exist, 72 are returned.
"""

from __future__ import annotations

import hashlib
import json
import re
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

NEWS_DIR = (
    BASE_DIR
    / "data"
    / "market_news"
)

HISTORY_DIR = (
    NEWS_DIR
    / "history"
)

CURRENT_FILE = (
    NEWS_DIR
    / "current.json"
)

RUN_SUMMARY_FILE = (
    NEWS_DIR
    / "run_summary.json"
)

# Rolling current-news window.
WINDOW_DAYS = 14

# Maximum number of relevant records exposed in current.json.
#
# This was increased from the previous 30-ish record result
# to a maximum of 100 records.
MAX_CURRENT_ARTICLES = 100

REQUEST_TIMEOUT = 25

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/154.0.0.0 Safari/537.36"
)


# ============================================================
# CNA RSS FEEDS
# ============================================================

CNA_FEEDS = {
    "latest":
        "https://www.channelnewsasia.com/api/v1/rss-outbound-feed?_format=xml",

    "asia":
        "https://www.channelnewsasia.com/api/v1/rss-outbound-feed?_format=xml&category=6511",

    "business":
        "https://www.channelnewsasia.com/api/v1/rss-outbound-feed?_format=xml&category=6936",

    "singapore":
        "https://www.channelnewsasia.com/api/v1/rss-outbound-feed?_format=xml&category=10416",

    "world":
        "https://www.channelnewsasia.com/api/v1/rss-outbound-feed?_format=xml&category=6311",
}


# ============================================================
# HARD EXCLUSION KEYWORDS
# ============================================================
#
# These categories are deliberately hard exclusions.
#
# A market keyword appearing in a sports or crime article
# must NOT be sufficient to make the article relevant.
# ============================================================

SPORTS_KEYWORDS = {
    "football",
    "soccer",
    "rugby",
    "cricket",
    "tennis",
    "golf",
    "basketball",
    "baseball",
    "badminton",
    "volleyball",
    "swimming",
    "cycling",
    "athletics",
    "formula 1",
    "formula one",
    "f1",
    "premier league",
    "champions league",
    "world cup",
    "asian games",
    "olympics",
    "olympic",
    "paralympic",
    "athlete",
    "athletes",
    "athletic",
    "coach",
    "coaches",
    "player",
    "players",
    "team",
    "teams",
    "match",
    "matches",
    "tournament",
    "tournaments",
    "championship",
    "championships",
    "medal",
    "medals",
    "gold medal",
    "silver medal",
    "bronze medal",
    "world no 1",
    "world number one",
    "wallabies",
    "battles",
    "grand slam",
    "league",
    "fixture",
    "fixtures",
    "goalkeeper",
    "striker",
    "midfielder",
    "defender",
    "batting",
    "bowling",
    "wicket",
    "racing",
    "race driver",
    "race drivers",
}


ENTERTAINMENT_KEYWORDS = {
    "movie",
    "movies",
    "film",
    "films",
    "actor",
    "actors",
    "actress",
    "actresses",
    "singer",
    "singers",
    "concert",
    "concerts",
    "celebrity",
    "celebrities",
    "music",
    "album",
    "albums",
    "television",
    "tv show",
    "tv shows",
    "reality show",
    "reality tv",
    "entertainment",
    "box office",
    "red carpet",
    "premiere",
    "director",
    "directors",
}


CRIME_KEYWORDS = {
    "murder",
    "murdered",
    "murderer",
    "homicide",
    "robbery",
    "robbed",
    "burglary",
    "burglar",
    "assault",
    "assaulted",
    "arrested",
    "arrest",
    "charged",
    "jail",
    "jailed",
    "prison",
    "crime",
    "criminal",
    "criminals",
    "police",
    "court",
    "trial",
    "sentenced",
    "sentence",
    "stole",
    "stolen",
    "thieves",
    "thief",
    "drugs",
    "drug trafficking",
    "drug trafficker",
    "vandalism",
    "fraud",
    "scam",
    "scams",
    "con artist",
}


ACCIDENT_KEYWORDS = {
    "accident",
    "accidents",
    "crash",
    "crashes",
    "crashed",
    "collision",
    "collided",
    "fire",
    "explosion",
    "exploded",
    "killed",
    "killed in",
    "dies",
    "died",
    "death",
    "deaths",
    "dead",
    "fatal",
    "fatality",
    "fatalities",
    "injured",
    "injury",
    "injuries",
    "rescue",
    "rescued",
    "missing",
    "missing person",
}


LIFESTYLE_KEYWORDS = {
    "restaurant",
    "restaurants",
    "bakery",
    "cafe",
    "coffee shop",
    "food",
    "recipe",
    "recipes",
    "travel",
    "tourism",
    "tourist attraction",
    "holiday",
    "holidays",
    "fashion",
    "beauty",
    "wellness",
    "lifestyle",
    "shopping",
    "viral",
    "viral video",
    "influencer",
    "social media star",
    "pets",
    "pet",
    "birthday",
    "wedding",
    "dating",
    "relationship",
}


WEATHER_KEYWORDS = {
    "rainy days",
    "rainfall",
    "weather",
    "weather forecast",
    "temperature",
    "temperatures",
    "heatwave",
    "heat wave",
    "storm",
    "storms",
    "thunderstorm",
    "typhoon",
    "hurricane",
    "tornado",
    "flood",
    "flooding",
    "floods",
    "drought",
    "droughts",
    "snowfall",
    "snowstorm",
}


# ============================================================
# MARKET KEYWORDS
# ============================================================

MARKET_KEYWORDS = {
    "stock": 5,
    "stocks": 5,
    "share": 5,
    "shares": 5,
    "share price": 6,
    "share prices": 6,
    "equities": 6,
    "equity": 5,
    "stock market": 7,
    "stock markets": 7,
    "financial markets": 6,
    "market": 3,
    "markets": 4,
    "investor": 4,
    "investors": 4,
    "investment": 4,
    "investments": 4,
    "portfolio": 4,
    "portfolios": 4,
    "asset management": 6,
    "asset managers": 6,
    "fund": 4,
    "funds": 4,
    "fundraising": 5,
    "fundraise": 5,
    "ipo": 7,
    "ipos": 7,
    "initial public offering": 7,
    "listing": 5,
    "listed company": 6,
    "earnings": 6,
    "profit": 5,
    "profits": 5,
    "revenue": 5,
    "revenues": 5,
    "dividend": 5,
    "dividends": 5,
    "merger": 6,
    "mergers": 6,
    "acquisition": 6,
    "acquisitions": 6,
    "takeover": 6,
    "takeovers": 6,
    "valuation": 5,
    "valuations": 5,
    "capital expenditure": 6,
    "capex": 6,
    "credit line": 5,
    "credit lines": 5,
    "debt": 4,
    "default": 5,
    "defaults": 5,
    "bank": 4,
    "banks": 4,
    "banking": 6,
    "financial institution": 6,
    "financial institutions": 6,
    "insurer": 5,
    "insurers": 5,
    "insurance": 5,
    "credit": 4,
    "loan": 4,
    "loans": 4,
    "bond": 6,
    "bonds": 6,
    "yield": 6,
    "yields": 6,
    "treasury": 5,
    "treasuries": 5,
    "interest rate": 7,
    "interest rates": 7,
    "rate cut": 7,
    "rate cuts": 7,
    "rate hike": 7,
    "rate hikes": 7,
    "central bank": 7,
    "central banks": 7,
    "monetary policy": 7,
    "stocks drop": 8,
    "stocks rise": 8,
    "stocks fall": 8,
    "oil rises": 8,
    "oil gains": 8,
    "oil falls": 8,
    "oil down": 8,
    "gold rises": 7,
    "gold falls": 7,
    "currency": 5,
    "currencies": 5,
    "forex": 7,
    "foreign exchange": 7,
    "fx": 5,
    "dollar": 4,
    "yuan": 5,
    "renminbi": 5,
    "yen": 5,
    "euro": 5,
    "sterling": 5,
    "commodity": 5,
    "commodities": 5,
    "oil": 5,
    "crude": 5,
    "gold": 5,
    "copper": 5,
    "natural gas": 6,
    "energy prices": 7,
    "property market": 6,
    "property prices": 6,
    "housing market": 6,
    "real estate": 5,
    "reit": 7,
    "reits": 7,
}


# ============================================================
# ECONOMIC KEYWORDS
# ============================================================

ECONOMIC_KEYWORDS = {
    "economy": 6,
    "economic": 6,
    "economic growth": 8,
    "economic outlook": 7,
    "gdp": 8,
    "gross domestic product": 8,
    "recession": 8,
    "inflation": 8,
    "deflation": 8,
    "cpi": 8,
    "consumer price index": 8,
    "ppi": 8,
    "producer price index": 8,
    "interest rate": 8,
    "interest rates": 8,
    "central bank": 8,
    "central banks": 8,
    "monetary policy": 8,
    "fiscal policy": 7,
    "government spending": 6,
    "budget": 6,
    "tax policy": 6,
    "tax reform": 6,
    "stimulus": 7,
    "economic policy": 7,
    "business confidence": 7,
    "consumer confidence": 7,
    "consumer spending": 6,
    "consumer prices": 7,
    "producer prices": 7,
    "retail sales": 7,
    "industrial production": 8,
    "industrial output": 8,
    "factory output": 8,
    "factory production": 8,
    "manufacturing output": 8,
    "manufacturing production": 8,
    "manufacturing": 5,
    "pmi": 8,
    "purchasing managers": 7,
    "exports": 7,
    "export growth": 8,
    "export growth": 8,
    "imports": 7,
    "import growth": 8,
    "trade surplus": 8,
    "trade deficit": 8,
    "trade balance": 8,
    "trade data": 7,
    "tariff": 7,
    "tariffs": 7,
    "trade war": 8,
    "supply chain": 6,
    "employment": 6,
    "unemployment": 7,
    "jobs data": 7,
    "nonfarm payroll": 8,
    "payrolls": 7,
    "wages": 5,
    "wage growth": 7,
    "productivity": 6,
    "business activity": 6,
    "economic activity": 7,
    "cost of living": 5,
    "electricity tariffs": 7,
    "gas tariffs": 7,
    "utility tariffs": 7,
    "household tariffs": 6,
    "energy prices": 7,
}


# ============================================================
# TECHNOLOGY / AI / SEMICONDUCTOR KEYWORDS
# ============================================================

TECHNOLOGY_KEYWORDS = {
    "semiconductor": 7,
    "semiconductors": 7,
    "chip": 6,
    "chips": 6,
    "chipmaking": 7,
    "chip maker": 7,
    "chipmakers": 7,
    "foundry": 6,
    "advanced chips": 7,
    "ai chip": 7,
    "ai chips": 7,
    "artificial intelligence": 5,
    "artificial intelligence sector": 7,
    "ai infrastructure": 7,
    "ai infra": 7,
    "data centre": 7,
    "data centres": 7,
    "data center": 7,
    "data centers": 7,
    "cloud computing": 6,
    "technology investment": 7,
    "technology company": 5,
    "technology companies": 5,
    "tech company": 5,
    "tech companies": 5,
    "venture capital": 7,
    "private equity": 7,
    "startup funding": 6,
    "technology funding": 7,
    "funding round": 6,
    "ipo": 7,
    "nvidia": 6,
    "deepseek": 6,
    "openai": 5,
    "anthropic": 6,
    "tesla": 5,
    "apple": 4,
    "microsoft": 4,
    "google": 4,
    "amazon": 4,
    "meta": 4,
}


# ============================================================
# GEOPOLITICAL KEYWORDS
# ============================================================

GEOPOLITICAL_KEYWORDS = {
    "sanction": 8,
    "sanctions": 8,
    "trade restrictions": 8,
    "export controls": 8,
    "import restrictions": 7,
    "trade war": 8,
    "tariff": 7,
    "tariffs": 7,
    "shipping disruption": 8,
    "shipping disruptions": 8,
    "shipping route": 6,
    "shipping routes": 6,
    "strait": 4,
    "hormuz": 8,
    "strait of hormuz": 9,
    "taiwan strait": 8,
    "south china sea": 7,
    "rare earth": 7,
    "rare earths": 7,
    "strategic minerals": 7,
    "energy security": 7,
    "oil supply": 8,
    "gas supply": 7,
    "pipeline": 6,
    "pipelines": 6,
    "oil loading": 7,
    "oil shipments": 8,
    "shipping": 4,
    "war": 2,
    "conflict": 3,
    "ceasefire": 4,
    "military escalation": 6,
    "ukraine": 2,
    "russia": 2,
    "iran": 3,
    "china": 2,
    "united states": 2,
    "european union": 3,
}


# ============================================================
# GENERAL POLITICAL / DIPLOMATIC WORDS
# ============================================================
#
# These do NOT create relevance by themselves.
#
# This prevents stories such as:
#
#   "politician says..."
#   "leader meets..."
#   "minister remains..."
#
# from entering the market dataset merely because they
# contain a country name or geopolitical vocabulary.
# ============================================================

GENERAL_POLITICAL_KEYWORDS = {
    "president",
    "prime minister",
    "minister",
    "lawmaker",
    "lawmakers",
    "politician",
    "politicians",
    "election",
    "elections",
    "political party",
    "political parties",
    "parliament",
    "parliamentary",
    "government",
    "diplomatic",
    "diplomacy",
    "ambassador",
    "bilateral",
    "summit",
    "meeting",
    "meets",
    "met",
    "statement",
    "says",
    "said",
    "pledge",
    "pledged",
    "resignation",
    "resigns",
    "appointment",
    "appointed",
}


# ============================================================
# HELPERS
# ============================================================

def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def iso_utc(
    dt: datetime | None = None,
) -> str:

    if dt is None:
        dt = utc_now()

    return (
        dt.astimezone(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )


def clean_text(
    value: Any,
) -> str:

    if value is None:
        return ""

    text = str(value)

    text = BeautifulSoup(
        text,
        "html.parser",
    ).get_text(
        " ",
        strip=True,
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def canonicalize_url(
    url: str,
) -> str:

    if not url:
        return ""

    try:
        parts = urlsplit(
            url.strip()
        )

        query_parts = []

        for item in parts.query.split("&"):

            if not item:
                continue

            key = item.split(
                "=",
                1,
            )[0].lower()

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

        normalized_query = "&".join(
            query_parts
        )

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


def article_id(
    source: str,
    url: str,
) -> str:

    raw = (
        f"{source}|"
        f"{canonicalize_url(url)}"
    ).encode("utf-8")

    return hashlib.sha256(
        raw
    ).hexdigest()


def shorten(
    text: str,
    max_chars: int = 320,
) -> str:

    text = clean_text(text)

    if len(text) <= max_chars:
        return text

    truncated = (
        text[:max_chars]
        .rsplit(" ", 1)[0]
    )

    return (
        truncated
        .rstrip(" .,;:")
        + "..."
    )


def keyword_matches(
    text: str,
    keywords: set[str],
) -> list[str]:

    lower = text.lower()

    return [
        keyword
        for keyword in keywords
        if keyword.lower() in lower
    ]


def weighted_score(
    text: str,
    keywords: dict[str, int],
) -> tuple[int, list[str]]:

    lower = text.lower()

    score = 0
    matches = []

    for keyword, weight in keywords.items():

        if keyword.lower() in lower:

            score += weight

            matches.append(
                keyword
            )

    return score, matches


# ============================================================
# DATE PARSING
# ============================================================

def parse_entry_datetime(
    entry: Any,
) -> datetime:

    parsed = getattr(
        entry,
        "published_parsed",
        None,
    )

    if parsed is None:

        parsed = getattr(
            entry,
            "updated_parsed",
            None,
        )

    if parsed is not None:

        try:
            from calendar import timegm

            timestamp = timegm(
                parsed
            )

            return datetime.fromtimestamp(
                timestamp,
                tz=timezone.utc,
            )

        except Exception:
            pass

    return utc_now()


# ============================================================
# RELEVANCE CLASSIFICATION
# ============================================================

def classify_relevance(
    title: str,
    description: str,
) -> dict[str, Any]:

    """
    Classify an article using title + RSS description.

    Classification order:

        1. Sports
        2. Entertainment
        3. Crime
        4. Accidents
        5. Lifestyle
        6. Weather
        7. Market
        8. Economic
        9. Technology
       10. Geopolitical + economic impact

    Hard exclusions always win.
    """

    text = clean_text(
        f"{title}. {description}"
    )

    lower = text.lower()

    # --------------------------------------------------------
    # HARD EXCLUSIONS
    # --------------------------------------------------------

    sports = keyword_matches(
        lower,
        SPORTS_KEYWORDS,
    )

    entertainment = keyword_matches(
        lower,
        ENTERTAINMENT_KEYWORDS,
    )

    crime = keyword_matches(
        lower,
        CRIME_KEYWORDS,
    )

    accidents = keyword_matches(
        lower,
        ACCIDENT_KEYWORDS,
    )

    lifestyle = keyword_matches(
        lower,
        LIFESTYLE_KEYWORDS,
    )

    weather = keyword_matches(
        lower,
        WEATHER_KEYWORDS,
    )

    # --------------------------------------------------------
    # Special handling:
    #
    # "market" is frequently used in sports articles
    # in descriptions such as player transfer markets.
    #
    # Therefore hard exclusions happen first.
    # --------------------------------------------------------

    if sports:

        return {
            "relevant": False,
            "category": "SPORTS",
            "score": -100,
            "reason": "hard_exclusion_sports",
            "matchedKeywords": sports,
            "excludedKeywords": sports,
        }

    if entertainment:

        return {
            "relevant": False,
            "category": "ENTERTAINMENT",
            "score": -100,
            "reason": "hard_exclusion_entertainment",
            "matchedKeywords": entertainment,
            "excludedKeywords": entertainment,
        }

    if crime:

        return {
            "relevant": False,
            "category": "CRIME",
            "score": -100,
            "reason": "hard_exclusion_crime",
            "matchedKeywords": crime,
            "excludedKeywords": crime,
        }

    if accidents:

        return {
            "relevant": False,
            "category": "ACCIDENT",
            "score": -100,
            "reason": "hard_exclusion_accident",
            "matchedKeywords": accidents,
            "excludedKeywords": accidents,
        }

    if lifestyle:

        return {
            "relevant": False,
            "category": "LIFESTYLE",
            "score": -100,
            "reason": "hard_exclusion_lifestyle",
            "matchedKeywords": lifestyle,
            "excludedKeywords": lifestyle,
        }

    # Weather alone is not market news.
    #
    # However, we do NOT hard-reject severe weather when the
    # same article contains a strong economic impact signal.
    if weather:

        economic_score, economic_matches = weighted_score(
            lower,
            ECONOMIC_KEYWORDS,
        )

        market_score, market_matches = weighted_score(
            lower,
            MARKET_KEYWORDS,
        )

        if (
            economic_score < 8
            and market_score < 8
        ):

            return {
                "relevant": False,
                "category": "WEATHER",
                "score": -100,
                "reason": "hard_exclusion_weather",
                "matchedKeywords": weather,
                "excludedKeywords": weather,
            }

    # --------------------------------------------------------
    # POSITIVE CLASSIFICATION
    # --------------------------------------------------------

    market_score, market_matches = weighted_score(
        lower,
        MARKET_KEYWORDS,
    )

    economic_score, economic_matches = weighted_score(
        lower,
        ECONOMIC_KEYWORDS,
    )

    technology_score, technology_matches = weighted_score(
        lower,
        TECHNOLOGY_KEYWORDS,
    )

    geopolitical_score, geopolitical_matches = weighted_score(
        lower,
        GEOPOLITICAL_KEYWORDS,
    )

    political_matches = keyword_matches(
        lower,
        GENERAL_POLITICAL_KEYWORDS,
    )

    # --------------------------------------------------------
    # MARKET
    # --------------------------------------------------------

    if market_score >= 6:

        return {
            "relevant": True,
            "category": "MARKET",
            "score": market_score,
            "reason": "strong_market_signal",
            "matchedKeywords": market_matches,
            "excludedKeywords": [],
        }

    # --------------------------------------------------------
    # ECONOMIC
    # --------------------------------------------------------

    if economic_score >= 6:

        return {
            "relevant": True,
            "category": "ECONOMIC",
            "score": economic_score,
            "reason": "strong_economic_signal",
            "matchedKeywords": economic_matches,
            "excludedKeywords": [],
        }

    # --------------------------------------------------------
    # TECHNOLOGY
    # --------------------------------------------------------
    #
    # Technology alone is not enough.
    #
    # We require meaningful technology/business/investment
    # signals.
    # --------------------------------------------------------

    strong_technology_terms = {
        "semiconductor",
        "semiconductors",
        "chipmaking",
        "chip maker",
        "chipmakers",
        "advanced chips",
        "ai chip",
        "ai chips",
        "ai infrastructure",
        "ai infra",
        "data centre",
        "data centres",
        "data center",
        "data centers",
        "venture capital",
        "private equity",
        "startup funding",
        "technology funding",
        "funding round",
    }

    strong_technology_matches = [
        item
        for item in technology_matches
        if item in strong_technology_terms
    ]

    if (
        technology_score >= 7
        and strong_technology_matches
    ):

        return {
            "relevant": True,
            "category": "TECHNOLOGY",
            "score": technology_score,
            "reason": "strong_technology_business_signal",
            "matchedKeywords": technology_matches,
            "excludedKeywords": [],
        }

    # --------------------------------------------------------
    # GEOPOLITICAL
    # --------------------------------------------------------
    #
    # Geopolitical stories must have a direct economic/
    # market consequence.
    #
    # A politician simply discussing Ukraine, Iran, Russia,
    # etc. is not sufficient.
    # --------------------------------------------------------

    economic_impact_terms = {
        "oil",
        "crude",
        "energy prices",
        "oil supply",
        "oil shipments",
        "oil loading",
        "gas supply",
        "pipeline",
        "pipelines",
        "shipping",
        "shipping disruption",
        "shipping disruptions",
        "trade",
        "trade restrictions",
        "tariff",
        "tariffs",
        "sanctions",
        "sanction",
        "export controls",
        "import restrictions",
        "rare earth",
        "rare earths",
        "strategic minerals",
        "semiconductor",
        "semiconductors",
        "chip",
        "chips",
        "energy security",
        "hormuz",
        "strait of hormuz",
    }

    direct_economic_impact = [
        item
        for item in geopolitical_matches
        if item in economic_impact_terms
    ]

    # Country-only geopolitical terms such as:
    #
    #   Ukraine
    #   Iran
    #   Russia
    #
    # do NOT qualify on their own.

    if (
        geopolitical_score >= 7
        and direct_economic_impact
    ):

        return {
            "relevant": True,
            "category": "GEOPOLITICAL",
            "score": geopolitical_score,
            "reason": "geopolitical_with_economic_impact",
            "matchedKeywords": geopolitical_matches,
            "excludedKeywords": [],
        }

    # --------------------------------------------------------
    # GENERAL POLITICAL ARTICLES
    # --------------------------------------------------------
    #
    # Explicitly reject political commentary or political
    # statements unless another classifier above qualified
    # the article.
    # --------------------------------------------------------

    if political_matches:

        return {
            "relevant": False,
            "category": "POLITICAL_GENERAL",
            "score": 0,
            "reason": "political_without_market_impact",
            "matchedKeywords": political_matches,
            "excludedKeywords": [],
        }

    # --------------------------------------------------------
    # DEFAULT
    # --------------------------------------------------------

    return {
        "relevant": False,
        "category": "GENERAL",
        "score": 0,
        "reason": "insufficient_market_relevance",
        "matchedKeywords": [],
        "excludedKeywords": [],
    }


# ============================================================
# RSS FETCH
# ============================================================

def fetch_feed(
    feed_name: str,
    feed_url: str,
) -> list[dict[str, Any]]:

    print(
        f"\nFetching CNA feed: {feed_name}"
    )

    try:

        response = requests.get(
            feed_url,
            headers={
                "User-Agent": USER_AGENT,
                "Accept": (
                    "application/rss+xml, "
                    "application/xml, "
                    "text/xml"
                ),
            },
            timeout=REQUEST_TIMEOUT,
        )

        response.raise_for_status()

        parsed = feedparser.parse(
            response.content
        )

        records = []

        for entry in parsed.entries:

            title = clean_text(
                getattr(
                    entry,
                    "title",
                    "",
                )
            )

            description = clean_text(
                getattr(
                    entry,
                    "summary",
                    getattr(
                        entry,
                        "description",
                        "",
                    ),
                )
            )

            url = clean_text(
                getattr(
                    entry,
                    "link",
                    "",
                )
            )

            if not title or not url:
                continue

            published_dt = parse_entry_datetime(
                entry
            )

            records.append(
                {
                    "feed": feed_name,
                    "title": title,
                    "description": description,
                    "url": canonicalize_url(
                        url
                    ),
                    "publishedAtUtc": iso_utc(
                        published_dt
                    ),
                }
            )

        print(
            f"  Entries: {len(records)}"
        )

        return records

    except Exception as exc:

        print(
            f"  ERROR fetching {feed_name}: "
            f"{type(exc).__name__}: {exc}"
        )

        return []


def collect_rss() -> list[dict[str, Any]]:

    records = []

    for feed_name, feed_url in CNA_FEEDS.items():

        records.extend(
            fetch_feed(
                feed_name,
                feed_url,
            )
        )

    return records


# ============================================================
# ARTICLE EXTRACTION
# ============================================================

def fetch_article_text(
    url: str,
) -> str:

    response = requests.get(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": (
                "text/html,"
                "application/xhtml+xml"
            ),
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

        return clean_text(
            extracted
        )

    # --------------------------------------------------------
    # Fallback BeautifulSoup extraction
    # --------------------------------------------------------

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

    for paragraph in soup.find_all(
        "p"
    ):

        text = clean_text(
            paragraph.get_text(
                " ",
                strip=True,
            )
        )

        if len(text) >= 40:

            paragraphs.append(
                text
            )

    return clean_text(
        " ".join(
            paragraphs
        )
    )


# ============================================================
# SUMMARY
# ============================================================

def split_sentences(
    text: str,
) -> list[str]:

    text = clean_text(
        text
    )

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

    candidates = split_sentences(
        article_text
    )

    if not candidates:

        description = clean_text(
            description
        )

        if description:

            return shorten(
                description,
                700,
            )

        return shorten(
            title,
            700,
        )

    priority_terms = [
        "market",
        "stock",
        "stocks",
        "share",
        "shares",
        "investor",
        "investors",
        "economy",
        "economic",
        "growth",
        "inflation",
        "interest rate",
        "central bank",
        "bank",
        "bond",
        "bonds",
        "yield",
        "currency",
        "trade",
        "tariff",
        "tariffs",
        "oil",
        "energy",
        "technology",
        "semiconductor",
        "semiconductors",
        "chip",
        "chips",
        "artificial intelligence",
        "ai",
        "gdp",
        "manufacturing",
        "factory output",
        "industrial production",
        "exports",
        "imports",
        "ipo",
        "funding",
        "investment",
    ]

    scored = []

    for index, sentence in enumerate(
        candidates
    ):

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

    selected = [
        item[2]
        for item in scored[:3]
    ]

    selected_set = set(
        selected
    )

    ordered = [
        sentence
        for sentence in candidates
        if sentence in selected_set
    ][:3]

    summary = " ".join(
        ordered
    )

    return shorten(
        summary,
        700,
    )


# ============================================================
# JSON HELPERS
# ============================================================

def load_json(
    path: Path,
    default: Any,
) -> Any:

    if not path.exists():
        return default

    try:

        with path.open(
            "r",
            encoding="utf-8",
        ) as handle:

            return json.load(
                handle
            )

    except Exception as exc:

        print(
            f"WARNING: Could not read "
            f"{path}: "
            f"{type(exc).__name__}: {exc}"
        )

        return default


# ============================================================
# HISTORY LOADING
# ============================================================

def load_existing_history_ids() -> set[str]:

    ids = set()

    if not HISTORY_DIR.exists():
        return ids

    for path in HISTORY_DIR.rglob(
        "*.json"
    ):

        data = load_json(
            path,
            {},
        )

        for article in data.get(
            "articles",
            [],
        ):

            if not isinstance(
                article,
                dict,
            ):
                continue

            value = article.get(
                "id"
            )

            if value:
                ids.add(
                    value
                )

    return ids


def load_all_history_articles() -> list[dict[str, Any]]:

    articles = []

    if not HISTORY_DIR.exists():
        return articles

    for path in HISTORY_DIR.rglob(
        "*.json"
    ):

        data = load_json(
            path,
            {},
        )

        for article in data.get(
            "articles",
            [],
        ):

            if isinstance(
                article,
                dict,
            ):

                articles.append(
                    article
                )

    return articles


# ============================================================
# HISTORY FILE
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

    return (
        directory
        / f"{dt.date().isoformat()}.json"
    )


def save_history_article(
    article: dict[str, Any],
) -> bool:

    path = history_file_for_date(
        article[
            "publishedAtUtc"
        ]
    )

    data = load_json(
        path,
        {
            "date": (
                article[
                    "publishedAtUtc"
                ][:10]
            ),
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
        if isinstance(
            item,
            dict,
        )
    }

    if article["id"] in existing_ids:

        return False

    articles.append(
        article
    )

    articles.sort(
        key=lambda item: item.get(
            "publishedAtUtc",
            "",
        ),
        reverse=True,
    )

    data[
        "articleCount"
    ] = len(
        articles
    )

    data[
        "updatedAtUtc"
    ] = iso_utc()

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

        handle.write(
            "\n"
        )

    return True


# ============================================================
# CURRENT 14-DAY INDEX
# ============================================================

def build_current_index(
    all_history_articles: list[dict[str, Any]],
) -> list[dict[str, Any]]:

    cutoff = (
        utc_now()
        - timedelta(
            days=WINDOW_DAYS
        )
    )

    current = []

    for article in all_history_articles:

        published = article.get(
            "publishedAtUtc",
            "",
        )

        try:

            published_dt = (
                datetime.fromisoformat(
                    published.replace(
                        "Z",
                        "+00:00",
                    )
                )
            )

        except Exception:

            continue

        if published_dt < cutoff:
            continue

        current.append(
            article
        )

    # --------------------------------------------------------
    # Deduplicate.
    # --------------------------------------------------------

    unique = {}

    for article in current:

        article_key = article.get(
            "id"
        )

        if not article_key:
            continue

        unique[
            article_key
        ] = article

    current = list(
        unique.values()
    )

    # --------------------------------------------------------
    # Newest first.
    # --------------------------------------------------------

    current.sort(
        key=lambda item: item.get(
            "publishedAtUtc",
            "",
        ),
        reverse=True,
    )

    # --------------------------------------------------------
    # Maximum Current News size.
    #
    # Historical archive remains untouched.
    # --------------------------------------------------------

    return current[
        :MAX_CURRENT_ARTICLES
    ]


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
        "maxArticles": MAX_CURRENT_ARTICLES,
        "source": SOURCE_NAME,
        "articleCount": len(
            articles
        ),
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

        handle.write(
            "\n"
        )


# ============================================================
# MAIN
# ============================================================

def main() -> None:

    print(
        "=" * 72
    )

    print(
        "VGrat FMS - CNA Market News Collector"
    )

    print(
        "=" * 72
    )

    print(
        f"Current News window: "
        f"{WINDOW_DAYS} days"
    )

    print(
        f"Current News maximum: "
        f"{MAX_CURRENT_ARTICLES} records"
    )

    print()

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
        "uniqueEntries": 0,
        "duplicates": 0,
        "rejectedSports": 0,
        "rejectedEntertainment": 0,
        "rejectedCrime": 0,
        "rejectedAccident": 0,
        "rejectedLifestyle": 0,
        "rejectedWeather": 0,
        "rejectedPolitical": 0,
        "rejectedGeneral": 0,
        "rejectedOther": 0,
        "relevantMarket": 0,
        "relevantEconomic": 0,
        "relevantTechnology": 0,
        "relevantGeopolitical": 0,
        "relevantArticles": 0,
        "articleExtractionAttempts": 0,
        "articleExtractionSuccess": 0,
        "articleExtractionFailures": 0,
        "newHistoricalRecords": 0,
    }

    # ========================================================
    # 1. RSS
    # ========================================================

    rss_records = collect_rss()

    stats[
        "rssEntries"
    ] = len(
        rss_records
    )

    print()
    print(
        f"RSS entries collected: "
        f"{stats['rssEntries']}"
    )

    # ========================================================
    # 2. Deduplicate
    # ========================================================

    unique_records = {}

    for record in rss_records:

        url = record.get(
            "url",
            "",
        )

        if not url:
            continue

        record_key = article_id(
            SOURCE_NAME,
            url,
        )

        if record_key in unique_records:

            stats[
                "duplicates"
            ] += 1

            existing = unique_records[
                record_key
            ]

            feeds = existing.setdefault(
                "feeds",
                [],
            )

            current_feed = record.get(
                "feed"
            )

            if (
                current_feed
                and current_feed
                not in feeds
            ):

                feeds.append(
                    current_feed
                )

            continue

        record["id"] = record_key

        record["feeds"] = [
            record.get(
                "feed",
                "unknown",
            )
        ]

        unique_records[
            record_key
        ] = record

    stats[
        "uniqueEntries"
    ] = len(
        unique_records
    )

    print(
        f"Unique RSS entries: "
        f"{stats['uniqueEntries']}"
    )

    print(
        f"RSS duplicates: "
        f"{stats['duplicates']}"
    )

    # ========================================================
    # 3. Relevance filter
    # ========================================================

    relevant_records = []

    print()
    print(
        "RELEVANCE FILTER"
    )
    print(
        "-" * 72
    )

    for record in unique_records.values():

        relevance = classify_relevance(
            record[
                "title"
            ],
            record[
                "description"
            ],
        )

        record[
            "relevance"
        ] = relevance

        if not relevance[
            "relevant"
        ]:

            category = relevance[
                "category"
            ]

            if category == "SPORTS":
                stats[
                    "rejectedSports"
                ] += 1

            elif category == "ENTERTAINMENT":
                stats[
                    "rejectedEntertainment"
                ] += 1

            elif category == "CRIME":
                stats[
                    "rejectedCrime"
                ] += 1

            elif category == "ACCIDENT":
                stats[
                    "rejectedAccident"
                ] += 1

            elif category == "LIFESTYLE":
                stats[
                    "rejectedLifestyle"
                ] += 1

            elif category == "WEATHER":
                stats[
                    "rejectedWeather"
                ] += 1

            elif category == "POLITICAL_GENERAL":
                stats[
                    "rejectedPolitical"
                ] += 1

            elif category == "GENERAL":
                stats[
                    "rejectedGeneral"
                ] += 1

            else:
                stats[
                    "rejectedOther"
                ] += 1

            print(
                "REJECT:",
                record[
                    "title"
                ],
            )

            continue

        category = relevance[
            "category"
        ]

        if category == "MARKET":
            stats[
                "relevantMarket"
            ] += 1

        elif category == "ECONOMIC":
            stats[
                "relevantEconomic"
            ] += 1

        elif category == "TECHNOLOGY":
            stats[
                "relevantTechnology"
            ] += 1

        elif category == "GEOPOLITICAL":
            stats[
                "relevantGeopolitical"
            ] += 1

        stats[
            "relevantArticles"
        ] += 1

        relevant_records.append(
            record
        )

        print(
            "KEEP:",
            record[
                "title"
            ],
            f"[{category}]",
            f"score={relevance['score']}",
        )

    # ========================================================
    # 4. Print filtering statistics
    # ========================================================

    print()
    print(
        "=" * 72
    )

    print(
        "FILTER RESULT"
    )

    print(
        "=" * 72
    )

    print(
        f"Relevant articles:        "
        f"{stats['relevantArticles']}"
    )

    print(
        f"  MARKET:                 "
        f"{stats['relevantMarket']}"
    )

    print(
        f"  ECONOMIC:               "
        f"{stats['relevantEconomic']}"
    )

    print(
        f"  TECHNOLOGY:             "
        f"{stats['relevantTechnology']}"
    )

    print(
        f"  GEOPOLITICAL:           "
        f"{stats['relevantGeopolitical']}"
    )

    print()

    print(
        f"Rejected sports:          "
        f"{stats['rejectedSports']}"
    )

    print(
        f"Rejected entertainment:   "
        f"{stats['rejectedEntertainment']}"
    )

    print(
        f"Rejected crime:           "
        f"{stats['rejectedCrime']}"
    )

    print(
        f"Rejected accidents:       "
        f"{stats['rejectedAccident']}"
    )

    print(
        f"Rejected lifestyle:       "
        f"{stats['rejectedLifestyle']}"
    )

    print(
        f"Rejected weather:         "
        f"{stats['rejectedWeather']}"
    )

    print(
        f"Rejected political:       "
        f"{stats['rejectedPolitical']}"
    )

    print(
        f"Rejected general:         "
        f"{stats['rejectedGeneral']}"
    )

    # ========================================================
    # 5. Existing history
    # ========================================================

    existing_ids = (
        load_existing_history_ids()
    )

    # ========================================================
    # 6. Process relevant articles
    # ========================================================

    for record in relevant_records:

        record_key = record[
            "id"
        ]

        # Already archived.
        if record_key in existing_ids:
            continue

        print()
        print(
            "Processing:",
            record[
                "title"
            ],
        )

        stats[
            "articleExtractionAttempts"
        ] += 1

        try:

            article_text = fetch_article_text(
                record[
                    "url"
                ]
            )

            if not article_text:

                raise RuntimeError(
                    "Article extraction "
                    "returned no text"
                )

            stats[
                "articleExtractionSuccess"
            ] += 1

            summary = generate_summary(
                record[
                    "title"
                ],
                record[
                    "description"
                ],
                article_text,
            )

        except Exception as exc:

            stats[
                "articleExtractionFailures"
            ] += 1

            print(
                "  Extraction failed:",
                type(exc).__name__,
                str(exc),
            )

            summary = shorten(
                record[
                    "description"
                ],
                700,
            )

            if not summary:
                summary = record[
                    "title"
                ]

        permanent_record = {
            "id": record[
                "id"
            ],
            "source": SOURCE_NAME,
            "title": record[
                "title"
            ],
            "publishedAtUtc": record[
                "publishedAtUtc"
            ],
            "url": record[
                "url"
            ],
            "category": record[
                "relevance"
            ][
                "category"
            ],
            "relevanceScore": record[
                "relevance"
            ][
                "score"
            ],
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

            stats[
                "newHistoricalRecords"
            ] += 1

            existing_ids.add(
                record_key
            )

            print(
                "  Saved historical record."
            )

    # ========================================================
    # 7. Build Current News
    # ========================================================

    all_history = (
        load_all_history_articles()
    )

    current_articles = (
        build_current_index(
            all_history
        )
    )

    write_current_index(
        current_articles
    )

    # ========================================================
    # 8. Run summary
    # ========================================================

    run_completed = utc_now()

    run_summary = {
        "status": "success",
        "source": SOURCE_NAME,
        "startedAtUtc": iso_utc(
            run_started
        ),
        "completedAtUtc": iso_utc(
            run_completed
        ),
        "windowDays": WINDOW_DAYS,
        "maxCurrentArticles": (
            MAX_CURRENT_ARTICLES
        ),
        "stats": stats,
        "currentArticleCount": len(
            current_articles
        ),
        "historicalArticleCount": len(
            all_history
        ),
    }

    with RUN_SUMMARY_FILE.open(
        "w",
        encoding="utf-8",
    ) as handle:

        json.dump(
            run_summary,
            handle,
            indent=2,
            ensure_ascii=False,
        )

        handle.write(
            "\n"
        )

    # ========================================================
    # 9. Final output
    # ========================================================

    print()
    print(
        "=" * 72
    )

    print(
        "NEWS COLLECTION COMPLETE"
    )

    print(
        "=" * 72
    )

    print(
        f"RSS entries:              "
        f"{stats['rssEntries']}"
    )

    print(
        f"Unique entries:           "
        f"{stats['uniqueEntries']}"
    )

    print(
        f"Duplicates:               "
        f"{stats['duplicates']}"
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

    print(
        f"Current maximum:          "
        f"{MAX_CURRENT_ARTICLES}"
    )

    print(
        "=" * 72
    )


if __name__ == "__main__":
    main()
