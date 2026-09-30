#!/usr/bin/env python3

"""
VGrat FMS - CNA Market News Collector
=====================================

Collect relevant CNA news for VGrat FMS.

PRIMARY OBJECTIVE
=================

Keep stories that materially relate to:

    - financial markets
    - investing
    - macroeconomics
    - business
    - technology with economic/business significance
    - semiconductors / AI infrastructure
    - commodities / energy
    - economically significant geopolitics

Reject:

    - sports
    - entertainment
    - lifestyle
    - crime / accidents
    - generic political commentary
    - generic diplomacy
    - generic human-interest stories
    - generic weather

ARCHITECTURE
============

RSS
 |
 v
deduplicate
 |
 v
TITLE-FIRST CLASSIFIER
 |
 +--> hard reject obvious non-market categories
 |
 +--> strong economic signal
 |
 +--> strong market signal
 |
 +--> strong technology/business signal
 |
 +--> geopolitical + economic impact
 |
 +--> reject
 |
 v
article extraction
 |
 v
summary
 |
 v
historical archive
 |
 v
rolling 14-day current.json

IMPORTANT
=========

The RSS description can contain unrelated words from the article
body or feed metadata.

Therefore:

    TITLE = highest weight
    DESCRIPTION = supporting evidence

This avoids rejecting genuine economic stories simply because
their descriptions contain words such as "team", "government",
"court", etc.

CURRENT NEWS
============

Maximum current records:

    100

This is a maximum, not a forced target.

Historical records are unlimited.
"""

from __future__ import annotations

import hashlib
import json
import re
from calendar import timegm
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

RUN_SUMMARY_FILE = NEWS_DIR / "run_summary.json"

WINDOW_DAYS = 14

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
# HARD / STRONG EXCLUSION TERMS
# ============================================================
#
# Deliberately narrow.
#
# DO NOT put broad words such as:
#
#     team
#     government
#     court
#     market
#     coach
#
# here because they can appear in legitimate financial stories.
# ============================================================

SPORTS_STRONG = {
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
    "formula 1",
    "formula one",
    "f1",
    "premier league",
    "champions league",
    "world cup",
    "asian games",
    "olympics",
    "olympic games",
    "paralympics",
    "grand slam",
    "bryce",
    "battles for gold",
    "gold medal",
    "silver medal",
    "bronze medal",
    "medal",
    "medals",
    "athlete",
    "athletes",
    "player",
    "players",
    "wallabies",
    "bledisloe",
    "goalkeeper",
    "striker",
    "midfielder",
    "defender",
    "wicket",
    "batting",
    "bowling",
    "tournament",
    "championship",
    "championships",
    "match",
    "matches",
}


ENTERTAINMENT_STRONG = {
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
    "album",
    "albums",
    "red carpet",
    "box office",
    "television series",
    "tv series",
    "reality show",
    "reality tv",
    "movie role",
    "film role",
}


LIFESTYLE_STRONG = {
    "bakery",
    "restaurant",
    "restaurants",
    "cafe",
    "coffee shop",
    "recipe",
    "recipes",
    "fashion",
    "beauty",
    "wellness",
    "dating",
    "wedding",
    "birthday",
    "viral video",
    "influencer",
    "pet",
    "pets",
}


ACCIDENT_STRONG = {
    "car crash",
    "car accident",
    "road accident",
    "traffic accident",
    "plane crash",
    "aircraft crash",
    "train crash",
    "train accident",
    "bus crash",
    "building collapse",
    "house fire",
    "flat fire",
    "fire kills",
    "fire killed",
    "dies after fire",
    "died after fire",
    "killed in crash",
    "killed in accident",
    "swept away",
    "avalanche",
}


CRIME_STRONG = {
    "murder",
    "murdered",
    "homicide",
    "robbery",
    "robbed",
    "burglary",
    "burglars",
    "drug trafficking",
    "drug trafficker",
    "drugs seized",
    "charged with",
    "jailed for",
    "sentenced to jail",
    "sentenced to prison",
    "arrested for",
    "arrested after",
    "police investigation",
    "police said",
    "criminal investigation",
    "crime investigation",
    "assaulted",
}


WEATHER_STRONG = {
    "weather forecast",
    "rainy days",
    "rainfall",
    "heatwave",
    "heat wave",
    "typhoon",
    "hurricane",
    "tornado",
    "avalanche",
}


# ============================================================
# MARKET SIGNALS
# ============================================================

MARKET_SIGNALS = {
    "stock": 6,
    "stocks": 6,
    "stock market": 9,
    "stock markets": 9,
    "shares": 6,
    "share price": 8,
    "share prices": 8,
    "equities": 7,
    "equity market": 8,
    "financial markets": 8,
    "investor": 5,
    "investors": 5,
    "investment": 5,
    "investments": 5,
    "portfolio": 5,
    "portfolios": 5,
    "asset manager": 7,
    "asset managers": 7,
    "asset management": 8,
    "fund manager": 7,
    "fund managers": 7,
    "mutual fund": 7,
    "mutual funds": 7,
    "fundraising": 7,
    "fundraise": 7,
    "funding round": 7,
    "ipo": 9,
    "ipos": 9,
    "initial public offering": 9,
    "listing": 6,
    "listed company": 7,
    "earnings": 7,
    "profit": 6,
    "profits": 6,
    "revenue": 6,
    "revenues": 6,
    "dividend": 6,
    "dividends": 6,
    "merger": 7,
    "mergers": 7,
    "acquisition": 7,
    "acquisitions": 7,
    "takeover": 7,
    "takeovers": 7,
    "valuation": 7,
    "valuations": 7,
    "capital expenditure": 7,
    "capex": 7,
    "credit line": 6,
    "credit lines": 6,
    "banking sector": 8,
    "banking": 7,
    "bank": 5,
    "banks": 5,
    "insurance": 7,
    "insurer": 7,
    "insurers": 7,
    "bond": 7,
    "bonds": 7,
    "bond yields": 9,
    "yield": 6,
    "yields": 6,
    "treasury yields": 9,
    "interest rate": 9,
    "interest rates": 9,
    "rate hike": 9,
    "rate hikes": 9,
    "rate cut": 9,
    "rate cuts": 9,
    "central bank": 9,
    "central banks": 9,
    "monetary policy": 9,
    "forex": 8,
    "foreign exchange": 8,
    "currency market": 8,
    "commodities": 7,
    "commodity prices": 8,
    "oil prices": 9,
    "oil rises": 9,
    "oil gains": 9,
    "oil falls": 9,
    "oil down": 9,
    "gold prices": 8,
    "energy prices": 8,
    "real estate market": 7,
    "property market": 7,
    "property prices": 7,
    "reit": 8,
    "reits": 8,
}


# ============================================================
# ECONOMIC SIGNALS
# ============================================================

ECONOMIC_SIGNALS = {
    "economy": 5,
    "economic": 5,
    "economic growth": 9,
    "economic outlook": 8,
    "gdp": 10,
    "gross domestic product": 10,
    "inflation": 10,
    "deflation": 10,
    "cpi": 10,
    "consumer price index": 10,
    "ppi": 10,
    "producer price index": 10,
    "interest rate": 9,
    "interest rates": 9,
    "central bank": 9,
    "central banks": 9,
    "monetary policy": 9,
    "fiscal policy": 8,
    "economic policy": 8,
    "business confidence": 8,
    "consumer confidence": 8,
    "consumer spending": 8,
    "retail sales": 9,
    "industrial production": 12,
    "industrial output": 12,
    "factory output": 12,
    "factory production": 12,
    "manufacturing output": 12,
    "manufacturing production": 12,
    "manufacturing activity": 9,
    "manufacturing": 6,
    "pmi": 10,
    "purchasing managers": 9,
    "exports": 8,
    "export growth": 10,
    "imports": 8,
    "import growth": 10,
    "trade surplus": 10,
    "trade deficit": 10,
    "trade balance": 9,
    "trade data": 8,
    "tariff": 8,
    "tariffs": 8,
    "trade war": 10,
    "supply chain": 7,
    "employment": 7,
    "unemployment": 9,
    "jobs data": 9,
    "payroll": 8,
    "payrolls": 8,
    "wage growth": 8,
    "productivity": 7,
    "business activity": 7,
    "economic activity": 8,
    "cost of living": 6,
    "electricity tariffs": 10,
    "gas tariffs": 10,
    "utility tariffs": 9,
    "household tariffs": 8,
}


# ============================================================
# TECHNOLOGY / AI / SEMICONDUCTOR SIGNALS
# ============================================================

TECHNOLOGY_SIGNALS = {
    "semiconductor": 10,
    "semiconductors": 10,
    "semiconductor industry": 12,
    "semiconductor sector": 12,
    "chip": 6,
    "chips": 6,
    "chipmaking": 10,
    "chip maker": 10,
    "chipmakers": 10,
    "chip hub": 10,
    "chip hubs": 10,
    "advanced chip": 10,
    "advanced chips": 10,
    "ai chip": 12,
    "ai chips": 12,
    "ai infrastructure": 10,
    "ai infra": 10,
    "data centre": 9,
    "data centres": 9,
    "data center": 9,
    "data centers": 9,
    "cloud computing": 8,
    "technology investment": 9,
    "technology investments": 9,
    "technology funding": 9,
    "technology standards": 9,
    "technology sector": 8,
    "technology company": 6,
    "technology companies": 6,
    "tech company": 6,
    "tech companies": 6,
    "venture capital": 9,
    "private equity": 9,
    "startup funding": 9,
    "funding round": 8,
    "artificial intelligence": 5,
    "ai sector": 8,
    "ai safety": 7,
    "ai policy": 8,
    "ai regulation": 9,
    "ai regulations": 9,
    "ai investment": 10,
    "ai investments": 10,
    "ai spending": 9,
    "ai capex": 10,
    "ai infrastructure": 10,
    "technology stocktake": 9,
    "digital infrastructure": 8,
    "digital economy": 8,
    "nvidia": 6,
    "deepseek": 7,
    "openai": 7,
    "anthropic": 8,
    "tesla": 6,
    "apple": 5,
    "microsoft": 5,
    "google": 5,
    "amazon": 5,
    "meta": 5,
}


# ============================================================
# FINANCIAL / BUSINESS LEGAL SIGNALS
# ============================================================
#
# These are useful for cases such as:
#
#     Goh Jin Hian false trading trial
#
# A legal/criminal term alone should not make it market news,
# but legal proceedings involving listed companies, market
# manipulation, securities or financial markets should qualify.
# ============================================================

FINANCIAL_LEGAL_SIGNALS = {
    "false trading": 10,
    "market manipulation": 12,
    "securities fraud": 12,
    "securities law": 9,
    "insider trading": 12,
    "insider dealing": 12,
    "shareholder": 7,
    "shareholders": 7,
    "listed company": 8,
    "listed companies": 8,
    "financial misconduct": 9,
    "market-making": 10,
    "market making": 10,
    "market maker": 10,
    "market makers": 10,
    "trading services": 7,
    "brokerage": 8,
    "broker": 7,
    "brokers": 7,
    "securities": 8,
    "exchange": 6,
    "stock exchange": 9,
}


# ============================================================
# GEOPOLITICAL SIGNALS
# ============================================================

GEOPOLITICAL_SIGNALS = {
    "sanction": 7,
    "sanctions": 7,
    "export controls": 8,
    "trade restrictions": 8,
    "import restrictions": 8,
    "trade war": 9,
    "tariff": 7,
    "tariffs": 7,
    "shipping disruption": 10,
    "shipping disruptions": 10,
    "shipping route": 7,
    "shipping routes": 7,
    "strait of hormuz": 12,
    "hormuz": 10,
    "taiwan strait": 9,
    "south china sea": 7,
    "rare earth": 9,
    "rare earths": 9,
    "strategic minerals": 9,
    "energy security": 8,
    "oil supply": 10,
    "oil shipments": 10,
    "oil loading": 10,
    "gas supply": 8,
    "pipeline": 7,
    "pipelines": 7,
    "shipping": 5,
    "trade": 5,
    "energy": 5,
    "semiconductor": 9,
    "semiconductors": 9,
}


# ============================================================
# GENERIC POLITICAL TERMS
# ============================================================

POLITICAL_TERMS = {
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
    "resignation",
    "resigns",
    "pledge",
    "pledged",
    "statement",
    "says",
    "said",
}


# ============================================================
# TEXT HELPERS
# ============================================================

def clean_text(value: Any) -> str:

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


def canonicalize_url(url: str) -> str:

    if not url:
        return ""

    try:

        parts = urlsplit(
            url.strip()
        )

        query = []

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

            query.append(item)

        return urlunsplit(
            (
                parts.scheme.lower(),
                parts.netloc.lower(),
                parts.path.rstrip("/"),
                "&".join(query),
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
    ).encode(
        "utf-8"
    )

    return hashlib.sha256(
        raw
    ).hexdigest()


def shorten(
    text: str,
    max_chars: int = 700,
) -> str:

    text = clean_text(
        text
    )

    if len(text) <= max_chars:
        return text

    value = (
        text[:max_chars]
        .rsplit(
            " ",
            1,
        )[0]
        .rstrip(
            " .,;:"
        )
    )

    return value + "..."


def find_matches(
    text: str,
    signals: dict[str, int] | set[str],
) -> list[str]:

    lower = text.lower()

    if isinstance(
        signals,
        dict,
    ):
        values = signals.keys()
    else:
        values = signals

    return [
        signal
        for signal in values
        if signal.lower() in lower
    ]


def score_signals(
    text: str,
    signals: dict[str, int],
) -> tuple[int, list[str]]:

    lower = text.lower()

    score = 0
    matches = []

    for signal, weight in signals.items():

        if signal.lower() in lower:

            score += weight

            matches.append(
                signal
            )

    return score, matches


# ============================================================
# CLASSIFIER
# ============================================================

def classify_relevance(
    title: str,
    description: str,
) -> dict[str, Any]:

    """
    TITLE-FIRST CLASSIFICATION

    Title gets 3x signal weight.

    Description gets 1x signal weight.

    This is important because CNA RSS descriptions may contain
    related but secondary material.

    Decision priority:

        1. obvious sports/entertainment/lifestyle
        2. strong economic
        3. strong market
        4. strong technology/business
        5. financial legal
        6. geopolitical economic impact
        7. reject
    """

    title = clean_text(
        title
    )

    description = clean_text(
        description
    )

    title_lower = title.lower()

    description_lower = description.lower()

    # --------------------------------------------------------
    # Hard category detection
    # --------------------------------------------------------

    sports_title = find_matches(
        title_lower,
        SPORTS_STRONG,
    )

    entertainment_title = find_matches(
        title_lower,
        ENTERTAINMENT_STRONG,
    )

    lifestyle_title = find_matches(
        title_lower,
        LIFESTYLE_STRONG,
    )

    accident_title = find_matches(
        title_lower,
        ACCIDENT_STRONG,
    )

    crime_title = find_matches(
        title_lower,
        CRIME_STRONG,
    )

    weather_title = find_matches(
        title_lower,
        WEATHER_STRONG,
    )

    # --------------------------------------------------------
    # Calculate all positive signals FIRST.
    #
    # This allows a genuine economic story to override generic
    # contextual words.
    # --------------------------------------------------------

    title_market, title_market_matches = score_signals(
        title_lower,
        MARKET_SIGNALS,
    )

    desc_market, desc_market_matches = score_signals(
        description_lower,
        MARKET_SIGNALS,
    )

    title_economic, title_economic_matches = score_signals(
        title_lower,
        ECONOMIC_SIGNALS,
    )

    desc_economic, desc_economic_matches = score_signals(
        description_lower,
        ECONOMIC_SIGNALS,
    )

    title_technology, title_technology_matches = score_signals(
        title_lower,
        TECHNOLOGY_SIGNALS,
    )

    desc_technology, desc_technology_matches = score_signals(
        description_lower,
        TECHNOLOGY_SIGNALS,
    )

    title_legal, title_legal_matches = score_signals(
        title_lower,
        FINANCIAL_LEGAL_SIGNALS,
    )

    desc_legal, desc_legal_matches = score_signals(
        description_lower,
        FINANCIAL_LEGAL_SIGNALS,
    )

    title_geo, title_geo_matches = score_signals(
        title_lower,
        GEOPOLITICAL_SIGNALS,
    )

    desc_geo, desc_geo_matches = score_signals(
        description_lower,
        GEOPOLITICAL_SIGNALS,
    )

    # Title gets 3x importance.
    market_score = (
        title_market * 3
        + desc_market
    )

    economic_score = (
        title_economic * 3
        + desc_economic
    )

    technology_score = (
        title_technology * 3
        + desc_technology
    )

    legal_score = (
        title_legal * 3
        + desc_legal
    )

    geopolitical_score = (
        title_geo * 3
        + desc_geo
    )

    # --------------------------------------------------------
    # 1. SPORTS
    #
    # A clear sports title remains excluded unless it contains
    # an extremely strong financial/business signal.
    # --------------------------------------------------------

    if sports_title:

        combined_financial = (
            market_score
            + economic_score
            + technology_score
            + legal_score
        )

        if combined_financial < 20:

            return {
                "relevant": False,
                "category": "REJECT",
                "score": 0,
                "reason": "sports",
                "matchedKeywords": sports_title,
            }

    # --------------------------------------------------------
    # 2. ENTERTAINMENT
    # --------------------------------------------------------

    if entertainment_title:

        if (
            market_score < 20
            and technology_score < 20
        ):

            return {
                "relevant": False,
                "category": "REJECT",
                "score": 0,
                "reason": "entertainment",
                "matchedKeywords": entertainment_title,
            }

    # --------------------------------------------------------
    # 3. LIFESTYLE
    # --------------------------------------------------------

    if lifestyle_title:

        if (
            market_score < 20
            and economic_score < 20
        ):

            return {
                "relevant": False,
                "category": "REJECT",
                "score": 0,
                "reason": "lifestyle",
                "matchedKeywords": lifestyle_title,
            }

    # --------------------------------------------------------
    # 4. ACCIDENT
    # --------------------------------------------------------

    if accident_title:

        if (
            market_score < 25
            and economic_score < 25
        ):

            return {
                "relevant": False,
                "category": "REJECT",
                "score": 0,
                "reason": "accident_human_interest",
                "matchedKeywords": accident_title,
            }

    # --------------------------------------------------------
    # 5. WEATHER
    # --------------------------------------------------------

    if weather_title:

        if (
            economic_score < 20
            and market_score < 20
        ):

            return {
                "relevant": False,
                "category": "REJECT",
                "score": 0,
                "reason": "weather",
                "matchedKeywords": weather_title,
            }

    # --------------------------------------------------------
    # 6. ECONOMIC
    #
    # Strong title signals alone are enough.
    #
    # This specifically fixes:
    #
    # Thai factory output
    # Japan factory output
    # semiconductor ambitions
    # tariff / inflation / GDP stories
    # --------------------------------------------------------

    if economic_score >= 20:

        return {
            "relevant": True,
            "category": "ECONOMIC",
            "score": economic_score,
            "reason": "strong_economic_signal",
            "matchedKeywords": (
                title_economic_matches
                + desc_economic_matches
            ),
        }

    # --------------------------------------------------------
    # 7. MARKET
    # --------------------------------------------------------

    if market_score >= 18:

        return {
            "relevant": True,
            "category": "MARKET",
            "score": market_score,
            "reason": "strong_market_signal",
            "matchedKeywords": (
                title_market_matches
                + desc_market_matches
            ),
        }

    # --------------------------------------------------------
    # 8. TECHNOLOGY
    #
    # Generic AI commentary is not enough.
    #
    # Business / infrastructure / semiconductor / funding /
    # policy / investment signals are required.
    # --------------------------------------------------------

    strong_technology_terms = {
        "semiconductor",
        "semiconductors",
        "semiconductor industry",
        "semiconductor sector",
        "chipmaking",
        "chip maker",
        "chipmakers",
        "chip hub",
        "chip hubs",
        "advanced chip",
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
        "funding round",
        "technology investment",
        "technology investments",
        "technology funding",
        "technology standards",
        "technology sector",
        "ai sector",
        "ai policy",
        "ai regulation",
        "ai regulations",
        "ai investment",
        "ai investments",
        "ai spending",
        "ai capex",
        "technology stocktake",
        "digital infrastructure",
        "digital economy",
    }

    title_strong_tech = [
        item
        for item in title_technology_matches
        if item in strong_technology_terms
    ]

    desc_strong_tech = [
        item
        for item in desc_technology_matches
        if item in strong_technology_terms
    ]

    if (
        technology_score >= 15
        and (
            title_strong_tech
            or desc_strong_tech
        )
    ):

        return {
            "relevant": True,
            "category": "TECHNOLOGY",
            "score": technology_score,
            "reason": "technology_business_or_infrastructure",
            "matchedKeywords": (
                title_technology_matches
                + desc_technology_matches
            ),
        }

    # --------------------------------------------------------
    # 9. FINANCIAL / BUSINESS LEGAL
    #
    # This fixes stories such as:
    #
    # Goh Jin Hian false trading trial
    #
    # while still rejecting ordinary criminal cases.
    # --------------------------------------------------------

    if legal_score >= 18:

        return {
            "relevant": True,
            "category": "MARKET",
            "score": legal_score,
            "reason": "financial_market_legal_story",
            "matchedKeywords": (
                title_legal_matches
                + desc_legal_matches
            ),
        }

    # --------------------------------------------------------
    # 10. GEOPOLITICAL + ECONOMIC IMPACT
    #
    # Country names and military activity alone are NOT enough.
    #
    # Require a direct economic transmission mechanism:
    #
    # oil
    # shipping
    # sanctions
    # trade
    # semiconductors
    # energy
    # strategic minerals
    # etc.
    # --------------------------------------------------------

    economic_geo_terms = {
        "sanction",
        "sanctions",
        "export controls",
        "trade restrictions",
        "import restrictions",
        "trade war",
        "tariff",
        "tariffs",
        "shipping disruption",
        "shipping disruptions",
        "shipping route",
        "shipping routes",
        "strait of hormuz",
        "hormuz",
        "taiwan strait",
        "rare earth",
        "rare earths",
        "strategic minerals",
        "energy security",
        "oil supply",
        "oil shipments",
        "oil loading",
        "gas supply",
        "pipeline",
        "pipelines",
        "shipping",
        "trade",
        "energy",
        "semiconductor",
        "semiconductors",
    }

    geo_economic_matches = [
        item
        for item in (
            title_geo_matches
            + desc_geo_matches
        )
        if item in economic_geo_terms
    ]

    if (
        geopolitical_score >= 18
        and geo_economic_matches
    ):

        return {
            "relevant": True,
            "category": "GEOPOLITICAL",
            "score": geopolitical_score,
            "reason": "geopolitical_with_direct_economic_impact",
            "matchedKeywords": (
                title_geo_matches
                + desc_geo_matches
            ),
        }

    # --------------------------------------------------------
    # 11. GENERIC POLITICAL / DIPLOMATIC
    # --------------------------------------------------------

    political_matches = find_matches(
        title_lower,
        POLITICAL_TERMS,
    )

    if political_matches:

        return {
            "relevant": False,
            "category": "REJECT",
            "score": 0,
            "reason": "political_without_market_or_economic_impact",
            "matchedKeywords": political_matches,
        }

    # --------------------------------------------------------
    # 12. CRIME
    #
    # A crime word alone in the description should NOT reject
    # an otherwise genuine market story.
    #
    # However, a clearly crime-focused title with no financial
    # relevance is rejected.
    # --------------------------------------------------------

    if crime_title:

        return {
            "relevant": False,
            "category": "REJECT",
            "score": 0,
            "reason": "crime_human_interest",
            "matchedKeywords": crime_title,
        }

    # --------------------------------------------------------
    # DEFAULT
    # --------------------------------------------------------

    return {
        "relevant": False,
        "category": "REJECT",
        "score": 0,
        "reason": "insufficient_market_relevance",
        "matchedKeywords": [],
    }


# ============================================================
# RSS
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

    if parsed:

        try:

            return datetime.fromtimestamp(
                timegm(parsed),
                tz=timezone.utc,
            )

        except Exception:
            pass

    return datetime.now(
        timezone.utc
    )


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

            published = parse_entry_datetime(
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
                    "publishedAtUtc": (
                        published
                        .astimezone(
                            timezone.utc
                        )
                        .replace(
                            microsecond=0
                        )
                        .isoformat()
                        .replace(
                            "+00:00",
                            "Z",
                        )
                    ),
                }
            )

        print(
            f"  Entries: {len(records)}"
        )

        return records

    except Exception as exc:

        print(
            f"  ERROR: "
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

        value = clean_text(
            paragraph.get_text(
                " ",
                strip=True,
            )
        )

        if len(value) >= 40:

            paragraphs.append(
                value
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

    return [
        item.strip()
        for item in re.split(
            r"(?<=[.!?])\s+",
            text,
        )
        if len(item.strip()) >= 45
    ]


def generate_summary(
    title: str,
    description: str,
    article_text: str,
) -> str:

    sentences = split_sentences(
        article_text
    )

    if not sentences:

        fallback = (
            description
            or title
        )

        return shorten(
            fallback
        )

    important_terms = [
        "stock",
        "stocks",
        "shares",
        "equities",
        "investor",
        "investors",
        "investment",
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
        sentences
    ):

        lower = sentence.lower()

        score = sum(
            1
            for term in important_terms
            if term in lower
        )

        scored.append(
            (
                score,
                -index,
                sentence,
            )
        )

    scored.sort(
        key=lambda item: (
            item[0],
            item[1],
        ),
        reverse=True,
    )

    selected = {
        item[2]
        for item in scored[:3]
    }

    ordered = [
        sentence
        for sentence in sentences
        if sentence in selected
    ][:3]

    return shorten(
        " ".join(
            ordered
        )
    )


# ============================================================
# JSON
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
            f"WARNING: Failed reading "
            f"{path}: "
            f"{type(exc).__name__}: {exc}"
        )

        return default


# ============================================================
# HISTORY
# ============================================================

def load_existing_history_ids() -> set[str]:

    result = set()

    if not HISTORY_DIR.exists():
        return result

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

                value = article.get(
                    "id"
                )

                if value:
                    result.add(
                        value
                    )

    return result


def load_all_history_articles() -> list[dict[str, Any]]:

    result = []

    if not HISTORY_DIR.exists():
        return result

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

                result.append(
                    article
                )

    return result


def history_file_for_date(
    published_at: str,
) -> Path:

    try:

        dt = datetime.fromisoformat(
            published_at.replace(
                "Z",
                "+00:00",
            )
        )

    except Exception:

        dt = datetime.now(
            timezone.utc
        )

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
            "date": article[
                "publishedAtUtc"
            ][:10],
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
    ] = datetime.now(
        timezone.utc
    ).replace(
        microsecond=0
    ).isoformat().replace(
        "+00:00",
        "Z",
    )

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
# CURRENT INDEX
# ============================================================

def build_current_index(
    articles: list[dict[str, Any]],
) -> list[dict[str, Any]]:

    cutoff = (
        datetime.now(
            timezone.utc
        )
        - timedelta(
            days=WINDOW_DAYS
        )
    )

    unique = {}

    for article in articles:

        try:

            published = datetime.fromisoformat(
                article[
                    "publishedAtUtc"
                ].replace(
                    "Z",
                    "+00:00",
                )
            )

        except Exception:

            continue

        if published < cutoff:
            continue

        article_key = article.get(
            "id"
        )

        if article_key:

            unique[
                article_key
            ] = article

    current = list(
        unique.values()
    )

    current.sort(
        key=lambda item: item.get(
            "publishedAtUtc",
            "",
        ),
        reverse=True,
    )

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
        "generatedAtUtc": (
            datetime.now(
                timezone.utc
            )
            .replace(
                microsecond=0
            )
            .isoformat()
            .replace(
                "+00:00",
                "Z",
            )
        ),
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

    print("=" * 80)

    print(
        "VGrat FMS - CNA Market News Collector"
    )

    print("=" * 80)

    print(
        f"Rolling window: "
        f"{WINDOW_DAYS} days"
    )

    print(
        f"Maximum current records: "
        f"{MAX_CURRENT_ARTICLES}"
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

    started = datetime.now(
        timezone.utc
    )

    stats = {
        "rssEntries": 0,
        "uniqueEntries": 0,
        "duplicates": 0,
        "relevantMarket": 0,
        "relevantEconomic": 0,
        "relevantTechnology": 0,
        "relevantGeopolitical": 0,
        "relevantArticles": 0,
        "rejectedSports": 0,
        "rejectedEntertainment": 0,
        "rejectedLifestyle": 0,
        "rejectedAccident": 0,
        "rejectedCrime": 0,
        "rejectedPolitical": 0,
        "rejectedGeneral": 0,
        "articleExtractionAttempts": 0,
        "articleExtractionSuccess": 0,
        "articleExtractionFailures": 0,
        "newHistoricalRecords": 0,
    }

    # ========================================================
    # RSS
    # ========================================================

    rss_records = collect_rss()

    stats[
        "rssEntries"
    ] = len(
        rss_records
    )

    # ========================================================
    # DEDUPLICATION
    # ========================================================

    unique_records = {}

    for record in rss_records:

        url = record.get(
            "url",
            "",
        )

        if not url:
            continue

        key = article_id(
            SOURCE_NAME,
            url,
        )

        if key in unique_records:

            stats[
                "duplicates"
            ] += 1

            feeds = unique_records[
                key
            ].setdefault(
                "feeds",
                [],
            )

            feed = record.get(
                "feed"
            )

            if (
                feed
                and feed not in feeds
            ):

                feeds.append(
                    feed
                )

            continue

        record[
            "id"
        ] = key

        record[
            "feeds"
        ] = [
            record.get(
                "feed",
                "unknown",
            )
        ]

        unique_records[
            key
        ] = record

    stats[
        "uniqueEntries"
    ] = len(
        unique_records
    )

    print()
    print(
        f"RSS entries: "
        f"{stats['rssEntries']}"
    )

    print(
        f"Unique entries: "
        f"{stats['uniqueEntries']}"
    )

    print(
        f"Duplicates: "
        f"{stats['duplicates']}"
    )

    # ========================================================
    # CLASSIFICATION
    # ========================================================

    relevant_records = []

    print()
    print("=" * 80)
    print("CLASSIFICATION")
    print("=" * 80)

    for record in unique_records.values():

        result = classify_relevance(
            record[
                "title"
            ],
            record[
                "description"
            ],
        )

        record[
            "classification"
        ] = result

        title = record[
            "title"
        ]

        if result[
            "relevant"
        ]:

            category = result[
                "category"
            ]

            stats[
                "relevantArticles"
            ] += 1

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

            relevant_records.append(
                record
            )

            print(
                "KEEP:",
                title,
                f"[{category}]",
                f"score={result['score']}",
                f"reason={result['reason']}",
            )

        else:

            reason = result[
                "reason"
            ]

            if reason == "sports":
                stats[
                    "rejectedSports"
                ] += 1

            elif reason == "entertainment":
                stats[
                    "rejectedEntertainment"
                ] += 1

            elif reason == "lifestyle":
                stats[
                    "rejectedLifestyle"
                ] += 1

            elif reason == "accident_human_interest":
                stats[
                    "rejectedAccident"
                ] += 1

            elif reason == "crime_human_interest":
                stats[
                    "rejectedCrime"
                ] += 1

            elif reason == (
                "political_without_market_or_economic_impact"
            ):
                stats[
                    "rejectedPolitical"
                ] += 1

            else:
                stats[
                    "rejectedGeneral"
                ] += 1

            print(
                "REJECT:",
                title,
                f"reason={reason}",
            )

    # ========================================================
    # EXISTING HISTORY
    # ========================================================

    existing_ids = (
        load_existing_history_ids()
    )

    # ========================================================
    # ARTICLE EXTRACTION
    # ========================================================

    for record in relevant_records:

        key = record[
            "id"
        ]

        # Already permanently archived.
        if key in existing_ids:
            continue

        title = record[
            "title"
        ]

        print()
        print(
            "PROCESS:",
            title
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
                    "No article text extracted"
                )

            stats[
                "articleExtractionSuccess"
            ] += 1

            summary = generate_summary(
                title,
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
                ]
            )

            if not summary:
                summary = title

        classification = record[
            "classification"
        ]

        permanent_record = {
            "id": key,
            "source": SOURCE_NAME,
            "title": title,
            "publishedAtUtc": record[
                "publishedAtUtc"
            ],
            "url": record[
                "url"
            ],
            "category": classification[
                "category"
            ],
            "relevanceScore": classification[
                "score"
            ],
            "relevanceReason": classification[
                "reason"
            ],
            "summary": summary,
            "feeds": record.get(
                "feeds",
                [],
            ),
            "collectedAtUtc": (
                datetime.now(
                    timezone.utc
                )
                .replace(
                    microsecond=0
                )
                .isoformat()
                .replace(
                    "+00:00",
                    "Z",
                )
            ),
        }

        if save_history_article(
            permanent_record
        ):

            stats[
                "newHistoricalRecords"
            ] += 1

            existing_ids.add(
                key
            )

            print(
                "  Saved."
            )

    # ========================================================
    # CURRENT INDEX
    # ========================================================

    history_articles = (
        load_all_history_articles()
    )

    current_articles = (
        build_current_index(
            history_articles
        )
    )

    write_current_index(
        current_articles
    )

    # ========================================================
    # RUN SUMMARY
    # ========================================================

    completed = datetime.now(
        timezone.utc
    )

    run_summary = {
        "status": "success",
        "source": SOURCE_NAME,
        "startedAtUtc": (
            started
            .replace(
                microsecond=0
            )
            .isoformat()
            .replace(
                "+00:00",
                "Z",
            )
        ),
        "completedAtUtc": (
            completed
            .replace(
                microsecond=0
            )
            .isoformat()
            .replace(
                "+00:00",
                "Z",
            )
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
            history_articles
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
    # FINAL
    # ========================================================

    print()
    print("=" * 80)
    print("COLLECTION COMPLETE")
    print("=" * 80)

    print(
        f"Relevant:              "
        f"{stats['relevantArticles']}"
    )

    print(
        f"  Market:              "
        f"{stats['relevantMarket']}"
    )

    print(
        f"  Economic:            "
        f"{stats['relevantEconomic']}"
    )

    print(
        f"  Technology:          "
        f"{stats['relevantTechnology']}"
    )

    print(
        f"  Geopolitical:        "
        f"{stats['relevantGeopolitical']}"
    )

    print()

    print(
        f"Rejected sports:       "
        f"{stats['rejectedSports']}"
    )

    print(
        f"Rejected entertainment:"
        f" {stats['rejectedEntertainment']}"
    )

    print(
        f"Rejected lifestyle:    "
        f"{stats['rejectedLifestyle']}"
    )

    print(
        f"Rejected accidents:    "
        f"{stats['rejectedAccident']}"
    )

    print(
        f"Rejected crime:        "
        f"{stats['rejectedCrime']}"
    )

    print(
        f"Rejected political:    "
        f"{stats['rejectedPolitical']}"
    )

    print(
        f"Rejected general:      "
        f"{stats['rejectedGeneral']}"
    )

    print()

    print(
        f"New historical:        "
        f"{stats['newHistoricalRecords']}"
    )

    print(
        f"Historical total:      "
        f"{len(history_articles)}"
    )

    print(
        f"Current 14-day total:  "
        f"{len(current_articles)}"
    )

    print(
        f"Current maximum:       "
        f"{MAX_CURRENT_ARTICLES}"
    )

    print("=" * 80)


if __name__ == "__main__":
    main()
