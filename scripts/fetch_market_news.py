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
    - financial-market legal/regulatory matters

Reject:

    - sports
    - entertainment
    - lifestyle
    - crime / accidents / human interest
    - generic political commentary
    - generic diplomacy
    - generic education/social policy
    - generic weather
    - generic technology commentary without economic/business relevance

CLASSIFICATION ARCHITECTURE
===========================

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
 +--> hard financial-market legal cases
 |
 +--> strong economic title
 |
 +--> strong market title
 |
 +--> strong technology/business title
 |
 +--> geopolitical title + direct economic mechanism
 |
 +--> description supporting evidence
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

The headline is the primary classification source.

The RSS description is supporting evidence only.

This prevents unrelated words appearing in an article description
from changing the meaning of an otherwise clear headline.

CURRENT NEWS
============

Maximum current records:

    100

This is a maximum, not a forced target.

Historical records are unlimited.

FINAL CATEGORIES
================

    MARKET
    ECONOMIC
    TECHNOLOGY
    GEOPOLITICAL
    REJECT
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


# ============================================================================
# CONFIGURATION
# ============================================================================

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


# ============================================================================
# CNA RSS FEEDS
# ============================================================================

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


# ============================================================================
# HARD EXCLUSIONS
# ============================================================================
#
# These are deliberately title-focused.
#
# A story should not be rejected merely because the description contains
# one of these words.
# ============================================================================

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
    "tournaments",
    "championship",
    "championships",
    "match",
    "matches",
    "squad",
    "international football",
    "international rugby",
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
    "festival",
    "writers festival",
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
    "food outlet",
    "opening first",
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
    "landslide",
    "drowning",
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
    "drug seized",
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
    "stole",
    "stolen",
    "thieves",
    "theft",
    "jail for",
    "prison for",
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
    "rain brings relief",
    "drought",
}


# ============================================================================
# STRONG ECONOMIC SIGNALS
# ============================================================================
#
# These are deliberately explicit.
#
# A direct macroeconomic indicator in the headline is sufficient by itself.
# ============================================================================

ECONOMIC_STRONG = {
    "gdp",
    "gross domestic product",
    "inflation",
    "deflation",
    "consumer price index",
    "producer price index",
    "cpi",
    "ppi",
    "pmi",
    "purchasing managers index",
    "economic growth",
    "economic contraction",
    "economic recession",
    "economic outlook",
    "economic activity",
    "economic data",
    "economic indicators",
    "industrial production",
    "industrial output",
    "factory output",
    "factory production",
    "manufacturing output",
    "manufacturing production",
    "manufacturing activity",
    "manufacturing sector",
    "retail sales",
    "consumer spending",
    "consumer confidence",
    "business confidence",
    "employment",
    "unemployment",
    "unemployment rate",
    "jobs data",
    "payroll",
    "payrolls",
    "wage growth",
    "productivity",
    "exports",
    "export growth",
    "imports",
    "import growth",
    "trade surplus",
    "trade deficit",
    "trade balance",
    "trade data",
    "industrial activity",
    "business activity",
    "cost of living",
    "electricity tariffs",
    "gas tariffs",
    "utility tariffs",
    "household tariffs",
    "energy prices",
}


ECONOMIC_SUPPORTING = {
    "economy",
    "economic",
    "growth",
    "output",
    "production",
    "manufacturing",
    "inflation",
    "prices",
    "consumer",
    "business",
    "trade",
    "exports",
    "imports",
    "employment",
    "unemployment",
    "wages",
    "productivity",
    "tariff",
    "tariffs",
}


# ============================================================================
# MARKET SIGNALS
# ============================================================================

MARKET_STRONG = {
    "stock market",
    "stock markets",
    "stocks",
    "shares",
    "share price",
    "share prices",
    "equities",
    "equity market",
    "financial markets",
    "investor",
    "investors",
    "investment",
    "investments",
    "portfolio",
    "portfolios",
    "asset manager",
    "asset managers",
    "asset management",
    "fund manager",
    "fund managers",
    "mutual fund",
    "mutual funds",
    "fundraising",
    "fundraise",
    "funding",
    "funding round",
    "raises",
    "raised",
    "raise",
    "ipo",
    "ipos",
    "initial public offering",
    "listing",
    "listed company",
    "listed companies",
    "earnings",
    "profit",
    "profits",
    "revenue",
    "revenues",
    "dividend",
    "dividends",
    "merger",
    "mergers",
    "acquisition",
    "acquisitions",
    "takeover",
    "takeovers",
    "valuation",
    "valuations",
    "capital expenditure",
    "capex",
    "credit line",
    "credit lines",
    "banking sector",
    "banking",
    "bank",
    "banks",
    "insurance",
    "insurer",
    "insurers",
    "bond",
    "bonds",
    "bond yields",
    "yield",
    "yields",
    "treasury yields",
    "interest rate",
    "interest rates",
    "rate hike",
    "rate hikes",
    "rate cut",
    "rate cuts",
    "central bank",
    "central banks",
    "monetary policy",
    "forex",
    "foreign exchange",
    "currency market",
    "commodities",
    "commodity prices",
    "oil prices",
    "oil rises",
    "oil gains",
    "oil falls",
    "oil down",
    "oil rises",
    "gold prices",
    "real estate market",
    "property market",
    "property prices",
    "reit",
    "reits",
    "payments",
    "payment",
    "digital payments",
    "mobile payments",
    "fintech",
}


MARKET_SUPPORTING = {
    "market",
    "markets",
    "finance",
    "financial",
    "bank",
    "banking",
    "investment",
    "investor",
    "company",
    "companies",
    "fund",
    "funds",
    "capital",
    "valuation",
    "business",
}


# ============================================================================
# TECHNOLOGY SIGNALS
# ============================================================================

TECHNOLOGY_STRONG = {
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
    "cloud computing",
    "technology investment",
    "technology investments",
    "technology funding",
    "technology standards",
    "technology sector",
    "technology company",
    "technology companies",
    "tech company",
    "tech companies",
    "tech executives",
    "technology executives",
    "venture capital",
    "private equity",
    "startup funding",
    "funding round",
    "artificial intelligence",
    "ai sector",
    "ai safety",
    "ai policy",
    "ai regulation",
    "ai regulations",
    "ai investment",
    "ai investments",
    "ai spending",
    "ai capex",
    "ai accord",
    "ai agreement",
    "ai pact",
    "ai safety pact",
    "ai document",
    "ai standards",
    "technology stocktake",
    "digital infrastructure",
    "digital economy",
    "digital payments",
    "mobile payments",
    "fintech",
    "chip programming",
    "chip programming tools",
    "chip development",
    "chip development tools",
    "nvidia",
    "deepseek",
    "openai",
    "anthropic",
    "tesla",
    "apple",
    "microsoft",
    "google",
    "amazon",
    "meta",
    "apple pay",
    "ai hack",
}


TECHNOLOGY_SUPPORTING = {
    "technology",
    "tech",
    "artificial intelligence",
    "ai",
    "digital",
    "software",
    "cyber",
    "cybersecurity",
    "chip",
    "chips",
    "semiconductor",
    "semiconductors",
    "cloud",
    "data centre",
    "data center",
}


# ============================================================================
# FINANCIAL LEGAL / REGULATORY SIGNALS
# ============================================================================
#
# These are MARKET stories even if the article is framed as a court case.
# ============================================================================

FINANCIAL_LEGAL_STRONG = {
    "false trading",
    "market manipulation",
    "securities fraud",
    "securities law",
    "insider trading",
    "insider dealing",
    "shareholder",
    "shareholders",
    "listed company",
    "listed companies",
    "financial misconduct",
    "market-making",
    "market making",
    "market maker",
    "market makers",
    "trading services",
    "brokerage",
    "broker",
    "brokers",
    "securities",
    "stock exchange",
    "exchange operator",
    "financial regulator",
    "financial regulators",
    "banking regulator",
    "banking regulators",
    "market regulator",
    "market regulators",
}


# ============================================================================
# GEOPOLITICAL ECONOMIC-IMPACT SIGNALS
# ============================================================================
#
# A geopolitical story is only accepted when there is a concrete economic,
# financial, energy, trade, shipping, or strategic-technology mechanism.
# ============================================================================

GEOPOLITICAL_ECONOMIC_STRONG = {
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
    "shipping lanes",
    "strait of hormuz",
    "hormuz",
    "taiwan strait",
    "south china sea",
    "rare earth",
    "rare earths",
    "strategic minerals",
    "energy security",
    "oil supply",
    "oil shipments",
    "oil loading",
    "gas supply",
    "gas shipments",
    "pipeline",
    "pipelines",
    "energy supply",
    "energy exports",
    "energy imports",
    "oil exports",
    "oil imports",
    "trade flows",
    "trade",
    "energy",
    "semiconductor",
    "semiconductors",
    "chip supply",
    "chip supplies",
    "chip exports",
    "chip imports",
}


# ============================================================================
# GENERIC POLITICAL TERMS
# ============================================================================

POLITICAL_TERMS = {
    "president",
    "prime minister",
    "minister",
    "ministers",
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
    "government's",
    "diplomatic",
    "diplomacy",
    "ambassador",
    "bilateral",
    "summit",
    "resignation",
    "resigns",
    "resigned",
    "pledge",
    "pledged",
    "statement",
    "statements",
    "commentary",
    "comment",
}


# ============================================================================
# GENERIC NON-ECONOMIC POLICY TOPICS
# ============================================================================

NON_ECONOMIC_POLICY_TERMS = {
    "psle",
    "primary school leaving examination",
    "dsa",
    "direct school admission",
    "school admission",
    "school admissions",
    "education policy",
    "education system",
    "school system",
    "hdb household",
    "hdb households",
    "u-save",
    "s&cc",
    "s&cc rebates",
    "rebates",
}


# ============================================================================
# TEXT HELPERS
# ============================================================================

def clean_text(value: Any) -> str:
    """
    Convert HTML/text input to normalized plain text.
    """
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
    """
    Remove common tracking parameters while preserving meaningful query data.
    """
    if not url:
        return ""

    try:
        parts = urlsplit(url.strip())

        query = []

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
    """
    Stable SHA-256 identifier based on source + canonical URL.
    """
    raw = (
        f"{source}|{canonicalize_url(url)}"
    ).encode(
        "utf-8"
    )

    return hashlib.sha256(raw).hexdigest()


def shorten(
    text: str,
    max_chars: int = 700,
) -> str:
    """
    Keep generated summaries compact.
    """
    text = clean_text(text)

    if len(text) <= max_chars:
        return text

    value = text[:max_chars]

    value = value.rsplit(
        " ",
        1,
    )[0].rstrip(
        " .,;:"
    )

    return value + "..."


def find_matches(
    text: str,
    signals: dict[str, int] | set[str],
) -> list[str]:
    """
    Return all signal phrases found in text.
    """
    lower = text.lower()

    if isinstance(signals, dict):
        values = signals.keys()
    else:
        values = signals

    return [
        signal
        for signal in values
        if signal.lower() in lower
    ]


def count_matches(
    text: str,
    signals: set[str] | dict[str, int],
) -> list[str]:
    """
    Alias used for semantic readability.
    """
    return find_matches(
        text,
        signals,
    )


# ============================================================================
# TITLE CLASSIFICATION HELPERS
# ============================================================================

def is_obvious_sports_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        SPORTS_STRONG,
    )


def is_obvious_entertainment_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        ENTERTAINMENT_STRONG,
    )


def is_obvious_lifestyle_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        LIFESTYLE_STRONG,
    )


def is_obvious_accident_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        ACCIDENT_STRONG,
    )


def is_obvious_crime_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        CRIME_STRONG,
    )


def is_obvious_weather_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        WEATHER_STRONG,
    )


def is_non_economic_policy_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        NON_ECONOMIC_POLICY_TERMS,
    )


def has_financial_context(
    title: str,
) -> bool:
    """
    Determine whether an otherwise criminal/legal headline is actually
    about a financial-market matter.

    Examples:

        false trading
        market manipulation
        insider trading
        securities fraud
        market-making
        shareholder dispute
    """
    matches = count_matches(
        title,
        FINANCIAL_LEGAL_STRONG,
    )

    if matches:
        return True

    return bool(
        count_matches(
            title,
            {
                "stock",
                "stocks",
                "shares",
                "share price",
                "market",
                "securities",
                "trading",
                "investor",
                "investors",
                "bank",
                "banking",
                "fund",
                "funds",
            },
        )
    )


def has_strong_economic_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        ECONOMIC_STRONG,
    )


def has_strong_market_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        MARKET_STRONG,
    )


def has_strong_technology_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        TECHNOLOGY_STRONG,
    )


def has_strong_geopolitical_economic_title(
    title: str,
) -> list[str]:
    return count_matches(
        title,
        GEOPOLITICAL_ECONOMIC_STRONG,
    )


# ============================================================================
# TECHNOLOGY CONTEXT RULES
# ============================================================================

def technology_business_context(
    title: str,
) -> bool:
    """
    Determine whether a technology headline has a concrete business,
    investment, infrastructure, semiconductor, payments, regulatory,
    or economic angle.

    This deliberately rejects generic AI commentary.
    """

    lower = title.lower()

    direct_contexts = [
        (
            {
                "apple pay",
                "digital payments",
                "mobile payments",
                "fintech",
            },
            {
                "bank",
                "banks",
                "payment",
                "payments",
                "launches",
                "partnership",
                "partnerships",
            },
        ),
        (
            {
                "deepseek",
                "huawei",
                "nvidia",
            },
            {
                "chip",
                "chips",
                "semiconductor",
                "semiconductors",
                "chipmaking",
                "programming",
                "technology",
            },
        ),
        (
            {
                "ai",
                "artificial intelligence",
            },
            {
                "accord",
                "agreement",
                "pact",
                "standards",
                "regulation",
                "regulations",
                "policy",
                "investment",
                "investments",
                "funding",
                "fund",
                "valuation",
                "ipo",
                "capex",
                "data centre",
                "data center",
                "infrastructure",
                "executives",
                "ceos",
                "safety pact",
                "safety",
                "technology standards",
            },
        ),
        (
            {
                "chip",
                "chips",
                "semiconductor",
                "semiconductors",
            },
            {
                "industry",
                "sector",
                "hub",
                "hubs",
                "race",
                "manufacturing",
                "manufacture",
                "production",
                "exports",
                "imports",
                "supply",
                "programming",
                "tools",
                "investment",
                "investments",
            },
        ),
    ]

    for subject_terms, context_terms in direct_contexts:
        if any(
            term in lower
            for term in subject_terms
        ) and any(
            term in lower
            for term in context_terms
        ):
            return True

    return False


# ============================================================================
# FINANCIAL MARKET CONTEXT
# ============================================================================

def market_business_context(
    title: str,
) -> bool:
    """
    Detect concrete financial-market/business events.
    """

    lower = title.lower()

    strong_pairs = [
        (
            {
                "funding",
                "fundraise",
                "fundraising",
                "raised",
                "raises",
                "raise",
            },
            {
                "$",
                "billion",
                "million",
                "valuation",
                "ipo",
                "invest",
                "investment",
                "fund",
            },
        ),
        (
            {
                "bank",
                "banks",
                "banking",
            },
            {
                "regulator",
                "regulators",
                "sector",
                "market",
                "customers",
                "customer impacts",
                "ai use",
            },
        ),
        (
            {
                "payment",
                "payments",
                "apple pay",
                "fintech",
            },
            {
                "bank",
                "banks",
                "partnership",
                "launches",
                "market",
            },
        ),
    ]

    for event_terms, context_terms in strong_pairs:
        if any(
            term in lower
            for term in event_terms
        ) and any(
            term in lower
            for term in context_terms
        ):
            return True

    return False


# ============================================================================
# MAIN CLASSIFIER
# ============================================================================

def classify_relevance(
    title: str,
    description: str,
) -> dict[str, Any]:
    """
    Classify one CNA article.

    Classification priority:

        1. hard exclusions
        2. financial-market legal
        3. strong economic title
        4. strong market title
        5. strong technology/business title
        6. geopolitical + economic mechanism
        7. supporting description evidence
        8. reject

    The title is always the primary evidence.
    """

    title = clean_text(title)
    description = clean_text(description)

    title_lower = title.lower()
    description_lower = description.lower()

    # ------------------------------------------------------------------
    # 1. HARD EXCLUSIONS
    # ------------------------------------------------------------------

    sports_matches = is_obvious_sports_title(
        title_lower
    )

    entertainment_matches = is_obvious_entertainment_title(
        title_lower
    )

    lifestyle_matches = is_obvious_lifestyle_title(
        title_lower
    )

    accident_matches = is_obvious_accident_title(
        title_lower
    )

    crime_matches = is_obvious_crime_title(
        title_lower
    )

    weather_matches = is_obvious_weather_title(
        title_lower
    )

    policy_matches = is_non_economic_policy_title(
        title_lower
    )

    financial_legal_matches = count_matches(
        title_lower,
        FINANCIAL_LEGAL_STRONG,
    )

    # Financial/legal stories override generic crime/legal exclusions.
    if financial_legal_matches:
        return {
            "relevant": True,
            "category": "MARKET",
            "score": 100,
            "reason": "financial_market_legal_story",
            "matchedKeywords": financial_legal_matches,
        }

    # Sports are rejected unless the title contains a genuine financial
    # market event. This protects against words like "player investment"
    # in unrelated sports stories.
    if sports_matches:
        market_title_matches = has_strong_market_title(
            title_lower
        )

        if not market_title_matches:
            return {
                "relevant": False,
                "category": "REJECT",
                "score": 0,
                "reason": "sports",
                "matchedKeywords": sports_matches,
            }

    # Entertainment stories are rejected unless there is an explicit
    # market/technology/business event.
    if entertainment_matches:
        market_title_matches = has_strong_market_title(
            title_lower
        )

        technology_title_matches = has_strong_technology_title(
            title_lower
        )

        if not (
            market_title_matches
            or technology_title_matches
            or market_business_context(title_lower)
        ):
            return {
                "relevant": False,
                "category": "REJECT",
                "score": 0,
                "reason": "entertainment",
                "matchedKeywords": entertainment_matches,
            }

    # Lifestyle stories are normally excluded.
    if lifestyle_matches:
        market_title_matches = has_strong_market_title(
            title_lower
        )

        technology_title_matches = has_strong_technology_title(
            title_lower
        )

        if not (
            market_title_matches
            or technology_title_matches
            or market_business_context(title_lower)
        ):
            return {
                "relevant": False,
                "category": "REJECT",
                "score": 0,
                "reason": "lifestyle",
                "matchedKeywords": lifestyle_matches,
            }

    # Accident/human-interest stories are excluded before description
    # scoring can introduce irrelevant economic words.
    if accident_matches:
        return {
            "relevant": False,
            "category": "REJECT",
            "score": 0,
            "reason": "accident_human_interest",
            "matchedKeywords": accident_matches,
        }

    # Generic crime is excluded before economic scoring.
    if crime_matches:
        return {
            "relevant": False,
            "category": "REJECT",
            "score": 0,
            "reason": "crime_human_interest",
            "matchedKeywords": crime_matches,
        }

    if weather_matches:
        economic_title_matches = has_strong_economic_title(
            title_lower
        )

        market_title_matches = has_strong_market_title(
            title_lower
        )

        if not (
            economic_title_matches
            or market_title_matches
        ):
            return {
                "relevant": False,
                "category": "REJECT",
                "score": 0,
                "reason": "weather",
                "matchedKeywords": weather_matches,
            }

    # Generic social/education policy is excluded.
    if policy_matches:
        return {
            "relevant": False,
            "category": "REJECT",
            "score": 0,
            "reason": "non_economic_social_policy",
            "matchedKeywords": policy_matches,
        }

    # ------------------------------------------------------------------
    # 2. STRONG ECONOMIC TITLE
    # ------------------------------------------------------------------

    economic_title_matches = has_strong_economic_title(
        title_lower
    )

    if economic_title_matches:
        return {
            "relevant": True,
            "category": "ECONOMIC",
            "score": 100 + len(economic_title_matches),
            "reason": "strong_economic_title",
            "matchedKeywords": economic_title_matches,
        }

    # ------------------------------------------------------------------
    # 3. STRONG MARKET TITLE
    # ------------------------------------------------------------------

    market_title_matches = has_strong_market_title(
        title_lower
    )

    if market_title_matches:
        return {
            "relevant": True,
            "category": "MARKET",
            "score": 90 + len(market_title_matches),
            "reason": "strong_market_title",
            "matchedKeywords": market_title_matches,
        }

    # ------------------------------------------------------------------
    # 4. SPECIFIC BUSINESS / MARKET CONTEXT
    # ------------------------------------------------------------------

    if market_business_context(title_lower):
        return {
            "relevant": True,
            "category": "MARKET",
            "score": 85,
            "reason": "business_market_event",
            "matchedKeywords": [],
        }

    # ------------------------------------------------------------------
    # 5. STRONG TECHNOLOGY / BUSINESS TITLE
    # ------------------------------------------------------------------

    technology_title_matches = has_strong_technology_title(
        title_lower
    )

    if technology_title_matches:
        if technology_business_context(title_lower):
            return {
                "relevant": True,
                "category": "TECHNOLOGY",
                "score": 80 + len(technology_title_matches),
                "reason": "technology_business_or_infrastructure",
                "matchedKeywords": technology_title_matches,
            }

    # ------------------------------------------------------------------
    # 6. GEOPOLITICAL + DIRECT ECONOMIC MECHANISM
    # ------------------------------------------------------------------

    geo_matches = has_strong_geopolitical_economic_title(
        title_lower
    )

    if geo_matches:
        return {
            "relevant": True,
            "category": "GEOPOLITICAL",
            "score": 70 + len(geo_matches),
            "reason": "geopolitical_with_direct_economic_impact",
            "matchedKeywords": geo_matches,
        }

    # ------------------------------------------------------------------
    # 7. DESCRIPTION SUPPORT
    # ------------------------------------------------------------------
    #
    # Description can rescue a borderline title only when it provides
    # concrete evidence AND the title itself does not indicate an
    # obviously excluded topic.
    #
    # Description alone cannot override hard exclusions.
    # ------------------------------------------------------------------

    description_economic_matches = count_matches(
        description_lower,
        ECONOMIC_STRONG,
    )

    description_market_matches = count_matches(
        description_lower,
        MARKET_STRONG,
    )

    description_technology_matches = count_matches(
        description_lower,
        TECHNOLOGY_STRONG,
    )

    description_geo_matches = count_matches(
        description_lower,
        GEOPOLITICAL_ECONOMIC_STRONG,
    )

    # A description with multiple strong economic indicators can rescue
    # a neutral business headline.
    if len(description_economic_matches) >= 2:
        return {
            "relevant": True,
            "category": "ECONOMIC",
            "score": 60 + len(description_economic_matches),
            "reason": "supporting_economic_description",
            "matchedKeywords": description_economic_matches,
        }

    # A concrete market description can rescue a neutral headline.
    if len(description_market_matches) >= 2:
        return {
            "relevant": True,
            "category": "MARKET",
            "score": 55 + len(description_market_matches),
            "reason": "supporting_market_description",
            "matchedKeywords": description_market_matches,
        }

    # Technology requires actual technology/business context.
    if len(description_technology_matches) >= 2:
        if (
            "technology"
            in description_lower
            or "artificial intelligence"
            in description_lower
            or "semiconductor"
            in description_lower
            or "chip"
            in description_lower
            or "ai"
            in description_lower
        ):
            return {
                "relevant": True,
                "category": "TECHNOLOGY",
                "score": 50 + len(description_technology_matches),
                "reason": "supporting_technology_description",
                "matchedKeywords": description_technology_matches,
            }

    # Geopolitical stories still require a concrete economic mechanism.
    if len(description_geo_matches) >= 2:
        return {
            "relevant": True,
            "category": "GEOPOLITICAL",
            "score": 50 + len(description_geo_matches),
            "reason": "supporting_geopolitical_economic_description",
            "matchedKeywords": description_geo_matches,
        }

    # ------------------------------------------------------------------
    # 8. GENERIC POLITICAL / COMMENTARY REJECTION
    # ------------------------------------------------------------------

    political_matches = count_matches(
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

    # ------------------------------------------------------------------
    # 9. DEFAULT REJECTION
    # ------------------------------------------------------------------

    return {
        "relevant": False,
        "category": "REJECT",
        "score": 0,
        "reason": "insufficient_market_relevance",
        "matchedKeywords": [],
    }


# ============================================================================
# RSS DATE PARSING
# ============================================================================

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


# ============================================================================
# RSS FETCHING
# ============================================================================

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
                    "url": canonicalize_url(url),
                    "publishedAtUtc": (
                        published
                        .astimezone(timezone.utc)
                        .replace(microsecond=0)
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
            "  ERROR:",
            type(exc).__name__,
            str(exc),
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


# ============================================================================
# ARTICLE EXTRACTION
# ============================================================================

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
        " ".join(paragraphs)
    )


# ============================================================================
# SUMMARY GENERATION
# ============================================================================

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
        "valuation",
        "payments",
        "fintech",
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
        " ".join(ordered)
    )


# ============================================================================
# JSON HELPERS
# ============================================================================

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
            f"WARNING: Failed reading {path}: "
            f"{type(exc).__name__}: {exc}"
        )

        return default


# ============================================================================
# HISTORY
# ============================================================================

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
        article["publishedAtUtc"]
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

    data["articleCount"] = len(
        articles
    )

    data["updatedAtUtc"] = (
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


# ============================================================================
# CURRENT 14-DAY INDEX
# ============================================================================

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


# ============================================================================
# MAIN
# ============================================================================

def main() -> None:
    print(
        "=" * 80
    )

    print(
        "VGrat FMS - CNA Market News Collector"
    )

    print(
        "=" * 80
    )

    print(
        f"Rolling window: {WINDOW_DAYS} days"
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
        "rejectedSocialPolicy": 0,
        "rejectedGeneral": 0,
        "articleExtractionAttempts": 0,
        "articleExtractionSuccess": 0,
        "articleExtractionFailures": 0,
        "newHistoricalRecords": 0,
    }

    # ------------------------------------------------------------------
    # RSS
    # ------------------------------------------------------------------

    rss_records = collect_rss()

    stats[
        "rssEntries"
    ] = len(
        rss_records
    )

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

        record["id"] = key

        record["feeds"] = [
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

    # ------------------------------------------------------------------
    # CLASSIFICATION
    # ------------------------------------------------------------------

    relevant_records = []

    print()

    print(
        "=" * 80
    )

    print(
        "CLASSIFICATION"
    )

    print(
        "=" * 80
    )

    for record in unique_records.values():
        result = classify_relevance(
            record["title"],
            record["description"],
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

            elif reason == "political_without_market_or_economic_impact":
                stats[
                    "rejectedPolitical"
                ] += 1

            elif reason == "non_economic_social_policy":
                stats[
                    "rejectedSocialPolicy"
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

    # ------------------------------------------------------------------
    # HISTORY
    # ------------------------------------------------------------------

    existing_ids = (
        load_existing_history_ids()
    )

    for record in relevant_records:
        key = record[
            "id"
        ]

        if key in existing_ids:
            continue

        title = record[
            "title"
        ]

        print()

        print(
            "PROCESS:",
            title,
        )

        stats[
            "articleExtractionAttempts"
        ] += 1

        try:
            article_text = fetch_article_text(
                record["url"]
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

    # ------------------------------------------------------------------
    # CURRENT INDEX
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # RUN SUMMARY
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # CONSOLE SUMMARY
    # ------------------------------------------------------------------

    print()

    print(
        "=" * 80
    )

    print(
        "COLLECTION COMPLETE"
    )

    print(
        "=" * 80
    )

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
        f"Rejected social policy:"
        f" {stats['rejectedSocialPolicy']}"
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

    print(
        "=" * 80
    )


if __name__ == "__main__":
    main()
