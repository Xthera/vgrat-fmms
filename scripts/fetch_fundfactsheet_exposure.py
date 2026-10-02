#!/usr/bin/env python3

"""
VGrat FMS - GOOGLE DISCOVERY FUND GEOGRAPHIC / SECTOR EXPOSURE

PURPOSE
=======

Search every PRULink fund in Funds Links.xlsm through Google
Programmable Search / Custom Search JSON API.

The search engine is used for DISCOVERY.

The actual geographic / sector values are accepted only when
they can be extracted explicitly from the discovered source.

SOURCE PRIORITY
===============

1. Official Prudential Singapore
2. Official underlying fund manager
3. Reputable secondary source

IMPORTANT
=========

Google search is NOT treated as the source of the allocation values.

The pipeline:

    Fund name
        |
        v
    Google Search
        |
        v
    Candidate sources
        |
        v
    Download source
        |
        v
    Verify exact fund / underlying fund
        |
        v
    Extract explicit allocation data
        |
        v
    JSON

NO INFERENCE
============

This script does NOT:

- infer country from holdings
- infer sector from holdings
- calculate country exposure from Top 10 holdings
- calculate sector exposure from Top 10 holdings
- normalize percentages to 100%
- manufacture missing values
- copy allocation data from a different fund merely because
  the names look similar

If explicit allocation data cannot be verified, the field remains
empty and the fund is marked partially_resolved or unresolved.

INPUT
=====

Preferred:

    Funds Links.xlsm

Column A:
    Prudential fund URL

Column B:
    Exact PruAccess / fund name reference

OUTPUT
======

data/fundfactsheet_exposure.json

data/fundfactsheet_exposure_diagnostics.json

ENVIRONMENT
===========

Required environment variables:

    GOOGLE_API_KEY
    GOOGLE_CSE_ID

These should be configured as GitHub Actions secrets.

"""

from __future__ import annotations

import io
import json
import logging
import os
import re
import sys
import time
import hashlib
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urlparse

import openpyxl
import pdfplumber
import requests
from bs4 import BeautifulSoup


# ============================================================
# CONFIGURATION
# ============================================================

ROOT = Path(__file__).resolve().parents[1]

MASTER_XLSM = ROOT / "Funds Links.xlsm"

OUTPUT_DIR = ROOT / "data"

OUTPUT_JSON = OUTPUT_DIR / "fundfactsheet_exposure.json"
DIAGNOSTICS_JSON = OUTPUT_DIR / "fundfactsheet_exposure_diagnostics.json"

GOOGLE_ENDPOINT = "https://www.googleapis.com/customsearch/v1"

REQUEST_TIMEOUT = 30

MAX_GOOGLE_RESULTS = 10

MAX_CANDIDATES_TO_FETCH = 8

MAX_HTML_BYTES = 8 * 1024 * 1024

MAX_PDF_BYTES = 20 * 1024 * 1024

GOOGLE_DELAY_SECONDS = 0.25

SOURCE_FETCH_DELAY_SECONDS = 0.15

USER_AGENT = (
    "Mozilla/5.0 "
    "(X11; Linux x86_64) "
    "AppleWebKit/537.36 "
    "(KHTML, like Gecko) "
    "Chrome/140.0 Safari/537.36 "
    "VGrat-FMS/1.0"
)

HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept-Language": "en-SG,en;q=0.9",
}

# Domains considered official / high priority.
OFFICIAL_DOMAIN_HINTS = {
    "prudential.com.sg": 100,
    "eastspring.com": 95,
    "fidelity.com": 95,
    "fidelityinternational.com": 95,
    "schroders.com": 95,
    "jpmorgan.com": 95,
    "jpmorganassetmanagement.com": 95,
    "blackrock.com": 95,
    "ishares.com": 95,
    "amundi.com": 95,
    "ubs.com": 95,
    "vanguard.com": 95,
    "franklintempleton.com": 95,
    "aberdeen.com": 95,
    "abrdn.com": 95,
    "manulifeim.com": 95,
    "pimco.com": 95,
    "columbiathreadneedle.com": 95,
    "invesco.com": 95,
    "wellington.com": 95,
    "bnpparibas-am.com": 95,
    "dws.com": 95,
    "allianzgi.com": 95,
    "axa-im.com": 95,
    "morganstanley.com": 95,
    "lgim.com": 95,
    "schroders.com": 95,
    "uobam.com.sg": 95,
    "uobam.com": 95,
    "dbsvickers.com": 50,
}

SECONDARY_DOMAIN_HINTS = {
    "morningstar.com": 70,
    "investing.com": 55,
    "markets.ft.com": 55,
    "lipperweb.com": 60,
    "fundsquare.net": 50,
    "bloomberg.com": 60,
    "marketscreener.com": 45,
}

GEOGRAPHIC_HEADINGS = [
    "country allocation",
    "country allocations",
    "geographical allocation",
    "geographic allocation",
    "geographical allocations",
    "geographic allocations",
    "regional allocation",
    "regional allocations",
    "country exposure",
    "geographic exposure",
    "geographical exposure",
    "regional exposure",
]

SECTOR_HEADINGS = [
    "sector allocation",
    "sector allocations",
    "industry allocation",
    "industry allocations",
    "sector exposure",
    "industry exposure",
    "industry breakdown",
    "sector breakdown",
]

STOP_HEADINGS = [
    "top 10 holdings",
    "top 10 holding",
    "top ten holdings",
    "top holdings",
    "performance",
    "important information",
    "investment objective",
    "fund details",
    "fund statistics",
    "risk classification",
    "benchmark",
    "calendar year performance",
    "past performance",
    "disclaimer",
]

GEOGRAPHIC_NAMES = {
    "united states",
    "usa",
    "us",
    "canada",
    "mexico",
    "brazil",
    "argentina",
    "chile",
    "peru",
    "colombia",
    "united kingdom",
    "uk",
    "great britain",
    "germany",
    "france",
    "italy",
    "spain",
    "netherlands",
    "switzerland",
    "sweden",
    "denmark",
    "norway",
    "finland",
    "ireland",
    "belgium",
    "austria",
    "portugal",
    "poland",
    "europe",
    "europe ex uk",
    "eurozone",
    "asia",
    "asia pacific",
    "asia ex japan",
    "japan",
    "china",
    "hong kong",
    "taiwan",
    "south korea",
    "korea",
    "india",
    "singapore",
    "australia",
    "new zealand",
    "malaysia",
    "indonesia",
    "thailand",
    "philippines",
    "vietnam",
    "middle east",
    "africa",
    "south africa",
    "emerging markets",
    "developed markets",
    "cash",
    "cash and others",
    "cash and cash equivalents",
    "others",
}

SECTOR_NAMES = {
    "information technology",
    "technology",
    "financials",
    "financial services",
    "health care",
    "healthcare",
    "industrials",
    "materials",
    "communication services",
    "telecommunication",
    "consumer discretionary",
    "consumer staples",
    "energy",
    "utilities",
    "real estate",
    "property",
    "semiconductors",
    "software",
    "internet",
    "electronics",
    "computers",
    "media",
    "banks",
    "insurance",
    "pharmaceuticals",
    "biotechnology",
    "diversified financials",
    "capital goods",
    "automobiles",
    "automotive",
    "transportation",
    "consumer services",
    "commercial services",
    "food and beverages",
    "retail",
    "cash",
    "cash and others",
    "cash and cash equivalents",
    "others",
}


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

logger = logging.getLogger("vgrat-factsheet-exposure")


# ============================================================
# GENERAL HELPERS
# ============================================================

def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def normalize_space(value: str) -> str:
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def normalize_name(value: str) -> str:
    value = value.lower()

    value = value.replace("–", "-")
    value = value.replace("—", "-")

    value = re.sub(r"\([^)]*\)", " ", value)

    replacements = {
        "prulink": "",
        "pru link": "",
        "pruprime": "",
        "accumulation": "",
        "acc": "",
        "distribution": "",
        "dist": "",
        "decu": "",
        "sgd": "",
        "usd": "",
    }

    for old, new in replacements.items():
        value = value.replace(old, new)

    value = re.sub(r"[^a-z0-9]+", " ", value)

    return normalize_space(value)


def name_tokens(value: str) -> List[str]:
    normalized = normalize_name(value)

    stop = {
        "fund",
        "portfolio",
        "the",
        "and",
        "of",
        "class",
    }

    return [
        token
        for token in normalized.split()
        if len(token) >= 3 and token not in stop
    ]


def filename_hash(url: str) -> str:
    return hashlib.sha256(url.encode("utf-8")).hexdigest()[:16]


def domain_of(url: str) -> str:
    try:
        host = urlparse(url).netloc.lower()
        if host.startswith("www."):
            host = host[4:]
        return host
    except Exception:
        return ""


def domain_score(url: str) -> Tuple[int, str]:
    domain = domain_of(url)

    for known, score in OFFICIAL_DOMAIN_HINTS.items():
        if domain == known or domain.endswith("." + known):
            return score, "official"

    for known, score in SECONDARY_DOMAIN_HINTS.items():
        if domain == known or domain.endswith("." + known):
            return score, "secondary"

    return 20, "other"


def safe_json_write(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    temporary = path.with_suffix(path.suffix + ".tmp")

    with temporary.open("w", encoding="utf-8") as handle:
        json.dump(
            data,
            handle,
            ensure_ascii=False,
            indent=2,
        )

    temporary.replace(path)


# ============================================================
# MASTER UNIVERSE
# ============================================================

def load_master_universe() -> List[Dict[str, str]]:
    if not MASTER_XLSM.exists():
        raise FileNotFoundError(
            f"Master universe not found: {MASTER_XLSM}"
        )

    workbook = openpyxl.load_workbook(
        MASTER_XLSM,
        read_only=True,
        data_only=True,
        keep_vba=True,
    )

    worksheet = workbook.active

    funds: List[Dict[str, str]] = []

    for row in worksheet.iter_rows(
        min_row=2,
        values_only=True,
    ):
        url = row[0] if len(row) >= 1 else None
        fund_name = row[1] if len(row) >= 2 else None

        if not fund_name:
            continue

        fund_name = normalize_space(str(fund_name))

        if url:
            url = normalize_space(str(url))
        else:
            url = ""

        funds.append(
            {
                "fundName": fund_name,
                "prudentialUrl": url,
            }
        )

    workbook.close()

    deduped: List[Dict[str, str]] = []
    seen = set()

    for fund in funds:
        key = normalize_name(fund["fundName"])

        if not key:
            continue

        if key in seen:
            continue

        seen.add(key)
        deduped.append(fund)

    return deduped


# ============================================================
# GOOGLE SEARCH
# ============================================================

def google_search(
    api_key: str,
    cse_id: str,
    query: str,
) -> Dict[str, Any]:

    params = {
        "key": api_key,
        "cx": cse_id,
        "q": query,
        "num": MAX_GOOGLE_RESULTS,
        "safe": "active",
        "hl": "en",
        "gl": "sg",
        "filter": "1",
    }

    response = requests.get(
        GOOGLE_ENDPOINT,
        params=params,
        headers=HEADERS,
        timeout=REQUEST_TIMEOUT,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Google API HTTP {response.status_code}: "
            f"{response.text[:500]}"
        )

    return response.json()


def build_queries(fund_name: str) -> List[str]:
    quoted = f'"{fund_name}"'

    return [
        (
            f"{quoted} "
            f'("country allocation" OR "geographic allocation" OR '
            f'"geographical allocation" OR "regional allocation") '
            f'("sector allocation" OR "industry allocation" OR '
            f'"sector exposure")'
        ),
        (
            f"{quoted} "
            f'("underlying fund" OR "investment manager" OR '
            f'"fund manager") '
            f'("country allocation" OR "sector allocation")'
        ),
    ]


def search_fund(
    fund_name: str,
    api_key: str,
    cse_id: str,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:

    queries = build_queries(fund_name)

    all_results: List[Dict[str, Any]] = []
    search_log: List[Dict[str, Any]] = []

    seen_urls = set()

    for query in queries:

        logger.info(
            "Google search: %s",
            query,
        )

        try:
            result = google_search(
                api_key=api_key,
                cse_id=cse_id,
                query=query,
            )

            items = result.get("items", [])

            search_log.append(
                {
                    "query": query,
                    "resultCount": len(items),
                    "totalResults": (
                        result.get("searchInformation", {})
                        .get("totalResults")
                    ),
                    "status": "success",
                }
            )

            for item in items:

                url = item.get("link", "")

                if not url:
                    continue

                if url in seen_urls:
                    continue

                seen_urls.add(url)

                all_results.append(
                    {
                        "query": query,
                        "title": item.get("title", ""),
                        "link": url,
                        "snippet": item.get("snippet", ""),
                        "displayLink": item.get(
                            "displayLink",
                            "",
                        ),
                        "mime": item.get("mime", ""),
                        "fileFormat": item.get(
                            "fileFormat",
                            "",
                        ),
                    }
                )

        except Exception as exc:

            logger.warning(
                "Google search failed for '%s': %s",
                query,
                exc,
            )

            search_log.append(
                {
                    "query": query,
                    "resultCount": 0,
                    "status": "error",
                    "error": str(exc),
                }
            )

        time.sleep(GOOGLE_DELAY_SECONDS)

    return rank_search_results(
        fund_name,
        all_results,
    ), search_log


# ============================================================
# GOOGLE RESULT RANKING
# ============================================================

def candidate_score(
    fund_name: str,
    result: Dict[str, Any],
) -> int:

    title = result.get("title", "")
    snippet = result.get("snippet", "")
    url = result.get("link", "")

    combined = (
        f"{title} {snippet} {url}"
    ).lower()

    score, _ = domain_score(url)

    normalized_fund = normalize_name(fund_name)

    if normalized_fund and normalized_fund in normalize_name(combined):
        score += 50

    tokens = name_tokens(fund_name)

    matched = 0

    for token in tokens:
        if token in combined:
            matched += 1

    score += min(30, matched * 5)

    allocation_terms = [
        "country allocation",
        "geographic allocation",
        "geographical allocation",
        "regional allocation",
        "sector allocation",
        "industry allocation",
        "sector exposure",
    ]

    for term in allocation_terms:
        if term in combined:
            score += 10

    if "underlying fund" in combined:
        score += 8

    if "investment manager" in combined:
        score += 8

    if url.lower().endswith(".pdf"):
        score += 8

    if "prudential.com.sg" in domain_of(url):
        score += 20

    return score


def rank_search_results(
    fund_name: str,
    results: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:

    ranked = []

    for result in results:

        enriched = dict(result)

        enriched["score"] = candidate_score(
            fund_name,
            result,
        )

        ranked.append(enriched)

    ranked.sort(
        key=lambda item: item["score"],
        reverse=True,
    )

    return ranked


# ============================================================
# SOURCE FETCHING
# ============================================================

def fetch_source(url: str) -> Tuple[str, bytes]:
    response = requests.get(
        url,
        headers=HEADERS,
        timeout=REQUEST_TIMEOUT,
        allow_redirects=True,
    )

    response.raise_for_status()

    content_type = (
        response.headers.get(
            "Content-Type",
            "",
        ).lower()
    )

    data = response.content

    if len(data) > MAX_PDF_BYTES:
        raise RuntimeError(
            f"Source too large: {len(data)} bytes"
        )

    if (
        "pdf" in content_type
        or url.lower().split("?")[0].endswith(".pdf")
    ):
        return "pdf", data

    return "html", data


# ============================================================
# PDF TEXT
# ============================================================

def extract_pdf_text(data: bytes) -> str:

    chunks: List[str] = []

    with pdfplumber.open(io.BytesIO(data)) as pdf:

        for page in pdf.pages:

            try:
                text = page.extract_text() or ""

                if text:
                    chunks.append(text)

            except Exception as exc:
                logger.debug(
                    "PDF page extraction failed: %s",
                    exc,
                )

    return "\n".join(chunks)


# ============================================================
# HTML TEXT
# ============================================================

def extract_html_text(data: bytes) -> str:

    if len(data) > MAX_HTML_BYTES:
        raise RuntimeError(
            f"HTML source too large: {len(data)} bytes"
        )

    soup = BeautifulSoup(
        data,
        "html.parser",
    )

    for tag in soup(
        [
            "script",
            "style",
            "noscript",
            "svg",
        ]
    ):
        tag.decompose()

    return soup.get_text(
        "\n",
        strip=True,
    )


# ============================================================
# FUND MATCHING
# ============================================================

def fund_match_score(
    fund_name: str,
    text: str,
) -> int:

    normalized_text = normalize_name(
        text[:50000]
    )

    tokens = name_tokens(fund_name)

    if not tokens:
        return 0

    matches = sum(
        1
        for token in tokens
        if token in normalized_text
    )

    return int(
        (matches / len(tokens)) * 100
    )


# ============================================================
# MANAGER EXTRACTION
# ============================================================

MANAGER_PATTERNS = [
    re.compile(
        r"(?:source|sourced)\s*:\s*"
        r"([A-Za-z0-9&.,'()\- ]{3,120})",
        re.IGNORECASE,
    ),
    re.compile(
        r"(?:investment manager|fund manager|manager of the fund)"
        r"\s*[:\-]\s*"
        r"([A-Za-z0-9&.,'()\- ]{3,150})",
        re.IGNORECASE,
    ),
]


def extract_manager(text: str) -> Optional[str]:

    for pattern in MANAGER_PATTERNS:

        matches = pattern.findall(text)

        for match in matches:

            value = normalize_space(match)

            value = re.split(
                r"\b(?:Important Information|The Fund|The underlying)\b",
                value,
                maxsplit=1,
                flags=re.IGNORECASE,
            )[0]

            value = value.strip(" -:;,.") 

            if len(value) >= 3:
                return value

    return None


# ============================================================
# SECTION DETECTION
# ============================================================

def heading_type(line: str) -> Optional[str]:

    normalized = normalize_space(
        line.lower()
    )

    for heading in GEOGRAPHIC_HEADINGS:

        if heading in normalized:
            return "geographic"

    for heading in SECTOR_HEADINGS:

        if heading in normalized:
            return "sector"

    return None


def is_stop_heading(line: str) -> bool:

    normalized = normalize_space(
        line.lower()
    )

    for heading in STOP_HEADINGS:

        if normalized.startswith(heading):
            return True

    return False


def clean_lines(text: str) -> List[str]:

    lines = []

    for raw in text.splitlines():

        line = normalize_space(raw)

        if not line:
            continue

        lines.append(line)

    return lines


def extract_sections(
    text: str,
) -> Dict[str, List[str]]:

    lines = clean_lines(text)

    sections = {
        "geographic": [],
        "sector": [],
    }

    current: Optional[str] = None

    for line in lines:

        detected = heading_type(line)

        if detected:
            current = detected
            continue

        if current and is_stop_heading(line):
            current = None
            continue

        if current:
            sections[current].append(line)

    return sections


# ============================================================
# ALLOCATION EXTRACTION
# ============================================================

PERCENT_RE = re.compile(
    r"(?<![\d.])"
    r"(-?\d+(?:[.,]\d+)?)"
    r"\s*%"
)


def normalize_percent(value: str) -> Optional[float]:

    try:
        value = value.replace(",", ".")
        number = float(value)

        if number < 0 or number > 100:
            return None

        return round(number, 4)

    except Exception:
        return None


def clean_category(value: str) -> str:

    value = normalize_space(value)

    value = re.sub(
        r"\s*[-–—:|]+\s*$",
        "",
        value,
    )

    return value.strip()


def category_is_valid(
    category: str,
    allocation_type: str,
) -> bool:

    value = normalize_name(category)

    if not value:
        return False

    if len(value) > 80:
        return False

    if re.search(
        r"\b(?:performance|benchmark|return|price|yield|"
        r"charge|fee|date|month|year|source)\b",
        value,
    ):
        return False

    if allocation_type == "geographic":

        if value in GEOGRAPHIC_NAMES:
            return True

        geographic_words = [
            "america",
            "europe",
            "asia",
            "africa",
            "pacific",
            "kingdom",
            "states",
            "korea",
            "arab",
        ]

        return any(
            word in value
            for word in geographic_words
        )

    if allocation_type == "sector":

        if value in SECTOR_NAMES:
            return True

        sector_words = [
            "technology",
            "financial",
            "health",
            "industrial",
            "material",
            "consumer",
            "energy",
            "utility",
            "real estate",
            "software",
            "semiconductor",
            "telecom",
            "communication",
            "bank",
            "insurance",
            "pharma",
            "automotive",
            "retail",
            "capital goods",
        ]

        return any(
            word in value
            for word in sector_words
        )

    return False


def parse_allocation_lines(
    lines: List[str],
    allocation_type: str,
) -> List[Dict[str, Any]]:

    results: List[Dict[str, Any]] = []

    for line in lines:

        matches = list(
            PERCENT_RE.finditer(line)
        )

        if not matches:
            continue

        # The normal case is one category + one percentage.
        if len(matches) == 1:

            match = matches[0]

            percentage = normalize_percent(
                match.group(1)
            )

            if percentage is None:
                continue

            category = clean_category(
                line[:match.start()]
            )

            if category_is_valid(
                category,
                allocation_type,
            ):
                results.append(
                    {
                        "name": category,
                        "weight": percentage,
                    }
                )

                continue

            # Some PDFs put the percentage before the label.
            category = clean_category(
                line[match.end():]
            )

            if category_is_valid(
                category,
                allocation_type,
            ):
                results.append(
                    {
                        "name": category,
                        "weight": percentage,
                    }
                )

            continue

        # Multiple percentages on a line.
        # Only accept when a clear table-like pattern exists.
        parts = re.split(
            r"\s{2,}|\t+|\|",
            line,
        )

        if len(parts) >= 2:

            for part in parts:

                m = PERCENT_RE.search(part)

                if not m:
                    continue

                percentage = normalize_percent(
                    m.group(1)
                )

                if percentage is None:
                    continue

                category = clean_category(
                    part[:m.start()]
                )

                if category_is_valid(
                    category,
                    allocation_type,
                ):
                    results.append(
                        {
                            "name": category,
                            "weight": percentage,
                        }
                    )

    return dedupe_allocations(results)


def dedupe_allocations(
    rows: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:

    result: List[Dict[str, Any]] = []

    seen = set()

    for row in rows:

        key = (
            normalize_name(
                str(row.get("name", ""))
            ),
            row.get("weight"),
        )

        if key in seen:
            continue

        seen.add(key)
        result.append(row)

    return result


# ============================================================
# TABLE EXTRACTION
# ============================================================

def extract_pdf_tables(
    data: bytes,
) -> Tuple[
    List[Dict[str, Any]],
    List[Dict[str, Any]],
]:

    geographic: List[Dict[str, Any]] = []
    sector: List[Dict[str, Any]] = []

    try:

        with pdfplumber.open(
            io.BytesIO(data)
        ) as pdf:

            for page_number, page in enumerate(
                pdf.pages,
                start=1,
            ):

                text = page.extract_text() or ""

                current = None

                for line in clean_lines(text):

                    detected = heading_type(line)

                    if detected:
                        current = detected
                        continue

                    if current and is_stop_heading(line):
                        current = None
                        continue

                    if not current:
                        continue

                    tables = page.extract_tables()

                    for table in tables or []:

                        for row in table or []:

                            if not row:
                                continue

                            cells = [
                                normalize_space(
                                    str(cell or "")
                                )
                                for cell in row
                            ]

                            row_text = " ".join(
                                cells
                            )

                            parsed = parse_allocation_lines(
                                [row_text],
                                current,
                            )

                            if current == "geographic":
                                geographic.extend(parsed)
                            else:
                                sector.extend(parsed)

    except Exception as exc:

        logger.debug(
            "PDF table extraction failed: %s",
            exc,
        )

    return (
        dedupe_allocations(geographic),
        dedupe_allocations(sector),
    )


# ============================================================
# ALLOCATION VALIDATION
# ============================================================

def allocation_quality(
    rows: List[Dict[str, Any]],
) -> int:

    if not rows:
        return 0

    score = 0

    score += min(
        50,
        len(rows) * 5,
    )

    total = sum(
        float(row["weight"])
        for row in rows
        if isinstance(
            row.get("weight"),
            (int, float),
        )
    )

    if 80 <= total <= 120:
        score += 30

    elif 50 <= total <= 150:
        score += 10

    if any(
        row.get("name", "").lower()
        in {
            "others",
            "cash",
            "cash and others",
            "cash and cash equivalents",
        }
        for row in rows
    ):
        score += 10

    return score


# ============================================================
# SOURCE ANALYSIS
# ============================================================

def analyze_source(
    fund_name: str,
    source: Dict[str, Any],
) -> Dict[str, Any]:

    url = source["link"]

    result = {
        "url": url,
        "title": source.get("title", ""),
        "domain": domain_of(url),
        "sourceScore": source.get("score", 0),
        "sourceClass": domain_score(url)[1],
        "contentType": None,
        "fundMatchScore": 0,
        "managerName": None,
        "factsheetDate": None,
        "geographicAllocation": [],
        "sectorAllocation": [],
        "extractionMethod": None,
        "status": "unusable",
        "reason": None,
    }

    try:

        content_type, data = fetch_source(url)

        result["contentType"] = content_type

        if content_type == "pdf":
            text = extract_pdf_text(data)

            geo_tables, sector_tables = (
                extract_pdf_tables(data)
            )

        else:
            text = extract_html_text(data)

            geo_tables = []
            sector_tables = []

        result["fundMatchScore"] = fund_match_score(
            fund_name,
            text,
        )

        manager = extract_manager(text)

        if manager:
            result["managerName"] = manager

        sections = extract_sections(text)

        geographic = parse_allocation_lines(
            sections["geographic"],
            "geographic",
        )

        sector = parse_allocation_lines(
            sections["sector"],
            "sector",
        )

        if geo_tables:
            geographic.extend(geo_tables)

        if sector_tables:
            sector.extend(sector_tables)

        geographic = dedupe_allocations(
            geographic
        )

        sector = dedupe_allocations(
            sector
        )

        result["geographicAllocation"] = geographic
        result["sectorAllocation"] = sector

        if geographic or sector:
            result["extractionMethod"] = (
                "explicit_text_or_table"
            )

        if (
            result["fundMatchScore"] < 35
            and "prudential.com.sg" not in domain_of(url)
        ):
            result["status"] = "rejected"
            result["reason"] = (
                "source_fund_name_match_too_low"
            )
            return result

        if geographic or sector:

            result["status"] = "candidate"

            if geographic and sector:
                result["reason"] = (
                    "geographic_and_sector_found"
                )
            elif geographic:
                result["reason"] = (
                    "geographic_found_only"
                )
            else:
                result["reason"] = (
                    "sector_found_only"
                )

        else:

            result["status"] = "unusable"
            result["reason"] = (
                "no_explicit_allocation_values_found"
            )

    except Exception as exc:

        result["status"] = "error"
        result["reason"] = str(exc)

    return result


# ============================================================
# SOURCE SELECTION
# ============================================================

def source_priority(
    source: Dict[str, Any],
) -> int:

    score = int(
        source.get(
            "sourceScore",
            0,
        )
    )

    geographic = source.get(
        "geographicAllocation",
        [],
    )

    sector = source.get(
        "sectorAllocation",
        [],
    )

    score += allocation_quality(
        geographic
    )

    score += allocation_quality(
        sector
    )

    if geographic and sector:
        score += 40

    if source.get(
        "sourceClass"
    ) == "official":
        score += 20

    return score


def select_best_source(
    candidates: List[Dict[str, Any]],
) -> Optional[Dict[str, Any]]:

    usable = [
        candidate
        for candidate in candidates
        if candidate.get("status") == "candidate"
    ]

    if not usable:
        return None

    usable.sort(
        key=source_priority,
        reverse=True,
    )

    return usable[0]


# ============================================================
# FUND PROCESSING
# ============================================================

def process_fund(
    index: int,
    total: int,
    fund: Dict[str, str],
    api_key: str,
    cse_id: str,
) -> Tuple[Dict[str, Any], Dict[str, Any]]:

    fund_name = fund["fundName"]

    logger.info(
        "[%d/%d] %s",
        index,
        total,
        fund_name,
    )

    ranked_results, search_log = search_fund(
        fund_name,
        api_key,
        cse_id,
    )

    candidates = []

    for source in ranked_results[
        :MAX_CANDIDATES_TO_FETCH
    ]:

        logger.info(
            "  Candidate: %s | score=%s | %s",
            source.get("title"),
            source.get("score"),
            source.get("link"),
        )

        try:

            analyzed = analyze_source(
                fund_name,
                source,
            )

            candidates.append(analyzed)

            logger.info(
                "    -> %s | geo=%d | sector=%d",
                analyzed.get("status"),
                len(
                    analyzed.get(
                        "geographicAllocation",
                        [],
                    )
                ),
                len(
                    analyzed.get(
                        "sectorAllocation",
                        [],
                    )
                ),
            )

        except Exception as exc:

            logger.warning(
                "Source analysis failed: %s",
                exc,
            )

        time.sleep(
            SOURCE_FETCH_DELAY_SECONDS
        )

    best = select_best_source(
        candidates
    )

    if best:

        geographic = best.get(
            "geographicAllocation",
            [],
        )

        sector = best.get(
            "sectorAllocation",
            [],
        )

        if geographic and sector:
            status = "resolved"
        else:
            status = "partially_resolved"

        record = {
            "fundName": fund_name,
            "prudentialUrl": fund.get(
                "prudentialUrl",
                "",
            ),
            "source": {
                "discovery": "google_custom_search",
                "type": best.get(
                    "sourceClass",
                    "other",
                ),
                "url": best.get(
                    "url",
                    "",
                ),
                "domain": best.get(
                    "domain",
                    "",
                ),
                "title": best.get(
                    "title",
                    "",
                ),
                "query": next(
                    (
                        item.get("query")
                        for item in ranked_results
                        if item.get("link")
                        == best.get("url")
                    ),
                    None,
                ),
            },
            "managerName": best.get(
                "managerName"
            ),
            "factsheetDate": best.get(
                "factsheetDate"
            ),
            "geographicAllocation": geographic,
            "sectorAllocation": sector,
            "status": status,
            "reason": best.get(
                "reason"
            ),
            "retrievedAt": utc_now(),
        }

    else:

        record = {
            "fundName": fund_name,
            "prudentialUrl": fund.get(
                "prudentialUrl",
                "",
            ),
            "source": {
                "discovery": "google_custom_search",
                "type": None,
                "url": None,
                "domain": None,
                "title": None,
                "query": None,
            },
            "managerName": None,
            "factsheetDate": None,
            "geographicAllocation": [],
            "sectorAllocation": [],
            "status": "unresolved",
            "reason": (
                "google_candidates_contained_no_verified_"
                "explicit_geographic_or_sector_values"
            ),
            "retrievedAt": utc_now(),
        }

    diagnostic = {
        "fundName": fund_name,
        "prudentialUrl": fund.get(
            "prudentialUrl",
            "",
        ),
        "searches": search_log,
        "googleResults": ranked_results,
        "analyzedCandidates": candidates,
        "selectedSource": (
            best.get("url")
            if best
            else None
        ),
        "completedAt": utc_now(),
    }

    return record, diagnostic


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    api_key = os.getenv(
        "GOOGLE_API_KEY"
    )

    cse_id = os.getenv(
        "GOOGLE_CSE_ID"
    )

    if not api_key:
        logger.error(
            "GOOGLE_API_KEY environment variable is missing."
        )
        return 2

    if not cse_id:
        logger.error(
            "GOOGLE_CSE_ID environment variable is missing."
        )
        return 2

    logger.info(
        "Using master universe: %s",
        MASTER_XLSM.name,
    )

    funds = load_master_universe()

    logger.info(
        "Loaded %d funds",
        len(funds),
    )

    if not funds:
        logger.error(
            "No funds found in master universe."
        )
        return 1

    output_funds: List[Dict[str, Any]] = []
    diagnostics: List[Dict[str, Any]] = []

    for index, fund in enumerate(
        funds,
        start=1,
    ):

        try:

            record, diagnostic = process_fund(
                index=index,
                total=len(funds),
                fund=fund,
                api_key=api_key,
                cse_id=cse_id,
            )

            output_funds.append(record)
            diagnostics.append(diagnostic)

        except Exception as exc:

            logger.exception(
                "Fatal fund-level error for %s",
                fund.get("fundName"),
            )

            output_funds.append(
                {
                    "fundName": fund.get(
                        "fundName"
                    ),
                    "prudentialUrl": fund.get(
                        "prudentialUrl",
                        "",
                    ),
                    "source": {
                        "discovery": "google_custom_search",
                        "type": None,
                        "url": None,
                        "domain": None,
                        "title": None,
                        "query": None,
                    },
                    "managerName": None,
                    "factsheetDate": None,
                    "geographicAllocation": [],
                    "sectorAllocation": [],
                    "status": "unresolved",
                    "reason": (
                        "fund_processing_error: "
                        + str(exc)
                    ),
                    "retrievedAt": utc_now(),
                }
            )

            diagnostics.append(
                {
                    "fundName": fund.get(
                        "fundName"
                    ),
                    "error": str(exc),
                    "completedAt": utc_now(),
                }
            )

    resolved = sum(
        1
        for fund in output_funds
        if fund.get("status")
        == "resolved"
    )

    partial = sum(
        1
        for fund in output_funds
        if fund.get("status")
        == "partially_resolved"
    )

    unresolved = sum(
        1
        for fund in output_funds
        if fund.get("status")
        == "unresolved"
    )

    geographic_count = sum(
        1
        for fund in output_funds
        if fund.get(
            "geographicAllocation"
        )
    )

    sector_count = sum(
        1
        for fund in output_funds
        if fund.get(
            "sectorAllocation"
        )
    )

    output = {
        "schemaVersion": "2.0",
        "generatedAt": utc_now(),
        "sourcePolicy": {
            "discovery": "Google Custom Search JSON API",
            "primary": "Official Prudential Singapore",
            "secondary": "Official underlying fund manager",
            "tertiary": "Reputable secondary source",
            "allowInference": False,
            "allowHoldingBasedEstimation": False,
        },
        "summary": {
            "fundCount": len(output_funds),
            "resolvedCount": resolved,
            "partiallyResolvedCount": partial,
            "unresolvedCount": unresolved,
            "geographicCount": geographic_count,
            "sectorCount": sector_count,
        },
        "funds": output_funds,
    }

    diagnostics_output = {
        "schemaVersion": "2.0",
        "generatedAt": utc_now(),
        "searchProvider": {
            "provider": "Google Custom Search JSON API",
            "queriesPerFund": 2,
            "maximumFunds": len(funds),
            "maximumSearchQueries": len(funds) * 2,
        },
        "funds": diagnostics,
    }

    safe_json_write(
        OUTPUT_JSON,
        output,
    )

    safe_json_write(
        DIAGNOSTICS_JSON,
        diagnostics_output,
    )

    logger.info(
        "========================================"
    )
    logger.info(
        "Completed Google discovery extraction"
    )
    logger.info(
        "Total funds: %d",
        len(output_funds),
    )
    logger.info(
        "Resolved: %d",
        resolved,
    )
    logger.info(
        "Partially resolved: %d",
        partial,
    )
    logger.info(
        "Unresolved: %d",
        unresolved,
    )
    logger.info(
        "Geographic data: %d",
        geographic_count,
    )
    logger.info(
        "Sector data: %d",
        sector_count,
    )
    logger.info(
        "Output: %s",
        OUTPUT_JSON,
    )
    logger.info(
        "Diagnostics: %s",
        DIAGNOSTICS_JSON,
    )
    logger.info(
        "========================================"
    )

    return 0


if __name__ == "__main__":
    sys.exit(main())
