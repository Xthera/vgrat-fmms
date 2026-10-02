#!/usr/bin/env python3

"""
VGrat FMS - PRUDENTIAL FUND FACTSHEET EXPOSURE EXTRACTOR

PURPOSE
=======

Build a completely separate dataset containing explicitly published:

    - Geographic / country allocation
    - Sector / industry allocation

for every fund in Funds Links.xlsm.

SOURCE PRIORITY
===============

1. Official Prudential Singapore fund page
2. Official Prudential Singapore factsheet PDF
3. Other documents linked directly from the Prudential fund page

IMPORTANT
=========

This script does NOT infer geography or sector from holdings.

It only accepts allocation information explicitly published by the
source document.

It also does NOT modify:

    - funds.json
    - BID history
    - holdings pipeline
    - Recovery 1
    - Recovery 2
    - Recovery 3

OUTPUT
======

data/fundfactsheet_exposure.json
data/fundfactsheet_exposure_diagnostics.json

The output is intentionally separate so it can be merged later.
"""

from __future__ import annotations

import io
import json
import logging
import math
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from pypdf import PdfReader


# ============================================================
# CONFIGURATION
# ============================================================

ROOT = Path(__file__).resolve().parents[1]

MASTER_XLSM = ROOT / "Funds Links.xlsm"

FALLBACK_FUNDS_JSON = ROOT / "data" / "funds.json"

OUTPUT_DIR = ROOT / "data"

OUTPUT_JSON = OUTPUT_DIR / "fundfactsheet_exposure.json"
DIAGNOSTICS_JSON = OUTPUT_DIR / "fundfactsheet_exposure_diagnostics.json"

REQUEST_TIMEOUT = 45

MAX_PDF_BYTES = 50 * 1024 * 1024

SLEEP_BETWEEN_FUNDS = 0.5

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/154.0.0.0 Safari/537.36"
)

HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "en-SG,en;q=0.9",
    "Cache-Control": "no-cache",
}

PRUDENTIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}

# Explicit headings we accept.
GEOGRAPHIC_HEADINGS = (
    "country allocation",
    "country allocation of underlying fund",
    "geographical allocation",
    "geographic allocation",
    "geographical breakdown",
    "geographic breakdown",
    "regional allocation",
    "regional allocation of underlying fund",
    "regional breakdown",
    "country exposure",
    "geographic exposure",
    "geographical exposure",
)

SECTOR_HEADINGS = (
    "sector allocation",
    "sector allocation of underlying fund",
    "sector breakdown",
    "sector exposure",
    "industry allocation",
    "industry allocation of underlying fund",
    "industry breakdown",
    "industry exposure",
    "economic sector",
    "sector weightings",
)

# Words which strongly suggest the document is an actual factsheet.
FACTSHEET_TERMS = (
    "fund factsheet",
    "factsheet",
    "all data as at",
    "investment objective",
    "fund details",
)

# Known allocation terminology.
GEOGRAPHY_TERMS = (
    "country",
    "countries",
    "regional",
    "region",
    "geographical",
    "geographic",
)

SECTOR_TERMS = (
    "sector",
    "industry",
    "industrials",
    "financials",
    "technology",
    "health care",
    "consumer",
    "energy",
    "utilities",
    "materials",
    "communication services",
    "real estate",
)

# Words that should never be treated as allocation categories.
BAD_CATEGORY_TERMS = {
    "country allocation",
    "sector allocation",
    "country allocation of underlying fund",
    "sector allocation of underlying fund",
    "fund",
    "underlying fund",
    "allocation",
    "source",
    "important information",
    "cash",
    "cash and cash equivalents",
    "total",
    "others",
}

# Percentages such as:
# 28.6%
# 28.6 %
# -2.5%
PERCENT_RE = re.compile(
    r"(?<![\d.])(-?\d+(?:\.\d+)?)\s*%"
)

DATE_PATTERNS = [
    re.compile(
        r"\b(\d{1,2})\s+"
        r"(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+"
        r"(\d{4})\b",
        re.I,
    ),
    re.compile(
        r"\b(\d{1,2})[-/]"
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)"
        r"[-/](\d{4})\b",
        re.I,
    ),
    re.compile(
        r"\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b"
    ),
]

MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

LOGGER = logging.getLogger("fundfactsheet-exposure")


# ============================================================
# HTTP SESSION
# ============================================================

SESSION = requests.Session()
SESSION.headers.update(HEADERS)


# ============================================================
# GENERAL HELPERS
# ============================================================

def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def clean_text(value: Any) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace("\xa0", " ")
    text = text.replace("\u2013", "-")
    text = text.replace("\u2014", "-")
    text = text.replace("\u2212", "-")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def normalize_name(value: str) -> str:
    value = clean_text(value).lower()

    value = value.replace("&", " and ")

    value = re.sub(r"\((?:sgd|usd|hkd|aud|gbp)\)", "", value)

    value = re.sub(
        r"\b(accumulation|accum|distribution|dis|decumulation|decum)\b",
        "",
        value,
    )

    value = re.sub(r"[^a-z0-9]+", " ", value)

    return re.sub(r"\s+", " ", value).strip()


def name_tokens(value: str) -> List[str]:
    return [
        token
        for token in normalize_name(value).split()
        if len(token) >= 3
    ]


def name_similarity(a: str, b: str) -> float:
    """
    Conservative identity comparison.

    Exact normalized match = 1.0.

    Otherwise calculates token overlap, but requires the core
    fund name to be reasonably represented.
    """

    na = normalize_name(a)
    nb = normalize_name(b)

    if not na or not nb:
        return 0.0

    if na == nb:
        return 1.0

    ta = set(name_tokens(a))
    tb = set(name_tokens(b))

    if not ta or not tb:
        return 0.0

    intersection = ta & tb

    score = len(intersection) / max(len(ta), len(tb))

    return round(score, 4)


def is_prudential_url(url: str) -> bool:
    try:
        host = urlparse(url).netloc.lower()
        return host in PRUDENTIAL_HOSTS or host.endswith(".prudential.com.sg")
    except Exception:
        return False


def absolute_url(base: str, href: str) -> str:
    return urljoin(base, href)


def is_pdf_url(url: str) -> bool:
    lower = url.lower()

    if ".pdf" in lower:
        return True

    return "format=pdf" in lower


def looks_like_factsheet_url(url: str) -> bool:
    lower = url.lower()

    keywords = (
        "factsheet",
        "fund-factsheet",
        "fund_factsheet",
        "fact-sheet",
        "prulink-funds",
        "/factsheets/",
    )

    return any(keyword in lower for keyword in keywords)


# ============================================================
# MASTER UNIVERSE
# ============================================================

def load_master_universe() -> List[Dict[str, str]]:
    """
    Preferred source:

        Funds Links.xlsm

    Column A:
        Prudential URL

    Column B:
        Exact PruAccess fund name
    """

    if MASTER_XLSM.exists():
        LOGGER.info("Using master universe: %s", MASTER_XLSM.name)

        wb = load_workbook(
            MASTER_XLSM,
            read_only=True,
            data_only=True,
            keep_links=True,
        )

        ws = wb.active

        funds: List[Dict[str, str]] = []

        for row in range(2, ws.max_row + 1):
            prudential_url = clean_text(ws.cell(row=row, column=1).value)
            fund_name = clean_text(ws.cell(row=row, column=2).value)

            if not prudential_url:
                continue

            if not fund_name:
                # Use URL as fallback identity, but keep it explicit.
                fund_name = Path(
                    urlparse(prudential_url).path
                ).stem.replace("-", " ").strip()

            funds.append(
                {
                    "fundName": fund_name,
                    "prudentialUrl": prudential_url,
                    "sourceUniverse": "Funds Links.xlsm",
                }
            )

        wb.close()

        LOGGER.info("Loaded %d funds", len(funds))

        return funds

    # --------------------------------------------------------
    # FALLBACK
    # --------------------------------------------------------

    if FALLBACK_FUNDS_JSON.exists():
        LOGGER.warning(
            "Funds Links.xlsm not found; using %s",
            FALLBACK_FUNDS_JSON,
        )

        payload = json.loads(
            FALLBACK_FUNDS_JSON.read_text(
                encoding="utf-8"
            )
        )

        if isinstance(payload, dict):
            items = payload.get("funds", [])
        else:
            items = payload

        funds = []

        for item in items:
            if not isinstance(item, dict):
                continue

            name = clean_text(
                item.get("fundName")
                or item.get("name")
            )

            url = clean_text(
                item.get("prudentialUrl")
                or item.get("url")
            )

            if name:
                funds.append(
                    {
                        "fundName": name,
                        "prudentialUrl": url,
                        "sourceUniverse": "data/funds.json",
                    }
                )

        LOGGER.info("Loaded %d funds", len(funds))

        return funds

    raise FileNotFoundError(
        "Neither Funds Links.xlsm nor data/funds.json exists."
    )


# ============================================================
# WEB FETCH
# ============================================================

def fetch_url(
    url: str,
    *,
    timeout: int = REQUEST_TIMEOUT,
) -> Optional[requests.Response]:

    try:
        response = SESSION.get(
            url,
            timeout=timeout,
            allow_redirects=True,
        )

        response.raise_for_status()

        return response

    except Exception as exc:
        LOGGER.debug(
            "GET failed: %s | %s",
            url,
            exc,
        )

        return None


# ============================================================
# PRUDENTIAL PAGE DISCOVERY
# ============================================================

def parse_html_links(
    page_url: str,
    html: str,
) -> List[Dict[str, str]]:

    soup = BeautifulSoup(html, "html.parser")

    results: List[Dict[str, str]] = []

    seen = set()

    for anchor in soup.find_all("a"):
        href = clean_text(anchor.get("href"))

        if not href:
            continue

        if href.startswith("#"):
            continue

        url = absolute_url(page_url, href)

        if url in seen:
            continue

        seen.add(url)

        text = clean_text(anchor.get_text(" ", strip=True))

        results.append(
            {
                "url": url,
                "text": text,
            }
        )

    return results


def discover_prudential_candidates(
    fund_name: str,
    prudential_url: str,
) -> Dict[str, Any]:

    result = {
        "pageUrl": prudential_url,
        "pageFetched": False,
        "pageStatus": None,
        "pageFinalUrl": None,
        "candidateLinks": [],
        "pdfCandidates": [],
        "factsheetCandidates": [],
        "errors": [],
    }

    if not prudential_url:
        result["errors"].append("missing_prudential_url")
        return result

    response = fetch_url(prudential_url)

    if response is None:
        result["errors"].append("prudential_page_fetch_failed")
        return result

    result["pageFetched"] = True
    result["pageStatus"] = response.status_code
    result["pageFinalUrl"] = response.url

    links = parse_html_links(
        response.url,
        response.text,
    )

    result["candidateLinks"] = links

    pdfs = []

    for link in links:
        url = link["url"]
        text = link["text"]

        if not is_pdf_url(url):
            continue

        pdfs.append(
            {
                "url": url,
                "linkText": text,
                "prudential": is_prudential_url(url),
                "factsheetUrl": looks_like_factsheet_url(url),
            }
        )

    # Factsheet-looking PDFs first.
    pdfs.sort(
        key=lambda item: (
            not item["factsheetUrl"],
            not item["prudential"],
        )
    )

    result["pdfCandidates"] = pdfs

    result["factsheetCandidates"] = [
        item
        for item in pdfs
        if item["factsheetUrl"]
    ]

    return result


# ============================================================
# OPTIONAL PLAYWRIGHT DISCOVERY
# ============================================================

def discover_with_playwright(
    prudential_url: str,
) -> Dict[str, Any]:

    """
    Prudential's current fund pages can expose document links through
    dynamically rendered page components.

    Playwright is therefore used only as a fallback when ordinary
    requests/HTML discovery does not find a factsheet.

    The function intentionally imports Playwright lazily.
    """

    result = {
        "available": False,
        "success": False,
        "links": [],
        "errors": [],
    }

    try:
        from playwright.sync_api import (
            sync_playwright,
        )
    except Exception as exc:
        result["errors"].append(
            f"playwright_import_failed: {exc}"
        )
        return result

    result["available"] = True

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True
            )

            page = browser.new_page(
                user_agent=USER_AGENT,
                locale="en-SG",
            )

            page.goto(
                prudential_url,
                wait_until="domcontentloaded",
                timeout=60000,
            )

            page.wait_for_timeout(2500)

            # Scroll to force lazy-loaded document components.
            page.evaluate(
                """
                () => {
                    window.scrollTo(
                        0,
                        document.body.scrollHeight
                    );
                }
                """
            )

            page.wait_for_timeout(1500)

            anchors = page.locator("a").all()

            seen = set()

            for anchor in anchors:
                try:
                    href = anchor.get_attribute("href")
                    text = anchor.inner_text()
                except Exception:
                    continue

                href = clean_text(href)
                text = clean_text(text)

                if not href:
                    continue

                url = absolute_url(
                    page.url,
                    href,
                )

                if url in seen:
                    continue

                seen.add(url)

                if (
                    is_pdf_url(url)
                    or looks_like_factsheet_url(url)
                    or "factsheet" in text.lower()
                ):
                    result["links"].append(
                        {
                            "url": url,
                            "text": text,
                        }
                    )

            browser.close()

        result["success"] = True

    except Exception as exc:
        result["errors"].append(
            f"playwright_failed: {exc}"
        )

    return result


# ============================================================
# PDF DOWNLOAD
# ============================================================

def download_pdf(
    url: str,
) -> Optional[bytes]:

    try:
        response = SESSION.get(
            url,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
            stream=True,
        )

        response.raise_for_status()

        content_type = (
            response.headers.get(
                "Content-Type",
                ""
            )
            .lower()
        )

        data = bytearray()

        for chunk in response.iter_content(
            chunk_size=64 * 1024
        ):
            if not chunk:
                continue

            data.extend(chunk)

            if len(data) > MAX_PDF_BYTES:
                LOGGER.warning(
                    "PDF exceeds MAX_PDF_BYTES: %s",
                    url,
                )
                return None

        blob = bytes(data)

        # Accept PDFs even if server reports a bad content type.
        if (
            blob.startswith(b"%PDF")
            or "application/pdf" in content_type
        ):
            return blob

        return None

    except Exception as exc:
        LOGGER.debug(
            "PDF download failed: %s | %s",
            url,
            exc,
        )

        return None


# ============================================================
# PDF TEXT
# ============================================================

def extract_pdf_pages(
    pdf_bytes: bytes,
) -> List[str]:

    reader = PdfReader(
        io.BytesIO(pdf_bytes)
    )

    pages: List[str] = []

    for page in reader.pages:
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""

        pages.append(text)

    return pages


def pdf_text(
    pages: List[str],
) -> str:

    return "\n".join(
        page
        for page in pages
        if page
    )


# ============================================================
# PDF IDENTITY
# ============================================================

def find_document_fund_names(
    text: str,
) -> List[str]:

    names = []

    lines = [
        clean_text(line)
        for line in text.splitlines()
    ]

    for line in lines:
        lower = line.lower()

        if "prulink" in lower:
            names.append(line)

        elif "pruprime" in lower:
            names.append(line)

    # Preserve order / uniqueness.
    output = []

    seen = set()

    for name in names:
        key = normalize_name(name)

        if key and key not in seen:
            seen.add(key)
            output.append(name)

    return output


def verify_document_identity(
    requested_fund: str,
    text: str,
) -> Dict[str, Any]:

    names = find_document_fund_names(text)

    best_name = ""
    best_score = 0.0

    for name in names:
        score = name_similarity(
            requested_fund,
            name,
        )

        if score > best_score:
            best_score = score
            best_name = name

    # Also permit exact phrase anywhere in the PDF.
    normalized_text = normalize_name(text)
    normalized_requested = normalize_name(
        requested_fund
    )

    exact_phrase = (
        normalized_requested
        and normalized_requested in normalized_text
    )

    if exact_phrase:
        best_score = max(
            best_score,
            1.0,
        )

    # Conservative acceptance.
    verified = (
        exact_phrase
        or best_score >= 0.70
    )

    return {
        "verified": verified,
        "score": round(best_score, 4),
        "matchedFundName": best_name,
        "detectedFundNames": names[:20],
    }


# ============================================================
# FACTSHEET DATE
# ============================================================

def extract_factsheet_date(
    text: str,
) -> Optional[str]:

    # Highest priority: "All data as at ..."
    match = re.search(
        r"all\s+data\s+as\s+at\s+"
        r"([^\n|]{6,40})",
        text,
        re.I,
    )

    if match:
        candidate = clean_text(
            match.group(1)
        )

        candidate = re.split(
            r"\bunless\b",
            candidate,
            maxsplit=1,
            flags=re.I,
        )[0].strip()

        parsed = parse_date_string(candidate)

        if parsed:
            return parsed

    # Generic date search.
    for pattern in DATE_PATTERNS:
        match = pattern.search(text)

        if not match:
            continue

        candidate = match.group(0)

        parsed = parse_date_string(candidate)

        if parsed:
            return parsed

    return None


def parse_date_string(
    value: str,
) -> Optional[str]:

    value = clean_text(value)

    # YYYY-MM-DD
    match = re.fullmatch(
        r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})",
        value,
    )

    if match:
        year, month, day = map(
            int,
            match.groups(),
        )

        try:
            return datetime(
                year,
                month,
                day,
            ).date().isoformat()
        except ValueError:
            return None

    # 30 Sep 2025 / 30 September 2025
    match = re.fullmatch(
        r"(\d{1,2})\s+"
        r"([A-Za-z]+)\s+"
        r"(\d{4})",
        value,
    )

    if match:
        day = int(match.group(1))
        month = MONTHS.get(
            match.group(2).lower()
        )
        year = int(match.group(3))

        if month:
            try:
                return datetime(
                    year,
                    month,
                    day,
                ).date().isoformat()
            except ValueError:
                return None

    return None


# ============================================================
# ALLOCATION SECTION DETECTION
# ============================================================

def normalized_heading_text(
    text: str,
) -> str:

    text = text.lower()

    text = text.replace(
        "\u00ad",
        "",
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def find_heading_positions(
    text: str,
    headings: Iterable[str],
) -> List[int]:

    lower = normalized_heading_text(text)

    positions = []

    for heading in headings:
        heading_norm = normalized_heading_text(
            heading
        )

        start = 0

        while True:
            index = lower.find(
                heading_norm,
                start,
            )

            if index < 0:
                break

            positions.append(index)

            start = index + len(
                heading_norm
            )

    return sorted(set(positions))


def extract_section_text(
    text: str,
    headings: Iterable[str],
    other_headings: Iterable[str],
    max_chars: int = 12000,
) -> Optional[str]:

    starts = find_heading_positions(
        text,
        headings,
    )

    if not starts:
        return None

    start = starts[0]

    ends = find_heading_positions(
        text[start + 1:],
        other_headings,
    )

    if ends:
        end = (
            start
            + 1
            + min(ends)
        )
    else:
        end = min(
            len(text),
            start + max_chars,
        )

    section = text[
        start:end
    ]

    return section.strip()


# ============================================================
# PERCENTAGE EXTRACTION
# ============================================================

def extract_percentage_values(
    text: str,
) -> List[float]:

    values = []

    for match in PERCENT_RE.finditer(text):
        try:
            value = float(
                match.group(1)
            )

            if math.isfinite(value):
                values.append(value)

        except Exception:
            continue

    return values


def clean_category(
    value: str,
) -> str:

    value = clean_text(value)

    value = re.sub(
        r"^[\-\u2022*•]+",
        "",
        value,
    )

    value = re.sub(
        r"\s+",
        " ",
        value,
    )

    value = value.strip(
        " :;,-"
    )

    return value


def category_is_valid(
    category: str,
    allocation_type: str,
) -> bool:

    category = clean_category(
        category
    )

    if not category:
        return False

    lower = category.lower()

    if lower in BAD_CATEGORY_TERMS:
        return False

    if len(category) < 2:
        return False

    if len(category) > 90:
        return False

    # Reject obvious narrative sentences.
    if len(category.split()) > 10:
        return False

    # Don't accept lines containing source/disclaimer prose.
    bad_fragments = (
        "important information",
        "there is no assurance",
        "past performance",
        "potential investor",
        "product summary",
        "copyright",
        "sustainability rating",
        "the fund is",
        "the underlying fund",
    )

    if any(
        fragment in lower
        for fragment in bad_fragments
    ):
        return False

    if allocation_type == "geographic":
        return any(
            term in lower
            for term in (
                GEOGRAPHY_TERMS
                + (
                    "united states",
                    "china",
                    "japan",
                    "singapore",
                    "hong kong",
                    "taiwan",
                    "india",
                    "australia",
                    "united kingdom",
                    "canada",
                    "europe",
                    "asia",
                    "others",
                )
            )
        ) or (
            # Country names frequently have no generic keyword.
            len(category.split()) <= 5
            and not any(
                word in lower
                for word in (
                    "allocation",
                    "source",
                    "data as at",
                )
            )
        )

    if allocation_type == "sector":
        return any(
            term in lower
            for term in SECTOR_TERMS
        ) or lower in {
            "financials",
            "industrials",
            "technology",
            "information technology",
            "health care",
            "consumer discretionary",
            "consumer staples",
            "communication services",
            "real estate",
            "materials",
            "energy",
            "utilities",
            "cash and cash equivalents",
            "others",
        }

    return False


# ============================================================
# TEXT-LINE EXTRACTION
# ============================================================

def extract_line_based_allocations(
    section: str,
    allocation_type: str,
) -> List[Dict[str, Any]]:

    lines = [
        clean_text(line)
        for line in section.splitlines()
    ]

    lines = [
        line
        for line in lines
        if line
    ]

    results = []

    # --------------------------------------------------------
    # Pattern A:
    #
    # United States 82.4%
    #
    # --------------------------------------------------------

    for line in lines:
        match = re.match(
            r"^(.*?)\s+(-?\d+(?:\.\d+)?)\s*%\s*$",
            line,
        )

        if not match:
            continue

        category = clean_category(
            match.group(1)
        )

        try:
            weight = float(
                match.group(2)
            )
        except Exception:
            continue

        if not category_is_valid(
            category,
            allocation_type,
        ):
            continue

        if not (
            -100.0
            <= weight
            <= 100.0
        ):
            continue

        results.append(
            {
                "name": category,
                "weight": weight,
            }
        )

    # --------------------------------------------------------
    # Pattern B:
    #
    # Category
    # 28.6%
    #
    # --------------------------------------------------------

    for index, line in enumerate(lines):
        match = PERCENT_RE.fullmatch(
            line
        )

        if not match:
            continue

        try:
            weight = float(
                match.group(1)
            )
        except Exception:
            continue

        # Search a few preceding lines for a category.
        for back in range(
            1,
            min(5, index + 1),
        ):
            category = clean_category(
                lines[index - back]
            )

            if not category_is_valid(
                category,
                allocation_type,
            ):
                continue

            results.append(
                {
                    "name": category,
                    "weight": weight,
                }
            )

            break

    return deduplicate_allocations(
        results
    )


# ============================================================
# TABLE-STYLE EXTRACTION
# ============================================================

def extract_adjacent_allocations(
    section: str,
    allocation_type: str,
) -> List[Dict[str, Any]]:

    """
    Handles PDF text extraction where labels and percentages
    are separated by whitespace rather than preserved as rows.
    """

    text = section

    matches = list(
        PERCENT_RE.finditer(text)
    )

    results = []

    for match in matches:
        try:
            weight = float(
                match.group(1)
            )
        except Exception:
            continue

        before = text[
            max(
                0,
                match.start() - 120,
            ):
            match.start()
        ]

        # Take the final line fragment before the percentage.
        fragments = re.split(
            r"[\n|]+",
            before,
        )

        if not fragments:
            continue

        category = clean_category(
            fragments[-1]
        )

        if not category_is_valid(
            category,
            allocation_type,
        ):
            continue

        results.append(
            {
                "name": category,
                "weight": weight,
            }
        )

    return deduplicate_allocations(
        results
    )


# ============================================================
# DEDUPLICATION
# ============================================================

def deduplicate_allocations(
    values: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:

    output = []

    seen = {}

    for item in values:
        name = clean_category(
            item.get("name")
        )

        if not name:
            continue

        key = normalize_name(name)

        if key in seen:
            # Keep the first explicit value.
            continue

        weight = item.get("weight")

        if not isinstance(
            weight,
            (int, float),
        ):
            continue

        if not math.isfinite(
            float(weight)
        ):
            continue

        seen[key] = True

        output.append(
            {
                "name": name,
                "weight": round(
                    float(weight),
                    4,
                ),
            }
        )

    return output


# ============================================================
# CHART-TEXT EXTRACTION
# ============================================================

def extract_chart_allocations(
    text: str,
    allocation_type: str,
) -> List[Dict[str, Any]]:

    """
    Conservative chart-text extraction.

    Some Prudential factsheets expose chart labels and percentage
    values in separate text blocks.

    This function only uses percentage values that are explicitly
    present in the PDF text.

    It never calculates missing percentages from holdings.
    """

    section = extract_section_text(
        text,
        GEOGRAPHIC_HEADINGS
        if allocation_type == "geographic"
        else SECTOR_HEADINGS,
        (
            SECTOR_HEADINGS
            if allocation_type == "geographic"
            else GEOGRAPHIC_HEADINGS
        ),
    )

    if not section:
        return []

    results = []

    lines = [
        clean_text(line)
        for line in section.splitlines()
    ]

    lines = [
        line
        for line in lines
        if line
    ]

    # First try conventional line-based extraction.
    results.extend(
        extract_line_based_allocations(
            section,
            allocation_type,
        )
    )

    if results:
        return deduplicate_allocations(
            results
        )

    # Then try adjacent text.
    results.extend(
        extract_adjacent_allocations(
            section,
            allocation_type,
        )
    )

    return deduplicate_allocations(
        results
    )


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_allocations_from_pdf(
    fund_name: str,
    pdf_url: str,
    pdf_bytes: bytes,
) -> Dict[str, Any]:

    result = {
        "success": False,
        "identity": {},
        "factsheetDate": None,
        "geographicAllocation": [],
        "sectorAllocation": [],
        "sourceUrl": pdf_url,
        "sourceDomain": urlparse(
            pdf_url
        ).netloc,
        "pdfBytes": len(pdf_bytes),
        "pages": 0,
        "hasGeographicHeading": False,
        "hasSectorHeading": False,
        "errors": [],
    }

    try:
        pages = extract_pdf_pages(
            pdf_bytes
        )

    except Exception as exc:
        result["errors"].append(
            f"pdf_parse_failed: {exc}"
        )
        return result

    result["pages"] = len(pages)

    text = pdf_text(
        pages
    )

    if not text.strip():
        result["errors"].append(
            "pdf_has_no_extractable_text"
        )
        return result

    identity = verify_document_identity(
        fund_name,
        text,
    )

    result["identity"] = identity

    if not identity["verified"]:
        result["errors"].append(
            "fund_identity_not_verified"
        )
        return result

    result["factsheetDate"] = (
        extract_factsheet_date(
            text
        )
    )

    normalized = normalized_heading_text(
        text
    )

    result["hasGeographicHeading"] = any(
        heading in normalized
        for heading in GEOGRAPHIC_HEADINGS
    )

    result["hasSectorHeading"] = any(
        heading in normalized
        for heading in SECTOR_HEADINGS
    )

    geographic = extract_chart_allocations(
        text,
        "geographic",
    )

    sector = extract_chart_allocations(
        text,
        "sector",
    )

    result["geographicAllocation"] = geographic
    result["sectorAllocation"] = sector

    if geographic or sector:
        result["success"] = True
    else:
        result["errors"].append(
            "no_explicit_allocation_values_extracted"
        )

    return result


# ============================================================
# MANAGER INFORMATION
# ============================================================

def extract_manager_names(
    text: str,
) -> List[str]:

    patterns = [
        r"Investment Manager of the Underlying Fund"
        r"\s+([A-Z][^\n]{3,100})",

        r"Investment Manager"
        r"\s+([A-Z][^\n]{3,100})",

        r"Manager of the Fund"
        r"\s+([A-Z][^\n]{3,100})",

        r"Source:\s*([^\n]{3,100})",
    ]

    results = []

    for pattern in patterns:
        for match in re.finditer(
            pattern,
            text,
            re.I,
        ):
            candidate = clean_text(
                match.group(1)
            )

            candidate = re.split(
                r"\b(?:Important Information|"
                r"Fund Details|"
                r"Investment Objective)\b",
                candidate,
                maxsplit=1,
                flags=re.I,
            )[0].strip()

            if (
                candidate
                and len(candidate) <= 120
            ):
                results.append(
                    candidate
                )

    output = []

    seen = set()

    for value in results:
        key = normalize_name(value)

        if key and key not in seen:
            seen.add(key)
            output.append(value)

    return output[:10]


# ============================================================
# CANDIDATE RANKING
# ============================================================

def candidate_score(
    fund_name: str,
    candidate: Dict[str, Any],
) -> int:

    score = 0

    url = candidate.get(
        "url",
        "",
    )

    text = candidate.get(
        "linkText",
        "",
    )

    if is_prudential_url(url):
        score += 30

    if looks_like_factsheet_url(url):
        score += 40

    lower_text = text.lower()

    if "factsheet" in lower_text:
        score += 30

    if "fund factsheet" in lower_text:
        score += 20

    return score


def rank_pdf_candidates(
    fund_name: str,
    candidates: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:

    for candidate in candidates:
        candidate["discoveryScore"] = candidate_score(
            fund_name,
            candidate,
        )

    return sorted(
        candidates,
        key=lambda item: (
            -item.get(
                "discoveryScore",
                0,
            ),
            item.get(
                "url",
                "",
            ),
        ),
    )


# ============================================================
# SINGLE FUND
# ============================================================

def process_fund(
    fund: Dict[str, str],
    index: int,
    total: int,
) -> Tuple[Dict[str, Any], Dict[str, Any]]:

    fund_name = fund["fundName"]
    prudential_url = fund.get(
        "prudentialUrl",
        "",
    )

    LOGGER.info(
        "[%d/%d] %s",
        index,
        total,
        fund_name,
    )

    retrieved_at = now_iso()

    output = {
        "fundName": fund_name,
        "prudentialUrl": prudential_url,
        "sourceUniverse": fund.get(
            "sourceUniverse",
            "Funds Links.xlsm",
        ),
        "source": None,
        "factsheetDate": None,
        "managerName": None,
        "geographicAllocation": [],
        "sectorAllocation": [],
        "status": "unresolved",
        "reason": "not_processed",
        "retrievedAt": retrieved_at,
    }

    diagnostics = {
        "fundName": fund_name,
        "prudentialUrl": prudential_url,
        "retrievedAt": retrieved_at,
        "discovery": {},
        "candidateAttempts": [],
        "managerNamesDetected": [],
        "errors": [],
    }

    discovery = discover_prudential_candidates(
        fund_name,
        prudential_url,
    )

    diagnostics["discovery"] = discovery

    candidates = list(
        discovery.get(
            "pdfCandidates",
            [],
        )
    )

    # --------------------------------------------------------
    # Playwright fallback
    # --------------------------------------------------------

    if not candidates:
        LOGGER.info(
            "  No PDF discovered through requests; "
            "trying rendered Prudential page"
        )

        rendered = discover_with_playwright(
            prudential_url
        )

        diagnostics[
            "playwrightDiscovery"
        ] = rendered

        for item in rendered.get(
            "links",
            [],
        ):
            url = item["url"]

            if (
                is_pdf_url(url)
                or looks_like_factsheet_url(url)
            ):
                candidates.append(
                    {
                        "url": url,
                        "linkText": item.get(
                            "text",
                            "",
                        ),
                        "prudential": is_prudential_url(
                            url
                        ),
                        "factsheetUrl": looks_like_factsheet_url(
                            url
                        ),
                    }
                )

    # --------------------------------------------------------
    # Deduplicate candidates
    # --------------------------------------------------------

    deduped = {}

    for candidate in candidates:
        url = candidate.get(
            "url",
            "",
        )

        if not url:
            continue

        deduped[url] = candidate

    candidates = rank_pdf_candidates(
        fund_name,
        list(deduped.values()),
    )

    # --------------------------------------------------------
    # Candidate processing
    # --------------------------------------------------------

    for candidate in candidates:

        url = candidate["url"]

        LOGGER.info(
            "  Trying factsheet: %s",
            url,
        )

        attempt = {
            "url": url,
            "discoveryScore": candidate.get(
                "discoveryScore",
                0,
            ),
            "downloaded": False,
            "identityVerified": False,
            "factsheetDate": None,
            "geographicCount": 0,
            "sectorCount": 0,
            "success": False,
            "errors": [],
        }

        pdf_bytes = download_pdf(
            url
        )

        if not pdf_bytes:
            attempt["errors"].append(
                "pdf_download_failed"
            )

            diagnostics[
                "candidateAttempts"
            ].append(attempt)

            continue

        attempt["downloaded"] = True

        extracted = extract_allocations_from_pdf(
            fund_name,
            url,
            pdf_bytes,
        )

        identity = extracted.get(
            "identity",
            {},
        )

        attempt[
            "identityVerified"
        ] = bool(
            identity.get(
                "verified",
                False,
            )
        )

        attempt["factsheetDate"] = (
            extracted.get(
                "factsheetDate"
            )
        )

        attempt[
            "geographicCount"
        ] = len(
            extracted.get(
                "geographicAllocation",
                [],
            )
        )

        attempt[
            "sectorCount"
        ] = len(
            extracted.get(
                "sectorAllocation",
                [],
            )
        )

        attempt["success"] = bool(
            extracted.get(
                "success",
                False,
            )
        )

        attempt["errors"].extend(
            extracted.get(
                "errors",
                [],
            )
        )

        diagnostics[
            "candidateAttempts"
        ].append(attempt)

        # Capture manager names from the
        # source PDF text where available.
        try:
            pages = extract_pdf_pages(
                pdf_bytes
            )

            manager_names = extract_manager_names(
                pdf_text(pages)
            )

            diagnostics[
                "managerNamesDetected"
            ].extend(manager_names)

        except Exception:
            pass

        if not extracted.get(
            "success",
            False,
        ):
            continue

        geo = extracted.get(
            "geographicAllocation",
            [],
        )

        sector = extracted.get(
            "sectorAllocation",
            [],
        )

        # ----------------------------------------------------
        # Require at least one allocation type.
        # ----------------------------------------------------

        if not geo and not sector:
            continue

        manager_names = (
            diagnostics.get(
                "managerNamesDetected",
                [],
            )
        )

        # Deduplicate managers.
        manager_output = None

        for manager in manager_names:
            if manager:
                manager_output = manager
                break

        output.update(
            {
                "source": {
                    "type": "prudential_factsheet",
                    "url": url,
                    "domain": urlparse(
                        url
                    ).netloc,
                    "title": candidate.get(
                        "linkText"
                    )
                    or "Prudential Fund Factsheet",
                },
                "factsheetDate": extracted.get(
                    "factsheetDate"
                ),
                "managerName": manager_output,
                "geographicAllocation": geo,
                "sectorAllocation": sector,
                "status": "resolved_prudential",
                "reason": (
                    "explicit allocation data "
                    "extracted from official "
                    "Prudential factsheet"
                ),
                "retrievedAt": retrieved_at,
            }
        )

        LOGGER.info(
            "  RESOLVED | geography=%d | sector=%d",
            len(geo),
            len(sector),
        )

        return output, diagnostics

    # --------------------------------------------------------
    # No source succeeded
    # --------------------------------------------------------

    if not candidates:
        output["reason"] = (
            "no_factsheet_pdf_discovered"
        )
    else:
        output["reason"] = (
            "factsheet_candidates_found_but_"
            "no_verified_explicit_allocations_extracted"
        )

    LOGGER.warning(
        "  UNRESOLVED | %s",
        output["reason"],
    )

    return output, diagnostics


# ============================================================
# CLEAN MANAGER DIAGNOSTICS
# ============================================================

def clean_diagnostics(
    diagnostics: Dict[str, Any],
) -> Dict[str, Any]:

    managers = []

    seen = set()

    for manager in diagnostics.get(
        "managerNamesDetected",
        [],
    ):
        key = normalize_name(
            manager
        )

        if key and key not in seen:
            seen.add(key)
            managers.append(
                manager
            )

    diagnostics[
        "managerNamesDetected"
    ] = managers

    return diagnostics


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    LOGGER.info(
        "VGrat FMS - Fund Factsheet Exposure Extraction"
    )

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        funds = load_master_universe()
    except Exception as exc:
        LOGGER.error(
            "Unable to load fund universe: %s",
            exc,
        )
        return 1

    if not funds:
        LOGGER.error(
            "Fund universe is empty."
        )
        return 1

    results = []

    diagnostics = []

    resolved = 0

    for index, fund in enumerate(
        funds,
        start=1,
    ):

        try:
            output, diagnostic = process_fund(
                fund,
                index,
                len(funds),
            )

        except Exception as exc:
            LOGGER.exception(
                "Unhandled error processing %s",
                fund.get("fundName"),
            )

            output = {
                "fundName": fund.get(
                    "fundName",
                    "",
                ),
                "prudentialUrl": fund.get(
                    "prudentialUrl",
                    "",
                ),
                "sourceUniverse": fund.get(
                    "sourceUniverse",
                    "Funds Links.xlsm",
                ),
                "source": None,
                "factsheetDate": None,
                "managerName": None,
                "geographicAllocation": [],
                "sectorAllocation": [],
                "status": "unresolved",
                "reason": (
                    f"unhandled_exception: {exc}"
                ),
                "retrievedAt": now_iso(),
            }

            diagnostic = {
                "fundName": fund.get(
                    "fundName",
                    "",
                ),
                "prudentialUrl": fund.get(
                    "prudentialUrl",
                    "",
                ),
                "retrievedAt": now_iso(),
                "errors": [
                    str(exc)
                ],
            }

        output["geographicAllocation"] = (
            output.get(
                "geographicAllocation",
                [],
            )
        )

        output["sectorAllocation"] = (
            output.get(
                "sectorAllocation",
                [],
            )
        )

        if output.get(
            "status"
        ) == "resolved_prudential":
            resolved += 1

        results.append(
            output
        )

        diagnostics.append(
            clean_diagnostics(
                diagnostic
            )
        )

        time.sleep(
            SLEEP_BETWEEN_FUNDS
        )

    # ========================================================
    # MAIN JSON
    # ========================================================

    payload = {
        "schemaVersion": "1.0",
        "generatedAt": now_iso(),
        "sourcePolicy": {
            "primary": (
                "Official Prudential Singapore "
                "fund pages and factsheet PDFs"
            ),
            "geographicRule": (
                "Only explicitly published "
                "geographic/country/regional "
                "allocation is accepted."
            ),
            "sectorRule": (
                "Only explicitly published "
                "sector/industry allocation "
                "is accepted."
            ),
            "holdingsInference": False,
            "forcedTotals": False,
        },
        "summary": {
            "fundCount": len(results),
            "resolved": resolved,
            "unresolved": (
                len(results) - resolved
            ),
        },
        "funds": results,
    }

    OUTPUT_JSON.write_text(
        json.dumps(
            payload,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    # ========================================================
    # DIAGNOSTICS JSON
    # ========================================================

    diagnostics_payload = {
        "schemaVersion": "1.0",
        "generatedAt": now_iso(),
        "summary": {
            "fundCount": len(results),
            "resolved": resolved,
            "unresolved": (
                len(results) - resolved
            ),
        },
        "funds": diagnostics,
    }

    DIAGNOSTICS_JSON.write_text(
        json.dumps(
            diagnostics_payload,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    LOGGER.info(
        "============================================================"
    )

    LOGGER.info(
        "Completed: %d resolved / %d unresolved",
        resolved,
        len(results) - resolved,
    )

    LOGGER.info(
        "Output: %s",
        OUTPUT_JSON,
    )

    LOGGER.info(
        "Diagnostics: %s",
        DIAGNOSTICS_JSON,
    )

    LOGGER.info(
        "============================================================"
    )

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )
