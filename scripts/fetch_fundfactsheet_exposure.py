#!/usr/bin/env python3

"""
VGrat FMS - INDEPENDENT FUND FACTSHEET EXPOSURE EXTRACTOR

PURPOSE
=======

Extract published:

    - Geographic allocation
    - Sector allocation

from fund factsheets.

SOURCE PRIORITY
===============

1. Official Prudential Singapore factsheet
2. Official underlying fund-manager factsheet
3. Unresolved

IMPORTANT
=========

This pipeline is independent from:

    - funds.json
    - BID history
    - holdings extraction
    - Holdings Recovery 1
    - Holdings Recovery 2
    - Holdings Recovery 3

It does NOT modify those pipelines.

OUTPUT
======

data/fundfactsheet_exposure.json

data/fundfactsheet_exposure_diagnostics.json

MASTER UNIVERSE
===============

Preferred:

    Funds Links.xlsm

Column A:
    Prudential fund URL

Column B:
    Exact PruAccess Fund Name

Fallback:

    data/funds.json

NO INFERENCE
===========

The extractor only records allocation percentages that are explicitly
published in a factsheet.

It does not:

    - calculate geographic exposure from holdings
    - calculate sector exposure from holdings
    - guess missing percentages
    - fill missing percentages
    - fabricate categories
    - force totals to 100%
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import logging
import re
import sys
import time

from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
import openpyxl
import pdfplumber


# ============================================================
# CONFIGURATION
# ============================================================

USER_AGENT = (
    "Mozilla/5.0 "
    "(compatible; VGrat-FMS-FundFactsheet/1.0)"
)

REQUEST_TIMEOUT = 30

MAX_PDF_BYTES = 25 * 1024 * 1024

DEFAULT_DELAY = 0.25


# ============================================================
# OFFICIAL PRUDENTIAL HOSTS
# ============================================================

PRUDENTIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
    "pruaccess.prudential.com.sg",
}


# ============================================================
# ALLOCATION SECTION HEADINGS
# ============================================================

GEOGRAPHIC_HEADINGS = (
    "geographical allocation",
    "geographic allocation",
    "geographical exposure",
    "geographic exposure",
    "country allocation",
    "country exposure",
    "regional allocation",
    "regional exposure",
    "allocation by country",
    "allocation by region",
    "asset allocation by country",
    "asset allocation by region",
)

SECTOR_HEADINGS = (
    "sector allocation",
    "sector exposure",
    "industry allocation",
    "industry exposure",
    "sector breakdown",
    "sector distribution",
    "allocation by sector",
    "allocation by industry",
)


# ============================================================
# REGEX
# ============================================================

PERCENT_RE = re.compile(
    r"(?<![\d.])"
    r"(\d{1,3}(?:[.,]\d+)?)"
    r"\s*%"
)

DATE_RE = re.compile(
    r"""
    \b
    (?:
        \d{1,2}
        [\s/-]+
        (?:
            Jan(?:uary)?
            |Feb(?:ruary)?
            |Mar(?:ch)?
            |Apr(?:il)?
            |May
            |Jun(?:e)?
            |Jul(?:y)?
            |Aug(?:ust)?
            |Sep(?:t(?:ember)?)?
            |Oct(?:ober)?
            |Nov(?:ember)?
            |Dec(?:ember)?
        )
        [\s/-]+
        \d{2,4}

        |

        20\d{2}
        [-/]
        \d{1,2}
        [-/]
        \d{1,2}
    )
    \b
    """,
    re.IGNORECASE | re.VERBOSE,
)


# ============================================================
# DATA CLASSES
# ============================================================

@dataclass
class Fund:
    row: int
    name: str
    prudential_url: str
    source_universe: str


@dataclass
class Candidate:
    url: str
    source_type: str
    discovered_from: str
    title: str = ""
    score: int = 0


# ============================================================
# HELPERS
# ============================================================

def now_utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def clean(value: Any) -> str:
    return re.sub(
        r"\s+",
        " ",
        str(value or "")
    ).strip()


def normalise(value: str) -> str:
    value = clean(value).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def is_prudential_url(url: str) -> bool:
    host = (urlparse(url).hostname or "").lower()

    return (
        host in PRUDENTIAL_HOSTS
        or any(
            host.endswith("." + domain)
            for domain in PRUDENTIAL_HOSTS
        )
    )


def is_pdf_url(url: str) -> bool:
    path = urlparse(url).path.lower()
    return path.endswith(".pdf")


def extract_percentage(text: str) -> float | None:

    match = PERCENT_RE.search(
        text.replace(",", "")
    )

    if not match:
        return None

    try:
        return float(match.group(1))
    except ValueError:
        return None


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ============================================================
# HTTP CLIENT
# ============================================================

class HttpClient:

    def __init__(self) -> None:

        self.session = requests.Session()

        self.session.headers.update({
            "User-Agent": USER_AGENT,
            "Accept": (
                "text/html,"
                "application/xhtml+xml,"
                "application/pdf,"
                "*/*"
            ),
            "Accept-Language": "en-SG,en;q=0.8",
        })

    def get(
        self,
        url: str,
    ) -> requests.Response | None:

        try:

            response = self.session.get(
                url,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
            )

            response.raise_for_status()

            return response

        except requests.RequestException as exc:

            logging.debug(
                "GET failed: %s: %s",
                url,
                exc,
            )

            return None


# ============================================================
# FUND UNIVERSE - XLSM
# ============================================================

def load_workbook_universe(
    path: Path,
) -> list[Fund]:

    workbook = openpyxl.load_workbook(
        path,
        read_only=True,
        data_only=True,
        keep_vba=True,
    )

    worksheet = workbook.active

    funds: list[Fund] = []

    for row_number, row in enumerate(
        worksheet.iter_rows(min_row=2),
        start=2,
    ):

        url = clean(
            row[0].value
            if len(row) >= 1
            else ""
        )

        name = clean(
            row[1].value
            if len(row) >= 2
            else ""
        )

        if not url.lower().startswith(
            ("http://", "https://")
        ):
            continue

        funds.append(
            Fund(
                row=row_number,
                name=name or f"Fund row {row_number}",
                prudential_url=url,
                source_universe=str(path),
            )
        )

    workbook.close()

    return funds


# ============================================================
# FUND UNIVERSE - JSON FALLBACK
# ============================================================

def load_json_universe(
    path: Path,
) -> list[Fund]:

    payload = json.loads(
        path.read_text(
            encoding="utf-8"
        )
    )

    if isinstance(payload, list):

        items = payload

    elif isinstance(payload, dict):

        items = []

        for key in (
            "funds",
            "data",
            "items",
        ):

            if isinstance(
                payload.get(key),
                list,
            ):

                items = payload[key]
                break

    else:

        items = []

    funds: list[Fund] = []

    for index, item in enumerate(
        items,
        start=1,
    ):

        if not isinstance(
            item,
            dict,
        ):
            continue

        name = (
            item.get("fundName")
            or item.get("name")
            or item.get("fund_name")
            or item.get("pruAccessFundName")
            or ""
        )

        url = (
            item.get("prudentialUrl")
            or item.get("prudential_url")
            or item.get("fundUrl")
            or item.get("fund_url")
            or item.get("url")
            or ""
        )

        if (
            name
            and str(url).startswith(
                ("http://", "https://")
            )
        ):

            funds.append(
                Fund(
                    row=index,
                    name=clean(name),
                    prudential_url=clean(url),
                    source_universe=str(path),
                )
            )

    return funds


# ============================================================
# LOAD MASTER UNIVERSE
# ============================================================

def load_universe(
    workbook: Path,
    json_path: Path,
) -> list[Fund]:

    if workbook.exists():

        funds = load_workbook_universe(
            workbook
        )

        if funds:
            logging.info(
                "Using master universe: %s",
                workbook,
            )

            return funds

    if json_path.exists():

        funds = load_json_universe(
            json_path
        )

        if funds:

            logging.info(
                "Using fallback universe: %s",
                json_path,
            )

            return funds

    raise RuntimeError(
        "No fund universe found. "
        f"Expected '{workbook}' or '{json_path}'."
    )


# ============================================================
# PDF CANDIDATE DISCOVERY
# ============================================================

def candidate_score(
    url: str,
    text: str,
) -> int:

    haystack = (
        f"{url} {text}"
    ).lower()

    score = 0

    if "factsheet" in haystack:
        score += 40

    if "fund factsheet" in haystack:
        score += 30

    if "fund-factsheet" in haystack:
        score += 30

    if "fund" in haystack:
        score += 5

    if "monthly" in haystack:
        score += 5

    return score


def discover_pdf_candidates(
    html: str,
    base_url: str,
    source_type: str,
) -> list[Candidate]:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    found: dict[str, Candidate] = {}

    # --------------------------------------------------------
    # Standard hyperlinks
    # --------------------------------------------------------

    for anchor in soup.find_all(
        "a",
        href=True,
    ):

        url = urljoin(
            base_url,
            anchor["href"],
        )

        if not is_pdf_url(url):
            continue

        text = clean(
            anchor.get_text(
                " ",
                strip=True,
            )
        )

        title = clean(
            anchor.get("title")
        )

        candidate = Candidate(
            url=url,
            source_type=source_type,
            discovered_from=base_url,
            title=text or title,
            score=candidate_score(
                url,
                text or title,
            ),
        )

        existing = found.get(url)

        if (
            existing is None
            or candidate.score > existing.score
        ):

            found[url] = candidate

    # --------------------------------------------------------
    # PDFs embedded in raw HTML / JavaScript
    # --------------------------------------------------------

    raw_pdf_urls = re.findall(
        r"""https?://[^"'\\\s<>]+?\.pdf(?:\?[^"'\\\s<>]*)?""",
        html,
        re.IGNORECASE,
    )

    for url in raw_pdf_urls:

        url = url.replace(
            "&amp;",
            "&",
        )

        candidate = Candidate(
            url=url,
            source_type=source_type,
            discovered_from=base_url,
            score=candidate_score(
                url,
                "",
            ),
        )

        existing = found.get(url)

        if (
            existing is None
            or candidate.score > existing.score
        ):

            found[url] = candidate

    return sorted(
        found.values(),
        key=lambda item: item.score,
        reverse=True,
    )


# ============================================================
# FETCH PDF
# ============================================================

def fetch_pdf(
    client: HttpClient,
    candidate: Candidate,
) -> tuple[bytes, str] | None:

    response = client.get(
        candidate.url
    )

    if response is None:
        return None

    data = response.content

    content_type = (
        response.headers.get(
            "Content-Type",
            "",
        ).lower()
    )

    if len(data) > MAX_PDF_BYTES:
        logging.warning(
            "Skipping oversized PDF: %s",
            candidate.url,
        )

        return None

    if (
        not data.startswith(b"%PDF")
        and "pdf" not in content_type
    ):
        return None

    return data, response.url


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf(
    pdf_bytes: bytes,
) -> tuple[str, list[dict[str, Any]]]:

    pages: list[dict[str, Any]] = []

    chunks: list[str] = []

    with pdfplumber.open(
        io.BytesIO(pdf_bytes)
    ) as pdf:

        for page_number, page in enumerate(
            pdf.pages,
            start=1,
        ):

            text = page.extract_text(
                x_tolerance=2,
                y_tolerance=3,
            ) or ""

            try:

                tables = (
                    page.extract_tables()
                    or []
                )

            except Exception:

                tables = []

            pages.append({
                "page": page_number,
                "text": text,
                "tables": tables,
            })

            if text:

                chunks.append(
                    f"[PAGE {page_number}]\n{text}"
                )

            for table in tables:

                for row in table:

                    cells = [
                        clean(cell)
                        for cell in row
                        if cell is not None
                    ]

                    if cells:

                        chunks.append(
                            " | ".join(cells)
                        )

    return (
        "\n".join(chunks),
        pages,
    )


# ============================================================
# CATEGORY VALIDATION
# ============================================================

def valid_category(
    value: str,
) -> bool:

    value = clean(value)

    if not value:
        return False

    if len(value) > 100:
        return False

    if not re.search(
        r"[A-Za-z]",
        value,
    ):
        return False

    bad_prefixes = (
        "total",
        "source",
        "benchmark",
        "portfolio",
        "as at",
        "date",
        "isin",
        "sedol",
        "factsheet",
        "performance",
        "dividend",
        "fund size",
        "investment objective",
    )

    low = value.lower()

    return not low.startswith(
        bad_prefixes
    )


# ============================================================
# PARSE LINE ALLOCATION
# ============================================================

def parse_allocation_line(
    line: str,
) -> tuple[str, float] | None:

    line = clean(line)

    match = re.search(
        r"^(.*?)"
        r"\s+"
        r"(?:[|:.\u00b7]+\s*)?"
        r"(\d{1,3}(?:[.,]\d+)?)"
        r"\s*%\s*$",
        line,
    )

    if match:

        category = clean(
            match.group(1)
        )

        try:

            pct = float(
                match.group(2).replace(
                    ",",
                    "",
                )
            )

        except ValueError:

            return None

    else:

        matches = list(
            PERCENT_RE.finditer(line)
        )

        if len(matches) != 1:
            return None

        match = matches[0]

        category = clean(
            line[:match.start()]
        )

        pct = extract_percentage(
            match.group(0)
        )

        if pct is None:
            return None

    category = re.sub(
        r"^[•·▪◦\-\u2013\u2014]+",
        "",
        category,
    ).strip()

    if not 0 <= pct <= 100:
        return None

    if not valid_category(category):
        return None

    return (
        category,
        round(pct, 4),
    )


# ============================================================
# FIND ALLOCATION HEADINGS
# ============================================================

def heading_positions(
    lines: list[str],
    headings: tuple[str, ...],
) -> list[int]:

    positions = []

    for index, line in enumerate(lines):

        low = line.lower()

        if any(
            heading in low
            for heading in headings
        ):

            positions.append(index)

    return positions


# ============================================================
# LINE-BASED EXTRACTION
# ============================================================

def extract_line_allocations(
    lines: list[str],
    headings: tuple[str, ...],
) -> list[dict[str, Any]]:

    results = []

    for start in heading_positions(
        lines,
        headings,
    ):

        for index in range(
            start + 1,
            min(
                start + 46,
                len(lines),
            ),
        ):

            line = clean(
                lines[index]
            )

            if not line:
                continue

            low = line.lower()

            if (
                index > start + 3
                and any(
                    heading in low
                    for heading in headings
                )
                and not PERCENT_RE.search(line)
            ):
                break

            parsed = parse_allocation_line(
                line
            )

            if parsed is None:
                continue

            category, pct = parsed

            results.append({
                "category": category,
                "percentage": pct,
                "rawLine": line,
            })

    return deduplicate(
        results
    )


# ============================================================
# TABLE-BASED EXTRACTION
# ============================================================

def extract_table_allocations(
    pages: list[dict[str, Any]],
    headings: tuple[str, ...],
) -> list[dict[str, Any]]:

    results = []

    for page in pages:

        page_text = clean(
            page.get("text", "")
        ).lower()

        if not any(
            heading in page_text
            for heading in headings
        ):
            continue

        for table in page.get(
            "tables",
            [],
        ):

            for row in table:

                cells = [
                    clean(cell)
                    for cell in row
                    if cell is not None
                ]

                if not cells:
                    continue

                percentages = [
                    extract_percentage(cell)
                    for cell in cells
                ]

                percentages = [
                    pct
                    for pct in percentages
                    if pct is not None
                ]

                if len(percentages) != 1:
                    continue

                pct = percentages[0]

                category_parts = [
                    cell
                    for cell in cells
                    if not PERCENT_RE.search(
                        cell
                    )
                ]

                category = clean(
                    " ".join(
                        category_parts
                    )
                )

                if (
                    valid_category(category)
                    and 0 <= pct <= 100
                ):

                    results.append({
                        "category": category,
                        "percentage": round(
                            pct,
                            4,
                        ),
                        "rawLine": " | ".join(
                            cells
                        ),
                    })

    return deduplicate(
        results
    )


# ============================================================
# DEDUPLICATION
# ============================================================

def deduplicate(
    items: list[dict[str, Any]],
) -> list[dict[str, Any]]:

    seen = set()

    output = []

    for item in items:

        key = (
            normalise(
                item["category"]
            ),
            item["percentage"],
        )

        if key in seen:
            continue

        seen.add(key)

        output.append(item)

    return output


# ============================================================
# DOCUMENT METADATA
# ============================================================

def extract_manager_name(
    text: str,
) -> str | None:

    patterns = (
        r"(?:fund|investment|asset)"
        r"\s+manager\s*[:\-]\s*(.+)",

        r"(?:managed|sub[- ]managed)"
        r"\s+by\s*[:\-]?\s*(.+)",
    )

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE,
        )

        if match:

            value = clean(
                match.group(1)
            )

            if 2 <= len(value) <= 120:
                return value

    return None


def extract_factsheet_date(
    text: str,
) -> str | None:

    match = DATE_RE.search(text)

    if match:
        return match.group(0)

    return None


# ============================================================
# PARSE COMPLETE PDF
# ============================================================

def parse_pdf(
    pdf_bytes: bytes,
    candidate: Candidate,
) -> dict[str, Any]:

    text, pages = extract_pdf(
        pdf_bytes
    )

    lines = [
        clean(line)
        for line in text.splitlines()
        if clean(line)
    ]

    geographic = extract_line_allocations(
        lines,
        GEOGRAPHIC_HEADINGS,
    )

    geographic.extend(
        extract_table_allocations(
            pages,
            GEOGRAPHIC_HEADINGS,
        )
    )

    geographic = deduplicate(
        geographic
    )

    sector = extract_line_allocations(
        lines,
        SECTOR_HEADINGS,
    )

    sector.extend(
        extract_table_allocations(
            pages,
            SECTOR_HEADINGS,
        )
    )

    sector = deduplicate(
        sector
    )

    return {
        "factsheetDate": extract_factsheet_date(
            text
        ),
        "managerName": extract_manager_name(
            text
        ),
        "geographic": geographic,
        "sector": sector,
        "sha256": sha256_bytes(
            pdf_bytes
        ),
        "candidate": asdict(
            candidate
        ),
    }


# ============================================================
# MANAGER FALLBACK DISCOVERY
# ============================================================

def discover_manager_candidates(
    client: HttpClient,
    html: str,
    base_url: str,
) -> list[Candidate]:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    candidates: dict[str, Candidate] = {}

    manager_pages = []

    for anchor in soup.find_all(
        "a",
        href=True,
    ):

        url = urljoin(
            base_url,
            anchor["href"],
        )

        text = clean(
            anchor.get_text(
                " ",
                strip=True,
            )
        )

        title = clean(
            anchor.get("title")
        )

        haystack = (
            f"{url} {text} {title}"
        ).lower()

        # Direct manager factsheet.
        if is_pdf_url(url):

            if any(
                term in haystack
                for term in (
                    "factsheet",
                    "fund manager",
                    "investment manager",
                )
            ):

                candidates[url] = Candidate(
                    url=url,
                    source_type="fund_manager",
                    discovered_from=base_url,
                    title=text or title,
                    score=candidate_score(
                        url,
                        text or title,
                    ),
                )

        # Manager-related landing page.
        elif any(
            term in text.lower()
            for term in (
                "fund manager",
                "investment manager",
                "manager",
            )
        ):

            manager_pages.append(url)

    # --------------------------------------------------------
    # Follow only manager-related links exposed by Prudential.
    # --------------------------------------------------------

    for manager_page in list(
        dict.fromkeys(manager_pages)
    )[:5]:

        response = client.get(
            manager_page
        )

        if response is None:
            continue

        page_candidates = discover_pdf_candidates(
            response.text,
            response.url,
            "fund_manager",
        )

        for candidate in page_candidates:

            candidate.source_type = (
                "fund_manager"
            )

            candidates[
                candidate.url
            ] = candidate

    return sorted(
        candidates.values(),
        key=lambda item: item.score,
        reverse=True,
    )


# ============================================================
# EMPTY RESULT
# ============================================================

def unresolved_result(
    fund: Fund,
    reason: str,
) -> dict[str, Any]:

    return {
        "fundName": fund.name,
        "prudentialUrl": fund.prudential_url,
        "sourceUniverse": fund.source_universe,

        "source": None,

        "factsheetDate": None,
        "managerName": None,

        "geographicAllocation": [],
        "sectorAllocation": [],

        "status": "unresolved",
        "reason": reason,

        "retrievedAt": now_utc(),
    }


# ============================================================
# PROCESS ONE FUND
# ============================================================

def process_fund(
    client: HttpClient,
    fund: Fund,
) -> tuple[
    dict[str, Any],
    dict[str, Any],
]:

    diagnostics = {
        "fundName": fund.name,
        "row": fund.row,
        "prudentialUrl": fund.prudential_url,
        "attemptedSources": [],
        "errors": [],
    }

    # --------------------------------------------------------
    # Fetch Prudential landing page.
    # --------------------------------------------------------

    prudential_page = client.get(
        fund.prudential_url
    )

    if prudential_page is None:

        diagnostics["errors"].append(
            "Unable to fetch Prudential fund page"
        )

        return (
            unresolved_result(
                fund,
                "Prudential fund page unavailable",
            ),
            diagnostics,
        )

    # --------------------------------------------------------
    # PRIMARY SOURCE
    # Official Prudential factsheet.
    # --------------------------------------------------------

    candidates = discover_pdf_candidates(
        prudential_page.text,
        prudential_page.url,
        "prudential",
    )

    prudential_candidates = [
        candidate
        for candidate in candidates
        if is_prudential_url(
            candidate.url
        )
    ]

    for candidate in prudential_candidates[:15]:

        diagnostics[
            "attemptedSources"
        ].append({
            "type": "prudential",
            "url": candidate.url,
        })

        fetched = fetch_pdf(
            client,
            candidate,
        )

        if fetched is None:
            continue

        pdf_bytes, final_url = fetched

        try:

            parsed = parse_pdf(
                pdf_bytes,
                candidate,
            )

        except Exception as exc:

            diagnostics[
                "errors"
            ].append(
                f"Prudential PDF parse failed: "
                f"{candidate.url}: {exc}"
            )

            continue

        if not (
            parsed["geographic"]
            or parsed["sector"]
        ):
            continue

        source = {
            "type": "prudential",
            "url": final_url,
            "title": candidate.title,
            "sha256": parsed["sha256"],
        }

        diagnostics[
            "selectedSource"
        ] = source

        return (
            {
                "fundName": fund.name,
                "prudentialUrl": fund.prudential_url,
                "sourceUniverse": fund.source_universe,

                "source": source,

                "factsheetDate": parsed[
                    "factsheetDate"
                ],

                "managerName": parsed[
                    "managerName"
                ],

                "geographicAllocation": parsed[
                    "geographic"
                ],

                "sectorAllocation": parsed[
                    "sector"
                ],

                "status": "resolved",

                "reason": (
                    "Published allocation data "
                    "extracted from official "
                    "Prudential factsheet"
                ),

                "retrievedAt": now_utc(),
            },
            diagnostics,
        )

    # --------------------------------------------------------
    # FALLBACK
    # Official underlying fund manager factsheet.
    # --------------------------------------------------------

    manager_candidates = (
        discover_manager_candidates(
            client,
            prudential_page.text,
            prudential_page.url,
        )
    )

    for candidate in manager_candidates[:15]:

        diagnostics[
            "attemptedSources"
        ].append({
            "type": "fund_manager",
            "url": candidate.url,
        })

        fetched = fetch_pdf(
            client,
            candidate,
        )

        if fetched is None:
            continue

        pdf_bytes, final_url = fetched

        try:

            parsed = parse_pdf(
                pdf_bytes,
                candidate,
            )

        except Exception as exc:

            diagnostics[
                "errors"
            ].append(
                f"Manager PDF parse failed: "
                f"{candidate.url}: {exc}"
            )

            continue

        if not (
            parsed["geographic"]
            or parsed["sector"]
        ):
            continue

        source = {
            "type": "fund_manager",
            "url": final_url,
            "title": candidate.title,
            "sha256": parsed["sha256"],
        }

        diagnostics[
            "selectedSource"
        ] = source

        return (
            {
                "fundName": fund.name,
                "prudentialUrl": fund.prudential_url,
                "sourceUniverse": fund.source_universe,

                "source": source,

                "factsheetDate": parsed[
                    "factsheetDate"
                ],

                "managerName": parsed[
                    "managerName"
                ],

                "geographicAllocation": parsed[
                    "geographic"
                ],

                "sectorAllocation": parsed[
                    "sector"
                ],

                "status": "resolved",

                "reason": (
                    "Published allocation data "
                    "extracted from official "
                    "underlying fund-manager factsheet"
                ),

                "retrievedAt": now_utc(),
            },
            diagnostics,
        )

    # --------------------------------------------------------
    # UNRESOLVED
    # --------------------------------------------------------

    return (
        unresolved_result(
            fund,
            (
                "No usable published geographic "
                "or sector allocation found"
            ),
        ),
        diagnostics,
    )


# ============================================================
# JSON WRITER
# ============================================================

def write_json(
    path: Path,
    payload: Any,
) -> None:

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    path.write_text(
        json.dumps(
            payload,
            indent=2,
            ensure_ascii=False,
        )
        + "\n",
        encoding="utf-8",
    )


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--workbook",
        default="Funds Links.xlsm",
    )

    parser.add_argument(
        "--funds-json",
        default="data/funds.json",
    )

    parser.add_argument(
        "--output",
        default=(
            "data/"
            "fundfactsheet_exposure.json"
        ),
    )

    parser.add_argument(
        "--diagnostics",
        default=(
            "data/"
            "fundfactsheet_exposure_diagnostics.json"
        ),
    )

    parser.add_argument(
        "--delay",
        type=float,
        default=DEFAULT_DELAY,
    )

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format=(
            "%(asctime)s | "
            "%(levelname)s | "
            "%(message)s"
        ),
    )

    funds = load_universe(
        Path(args.workbook),
        Path(args.funds_json),
    )

    logging.info(
        "Loaded %d funds",
        len(funds),
    )

    client = HttpClient()

    results = []
    diagnostics = []

    for index, fund in enumerate(
        funds,
        start=1,
    ):

        logging.info(
            "[%d/%d] %s",
            index,
            len(funds),
            fund.name,
        )

        try:

            result, diagnostic = process_fund(
                client,
                fund,
            )

        except Exception as exc:

            logging.exception(
                "Unhandled error for %s",
                fund.name,
            )

            result = unresolved_result(
                fund,
                f"Unhandled extraction error: {exc}",
            )

            diagnostic = {
                "fundName": fund.name,
                "row": fund.row,
                "prudentialUrl": fund.prudential_url,
                "attemptedSources": [],
                "errors": [str(exc)],
            }

        results.append(result)
        diagnostics.append(diagnostic)

        time.sleep(
            max(
                args.delay,
                0,
            )
        )

    # ========================================================
    # SUMMARY
    # ========================================================

    resolved = sum(
        item["status"] == "resolved"
        for item in results
    )

    unresolved = (
        len(results)
        - resolved
    )

    prudential_count = sum(
        isinstance(
            item.get("source"),
            dict,
        )
        and item["source"].get(
            "type"
        ) == "prudential"
        for item in results
    )

    manager_count = sum(
        isinstance(
            item.get("source"),
            dict,
        )
        and item["source"].get(
            "type"
        ) == "fund_manager"
        for item in results
    )

    # ========================================================
    # MAIN OUTPUT
    # ========================================================

    output = {
        "schemaVersion": "1.0.0",

        "generatedAt": now_utc(),

        "sourcePolicy": {
            "primary": (
                "Official Prudential "
                "Singapore factsheet"
            ),

            "fallback": (
                "Official underlying "
                "fund-manager factsheet"
            ),

            "unresolvedRule": (
                "No usable published "
                "allocation data; "
                "no inference"
            ),
        },

        "summary": {
            "totalFunds": len(results),
            "resolved": resolved,
            "unresolved": unresolved,
            "prudentialFactsheets": (
                prudential_count
            ),
            "fundManagerFactsheets": (
                manager_count
            ),
        },

        "funds": results,
    }

    # ========================================================
    # DIAGNOSTICS OUTPUT
    # ========================================================

    diagnostic_output = {
        "schemaVersion": "1.0.0",

        "generatedAt": now_utc(),

        "summary": {
            "totalFunds": len(results),
            "resolved": resolved,
            "unresolved": unresolved,
            "prudentialFactsheets": (
                prudential_count
            ),
            "fundManagerFactsheets": (
                manager_count
            ),
        },

        "funds": diagnostics,
    }

    write_json(
        Path(args.output),
        output,
    )

    write_json(
        Path(args.diagnostics),
        diagnostic_output,
    )

    logging.info(
        "Completed: %d resolved / %d unresolved",
        resolved,
        unresolved,
    )

    return 0


if __name__ == "__main__":
    sys.exit(main())
