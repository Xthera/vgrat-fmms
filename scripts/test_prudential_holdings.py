#!/usr/bin/env python3

"""
Prudential Singapore official fund Top Holdings extractor.

MASTER SOURCE
=============

Excel:
    Funds Links.xlsm

Column A:
    Prudential Singapore fund URL.

Column B:
    Exact PruAccess fund name.

SCOPE
=====

This script processes every populated URL in Excel Column A.

For each fund it:

1. Opens the official Prudential Singapore fund page.
2. Finds the official Prudential factsheet link.
3. Downloads the official Prudential factsheet PDF.
4. Extracts the "Top 10 Holdings" section.
5. Parses up to 10 holdings.
6. Preserves the exact published holding order.
7. Preserves multiline holding names.
8. Preserves published percentages when available.
9. Stores null when Prudential publishes a holding without a percentage.
10. Performs strict validation.
11. Saves the raw PDF, extracted text, JSON and metadata.
12. Produces a run summary.

HARD RULES
==========

- Excel Column A controls the fund universe.
- No hardcoded fund count.
- Only official Prudential Singapore sources.
- No third-party holdings sources.
- No inferred holding names.
- No fabricated holding names.
- No calculated holding percentages.
- No inferred holding percentages.
- No forced 10 holdings.
- If Prudential publishes fewer than 10 holdings, store exactly that count.
- Holdings remain in Prudential's published order.
- Multiline holding names are preserved/joined.
- Published percentage is optional.
- Missing published percentage = null.
- Rank is the primary holding-row boundary.
- Percentage does NOT define the holding boundary.
- Duplicate percentages are allowed.
- Duplicate holding names are rejected.
- If a Top Holdings section exists but holdings cannot be identified reliably,
  the fund fails.
- If no Top Holdings section exists, the fund is classified as
  "no_holdings_section".
- No synthetic data.
- No interpolation.
- No estimates.
- No carry-forward.
- This script does not modify PruAccess extraction.
- This script does not create data.json.
- This script does not modify frontend files.
- This is a holdings collector/validator only.
"""

from __future__ import annotations

import json
import re
import sys
import time
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any, Optional
from urllib.parse import urljoin

import openpyxl
import requests
from bs4 import BeautifulSoup
from pypdf import PdfReader


# ============================================================================
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_holdings")

MAX_HOLDINGS = 10

REQUEST_TIMEOUT_SECONDS = 60

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/153.0.0.0 Safari/537.36"
)

HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,*/*;q=0.8"
    ),
}

PDF_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "application/pdf,*/*",
}

TOP_HOLDINGS_PATTERNS = [
    re.compile(r"\btop\s*10\s*holdings\b", re.IGNORECASE),
    re.compile(r"\btop\s*ten\s*holdings\b", re.IGNORECASE),
]

SECTION_END_PATTERNS = [
    re.compile(r"^\s*asset\s+allocation\s*$", re.IGNORECASE),
    re.compile(r"^\s*asset\s+allocation\b", re.IGNORECASE),
    re.compile(r"^\s*geographical\s+allocation\b", re.IGNORECASE),
    re.compile(r"^\s*sector\s+allocation\b", re.IGNORECASE),
    re.compile(r"^\s*country\s+allocation\b", re.IGNORECASE),
    re.compile(r"^\s*fund\s+characteristics\b", re.IGNORECASE),
    re.compile(r"^\s*fund\s+information\b", re.IGNORECASE),
    re.compile(r"^\s*investment\s+objective\b", re.IGNORECASE),
    re.compile(r"^\s*risk\s+classification\b", re.IGNORECASE),
    re.compile(r"^\s*performance\b", re.IGNORECASE),
    re.compile(r"^\s*past\s+performance\b", re.IGNORECASE),
    re.compile(r"^\s*important\s+information\b", re.IGNORECASE),
    re.compile(r"^\s*disclaimer\b", re.IGNORECASE),
    re.compile(r"^\s*source\b", re.IGNORECASE),
]


# ============================================================================
# GENERAL HELPERS
# ============================================================================

def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def clean_text(value: Any) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")

    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s+\n", "\n", text)
    text = re.sub(r"\n\s+", "\n", text)

    return text.strip()


def safe_filename(
    value: str,
    max_length: int = 180,
) -> str:
    value = clean_text(value)

    value = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        "_",
        value,
    )

    value = re.sub(r"\s+", "_", value)
    value = re.sub(r"_+", "_", value)

    value = value.strip(" ._")

    if not value:
        value = "fund"

    return value[:max_length]


def normalize_name(value: str) -> str:
    value = clean_text(value)

    value = value.lower()

    value = value.replace("–", "-")
    value = value.replace("—", "-")
    value = value.replace("−", "-")

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def contains_percentage(text: str) -> bool:
    return bool(
        re.search(
            r"[+-]?\d+(?:\.\d+)?\s*%",
            text or "",
        )
    )


def is_rank_line(
    line: str,
) -> Optional[int]:
    """
    Recognise standalone holding ranks.

    Examples:
        1
        2
        10
        1.
        2)
        10 -
    """

    text = clean_text(line)

    if not text:
        return None

    match = re.fullmatch(
        r"(\d{1,2})\s*[\.\):\-]?\s*",
        text,
    )

    if not match:
        return None

    rank = int(match.group(1))

    if rank < 1 or rank > MAX_HOLDINGS:
        return None

    return rank


def split_rank_and_text(
    line: str,
) -> tuple[Optional[int], str]:
    """
    Recognise:

        1 Holding Name
        1. Holding Name
        1) Holding Name
        10 Holding Name
    """

    text = clean_text(line)

    if not text:
        return None, ""

    match = re.match(
        r"^(\d{1,2})\s*[\.\):\-]\s*(.+)$",
        text,
    )

    if match:
        rank = int(match.group(1))

        if 1 <= rank <= MAX_HOLDINGS:
            return rank, clean_text(match.group(2))

    match = re.match(
        r"^(\d{1,2})\s+(.+)$",
        text,
    )

    if match:
        rank = int(match.group(1))

        if 1 <= rank <= MAX_HOLDINGS:
            remainder = clean_text(match.group(2))

            if remainder and not contains_percentage(
                remainder
            ):
                return rank, remainder

    return None, text


def is_probable_header_or_noise(
    line: str,
) -> bool:
    text = clean_text(line)

    if not text:
        return True

    lower = text.lower()

    noise_exact = {
        "top 10 holdings",
        "top ten holdings",
        "holdings",
        "holding",
        "source",
        "sources",
        "asset allocation",
        "sector allocation",
        "geographical allocation",
        "country allocation",
        "name",
        "names",
        "weight",
        "weights",
        "%",
        "percentage",
        "market value",
        "quantity",
        "shares",
        "total",
    }

    if lower in noise_exact:
        return True

    if lower.startswith("top 10 holdings"):
        return True

    if lower.startswith("top ten holdings"):
        return True

    return False


def looks_like_footer(
    line: str,
) -> bool:
    text = clean_text(line)

    if not text:
        return False

    lower = text.lower()

    footer_markers = [
        "past performance",
        "investment involves risk",
        "important information",
        "this document",
        "prudential assurance",
        "prudential corporation asia",
        "source:",
        "disclaimer",
        "not an offer",
        "not investment advice",
    ]

    return any(
        marker in lower
        for marker in footer_markers
    )


# ============================================================================
# EXCEL
# ============================================================================

def read_excel_funds() -> list[dict[str, Any]]:
    """
    Read every populated URL in Excel Column A.

    Column B is retained as the exact PruAccess fund name.

    Excel Column A is the sole universe controller.
    """

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = openpyxl.load_workbook(
        EXCEL_FILE,
        data_only=True,
        read_only=True,
    )

    worksheet = workbook.active

    funds: list[dict[str, Any]] = []

    for row_number, row in enumerate(
        worksheet.iter_rows(
            min_col=1,
            max_col=2,
            values_only=True,
        ),
        start=1,
    ):
        url_value = row[0]
        pruaccess_name = row[1]

        if url_value is None:
            continue

        url = clean_text(url_value)

        if not url:
            continue

        if not url.lower().startswith(
            ("http://", "https://")
        ):
            continue

        funds.append(
            {
                "excelRow": row_number,
                "url": url,
                "pruAccessFundName": clean_text(
                    pruaccess_name
                ),
            }
        )

    workbook.close()

    return funds


# ============================================================================
# HTTP
# ============================================================================

def create_session() -> requests.Session:
    session = requests.Session()

    session.headers.update(
        HEADERS
    )

    return session


def download_url(
    session: requests.Session,
    url: str,
    *,
    expect_pdf: bool = False,
) -> tuple[str, bytes]:
    response = session.get(
        url,
        headers=(
            PDF_HEADERS
            if expect_pdf
            else HEADERS
        ),
        timeout=REQUEST_TIMEOUT_SECONDS,
        allow_redirects=True,
    )

    response.raise_for_status()

    final_url = response.url

    content = response.content

    if not content:
        raise RuntimeError(
            f"Downloaded empty response from {url}"
        )

    if expect_pdf:
        content_type = (
            response.headers.get(
                "content-type"
            )
            or ""
        ).lower()

        if (
            not content.startswith(b"%PDF")
            and "pdf" not in content_type
        ):
            raise RuntimeError(
                f"Expected PDF but received "
                f"content-type={content_type!r} "
                f"from {final_url}"
            )

    return final_url, content


# ============================================================================
# PRUDENTIAL FUND PAGE
# ============================================================================

def find_factsheet_url(
    page_url: str,
    html: str,
) -> Optional[str]:
    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    candidates: list[
        tuple[int, str]
    ] = []

    for anchor in soup.find_all("a"):
        href = anchor.get("href")

        if not href:
            continue

        absolute = urljoin(
            page_url,
            href,
        )

        text = clean_text(
            anchor.get_text(
                " ",
                strip=True,
            )
        )

        combined = (
            f"{text} {absolute}"
        ).lower()

        score = 0

        if "factsheet" in combined:
            score += 100

        if ".pdf" in absolute.lower():
            score += 50

        if "fund" in combined:
            score += 10

        if score > 0:
            candidates.append(
                (score, absolute)
            )

    if not candidates:
        return None

    candidates.sort(
        key=lambda item: (
            -item[0],
            item[1],
        )
    )

    return candidates[0][1]


def extract_fund_page_metadata(
    html: str,
) -> dict[str, Any]:
    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    title = clean_text(
        soup.title.get_text(
            " ",
            strip=True,
        )
        if soup.title
        else ""
    )

    return {
        "pageTitle": title,
    }


# ============================================================================
# PDF EXTRACTION
# ============================================================================

def extract_pdf_text(
    pdf_bytes: bytes,
    *,
    extraction_mode: Optional[str] = None,
) -> str:
    """
    Extract PDF text using pypdf.

    Both default and layout extraction are tested.
    """

    reader = PdfReader(
        BytesIO(pdf_bytes)
    )

    pages: list[str] = []

    for page in reader.pages:
        if extraction_mode == "layout":
            try:
                text = page.extract_text(
                    extraction_mode="layout"
                )
            except TypeError:
                text = page.extract_text()
        else:
            text = page.extract_text()

        if text:
            pages.append(text)

    return "\n".join(pages)


def extract_top_holdings_section(
    pdf_text: str,
) -> tuple[
    Optional[str],
    Optional[int],
]:
    """
    Locate the Top 10 Holdings section.
    """

    if not pdf_text:
        return None, None

    lines = pdf_text.splitlines()

    start_index: Optional[int] = None

    for index, raw_line in enumerate(
        lines
    ):
        line = clean_text(raw_line)

        if not line:
            continue

        if any(
            pattern.search(line)
            for pattern in TOP_HOLDINGS_PATTERNS
        ):
            start_index = index
            break

    if start_index is None:
        return None, None

    section_lines: list[str] = []

    for index in range(
        start_index + 1,
        len(lines),
    ):
        line = clean_text(lines[index])

        if not line:
            continue

        if any(
            pattern.search(line)
            for pattern in SECTION_END_PATTERNS
        ):
            break

        if looks_like_footer(line):
            break

        section_lines.append(line)

    section_text = "\n".join(
        section_lines
    ).strip()

    return (
        section_text,
        start_index,
    )


# ============================================================================
# HOLDING PARSER
# ============================================================================

def extract_percentage_from_line(
    line: str,
) -> tuple[
    Optional[float],
    Optional[str],
    str,
]:
    """
    Extract an explicitly published percentage.

    No calculation or inference is performed.
    """

    text = clean_text(line)

    match = re.search(
        r"([+-]?\d+(?:\.\d+)?)\s*%",
        text,
    )

    if not match:
        return (
            None,
            None,
            text,
        )

    raw_percent = match.group(0)

    try:
        value = float(
            match.group(1)
        )
    except ValueError:
        return (
            None,
            None,
            text,
        )

    remaining = (
        text[:match.start()]
        + " "
        + text[match.end():]
    )

    remaining = clean_text(
        remaining
    )

    return (
        value,
        raw_percent,
        remaining,
    )


def clean_holding_name(
    fragments: list[str],
) -> str:
    cleaned: list[str] = []

    for fragment in fragments:
        text = clean_text(
            fragment
        )

        if not text:
            continue

        if is_probable_header_or_noise(
            text
        ):
            continue

        cleaned.append(text)

    if not cleaned:
        return ""

    name = " ".join(
        cleaned
    )

    name = re.sub(
        r"\s+",
        " ",
        name,
    )

    return name.strip()


def parse_holdings(
    section_text: str,
) -> dict[str, Any]:
    """
    Parse the Top 10 Holdings section.

    IMPORTANT:

    Rank is the primary holding boundary.

    Percentage is optional.

    Therefore this is valid:

        1
        Holding A

        2
        Holding B

    and produces null weights.

    This is also valid:

        1
        Holding A
        9.7%

        2
        Holding B
        33.5%
    """

    if not section_text:
        return {
            "status": "failed",
            "reason": (
                "empty_holdings_section"
            ),
            "holdings": [],
            "diagnostics": {},
        }

    lines = [
        clean_text(line)
        for line in section_text.splitlines()
    ]

    lines = [
        line
        for line in lines
        if line
    ]

    holdings: list[
        dict[str, Any]
    ] = []

    current_rank: Optional[int] = None
    current_fragments: list[str] = []
    current_weight_percent: Optional[
        float
    ] = None
    current_weight_text: Optional[
        str
    ] = None

    explicit_ranks: list[int] = []
    percentage_count = 0

    def finalize_current() -> None:
        nonlocal current_rank
        nonlocal current_fragments
        nonlocal current_weight_percent
        nonlocal current_weight_text

        if current_rank is None:
            return

        name = clean_holding_name(
            current_fragments
        )

        if not name:
            raise ValueError(
                f"Holding rank "
                f"{current_rank} has no "
                f"identifiable holding name."
            )

        holdings.append(
            {
                "rank": current_rank,
                "name": name,
                "weightPercent": (
                    current_weight_percent
                ),
                "weightText": (
                    current_weight_text
                ),
            }
        )

        current_rank = None
        current_fragments = []
        current_weight_percent = None
        current_weight_text = None

    for raw_line in lines:
        line = clean_text(
            raw_line
        )

        if not line:
            continue

        # ------------------------------------------------------------
        # Standalone rank
        # ------------------------------------------------------------

        standalone_rank = is_rank_line(
            line
        )

        if standalone_rank is not None:
            if current_rank is not None:
                finalize_current()

            current_rank = (
                standalone_rank
            )

            explicit_ranks.append(
                standalone_rank
            )

            continue

        # ------------------------------------------------------------
        # Inline rank + holding
        # ------------------------------------------------------------

        inline_rank, inline_text = (
            split_rank_and_text(line)
        )

        if inline_rank is not None:
            if current_rank is not None:
                finalize_current()

            current_rank = inline_rank

            explicit_ranks.append(
                inline_rank
            )

            if inline_text:
                (
                    weight,
                    weight_text,
                    remaining,
                ) = extract_percentage_from_line(
                    inline_text
                )

                if weight is not None:
                    current_weight_percent = (
                        weight
                    )
                    current_weight_text = (
                        weight_text
                    )
                    percentage_count += 1

                if remaining:
                    current_fragments.append(
                        remaining
                    )

            continue

        # ------------------------------------------------------------
        # Ignore obvious table headings/noise
        # ------------------------------------------------------------

        if is_probable_header_or_noise(
            line
        ):
            continue

        # ------------------------------------------------------------
        # Explicit percentage
        # ------------------------------------------------------------

        (
            weight,
            weight_text,
            remaining,
        ) = extract_percentage_from_line(
            line
        )

        if (
            current_rank is not None
            and weight is not None
        ):
            if (
                current_weight_percent
                is not None
            ):
                raise ValueError(
                    f"Holding rank "
                    f"{current_rank} contains "
                    f"multiple published "
                    f"percentages."
                )

            current_weight_percent = (
                weight
            )

            current_weight_text = (
                weight_text
            )

            percentage_count += 1

            if remaining:
                if not is_probable_header_or_noise(
                    remaining
                ):
                    current_fragments.append(
                        remaining
                    )

            continue

        # ------------------------------------------------------------
        # Normal holding-name fragment
        # ------------------------------------------------------------

        if current_rank is not None:
            if looks_like_footer(line):
                finalize_current()
                break

            current_fragments.append(
                line
            )

    if current_rank is not None:
        finalize_current()

    # ------------------------------------------------------------
    # Structural validation
    # ------------------------------------------------------------

    if not holdings:
        return {
            "status": "failed",
            "reason": (
                "Top Holdings section exists "
                "but no ranked holdings were "
                "identified."
            ),
            "holdings": [],
            "diagnostics": {
                "explicitRanks": explicit_ranks,
                "percentageCount": (
                    percentage_count
                ),
            },
        }

    if len(holdings) > MAX_HOLDINGS:
        raise ValueError(
            f"Parsed {len(holdings)} holdings, "
            f"exceeding maximum "
            f"{MAX_HOLDINGS}."
        )

    expected_rank = 1

    for holding in holdings:
        rank = holding["rank"]

        if rank != expected_rank:
            raise ValueError(
                "Holding ranks are not "
                "sequential. "
                f"Expected {expected_rank}, "
                f"received {rank}."
            )

        expected_rank += 1

    # Duplicate holding names remain a hard failure.
    normalized_names: set[str] = set()

    for holding in holdings:
        normalized = normalize_name(
            holding["name"]
        )

        if not normalized:
            raise ValueError(
                f"Holding rank "
                f"{holding['rank']} has an "
                f"empty normalized name."
            )

        if normalized in normalized_names:
            raise ValueError(
                "Duplicate holding name "
                f"detected: "
                f"{holding['name']!r}"
            )

        normalized_names.add(
            normalized
        )

    return {
        "status": "success",
        "reason": None,
        "holdings": holdings,
        "diagnostics": {
            "explicitRanks": explicit_ranks,
            "holdingCount": len(holdings),
            "percentageCount": (
                percentage_count
            ),
            "holdingsWithoutPublishedPercentage": sum(
                1
                for holding in holdings
                if holding[
                    "weightPercent"
                ]
                is None
            ),
            "holdingRankDefinesRowBoundary": True,
            "publishedPercentageOptional": True,
        },
    }


# ============================================================================
# VALIDATION
# ============================================================================

def validate_holdings(
    holdings: list[
        dict[str, Any]
    ],
) -> None:
    """
    Final strict validation.

    Missing published percentages are valid.
    """

    if not holdings:
        raise ValueError(
            "No holdings supplied for validation."
        )

    if len(holdings) > MAX_HOLDINGS:
        raise ValueError(
            f"Too many holdings: "
            f"{len(holdings)}"
        )

    seen_names: set[str] = set()

    for expected_rank, holding in enumerate(
        holdings,
        start=1,
    ):
        if holding.get("rank") != expected_rank:
            raise ValueError(
                f"Invalid holding rank. "
                f"Expected {expected_rank}, "
                f"received "
                f"{holding.get('rank')!r}."
            )

        name = clean_text(
            holding.get("name")
        )

        if not name:
            raise ValueError(
                f"Holding {expected_rank} "
                f"has no name."
            )

        normalized = normalize_name(
            name
        )

        if normalized in seen_names:
            raise ValueError(
                f"Duplicate holding name: "
                f"{name!r}"
            )

        seen_names.add(
            normalized
        )

        weight = holding.get(
            "weightPercent"
        )

        weight_text = holding.get(
            "weightText"
        )

        if weight is None:
            if weight_text not in (
                None,
                "",
            ):
                raise ValueError(
                    f"Holding "
                    f"{expected_rank} has "
                    f"weightText without "
                    f"weightPercent."
                )

        else:
            if not isinstance(
                weight,
                (int, float),
            ):
                raise ValueError(
                    f"Holding "
                    f"{expected_rank} has "
                    f"invalid "
                    f"weightPercent."
                )

            if weight < 0:
                raise ValueError(
                    f"Holding "
                    f"{expected_rank} has "
                    f"negative "
                    f"weightPercent."
                )

            if weight > 100:
                raise ValueError(
                    f"Holding "
                    f"{expected_rank} has "
                    f"weightPercent "
                    f"above 100."
                )

            if not weight_text:
                raise ValueError(
                    f"Holding "
                    f"{expected_rank} has "
                    f"weightPercent but "
                    f"no weightText."
                )


# ============================================================================
# OUTPUT
# ============================================================================

def write_json(
    path: Path,
    payload: Any,
) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with path.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            payload,
            handle,
            ensure_ascii=False,
            indent=2,
        )


def write_text(
    path: Path,
    text: str,
) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    path.write_text(
        text,
        encoding="utf-8",
    )


# ============================================================================
# FUND PROCESSING
# ============================================================================

def process_fund(
    session: requests.Session,
    fund: dict[str, Any],
    fund_index: int,
    total_funds: int,
) -> dict[str, Any]:
    started_at = utc_now_iso()

    excel_row = fund["excelRow"]
    page_url = fund["url"]
    pruaccess_name = fund[
        "pruAccessFundName"
    ]

    fund_label = (
        pruaccess_name
        or f"Excel Row {excel_row}"
    )

    print(
        f"[{fund_index}/{total_funds}] "
        f"Processing: {fund_label}"
    )

    result: dict[str, Any] = {
        "excelRow": excel_row,
        "url": page_url,
        "pruAccessFundName": (
            pruaccess_name
        ),
        "status": "failed",
        "reason": None,
        "factsheetUrl": None,
        "factsheetFinalUrl": None,
        "factsheetBytes": 0,
        "pdfExtractionMode": None,
        "holdings": [],
        "diagnostics": {},
        "startedAtUtc": started_at,
        "completedAtUtc": None,
    }

    base_filename = safe_filename(
        pruaccess_name
        or f"row_{excel_row}"
    )

    fund_dir = (
        OUTPUT_DIR
        / "funds"
        / f"{excel_row:03d}_{base_filename}"
    )

    fund_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        # ------------------------------------------------------------
        # 1. Fund page
        # ------------------------------------------------------------

        print(
            "  Downloading fund page..."
        )

        fund_final_url, html_bytes = (
            download_url(
                session,
                page_url,
                expect_pdf=False,
            )
        )

        html = html_bytes.decode(
            "utf-8",
            errors="replace",
        )

        result[
            "fundPageFinalUrl"
        ] = fund_final_url

        result[
            "fundPageMetadata"
        ] = extract_fund_page_metadata(
            html
        )

        write_text(
            fund_dir / "fund_page.html",
            html,
        )

        # ------------------------------------------------------------
        # 2. Find official factsheet
        # ------------------------------------------------------------

        factsheet_url = (
            find_factsheet_url(
                fund_final_url,
                html,
            )
        )

        if not factsheet_url:
            result["reason"] = (
                "official_factsheet_link_not_found"
            )

            print(
                "  FAILED: official factsheet "
                "link not found."
            )

            return result

        result[
            "factsheetUrl"
        ] = factsheet_url

        print(
            f"  Factsheet: "
            f"{factsheet_url}"
        )

        # ------------------------------------------------------------
        # 3. Download factsheet
        # ------------------------------------------------------------

        print(
            "  Downloading factsheet PDF..."
        )

        (
            factsheet_final_url,
            pdf_bytes,
        ) = download_url(
            session,
            factsheet_url,
            expect_pdf=True,
        )

        result[
            "factsheetFinalUrl"
        ] = factsheet_final_url

        result[
            "factsheetBytes"
        ] = len(pdf_bytes)

        pdf_path = (
            fund_dir
            / "factsheet.pdf"
        )

        pdf_path.write_bytes(
            pdf_bytes
        )

        # ------------------------------------------------------------
        # 4. Extract PDF text
        # ------------------------------------------------------------

        extraction_candidates: list[
            tuple[str, str, int]
        ] = []

        default_text = extract_pdf_text(
            pdf_bytes,
            extraction_mode=None,
        )

        (
            default_section,
            default_start,
        ) = extract_top_holdings_section(
            default_text
        )

        if default_section:
            extraction_candidates.append(
                (
                    "default",
                    default_section,
                    len(
                        default_section.splitlines()
                    ),
                )
            )

        layout_text = extract_pdf_text(
            pdf_bytes,
            extraction_mode="layout",
        )

        (
            layout_section,
            layout_start,
        ) = extract_top_holdings_section(
            layout_text
        )

        if layout_section:
            extraction_candidates.append(
                (
                    "layout",
                    layout_section,
                    len(
                        layout_section.splitlines()
                    ),
                )
            )

        write_text(
            fund_dir
            / "extracted_text_default.txt",
            default_text,
        )

        write_text(
            fund_dir
            / "extracted_text_layout.txt",
            layout_text,
        )

        # ------------------------------------------------------------
        # 5. Detect Top Holdings
        # ------------------------------------------------------------

        if not extraction_candidates:
            result["status"] = (
                "no_holdings_section"
            )

            result["reason"] = (
                "No Top 10 Holdings section "
                "was found in the official "
                "factsheet."
            )

            print(
                "  NO HOLDINGS SECTION"
            )

            return result

        # ------------------------------------------------------------
        # 6. Parse candidate extraction modes
        # ------------------------------------------------------------

        parsed_candidates: list[
            tuple[str, dict[str, Any]]
        ] = []

        parser_errors: dict[
            str,
            str,
        ] = {}

        for (
            mode,
            section,
            _line_count,
        ) in extraction_candidates:
            try:
                parsed = parse_holdings(
                    section
                )

                if (
                    parsed.get("status")
                    == "success"
                ):
                    validate_holdings(
                        parsed["holdings"]
                    )

                    parsed_candidates.append(
                        (
                            mode,
                            parsed,
                        )
                    )

                else:
                    parser_errors[
                        mode
                    ] = (
                        parsed.get(
                            "reason"
                        )
                        or "parser_failed"
                    )

            except Exception as exc:
                parser_errors[
                    mode
                ] = str(exc)

        # ------------------------------------------------------------
        # 7. Select valid parser result
        # ------------------------------------------------------------

        if not parsed_candidates:
            result["reason"] = (
                "Top Holdings section exists, "
                "but no structurally valid "
                "ranked holdings could be "
                "extracted."
            )

            result["diagnostics"] = {
                "parserErrors": (
                    parser_errors
                ),
                "availableExtractionModes": [
                    item[0]
                    for item in extraction_candidates
                ],
            }

            print(
                "  FAILED: holdings section "
                "found but parsing failed."
            )

            for mode, error in (
                parser_errors.items()
            ):
                print(
                    f"    {mode}: {error}"
                )

            return result

        # Prefer the result with the most
        # successfully parsed holdings.
        parsed_candidates.sort(
            key=lambda item: (
                -len(
                    item[1]["holdings"]
                ),
                0
                if item[0] == "layout"
                else 1,
            )
        )

        (
            selected_mode,
            selected,
        ) = parsed_candidates[0]

        holdings = selected[
            "holdings"
        ]

        validate_holdings(
            holdings
        )

        result["status"] = "success"
        result["reason"] = None
        result[
            "pdfExtractionMode"
        ] = selected_mode
        result["holdings"] = holdings

        result["diagnostics"] = {
            **selected.get(
                "diagnostics",
                {},
            ),
            "parserErrors": (
                parser_errors
            ),
            "availableExtractionModes": [
                item[0]
                for item in extraction_candidates
            ],
        }

        selected_section = next(
            (
                section
                for (
                    mode,
                    section,
                    _count,
                ) in extraction_candidates
                if mode == selected_mode
            ),
            "",
        )

        write_text(
            fund_dir
            / "top_holdings_section.txt",
            selected_section,
        )

        write_json(
            fund_dir
            / "holdings.json",
            result,
        )

        print(
            f"  SUCCESS: "
            f"{len(holdings)} holdings "
            f"using {selected_mode} extraction."
        )

        for holding in holdings:
            weight = holding.get(
                "weightText"
            )

            if weight:
                print(
                    f"    {holding['rank']}. "
                    f"{holding['name']} "
                    f"({weight})"
                )
            else:
                print(
                    f"    {holding['rank']}. "
                    f"{holding['name']} "
                    f"(percentage not published)"
                )

        return result

    except Exception as exc:
        result["status"] = "failed"

        result["reason"] = (
            f"{type(exc).__name__}: {exc}"
        )

        print(
            f"  FAILED: "
            f"{result['reason']}"
        )

        try:
            write_json(
                fund_dir
                / "holdings.json",
                result,
            )
        except Exception:
            pass

        return result

    finally:
        result[
            "completedAtUtc"
        ] = utc_now_iso()

        try:
            write_json(
                fund_dir
                / "holdings.json",
                result,
            )
        except Exception:
            pass


# ============================================================================
# RUN SUMMARY
# ============================================================================

def build_run_summary(
    started_at: str,
    completed_at: str,
    funds: list[
        dict[str, Any]
    ],
    results: list[
        dict[str, Any]
    ],
) -> dict[str, Any]:

    successful = [
        result
        for result in results
        if result.get("status")
        == "success"
    ]

    no_holdings_section = [
        result
        for result in results
        if result.get("status")
        == "no_holdings_section"
    ]

    failed = [
        result
        for result in results
        if result.get("status")
        == "failed"
    ]

    total_holdings = sum(
        len(
            result.get(
                "holdings",
                [],
            )
        )
        for result in successful
    )

    total_published_percentages = sum(
        1
        for result in successful
        for holding in result.get(
            "holdings",
            [],
        )
        if holding.get(
            "weightPercent"
        )
        is not None
    )

    total_missing_percentages = sum(
        1
        for result in successful
        for holding in result.get(
            "holdings",
            [],
        )
        if holding.get(
            "weightPercent"
        )
        is None
    )

    return {
        "status": (
            "success"
            if not failed
            else "failed"
        ),
        "startedAtUtc": started_at,
        "completedAtUtc": completed_at,
        "excelFile": str(
            EXCEL_FILE
        ),
        "fundUniverseCount": len(
            funds
        ),
        "successfulFundCount": len(
            successful
        ),
        "noHoldingsSectionFundCount": len(
            no_holdings_section
        ),
        "failedFundCount": len(
            failed
        ),
        "totalHoldings": (
            total_holdings
        ),
        "totalPublishedPercentages": (
            total_published_percentages
        ),
        "totalMissingPublishedPercentages": (
            total_missing_percentages
        ),
        "failedFunds": [
            {
                "excelRow": result.get(
                    "excelRow"
                ),
                "url": result.get(
                    "url"
                ),
                "pruAccessFundName": result.get(
                    "pruAccessFundName"
                ),
                "reason": result.get(
                    "reason"
                ),
            }
            for result in failed
        ],
        "noHoldingsSectionFunds": [
            {
                "excelRow": result.get(
                    "excelRow"
                ),
                "url": result.get(
                    "url"
                ),
                "pruAccessFundName": result.get(
                    "pruAccessFundName"
                ),
                "reason": result.get(
                    "reason"
                ),
            }
            for result in no_holdings_section
        ],
        "rules": {
            "excelColumnAControlsUniverse": True,
            "officialPrudentialSourcesOnly": True,
            "thirdPartySourcesAllowed": False,
            "maximumHoldings": MAX_HOLDINGS,
            "publishedHoldingCountIsAuthoritative": True,
            "holdingRankDefinesRowBoundary": True,
            "publishedPercentageOptional": True,
            "missingPublishedPercentageStoredAsNull": True,
            "percentageCalculationAllowed": False,
            "percentageInferenceAllowed": False,
            "holdingNameInferenceAllowed": False,
            "duplicatePercentagesAllowed": True,
            "duplicateHoldingNamesAllowed": False,
            "syntheticDataAllowed": False,
            "estimatedDataAllowed": False,
            "interpolationAllowed": False,
            "forcedTenHoldings": False,
        },
    }


# ============================================================================
# MAIN
# ============================================================================

def main() -> int:
    started_at = utc_now_iso()

    print("=" * 78)
    print(
        "PRUDENTIAL OFFICIAL TOP HOLDINGS EXTRACTOR"
    )
    print("=" * 78)

    print(
        f"Excel file: {EXCEL_FILE}"
    )

    print(
        f"Output directory: {OUTPUT_DIR}"
    )

    print()

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        funds = read_excel_funds()

    except Exception as exc:
        print(
            f"ERROR reading Excel: {exc}"
        )

        return 1

    if not funds:
        print(
            "ERROR: Excel Column A contains "
            "no populated fund URLs."
        )

        return 1

    print(
        f"Excel fund universe: "
        f"{len(funds)}"
    )

    print()

    session = create_session()

    results: list[
        dict[str, Any]
    ] = []

    for index, fund in enumerate(
        funds,
        start=1,
    ):
        result = process_fund(
            session=session,
            fund=fund,
            fund_index=index,
            total_funds=len(funds),
        )

        results.append(
            result
        )

        print()

        if index < len(funds):
            time.sleep(0.25)

    completed_at = utc_now_iso()

    all_holdings_payload = {
        "status": (
            "success"
            if all(
                result.get(
                    "status"
                )
                in (
                    "success",
                    "no_holdings_section",
                )
                for result in results
            )
            else "failed"
        ),
        "generatedAtUtc": (
            completed_at
        ),
        "fundUniverseCount": len(
            funds
        ),
        "funds": results,
    }

    write_json(
        OUTPUT_DIR
        / "all_holdings.json",
        all_holdings_payload,
    )

    summary = build_run_summary(
        started_at=started_at,
        completed_at=completed_at,
        funds=funds,
        results=results,
    )

    write_json(
        OUTPUT_DIR
        / "run_summary.json",
        summary,
    )

    print("=" * 78)
    print("RUN SUMMARY")
    print("=" * 78)

    print(
        f"Status: {summary['status']}"
    )

    print(
        "Excel fund universe: "
        f"{summary['fundUniverseCount']}"
    )

    print(
        "Successful: "
        f"{summary['successfulFundCount']}"
    )

    print(
        "No holdings section: "
        f"{summary['noHoldingsSectionFundCount']}"
    )

    print(
        "Failed: "
        f"{summary['failedFundCount']}"
    )

    print(
        "Total holdings: "
        f"{summary['totalHoldings']}"
    )

    print(
        "Published percentages: "
        f"{summary['totalPublishedPercentages']}"
    )

    print(
        "Missing published percentages: "
        f"{summary['totalMissingPublishedPercentages']}"
    )

    if summary["failedFunds"]:
        print()
        print("FAILED FUNDS:")

        for failed in summary[
            "failedFunds"
        ]:
            print(
                f"  Row {failed['excelRow']}: "
                f"{failed['pruAccessFundName'] or failed['url']}"
            )

            print(
                f"    {failed['reason']}"
            )

    print()

    print(
        "Combined output: "
        f"{OUTPUT_DIR / 'all_holdings.json'}"
    )

    print(
        "Run summary: "
        f"{OUTPUT_DIR / 'run_summary.json'}"
    )

    print("=" * 78)

    if summary[
        "failedFundCount"
    ] > 0:
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )
