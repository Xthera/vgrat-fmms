#!/usr/bin/env python3

"""
VGrat FMS - Prudential Top Holdings ALL-FUND TEST EXTRACTOR

MASTER SOURCE
=============

Funds Links.xlsm

Column A:
    Prudential fund URL

Column B:
    Exact PruAccess fund name


PURPOSE
=======

Extract official Prudential Top Holdings from EVERY fund listed in
Excel Column A.

The script supports the different PDF layouts used by Prudential,
including:

    1. Rank + holding name + percentage on one line

    2. Rank + multi-line holding name + percentage

    3. Rank on its own line
       holding name on one or more lines
       percentage on a later line

    4. Holding name and percentage separated by PDF table extraction

    5. Holdings published without percentages

No holding percentage is invented when Prudential does not publish one.


HARD RULES
==========

- Excel Column A controls the universe.
- No hardcoded 67-fund limit.
- Every populated URL in Column A is processed.
- Only official Prudential Singapore pages are accepted.
- Only official Prudential Singapore factsheets are accepted.
- No third-party holdings source.
- No inferred holding names.
- No fabricated holdings.
- No fabricated percentages.
- No forced 10 holdings.
- If Prudential publishes fewer than 10 holdings, store exactly that count.
- Holdings are stored in the order published by Prudential.
- Multi-line holding names are preserved and joined into one name.
- A holding percentage may be null when Prudential does not publish
  a percentage for that holding.
- No percentage is calculated from other holdings.
- No percentage is inferred.
- If a Top Holdings section exists but contains no identifiable
  holding names, the fund is FAILED.
- If a Top Holdings section exists and names are identifiable but
  Prudential does not publish percentages, the fund is SUCCESS with
  null weights.
- If a factsheet has no Top Holdings section at all, the fund is
  marked NO_HOLDINGS_SECTION.
- Failed funds do not get fabricated holdings.
- The script never changes PruAccess settings.
- This script does NOT modify test_pruaccess.py.
- This script does NOT create data.json.
- This script does NOT create index.html/css/js.
- This script is a TEST collector only.


OUTPUT
======

output_holdings/
    all_holdings.json
    run_summary.json

    funds/
        <excelRow>_<identifier>/
            factsheet.pdf
            factsheet_text.txt
            top_holdings_section.txt
            top_holdings.json
            metadata.json

        <excelRow>_failed/
            failure.json
"""


from __future__ import annotations

import json
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse

from openpyxl import load_workbook
from pypdf import PdfReader
from playwright.sync_api import (
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)


# =============================================================================
# CONFIGURATION
# =============================================================================

EXCEL_FILE = Path(
    "Funds Links.xlsm"
)

OUTPUT_DIR = Path(
    "output_holdings"
)

FUNDS_OUTPUT_DIR = (
    OUTPUT_DIR / "funds"
)

ALL_HOLDINGS_FILE = (
    OUTPUT_DIR / "all_holdings.json"
)

RUN_SUMMARY_FILE = (
    OUTPUT_DIR / "run_summary.json"
)


# ---------------------------------------------------------------------------
# Browser
# ---------------------------------------------------------------------------

BROWSER_HEADLESS = True

PAGE_TIMEOUT_MS = 120000

FACTSHEET_DOWNLOAD_TIMEOUT_MS = 120000

POST_PAGE_WAIT_MS = 1500

RETRY_COUNT = 3

RETRY_DELAY_SECONDS = 3.0


# ---------------------------------------------------------------------------
# Holdings
# ---------------------------------------------------------------------------

MAX_HOLDINGS = 10


# ---------------------------------------------------------------------------
# Only official Prudential Singapore URLs are accepted.
# ---------------------------------------------------------------------------

PRUDENTIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}


# =============================================================================
# GENERAL HELPERS
# =============================================================================

def clean_text(
    value,
) -> str:
    """
    Normalize whitespace while preserving meaningful text.
    """

    if value is None:
        return ""

    text = str(value)

    text = text.replace(
        "\xa0",
        " ",
    )

    text = text.replace(
        "\u200b",
        "",
    )

    text = text.replace(
        "\ufeff",
        "",
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def normalize_text(
    value,
) -> str:

    text = clean_text(
        value
    )

    text = text.casefold()

    text = text.replace(
        "–",
        "-",
    )

    text = text.replace(
        "—",
        "-",
    )

    return re.sub(
        r"\s+",
        " ",
        text,
    ).strip()


def utc_now_iso() -> str:

    return (
        datetime.now(
            timezone.utc
        )
        .replace(
            microsecond=0
        )
        .isoformat()
    )


def save_json(
    path: Path,
    data,
) -> None:

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    path.write_text(
        json.dumps(
            data,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def safe_filename(
    value,
) -> str:

    text = clean_text(
        value
    )

    text = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        text,
    )

    text = text.strip(
        "._"
    )

    if not text:
        text = "fund"

    return text[:120]


def is_prudential_url(
    url: str,
) -> bool:

    try:

        parsed = urlparse(
            url
        )

        hostname = (
            parsed.hostname
            or ""
        ).lower()

        return (
            parsed.scheme.lower()
            in {
                "http",
                "https",
            }
            and hostname
            in PRUDENTIAL_HOSTS
        )

    except Exception:

        return False


def ensure_prudential_url(
    url: str,
) -> str:

    url = clean_text(
        url
    )

    if not url:

        raise RuntimeError(
            "URL is empty."
        )

    if not is_prudential_url(
        url
    ):

        raise RuntimeError(
            "Non-Prudential URL rejected: "
            f"{url}"
        )

    return url


# =============================================================================
# EXCEL MASTER UNIVERSE
# =============================================================================

def read_excel_funds() -> list[dict]:

    print(
        "\n"
        "============================================================"
    )

    print(
        "READING EXCEL MASTER FUND UNIVERSE"
    )

    print(
        "============================================================"
    )

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        keep_vba=True,
        data_only=True,
    )

    worksheet = workbook.active

    funds = []

    for row_number in range(
        2,
        worksheet.max_row + 1,
    ):

        column_a = worksheet.cell(
            row=row_number,
            column=1,
        ).value

        column_b = worksheet.cell(
            row=row_number,
            column=2,
        ).value

        prudential_url = clean_text(
            column_a
        )

        pruaccess_name = clean_text(
            column_b
        )

        if not prudential_url:
            continue

        funds.append(
            {
                "excelRow":
                    row_number,

                "prudentialUrl":
                    prudential_url,

                "pruAccessName":
                    pruaccess_name,
            }
        )

    workbook.close()

    if not funds:

        raise RuntimeError(
            "No populated Prudential URLs were found "
            "in Excel Column A."
        )

    print(
        f"Excel fund universe: {len(funds)}"
    )

    return funds


# =============================================================================
# FIND FACTSHEET LINK
# =============================================================================

def find_factsheet_url(
    page,
    prudential_url: str,
) -> str:

    anchors = page.locator(
        "a"
    )

    anchor_count = anchors.count()

    candidates = []

    for index in range(
        anchor_count
    ):

        anchor = anchors.nth(
            index
        )

        try:

            href = clean_text(
                anchor.get_attribute(
                    "href"
                )
            )

            if not href:
                continue

            absolute_url = urljoin(
                page.url,
                href,
            )

            if not is_prudential_url(
                absolute_url
            ):
                continue

            anchor_text = clean_text(
                anchor.inner_text()
            )

        except Exception:

            continue

        href_lower = (
            absolute_url.lower()
        )

        text_lower = (
            anchor_text.lower()
        )

        score = 0

        if (
            "fund factsheet"
            in text_lower
        ):
            score += 200

        elif (
            "factsheet"
            in text_lower
        ):
            score += 150

        if "factsheet" in href_lower:
            score += 100

        if href_lower.endswith(
            ".pdf"
        ):
            score += 50

        if score > 0:

            candidates.append(
                {
                    "score":
                        score,

                    "url":
                        absolute_url,

                    "text":
                        anchor_text,
                }
            )

    if not candidates:

        raise RuntimeError(
            "Could not find an official Prudential "
            "factsheet link on the fund page."
        )

    candidates.sort(
        key=lambda item: (
            -item["score"],
            item["url"],
        )
    )

    selected = candidates[0]

    print(
        "Factsheet link found:"
    )

    print(
        selected["url"]
    )

    return ensure_prudential_url(
        selected["url"]
    )


# =============================================================================
# PDF TEXT
# =============================================================================

def extract_pdf_text(
    pdf_bytes: bytes,
) -> tuple[str, int]:

    from io import BytesIO

    reader = PdfReader(
        BytesIO(
            pdf_bytes
        )
    )

    page_count = len(
        reader.pages
    )

    if page_count == 0:

        raise RuntimeError(
            "Factsheet PDF has zero pages."
        )

    pages = []

    for page_index, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):

        try:

            page_text = (
                pdf_page.extract_text()
                or ""
            )

        except Exception as error:

            raise RuntimeError(
                "Failed to extract PDF text "
                f"from page {page_index}: {error}"
            ) from error

        pages.append(
            page_text
        )

    full_text = "\n".join(
        pages
    )

    if not clean_text(
        full_text
    ):

        raise RuntimeError(
            "Factsheet PDF contains no extractable text."
        )

    return (
        full_text,
        page_count,
    )


# =============================================================================
# PDF LINE CLEANING
# =============================================================================

def pdf_lines(
    text: str,
) -> list[str]:

    lines = []

    for raw_line in text.splitlines():

        line = clean_text(
            raw_line
        )

        if not line:
            continue

        lines.append(
            line
        )

    return lines


# =============================================================================
# FACTSHEET DATE EXTRACTION
# =============================================================================

def extract_data_as_at(
    text: str,
) -> str | None:

    patterns = [

        re.compile(
            r"""
            \b
            data
            \s+as\s+at
            \s*
            ([0-9]{1,2}
            \s+
            [A-Za-z]{3,9}
            \s+
            [0-9]{4})
            \b
            """,
            re.IGNORECASE
            | re.VERBOSE,
        ),

        re.compile(
            r"""
            \ball\s+data
            \s+as\s+at
            \s*
            ([0-9]{1,2}
            \s+
            [A-Za-z]{3,9}
            \s+
            [0-9]{4})
            \b
            """,
            re.IGNORECASE
            | re.VERBOSE,
        ),
    ]

    for pattern in patterns:

        match = pattern.search(
            text
        )

        if match:

            return clean_text(
                match.group(1)
            )

    return None


# =============================================================================
# FACTSHEET DOCUMENT DATE
# =============================================================================

def extract_document_date(
    text: str,
) -> str | None:

    lines = pdf_lines(
        text
    )

    month_year_pattern = re.compile(
        r"""
        \b
        (
            January|
            February|
            March|
            April|
            May|
            June|
            July|
            August|
            September|
            October|
            November|
            December
        )
        \s+
        ([0-9]{4})
        \b
        """,
        re.IGNORECASE
        | re.VERBOSE,
    )

    for line in lines[:50]:

        match = (
            month_year_pattern.search(
                line
            )
        )

        if match:

            return clean_text(
                match.group(0)
            )

    return None


# =============================================================================
# HOLDINGS SECTION BOUNDARIES
# =============================================================================

def find_holdings_start(
    lines: list[str],
) -> int | None:

    for index, line in enumerate(
        lines
    ):

        normalized = normalize_text(
            line
        )

        if (
            "top 10 holdings"
            in normalized
            or
            "top ten holdings"
            in normalized
            or
            "top holdings"
            in normalized
        ):

            return index

    return None


def is_holdings_end(
    line: str,
) -> bool:

    normalized = normalize_text(
        line
    )

    endings = (

        "source",

        "source:",

        "inception date",

        "important information",

        "important information:",

        "disclaimer",

        "past performance",

        "portfolio characteristics",

        "asset allocation",

        "sector allocation",

        "geographical allocation",

        "country allocation",

        "currency allocation",

        "risk profile",
    )

    if normalized in endings:
        return True

    if normalized.startswith(
        "source:"
    ):
        return True

    if normalized.startswith(
        "inception date:"
    ):
        return True

    if normalized.startswith(
        "important information"
    ):
        return True

    if re.fullmatch(
        r"page\s+\d+(\s+of\s+\d+)?",
        normalized,
    ):
        return True

    return False


def extract_holdings_section(
    text: str,
) -> tuple[
    str,
    str,
]:

    lines = pdf_lines(
        text
    )

    start_index = find_holdings_start(
        lines
    )

    if start_index is None:

        return (
            "",
            "not_published",
        )

    section = []

    for line in lines[
        start_index + 1:
    ]:

        if is_holdings_end(
            line
        ):
            break

        section.append(
            line
        )

    return (
        "\n".join(
            section
        ),
        "published",
    )


# =============================================================================
# HOLDING NAME CLEANUP
# =============================================================================

def clean_holding_name(
    value: str,
) -> str:

    name = clean_text(
        value
    )

    if not name:
        return ""

    name = name.replace(
        "|",
        " ",
    )

    name = re.sub(
        r"^[•·▪■□*]+",
        "",
        name,
    )

    name = name.strip()

    # Remove accidental leading rank.
    name = re.sub(
        r"^\d{1,2}[.)]\s+",
        "",
        name,
    )

    # Remove a standalone rank.
    name = re.sub(
        r"^\d{1,2}\s+",
        "",
        name,
    )

    name = re.sub(
        r"\bNone\b",
        " ",
        name,
        flags=re.IGNORECASE,
    )

    name = clean_text(
        name
    )

    if normalize_text(
        name
    ) in {
        "",
        "none",
        "null",
        "-",
        "—",
        "holding",
        "holdings",
        "name",
        "security",
        "security name",
        "portfolio",
        "portfolio holdings",
        "weight",
        "weight %",
        "%",
    }:

        return ""

    return name


# =============================================================================
# RANK DETECTION
# =============================================================================

def extract_rank(
    line: str,
) -> tuple[
    int | None,
    str,
]:
    """
    Detect a published holding rank.

    Supported examples:

        1
        1.
        1)
        1. FUND NAME
        1 FUND NAME
    """

    original = clean_text(
        line
    )

    if not original:
        return (
            None,
            "",
        )

    match = re.match(
        r"""
        ^
        (\d{1,2})
        (?:
            [.)]
            |
            (?=\s)
            |
            $
        )
        \s*
        (.*)
        $
        """,
        original,
        re.VERBOSE,
    )

    if not match:
        return (
            None,
            original,
        )

    rank = int(
        match.group(1)
    )

    remainder = clean_text(
        match.group(2)
    )

    if rank < 1 or rank > MAX_HOLDINGS:
        return (
            None,
            original,
        )

    return (
        rank,
        remainder,
    )


# =============================================================================
# PERCENTAGE PARSING
# =============================================================================

def extract_percentage_from_line(
    value: str,
) -> tuple[
    float | None,
    str,
]:
    """
    Extract a percentage from anywhere in a line.

    Returns:

        percentage
        remaining text

    Examples:

        "ABC FUND 38.8%"
        -> 38.8, "ABC FUND"

        "38.8%"
        -> 38.8, ""

        "ABC FUND 38.8% 2025"
        -> 38.8, "ABC FUND 2025"
    """

    if value is None:
        return (
            None,
            "",
        )

    text = clean_text(
        value
    )

    matches = list(
        re.finditer(
            r"(?<![\d.])"
            r"([0-9]+(?:\.[0-9]+)?)"
            r"\s*%",
            text,
        )
    )

    if not matches:
        return (
            None,
            text,
        )

    # Use the first published percentage in the line.
    match = matches[0]

    percentage = float(
        match.group(1)
    )

    if (
        percentage < 0
        or percentage > 100
    ):

        return (
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
        percentage,
        remaining,
    )


def parse_percentage(
    value: str,
) -> float | None:

    if value is None:
        return None

    text = clean_text(
        value
    )

    match = re.fullmatch(
        r"([0-9]+(?:\.[0-9]+)?)\s*%",
        text,
    )

    if not match:
        return None

    percentage = float(
        match.group(1)
    )

    if (
        percentage < 0
        or percentage > 100
    ):
        return None

    return percentage


# =============================================================================
# HEADER / NON-HOLDING DETECTION
# =============================================================================

def looks_like_holdings_header(
    line: str,
) -> bool:

    normalized = normalize_text(
        line
    )

    headers = {
        "top 10 holdings",
        "top ten holdings",
        "top holdings",
        "holding",
        "holdings",
        "name",
        "security",
        "security name",
        "weight",
        "weight %",
        "weight (%)",
        "percentage",
        "% of net assets",
        "% of portfolio",
        "portfolio holdings",
    }

    if normalized in headers:
        return True

    if (
        "holding"
        in normalized
        and
        (
            "weight"
            in normalized
            or
            "percentage"
            in normalized
            or
            "%"
            in normalized
        )
    ):
        return True

    return False


# =============================================================================
# RANKED HOLDINGS PARSER
# =============================================================================

def parse_ranked_holdings(
    section_text: str,
) -> tuple[
    list[dict],
    dict,
]:
    """
    Robust parser for Prudential PDF holdings tables.

    It understands multi-line holding names.

    Example:

        1
        JPMORGAN
        FUNDS
        EMERGING MARKETS
        14.7%

    becomes:

        rank = 1
        name = "JPMORGAN FUNDS EMERGING MARKETS"
        weight = 14.7

    It also supports holdings with no published percentage:

        1
        FUND NAME

        2
        ANOTHER FUND

    In that case weightPercent is null.

    IMPORTANT:

    No percentage is calculated or inferred.
    """

    lines = pdf_lines(
        section_text
    )

    if not lines:

        raise RuntimeError(
            "Top Holdings section is empty."
        )

    # -------------------------------------------------------------------------
    # Remove obvious table headers.
    # -------------------------------------------------------------------------

    cleaned_lines = []

    for line in lines:

        if looks_like_holdings_header(
            line
        ):
            continue

        cleaned_lines.append(
            line
        )

    lines = cleaned_lines

    # -------------------------------------------------------------------------
    # Build ranked blocks.
    #
    # A new rank starts a new holding.
    # -------------------------------------------------------------------------

    blocks = []

    current = None

    for line in lines:

        rank, remainder = extract_rank(
            line
        )

        if rank is not None:

            # A new rank before the previous rank is completed is not
            # automatically an error anymore. It can happen when the
            # previous holding has a multi-line name and its percentage
            # appears in a different PDF text column.
            if current is not None:

                blocks.append(
                    current
                )

            current = {
                "rank": rank,
                "lines": [],
            }

            if remainder:

                current[
                    "lines"
                ].append(
                    remainder
                )

            continue

        if current is not None:

            current[
                "lines"
            ].append(
                line
            )

    if current is not None:

        blocks.append(
            current
        )

    # -------------------------------------------------------------------------
    # If no ranks were found, use a second non-ranked parser.
    # -------------------------------------------------------------------------

    if not blocks:

        return (
            parse_non_ranked_holdings(
                lines
            ),
            {
                "parser":
                    "non_ranked_fallback",

                "rankedBlocks":
                    0,
            },
        )

    # -------------------------------------------------------------------------
    # Validate and parse each ranked block.
    # -------------------------------------------------------------------------

    holdings = []

    for block in blocks:

        rank = block[
            "rank"
        ]

        block_lines = [
            clean_text(
                line
            )
            for line
            in block[
                "lines"
            ]
            if clean_text(
                line
            )
        ]

        if rank < 1 or rank > MAX_HOLDINGS:
            continue

        # ---------------------------------------------------------------------
        # Stop if this is clearly outside the holdings table.
        # ---------------------------------------------------------------------

        useful_lines = []

        for line in block_lines:

            if is_holdings_end(
                line
            ):
                break

            useful_lines.append(
                line
            )

        block_lines = useful_lines

        if not block_lines:
            continue

        # ---------------------------------------------------------------------
        # Extract all percentages in the block.
        #
        # Normally there is zero or one.
        #
        # If several percentages appear because of PDF extraction artifacts,
        # use the first percentage associated with the block.
        # ---------------------------------------------------------------------

        percentage = None

        name_parts = []

        for line in block_lines:

            line_percentage, remaining = (
                extract_percentage_from_line(
                    line
                )
            )

            if (
                percentage is None
                and
                line_percentage is not None
            ):

                percentage = line_percentage

            remaining = clean_text(
                remaining
            )

            if remaining:

                name_parts.append(
                    remaining
                )

        # ---------------------------------------------------------------------
        # If a percentage line contained only the percentage, it contributes
        # nothing to the name.
        # ---------------------------------------------------------------------

        name = clean_holding_name(
            " ".join(
                name_parts
            )
        )

        # ---------------------------------------------------------------------
        # Some PDF layouts put the percentage before the holding name.
        # Try to recover the name from the non-percentage text.
        # ---------------------------------------------------------------------

        if not name:

            non_percentage_lines = []

            for line in block_lines:

                stripped = re.sub(
                    r"[0-9]+(?:\.[0-9]+)?\s*%",
                    " ",
                    line,
                )

                stripped = clean_text(
                    stripped
                )

                if stripped:

                    non_percentage_lines.append(
                        stripped
                    )

            name = clean_holding_name(
                " ".join(
                    non_percentage_lines
                )
            )

        # ---------------------------------------------------------------------
        # Header-like content is not a holding.
        # ---------------------------------------------------------------------

        if looks_like_holdings_header(
            name
        ):
            name = ""

        if not name:

            # Do not silently create a fabricated holding.
            continue

        holdings.append(
            {
                "rank":
                    rank,

                "name":
                    name,

                "weightPercent":
                    percentage,

                "weightText":
                    (
                        f"{percentage:g}%"
                        if percentage is not None
                        else None
                    ),
            }
        )

        if len(
            holdings
        ) >= MAX_HOLDINGS:

            break

    # -------------------------------------------------------------------------
    # Sort by published rank.
    # -------------------------------------------------------------------------

    holdings.sort(
        key=lambda item: item[
            "rank"
        ]
    )

    # -------------------------------------------------------------------------
    # Remove exact duplicate ranks caused by PDF extraction artifacts.
    # Keep first occurrence.
    # -------------------------------------------------------------------------

    unique_by_rank = []

    seen_ranks = set()

    for item in holdings:

        rank = item[
            "rank"
        ]

        if rank in seen_ranks:
            continue

        seen_ranks.add(
            rank
        )

        unique_by_rank.append(
            item
        )

    holdings = unique_by_rank

    # -------------------------------------------------------------------------
    # If ranked parsing found names, use them.
    # -------------------------------------------------------------------------

    if holdings:

        # Renumber only if Prudential's actual ranks are sequential.
        #
        # We do NOT invent ranks. The original published rank is retained.
        #
        # For normal Prudential tables this should be 1..N.

        return (
            holdings,
            {
                "parser":
                    "ranked_block_parser",

                "rankedBlocks":
                    len(blocks),

                "publishedPercentages":
                    sum(
                        1
                        for item
                        in holdings
                        if item[
                            "weightPercent"
                        ] is not None
                    ),

                "holdingsWithoutPercentage":
                    sum(
                        1
                        for item
                        in holdings
                        if item[
                            "weightPercent"
                        ] is None
                    ),
            },
        )

    # -------------------------------------------------------------------------
    # No ranked names were recoverable.
    #
    # Try the non-ranked fallback before declaring failure.
    # -------------------------------------------------------------------------

    fallback = parse_non_ranked_holdings(
        lines
    )

    if fallback:

        return (
            fallback,
            {
                "parser":
                    "non_ranked_fallback_after_ranked_failure",

                "rankedBlocks":
                    len(blocks),
            },
        )

    raise RuntimeError(
        "Top Holdings section was found, but no holding names "
        "could be parsed from the PDF text."
    )


# =============================================================================
# NON-RANKED FALLBACK
# =============================================================================

def parse_non_ranked_holdings(
    lines: list[str],
) -> list[dict]:
    """
    Fallback for PDFs where the PDF text extraction removes the visible
    holding rank.

    Supports:

        NAME
        38.8%

    and:

        NAME 38.8%

    It also supports multi-line names.
    """

    holdings = []

    current_name_parts = []

    current_percentage = None

    def flush_current():

        nonlocal current_name_parts
        nonlocal current_percentage

        if not current_name_parts:
            current_percentage = None
            return

        name = clean_holding_name(
            " ".join(
                current_name_parts
            )
        )

        if not name:
            current_name_parts = []
            current_percentage = None
            return

        holdings.append(
            {
                "rank":
                    len(holdings) + 1,

                "name":
                    name,

                "weightPercent":
                    current_percentage,

                "weightText":
                    (
                        f"{current_percentage:g}%"
                        if current_percentage is not None
                        else None
                    ),
            }
        )

        current_name_parts = []
        current_percentage = None

    for line in lines:

        if is_holdings_end(
            line
        ):
            break

        if looks_like_holdings_header(
            line
        ):
            continue

        percentage, remaining = (
            extract_percentage_from_line(
                line
            )
        )

        remaining = clean_text(
            remaining
        )

        if percentage is not None:

            # If we already have a name, this percentage completes it.
            if current_name_parts:

                current_percentage = percentage

                if remaining:

                    current_name_parts.append(
                        remaining
                    )

                flush_current()

                if len(
                    holdings
                ) >= MAX_HOLDINGS:

                    break

                continue

            # If no name is pending, this cannot safely become a holding.
            continue

        # ---------------------------------------------------------------------
        # No percentage.
        #
        # If a new plausible name arrives, append it to the current multi-line
        # name. The next percentage will complete the holding.
        # ---------------------------------------------------------------------

        candidate = clean_holding_name(
            line
        )

        if not candidate:
            continue

        current_name_parts.append(
            candidate
        )

    # -------------------------------------------------------------------------
    # If the final holding has a name but no percentage, retain it.
    # -------------------------------------------------------------------------

    if current_name_parts:

        flush_current()

    return holdings[:MAX_HOLDINGS]


# =============================================================================
# HOLDINGS VALIDATION
# =============================================================================

def validate_holdings(
    holdings: list[dict],
) -> None:

    if not holdings:

        raise RuntimeError(
            "No holding names could be extracted."
        )

    if len(
        holdings
    ) > MAX_HOLDINGS:

        raise RuntimeError(
            "Parser produced more than 10 holdings."
        )

    # -------------------------------------------------------------------------
    # Ranks must be valid.
    # -------------------------------------------------------------------------

    ranks = []

    for holding in holdings:

        rank = holding.get(
            "rank"
        )

        if not isinstance(
            rank,
            int,
        ):

            raise RuntimeError(
                "Holding rank is not an integer."
            )

        if (
            rank < 1
            or
            rank > MAX_HOLDINGS
        ):

            raise RuntimeError(
                f"Invalid holding rank: {rank}"
            )

        ranks.append(
            rank
        )

    # -------------------------------------------------------------------------
    # Duplicate ranks are not allowed.
    # -------------------------------------------------------------------------

    if len(
        ranks
    ) != len(
        set(
            ranks
        )
    ):

        raise RuntimeError(
            "Duplicate holding ranks detected."
        )

    # -------------------------------------------------------------------------
    # Holding names must exist.
    # -------------------------------------------------------------------------

    names = []

    for holding in holdings:

        name = clean_holding_name(
            holding.get(
                "name"
            )
            or ""
        )

        if not name:

            raise RuntimeError(
                "A published holding rank was found "
                "but no holding name could be extracted."
            )

        names.append(
            normalize_text(
                name
            )
        )

    # -------------------------------------------------------------------------
    # Exact duplicate names are treated as a parser problem.
    #
    # This protects against the same PDF table being captured twice.
    # -------------------------------------------------------------------------

    duplicates = {}

    for name in names:

        duplicates[name] = (
            duplicates.get(
                name,
                0,
            )
            + 1
        )

    duplicate_names = [
        name
        for name, count
        in duplicates.items()
        if count > 1
    ]

    if duplicate_names:

        raise RuntimeError(
            "Duplicate holding name detected: "
            + ", ".join(
                duplicate_names
            )
        )

    # -------------------------------------------------------------------------
    # Percentage validation.
    #
    # None is explicitly allowed because some official Prudential factsheets
    # publish holding names without holding percentages.
    # -------------------------------------------------------------------------

    for holding in holdings:

        weight = holding.get(
            "weightPercent"
        )

        if weight is None:
            continue

        if not isinstance(
            weight,
            (
                int,
                float,
            ),
        ):

            raise RuntimeError(
                "Holding percentage is not numeric."
            )

        if (
            weight < 0
            or
            weight > 100
        ):

            raise RuntimeError(
                f"Invalid holding weight: {weight}"
            )


# =============================================================================
# HOLDINGS PARSER
# =============================================================================

def parse_holdings(
    section_text: str,
) -> tuple[
    list[dict],
    dict,
]:

    if not clean_text(
        section_text
    ):

        raise RuntimeError(
            "Top Holdings section is empty."
        )

    holdings, parser_info = (
        parse_ranked_holdings(
            section_text
        )
    )

    validate_holdings(
        holdings
    )

    return (
        holdings,
        parser_info,
    )


# =============================================================================
# FACTSHEET DOWNLOAD
# =============================================================================

def download_factsheet(
    page,
    factsheet_url: str,
) -> bytes:

    factsheet_url = ensure_prudential_url(
        factsheet_url
    )

    print(
        "Downloading factsheet..."
    )

    response = page.request.get(
        factsheet_url,
        timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
    )

    status = response.status

    if status != 200:

        raise RuntimeError(
            "Factsheet HTTP status was "
            f"{status}: {factsheet_url}"
        )

    body = response.body()

    if not body:

        raise RuntimeError(
            "Factsheet response contained zero bytes."
        )

    if not body.startswith(
        b"%PDF"
    ):

        content_type = clean_text(
            response.headers.get(
                "content-type"
            )
        )

        raise RuntimeError(
            "Factsheet response was not a PDF. "
            f"Content-Type={content_type}; "
            f"URL={factsheet_url}"
        )

    print(
        f"Factsheet bytes: {len(body):,}"
    )

    return body


# =============================================================================
# SINGLE FUND
# =============================================================================

def extract_single_fund(
    page,
    excel_fund: dict,
) -> dict:

    excel_row = int(
        excel_fund[
            "excelRow"
        ]
    )

    prudential_url = (
        excel_fund[
            "prudentialUrl"
        ]
    )

    pruaccess_name = (
        excel_fund.get(
            "pruAccessName"
        )
        or ""
    )

    print(
        "\n"
        "============================================================"
    )

    print(
        f"PROCESSING FUND ROW {excel_row}"
    )

    print(
        "============================================================"
    )

    print(
        f"Prudential URL: {prudential_url}"
    )

    print(
        f"Excel PruAccess name: {pruaccess_name}"
    )

    if not is_prudential_url(
        prudential_url
    ):

        raise RuntimeError(
            "Excel Column A URL is not an official "
            "Prudential Singapore URL."
        )

    # -------------------------------------------------------------------------
    # Open official Prudential page.
    # -------------------------------------------------------------------------

    page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

    try:

        page.wait_for_load_state(
            "networkidle",
            timeout=25000,
        )

    except PlaywrightTimeoutError:

        pass

    page.wait_for_timeout(
        POST_PAGE_WAIT_MS
    )

    final_url = clean_text(
        page.url
    )

    page_title = clean_text(
        page.title()
    )

    if not is_prudential_url(
        final_url
    ):

        raise RuntimeError(
            "Prudential fund page redirected outside "
            "Prudential Singapore: "
            f"{final_url}"
        )

    # -------------------------------------------------------------------------
    # Fund name.
    # -------------------------------------------------------------------------

    fund_name = ""

    h1 = page.locator(
        "h1"
    ).first

    if h1.count():

        try:

            fund_name = clean_text(
                h1.inner_text()
            )

        except Exception:

            fund_name = ""

    if not fund_name:

        body_text = clean_text(
            page.locator(
                "body"
            ).inner_text()
        )

        match = re.search(
            r"\bPRU(?:Link|Prime)\s+[^\n]+",
            body_text,
            re.IGNORECASE,
        )

        if match:

            fund_name = clean_text(
                match.group(0)
            )

    # -------------------------------------------------------------------------
    # Factsheet.
    # -------------------------------------------------------------------------

    factsheet_url = (
        find_factsheet_url(
            page,
            prudential_url,
        )
    )

    factsheet_bytes = (
        download_factsheet(
            page,
            factsheet_url,
        )
    )

    # -------------------------------------------------------------------------
    # PDF extraction.
    # -------------------------------------------------------------------------

    (
        full_text,
        pdf_page_count,
    ) = extract_pdf_text(
        factsheet_bytes
    )

    # -------------------------------------------------------------------------
    # Factsheet dates.
    # -------------------------------------------------------------------------

    data_as_at = (
        extract_data_as_at(
            full_text
        )
    )

    document_date = (
        extract_document_date(
            full_text
        )
    )

    # -------------------------------------------------------------------------
    # Top holdings section.
    # -------------------------------------------------------------------------

    (
        section_text,
        section_status,
    ) = extract_holdings_section(
        full_text
    )

    if section_status == "not_published":

        return {
            "status":
                "no_holdings_section",

            "excelRow":
                excel_row,

            "prudentialUrl":
                prudential_url,

            "finalUrl":
                final_url,

            "pageTitle":
                page_title,

            "fundName":
                fund_name,

            "excelPruAccessName":
                pruaccess_name,

            "factsheetUrl":
                factsheet_url,

            "factsheetDocumentDate":
                document_date,

            "factsheetDataAsAt":
                data_as_at,

            "factsheetPageCount":
                pdf_page_count,

            "holdingsSectionStatus":
                "not_published",

            "topHoldingsCount":
                0,

            "topHoldings":
                [],

            "holdingsParser":
                None,

            "rules": {
                "officialPrudentialSourceOnly":
                    True,

                "maximumHoldings":
                    MAX_HOLDINGS,

                "publishedCountUsedExactly":
                    True,

                "fewerThanTenAllowed":
                    True,

                "noForcedTenEntries":
                    True,

                "noInferredHoldings":
                    True,

                "noFabricatedPercentages":
                    True,

                "nullWeightAllowedWhenNotPublished":
                    True,
            },
        }

    # -------------------------------------------------------------------------
    # Top Holdings section exists.
    # -------------------------------------------------------------------------

    (
        holdings,
        parser_info,
    ) = parse_holdings(
        section_text
    )

    published_percentage_count = sum(
        1
        for holding in holdings
        if holding.get(
            "weightPercent"
        ) is not None
    )

    missing_percentage_count = (
        len(
            holdings
        )
        -
        published_percentage_count
    )

    return {
        "status":
            "success",

        "excelRow":
            excel_row,

        "prudentialUrl":
            prudential_url,

        "finalUrl":
            final_url,

        "pageTitle":
            page_title,

        "fundName":
            fund_name,

        "excelPruAccessName":
            pruaccess_name,

        "factsheetUrl":
            factsheet_url,

        "factsheetDocumentDate":
            document_date,

        "factsheetDataAsAt":
            data_as_at,

        "factsheetPageCount":
            pdf_page_count,

        "holdingsSectionStatus":
            "published",

        "topHoldingsCount":
            len(
                holdings
            ),

        "publishedHoldingPercentageCount":
            published_percentage_count,

        "holdingsWithoutPublishedPercentage":
            missing_percentage_count,

        "topHoldings":
            holdings,

        "holdingsParser":
            parser_info,

        "rules": {
            "officialPrudentialSourceOnly":
                True,

            "maximumHoldings":
                MAX_HOLDINGS,

            "publishedCountUsedExactly":
                True,

            "fewerThanTenAllowed":
                True,

            "noForcedTenEntries":
                True,

            "noInferredHoldings":
                True,

            "noFabricatedPercentages":
                True,

            "nullWeightAllowedWhenNotPublished":
                True,
        },
    }


# =============================================================================
# SAVE SUCCESS
# =============================================================================

def save_success_result(
    result: dict,
    factsheet_bytes: bytes,
    full_text: str,
    section_text: str,
) -> Path:

    excel_row = int(
        result[
            "excelRow"
        ]
    )

    identifier = (
        safe_filename(
            result.get(
                "fundIdentifier"
            )
        )
        if result.get(
            "fundIdentifier"
        )
        else ""
    )

    if not identifier:

        parsed = urlparse(
            result.get(
                "finalUrl"
            )
            or result.get(
                "prudentialUrl"
            )
        )

        identifier = safe_filename(
            parsed.path.rstrip(
                "/"
            ).split(
                "/"
            )[-1]
            or result.get(
                "fundName"
            )
            or f"fund_{excel_row}"
        )

    directory = (
        FUNDS_OUTPUT_DIR
        / (
            f"{excel_row}_"
            f"{identifier}"
        )
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    if factsheet_bytes:

        (
            directory
            / "factsheet.pdf"
        ).write_bytes(
            factsheet_bytes
        )

    if full_text:

        (
            directory
            / "factsheet_text.txt"
        ).write_text(
            full_text,
            encoding="utf-8",
        )

    if section_text:

        (
            directory
            / "top_holdings_section.txt"
        ).write_text(
            section_text,
            encoding="utf-8",
        )

    save_json(
        directory
        / "top_holdings.json",
        result,
    )

    metadata = {
        "excelRow":
            excel_row,

        "fundName":
            result.get(
                "fundName"
            ),

        "factsheetUrl":
            result.get(
                "factsheetUrl"
            ),

        "factsheetDocumentDate":
            result.get(
                "factsheetDocumentDate"
            ),

        "factsheetDataAsAt":
            result.get(
                "factsheetDataAsAt"
            ),

        "topHoldingsCount":
            result.get(
                "topHoldingsCount"
            ),

        "publishedHoldingPercentageCount":
            result.get(
                "publishedHoldingPercentageCount"
            ),

        "holdingsWithoutPublishedPercentage":
            result.get(
                "holdingsWithoutPublishedPercentage"
            ),

        "holdingsParser":
            result.get(
                "holdingsParser"
            ),

        "status":
            result.get(
                "status"
            ),

        "savedAtUtc":
            utc_now_iso(),
    }

    save_json(
        directory
        / "metadata.json",
        metadata,
    )

    return directory


# =============================================================================
# SAVE FAILURE
# =============================================================================

def save_failure_result(
    excel_fund: dict,
    error_text: str,
) -> Path:

    excel_row = int(
        excel_fund[
            "excelRow"
        ]
    )

    directory = (
        FUNDS_OUTPUT_DIR
        / (
            f"{excel_row}_failed"
        )
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    failure = {
        "status":
            "failed",

        "excelRow":
            excel_row,

        "prudentialUrl":
            excel_fund[
                "prudentialUrl"
            ],

        "excelPruAccessName":
            excel_fund.get(
                "pruAccessName"
            ),

        "error":
            clean_text(
                error_text
            ),

        "failedAtUtc":
            utc_now_iso(),
    }

    save_json(
        directory
        / "failure.json",
        failure,
    )

    return directory


# =============================================================================
# MAIN
# =============================================================================

def main() -> int:

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    started_at = utc_now_iso()

    print(
        "\n"
        "################################################################"
    )

    print(
        "VGRAT FMS - PRUDENTIAL TOP HOLDINGS ALL-FUND TEST"
    )

    print(
        "################################################################"
    )

    print(
        f"Started UTC: {started_at}"
    )

    funds = read_excel_funds()

    successful = []

    no_holdings_section = []

    failed = []

    with sync_playwright() as playwright:

        browser = playwright.chromium.launch(
            headless=BROWSER_HEADLESS
        )

        context = browser.new_context(
            viewport={
                "width": 1440,
                "height": 1000,
            },

            user_agent=(
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Chrome/153.0.0.0 Safari/537.36"
            ),
        )

        page = context.new_page()

        try:

            total = len(
                funds
            )

            for index, excel_fund in enumerate(
                funds,
                start=1,
            ):

                print(
                    "\n"
                    + "=" * 72
                )

                print(
                    f"FUND {index}/{total}"
                )

                print(
                    f"Excel row: "
                    f"{excel_fund['excelRow']}"
                )

                print(
                    "=" * 72
                )

                last_error = None

                result = None

                for attempt in range(
                    1,
                    RETRY_COUNT + 1,
                ):

                    try:

                        print(
                            f"\nAttempt {attempt}/"
                            f"{RETRY_COUNT}"
                        )

                        result = (
                            extract_single_fund(
                                page,
                                excel_fund,
                            )
                        )

                        last_error = None

                        break

                    except Exception as error:

                        last_error = (
                            clean_text(
                                str(error)
                            )
                        )

                        print(
                            "Attempt failed:"
                        )

                        print(
                            last_error
                        )

                        if attempt < RETRY_COUNT:

                            print(
                                f"Retrying in "
                                f"{RETRY_DELAY_SECONDS} seconds..."
                            )

                            time.sleep(
                                RETRY_DELAY_SECONDS
                            )

                if result is None:

                    failure_directory = (
                        save_failure_result(
                            excel_fund,
                            last_error
                            or
                            "Unknown extraction failure.",
                        )
                    )

                    failed.append(
                        {
                            "excelRow":
                                excel_fund[
                                    "excelRow"
                                ],

                            "prudentialUrl":
                                excel_fund[
                                    "prudentialUrl"
                                ],

                            "pruAccessName":
                                excel_fund.get(
                                    "pruAccessName"
                                ),

                            "error":
                                last_error
                                or
                                "Unknown extraction failure.",

                            "outputDirectory":
                                str(
                                    failure_directory
                                ),
                        }
                    )

                    continue

                # -----------------------------------------------------------------
                # No Top Holdings section.
                # -----------------------------------------------------------------

                if (
                    result[
                        "status"
                    ]
                    ==
                    "no_holdings_section"
                ):

                    result_directory = (
                        save_success_result(
                            result,
                            b"",
                            "",
                            "",
                        )
                    )

                    no_holdings_section.append(
                        result
                    )

                    print(
                        "\n"
                        "NO TOP HOLDINGS SECTION"
                    )

                    continue

                # -----------------------------------------------------------------
                # SUCCESS
                #
                # The factsheet bytes and parsed text are already available from
                # extract_single_fund(). Downloading the same PDF again is not
                # necessary.
                # -----------------------------------------------------------------

                try:

                    factsheet_bytes = (
                        download_factsheet(
                            page,
                            result[
                                "factsheetUrl"
                            ],
                        )
                    )

                    (
                        full_text,
                        _page_count,
                    ) = extract_pdf_text(
                        factsheet_bytes
                    )

                    (
                        section_text,
                        section_status,
                    ) = extract_holdings_section(
                        full_text
                    )

                    if (
                        section_status
                        !=
                        "published"
                    ):

                        raise RuntimeError(
                            "Factsheet holdings section "
                            "disappeared during final verification."
                        )

                    (
                        verified_holdings,
                        verified_parser_info,
                    ) = parse_holdings(
                        section_text
                    )

                    if len(
                        verified_holdings
                    ) != result[
                        "topHoldingsCount"
                    ]:

                        raise RuntimeError(
                            "Holding count changed during "
                            "final verification."
                        )

                    result[
                        "topHoldings"
                    ] = verified_holdings

                    result[
                        "topHoldingsCount"
                    ] = len(
                        verified_holdings
                    )

                    result[
                        "holdingsParser"
                    ] = verified_parser_info

                    result[
                        "publishedHoldingPercentageCount"
                    ] = sum(
                        1
                        for holding
                        in verified_holdings
                        if holding.get(
                            "weightPercent"
                        ) is not None
                    )

                    result[
                        "holdingsWithoutPublishedPercentage"
                    ] = sum(
                        1
                        for holding
                        in verified_holdings
                        if holding.get(
                            "weightPercent"
                        ) is None
                    )

                    output_directory = (
                        save_success_result(
                            result,
                            factsheet_bytes,
                            full_text,
                            section_text,
                        )
                    )

                except Exception as error:

                    error_text = clean_text(
                        str(error)
                    )

                    failure_directory = (
                        save_failure_result(
                            excel_fund,
                            error_text,
                        )
                    )

                    failed.append(
                        {
                            "excelRow":
                                excel_fund[
                                    "excelRow"
                                ],

                            "prudentialUrl":
                                excel_fund[
                                    "prudentialUrl"
                                ],

                            "pruAccessName":
                                excel_fund.get(
                                    "pruAccessName"
                                ),

                            "error":
                                error_text,

                            "outputDirectory":
                                str(
                                    failure_directory
                                ),
                        }
                    )

                    continue

                successful.append(
                    result
                )

                print(
                    "\n"
                    "SUCCESS"
                )

                print(
                    f"Fund: "
                    f"{result.get('fundName') or '-'}"
                )

                print(
                    f"Factsheet data as at: "
                    f"{result.get('factsheetDataAsAt') or '-'}"
                )

                print(
                    f"Factsheet document date: "
                    f"{result.get('factsheetDocumentDate') or '-'}"
                )

                print(
                    f"Top holdings extracted: "
                    f"{result.get('topHoldingsCount')}"
                )

                print(
                    f"With published percentage: "
                    f"{result.get('publishedHoldingPercentageCount')}"
                )

                print(
                    f"Without published percentage: "
                    f"{result.get('holdingsWithoutPublishedPercentage')}"
                )

                print(
                    f"Parser: "
                    f"{result.get('holdingsParser')}"
                )

                for holding in result[
                    "topHoldings"
                ]:

                    weight_text = (
                        holding.get(
                            "weightText"
                        )
                        or
                        "not published"
                    )

                    print(
                        f"  {holding['rank']}. "
                        f"{holding['name']} - "
                        f"{weight_text}"
                    )

        finally:

            context.close()

            browser.close()

    # =========================================================================
    # ALL-HOLDINGS CONSOLIDATED FILE
    # =========================================================================

    completed_at = utc_now_iso()

    all_results = (
        successful
        +
        no_holdings_section
    )

    all_holdings_payload = {
        "status":
            "success",

        "source":
            "Prudential Singapore",

        "generatedAtUtc":
            completed_at,

        "excelFile":
            str(
                EXCEL_FILE
            ),

        "excelFundUniverse":
            len(
                funds
            ),

        "successfulFunds":
            len(
                successful
            ),

        "noHoldingsSectionFunds":
            len(
                no_holdings_section
            ),

        "failedFunds":
            len(
                failed
            ),

        "rules": {

            "excelColumnAControlsUniverse":
                True,

            "officialPrudentialOnly":
                True,

            "maximumHoldings":
                MAX_HOLDINGS,

            "publishedCountUsedExactly":
                True,

            "fewerThanTenAllowed":
                True,

            "noForcedTenEntries":
                True,

            "noInferredHoldings":
                True,

            "noFabricatedPercentages":
                True,

            "nullWeightAllowedWhenNotPublished":
                True,

            "thirdPartyHoldings":
                False,
        },

        "funds":
            all_results,

        "failed":
            failed,
    }

    save_json(
        ALL_HOLDINGS_FILE,
        all_holdings_payload,
    )

    # =========================================================================
    # RUN SUMMARY
    # =========================================================================

    total_holdings = sum(
        int(
            item.get(
                "topHoldingsCount",
                0,
            )
            or 0
        )
        for item
        in successful
    )

    total_published_percentages = sum(
        int(
            item.get(
                "publishedHoldingPercentageCount",
                0,
            )
            or 0
        )
        for item
        in successful
    )

    total_without_percentages = sum(
        int(
            item.get(
                "holdingsWithoutPublishedPercentage",
                0,
            )
            or 0
        )
        for item
        in successful
    )

    run_summary = {
        "status":
            (
                "success"
                if not failed
                else "partial"
            ),

        "startedAtUtc":
            started_at,

        "completedAtUtc":
            completed_at,

        "excelFile":
            str(
                EXCEL_FILE
            ),

        "excelFundUniverse":
            len(
                funds
            ),

        "successfulFunds":
            len(
                successful
            ),

        "noHoldingsSectionFunds":
            len(
                no_holdings_section
            ),

        "failedFunds":
            len(
                failed
            ),

        "totalPublishedTopHoldings":
            total_holdings,

        "totalHoldingsWithPublishedPercentage":
            total_published_percentages,

        "totalHoldingsWithoutPublishedPercentage":
            total_without_percentages,

        "successfulFundsDetail":
            [
                {
                    "excelRow":
                        item.get(
                            "excelRow"
                        ),

                    "fundName":
                        item.get(
                            "fundName"
                        ),

                    "factsheetUrl":
                        item.get(
                            "factsheetUrl"
                        ),

                    "factsheetDocumentDate":
                        item.get(
                            "factsheetDocumentDate"
                        ),

                    "factsheetDataAsAt":
                        item.get(
                            "factsheetDataAsAt"
                        ),

                    "topHoldingsCount":
                        item.get(
                            "topHoldingsCount"
                        ),

                    "publishedHoldingPercentageCount":
                        item.get(
                            "publishedHoldingPercentageCount"
                        ),

                    "holdingsWithoutPublishedPercentage":
                        item.get(
                            "holdingsWithoutPublishedPercentage"
                        ),

                    "holdingsParser":
                        item.get(
                            "holdingsParser"
                        ),
                }

                for item
                in successful
            ],

        "noHoldingsSectionDetail":
            [
                {
                    "excelRow":
                        item.get(
                            "excelRow"
                        ),

                    "fundName":
                        item.get(
                            "fundName"
                        ),

                    "factsheetUrl":
                        item.get(
                            "factsheetUrl"
                        ),
                }

                for item
                in no_holdings_section
            ],

        "failedFundsDetail":
            failed,

        "rules": {

            "excelColumnAControlsUniverse":
                True,

            "officialPrudentialFactsheetOnly":
                True,

            "maximumHoldings":
                MAX_HOLDINGS,

            "publishedHoldingCountUsedExactly":
                True,

            "fewerThanTenHoldingsAllowed":
                True,

            "noForcedTenEntries":
                True,

            "noInferredHoldings":
                True,

            "noFabricatedHoldingWeights":
                True,

            "nullWeightAllowedWhenNotPublished":
                True,

            "noThirdPartyHoldings":
                True,
        },
    }

    save_json(
        RUN_SUMMARY_FILE,
        run_summary,
    )

    # =========================================================================
    # CONSOLE SUMMARY
    # =========================================================================

    print(
        "\n\n"
        + "=" * 78
    )

    print(
        "PRUDENTIAL TOP HOLDINGS ALL-FUND TEST COMPLETE"
    )

    print(
        "=" * 78
    )

    print(
        f"Excel fund universe: "
        f"{len(funds)}"
    )

    print(
        f"Successful with holdings: "
        f"{len(successful)}"
    )

    print(
        f"No Top Holdings section: "
        f"{len(no_holdings_section)}"
    )

    print(
        f"Failed: "
        f"{len(failed)}"
    )

    print(
        f"Total published holdings extracted: "
        f"{total_holdings}"
    )

    print(
        f"Holdings with published percentage: "
        f"{total_published_percentages}"
    )

    print(
        f"Holdings without published percentage: "
        f"{total_without_percentages}"
    )

    print(
        "\nOutput:"
    )

    print(
        f" - {ALL_HOLDINGS_FILE}"
    )

    print(
        f" - {RUN_SUMMARY_FILE}"
    )

    print(
        f" - {FUNDS_OUTPUT_DIR}"
    )

    if failed:

        print(
            "\nFAILED FUNDS:"
        )

        for failure in failed:

            print(
                f" - Row "
                f"{failure['excelRow']}: "
                f"{failure['error']}"
            )

    print(
        "\nDone."
    )

    return 0


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":

    try:

        raise SystemExit(
            main()
        )

    except KeyboardInterrupt:

        print(
            "\nInterrupted.",
            file=sys.stderr,
        )

        raise SystemExit(
            130
        )
