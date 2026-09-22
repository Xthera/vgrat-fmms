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

Extract official Prudential Singapore Top Holdings information for every
fund listed in Excel Column A.

The script:

    Funds Links.xlsm
          |
          v
    Official Prudential fund page
          |
          v
    Official Prudential Fund Factsheet PDF
          |
          v
    Multiple PDF extraction modes
          |
          +--> normal text extraction
          |
          +--> layout-preserving extraction
          |
          +--> coordinate/row reconstruction
          |
          v
    Top Holdings section
          |
          v
    Validated holdings
          |
          v
    JSON output


HARD RULES
==========

- Excel Column A controls the universe.
- No hardcoded 67-fund limit.
- Every populated URL in Column A is processed.
- Only official Prudential Singapore pages are accepted.
- Only official Prudential Singapore factsheets are accepted.
- No third-party holdings source.
- No inferred holdings.
- No fabricated holdings.
- No fabricated percentages.
- No forced 10 holdings.
- If Prudential publishes fewer than 10 holdings, store exactly that count.
- Holdings are stored in published rank/order.
- Multi-line holding names are supported.
- Holdings may have no published percentage.
- Missing published percentage is stored as null.
- No percentage is calculated or inferred.
- A new holding rank may define the boundary of the previous holding.
- A published percentage may also define the boundary.
- Ambiguous PDF structures are rejected rather than guessed.
- Duplicate holding names are rejected.
- Explicit holding ranks must be sequential.
- More than 10 holdings are rejected.
- If a Top Holdings heading exists but contains no actual data rows,
  status is "no_holdings_data".
- If no Top Holdings section exists, status is "no_holdings_section".
- If Top Holdings data exists but cannot be parsed safely, the fund fails.
- Failed funds do not receive fabricated holdings.
- This script does not modify test_pruaccess.py.
- This script does not create data.json.
- This script does not create index.html/css/js.
"""


from __future__ import annotations

import json
import re
import sys
import time
from datetime import datetime, timezone
from io import BytesIO
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


# =============================================================================
# BROWSER
# =============================================================================

BROWSER_HEADLESS = True

PAGE_TIMEOUT_MS = 120000

FACTSHEET_DOWNLOAD_TIMEOUT_MS = 120000

POST_PAGE_WAIT_MS = 1500

RETRY_COUNT = 3

RETRY_DELAY_SECONDS = 3.0


# =============================================================================
# HOLDINGS
# =============================================================================

MAX_HOLDINGS = 10


# =============================================================================
# PDF EXTRACTION
# =============================================================================

# Y-coordinate tolerance used by coordinate-based PDF reconstruction.
PDF_ROW_Y_TOLERANCE = 3.5


# =============================================================================
# OFFICIAL PRUDENTIAL HOSTS
# =============================================================================

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
# FACTSHEET DISCOVERY
# =============================================================================

def find_factsheet_url(
    page,
    prudential_url: str,
) -> str:
    """
    Locate the official Fund Factsheet link.

    This intentionally only considers visible/actual anchor links exposed
    by the rendered Prudential page.

    It does NOT guess a PDF URL from the fund name.
    """

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
# PDF DOWNLOAD
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
# PDF TEXT EXTRACTION
# =============================================================================

def extract_pdf_text_plain(
    pdf_bytes: bytes,
) -> tuple[str, int]:

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
                pdf_page.extract_text(
                    extraction_mode="plain"
                )
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

    return (
        full_text,
        page_count,
    )


def extract_pdf_text_layout(
    pdf_bytes: bytes,
) -> tuple[str, int]:

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
                pdf_page.extract_text(
                    extraction_mode="layout"
                )
                or ""
            )

        except Exception as error:

            raise RuntimeError(
                "Failed to perform layout extraction "
                f"from page {page_index}: {error}"
            ) from error

        pages.append(
            page_text
        )

    return (
        "\n".join(
            pages
        ),
        page_count,
    )


def extract_pdf_text_coordinates(
    pdf_bytes: bytes,
) -> tuple[str, int]:

    """
    Reconstruct PDF text rows using text coordinates.

    This is useful when a factsheet contains a table whose left/right
    columns are extracted in an undesirable reading order.

    Text chunks are grouped approximately by their Y coordinate and
    then sorted from left to right.
    """

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

    all_pages = []

    for page_index, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):

        chunks = []

        def visitor_text(
            text,
            cm,
            tm,
            font_dict,
            font_size,
        ):

            if not text:
                return

            value = str(
                text
            )

            value = value.replace(
                "\r",
                "",
            )

            value = value.replace(
                "\x00",
                "",
            )

            if not value.strip():
                return

            try:

                x = float(
                    tm[4]
                )

                y = float(
                    tm[5]
                )

            except Exception:

                return

            chunks.append(
                {
                    "text":
                        value,

                    "x":
                        x,

                    "y":
                        y,
                }
            )

        try:

            pdf_page.extract_text(
                extraction_mode="plain",
                visitor_text=visitor_text,
            )

        except Exception as error:

            raise RuntimeError(
                "Coordinate PDF extraction failed "
                f"on page {page_index}: {error}"
            ) from error

        if not chunks:

            all_pages.append(
                ""
            )

            continue

        # Sort approximately top-to-bottom first.
        chunks.sort(
            key=lambda item: (
                -item["y"],
                item["x"],
            )
        )

        rows = []

        for chunk in chunks:

            assigned = False

            for row in rows:

                if abs(
                    row["y"]
                    - chunk["y"]
                ) <= PDF_ROW_Y_TOLERANCE:

                    row["chunks"].append(
                        chunk
                    )

                    # Keep the average Y for subsequent comparisons.
                    row["y"] = (
                        row["y"]
                        + chunk["y"]
                    ) / 2.0

                    assigned = True

                    break

            if not assigned:

                rows.append(
                    {
                        "y":
                            chunk["y"],

                        "chunks":
                            [
                                chunk
                            ],
                    }
                )

        # Re-sort rows top-to-bottom.
        rows.sort(
            key=lambda row: -row["y"]
        )

        page_lines = []

        for row in rows:

            row["chunks"].sort(
                key=lambda item: item["x"]
            )

            fragments = []

            for chunk in row[
                "chunks"
            ]:

                value = chunk[
                    "text"
                ]

                value = value.replace(
                    "\n",
                    " ",
                )

                value = clean_text(
                    value
                )

                if value:
                    fragments.append(
                        value
                    )

            if not fragments:
                continue

            line = " ".join(
                fragments
            )

            line = clean_text(
                line
            )

            if line:
                page_lines.append(
                    line
                )

        all_pages.append(
            "\n".join(
                page_lines
            )
        )

    return (
        "\n\n".join(
            all_pages
        ),
        page_count,
    )


def extract_pdf_text_candidates(
    pdf_bytes: bytes,
) -> list[dict]:

    candidates = []

    extraction_functions = [
        (
            "plain",
            extract_pdf_text_plain,
        ),
        (
            "layout",
            extract_pdf_text_layout,
        ),
        (
            "coordinates",
            extract_pdf_text_coordinates,
        ),
    ]

    for mode, function in extraction_functions:

        try:

            text, page_count = function(
                pdf_bytes
            )

            if clean_text(
                text
            ):

                candidates.append(
                    {
                        "mode":
                            mode,

                        "text":
                            text,

                        "pageCount":
                            page_count,
                    }
                )

        except Exception as error:

            print(
                f"PDF extraction mode '{mode}' failed: "
                f"{clean_text(str(error))}"
            )

    if not candidates:

        raise RuntimeError(
            "All PDF extraction modes failed."
        )

    return candidates


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
# FACTSHEET DATES
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
# HOLDINGS SECTION
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


def is_holdings_header_or_noise(
    line: str,
) -> bool:

    normalized = normalize_text(
        line
    )

    if not normalized:
        return True

    noise = {
        "holding",
        "holdings",
        "name",
        "names",
        "weight",
        "weights",
        "allocation",
        "allocations",
        "%",
        "portfolio holdings",
        "top 10 holdings",
        "top ten holdings",
    }

    if normalized in noise:
        return True

    if normalized.startswith(
        "top 10 holdings"
    ):
        return True

    if normalized.startswith(
        "top ten holdings"
    ):
        return True

    return False


def extract_holdings_section(
    text: str,
) -> tuple[str, str]:

    """
    Returns:

        section_text
        section_status

    Status:

        published
        not_published
        empty
    """

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

    meaningful_lines = []

    for line in section:

        if not is_holdings_header_or_noise(
            line
        ):

            meaningful_lines.append(
                line
            )

    if not meaningful_lines:

        return (
            "",
            "empty",
        )

    return (
        "\n".join(
            meaningful_lines
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

    name = re.sub(
        r"^[•·▪■□*]+",
        "",
        name,
    )

    name = name.strip()

    name = re.sub(
        r"^\d{1,2}[.)]\s+",
        "",
        name,
    )

    name = name.replace(
        "|",
        " ",
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

    name = name.strip(
        " -|"
    )

    if normalize_text(
        name
    ) in {
        "none",
        "null",
        "-",
        "—",
        "",
    }:
        return ""

    return name


def clean_holding_fragment(
    value: str,
) -> str:

    fragment = clean_text(
        value
    )

    if not fragment:
        return ""

    fragment = fragment.replace(
        "|",
        " ",
    )

    fragment = re.sub(
        r"\bNone\b",
        " ",
        fragment,
        flags=re.IGNORECASE,
    )

    fragment = re.sub(
        r"^[•·▪■□*]+",
        "",
        fragment,
    )

    return clean_text(
        fragment
    )


def combine_holding_name_fragments(
    fragments: list[str],
) -> str:

    cleaned_fragments = []

    for fragment in fragments:

        fragment = clean_holding_fragment(
            fragment
        )

        if not fragment:
            continue

        if re.fullmatch(
            r"\d{1,2}",
            fragment,
        ):
            continue

        cleaned_fragments.append(
            fragment
        )

    if not cleaned_fragments:
        return ""

    combined = " ".join(
        cleaned_fragments
    )

    combined = re.sub(
        r"\s+",
        " ",
        combined,
    ).strip()

    return clean_holding_name(
        combined
    )


# =============================================================================
# PERCENTAGE
# =============================================================================

def find_percentage_in_line(
    line: str,
) -> tuple[
    float,
    str,
    int,
    int,
] | None:

    if not line:
        return None

    match = re.search(
        r"""
        (?<![\d.])
        ([0-9]+(?:\.[0-9]+)?)
        \s*%
        """,
        line,
        re.VERBOSE,
    )

    if not match:
        return None

    try:

        percentage = float(
            match.group(1)
        )

    except ValueError:

        return None

    if (
        percentage < 0
        or percentage > 100
    ):

        return None

    return (
        percentage,
        clean_text(
            match.group(0)
        ),
        match.start(),
        match.end(),
    )


# =============================================================================
# HOLDING RANK
# =============================================================================

def extract_leading_rank(
    line: str,
) -> tuple[
    int | None,
    str,
]:

    text = clean_text(
        line
    )

    if not text:
        return (
            None,
            "",
        )

    # Explicit forms:
    #
    # 1 NAME
    # 1. NAME
    # 1) NAME
    # 1 - NAME
    # 1: NAME
    #
    match = re.match(
        r"^\s*(\d{1,2})(?:[.)\-:]|\s)\s*(.*)$",
        text,
    )

    if not match:
        return (
            None,
            text,
        )

    try:

        rank = int(
            match.group(1)
        )

    except ValueError:

        return (
            None,
            text,
        )

    if rank < 1 or rank > 99:

        return (
            None,
            text,
        )

    remainder = clean_text(
        match.group(2)
    )

    return (
        rank,
        remainder,
    )


# =============================================================================
# HOLDINGS PARSER
# =============================================================================

def parse_holdings(
    section_text: str,
) -> list[dict]:

    """
    Parse one extracted Top Holdings section.

    Supported structures include:

        1 HOLDING NAME 10.5%
        2 HOLDING NAME 9.7%

    and:

        1
        HOLDING NAME
        10.5%

    and:

        1
        HOLDING NAME
        2
        NEXT HOLDING

    where the new rank closes the previous holding.

    A percentage is optional.

    If no percentage is published:

        weightPercent = null
        weightText = null

    No percentage is inferred.
    """

    if not clean_text(
        section_text
    ):

        raise RuntimeError(
            "Top Holdings section is empty."
        )

    lines = pdf_lines(
        section_text
    )

    holdings = []

    seen_names = set()

    pending_fragments: list[str] = []

    pending_rank: int | None = None

    explicit_rank_count = 0

    def commit_pending():

        nonlocal pending_fragments
        nonlocal pending_rank

        name = combine_holding_name_fragments(
            pending_fragments
        )

        if not name:

            raise RuntimeError(
                "A holding boundary was found but "
                "no holding name could be extracted."
            )

        name_key = normalize_text(
            name
        )

        if name_key in seen_names:

            raise RuntimeError(
                "Duplicate holding name detected: "
                f"{name}"
            )

        if len(
            holdings
        ) >= MAX_HOLDINGS:

            raise RuntimeError(
                "More than 10 published holdings were parsed."
            )

        holding = {
            "rank":
                (
                    pending_rank
                    if pending_rank is not None
                    else len(holdings) + 1
                ),

            "name":
                name,

            "weightPercent":
                None,

            "weightText":
                None,
        }

        holdings.append(
            holding
        )

        seen_names.add(
            name_key
        )

        pending_fragments = []

        pending_rank = None

    for raw_line in lines:

        line = clean_text(
            raw_line
        )

        if not line:
            continue

        if is_holdings_header_or_noise(
            line
        ):
            continue

        detected_rank, remainder = (
            extract_leading_rank(
                line
            )
        )

        # ---------------------------------------------------------------------
        # NEW EXPLICIT RANK
        # ---------------------------------------------------------------------

        if detected_rank is not None:

            explicit_rank_count += 1

            # If another holding was already being accumulated, the new rank
            # safely closes it even if there was no percentage.
            if pending_fragments:

                commit_pending()

            pending_rank = (
                detected_rank
            )

            line = remainder

            # Standalone rank.
            if not line:
                continue

        # ---------------------------------------------------------------------
        # PUBLISHED PERCENTAGE
        # ---------------------------------------------------------------------

        percentage_info = (
            find_percentage_in_line(
                line
            )
        )

        if percentage_info is not None:

            (
                percentage,
                percentage_text,
                start_index,
                _end_index,
            ) = percentage_info

            name_fragment = clean_text(
                line[:start_index]
            )

            if name_fragment:

                pending_fragments.append(
                    name_fragment
                )

            if not pending_fragments:

                raise RuntimeError(
                    "A published holding percentage was found "
                    "without a holding name."
                )

            name = combine_holding_name_fragments(
                pending_fragments
            )

            if not name:

                raise RuntimeError(
                    "A published holding percentage was found "
                    "but no holding name could be extracted."
                )

            name_key = normalize_text(
                name
            )

            if name_key in seen_names:

                raise RuntimeError(
                    "Duplicate holding name detected: "
                    f"{name}"
                )

            if len(
                holdings
            ) >= MAX_HOLDINGS:

                raise RuntimeError(
                    "More than 10 published holdings were parsed."
                )

            rank = (
                pending_rank
                if pending_rank is not None
                else len(holdings) + 1
            )

            holdings.append(
                {
                    "rank":
                        rank,

                    "name":
                        name,

                    "weightPercent":
                        percentage,

                    "weightText":
                        percentage_text,
                }
            )

            seen_names.add(
                name_key
            )

            pending_fragments = []

            pending_rank = None

            continue

        # ---------------------------------------------------------------------
        # NO PERCENTAGE
        #
        # This may be a wrapped holding-name line.
        # ---------------------------------------------------------------------

        fragment = clean_holding_fragment(
            line
        )

        if not fragment:
            continue

        if is_holdings_header_or_noise(
            fragment
        ):
            continue

        # ---------------------------------------------------------------------
        # Ignore obvious metadata lines that can appear around the table.
        # ---------------------------------------------------------------------

        normalized_fragment = normalize_text(
            fragment
        )

        if normalized_fragment in {
            "source",
            "source:",
            "data as at",
            "weight",
            "weights",
            "holding",
            "holdings",
        }:
            continue

        pending_fragments.append(
            fragment
        )

    # -------------------------------------------------------------------------
    # Commit final holding.
    #
    # This is important for factsheets where Prudential publishes the holding
    # name but does not publish a percentage.
    # -------------------------------------------------------------------------

    if pending_fragments:

        commit_pending()

    if not holdings:

        raise RuntimeError(
            "Top Holdings section was found, "
            "but no holding rows could be safely parsed."
        )

    if len(
        holdings
    ) > MAX_HOLDINGS:

        raise RuntimeError(
            "Parser produced more than 10 holdings."
        )

    # -------------------------------------------------------------------------
    # Validate ranks.
    # -------------------------------------------------------------------------

    ranks = [
        item[
            "rank"
        ]
        for item in holdings
    ]

    if explicit_rank_count > 0:

        expected_ranks = list(
            range(
                1,
                len(
                    holdings
                ) + 1,
            )
        )

        if ranks != expected_ranks:

            raise RuntimeError(
                "Holding ranks are not sequential. "
                f"Parsed={ranks}; "
                f"Expected={expected_ranks}"
            )

    # -------------------------------------------------------------------------
    # Validate names and optional weights.
    # -------------------------------------------------------------------------

    names = []

    for holding in holdings:

        name = clean_holding_name(
            holding.get(
                "name"
            )
        )

        if not name:

            raise RuntimeError(
                "Holding name is empty."
            )

        holding[
            "name"
        ] = name

        names.append(
            normalize_text(
                name
            )
        )

        weight = holding.get(
            "weightPercent"
        )

        if weight is None:

            holding[
                "weightText"
            ] = None

            continue

        if not isinstance(
            weight,
            (int, float),
        ):

            raise RuntimeError(
                "Holding weight is invalid."
            )

        if (
            weight < 0
            or weight > 100
        ):

            raise RuntimeError(
                "Holding weight is outside 0-100%."
            )

    if len(
        names
    ) != len(
        set(
            names
        )
    ):

        raise RuntimeError(
            "Duplicate holding names detected."
        )

    return holdings


# =============================================================================
# HOLDINGS CANDIDATE VALIDATION
# =============================================================================

def evaluate_holdings_candidates(
    pdf_candidates: list[dict],
) -> dict:

    """
    Run the holdings parser against every PDF extraction mode.

    The best valid candidate is selected using:

        1. more holdings
        2. more published percentages
        3. extraction preference:
             coordinates
             layout
             plain

    We never merge candidates together.
    """

    valid_candidates = []

    no_data_candidates = []

    for candidate in pdf_candidates:

        mode = candidate[
            "mode"
        ]

        text = candidate[
            "text"
        ]

        section_text, section_status = (
            extract_holdings_section(
                text
            )
        )

        if section_status == "not_published":

            no_data_candidates.append(
                {
                    "mode":
                        mode,

                    "status":
                        "no_holdings_section",

                    "text":
                        text,

                    "pageCount":
                        candidate[
                            "pageCount"
                        ],

                    "sectionText":
                        "",
                }
            )

            continue

        if section_status == "empty":

            no_data_candidates.append(
                {
                    "mode":
                        mode,

                    "status":
                        "no_holdings_data",

                    "text":
                        text,

                    "pageCount":
                        candidate[
                            "pageCount"
                        ],

                    "sectionText":
                        "",
                }
            )

            continue

        try:

            holdings = parse_holdings(
                section_text
            )

        except Exception as error:

            print(
                f"Holdings parser rejected "
                f"PDF mode '{mode}': "
                f"{clean_text(str(error))}"
            )

            continue

        published_percentages = sum(
            1
            for holding
            in holdings
            if holding.get(
                "weightPercent"
            ) is not None
        )

        valid_candidates.append(
            {
                "mode":
                    mode,

                "status":
                    "success",

                "text":
                    text,

                "pageCount":
                    candidate[
                        "pageCount"
                    ],

                "sectionText":
                    section_text,

                "holdings":
                    holdings,

                "holdingCount":
                    len(
                        holdings
                    ),

                "publishedPercentages":
                    published_percentages,
            }
        )

    if valid_candidates:

        mode_priority = {
            "coordinates":
                3,

            "layout":
                2,

            "plain":
                1,
        }

        valid_candidates.sort(
            key=lambda item: (
                -item[
                    "holdingCount"
                ],

                -item[
                    "publishedPercentages"
                ],

                -mode_priority.get(
                    item[
                        "mode"
                    ],
                    0,
                ),
            )
        )

        return valid_candidates[0]

    # -------------------------------------------------------------------------
    # If every extraction mode says there is no holdings section/data, preserve
    # that information instead of marking the fund as a parser failure.
    # -------------------------------------------------------------------------

    if no_data_candidates:

        # Prefer a genuine empty Top Holdings section over "not published"
        # when at least one extraction mode can identify it.
        empty_candidates = [
            item
            for item
            in no_data_candidates
            if item[
                "status"
            ]
            == "no_holdings_data"
        ]

        if empty_candidates:

            return empty_candidates[0]

        return no_data_candidates[0]

    raise RuntimeError(
        "Top Holdings section was found, but none of the "
        "available PDF extraction modes could parse it safely."
    )


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
            "Prudential fund page redirected "
            "outside Prudential Singapore: "
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
    # Extract PDF using all supported modes.
    # -------------------------------------------------------------------------

    pdf_candidates = (
        extract_pdf_text_candidates(
            factsheet_bytes
        )
    )

    print(
        "PDF extraction modes available:"
    )

    for candidate in pdf_candidates:

        print(
            f"  - {candidate['mode']}"
        )

    selected = (
        evaluate_holdings_candidates(
            pdf_candidates
        )
    )

    selected_mode = selected[
        "mode"
    ]

    selected_text = selected[
        "text"
    ]

    pdf_page_count = selected[
        "pageCount"
    ]

    print(
        f"Selected PDF extraction mode: "
        f"{selected_mode}"
    )

    # -------------------------------------------------------------------------
    # Factsheet dates.
    # -------------------------------------------------------------------------

    data_as_at = (
        extract_data_as_at(
            selected_text
        )
    )

    document_date = (
        extract_document_date(
            selected_text
        )
    )

    # -------------------------------------------------------------------------
    # NO HOLDINGS SECTION.
    # -------------------------------------------------------------------------

    if selected[
        "status"
    ] == "no_holdings_section":

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

            "pdfExtractionMode":
                selected_mode,

            "holdingsSectionStatus":
                "not_published",

            "topHoldingsCount":
                0,

            "topHoldings":
                [],

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

                "missingPublishedPercentageStoredAsNull":
                    True,

                "multilineHoldingNamesSupported":
                    True,

                "rankDefinesHoldingBoundary":
                    True,

                "percentageDefinesHoldingBoundary":
                    True,
            },
        }

    # -------------------------------------------------------------------------
    # EMPTY TOP HOLDINGS SECTION.
    # -------------------------------------------------------------------------

    if selected[
        "status"
    ] == "no_holdings_data":

        return {
            "status":
                "no_holdings_data",

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

            "pdfExtractionMode":
                selected_mode,

            "holdingsSectionStatus":
                "empty",

            "topHoldingsCount":
                0,

            "topHoldings":
                [],

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

                "missingPublishedPercentageStoredAsNull":
                    True,

                "multilineHoldingNamesSupported":
                    True,

                "rankDefinesHoldingBoundary":
                    True,

                "percentageDefinesHoldingBoundary":
                    True,
            },
        }

    # -------------------------------------------------------------------------
    # SUCCESS.
    # -------------------------------------------------------------------------

    holdings = selected[
        "holdings"
    ]

    section_text = selected[
        "sectionText"
    ]

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

        "pdfExtractionMode":
            selected_mode,

        "holdingsSectionStatus":
            "published",

        "topHoldingsCount":
            len(
                holdings
            ),

        "topHoldings":
            holdings,

        "_sectionText":
            section_text,

        "_fullText":
            selected_text,

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

            "missingPublishedPercentageStoredAsNull":
                True,

            "multilineHoldingNamesSupported":
                True,

            "rankDefinesHoldingBoundary":
                True,

            "percentageDefinesHoldingBoundary":
                True,

            "multiplePdfExtractionModesTested":
                True,

            "coordinateBasedExtractionSupported":
                True,
        },
    }


# =============================================================================
# SAVE SUCCESS
# =============================================================================

def save_success_result(
    result: dict,
    factsheet_bytes: bytes,
) -> Path:

    excel_row = int(
        result[
            "excelRow"
        ]
    )

    identifier = ""

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

    (
        directory
        / "factsheet.pdf"
    ).write_bytes(
        factsheet_bytes
    )

    full_text = result.pop(
        "_fullText",
        "",
    )

    section_text = result.pop(
        "_sectionText",
        "",
    )

    (
        directory
        / "factsheet_text.txt"
    ).write_text(
        full_text,
        encoding="utf-8",
    )

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

        "pdfExtractionMode":
            result.get(
                "pdfExtractionMode"
            ),

        "topHoldingsCount":
            result.get(
                "topHoldingsCount"
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
# SAVE NO-HOLDINGS RESULT
# =============================================================================

def save_no_holdings_result(
    result: dict,
) -> Path:

    excel_row = int(
        result[
            "excelRow"
        ]
    )

    identifier = safe_filename(
        result.get(
            "fundName"
        )
        or
        urlparse(
            result.get(
                "finalUrl"
            )
            or result.get(
                "prudentialUrl"
            )
        ).path.rstrip(
            "/"
        ).split(
            "/"
        )[-1]
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

    save_json(
        directory
        / "top_holdings.json",
        result,
    )

    metadata = {
        "excelRow":
            result.get(
                "excelRow"
            ),

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

        "pdfExtractionMode":
            result.get(
                "pdfExtractionMode"
            ),

        "topHoldingsCount":
            0,

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

    no_holdings_data = []

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

                        last_error = clean_text(
                            str(error)
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

                # -----------------------------------------------------------------
                # Complete failure.
                # -----------------------------------------------------------------

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

                if result[
                    "status"
                ] == "no_holdings_section":

                    output_directory = (
                        save_no_holdings_result(
                            result
                        )
                    )

                    result[
                        "outputDirectory"
                    ] = str(
                        output_directory
                    )

                    no_holdings_section.append(
                        result
                    )

                    print(
                        "\nNO TOP HOLDINGS SECTION"
                    )

                    continue

                # -----------------------------------------------------------------
                # Top Holdings heading exists but no rows.
                # -----------------------------------------------------------------

                if result[
                    "status"
                ] == "no_holdings_data":

                    output_directory = (
                        save_no_holdings_result(
                            result
                        )
                    )

                    result[
                        "outputDirectory"
                    ] = str(
                        output_directory
                    )

                    no_holdings_data.append(
                        result
                    )

                    print(
                        "\nTOP HOLDINGS SECTION PRESENT "
                        "BUT NO HOLDINGS DATA"
                    )

                    continue

                # -----------------------------------------------------------------
                # SUCCESS
                #
                # Re-download official PDF for audit output.
                # -----------------------------------------------------------------

                try:

                    factsheet_response = (
                        page.request.get(
                            result[
                                "factsheetUrl"
                            ],
                            timeout=(
                                FACTSHEET_DOWNLOAD_TIMEOUT_MS
                            ),
                        )
                    )

                    if factsheet_response.status != 200:

                        raise RuntimeError(
                            "Factsheet re-download returned "
                            f"HTTP {factsheet_response.status}."
                        )

                    factsheet_bytes = (
                        factsheet_response.body()
                    )

                    if not factsheet_bytes.startswith(
                        b"%PDF"
                    ):

                        raise RuntimeError(
                            "Factsheet re-download did not "
                            "return a valid PDF."
                        )

                    # Re-run every extraction mode and verify that at least
                    # one produces the same or a valid holding result.
                    verification_candidates = (
                        extract_pdf_text_candidates(
                            factsheet_bytes
                        )
                    )

                    verified = (
                        evaluate_holdings_candidates(
                            verification_candidates
                        )
                    )

                    if verified[
                        "status"
                    ] != "success":

                        raise RuntimeError(
                            "Factsheet holdings section "
                            "could not be verified."
                        )

                    verified_holdings = verified[
                        "holdings"
                    ]

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
                        "pdfExtractionMode"
                    ] = verified[
                        "mode"
                    ]

                    result[
                        "_fullText"
                    ] = verified[
                        "text"
                    ]

                    result[
                        "_sectionText"
                    ] = verified[
                        "sectionText"
                    ]

                    output_directory = (
                        save_success_result(
                            result,
                            factsheet_bytes,
                        )
                    )

                    result[
                        "outputDirectory"
                    ] = str(
                        output_directory
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
                    "\nSUCCESS"
                )

                print(
                    f"Fund: "
                    f"{result.get('fundName') or '-'}"
                )

                print(
                    f"PDF extraction mode: "
                    f"{result.get('pdfExtractionMode') or '-'}"
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

                for holding in result[
                    "topHoldings"
                ]:

                    weight_text = (
                        holding.get(
                            "weightText"
                        )
                        or "-"
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
    # CONSOLIDATED OUTPUT
    # =========================================================================

    completed_at = utc_now_iso()

    all_results = (
        successful
        + no_holdings_section
        + no_holdings_data
    )

    all_holdings_payload = {
        "status":
            (
                "success"
                if not failed
                else "partial"
            ),

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

        "noHoldingsDataFunds":
            len(
                no_holdings_data
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

            "missingPublishedPercentageStoredAsNull":
                True,

            "multilineHoldingNamesSupported":
                True,

            "rankDefinesHoldingBoundary":
                True,

            "percentageDefinesHoldingBoundary":
                True,

            "multiplePdfExtractionModesTested":
                True,

            "coordinateBasedExtractionSupported":
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
    # SUMMARY
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
        sum(
            1
            for holding
            in item.get(
                "topHoldings",
                [],
            )
            if holding.get(
                "weightPercent"
            ) is not None
        )
        for item
        in successful
    )

    total_missing_percentages = (
        total_holdings
        -
        total_published_percentages
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

        "noHoldingsDataFunds":
            len(
                no_holdings_data
            ),

        "failedFunds":
            len(
                failed
            ),

        "totalPublishedTopHoldings":
            total_holdings,

        "totalPublishedPercentages":
            total_published_percentages,

        "totalMissingPublishedPercentages":
            total_missing_percentages,

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

                    "pdfExtractionMode":
                        item.get(
                            "pdfExtractionMode"
                        ),

                    "topHoldingsCount":
                        item.get(
                            "topHoldingsCount"
                        ),

                    "missingPublishedPercentages":
                        sum(
                            1
                            for holding
                            in item.get(
                                "topHoldings",
                                [],
                            )
                            if holding.get(
                                "weightPercent"
                            ) is None
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

        "noHoldingsDataDetail":
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
                in no_holdings_data
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

            "missingPublishedPercentageStoredAsNull":
                True,

            "noThirdPartyHoldings":
                True,

            "multilineHoldingNamesSupported":
                True,

            "wrappedPdfLinesJoined":
                True,

            "rankDefinesHoldingBoundary":
                True,

            "percentageDefinesHoldingBoundary":
                True,

            "multiplePdfExtractionModesTested":
                True,

            "coordinateBasedExtractionSupported":
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
        f"Top Holdings section but no data: "
        f"{len(no_holdings_data)}"
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
        f"Published percentages: "
        f"{total_published_percentages}"
    )

    print(
        f"Missing published percentages: "
        f"{total_missing_percentages}"
    )

    print(
        "\nPDF extraction modes:"
    )

    print(
        " - Plain"
    )

    print(
        " - Layout"
    )

    print(
        " - Coordinate/row reconstruction"
    )

    print(
        "\nHolding boundary:"
    )

    print(
        " - Explicit rank"
    )

    print(
        " - OR published percentage"
    )

    print(
        "\nMissing published percentage:"
    )

    print(
        " - Stored as null"
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
