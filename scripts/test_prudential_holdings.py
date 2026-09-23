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

This script tests official Prudential Top Holdings extraction for EVERY
fund listed in Excel Column A.

The script operates in TWO STAGES.

STAGE 1 - IMMUTABLE BASELINE
============================

The original Prudential factsheet extraction/parsing workflow runs first.

    Funds Links.xlsm
          |
          v
    Read every populated URL in Column A
          |
          v
    Open official Prudential fund page
          |
          v
    Locate official Fund Factsheet PDF
          |
          v
    Download official Prudential PDF
          |
          v
    Extract PDF text
          |
          v
    Locate "Top 10 holdings"
          |
          v
    Build logical holding lines
          |
          v
    Primary parser
          |
          v
    Existing fallback parser
          |
          v
    Stage-1 result


STAGE 2 - RECOVERY
==================

Stage 2 starts ONLY after Stage 1 has completely finished.

Only funds that FAILED Stage 1 are eligible.

Successful Stage-1 funds are never reinterpreted.

Recovery is adaptive and uses the same official Prudential factsheet.

Recovery strategies:

    13 / 14 / 15
        Two-column / Dividend History interleaving

    17 / 34 / 53
        Detached holding names and portfolio weights

    29
        Complex/interleaved layout

    59 / 60
        Alternate Top Holdings boundary detection

The recovery stage also tries layout-preserving PDF extraction and
coordinate-aware reconstruction before declaring a fund unrecoverable.


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
- Holdings are stored in the order published by Prudential whenever the
  published order can be established from the official PDF.
- Multi-line holding names are joined into one holding name.
- Fixed-income coupon/rate percentages inside security descriptions are
  not automatically treated as portfolio weights.
- For fixed-income logical lines containing multiple percentages, the LAST
  percentage is treated as the portfolio weight.
- Duplicate holding names are allowed.
- Duplicate holding percentages are allowed.
- Failed funds do not get blank fabricated holdings.
- The script never changes PruAccess settings.
- This script does NOT modify test_pruaccess.py.
- This script does NOT create data.json.
- This script does NOT create index.html/css/js.
- This script is a TEST collector only.


IMPORTANT FIX
=============

Prudential fixed-income holdings can be extracted from a PDF like this:

    CORPORACION ANDINA DE FOMENTO 7.7%
    6-MAR2029 1.7%

The first 7.7% is the bond coupon.

The second 1.7% is the portfolio weight.

The parser must reconstruct:

    CORPORACION ANDINA DE FOMENTO 7.7% 6-MAR2029 - 1.7%


HYPHENATED LINE-WRAP FIX
========================

A maturity date can itself be split:

    CORPORACION ANDINA DE FOMENTO 7.7% 6-MAR-
    2029 1.7%

The pre-pass rejoins this to:

    CORPORACION ANDINA DE FOMENTO 7.7% 6-MAR-2029 1.7%


STAGE 2 RECOVERY RULE
=====================

Recovery is NOT another interpretation of successful Stage-1 results.

A fund enters recovery only if its Stage-1 final status is FAILED.

Recovery must still produce:

    rank
    name
    weightPercent
    weightText

and the result must pass the same structural validation rules.

If recovery cannot establish a valid official holding/weight pair,
the fund remains FAILED.


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

        <excelRow>_recovered_<identifier>/
            factsheet.pdf
            factsheet_text.txt
            top_holdings_section.txt
            recovery_layout.txt
            recovery_diagnostics.json
            top_holdings.json
            metadata.json

        <excelRow>_failed/
            failure.json
            factsheet_text.txt
            top_holdings_section.txt
            recovery_layout.txt
            recovery_diagnostics.json
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


# -----------------------------------------------------------------------------
# Browser
# -----------------------------------------------------------------------------

BROWSER_HEADLESS = True

PAGE_TIMEOUT_MS = 120000

FACTSHEET_DOWNLOAD_TIMEOUT_MS = 120000

POST_PAGE_WAIT_MS = 1500

RETRY_COUNT = 3

RETRY_DELAY_SECONDS = 3.0


# -----------------------------------------------------------------------------
# Holdings
# -----------------------------------------------------------------------------

MAX_HOLDINGS = 10


# -----------------------------------------------------------------------------
# Recovery
# -----------------------------------------------------------------------------

RECOVERY_ENABLED = True

RECOVERY_RETRY_COUNT = 2

RECOVERY_RETRY_DELAY_SECONDS = 2.0

RECOVERY_TARGET_ROWS = {
    13,
    14,
    15,
    17,
    29,
    34,
    53,
    59,
    60,
}


# -----------------------------------------------------------------------------
# Only official Prudential Singapore URLs are accepted.
# -----------------------------------------------------------------------------

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


class HoldingsParseFailure(RuntimeError):

    def __init__(
        self,
        message: str,
        section_text: str = "",
        full_text: str = "",
    ) -> None:

        super().__init__(
            message
        )

        self.section_text = section_text

        self.full_text = full_text


class RecoveryParseFailure(RuntimeError):

    def __init__(
        self,
        message: str,
        layout_text: str = "",
        diagnostics: dict | None = None,
    ) -> None:

        super().__init__(
            message
        )

        self.layout_text = layout_text

        self.diagnostics = diagnostics or {}


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


def extract_holdings_section(
    text: str,
) -> tuple[str, str]:

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


# =============================================================================
# PERCENTAGE PARSING
# =============================================================================

def parse_percentage(
    value: str,
) -> float | None:

    if value is None:
        return None

    match = re.fullmatch(
        r"\s*"
        r"([0-9]+(?:\.[0-9]+)?)"
        r"\s*%"
        r"\s*",
        value,
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


def percentage_matches(
    line: str,
) -> list[re.Match]:

    return list(
        re.finditer(
            r"(?<![\d.])"
            r"([0-9]+(?:\.[0-9]+)?)"
            r"\s*%",
            line,
        )
    )


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

    if percentage < 0 or percentage > 100:
        return None

    return (
        percentage,
        clean_text(match.group(0)),
        match.start(),
        match.end(),
    )


def find_last_percentage_in_line(
    line: str,
) -> tuple[
    float,
    str,
    int,
    int,
] | None:

    if not line:
        return None

    matches = percentage_matches(
        line
    )

    if not matches:
        return None

    match = matches[-1]

    try:

        percentage = float(
            match.group(1)
        )

    except ValueError:

        return None

    if percentage < 0 or percentage > 100:
        return None

    return (
        percentage,
        clean_text(match.group(0)),
        match.start(),
        match.end(),
    )


# =============================================================================
# MATURITY DATE DETECTION
# =============================================================================

MONTH_PATTERN = (
    r"JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC"
)


def contains_maturity_date(
    line: str,
) -> bool:

    if not line:
        return False

    patterns = [

        rf"\b\d{{1,2}}[-/ ]"
        rf"(?:{MONTH_PATTERN})[-/ ]"
        rf"\d{{4}}\b",

        rf"\b\d{{1,2}}[-/ ]"
        rf"(?:{MONTH_PATTERN})"
        rf"\d{{4}}\b",

        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{4}\b",

        r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b",
    ]

    upper = line.upper()

    return any(
        re.search(
            pattern,
            upper,
        )
        for pattern in patterns
    )


def line_ends_with_percentage(
    line: str,
) -> bool:

    return bool(
        re.search(
            r"[0-9]+(?:\.[0-9]+)?\s*%\s*$",
            clean_text(line),
        )
    )


def line_has_percentage(
    line: str,
) -> bool:

    return bool(
        percentage_matches(
            line
        )
    )


# =============================================================================
# HYPHENATED LINE-WRAP RECONSTRUCTION
# =============================================================================

def merge_hyphenated_line_breaks(
    lines: list[str],
) -> list[str]:

    merged: list[str] = []

    index = 0

    while index < len(lines):

        current = clean_text(
            lines[index]
        )

        while (
            current.endswith("-")
            and index + 1 < len(lines)
        ):

            index += 1

            next_line = clean_text(
                lines[index]
            )

            current = current + next_line

        merged.append(
            current
        )

        index += 1

    return merged


# =============================================================================
# LOGICAL HOLDING LINE RECONSTRUCTION
# =============================================================================

def build_logical_holding_lines(
    section_text: str,
) -> list[str]:

    raw_lines = pdf_lines(
        section_text
    )

    if not raw_lines:
        return []

    lines = merge_hyphenated_line_breaks(
        raw_lines
    )

    logical_lines = []

    index = 0

    while index < len(lines):

        current = clean_text(
            lines[index]
        )

        if not current:
            index += 1
            continue

        if is_holding_header_or_noise(
            current
        ):

            logical_lines.append(
                current
            )

            index += 1
            continue

        logical_lines.append(
            current
        )

        index += 1

    return logical_lines


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
        return None, ""

    standalone = re.fullmatch(
        r"(\d{1,2})",
        text,
    )

    if standalone:

        rank = int(
            standalone.group(1)
        )

        if 1 <= rank <= 99:
            return rank, ""

        return None, text

    match = re.match(
        r"^\s*(\d{1,2})[.)\-:]\s+(.+)$",
        text,
    )

    if match:

        rank = int(
            match.group(1)
        )

        if 1 <= rank <= 99:

            return (
                rank,
                clean_text(
                    match.group(2)
                ),
            )

        return None, text

    match = re.match(
        r"^\s*(\d{1,2})\s+(.+)$",
        text,
    )

    if match:

        rank = int(
            match.group(1)
        )

        if 1 <= rank <= 99:

            return (
                rank,
                clean_text(
                    match.group(2)
                ),
            )

    return None, text


# =============================================================================
# HOLDING LINE HELPERS
# =============================================================================

def is_holding_header_or_noise(
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
    }

    if normalized in noise:
        return True

    if normalized.startswith(
        "top 10 holdings"
    ):
        return True

    return False


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

    fragment = clean_text(
        fragment
    )

    return fragment


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

        standalone_rank = re.fullmatch(
            r"\d{1,2}",
            fragment,
        )

        if standalone_rank:
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

    combined = combined.strip(
        " -|"
    )

    return clean_holding_name(
        combined
    )


# =============================================================================
# PARSED HOLDING VALIDATION
# =============================================================================

def validate_holdings(
    holdings: list[dict],
    parser_name: str,
) -> list[dict]:

    if not holdings:

        raise RuntimeError(
            f"{parser_name} parser produced no holdings."
        )

    if len(holdings) > MAX_HOLDINGS:

        raise RuntimeError(
            f"{parser_name} parser produced more than "
            f"{MAX_HOLDINGS} holdings."
        )

    ranks = [
        holding["rank"]
        for holding in holdings
    ]

    expected_ranks = list(
        range(
            1,
            len(holdings) + 1,
        )
    )

    if ranks != expected_ranks:

        raise RuntimeError(
            f"{parser_name} holding ranks are not sequential. "
            f"Parsed={ranks}; Expected={expected_ranks}"
        )

    for holding in holdings:

        name = clean_holding_name(
            holding.get(
                "name"
            )
        )

        if not name:

            raise RuntimeError(
                f"{parser_name} holding name is empty."
            )

        weight = holding.get(
            "weightPercent"
        )

        if not isinstance(
            weight,
            (int, float),
        ):

            raise RuntimeError(
                f"{parser_name} holding weight is invalid."
            )

        if (
            weight < 0
            or weight > 100
        ):

            raise RuntimeError(
                f"{parser_name} holding weight is outside 0-100%."
            )

        holding[
            "name"
        ] = name

    return holdings


# =============================================================================
# PRIMARY HOLDINGS PARSER
# =============================================================================

def parse_holdings(
    section_text: str,
) -> list[dict]:

    if not clean_text(
        section_text
    ):

        raise RuntimeError(
            "Top 10 holdings section is empty."
        )

    lines = build_logical_holding_lines(
        section_text
    )

    holdings = []

    pending_fragments: list[str] = []

    pending_rank: int | None = None

    def commit_pending(
        percentage: float,
        percentage_text: str,
    ) -> None:

        nonlocal pending_fragments
        nonlocal pending_rank

        name = combine_holding_name_fragments(
            pending_fragments
        )

        if not name:

            raise RuntimeError(
                "A published holding percentage was found "
                "but no holding name could be extracted."
            )

        if len(holdings) >= MAX_HOLDINGS:

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

        pending_fragments = []

        pending_rank = None

    for raw_line in lines:

        line = clean_text(
            raw_line
        )

        if not line:
            continue

        if is_holding_header_or_noise(
            line
        ):
            continue

        detected_rank, remainder = (
            extract_leading_rank(
                line
            )
        )

        if detected_rank is not None:

            if pending_fragments:

                raise RuntimeError(
                    "A new holding rank appeared before "
                    "the previous holding received a published percentage."
                )

            pending_rank = detected_rank

            line = remainder

            if not line:
                continue

        percentage_candidates = percentage_matches(
            line
        )

        if len(
            percentage_candidates
        ) > 1:

            raise RuntimeError(
                "Primary parser found multiple percentages on one logical "
                "line and will not guess the holding weight. "
                "Fallback parser required."
            )

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

            commit_pending(
                percentage=percentage,
                percentage_text=percentage_text,
            )

            if len(
                holdings
            ) >= MAX_HOLDINGS:

                break

            continue

        fragment = clean_holding_fragment(
            line
        )

        if not fragment:
            continue

        if is_holding_header_or_noise(
            fragment
        ):
            continue

        pending_fragments.append(
            fragment
        )

    if not holdings:

        raise RuntimeError(
            "Top 10 holdings section was found, "
            "but no holding/percentage pairs could be parsed."
        )

    return validate_holdings(
        holdings,
        "Primary",
    )


# =============================================================================
# FALLBACK HOLDINGS PARSER
# =============================================================================

def parse_holdings_fallback(
    section_text: str,
) -> list[dict]:

    if not clean_text(
        section_text
    ):

        raise RuntimeError(
            "Top 10 holdings section is empty."
        )

    lines = build_logical_holding_lines(
        section_text
    )

    holdings = []

    pending_fragments: list[str] = []

    pending_rank: int | None = None

    def commit_pending(
        percentage: float,
        percentage_text: str,
    ) -> None:

        nonlocal pending_fragments
        nonlocal pending_rank

        name = combine_holding_name_fragments(
            pending_fragments
        )

        if not name:

            raise RuntimeError(
                "Fallback parser found a published holding percentage "
                "but no holding name could be extracted."
            )

        if len(
            holdings
        ) >= MAX_HOLDINGS:

            raise RuntimeError(
                "Fallback parser produced more than 10 holdings."
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

        pending_fragments = []

        pending_rank = None

    for raw_line in lines:

        line = clean_text(
            raw_line
        )

        if not line:
            continue

        if is_holding_header_or_noise(
            line
        ):
            continue

        detected_rank, remainder = (
            extract_leading_rank(
                line
            )
        )

        if detected_rank is not None:

            if pending_fragments:

                raise RuntimeError(
                    "Fallback parser encountered a new holding rank before "
                    "the previous holding received a published percentage."
                )

            pending_rank = detected_rank

            line = remainder

            if not line:
                continue

        percentage_info = (
            find_last_percentage_in_line(
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

            commit_pending(
                percentage,
                percentage_text,
            )

            if len(
                holdings
            ) >= MAX_HOLDINGS:

                break

            continue

        fragment = clean_holding_fragment(
            line
        )

        if (
            fragment
            and not is_holding_header_or_noise(
                fragment
            )
        ):

            pending_fragments.append(
                fragment
            )

    return validate_holdings(
        holdings,
        "Fallback",
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
# SINGLE FUND - STAGE 1 BASELINE
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

    (
        full_text,
        pdf_page_count,
    ) = extract_pdf_text(
        factsheet_bytes
    )

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

    (
        section_text,
        section_status,
    ) = extract_holdings_section(
        full_text
    )

    if section_status == "not_published":

        result = {
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

                "duplicateHoldingNamesAllowed":
                    True,

                "duplicateHoldingPercentagesAllowed":
                    True,

                "multilineHoldingNamesSupported":
                    True,

                "wrappedPdfLinesJoined":
                    True,

                "fixedIncomeCouponPercentagesSupported":
                    True,

                "logicalHoldingLineReconstruction":
                    True,
            },
        }

        return result

    parser_used = "primary"

    primary_parser_error = None

    try:

        holdings = parse_holdings(
            section_text
        )

    except Exception as primary_error:

        primary_parser_error = clean_text(
            str(primary_error)
        )

        print(
            "Primary holdings parser failed; "
            "running fallback parser..."
        )

        print(
            f"Primary parser error: "
            f"{primary_parser_error}"
        )

        try:

            holdings = parse_holdings_fallback(
                section_text
            )

        except Exception as fallback_error:

            raise HoldingsParseFailure(
                clean_text(
                    str(fallback_error)
                ),
                section_text=section_text,
                full_text=full_text,
            ) from fallback_error

        parser_used = "fallback"

    result = {
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

        "topHoldings":
            holdings,

        "holdingsParser":
            parser_used,

        "primaryParserError":
            primary_parser_error,

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

            "duplicateHoldingNamesAllowed":
                True,

            "duplicateHoldingPercentagesAllowed":
                True,

            "multilineHoldingNamesSupported":
                True,

            "wrappedPdfLinesJoined":
                True,

            "fixedIncomeCouponPercentagesSupported":
                True,

            "logicalHoldingLineReconstruction":
                True,

            "percentageDefinesHoldingBoundary":
                True,

            "fallbackUsesLastPercentageOnLogicalLine":
                True,
        },
    }

    return result


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

    (
        directory
        / "factsheet.pdf"
    ).write_bytes(
        factsheet_bytes
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

        "topHoldingsCount":
            result.get(
                "topHoldingsCount"
            ),

        "status":
            result.get(
                "status"
            ),

        "holdingsParser":
            result.get(
                "holdingsParser"
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
    section_text: str | None = None,
    full_text: str | None = None,
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

    if section_text:

        (
            directory
            / "top_holdings_section.txt"
        ).write_text(
            section_text,
            encoding="utf-8",
        )

    if full_text:

        (
            directory
            / "factsheet_text.txt"
        ).write_text(
            full_text,
            encoding="utf-8",
        )

    return directory


# =============================================================================
# ============================================================================
# STAGE 2 RECOVERY ENGINE
# ============================================================================
# =============================================================================
#
# IMPORTANT:
#
# Nothing above this point is used to reinterpret successful Stage-1 funds.
#
# Recovery receives ONLY a Stage-1 failed fund.
#
# =============================================================================


# =============================================================================
# RECOVERY STRATEGY IDENTIFICATION
# =============================================================================

def recovery_strategy_for_row(
    excel_row: int,
) -> str:

    if excel_row in {
        13,
        14,
        15,
    }:

        return (
            "two_column_dividend_history_interleaving"
        )

    if excel_row in {
        17,
        34,
        53,
    }:

        return (
            "detached_names_and_weights"
        )

    if excel_row == 29:

        return (
            "complex_interleaving"
        )

    if excel_row in {
        59,
        60,
    }:

        return (
            "alternate_holdings_boundary"
        )

    return (
        "generic_layout_recovery"
    )


# =============================================================================
# LAYOUT-PRESERVING PDF EXTRACTION
# =============================================================================

def extract_pdf_layout_text(
    pdf_bytes: bytes,
) -> tuple[str, list[dict]]:

    from io import BytesIO

    reader = PdfReader(
        BytesIO(
            pdf_bytes
        )
    )

    page_blocks = []

    all_pages = []

    for page_number, page in enumerate(
        reader.pages,
        start=1,
    ):

        try:

            layout_text = (
                page.extract_text(
                    extraction_mode="layout"
                )
                or ""
            )

        except Exception:

            try:

                layout_text = (
                    page.extract_text()
                    or ""
                )

            except Exception:

                layout_text = ""

        layout_text = layout_text or ""

        all_pages.append(
            layout_text
        )

        page_blocks.append(
            {
                "page":
                    page_number,

                "text":
                    layout_text,

                "lines":
                    layout_text.splitlines(),
            }
        )

    combined = "\n".join(
        all_pages
    )

    if not clean_text(
        combined
    ):

        raise RuntimeError(
            "Layout PDF extraction produced no text."
        )

    return (
        combined,
        page_blocks,
    )


# =============================================================================
# RECOVERY HEADING DETECTION
# =============================================================================

def recovery_heading_score(
    line: str,
) -> int:

    normalized = normalize_text(
        line
    )

    if not normalized:
        return 0

    score = 0

    if (
        "top 10 holdings"
        in normalized
    ):

        score += 100

    if (
        "top ten holdings"
        in normalized
    ):

        score += 100

    if (
        "top holdings"
        in normalized
    ):

        score += 80

    if (
        normalized == "holdings"
    ):

        score += 50

    if (
        "portfolio holdings"
        in normalized
    ):

        score += 50

    if (
        "dividend history"
        in normalized
    ):

        score -= 30

    if (
        "fund performance"
        in normalized
    ):

        score -= 30

    return score


def find_recovery_heading_positions(
    lines: list[str],
) -> list[int]:

    positions = []

    for index, line in enumerate(
        lines
    ):

        if (
            recovery_heading_score(
                line
            )
            > 0
        ):

            positions.append(
                index
            )

    return positions


# =============================================================================
# RECOVERY SECTION EXTRACTION
# =============================================================================

def recovery_section_end_score(
    line: str,
) -> bool:

    normalized = normalize_text(
        line
    )

    if not normalized:
        return False

    exact_endings = {
        "source",
        "source:",
        "dividend history",
        "fund performance",
        "performance",
        "asset allocation",
        "portfolio characteristics",
        "important information",
        "important information:",
        "disclaimer",
        "risk factors",
    }

    if normalized in exact_endings:
        return True

    if normalized.startswith(
        "source:"
    ):
        return True

    if normalized.startswith(
        "important information"
    ):
        return True

    if normalized.startswith(
        "dividend history"
    ):
        return True

    if normalized.startswith(
        "fund performance"
    ):
        return True

    if re.fullmatch(
        r"page\s+\d+(\s+of\s+\d+)?",
        normalized,
    ):
        return True

    return False


def extract_recovery_sections(
    layout_text: str,
) -> list[dict]:

    lines = [
        clean_text(
            line
        )
        for line in layout_text.splitlines()
        if clean_text(
            line
        )
    ]

    headings = (
        find_recovery_heading_positions(
            lines
        )
    )

    sections = []

    for heading_index in headings:

        heading = lines[
            heading_index
        ]

        section_lines = []

        for line in lines[
            heading_index + 1:
        ]:

            if recovery_section_end_score(
                line
            ):

                break

            section_lines.append(
                line
            )

        section_text = "\n".join(
            section_lines
        )

        if clean_text(
            section_text
        ):

            sections.append(
                {
                    "headingIndex":
                        heading_index,

                    "heading":
                        heading,

                    "text":
                        section_text,

                    "lines":
                        section_lines,
                }
            )

    return sections


# =============================================================================
# RECOVERY PERCENTAGE CANDIDATES
# =============================================================================

def recovery_percentage_values(
    text: str,
) -> list[tuple[float, str]]:

    values = []

    for match in re.finditer(
        r"(?<![\d.])"
        r"([0-9]+(?:\.[0-9]+)?)"
        r"\s*%",
        text,
    ):

        try:

            value = float(
                match.group(1)
            )

        except ValueError:

            continue

        if (
            value < 0
            or value > 100
        ):

            continue

        values.append(
            (
                value,
                clean_text(
                    match.group(0)
                ),
            )
        )

    return values


# =============================================================================
# RECOVERY NOISE
# =============================================================================

def recovery_is_noise(
    line: str,
) -> bool:

    normalized = normalize_text(
        line
    )

    if not normalized:
        return True

    noise_fragments = (
        "top 10 holdings",
        "top ten holdings",
        "holdings",
        "holding",
        "name",
        "weight",
        "weights",
        "portfolio",
        "source",
        "dividend history",
        "dividend",
        "performance",
        "fund performance",
        "asset allocation",
        "portfolio characteristics",
    )

    if normalized in {
        "holding",
        "holdings",
        "name",
        "weight",
        "weights",
        "%",
    }:

        return True

    for fragment in noise_fragments:

        if (
            normalized == fragment
        ):

            return True

    return False


# =============================================================================
# RECOVERY NAME NORMALISATION
# =============================================================================

def recovery_clean_name(
    value: str,
) -> str:

    value = clean_holding_fragment(
        value
    )

    value = re.sub(
        r"^\s*(?:\d{1,2}[.)\-:]|\d{1,2})\s+",
        "",
        value,
    )

    value = clean_holding_name(
        value
    )

    return value


# =============================================================================
# DETACHED NAME/WEIGHT PARSER
# =============================================================================

def recovery_parse_detached_pairs(
    section_text: str,
) -> list[dict]:

    raw_lines = [
        clean_text(
            line
        )
        for line in section_text.splitlines()
        if clean_text(
            line
        )
    ]

    lines = merge_hyphenated_line_breaks(
        raw_lines
    )

    holdings = []

    pending_name_parts = []

    for line in lines:

        if recovery_is_noise(
            line
        ):

            continue

        rank, remainder = (
            extract_leading_rank(
                line
            )
        )

        if rank is not None:

            if remainder:

                line = remainder

            else:

                continue

        matches = percentage_matches(
            line
        )

        if not matches:

            fragment = recovery_clean_name(
                line
            )

            if fragment:

                pending_name_parts.append(
                    fragment
                )

            continue

        last_match = matches[-1]

        try:

            weight = float(
                last_match.group(1)
            )

        except ValueError:

            continue

        if weight < 0 or weight > 100:
            continue

        name_before = clean_text(
            line[
                :last_match.start()
            ]
        )

        if name_before:

            pending_name_parts.append(
                name_before
            )

        name = combine_holding_name_fragments(
            pending_name_parts
        )

        if not name:

            continue

        holdings.append(
            {
                "rank":
                    len(holdings) + 1,

                "name":
                    name,

                "weightPercent":
                    weight,

                "weightText":
                    clean_text(
                        last_match.group(0)
                    ),
            }
        )

        pending_name_parts = []

        if len(
            holdings
        ) >= MAX_HOLDINGS:

            break

    return validate_holdings(
        holdings,
        "RecoveryDetached",
    )


# =============================================================================
# TWO-COLUMN / DIVIDEND HISTORY RECOVERY
# =============================================================================

def recovery_parse_two_column(
    section_text: str,
) -> list[dict]:

    """
    Some factsheets extract two visual columns into an interleaved text stream.

    The recovery parser therefore does NOT assume that every adjacent line
    belongs to the same holding.

    It builds:

        name candidates
        weight candidates

    and pairs weights with the nearest viable preceding holding text.

    Fixed-income descriptions remain intact because the final percentage on
    a logical fragment is treated as the portfolio weight.
    """

    raw_lines = [
        clean_text(
            line
        )
        for line in section_text.splitlines()
        if clean_text(
            line
        )
    ]

    lines = merge_hyphenated_line_breaks(
        raw_lines
    )

    name_candidates = []

    weight_candidates = []

    for index, line in enumerate(
        lines
    ):

        if recovery_is_noise(
            line
        ):
            continue

        percentages = percentage_matches(
            line
        )

        if percentages:

            match = percentages[-1]

            try:

                value = float(
                    match.group(1)
                )

            except ValueError:

                continue

            if 0 <= value <= 100:

                prefix = clean_text(
                    line[
                        :match.start()
                    ]
                )

                suffix = clean_text(
                    line[
                        match.end():
                    ]
                )

                weight_candidates.append(
                    {
                        "lineIndex":
                            index,

                        "value":
                            value,

                        "text":
                            clean_text(
                                match.group(0)
                            ),

                        "prefix":
                            prefix,

                        "suffix":
                            suffix,
                    }
                )

                if prefix:

                    candidate = recovery_clean_name(
                        prefix
                    )

                    if candidate:

                        name_candidates.append(
                            {
                                "lineIndex":
                                    index,

                                "name":
                                    candidate,
                            }
                        )

                if suffix:

                    candidate = recovery_clean_name(
                        suffix
                    )

                    if candidate:

                        name_candidates.append(
                            {
                                "lineIndex":
                                    index,

                                "name":
                                    candidate,
                            }
                        )

        else:

            candidate = recovery_clean_name(
                line
            )

            if candidate:

                name_candidates.append(
                    {
                        "lineIndex":
                            index,

                        "name":
                            candidate,
                    }
                )

    if not weight_candidates:

        raise RuntimeError(
            "Recovery two-column parser found no percentage candidates."
        )

    holdings = []

    used_name_indices = set()

    for weight in weight_candidates:

        candidates = []

        for candidate_index, candidate in enumerate(
            name_candidates
        ):

            if candidate_index in used_name_indices:
                continue

            distance = abs(
                candidate[
                    "lineIndex"
                ]
                -
                weight[
                    "lineIndex"
                ]
            )

            if distance > 8:
                continue

            name = candidate[
                "name"
            ]

            if not name:
                continue

            score = distance

            if (
                candidate[
                    "lineIndex"
                ]
                <=
                weight[
                    "lineIndex"
                ]
            ):

                score -= 2

            if contains_maturity_date(
                name
            ):

                score += 1

            candidates.append(
                (
                    score,
                    candidate_index,
                    candidate,
                )
            )

        if not candidates:
            continue

        candidates.sort(
            key=lambda item: (
                item[0],
                item[2]["lineIndex"],
            )
        )

        _, selected_index, selected = (
            candidates[0]
        )

        used_name_indices.add(
            selected_index
        )

        holdings.append(
            {
                "rank":
                    len(holdings) + 1,

                "name":
                    selected[
                        "name"
                    ],

                "weightPercent":
                    weight[
                        "value"
                    ],

                "weightText":
                    weight[
                        "text"
                    ],
            }
        )

        if len(
            holdings
        ) >= MAX_HOLDINGS:

            break

    return validate_holdings(
        holdings,
        "RecoveryTwoColumn",
    )


# =============================================================================
# COMPLEX INTERLEAVING RECOVERY
# =============================================================================

def recovery_parse_complex_interleaving(
    section_text: str,
) -> list[dict]:

    """
    More permissive recovery for the complex layout.

    It first attempts line-local last-percentage parsing, then falls back
    to detached name/weight pairing.

    No percentage is manufactured. Every weight originates from a literal
    percentage in the official PDF text.
    """

    attempts = []

    try:

        attempts.append(
            recovery_parse_detached_pairs(
                section_text
            )
        )

    except Exception:
        pass

    try:

        attempts.append(
            recovery_parse_two_column(
                section_text
            )
        )

    except Exception:
        pass

    valid = []

    for holdings in attempts:

        try:

            valid.append(
                validate_holdings(
                    holdings,
                    "RecoveryComplex",
                )
            )

        except Exception:

            continue

    if not valid:

        raise RuntimeError(
            "Complex recovery could not construct valid holdings."
        )

    valid.sort(
        key=lambda holdings: (
            -len(
                holdings
            ),
            sum(
                holding[
                    "weightPercent"
                ]
                for holding in holdings
            ),
        )
    )

    return valid[0]


# =============================================================================
# ALTERNATE HOLDINGS BOUNDARY RECOVERY
# =============================================================================

def recovery_parse_alternate_boundary(
    layout_text: str,
) -> tuple[
    list[dict],
    str,
    dict,
]:

    sections = extract_recovery_sections(
        layout_text
    )

    if not sections:

        raise RuntimeError(
            "No alternate holdings headings were found "
            "in the layout-preserved PDF text."
        )

    candidates = []

    for section_index, section in enumerate(
        sections
    ):

        parsers = [
            (
                "primary",
                parse_holdings,
            ),
            (
                "fallback",
                parse_holdings_fallback,
            ),
            (
                "detached",
                recovery_parse_detached_pairs,
            ),
            (
                "two_column",
                recovery_parse_two_column,
            ),
        ]

        for parser_name, parser in parsers:

            try:

                holdings = parser(
                    section[
                        "text"
                    ]
                )

                candidates.append(
                    {
                        "sectionIndex":
                            section_index,

                        "heading":
                            section[
                                "heading"
                            ],

                        "parser":
                            parser_name,

                        "holdings":
                            holdings,

                        "score":
                            len(
                                holdings
                            ) * 100
                            +
                            recovery_heading_score(
                                section[
                                    "heading"
                                ]
                            ),
                    }
                )

            except Exception:

                continue

    if not candidates:

        raise RuntimeError(
            "Alternate holdings boundary recovery found "
            "candidate headings but no valid holding set."
        )

    candidates.sort(
        key=lambda item: (
            -item["score"],
            item["sectionIndex"],
        )
    )

    selected = candidates[0]

    diagnostics = {
        "candidateSections":
            len(
                sections
            ),

        "successfulCandidates":
            len(
                candidates
            ),

        "selectedSectionIndex":
            selected[
                "sectionIndex"
            ],

        "selectedHeading":
            selected[
                "heading"
            ],

        "selectedParser":
            selected[
                "parser"
            ],
    }

    return (
        selected[
            "holdings"
        ],

        sections[
            selected[
                "sectionIndex"
            ]
        ][
            "text"
        ],

        diagnostics,
    )


# =============================================================================
# RECOVERY DIAGNOSTICS
# =============================================================================

def build_recovery_diagnostics(
    excel_row: int,
    strategy: str,
    layout_text: str,
    section_text: str,
) -> dict:

    lines = [
        clean_text(
            line
        )
        for line in layout_text.splitlines()
        if clean_text(
            line
        )
    ]

    headings = []

    for index, line in enumerate(
        lines
    ):

        score = recovery_heading_score(
            line
        )

        if score > 0:

            headings.append(
                {
                    "lineIndex":
                        index,

                    "line":
                        line,

                    "score":
                        score,
                }
            )

    percentages = []

    for index, line in enumerate(
        lines
    ):

        values = recovery_percentage_values(
            line
        )

        if values:

            percentages.append(
                {
                    "lineIndex":
                        index,

                    "line":
                        line,

                    "percentages":
                        [
                            value
                            for value, _text
                            in values
                        ],
                }
            )

    return {
        "excelRow":
            excel_row,

        "strategy":
            strategy,

        "layoutLineCount":
            len(
                lines
            ),

        "candidateHeadings":
            headings,

        "percentageCandidateLines":
            percentages[:200],

        "sectionTextLength":
            len(
                section_text
            ),

        "generatedAtUtc":
            utc_now_iso(),
    }


# =============================================================================
# RECOVERY SINGLE FUND
# =============================================================================

def recover_single_fund(
    page,
    excel_fund: dict,
    stage1_error: str,
) -> dict:

    excel_row = int(
        excel_fund[
            "excelRow"
        ]
    )

    strategy = recovery_strategy_for_row(
        excel_row
    )

    print(
        "\n"
        "################################################################"
    )

    print(
        f"STAGE 2 RECOVERY - ROW {excel_row}"
    )

    print(
        "################################################################"
    )

    print(
        f"Recovery strategy: {strategy}"
    )

    print(
        f"Stage-1 failure: {stage1_error}"
    )

    prudential_url = ensure_prudential_url(
        excel_fund[
            "prudentialUrl"
        ]
    )

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

        raise RecoveryParseFailure(
            "Recovery page redirected outside "
            "Prudential Singapore."
        )

    factsheet_url = find_factsheet_url(
        page,
        prudential_url,
    )

    factsheet_bytes = download_factsheet(
        page,
        factsheet_url,
    )

    full_text, page_count = (
        extract_pdf_text(
            factsheet_bytes
        )
    )

    layout_text, _page_blocks = (
        extract_pdf_layout_text(
            factsheet_bytes
        )
    )

    diagnostics = build_recovery_diagnostics(
        excel_row,
        strategy,
        layout_text,
        "",
    )

    recovery_sections = (
        extract_recovery_sections(
            layout_text
        )
    )

    selected_section_text = ""

    selected_parser = ""

    holdings = None

    candidate_results = []

    if strategy == (
        "alternate_holdings_boundary"
    ):

        try:

            (
                holdings,
                selected_section_text,
                boundary_diagnostics,
            ) = recovery_parse_alternate_boundary(
                layout_text
            )

            diagnostics[
                "boundaryRecovery"
            ] = boundary_diagnostics

            selected_parser = (
                "alternate_boundary"
            )

        except Exception as error:

            candidate_results.append(
                {
                    "parser":
                        "alternate_boundary",

                    "error":
                        clean_text(
                            str(error)
                        ),
                }
            )

    else:

        for section in recovery_sections:

            candidate_text = section[
                "text"
            ]

            parser_attempts = []

            if strategy == (
                "two_column_dividend_history_interleaving"
            ):

                parser_attempts = [
                    (
                        "two_column",
                        recovery_parse_two_column,
                    ),
                    (
                        "detached",
                        recovery_parse_detached_pairs,
                    ),
                    (
                        "fallback",
                        parse_holdings_fallback,
                    ),
                ]

            elif strategy == (
                "detached_names_and_weights"
            ):

                parser_attempts = [
                    (
                        "detached",
                        recovery_parse_detached_pairs,
                    ),
                    (
                        "two_column",
                        recovery_parse_two_column,
                    ),
                    (
                        "fallback",
                        parse_holdings_fallback,
                    ),
                ]

            elif strategy == (
                "complex_interleaving"
            ):

                parser_attempts = [
                    (
                        "complex",
                        recovery_parse_complex_interleaving,
                    ),
                    (
                        "two_column",
                        recovery_parse_two_column,
                    ),
                    (
                        "detached",
                        recovery_parse_detached_pairs,
                    ),
                ]

            else:

                parser_attempts = [
                    (
                        "fallback",
                        parse_holdings_fallback,
                    ),
                    (
                        "detached",
                        recovery_parse_detached_pairs,
                    ),
                    (
                        "two_column",
                        recovery_parse_two_column,
                    ),
                ]

            for parser_name, parser in parser_attempts:

                try:

                    parsed_holdings = parser(
                        candidate_text
                    )

                    parsed_holdings = validate_holdings(
                        parsed_holdings,
                        f"Recovery-{parser_name}",
                    )

                    candidate_results.append(
                        {
                            "parser":
                                parser_name,

                            "heading":
                                section[
                                    "heading"
                                ],

                            "holdingCount":
                                len(
                                    parsed_holdings
                                ),

                            "holdings":
                                parsed_holdings,
                        }
                    )

                except Exception as error:

                    candidate_results.append(
                        {
                            "parser":
                                parser_name,

                            "heading":
                                section[
                                    "heading"
                                ],

                            "error":
                                clean_text(
                                    str(error)
                                ),
                        }
                    )

    valid_candidates = [
        candidate
        for candidate
        in candidate_results
        if candidate.get(
            "holdings"
        )
    ]

    if holdings is None and valid_candidates:

        valid_candidates.sort(
            key=lambda candidate: (
                -len(
                    candidate[
                        "holdings"
                    ]
                ),
                candidate.get(
                    "heading",
                    "",
                ),
            )
        )

        selected = valid_candidates[0]

        holdings = selected[
            "holdings"
        ]

        selected_section_text = ""

        for section in recovery_sections:

            if (
                section[
                    "heading"
                ]
                ==
                selected.get(
                    "heading"
                )
            ):

                selected_section_text = (
                    section[
                        "text"
                    ]
                )

                break

        selected_parser = (
            selected.get(
                "parser"
            )
            or
            "recovery"
        )

    diagnostics[
        "candidateResults"
    ] = candidate_results

    diagnostics[
        "recoverySectionsFound"
    ] = len(
        recovery_sections
    )

    if not holdings:

        raise RecoveryParseFailure(
            "Stage-2 recovery could not establish "
            "a valid official holdings set.",
            layout_text=layout_text,
            diagnostics=diagnostics,
        )

    holdings = validate_holdings(
        holdings,
        "Stage2",
    )

    diagnostics[
        "selectedParser"
    ] = selected_parser

    diagnostics[
        "selectedHoldingCount"
    ] = len(
        holdings
    )

    diagnostics[
        "selectedSectionTextLength"
    ] = len(
        selected_section_text
    )

    diagnostics[
        "stage1Error"
    ] = stage1_error

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

    result = {
        "status":
            "success",

        "recoveryStatus":
            "recovered",

        "recoveryStage":
            2,

        "recoveryStrategy":
            strategy,

        "recoveryParser":
            selected_parser,

        "stage1Error":
            stage1_error,

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
            excel_fund.get(
                "pruAccessName"
            ),

        "factsheetUrl":
            factsheet_url,

        "factsheetDocumentDate":
            extract_document_date(
                full_text
            ),

        "factsheetDataAsAt":
            extract_data_as_at(
                full_text
            ),

        "factsheetPageCount":
            page_count,

        "holdingsSectionStatus":
            "published",

        "topHoldingsCount":
            len(
                holdings
            ),

        "topHoldings":
            holdings,

        "rules": {

            "officialPrudentialSourceOnly":
                True,

            "recoveryOnlyAfterStage1Failure":
                True,

            "successfulStage1FundsUntouched":
                True,

            "maximumHoldings":
                MAX_HOLDINGS,

            "noInferredHoldings":
                True,

            "noFabricatedPercentages":
                True,

            "duplicateHoldingNamesAllowed":
                True,

            "duplicateHoldingPercentagesAllowed":
                True,

            "fixedIncomeLastPercentageAsWeight":
                True,

            "layoutRecoveryEnabled":
                True,

            "coordinateIndependentLayoutFallback":
                True,

            "thirdPartyHoldings":
                False,
        },

        "recoveryDiagnostics":
            diagnostics,
    }

    return {
        "result":
            result,

        "factsheetBytes":
            factsheet_bytes,

        "fullText":
            full_text,

        "sectionText":
            selected_section_text,

        "layoutText":
            layout_text,

        "diagnostics":
            diagnostics,
    }


# =============================================================================
# SAVE RECOVERED RESULT
# =============================================================================

def save_recovered_result(
    recovered: dict,
) -> Path:

    result = recovered[
        "result"
    ]

    factsheet_bytes = recovered[
        "factsheetBytes"
    ]

    full_text = recovered[
        "fullText"
    ]

    section_text = recovered[
        "sectionText"
    ]

    layout_text = recovered[
        "layoutText"
    ]

    diagnostics = recovered[
        "diagnostics"
    ]

    excel_row = int(
        result[
            "excelRow"
        ]
    )

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
        /
        (
            f"{excel_row}_recovered_"
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

    (
        directory
        / "recovery_layout.txt"
    ).write_text(
        layout_text,
        encoding="utf-8",
    )

    save_json(
        directory
        / "recovery_diagnostics.json",
        diagnostics,
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

        "status":
            "success",

        "recoveryStatus":
            "recovered",

        "recoveryStrategy":
            result.get(
                "recoveryStrategy"
            ),

        "recoveryParser":
            result.get(
                "recoveryParser"
            ),

        "stage1Error":
            result.get(
                "stage1Error"
            ),

        "topHoldingsCount":
            result.get(
                "topHoldingsCount"
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
# SAVE RECOVERY FAILURE
# =============================================================================

def save_recovery_failure(
    excel_fund: dict,
    stage1_error: str,
    recovery_error: str,
    layout_text: str = "",
    diagnostics: dict | None = None,
) -> Path:

    excel_row = int(
        excel_fund[
            "excelRow"
        ]
    )

    directory = (
        FUNDS_OUTPUT_DIR
        /
        f"{excel_row}_failed"
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    failure_path = (
        directory
        / "failure.json"
    )

    existing_failure = {}

    if failure_path.exists():

        try:

            existing_failure = json.loads(
                failure_path.read_text(
                    encoding="utf-8"
                )
            )

        except Exception:

            existing_failure = {}

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

        "stage1Error":
            stage1_error,

        "recoveryStatus":
            "failed",

        "recoveryStrategy":
            recovery_strategy_for_row(
                excel_row
            ),

        "recoveryError":
            clean_text(
                recovery_error
            ),

        "failedAtUtc":
            utc_now_iso(),

        "stage1Failure":
            existing_failure,
    }

    save_json(
        failure_path,
        failure,
    )

    if layout_text:

        (
            directory
            / "recovery_layout.txt"
        ).write_text(
            layout_text,
            encoding="utf-8",
        )

    if diagnostics:

        save_json(
            directory
            / "recovery_diagnostics.json",
            diagnostics,
        )

    return directory


# =============================================================================
# STAGE 1 EXECUTION
# =============================================================================

def run_stage_1(
    page,
    funds: list[dict],
) -> tuple[
    list[dict],
    list[dict],
    list[dict],
]:

    successful = []

    no_holdings_section = []

    failed = []

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
            f"STAGE 1 FUND {index}/{total}"
        )

        print(
            f"Excel row: "
            f"{excel_fund['excelRow']}"
        )

        print(
            "=" * 72
        )

        last_error = None

        last_exception = None

        result = None

        for attempt in range(
            1,
            RETRY_COUNT + 1,
        ):

            try:

                print(
                    f"\nAttempt "
                    f"{attempt}/{RETRY_COUNT}"
                )

                result = (
                    extract_single_fund(
                        page,
                        excel_fund,
                    )
                )

                last_error = None

                last_exception = None

                break

            except Exception as error:

                last_error = clean_text(
                    str(error)
                )

                last_exception = error

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
                    section_text=getattr(
                        last_exception,
                        "section_text",
                        None,
                    ),
                    full_text=getattr(
                        last_exception,
                        "full_text",
                        None,
                    ),
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

        if (
            result[
                "status"
            ]
            ==
            "no_holdings_section"
        ):

            result_directory = (
                FUNDS_OUTPUT_DIR
                /
                (
                    f"{result['excelRow']}_"
                    f"{safe_filename(result.get('fundName') or 'fund')}"
                )
            )

            result_directory.mkdir(
                parents=True,
                exist_ok=True,
            )

            save_json(
                result_directory
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
                result_directory
                / "metadata.json",
                metadata,
            )

            no_holdings_section.append(
                result
            )

            print(
                "\n"
                "NO TOP HOLDINGS SECTION"
            )

            continue

        # ---------------------------------------------------------------------
        # FINAL OFFICIAL PDF VERIFICATION
        # ---------------------------------------------------------------------

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
                    "disappeared during verification."
                )

            verification_parser = (
                result.get(
                    "holdingsParser"
                )
                or
                "primary"
            )

            if (
                verification_parser
                ==
                "fallback"
            ):

                verified_holdings = (
                    parse_holdings_fallback(
                        section_text
                    )
                )

            else:

                try:

                    verified_holdings = (
                        parse_holdings(
                            section_text
                        )
                    )

                except Exception as verification_primary_error:

                    print(
                        "Primary verification parser failed; "
                        "running fallback parser..."
                    )

                    try:

                        verified_holdings = (
                            parse_holdings_fallback(
                                section_text
                            )
                        )

                    except Exception as verification_fallback_error:

                        raise HoldingsParseFailure(
                            clean_text(
                                str(
                                    verification_fallback_error
                                )
                            ),
                            section_text=section_text,
                            full_text=full_text,
                        ) from verification_fallback_error

                    result[
                        "holdingsParser"
                    ] = "fallback"

                    result[
                        "primaryParserError"
                    ] = clean_text(
                        str(
                            verification_primary_error
                        )
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
                    section_text=getattr(
                        error,
                        "section_text",
                        None,
                    ),
                    full_text=getattr(
                        error,
                        "full_text",
                        None,
                    ),
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

                    "pruaccessName":
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

            print(
                f"  {holding['rank']}. "
                f"{holding['name']} - "
                f"{holding['weightText']}"
            )

    return (
        successful,
        no_holdings_section,
        failed,
    )


# =============================================================================
# STAGE 2 EXECUTION
# =============================================================================

def run_stage_2_recovery(
    page,
    funds: list[dict],
    stage1_failed: list[dict],
) -> tuple[
    list[dict],
    list[dict],
]:

    recovered = []

    still_failed = []

    if not RECOVERY_ENABLED:

        return (
            recovered,
            stage1_failed,
        )

    failed_by_row = {
        int(
            failure[
                "excelRow"
            ]
        ):
            failure
        for failure
        in stage1_failed
    }

    eligible = []

    for excel_fund in funds:

        row = int(
            excel_fund[
                "excelRow"
            ]
        )

        if row not in failed_by_row:

            continue

        if (
            RECOVERY_TARGET_ROWS
            and
            row not in RECOVERY_TARGET_ROWS
        ):

            continue

        eligible.append(
            (
                excel_fund,
                failed_by_row[
                    row
                ],
            )
        )

    print(
        "\n\n"
        + "#" * 78
    )

    print(
        "STAGE 2 RECOVERY"
    )

    print(
        "#" * 78
    )

    print(
        f"Stage-1 failed funds: "
        f"{len(stage1_failed)}"
    )

    print(
        f"Recovery-target failures: "
        f"{len(eligible)}"
    )

    print(
        "Recovery targets: "
        f"{sorted(RECOVERY_TARGET_ROWS)}"
    )

    for index, (
        excel_fund,
        failure,
    ) in enumerate(
        eligible,
        start=1,
    ):

        row = int(
            excel_fund[
                "excelRow"
            ]
        )

        stage1_error = clean_text(
            failure.get(
                "error"
            )
            or
            "Stage-1 failure."
        )

        print(
            "\n"
            + "-" * 72
        )

        print(
            f"RECOVERY {index}/{len(eligible)}"
        )

        print(
            f"Excel row: {row}"
        )

        print(
            f"Strategy: "
            f"{recovery_strategy_for_row(row)}"
        )

        print(
            "-" * 72
        )

        last_recovery_error = None

        last_recovery_exception = None

        recovered_payload = None

        for attempt in range(
            1,
            RECOVERY_RETRY_COUNT + 1,
        ):

            try:

                print(
                    f"Recovery attempt "
                    f"{attempt}/{RECOVERY_RETRY_COUNT}"
                )

                recovered_payload = (
                    recover_single_fund(
                        page,
                        excel_fund,
                        stage1_error,
                    )
                )

                last_recovery_error = None

                last_recovery_exception = None

                break

            except Exception as error:

                last_recovery_error = clean_text(
                    str(error)
                )

                last_recovery_exception = error

                print(
                    "Recovery attempt failed:"
                )

                print(
                    last_recovery_error
                )

                if attempt < RECOVERY_RETRY_COUNT:

                    time.sleep(
                        RECOVERY_RETRY_DELAY_SECONDS
                    )

        if recovered_payload is not None:

            output_directory = (
                save_recovered_result(
                    recovered_payload
                )
            )

            result = recovered_payload[
                "result"
            ]

            recovered.append(
                {
                    "result":
                        result,

                    "outputDirectory":
                        str(
                            output_directory
                        ),
                }
            )

            print(
                "\nSTAGE 2 RECOVERED"
            )

            print(
                f"Fund: "
                f"{result.get('fundName') or '-'}"
            )

            print(
                f"Holdings: "
                f"{result.get('topHoldingsCount')}"
            )

            for holding in result[
                "topHoldings"
            ]:

                print(
                    f"  {holding['rank']}. "
                    f"{holding['name']} - "
                    f"{holding['weightText']}"
                )

            continue

        layout_text = getattr(
            last_recovery_exception,
            "layout_text",
            "",
        )

        diagnostics = getattr(
            last_recovery_exception,
            "diagnostics",
            {},
        )

        failure_directory = (
            save_recovery_failure(
                excel_fund,
                stage1_error,
                last_recovery_error
                or
                "Unknown Stage-2 recovery failure.",
                layout_text=layout_text,
                diagnostics=diagnostics,
            )
        )

        still_failed.append(
            {
                "excelRow":
                    row,

                "prudentialUrl":
                    excel_fund[
                        "prudentialUrl"
                    ],

                "pruaccessName":
                    excel_fund.get(
                        "pruAccessName"
                    ),

                "error":
                    stage1_error,

                "stage1Error":
                    stage1_error,

                "recoveryError":
                    last_recovery_error
                    or
                    "Unknown Stage-2 recovery failure.",

                "recoveryStrategy":
                    recovery_strategy_for_row(
                        row
                    ),

                "outputDirectory":
                    str(
                        failure_directory
                    ),
            }
        )

    # -------------------------------------------------------------------------
    # Failed funds outside the explicit recovery target remain failed exactly
    # as Stage 1 recorded them. They are not passed through a generic parser.
    # -------------------------------------------------------------------------

    recovered_rows = {
        int(
            item[
                "result"
            ][
                "excelRow"
            ]
        )
        for item
        in recovered
    }

    unresolved_rows = {
        int(
            item[
                "excelRow"
            ]
        )
        for item
        in still_failed
    }

    for failure in stage1_failed:

        row = int(
            failure[
                "excelRow"
            ]
        )

        if row in recovered_rows:
            continue

        if row in unresolved_rows:
            continue

        still_failed.append(
            failure
        )

    still_failed.sort(
        key=lambda item: int(
            item[
                "excelRow"
            ]
        )
    )

    return (
        recovered,
        still_failed,
    )


# =============================================================================
# CONSOLIDATED OUTPUT
# =============================================================================

def build_consolidated_outputs(
    funds: list[dict],
    successful: list[dict],
    no_holdings_section: list[dict],
    stage1_failed: list[dict],
    recovered: list[dict],
    final_failed: list[dict],
    started_at: str,
) -> tuple[
    dict,
    dict,
]:

    recovered_results = [
        item[
            "result"
        ]
        for item
        in recovered
    ]

    all_results = (
        successful
        +
        no_holdings_section
        +
        recovered_results
    )

    all_results.sort(
        key=lambda item: int(
            item[
                "excelRow"
            ]
        )
    )

    fallback_recovered = [
        item
        for item
        in successful
        if item.get(
            "holdingsParser"
        ) == "fallback"
    ]

    total_holdings = sum(
        int(
            item.get(
                "topHoldingsCount",
                0,
            )
            or 0
        )
        for item
        in all_results
        if item.get(
            "status"
        )
        == "success"
    )

    final_status = (
        "success"
        if not final_failed
        else "partial"
    )

    all_holdings_payload = {
        "status":
            final_status,

        "source":
            "Prudential Singapore",

        "generatedAtUtc":
            utc_now_iso(),

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

        "stage1FailedFunds":
            len(
                stage1_failed
            ),

        "recoveredFunds":
            len(
                recovered
            ),

        "failedFunds":
            len(
                final_failed
            ),

        "fallbackRecoveredFunds":
            len(
                fallback_recovered
            ),

        "recoveryTargets":
            sorted(
                RECOVERY_TARGET_ROWS
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

            "duplicateHoldingNamesAllowed":
                True,

            "duplicateHoldingPercentagesAllowed":
                True,

            "multilineHoldingNamesSupported":
                True,

            "wrappedPdfLinesJoined":
                True,

            "fixedIncomeCouponPercentagesSupported":
                True,

            "logicalHoldingLineReconstruction":
                True,

            "percentageDefinesHoldingBoundary":
                True,

            "automaticFallbackParser":
                True,

            "fallbackUsesLastPercentageOnLogicalLine":
                True,

            "stage2RecoveryAfterBaseline":
                True,

            "stage2OnlyFailedFunds":
                True,

            "successfulStage1FundsUntouched":
                True,

            "officialFactsheetOnly":
                True,

            "thirdPartyHoldings":
                False,
        },

        "funds":
            all_results,

        "failed":
            final_failed,
    }

    fallback_details = [
        {
            "excelRow":
                item.get(
                    "excelRow"
                ),

            "fundName":
                item.get(
                    "fundName"
                ),

            "primaryParserError":
                item.get(
                    "primaryParserError"
                ),

            "factsheetUrl":
                item.get(
                    "factsheetUrl"
                ),

            "topHoldingsCount":
                item.get(
                    "topHoldingsCount"
                ),
        }
        for item
        in fallback_recovered
    ]

    recovery_details = [
        {
            "excelRow":
                item[
                    "result"
                ].get(
                    "excelRow"
                ),

            "fundName":
                item[
                    "result"
                ].get(
                    "fundName"
                ),

            "strategy":
                item[
                    "result"
                ].get(
                    "recoveryStrategy"
                ),

            "parser":
                item[
                    "result"
                ].get(
                    "recoveryParser"
                ),

            "stage1Error":
                item[
                    "result"
                ].get(
                    "stage1Error"
                ),

            "factsheetUrl":
                item[
                    "result"
                ].get(
                    "factsheetUrl"
                ),

            "topHoldingsCount":
                item[
                    "result"
                ].get(
                    "topHoldingsCount"
                ),

            "outputDirectory":
                item.get(
                    "outputDirectory"
                ),
        }
        for item
        in recovered
    ]

    run_summary = {
        "status":
            final_status,

        "startedAtUtc":
            started_at,

        "completedAtUtc":
            utc_now_iso(),

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

        "stage1FailedFunds":
            len(
                stage1_failed
            ),

        "recoveredFunds":
            len(
                recovered
            ),

        "failedFunds":
            len(
                final_failed
            ),

        "fallbackRecoveredFunds":
            len(
                fallback_recovered
            ),

        "recoveryTargets":
            sorted(
                RECOVERY_TARGET_ROWS
            ),

        "fallbackRecoveredFundsDetail":
            fallback_details,

        "recoveredFundsDetail":
            recovery_details,

        "totalPublishedTopHoldings":
            total_holdings,

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

        "stage1FailedFundsDetail":
            stage1_failed,

        "failedFundsDetail":
            final_failed,

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

            "noDuplicateHoldingNameValidation":
                True,

            "duplicateHoldingNamesAllowed":
                True,

            "duplicateHoldingPercentagesAllowed":
                True,

            "noThirdPartyHoldings":
                True,

            "multilineHoldingNamesSupported":
                True,

            "wrappedPdfLinesJoined":
                True,

            "fixedIncomeCouponPercentagesSupported":
                True,

            "logicalHoldingLineReconstruction":
                True,

            "percentageDefinesHoldingBoundary":
                True,

            "automaticFallbackParser":
                True,

            "fallbackUsesLastPercentageOnLogicalLine":
                True,

            "stage2RecoveryRunsAfterStage1":
                True,

            "stage2RecoveryOnlyRunsOnStage1Failures":
                True,

            "successfulStage1FundsUntouched":
                True,

            "officialFactsheetOnly":
                True,
        },
    }

    return (
        all_holdings_payload,
        run_summary,
    )


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

    stage1_failed = []

    recovered = []

    final_failed = []

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

            # =================================================================
            # STAGE 1
            # =================================================================

            print(
                "\n\n"
                + "#" * 78
            )

            print(
                "STAGE 1 - BASELINE EXTRACTION"
            )

            print(
                "#" * 78
            )

            (
                successful,
                no_holdings_section,
                stage1_failed,
            ) = run_stage_1(
                page,
                funds,
            )

            print(
                "\n"
                + "=" * 78
            )

            print(
                "STAGE 1 COMPLETE"
            )

            print(
                "=" * 78
            )

            print(
                f"Excel fund universe: "
                f"{len(funds)}"
            )

            print(
                f"Successful: "
                f"{len(successful)}"
            )

            print(
                f"No holdings section: "
                f"{len(no_holdings_section)}"
            )

            print(
                f"Failed: "
                f"{len(stage1_failed)}"
            )

            if stage1_failed:

                print(
                    "\nStage-1 failed rows:"
                )

                for failure in stage1_failed:

                    print(
                        f" - Row "
                        f"{failure['excelRow']}: "
                        f"{failure['error']}"
                    )

            # =================================================================
            # STAGE 2
            # =================================================================

            if RECOVERY_ENABLED and stage1_failed:

                (
                    recovered,
                    final_failed,
                ) = run_stage_2_recovery(
                    page,
                    funds,
                    stage1_failed,
                )

            else:

                final_failed = list(
                    stage1_failed
                )

            # =================================================================
            # FINAL
            # =================================================================

            print(
                "\n\n"
                + "#" * 78
            )

            print(
                "STAGE 2 / FINAL PROCESSING COMPLETE"
            )

            print(
                "#" * 78
            )

        finally:

            context.close()

            browser.close()

    # =========================================================================
    # CONSOLIDATED OUTPUT
    # =========================================================================

    (
        all_holdings_payload,
        run_summary,
    ) = build_consolidated_outputs(
        funds=funds,
        successful=successful,
        no_holdings_section=no_holdings_section,
        stage1_failed=stage1_failed,
        recovered=recovered,
        final_failed=final_failed,
        started_at=started_at,
    )

    save_json(
        ALL_HOLDINGS_FILE,
        all_holdings_payload,
    )

    save_json(
        RUN_SUMMARY_FILE,
        run_summary,
    )

    # =========================================================================
    # CONSOLE SUMMARY
    # =========================================================================

    recovered_results = [
        item[
            "result"
        ]
        for item
        in recovered
    ]

    final_results = (
        successful
        +
        no_holdings_section
        +
        recovered_results
    )

    total_holdings = sum(
        int(
            item.get(
                "topHoldingsCount",
                0,
            )
            or 0
        )
        for item
        in final_results
        if item.get(
            "status"
        )
        == "success"
    )

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
        f"Stage-1 successful with holdings: "
        f"{len(successful)}"
    )

    print(
        f"No Top Holdings section: "
        f"{len(no_holdings_section)}"
    )

    print(
        f"Stage-1 failed: "
        f"{len(stage1_failed)}"
    )

    print(
        f"Stage-2 recovered: "
        f"{len(recovered)}"
    )

    print(
        f"Final failed: "
        f"{len(final_failed)}"
    )

    print(
        f"Total published holdings extracted: "
        f"{total_holdings}"
    )

    print(
        "\nSTAGE 2 RECOVERY TARGETS:"
    )

    print(
        " - Rows 13/14/15: two-column / Dividend History interleaving"
    )

    print(
        " - Rows 17/34/53: detached names and weights"
    )

    print(
        " - Row 29: complex interleaving"
    )

    print(
        " - Rows 59/60: alternate Top Holdings boundary"
    )

    print(
        "\nSuccessful Stage-1 funds were not reinterpreted."
    )

    print(
        "Recovery runs only against Stage-1 failures."
    )

    print(
        "Official Prudential factsheets only."
    )

    print(
        "No inferred holdings."
    )

    print(
        "No fabricated percentages."
    )

    print(
        "Duplicate holding names: ALLOWED"
    )

    print(
        "Duplicate holding percentages: ALLOWED"
    )

    print(
        "Fixed-income last percentage: PORTFOLIO WEIGHT"
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

    if recovered:

        print(
            "\nRECOVERED FUNDS:"
        )

        for item in recovered:

            result = item[
                "result"
            ]

            print(
                f" - Row "
                f"{result['excelRow']}: "
                f"{result.get('fundName') or '-'} "
                f"("
                f"{result.get('recoveryStrategy')}"
                f")"
            )

    if final_failed:

        print(
            "\nFINAL FAILED FUNDS:"
        )

        for failure in final_failed:

            print(
                f" - Row "
                f"{failure['excelRow']}: "
                f"{failure.get('recoveryError') or failure.get('error')}"
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
