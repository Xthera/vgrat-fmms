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

IMPORTANT DESIGN PRINCIPLE
==========================

This extractor is deliberately CONSERVATIVE.

Wrong holding data is worse than missing holding data.

Therefore:

- Official Prudential factsheet only.
- No third-party holdings.
- No inferred holdings.
- No fabricated holdings.
- No fabricated percentages.
- No calculated percentages.
- No forced 10 holdings.
- Multi-line holding names are supported.
- Holdings without published percentages are allowed.
- If the parser cannot establish that a name belongs to the
  Top Holdings table, it is NOT accepted.
- If the parser cannot reliably identify holdings, the fund FAILS.
- Raw PDF text and the detected Top Holdings section are retained
  for auditing.

The extractor supports common Prudential PDF layouts including:

    1. Rank + name + percentage on one line

    2. Rank + multi-line name + percentage

    3. Rank on separate line
       name on one or more lines
       percentage on separate line

    4. Name + percentage separated by PDF extraction

    5. Published holding names without percentages

    6. PDFs where pypdf produces different reading orders

It deliberately DOES NOT try to "guess" a holding when the PDF
structure is ambiguous.


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
- If Prudential publishes fewer than 10 holdings, store exactly
  that published count.
- Holdings are stored in published order whenever that order can
  be established.
- Multi-line names are joined only when they belong to the same
  detected holding row.
- Holding percentage may be null when Prudential does not publish it.
- No percentage is calculated.
- No percentage is inferred.
- If a Top Holdings section exists but no reliable holding rows can
  be established, the fund is FAILED.
- If a Top Holdings section exists and reliable names are published
  without percentages, the fund is SUCCESS with null weights.
- If no Top Holdings section exists, status is NO_HOLDINGS_SECTION.
- Duplicate ranks are never silently deduplicated.
- Duplicate names at different published ranks are retained and
  flagged as a warning.
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
# OFFICIAL PRUDENTIAL HOSTS
# =============================================================================

PRUDENTIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}


# =============================================================================
# GENERAL HELPERS
# =============================================================================

def clean_text(value) -> str:

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


def normalize_text(value) -> str:

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


def safe_filename(value) -> str:

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
                "excelRow": row_number,
                "prudentialUrl": prudential_url,
                "pruAccessName": pruaccess_name,
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
# FACTSHEET LINK
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

        if "fund factsheet" in text_lower:
            score += 200

        elif "factsheet" in text_lower:
            score += 150

        if "factsheet" in href_lower:
            score += 100

        if href_lower.endswith(".pdf"):
            score += 50

        if score > 0:

            candidates.append(
                {
                    "score": score,
                    "url": absolute_url,
                    "text": anchor_text,
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
# PDF EXTRACTION
# =============================================================================

def extract_pdf_text_variants(
    pdf_bytes: bytes,
) -> tuple[list[dict], int]:

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

    variants = [
        {
            "name":
                "default",

            "pages":
                [],
        }
    ]

    has_layout_mode = True

    try:
        # Check whether the installed pypdf supports extraction_mode.
        reader.pages[0].extract_text(
            extraction_mode="layout"
        )

    except TypeError:

        has_layout_mode = False

    except Exception:

        # Layout extraction exists but may not work on this PDF.
        has_layout_mode = True

    if has_layout_mode:

        variants.append(
            {
                "name":
                    "layout",

                "pages":
                    [],
            }
        )

    for page_index, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):

        try:

            default_text = (
                pdf_page.extract_text()
                or ""
            )

        except Exception as error:

            raise RuntimeError(
                "Failed to extract PDF text "
                f"from page {page_index}: {error}"
            ) from error

        variants[0][
            "pages"
        ].append(
            default_text
        )

        if has_layout_mode:

            try:

                layout_text = (
                    pdf_page.extract_text(
                        extraction_mode="layout"
                    )
                    or ""
                )

            except Exception:

                layout_text = ""

            variants[1][
                "pages"
            ].append(
                layout_text
            )

    output = []

    for variant in variants:

        full_text = "\n".join(
            variant["pages"]
        )

        if clean_text(
            full_text
        ):

            output.append(
                {
                    "name":
                        variant["name"],

                    "text":
                        full_text,
                }
            )

    if not output:

        raise RuntimeError(
            "Factsheet PDF contains no extractable text."
        )

    return (
        output,
        page_count,
    )


def extract_pdf_text(
    pdf_bytes: bytes,
) -> tuple[str, int]:

    variants, page_count = (
        extract_pdf_text_variants(
            pdf_bytes
        )
    )

    return (
        variants[0]["text"],
        page_count,
    )


# =============================================================================
# PDF LINES
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

    pattern = re.compile(
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

        match = pattern.search(
            line
        )

        if match:

            return clean_text(
                match.group(0)
            )

    return None


# =============================================================================
# HOLDINGS SECTION DETECTION
# =============================================================================

HOLDINGS_SECTION_PATTERNS = [
    re.compile(
        r"^\s*top\s+10\s+holdings\s*$",
        re.IGNORECASE,
    ),
    re.compile(
        r"^\s*top\s+ten\s+holdings\s*$",
        re.IGNORECASE,
    ),
    re.compile(
        r"^\s*top\s+holdings\s*$",
        re.IGNORECASE,
    ),
]


def is_exact_holdings_heading(
    line: str,
) -> bool:

    normalized = normalize_text(
        line
    )

    for pattern in HOLDINGS_SECTION_PATTERNS:

        if pattern.fullmatch(
            normalized
        ):

            return True

    return False


def find_holdings_start(
    lines: list[str],
) -> int | None:

    for index, line in enumerate(
        lines
    ):

        if is_exact_holdings_heading(
            line
        ):

            return index

    return None


def is_holdings_end(
    line: str,
) -> bool:

    normalized = normalize_text(
        line
    )

    exact_endings = {
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
        "fund performance",
        "fund information",
    }

    if normalized in exact_endings:
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

    if normalized.startswith(
        "disclaimer"
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

    section_text = "\n".join(
        section
    )

    if not clean_text(
        section_text
    ):

        return (
            "",
            "published_empty",
        )

    return (
        section_text,
        "published",
    )


# =============================================================================
# HOLDINGS TABLE HEADERS
# =============================================================================

HEADER_WORDS = {
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


def looks_like_holdings_header(
    line: str,
) -> bool:

    normalized = normalize_text(
        line
    )

    if normalized in HEADER_WORDS:
        return True

    if (
        "holding" in normalized
        and (
            "weight" in normalized
            or "percentage" in normalized
            or "%" in normalized
        )
    ):
        return True

    return False


# =============================================================================
# RANK DETECTION
# =============================================================================

RANK_RE = re.compile(
    r"""
    ^
    (?:
        rank
        \s*
    )?
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
    re.IGNORECASE
    | re.VERBOSE,
)


def extract_rank(
    line: str,
) -> tuple[int | None, str]:

    original = clean_text(
        line
    )

    if not original:
        return (
            None,
            "",
        )

    match = RANK_RE.match(
        original
    )

    if not match:

        return (
            None,
            original,
        )

    rank = int(
        match.group(1)
    )

    if (
        rank < 1
        or rank > MAX_HOLDINGS
    ):

        return (
            None,
            original,
        )

    remainder = clean_text(
        match.group(2)
    )

    return (
        rank,
        remainder,
    )


# =============================================================================
# PERCENTAGE
# =============================================================================

PERCENT_RE = re.compile(
    r"""
    (?<![\d.])
    ([0-9]+(?:\.[0-9]+)?)
    \s*%
    """,
    re.VERBOSE,
)


def extract_percentage_from_line(
    value: str,
) -> tuple[float | None, str, str | None]:

    text = clean_text(
        value
    )

    if not text:

        return (
            None,
            "",
            None,
        )

    matches = list(
        PERCENT_RE.finditer(
            text
        )
    )

    if not matches:

        return (
            None,
            text,
            None,
        )

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
            None,
        )

    raw_percentage = (
        match.group(0)
        .strip()
    )

    remaining = clean_text(
        text[:match.start()]
        + " "
        + text[match.end():]
    )

    return (
        percentage,
        remaining,
        raw_percentage,
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

    name = re.sub(
        r"^\d{1,2}[.)]\s+",
        "",
        name,
    )

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
        "weight (%)",
        "percentage",
        "%",
        "rank",
    }:

        return ""

    return name


# =============================================================================
# CANDIDATE NAME VALIDATION
# =============================================================================

def looks_like_number_or_metadata(
    text: str,
) -> bool:

    normalized = normalize_text(
        text
    )

    if not normalized:
        return True

    if re.fullmatch(
        r"[\d.,%+\-/]+",
        normalized,
    ):
        return True

    metadata_patterns = [
        r"^as\s+at\b",
        r"^data\s+as\s+at\b",
        r"^date\b",
        r"^page\s+\d+",
        r"^source\b",
        r"^isin\b",
        r"^sedol\b",
        r"^isin\s*:",
        r"^fund\s+price\b",
        r"^fund\s+size\b",
        r"^nav\b",
        r"^currency\b",
        r"^asset\s+class\b",
        r"^risk\b",
        r"^inception\b",
        r"^performance\b",
        r"^return\b",
        r"^returns\b",
        r"^important\s+information\b",
        r"^disclaimer\b",
    ]

    for pattern in metadata_patterns:

        if re.search(
            pattern,
            normalized,
        ):
            return True

    return False


def is_plausible_holding_name(
    name: str,
) -> bool:

    name = clean_holding_name(
        name
    )

    if not name:
        return False

    if looks_like_number_or_metadata(
        name
    ):
        return False

    # A holding name should contain at least one alphabetic character.
    if not re.search(
        r"[A-Za-z]",
        name,
    ):
        return False

    # Avoid obvious table/document fragments.
    normalized = normalize_text(
        name
    )

    bad_exact = {
        "none",
        "n/a",
        "na",
        "not available",
        "not applicable",
        "total",
        "total holdings",
        "weight",
        "weights",
        "portfolio",
        "portfolio characteristics",
        "asset allocation",
        "sector allocation",
        "country allocation",
        "currency allocation",
        "geographical allocation",
    }

    if normalized in bad_exact:
        return False

    # A single very short word is usually a PDF header, not a security.
    if len(name) < 3:
        return False

    return True


# =============================================================================
# ROW / BLOCK UTILITIES
# =============================================================================

def extract_percentage_from_block(
    lines: list[str],
) -> tuple[
    float | None,
    str | None,
    list[str],
]:

    percentage = None
    percentage_text = None
    remaining_lines = []

    for line in lines:

        line_percentage, remaining, raw = (
            extract_percentage_from_line(
                line
            )
        )

        if (
            percentage is None
            and line_percentage is not None
        ):

            percentage = line_percentage
            percentage_text = raw

        if clean_text(
            remaining
        ):

            remaining_lines.append(
                clean_text(
                    remaining
                )
            )

    return (
        percentage,
        percentage_text,
        remaining_lines,
    )


def make_holding(
    rank: int,
    name_lines: list[str],
    percentage: float | None,
    percentage_text: str | None,
) -> dict | None:

    cleaned_parts = []

    for line in name_lines:

        candidate = clean_holding_name(
            line
        )

        if not candidate:
            continue

        if looks_like_holdings_header(
            candidate
        ):
            continue

        if looks_like_number_or_metadata(
            candidate
        ):
            continue

        cleaned_parts.append(
            candidate
        )

    if not cleaned_parts:
        return None

    name = clean_holding_name(
        " ".join(
            cleaned_parts
        )
    )

    if not is_plausible_holding_name(
        name
    ):
        return None

    return {
        "rank":
            rank,

        "name":
            name,

        "weightPercent":
            percentage,

        "weightText":
            percentage_text,
    }


# =============================================================================
# RANKED BLOCK PARSER
# =============================================================================

def parse_ranked_candidate(
    lines: list[str],
) -> tuple[list[dict], dict]:

    blocks = []

    current = None

    for line in lines:

        line = clean_text(
            line
        )

        if not line:
            continue

        if is_holdings_end(
            line
        ):
            break

        if looks_like_holdings_header(
            line
        ):
            continue

        rank, remainder = extract_rank(
            line
        )

        if rank is not None:

            if current is not None:

                blocks.append(
                    current
                )

            current = {
                "rank":
                    rank,

                "lines":
                    [],
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

    holdings = []

    diagnostics = {
        "parser":
            "ranked_block_parser",

        "rankedBlockCount":
            len(blocks),

        "discardedBlocks":
            [],

        "orphanPercentages":
            0,
    }

    # -------------------------------------------------------------------------
    # Critical protection:
    #
    # A ranked table should normally start at rank 1.
    # If the first detected rank is something else, it is much more likely
    # that unrelated numbered PDF content has been captured.
    # -------------------------------------------------------------------------

    if not blocks:

        return (
            [],
            diagnostics,
        )

    ranks = [
        block["rank"]
        for block
        in blocks
    ]

    if 1 not in ranks:

        diagnostics[
            "reason"
        ] = (
            "Ranked content did not contain "
            "a rank-1 row."
        )

        return (
            [],
            diagnostics,
        )

    # -------------------------------------------------------------------------
    # Only use the first rank-1 occurrence as the beginning of the table.
    # -------------------------------------------------------------------------

    first_rank_one_index = ranks.index(
        1
    )

    blocks = blocks[
        first_rank_one_index:
    ]

    previous_rank = None

    for block in blocks:

        rank = block[
            "rank"
        ]

        # A rank going backwards is strong evidence that PDF columns or
        # another table have been mixed into the section.
        if (
            previous_rank is not None
            and rank <= previous_rank
        ):

            diagnostics[
                "discardedBlocks"
            ].append(
                {
                    "rank":
                        rank,

                    "reason":
                        "rank_not_increasing",

                    "lines":
                        block["lines"],
                }
            )

            continue

        previous_rank = rank

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

        if not block_lines:

            diagnostics[
                "discardedBlocks"
            ].append(
                {
                    "rank":
                        rank,

                    "reason":
                        "empty_block",
                }
            )

            continue

        percentage, percentage_text, name_lines = (
            extract_percentage_from_block(
                block_lines
            )
        )

        holding = make_holding(
            rank,
            name_lines,
            percentage,
            percentage_text,
        )

        if holding is None:

            diagnostics[
                "discardedBlocks"
            ].append(
                {
                    "rank":
                        rank,

                    "reason":
                        "no_plausible_name",

                    "lines":
                        block_lines,
                }
            )

            continue

        holdings.append(
            holding
        )

        if len(
            holdings
        ) >= MAX_HOLDINGS:

            break

    # -------------------------------------------------------------------------
    # Strong structural validation.
    #
    # The detected ranks should correspond to 1..N without unexplained
    # backward jumps. Missing ranks are allowed only when the PDF extraction
    # genuinely omitted a row, but we record that diagnostic.
    # -------------------------------------------------------------------------

    extracted_ranks = [
        item["rank"]
        for item
        in holdings
    ]

    expected_ranks = list(
        range(
            1,
            len(
                extracted_ranks
            ) + 1,
        )
    )

    diagnostics[
        "extractedRanks"
    ] = extracted_ranks

    diagnostics[
        "expectedRanksForCount"
    ] = expected_ranks

    diagnostics[
        "missingRanks"
    ] = [
        rank
        for rank in range(
            1,
            MAX_HOLDINGS + 1,
        )
        if rank in ranks
        and rank not in extracted_ranks
    ]

    diagnostics[
        "publishedPercentages"
    ] = sum(
        1
        for item
        in holdings
        if item.get(
            "weightPercent"
        ) is not None
    )

    diagnostics[
        "namesWithoutPercentage"
    ] = sum(
        1
        for item
        in holdings
        if item.get(
            "weightPercent"
        ) is None
    )

    return (
        holdings,
        diagnostics,
    )


# =============================================================================
# TABLE-STYLE NON-RANKED PARSER
# =============================================================================

def parse_table_rows_candidate(
    lines: list[str],
) -> tuple[list[dict], dict]:

    diagnostics = {
        "parser":
            "table_row_parser",

        "candidateLines":
            len(lines),

        "discardedLines":
            [],

        "percentageLines":
            0,
    }

    rows = []

    # -------------------------------------------------------------------------
    # First attempt:
    #
    # A line containing a percentage is treated as the end of a row.
    # Text immediately preceding it becomes the holding name.
    #
    # This is deliberately conservative. We do NOT combine arbitrary text
    # after a percentage with the next holding.
    # -------------------------------------------------------------------------

    pending_name_parts = []

    for line in lines:

        line = clean_text(
            line
        )

        if not line:
            continue

        if is_holdings_end(
            line
        ):
            break

        if looks_like_holdings_header(
            line
        ):
            continue

        percentage, remaining, percentage_text = (
            extract_percentage_from_line(
                line
            )
        )

        if percentage is not None:

            diagnostics[
                "percentageLines"
            ] += 1

            if remaining:

                pending_name_parts.append(
                    remaining
                )

            if pending_name_parts:

                name = clean_holding_name(
                    " ".join(
                        pending_name_parts
                    )
                )

                if is_plausible_holding_name(
                    name
                ):

                    rows.append(
                        {
                            "rank":
                                len(rows) + 1,

                            "name":
                                name,

                            "weightPercent":
                                percentage,

                            "weightText":
                                percentage_text,
                        }
                    )

                else:

                    diagnostics[
                        "discardedLines"
                    ].append(
                        {
                            "text":
                                name,

                            "reason":
                                "not_plausible_name",
                        }
                    )

            pending_name_parts = []

            if len(
                rows
            ) >= MAX_HOLDINGS:

                break

            continue

        # ---------------------------------------------------------------------
        # Lines without a percentage.
        #
        # These can be name continuations.
        # ---------------------------------------------------------------------

        candidate = clean_holding_name(
            line
        )

        if not candidate:
            continue

        if looks_like_number_or_metadata(
            candidate
        ):
            continue

        pending_name_parts.append(
            candidate
        )

    diagnostics[
        "namesWithoutPercentage"
    ] = 0

    return (
        rows[:MAX_HOLDINGS],
        diagnostics,
    )


# =============================================================================
# NAMES-ONLY PARSER
# =============================================================================

def parse_names_only_candidate(
    lines: list[str],
) -> tuple[list[dict], dict]:

    """
    Used ONLY when the factsheet genuinely publishes holding names without
    percentages.

    This parser is intentionally stricter than the old fallback.

    It requires explicit ranked rows.

    We do NOT guess arbitrary unranked lines as holdings.
    """

    diagnostics = {
        "parser":
            "names_only_ranked_parser",

        "rankedRows":
            0,

        "discardedRows":
            [],
    }

    blocks = []

    current = None

    for line in lines:

        line = clean_text(
            line
        )

        if not line:
            continue

        if is_holdings_end(
            line
        ):
            break

        if looks_like_holdings_header(
            line
        ):
            continue

        rank, remainder = extract_rank(
            line
        )

        if rank is not None:

            if current is not None:

                blocks.append(
                    current
                )

            current = {
                "rank":
                    rank,

                "lines":
                    [],
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

    if not blocks:

        return (
            [],
            diagnostics,
        )

    ranks = [
        block["rank"]
        for block
        in blocks
    ]

    if 1 not in ranks:

        return (
            [],
            diagnostics,
        )

    first = ranks.index(
        1
    )

    blocks = blocks[
        first:
    ]

    previous_rank = None

    for block in blocks:

        rank = block[
            "rank"
        ]

        if (
            previous_rank is not None
            and rank <= previous_rank
        ):
            continue

        previous_rank = rank

        name_parts = []

        for line in block[
            "lines"
        ]:

            candidate = clean_holding_name(
                line
            )

            if not candidate:
                continue

            if looks_like_holdings_header(
                candidate
            ):
                continue

            if looks_like_number_or_metadata(
                candidate
            ):
                continue

            name_parts.append(
                candidate
            )

        name = clean_holding_name(
            " ".join(
                name_parts
            )
        )

        if not is_plausible_holding_name(
            name
        ):
            continue

        # If the block contains an actual percentage, this is not a
        # names-only table; leave percentage parsing to the ranked parser.
        block_percentage_found = any(
            PERCENT_RE.search(
                line
            )
            for line
            in block[
                "lines"
            ]
        )

        if block_percentage_found:
            continue

        rows = {
            "rank":
                rank,

            "name":
                name,

            "weightPercent":
                None,

            "weightText":
                None,
        }

        blocks_count = len(
            blocks
        )

        if blocks_count:
            diagnostics[
                "rankedRows"
            ] += 1

        if diagnostics[
            "rankedRows"
        ] >= MAX_HOLDINGS:
            break

        # Store temporarily.
        if "holdings" not in diagnostics:
            diagnostics[
                "holdings"
            ] = []

        diagnostics[
            "holdings"
        ].append(
            rows
        )

    holdings = diagnostics.pop(
        "holdings",
        [],
    )

    return (
        holdings[:MAX_HOLDINGS],
        diagnostics,
    )


# =============================================================================
# DUPLICATE WARNING
# =============================================================================

def duplicate_name_warnings(
    holdings: list[dict],
) -> list[str]:

    seen = {}

    for holding in holdings:

        normalized = normalize_text(
            holding.get(
                "name"
            )
            or ""
        )

        if not normalized:
            continue

        seen.setdefault(
            normalized,
            [],
        ).append(
            holding.get(
                "rank"
            )
        )

    warnings = []

    for normalized_name, ranks in seen.items():

        if len(
            ranks
        ) > 1:

            warnings.append(
                (
                    f"Duplicate published holding name "
                    f"at ranks {', '.join(map(str, ranks))}: "
                    f"{normalized_name}"
                )
            )

    return warnings


# =============================================================================
# HOLDINGS VALIDATION
# =============================================================================

def validate_holdings(
    holdings: list[dict],
    require_rank_one: bool = True,
) -> dict:

    if not holdings:

        raise RuntimeError(
            "No reliable holding rows could be extracted."
        )

    if len(
        holdings
    ) > MAX_HOLDINGS:

        raise RuntimeError(
            "Parser produced more than 10 holdings."
        )

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
            or rank > MAX_HOLDINGS
        ):

            raise RuntimeError(
                f"Invalid holding rank: {rank}"
            )

        ranks.append(
            rank
        )

        name = clean_holding_name(
            holding.get(
                "name"
            )
            or ""
        )

        if not is_plausible_holding_name(
            name
        ):

            raise RuntimeError(
                "Invalid holding name: "
                f"{name}"
            )

        weight = holding.get(
            "weightPercent"
        )

        if weight is not None:

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
                or weight > 100
            ):

                raise RuntimeError(
                    f"Invalid holding weight: {weight}"
                )

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

    if require_rank_one and 1 not in ranks:

        raise RuntimeError(
            "Reliable Top Holdings table could not be established "
            "because rank 1 was not detected."
        )

    warnings = duplicate_name_warnings(
        holdings
    )

    return {
        "duplicateNameWarnings":
            warnings,

        "publishedPercentageCount":
            sum(
                1
                for holding
                in holdings
                if holding.get(
                    "weightPercent"
                ) is not None
            ),

        "missingPercentageCount":
            sum(
                1
                for holding
                in holdings
                if holding.get(
                    "weightPercent"
                ) is None
            ),
    }


# =============================================================================
# MASTER HOLDINGS PARSER
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

    lines = pdf_lines(
        section_text
    )

    if not lines:

        raise RuntimeError(
            "Top Holdings section contains no usable text."
        )

    candidates = []

    # -------------------------------------------------------------------------
    # Candidate 1: ranked parser
    # -------------------------------------------------------------------------

    ranked_holdings, ranked_info = (
        parse_ranked_candidate(
            lines
        )
    )

    if ranked_holdings:

        try:

            validation = validate_holdings(
                ranked_holdings,
                require_rank_one=True,
            )

            candidates.append(
                {
                    "holdings":
                        ranked_holdings,

                    "parser":
                        ranked_info,

                    "validation":
                        validation,

                    "score":
                        score_candidate(
                            ranked_holdings,
                            ranked_info,
                            validation,
                        ),
                }
            )

        except Exception:
            pass

    # -------------------------------------------------------------------------
    # Candidate 2: table-row parser
    # -------------------------------------------------------------------------

    table_holdings, table_info = (
        parse_table_rows_candidate(
            lines
        )
    )

    if table_holdings:

        try:

            validation = validate_holdings(
                table_holdings,
                require_rank_one=False,
            )

            candidates.append(
                {
                    "holdings":
                        table_holdings,

                    "parser":
                        table_info,

                    "validation":
                        validation,

                    "score":
                        score_candidate(
                            table_holdings,
                            table_info,
                            validation,
                        ),
                }
            )

        except Exception:
            pass

    # -------------------------------------------------------------------------
    # Candidate 3: explicit names-only ranked parser
    # -------------------------------------------------------------------------

    names_only_holdings, names_only_info = (
        parse_names_only_candidate(
            lines
        )
    )

    if names_only_holdings:

        try:

            validation = validate_holdings(
                names_only_holdings,
                require_rank_one=True,
            )

            candidates.append(
                {
                    "holdings":
                        names_only_holdings,

                    "parser":
                        names_only_info,

                    "validation":
                        validation,

                    "score":
                        score_candidate(
                            names_only_holdings,
                            names_only_info,
                            validation,
                        ),
                }
            )

        except Exception:
            pass

    if not candidates:

        raise RuntimeError(
            "Top Holdings section was found, but no reliable "
            "holding rows could be established from the PDF text."
        )

    # -------------------------------------------------------------------------
    # Select the strongest structurally valid candidate.
    #
    # We do NOT simply choose the candidate with the most rows.
    # Structural evidence is weighted first.
    # -------------------------------------------------------------------------

    candidates.sort(
        key=lambda candidate: candidate[
            "score"
        ],
        reverse=True,
    )

    selected = candidates[0]

    holdings = selected[
        "holdings"
    ]

    parser_info = selected[
        "parser"
    ]

    validation = selected[
        "validation"
    ]

    parser_info = dict(
        parser_info
    )

    parser_info[
        "selectionScore"
    ] = selected[
        "score"
    ]

    parser_info[
        "candidateCount"
    ] = len(
        candidates
    )

    parser_info[
        "duplicateNameWarnings"
    ] = validation[
        "duplicateNameWarnings"
    ]

    parser_info[
        "publishedPercentages"
    ] = validation[
        "publishedPercentageCount"
    ]

    parser_info[
        "namesWithoutPercentage"
    ] = validation[
        "missingPercentageCount"
    ]

    return (
        holdings,
        parser_info,
    )


# =============================================================================
# CANDIDATE SCORING
# =============================================================================

def score_candidate(
    holdings: list[dict],
    parser_info: dict,
    validation: dict,
) -> int:

    if not holdings:
        return -100000

    score = 0

    parser_name = parser_info.get(
        "parser"
    )

    # Explicit ranked parsing is preferred.
    if parser_name == "ranked_block_parser":
        score += 100

    elif parser_name == "names_only_ranked_parser":
        score += 90

    elif parser_name == "table_row_parser":
        score += 40

    # Rank 1 is strong evidence.
    ranks = [
        item.get(
            "rank"
        )
        for item
        in holdings
    ]

    if 1 in ranks:
        score += 50

    # Sequential ranks are stronger than arbitrary ranks.
    if ranks:

        expected = list(
            range(
                1,
                len(ranks) + 1,
            )
        )

        if ranks == expected:
            score += 80

        else:
            score -= 20

    # More holdings can be evidence, but not overwhelmingly so.
    score += min(
        len(holdings) * 5,
        50,
    )

    # Published percentages are useful evidence.
    percentage_count = sum(
        1
        for item
        in holdings
        if item.get(
            "weightPercent"
        ) is not None
    )

    score += min(
        percentage_count * 3,
        30,
    )

    # Penalize duplicate names slightly, but do not discard them.
    duplicate_warnings = validation.get(
        "duplicateNameWarnings",
        [],
    )

    score -= len(
        duplicate_warnings
    ) * 5

    return score


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
    # PDF extraction variants.
    # -------------------------------------------------------------------------

    (
        text_variants,
        pdf_page_count,
    ) = extract_pdf_text_variants(
        factsheet_bytes
    )

    primary_text = text_variants[0][
        "text"
    ]

    data_as_at = (
        extract_data_as_at(
            primary_text
        )
    )

    document_date = (
        extract_document_date(
            primary_text
        )
    )

    # -------------------------------------------------------------------------
    # Try every available PDF extraction layout.
    #
    # This is important because pypdf's normal text order and layout order
    # can produce completely different table structures.
    # -------------------------------------------------------------------------

    published_sections = []

    for variant in text_variants:

        (
            section_text,
            section_status,
        ) = extract_holdings_section(
            variant["text"]
        )

        if section_status == "published":

            published_sections.append(
                {
                    "extractionMode":
                        variant["name"],

                    "fullText":
                        variant["text"],

                    "sectionText":
                        section_text,
                }
            )

    # -------------------------------------------------------------------------
    # No section at all.
    # -------------------------------------------------------------------------

    if not published_sections:

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

            "rules":
                rules_payload(),
        }

    # -------------------------------------------------------------------------
    # Parse every extraction mode.
    # -------------------------------------------------------------------------

    successful_candidates = []

    parser_errors = []

    for published_section in published_sections:

        try:

            (
                holdings,
                parser_info,
            ) = parse_holdings(
                published_section[
                    "sectionText"
                ]
            )

            successful_candidates.append(
                {
                    "holdings":
                        holdings,

                    "parser":
                        parser_info,

                    "extractionMode":
                        published_section[
                            "extractionMode"
                        ],

                    "sectionText":
                        published_section[
                            "sectionText"
                        ],

                    "fullText":
                        published_section[
                            "fullText"
                        ],
                }
            )

        except Exception as error:

            parser_errors.append(
                {
                    "extractionMode":
                        published_section[
                            "extractionMode"
                        ],

                    "error":
                        clean_text(
                            str(error)
                        ),
                }
            )

    if not successful_candidates:

        raise RuntimeError(
            "Top Holdings section was found, but no reliable "
            "holding rows could be established. "
            f"Parser diagnostics: {parser_errors}"
        )

    # -------------------------------------------------------------------------
    # Select the best extraction-mode result.
    # -------------------------------------------------------------------------

    def candidate_score(
        candidate
    ):

        holdings = candidate[
            "holdings"
        ]

        parser_info = candidate[
            "parser"
        ]

        validation = {
            "duplicateNameWarnings":
                parser_info.get(
                    "duplicateNameWarnings",
                    [],
                ),
        }

        return score_candidate(
            holdings,
            parser_info,
            validation,
        )

    successful_candidates.sort(
        key=candidate_score,
        reverse=True,
    )

    selected = successful_candidates[0]

    holdings = selected[
        "holdings"
    ]

    parser_info = dict(
        selected[
            "parser"
        ]
    )

    parser_info[
        "extractionMode"
    ] = selected[
        "extractionMode"
    ]

    parser_info[
        "alternativeParserErrors"
    ] = parser_errors

    parser_info[
        "availableExtractionModes"
    ] = [
        item[
            "extractionMode"
        ]
        for item
        in published_sections
    ]

    percentage_count = sum(
        1
        for holding
        in holdings
        if holding.get(
            "weightPercent"
        ) is not None
    )

    missing_percentage_count = (
        len(
            holdings
        )
        -
        percentage_count
    )

    duplicate_warnings = (
        duplicate_name_warnings(
            holdings
        )
    )

    parser_info[
        "duplicateNameWarnings"
    ] = duplicate_warnings

    # -------------------------------------------------------------------------
    # IMPORTANT:
    #
    # Duplicate names are NOT automatically failures.
    #
    # The published rank is the authoritative row identity. If Prudential
    # legitimately publishes the same name more than once, we preserve it.
    # -------------------------------------------------------------------------

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
            percentage_count,

        "holdingsWithoutPublishedPercentage":
            missing_percentage_count,

        "topHoldings":
            holdings,

        "holdingsParser":
            parser_info,

        "rules":
            rules_payload(),
    }


# =============================================================================
# RULES
# =============================================================================

def rules_payload() -> dict:

    return {
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

        "noFabricatedHoldings":
            True,

        "noFabricatedPercentages":
            True,

        "noCalculatedPercentages":
            True,

        "nullWeightAllowedWhenNotPublished":
            True,

        "ambiguousRowsFail":
            True,

        "duplicateRanksFail":
            True,

        "duplicateNamesAreWarnings":
            True,

        "thirdPartyHoldings":
            False,
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

    identifier = ""

    if result.get(
        "fundIdentifier"
    ):

        identifier = safe_filename(
            result.get(
                "fundIdentifier"
            )
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
                # Final verification.
                #
                # Download the official PDF again and independently parse it.
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
                        text_variants,
                        verification_page_count,
                    ) = extract_pdf_text_variants(
                        factsheet_bytes
                    )

                    verification_candidates = []

                    for variant in text_variants:

                        (
                            section_text,
                            section_status,
                        ) = extract_holdings_section(
                            variant[
                                "text"
                            ]
                        )

                        if section_status != "published":
                            continue

                        try:

                            (
                                verified_holdings,
                                verified_parser_info,
                            ) = parse_holdings(
                                section_text
                            )

                            verification_candidates.append(
                                {
                                    "holdings":
                                        verified_holdings,

                                    "parser":
                                        verified_parser_info,

                                    "extractionMode":
                                        variant[
                                            "name"
                                        ],

                                    "fullText":
                                        variant[
                                            "text"
                                        ],

                                    "sectionText":
                                        section_text,
                                }
                            )

                        except Exception:
                            continue

                    if not verification_candidates:

                        raise RuntimeError(
                            "Final verification could not establish "
                            "reliable Top Holdings rows."
                        )

                    verification_candidates.sort(
                        key=lambda candidate:
                            score_candidate(
                                candidate[
                                    "holdings"
                                ],
                                candidate[
                                    "parser"
                                ],
                                {
                                    "duplicateNameWarnings":
                                        duplicate_name_warnings(
                                            candidate[
                                                "holdings"
                                            ]
                                        ),
                                },
                            ),
                        reverse=True,
                    )

                    verified = (
                        verification_candidates[0]
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
                            "final verification: "
                            f"initial={result['topHoldingsCount']} "
                            f"verified={len(verified_holdings)}"
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
                    ] = verified[
                        "parser"
                    ]

                    result[
                        "holdingsParser"
                    ] = dict(
                        result[
                            "holdingsParser"
                        ]
                    )

                    result[
                        "holdingsParser"
                    ][
                        "verificationExtractionMode"
                    ] = verified[
                        "extractionMode"
                    ]

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
                            verified[
                                "fullText"
                            ],
                            verified[
                                "sectionText"
                            ],
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

                duplicate_warnings = (
                    result.get(
                        "holdingsParser",
                        {}
                    ).get(
                        "duplicateNameWarnings",
                        [],
                    )
                )

                if duplicate_warnings:

                    print(
                        "\n"
                        "WARNING - DUPLICATE PUBLISHED "
                        "HOLDING NAMES:"
                    )

                    for warning in duplicate_warnings:

                        print(
                            f"  {warning}"
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

        "rules":
            rules_payload(),

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

    duplicate_warning_funds = []

    for item in successful:

        warnings = (
            item.get(
                "holdingsParser",
                {}
            ).get(
                "duplicateNameWarnings",
                [],
            )
        )

        if warnings:

            duplicate_warning_funds.append(
                {
                    "excelRow":
                        item.get(
                            "excelRow"
                        ),

                    "fundName":
                        item.get(
                            "fundName"
                        ),

                    "warnings":
                        warnings,
                }
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

        "duplicateHoldingNameWarningFunds":
            duplicate_warning_funds,

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

        "rules":
            rules_payload(),
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
        f"Funds with duplicate-name warnings: "
        f"{len(duplicate_warning_funds)}"
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
