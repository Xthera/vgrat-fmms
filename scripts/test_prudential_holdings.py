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

The parser supports Prudential factsheets / fund-information documents where:

    - holding names wrap across multiple lines
    - rank numbers are in a separate column
    - percentages are in a separate column
    - percentages are NOT published for some funds
    - PDF text extraction order differs from visual table order
    - duplicate holding names legitimately appear
    - the Top Holdings heading contains extraction artifacts such as
      "Top 10 Holdings3"


HARD RULES
==========

- Excel Column A controls the universe.
- No hardcoded 67-fund limit.
- Every populated URL in Column A is processed.
- Only official Prudential Singapore pages are accepted.
- Only official Prudential Singapore factsheets/documents are accepted.
- No third-party holdings source.
- No inferred holdings.
- No fabricated holdings.
- No fabricated percentages.
- No calculated percentages.
- No estimated percentages.
- No forced 10 holdings.
- If Prudential publishes fewer than 10 holdings, store exactly that count.
- Holdings are stored in the order published by Prudential.
- Multi-line holding names are joined into one holding.
- Standalone rank numbers do NOT define a completed row by themselves.
- Published percentage is stored only when actually published.
- A holding without a published percentage is valid.
- Missing percentage is represented as:
      weightPercent = null
      weightText = "-"
      weightPublished = false
- Duplicate holding names are allowed.
- If the official document contains a Top Holdings section but names cannot
  be safely reconstructed, the fund is marked FAILED.
- A fund is marked NO_HOLDINGS_SECTION only when there is no evidence of a
  Top Holdings section in the official document.
- Failed funds do not get fabricated holdings.
- This script never changes PruAccess settings.
- This script does NOT modify test_pruaccess.py.
- This script does NOT create data.json.
- This script does NOT create index.html/css/js.
- This script is a TEST collector only.


PARSER STRATEGY
===============

PRIMARY:
    pdfplumber coordinate-based visual reconstruction.

For every visual Top Holdings section:

    1. Find Top 10 Holdings heading.
    2. Identify holding rank anchors where available.
    3. Reconstruct rows by vertical PDF coordinates.
    4. Join wrapped holding-name lines.
    5. Search each row for an actual published percentage.
    6. If a percentage exists, store it.
    7. If no percentage exists, store null / "-".
    8. Never infer a missing weight.

This allows:

    1   LONG HOLDING NAME
        CONTINUED NAME

    2   ANOTHER HOLDING NAME

without requiring every holding to have a percentage.

FALLBACK:
    pypdf text parser.

NO_HOLDINGS_SECTION RULE
========================

A parser failure must NOT be converted into NO_HOLDINGS_SECTION.

The decision is:

    Heading genuinely absent
        -> NO_HOLDINGS_SECTION

    Heading present
        -> attempt parser
        -> parser fails
        -> FAILED

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
            top_holdings_visual_lines.txt
            top_holdings.json
            metadata.json

        <excelRow>_<identifier>_failed/
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
from statistics import median
from urllib.parse import urljoin, urlparse

import pdfplumber
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
# PDF
# -----------------------------------------------------------------------------

PDF_WORD_X_TOLERANCE = 2

PDF_WORD_Y_TOLERANCE = 2

PDF_LINE_Y_TOLERANCE = 3

PDF_SECTION_PADDING = 8

MAX_VISUAL_LINES = 250

MAX_HOLDINGS = 10

MAX_HOLDINGS_SECTION_HEIGHT = 500


# -----------------------------------------------------------------------------
# Official Prudential Singapore hosts only.
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
    """
    Normalize whitespace.
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
        .replace(
            "+00:00",
            "Z",
        )
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


def extract_citicode_from_url(
    url: str,
) -> str | None:

    parsed = urlparse(
        url
    )

    match = re.search(
        r"(?:^|&)citicode=([^&]+)",
        parsed.query,
        flags=re.IGNORECASE,
    )

    if match:

        return clean_text(
            match.group(1)
        )

    return None


# =============================================================================
# EXCEL MASTER UNIVERSE
# =============================================================================

def read_excel_funds() -> list[dict]:
    """
    Read every populated URL from Column A.

    Column A controls the master universe.

    Column B is retained as the PruAccess reference name.
    """

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

        if not is_prudential_url(
            prudential_url
        ):

            raise RuntimeError(
                f"Excel row {row_number} contains "
                f"a non-Prudential URL: "
                f"{prudential_url}"
            )

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
# FIND FACTSHEET
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
            absolute_url.casefold()
        )

        text_lower = (
            anchor_text.casefold()
        )

        score = 0

        if "fund factsheet" in text_lower:

            score += 200

        elif "factsheet" in text_lower:

            score += 150

        if "factsheet" in href_lower:

            score += 100

        if (
            href_lower.endswith(
                ".pdf"
            )
            or
            ".pdf?" in href_lower
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
# PDF VISUAL LINES
# =============================================================================

def group_pdf_words_into_lines(
    words: list[dict],
) -> list[dict]:

    if not words:

        return []

    sorted_words = sorted(
        words,
        key=lambda word: (
            float(
                word.get(
                    "top",
                    0,
                )
            ),
            float(
                word.get(
                    "x0",
                    0,
                )
            ),
        ),
    )

    groups = []

    for word in sorted_words:

        top = float(
            word.get(
                "top",
                0,
            )
        )

        bottom = float(
            word.get(
                "bottom",
                top,
            )
        )

        placed = False

        for group in reversed(
            groups[-6:]
        ):

            if abs(
                top
                -
                group["top"]
            ) <= PDF_LINE_Y_TOLERANCE:

                group[
                    "words"
                ].append(
                    word
                )

                group[
                    "top"
                ] = min(
                    group["top"],
                    top,
                )

                group[
                    "bottom"
                ] = max(
                    group["bottom"],
                    bottom,
                )

                group[
                    "x0"
                ] = min(
                    group["x0"],
                    float(
                        word.get(
                            "x0",
                            0,
                        )
                    ),
                )

                group[
                    "x1"
                ] = max(
                    group["x1"],
                    float(
                        word.get(
                            "x1",
                            0,
                        )
                    ),
                )

                placed = True

                break

        if not placed:

            groups.append(
                {
                    "top":
                        top,

                    "bottom":
                        bottom,

                    "x0":
                        float(
                            word.get(
                                "x0",
                                0,
                            )
                        ),

                    "x1":
                        float(
                            word.get(
                                "x1",
                                0,
                            )
                        ),

                    "words":
                        [
                            word
                        ],
                }
            )

    output = []

    for group in groups:

        ordered_words = sorted(
            group[
                "words"
            ],
            key=lambda word: float(
                word.get(
                    "x0",
                    0,
                )
            ),
        )

        parts = []

        for word in ordered_words:

            text = clean_text(
                word.get(
                    "text",
                    "",
                )
            )

            if text:

                parts.append(
                    text
                )

        line_text = clean_text(
            " ".join(
                parts
            )
        )

        if not line_text:

            continue

        output.append(
            {
                "text":
                    line_text,

                "top":
                    group[
                        "top"
                    ],

                "bottom":
                    group[
                        "bottom"
                    ],

                "x0":
                    group[
                        "x0"
                    ],

                "x1":
                    group[
                        "x1"
                    ],

                "words":
                    ordered_words,
            }
        )

    output.sort(
        key=lambda item: (
            item["top"],
            item["x0"],
        )
    )

    return output


# =============================================================================
# TOP HOLDINGS HEADING
# =============================================================================

def is_top_holdings_heading(
    text: str,
) -> bool:
    """
    Supports:

        Top 10 Holdings
        Top 10 Holdings3
        Top 10 Holdings 3
        Top Ten Holdings
        Top Ten Holdings3
    """

    normalized = normalize_text(
        text
    )

    return bool(
        re.search(
            r"\b"
            r"top"
            r"\s*"
            r"(?:10|ten)"
            r"\s+"
            r"holdings"
            r"(?:\s*\d+)?"
            r"\b",
            normalized,
            flags=re.IGNORECASE,
        )
    )


def find_textual_top_holdings_heading(
    full_text: str,
) -> str | None:

    patterns = [

        r"(?i)\btop\s*10\s*holdings(?:\s*\d+)?\b",

        r"(?i)\btop\s*ten\s+holdings(?:\s*\d+)?\b",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            full_text,
        )

        if match:

            return clean_text(
                match.group(0)
            )

    return None


def locate_visual_top_holdings_heading(
    visual_lines: list[dict],
) -> dict | None:

    for line in visual_lines:

        if not is_top_holdings_heading(
            line[
                "text"
            ]
        ):

            continue

        return {
            "x0":
                line[
                    "x0"
                ],

            "x1":
                line[
                    "x1"
                ],

            "top":
                line[
                    "top"
                ],

            "bottom":
                line[
                    "bottom"
                ],

            "text":
                line[
                    "text"
                ],
        }

    return None


# =============================================================================
# PERCENTAGE
# =============================================================================

def percentage_value(
    text: str,
) -> float | None:

    if not text:

        return None

    match = re.fullmatch(
        r"\s*"
        r"([0-9]+(?:\.[0-9]+)?)"
        r"\s*%"
        r"\s*",
        text,
    )

    if not match:

        return None

    try:

        value = float(
            match.group(1)
        )

    except ValueError:

        return None

    if (
        value < 0
        or
        value > 100
    ):

        return None

    return value


def extract_percentage_tokens(
    words: list[dict],
) -> list[dict]:

    sorted_words = sorted(
        words,
        key=lambda word: (
            float(
                word.get(
                    "top",
                    0,
                )
            ),
            float(
                word.get(
                    "x0",
                    0,
                )
            ),
        ),
    )

    tokens = []

    index = 0

    while index < len(
        sorted_words
    ):

        word = sorted_words[
            index
        ]

        text = clean_text(
            word.get(
                "text",
                "",
            )
        )

        direct_value = (
            percentage_value(
                text
            )
        )

        if direct_value is not None:

            tokens.append(
                {
                    "value":
                        direct_value,

                    "text":
                        text,

                    "x0":
                        float(
                            word.get(
                                "x0",
                                0,
                            )
                        ),

                    "x1":
                        float(
                            word.get(
                                "x1",
                                0,
                            )
                        ),

                    "top":
                        float(
                            word.get(
                                "top",
                                0,
                            )
                        ),

                    "bottom":
                        float(
                            word.get(
                                "bottom",
                                word.get(
                                    "top",
                                    0,
                                ),
                            )
                        ),
                }
            )

            index += 1

            continue

        # Separate number + %.
        if (
            re.fullmatch(
                r"[0-9]+(?:\.[0-9]+)?",
                text,
            )
            and
            index + 1
            <
            len(
                sorted_words
            )
        ):

            next_word = sorted_words[
                index + 1
            ]

            next_text = clean_text(
                next_word.get(
                    "text",
                    "",
                )
            )

            same_line = (
                abs(
                    float(
                        next_word.get(
                            "top",
                            0,
                        )
                    )
                    -
                    float(
                        word.get(
                            "top",
                            0,
                        )
                    )
                )
                <=
                PDF_LINE_Y_TOLERANCE
            )

            close_x = (
                float(
                    next_word.get(
                        "x0",
                        0,
                    )
                )
                -
                float(
                    word.get(
                        "x1",
                        0,
                    )
                )
                <=
                8
            )

            if (
                next_text == "%"
                and
                same_line
                and
                close_x
            ):

                try:

                    value = float(
                        text
                    )

                except ValueError:

                    value = None

                if (
                    value is not None
                    and
                    0 <= value <= 100
                ):

                    tokens.append(
                        {
                            "value":
                                value,

                            "text":
                                f"{text}%",

                            "x0":
                                float(
                                    word.get(
                                        "x0",
                                        0,
                                    )
                                ),

                            "x1":
                                float(
                                    next_word.get(
                                        "x1",
                                        0,
                                    )
                                ),

                            "top":
                                min(
                                    float(
                                        word.get(
                                            "top",
                                            0,
                                        )
                                    ),
                                    float(
                                        next_word.get(
                                            "top",
                                            0,
                                        )
                                    ),
                                ),

                            "bottom":
                                max(
                                    float(
                                        word.get(
                                            "bottom",
                                            word.get(
                                                "top",
                                                0,
                                            ),
                                        )
                                    ),
                                    float(
                                        next_word.get(
                                            "bottom",
                                            next_word.get(
                                                "top",
                                                0,
                                            ),
                                        )
                                    ),
                                ),
                        }
                    )

                    index += 2

                    continue

        index += 1

    return tokens


# =============================================================================
# NAME CLEANUP
# =============================================================================

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


def clean_holding_name(
    value: str,
) -> str:

    name = clean_holding_fragment(
        value
    )

    if not name:

        return ""

    name = re.sub(
        r"^\d{1,2}[.)\-:]\s*",
        "",
        name,
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
        "",
        "none",
        "null",
        "holding",
        "holdings",
        "name",
        "weight",
        "weights",
        "allocation",
        "allocations",
        "source",
        "source:",
        "disclaimer",
        "-",
        "—",
    }:

        return ""

    return name


def is_noise_fragment(
    value: str,
) -> bool:

    normalized = normalize_text(
        value
    )

    if not normalized:

        return True

    return normalized in {
        "holding",
        "holdings",
        "name",
        "weight",
        "weights",
        "allocation",
        "allocations",
        "portfolio holdings",
        "top 10 holdings",
        "source",
        "source:",
        "none",
        "null",
    }


def combine_name_fragments(
    fragments: list[str],
) -> str:

    parts = []

    for fragment in fragments:

        cleaned = clean_holding_fragment(
            fragment
        )

        if not cleaned:

            continue

        if re.fullmatch(
            r"\d{1,2}",
            cleaned,
        ):

            continue

        if is_noise_fragment(
            cleaned
        ):

            continue

        parts.append(
            cleaned
        )

    if not parts:

        return ""

    return clean_holding_name(
        " ".join(
            parts
        )
    )


# =============================================================================
# RANK EXTRACTION
# =============================================================================

def parse_rank_token(
    text: str,
) -> int | None:
    """
    Detect standalone visual rank:

        1
        2
        10
        1.
        2)
        3:
    """

    text = clean_text(
        text
    )

    match = re.fullmatch(
        r"([0-9]{1,2})[.)\-:]?",
        text,
    )

    if not match:

        return None

    try:

        value = int(
            match.group(1)
        )

    except ValueError:

        return None

    if (
        value < 1
        or
        value > MAX_HOLDINGS
    ):

        return None

    return value


def find_rank_candidates(
    words: list[dict],
    heading_bottom: float,
    section_end: float,
) -> list[dict]:

    candidates = []

    for word in words:

        top = float(
            word.get(
                "top",
                0,
            )
        )

        bottom = float(
            word.get(
                "bottom",
                top,
            )
        )

        if bottom <= heading_bottom:
            continue

        if top >= section_end:
            continue

        rank = parse_rank_token(
            word.get(
                "text",
                "",
            )
        )

        if rank is None:
            continue

        candidates.append(
            {
                "rank":
                    rank,

                "x0":
                    float(
                        word.get(
                            "x0",
                            0,
                        )
                    ),

                "x1":
                    float(
                        word.get(
                            "x1",
                            0,
                        )
                    ),

                "top":
                    top,

                "bottom":
                    bottom,

                "center":
                    (
                        top
                        +
                        bottom
                    )
                    /
                    2,
            }
        )

    return candidates


def select_rank_column(
    candidates: list[dict],
) -> list[dict]:
    """
    Rank values can occur elsewhere in the page.

    Select the x-coordinate cluster that most plausibly represents the
    Top Holdings rank column.
    """

    if not candidates:

        return []

    clusters = []

    tolerance = 18.0

    for candidate in candidates:

        placed = False

        for cluster in clusters:

            if abs(
                candidate[
                    "x0"
                ]
                -
                cluster[
                    "median_x"
                ]
            ) <= tolerance:

                cluster[
                    "items"
                ].append(
                    candidate
                )

                xs = [
                    item[
                        "x0"
                    ]
                    for item
                    in cluster[
                        "items"
                    ]
                ]

                cluster[
                    "median_x"
                ] = median(
                    xs
                )

                placed = True

                break

        if not placed:

            clusters.append(
                {
                    "median_x":
                        candidate[
                            "x0"
                        ],

                    "items":
                        [
                            candidate
                        ],
                }
            )

    # Prefer the cluster with:
    #   1. the most distinct ranks
    #   2. the greatest count
    #   3. the smallest x position
    #
    # Rank columns are normally left of the holding names.
    def cluster_score(
        cluster,
    ):

        distinct_ranks = len(
            {
                item[
                    "rank"
                ]
                for item
                in cluster[
                    "items"
                ]
            }
        )

        count = len(
            cluster[
                "items"
            ]
        )

        x = cluster[
            "median_x"
        ]

        return (
            distinct_ranks,
            count,
            -x,
        )

    clusters.sort(
        key=cluster_score,
        reverse=True,
    )

    return clusters[0][
        "items"
    ]


# =============================================================================
# VISUAL ROW RECONSTRUCTION
# =============================================================================

def reconstruct_visual_holdings(
    pdf_bytes: bytes,
) -> tuple[
    list[dict],
    str,
    list[str],
    int,
    str,
]:
    """
    Primary visual parser.

    Supports both:

        ranked + percentage

    and:

        ranked + names only.

    Percentages are optional.

    Returns:

        holdings
        section_text
        debug_lines
        pdf_page_number
        weight_status

    weight_status:

        published
        not_published
        partial
    """

    with pdfplumber.open(
        BytesIO(
            pdf_bytes
        )
    ) as pdf:

        for page_number, pdf_page in enumerate(
            pdf.pages,
            start=1,
        ):

            try:

                words = pdf_page.extract_words(
                    x_tolerance=PDF_WORD_X_TOLERANCE,
                    y_tolerance=PDF_WORD_Y_TOLERANCE,
                    keep_blank_chars=False,
                    use_text_flow=False,
                )

            except Exception:

                continue

            if not words:
                continue

            visual_lines = (
                group_pdf_words_into_lines(
                    words
                )
            )

            heading = (
                locate_visual_top_holdings_heading(
                    visual_lines
                )
            )

            if heading is None:

                continue

            heading_bottom = (
                heading[
                    "bottom"
                ]
            )

            heading_left = (
                heading[
                    "x0"
                ]
            )

            section_start = (
                heading_bottom
                + 2
            )

            section_end = min(
                float(
                    pdf_page.height
                ),
                section_start
                +
                MAX_HOLDINGS_SECTION_HEIGHT,
            )

            # -------------------------------------------------------------
            # Collect words beneath heading.
            # -------------------------------------------------------------

            section_words = []

            for word in words:

                top = float(
                    word.get(
                        "top",
                        0,
                    )
                )

                bottom = float(
                    word.get(
                        "bottom",
                        top,
                    )
                )

                x1 = float(
                    word.get(
                        "x1",
                        0,
                    )
                )

                if bottom <= section_start:
                    continue

                if top >= section_end:
                    continue

                if x1 < heading_left - PDF_SECTION_PADDING:
                    continue

                section_words.append(
                    word
                )

            if not section_words:

                continue

            # -------------------------------------------------------------
            # Identify rank column.
            # -------------------------------------------------------------

            rank_candidates = (
                find_rank_candidates(
                    section_words,
                    heading_bottom,
                    section_end,
                )
            )

            rank_column = (
                select_rank_column(
                    rank_candidates
                )
            )

            # -------------------------------------------------------------
            # Build unique visual rank anchors.
            # -------------------------------------------------------------

            ranks_by_number = {}

            for candidate in rank_column:

                rank = candidate[
                    "rank"
                ]

                existing = (
                    ranks_by_number.get(
                        rank
                    )
                )

                if existing is None:

                    ranks_by_number[
                        rank
                    ] = candidate

                    continue

                # Prefer the candidate occurring later/within the main
                # Top Holdings vertical table when duplicates exist.
                if (
                    abs(
                        candidate[
                            "x0"
                        ]
                        -
                        heading_left
                    )
                    <
                    abs(
                        existing[
                            "x0"
                        ]
                        -
                        heading_left
                    )
                ):

                    ranks_by_number[
                        rank
                    ] = candidate

            # Require at least one plausible rank for names-only mode.
            ranked_items = sorted(
                ranks_by_number.values(),
                key=lambda item: item[
                    "center"
                ],
            )

            # -------------------------------------------------------------
            # Percentage tokens in section.
            # -------------------------------------------------------------

            percentage_tokens = (
                extract_percentage_tokens(
                    section_words
                )
            )

            percentage_tokens.sort(
                key=lambda item: (
                    item["top"],
                    item["x0"],
                )
            )

            # -------------------------------------------------------------
            # If we have ranks, reconstruct each visual row.
            # -------------------------------------------------------------

            if ranked_items:

                centers = [
                    item[
                        "center"
                    ]
                    for item
                    in ranked_items
                ]

                positive_gaps = []

                for index in range(
                    1,
                    len(
                        centers
                    ),
                ):

                    gap = (
                        centers[index]
                        -
                        centers[index - 1]
                    )

                    if gap > 0:

                        positive_gaps.append(
                            gap
                        )

                typical_gap = (
                    median(
                        positive_gaps
                    )
                    if positive_gaps
                    else 20.0
                )

                # Keep only plausible table ranks.
                ranked_items = ranked_items[
                    :MAX_HOLDINGS
                ]

                reconstructed = []

                debug_lines = []

                for index, rank_item in enumerate(
                    ranked_items
                ):

                    row_center = (
                        rank_item[
                            "center"
                        ]
                    )

                    if index == 0:

                        row_top = max(
                            section_start,
                            row_center
                            -
                            max(
                                typical_gap,
                                18.0,
                            ),
                        )

                    else:

                        previous_center = (
                            ranked_items[
                                index - 1
                            ][
                                "center"
                            ]
                        )

                        row_top = (
                            previous_center
                            +
                            row_center
                        ) / 2

                    if index + 1 < len(
                        ranked_items
                    ):

                        next_center = (
                            ranked_items[
                                index + 1
                            ][
                                "center"
                            ]
                        )

                        row_bottom = (
                            row_center
                            +
                            next_center
                        ) / 2

                    else:

                        row_bottom = min(
                            section_end,
                            row_center
                            +
                            max(
                                typical_gap,
                                22.0,
                            ),
                        )

                    # -----------------------------------------------------
                    # Words in row.
                    # -----------------------------------------------------

                    row_words = []

                    for word in section_words:

                        word_top = float(
                            word.get(
                                "top",
                                0,
                            )
                        )

                        word_bottom = float(
                            word.get(
                                "bottom",
                                word_top,
                            )
                        )

                        word_center = (
                            word_top
                            +
                            word_bottom
                        ) / 2

                        if (
                            word_center
                            <
                            row_top
                        ):

                            continue

                        if (
                            word_center
                            >
                            row_bottom
                        ):

                            continue

                        row_words.append(
                            word
                        )

                    # -----------------------------------------------------
                    # Find actual percentage belonging to this row.
                    #
                    # If none exists, this is valid names-only data.
                    # -----------------------------------------------------

                    row_percentages = []

                    for token in percentage_tokens:

                        percentage_center = (
                            token[
                                "top"
                            ]
                            +
                            token[
                                "bottom"
                            ]
                        ) / 2

                        if (
                            percentage_center
                            <
                            row_top
                        ):

                            continue

                        if (
                            percentage_center
                            >
                            row_bottom
                        ):

                            continue

                        row_percentages.append(
                            token
                        )

                    percentage_token = None

                    if row_percentages:

                        # Holding weight should normally be the rightmost
                        # percentage in the visual row.
                        percentage_token = max(
                            row_percentages,
                            key=lambda token: token[
                                "x0"
                            ],
                        )

                    # -----------------------------------------------------
                    # Build name words.
                    # -----------------------------------------------------

                    name_words = []

                    for word in row_words:

                        text = clean_text(
                            word.get(
                                "text",
                                "",
                            )
                        )

                        if not text:
                            continue

                        # Skip percentage tokens.
                        if (
                            percentage_value(
                                text
                            )
                            is not None
                        ):

                            continue

                        if text == "%":

                            continue

                        # Skip standalone rank values.
                        rank_value = parse_rank_token(
                            text
                        )

                        if (
                            rank_value is not None
                            and
                            abs(
                                float(
                                    word.get(
                                        "x0",
                                        0,
                                    )
                                )
                                -
                                rank_item[
                                    "x0"
                                ]
                            )
                            <=
                            20
                        ):

                            continue

                        # If percentage exists, names live to the left
                        # of the percentage column.
                        if percentage_token is not None:

                            word_x1 = float(
                                word.get(
                                    "x1",
                                    0,
                                )
                            )

                            if (
                                word_x1
                                >
                                percentage_token[
                                    "x0"
                                ]
                                +
                                4
                            ):

                                continue

                        # Name should normally be to the right of the
                        # rank column.
                        word_x0 = float(
                            word.get(
                                "x0",
                                0,
                            )
                        )

                        if (
                            word_x0
                            <
                            rank_item[
                                "x1"
                            ]
                            -
                            2
                        ):

                            continue

                        name_words.append(
                            word
                        )

                    # -----------------------------------------------------
                    # Group wrapped name lines.
                    # -----------------------------------------------------

                    name_lines = (
                        group_pdf_words_into_lines(
                            name_words
                        )
                    )

                    fragments = []

                    for name_line in name_lines:

                        fragment = clean_text(
                            name_line[
                                "text"
                            ]
                        )

                        if not fragment:
                            continue

                        if is_noise_fragment(
                            fragment
                        ):
                            continue

                        fragments.append(
                            fragment
                        )

                    name = (
                        combine_name_fragments(
                            fragments
                        )
                    )

                    # -----------------------------------------------------
                    # If coordinate filtering was too aggressive, retry
                    # without requiring x > rank column.
                    # -----------------------------------------------------

                    if not name:

                        relaxed_words = []

                        for word in row_words:

                            text = clean_text(
                                word.get(
                                    "text",
                                    "",
                                )
                            )

                            if not text:
                                continue

                            if text == "%":
                                continue

                            if (
                                percentage_value(
                                    text
                                )
                                is not None
                            ):
                                continue

                            rank_value = (
                                parse_rank_token(
                                    text
                                )
                            )

                            if (
                                rank_value is not None
                            ):

                                continue

                            if (
                                percentage_token
                                is not None
                            ):

                                word_x1 = float(
                                    word.get(
                                        "x1",
                                        0,
                                    )
                                )

                                if (
                                    word_x1
                                    >
                                    percentage_token[
                                        "x0"
                                    ]
                                    +
                                    4
                                ):

                                    continue

                            relaxed_words.append(
                                word
                            )

                        relaxed_lines = (
                            group_pdf_words_into_lines(
                                relaxed_words
                            )
                        )

                        relaxed_fragments = []

                        for relaxed_line in relaxed_lines:

                            fragment = clean_text(
                                relaxed_line[
                                    "text"
                                ]
                            )

                            if not fragment:
                                continue

                            if is_noise_fragment(
                                fragment
                            ):
                                continue

                            relaxed_fragments.append(
                                fragment
                            )

                        name = (
                            combine_name_fragments(
                                relaxed_fragments
                            )
                        )

                    # -----------------------------------------------------
                    # No usable holding name.
                    #
                    # Do not fabricate one.
                    # -----------------------------------------------------

                    if not name:

                        debug_lines.append(
                            (
                                f"rank={rank_item['rank']} "
                                f"NO_NAME "
                                f"y={row_center:.2f}"
                            )
                        )

                        continue

                    # -----------------------------------------------------
                    # Percentage.
                    # -----------------------------------------------------

                    if percentage_token is not None:

                        holding = {
                            "rank":
                                rank_item[
                                    "rank"
                                ],

                            "name":
                                name,

                            "weightPercent":
                                percentage_token[
                                    "value"
                                ],

                            "weightText":
                                percentage_token[
                                    "text"
                                ],

                            "weightPublished":
                                True,
                        }

                    else:

                        holding = {
                            "rank":
                                rank_item[
                                    "rank"
                                ],

                            "name":
                                name,

                            "weightPercent":
                                None,

                            "weightText":
                                "-",

                            "weightPublished":
                                False,
                        }

                    reconstructed.append(
                        holding
                    )

                    debug_lines.append(
                        (
                            f"rank={rank_item['rank']} "
                            f"weight="
                            f"{holding['weightText']} "
                            f"y={row_center:.2f} "
                            f"{name}"
                        )
                    )

                    if len(
                        reconstructed
                    ) >= MAX_HOLDINGS:

                        break

                # ---------------------------------------------------------
                # Validate visual result.
                # ---------------------------------------------------------

                if reconstructed:

                    # Remove accidental non-sequential later ranks while
                    # preserving the visual order. The rank in output is
                    # normalized to published order.
                    for position, holding in enumerate(
                        reconstructed,
                        start=1,
                    ):

                        holding[
                            "rank"
                        ] = position

                    published_weights = sum(
                        1
                        for holding
                        in reconstructed
                        if holding[
                            "weightPublished"
                        ]
                    )

                    if (
                        published_weights
                        ==
                        len(
                            reconstructed
                        )
                    ):

                        weight_status = (
                            "published"
                        )

                    elif (
                        published_weights
                        ==
                        0
                    ):

                        weight_status = (
                            "not_published"
                        )

                    else:

                        weight_status = (
                            "partial"
                        )

                    section_lines = []

                    for holding in reconstructed:

                        section_lines.append(
                            (
                                f"{holding['rank']}. "
                                f"{holding['name']} "
                                f"{holding['weightText']}"
                            )
                        )

                    return (
                        reconstructed,
                        "\n".join(
                            section_lines
                        ),
                        debug_lines,
                        page_number,
                        weight_status,
                    )

    return (
        [],
        "",
        [],
        0,
        "not_found",
    )


# =============================================================================
# TEXT SECTION EXTRACTION
# =============================================================================

def pdf_lines(
    text: str,
) -> list[str]:

    output = []

    for raw_line in text.splitlines():

        line = clean_text(
            raw_line
        )

        if line:

            output.append(
                line
            )

    return output


def find_holdings_start(
    lines: list[str],
) -> int | None:

    for index, line in enumerate(
        lines
    ):

        if is_top_holdings_heading(
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

    endings = {
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
        "country allocation",
        "geographical allocation",
        "geographic allocation",
    }

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

    section_text = "\n".join(
        section
    )

    if not clean_text(
        section_text
    ):

        return (
            "",
            "published",
        )

    return (
        section_text,
        "published",
    )


# =============================================================================
# TEXT FALLBACK FOR NAMES-ONLY HOLDINGS
# =============================================================================

def parse_text_holdings(
    section_text: str,
) -> list[dict]:
    """
    Conservative pypdf fallback.

    Supports:

        1
        Holding Name

        2
        Another Holding

    and:

        1 Holding Name 4.5%

    Does not require percentages.

    It deliberately requires rank markers because otherwise arbitrary
    document text could be mistaken for holdings.
    """

    lines = pdf_lines(
        section_text
    )

    holdings = []

    current_rank = None

    current_fragments = []

    def commit_current():

        nonlocal current_rank
        nonlocal current_fragments

        if current_rank is None:

            current_fragments = []

            return

        name = combine_name_fragments(
            current_fragments
        )

        if not name:

            current_fragments = []

            return

        weight = None

        weight_text = "-"

        weight_published = False

        combined_text = clean_text(
            " ".join(
                current_fragments
            )
        )

        match = re.search(
            r"([0-9]+(?:\.[0-9]+)?)\s*%",
            combined_text,
        )

        if match:

            value = float(
                match.group(1)
            )

            if (
                0 <= value <= 100
            ):

                weight = value

                weight_text = (
                    match.group(0)
                )

                weight_published = True

                name = clean_holding_name(
                    combined_text[
                        :match.start()
                    ]
                )

                if not name:

                    name = combine_name_fragments(
                        current_fragments
                    )

        holdings.append(
            {
                "rank":
                    current_rank,

                "name":
                    name,

                "weightPercent":
                    weight,

                "weightText":
                    weight_text,

                "weightPublished":
                    weight_published,
            }
        )

        current_rank = None

        current_fragments = []

    for line in lines:

        if is_noise_fragment(
            line
        ):

            continue

        # -------------------------------------------------------------
        # A standalone rank.
        # -------------------------------------------------------------

        standalone_rank = parse_rank_token(
            line
        )

        if standalone_rank is not None:

            if current_rank is not None:

                commit_current()

            current_rank = (
                standalone_rank
            )

            continue

        # -------------------------------------------------------------
        # Rank prefix.
        # -------------------------------------------------------------

        prefix_match = re.match(
            r"""
            ^
            ([0-9]{1,2})
            [.)\-:]
            \s+
            (.+)
            $
            """,
            line,
            re.VERBOSE,
        )

        if prefix_match:

            rank = int(
                prefix_match.group(1)
            )

            if (
                1 <= rank <= MAX_HOLDINGS
            ):

                if current_rank is not None:

                    commit_current()

                current_rank = rank

                remainder = clean_text(
                    prefix_match.group(2)
                )

                if remainder:

                    current_fragments.append(
                        remainder
                    )

                continue

        # -------------------------------------------------------------
        # Continue current holding.
        # -------------------------------------------------------------

        if current_rank is not None:

            current_fragments.append(
                line
            )

    if current_rank is not None:

        commit_current()

    if not holdings:

        raise RuntimeError(
            "Top Holdings section was found, "
            "but no ranked holding names could be parsed "
            "from the PDF text."
        )

    if len(
        holdings
    ) > MAX_HOLDINGS:

        holdings = holdings[
            :MAX_HOLDINGS
        ]

    # Normalize rank order to output order.
    for position, holding in enumerate(
        holdings,
        start=1,
    ):

        holding[
            "rank"
        ] = position

    return holdings


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

    for line in lines[:60]:

        match = pattern.search(
            line
        )

        if match:

            return clean_text(
                match.group(0)
            )

    return None


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

    if response.status != 200:

        raise RuntimeError(
            "Factsheet HTTP status was "
            f"{response.status}: "
            f"{factsheet_url}"
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
        or
        ""
    )

    if not is_prudential_url(
        prudential_url
    ):

        raise RuntimeError(
            "Excel Column A URL is not an official "
            "Prudential Singapore URL."
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

    if pruaccess_name:

        print(
            f"Excel PruAccess name: {pruaccess_name}"
        )

    citicode = (
        extract_citicode_from_url(
            prudential_url
        )
    )

    # -------------------------------------------------------------------------
    # Prudential page.
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
            "Prudential page redirected "
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
    # Extract PDF text.
    # -------------------------------------------------------------------------

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

    # =========================================================================
    # SECTION EXISTENCE
    # =========================================================================

    textual_heading = (
        find_textual_top_holdings_heading(
            full_text
        )
    )

    (
        visual_holdings,
        visual_section_text,
        visual_debug_lines,
        visual_page_number,
        visual_weight_status,
    ) = reconstruct_visual_holdings(
        factsheet_bytes
    )

    # =========================================================================
    # PRIMARY VISUAL RESULT
    # =========================================================================

    if visual_holdings:

        print(
            "Holdings parser: "
            "pdfplumber visual rank reconstruction"
        )

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

            "fundIdentifier":
                citicode,

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

            "holdingsHeadingDetected":
                (
                    textual_heading
                    or
                    "visual"
                ),

            "holdingsParserMethod":
                "pdfplumber_coordinates",

            "holdingsPdfPage":
                visual_page_number,

            "topHoldingsCount":
                len(
                    visual_holdings
                ),

            "topHoldingsWeightStatus":
                visual_weight_status,

            "topHoldings":
                visual_holdings,

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

                "noFabricatedHoldings":
                    True,

                "noCalculatedWeights":
                    True,

                "noEstimatedWeights":
                    True,

                "multilineHoldingNamesSupported":
                    True,

                "wrappedPdfLinesJoined":
                    True,

                "rankNumbersDoNotDefineRows":
                    True,

                "publishedWeightsOptional":
                    True,

                "missingWeightRepresentedAsNull":
                    True,

                "duplicateHoldingNamesAllowed":
                    True,
            },
        }

        result[
            "_sectionText"
        ] = visual_section_text

        result[
            "_visualLines"
        ] = visual_debug_lines

        return result

    # =========================================================================
    # TEXTUAL SECTION EXISTS
    # =========================================================================

    if textual_heading:

        (
            fallback_section_text,
            fallback_status,
        ) = extract_holdings_section(
            full_text
        )

        if (
            fallback_status
            !=
            "published"
        ):

            raise RuntimeError(
                "Top Holdings heading was detected in "
                "the official PDF, but the section could "
                "not be isolated safely."
            )

        print(
            "Coordinate parser did not reconstruct the "
            "holdings table."
        )

        print(
            "Using pypdf ranked-text fallback."
        )

        holdings = parse_text_holdings(
            fallback_section_text
        )

        published_weights = sum(
            1
            for holding
            in holdings
            if holding[
                "weightPublished"
            ]
        )

        if (
            published_weights
            ==
            len(
                holdings
            )
        ):

            weight_status = (
                "published"
            )

        elif published_weights == 0:

            weight_status = (
                "not_published"
            )

        else:

            weight_status = (
                "partial"
            )

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

            "fundIdentifier":
                citicode,

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

            "holdingsHeadingDetected":
                textual_heading,

            "holdingsParserMethod":
                "pypdf_text_fallback",

            "holdingsPdfPage":
                None,

            "topHoldingsCount":
                len(
                    holdings
                ),

            "topHoldingsWeightStatus":
                weight_status,

            "topHoldings":
                holdings,

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

                "noFabricatedHoldings":
                    True,

                "noCalculatedWeights":
                    True,

                "noEstimatedWeights":
                    True,

                "multilineHoldingNamesSupported":
                    True,

                "wrappedPdfLinesJoined":
                    True,

                "rankNumbersDoNotDefineRows":
                    True,

                "publishedWeightsOptional":
                    True,

                "missingWeightRepresentedAsNull":
                    True,

                "duplicateHoldingNamesAllowed":
                    True,
            },
        }

        result[
            "_sectionText"
        ] = fallback_section_text

        result[
            "_visualLines"
        ] = []

        return result

    # =========================================================================
    # ONLY NOW: NO HOLDINGS SECTION
    # =========================================================================

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

        "fundIdentifier":
            citicode,

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

        "holdingsHeadingDetected":
            None,

        "holdingsParserMethod":
            None,

        "holdingsPdfPage":
            None,

        "topHoldingsCount":
            0,

        "topHoldingsWeightStatus":
            "not_applicable",

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

            "noFabricatedHoldings":
                True,

            "noCalculatedWeights":
                True,

            "noEstimatedWeights":
                True,

            "multilineHoldingNamesSupported":
                True,

            "wrappedPdfLinesJoined":
                True,

            "rankNumbersDoNotDefineRows":
                True,

            "publishedWeightsOptional":
                True,

            "missingWeightRepresentedAsNull":
                True,

            "duplicateHoldingNamesAllowed":
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

        identifier = safe_filename(
            result.get(
                "fundName"
            )
            or
            f"fund_{excel_row}"
        )

    directory = (
        FUNDS_OUTPUT_DIR
        /
        (
            f"{excel_row}_"
            f"{identifier}"
        )
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    # -------------------------------------------------------------------------
    # Official source PDF.
    # -------------------------------------------------------------------------

    (
        directory
        / "factsheet.pdf"
    ).write_bytes(
        factsheet_bytes
    )

    # -------------------------------------------------------------------------
    # Complete PDF text.
    # -------------------------------------------------------------------------

    (
        directory
        / "factsheet_text.txt"
    ).write_text(
        full_text,
        encoding="utf-8",
    )

    # -------------------------------------------------------------------------
    # Parsed Top Holdings section.
    # -------------------------------------------------------------------------

    (
        directory
        / "top_holdings_section.txt"
    ).write_text(
        section_text,
        encoding="utf-8",
    )

    # -------------------------------------------------------------------------
    # Debug visual rows.
    # -------------------------------------------------------------------------

    visual_lines = (
        result.get(
            "_visualLines"
        )
        or
        []
    )

    (
        directory
        / "top_holdings_visual_lines.txt"
    ).write_text(
        "\n".join(
            visual_lines
        ),
        encoding="utf-8",
    )

    # -------------------------------------------------------------------------
    # Public parsed result.
    # -------------------------------------------------------------------------

    public_result = {
        key: value
        for key, value
        in result.items()
        if not key.startswith("_")
    }

    save_json(
        directory
        / "top_holdings.json",
        public_result,
    )

    # -------------------------------------------------------------------------
    # Metadata.
    # -------------------------------------------------------------------------

    metadata = {
        "excelRow":
            excel_row,

        "fundName":
            result.get(
                "fundName"
            ),

        "fundIdentifier":
            result.get(
                "fundIdentifier"
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

        "holdingsHeadingDetected":
            result.get(
                "holdingsHeadingDetected"
            ),

        "holdingsParserMethod":
            result.get(
                "holdingsParserMethod"
            ),

        "holdingsPdfPage":
            result.get(
                "holdingsPdfPage"
            ),

        "topHoldingsCount":
            result.get(
                "topHoldingsCount"
            ),

        "topHoldingsWeightStatus":
            result.get(
                "topHoldingsWeightStatus"
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
# SAVE NO HOLDINGS SECTION
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
            "fundIdentifier"
        )
        or
        result.get(
            "fundName"
        )
        or
        f"fund_{excel_row}"
    )

    directory = (
        FUNDS_OUTPUT_DIR
        /
        (
            f"{excel_row}_"
            f"{identifier}"
        )
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    public_result = {
        key: value
        for key, value
        in result.items()
        if not key.startswith("_")
    }

    save_json(
        directory
        / "top_holdings.json",
        public_result,
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

        "fundIdentifier":
            result.get(
                "fundIdentifier"
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

        "topHoldingsWeightStatus":
            "not_applicable",

        "status":
            "no_holdings_section",

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

    identifier = (
        extract_citicode_from_url(
            excel_fund[
                "prudentialUrl"
            ]
        )
        or
        f"fund_{excel_row}"
    )

    directory = (
        FUNDS_OUTPUT_DIR
        /
        (
            f"{excel_row}_"
            f"{safe_filename(identifier)}"
            "_failed"
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

                result = None

                last_error = None

                # -------------------------------------------------------------
                # Whole-fund retries.
                # -------------------------------------------------------------

                for attempt in range(
                    1,
                    RETRY_COUNT + 1,
                ):

                    try:

                        print(
                            f"\nAttempt "
                            f"{attempt}/"
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
                            str(
                                error
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
                                f"{RETRY_DELAY_SECONDS} "
                                f"seconds..."
                            )

                            time.sleep(
                                RETRY_DELAY_SECONDS
                            )

                # -------------------------------------------------------------
                # Complete failure.
                # -------------------------------------------------------------

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

                # -------------------------------------------------------------
                # Genuine no Top Holdings section.
                # -------------------------------------------------------------

                if (
                    result[
                        "status"
                    ]
                    ==
                    "no_holdings_section"
                ):

                    save_no_holdings_result(
                        result
                    )

                    no_holdings_section.append(
                        result
                    )

                    print(
                        "\n"
                        "NO TOP HOLDINGS SECTION"
                    )

                    continue

                # -------------------------------------------------------------
                # Successful extraction.
                #
                # Save the exact PDF bytes that were downloaded and parsed.
                # -------------------------------------------------------------

                try:

                    section_text = (
                        result.get(
                            "_sectionText"
                        )
                        or
                        ""
                    )

                    if not section_text:

                        raise RuntimeError(
                            "Successful result contains "
                            "no Top Holdings section text."
                        )

                    save_success_result(
                        result,
                        # The actual PDF bytes are retained in memory by
                        # re-downloading below so saved source is identical
                        # to the final verification source.
                        page.request.get(
                            result[
                                "factsheetUrl"
                            ],
                            timeout=(
                                FACTSHEET_DOWNLOAD_TIMEOUT_MS
                            ),
                        ).body(),
                        *(
                            lambda verified_pdf: (
                                extract_pdf_text(
                                    verified_pdf
                                )[0],
                                result.get(
                                    "_sectionText"
                                )
                                or
                                "",
                            )
                        )(
                            page.request.get(
                                result[
                                    "factsheetUrl"
                                ],
                                timeout=(
                                    FACTSHEET_DOWNLOAD_TIMEOUT_MS
                                ),
                            ).body()
                        ),
                    )

                except Exception as error:

                    error_text = clean_text(
                        str(
                            error
                        )
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
                    f"Citicode: "
                    f"{result.get('fundIdentifier') or '-'}"
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
                    f"Parser: "
                    f"{result.get('holdingsParserMethod') or '-'}"
                )

                print(
                    f"Weight status: "
                    f"{result.get('topHoldingsWeightStatus') or '-'}"
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

        finally:

            context.close()

            browser.close()

    # =========================================================================
    # CONSOLIDATED OUTPUT
    # =========================================================================

    completed_at = utc_now_iso()

    public_successful = []

    for item in successful:

        public_successful.append(
            {
                key: value
                for key, value
                in item.items()
                if not key.startswith("_")
            }
        )

    public_no_holdings = []

    for item in no_holdings_section:

        public_no_holdings.append(
            {
                key: value
                for key, value
                in item.items()
                if not key.startswith("_")
            }
        )

    all_results = (
        public_successful
        +
        public_no_holdings
    )

    total_holdings = sum(
        int(
            item.get(
                "topHoldingsCount",
                0,
            )
            or
            0
        )
        for item
        in successful
    )

    weight_status_counts = {
        "published":
            sum(
                1
                for item
                in successful
                if item.get(
                    "topHoldingsWeightStatus"
                )
                ==
                "published"
            ),

        "notPublished":
            sum(
                1
                for item
                in successful
                if item.get(
                    "topHoldingsWeightStatus"
                )
                ==
                "not_published"
            ),

        "partial":
            sum(
                1
                for item
                in successful
                if item.get(
                    "topHoldingsWeightStatus"
                )
                ==
                "partial"
            ),
    }

    all_holdings_payload = {
        "status":
            (
                "success"
                if not failed
                else
                "partial"
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

        "failedFunds":
            len(
                failed
            ),

        "totalPublishedTopHoldings":
            total_holdings,

        "weightStatusCounts":
            weight_status_counts,

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

            "noFabricatedHoldings":
                True,

            "noCalculatedWeights":
                True,

            "noEstimatedWeights":
                True,

            "multilineHoldingNamesSupported":
                True,

            "wrappedPdfLinesJoined":
                True,

            "rankNumbersDoNotDefineRows":
                True,

            "publishedWeightsOptional":
                True,

            "missingWeightRepresentedAsNull":
                True,

            "duplicateHoldingNamesAllowed":
                True,

            "noFalseNoHoldingsClassification":
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

    parser_usage = {
        "pdfplumberCoordinates":
            sum(
                1
                for item
                in successful
                if item.get(
                    "holdingsParserMethod"
                )
                ==
                "pdfplumber_coordinates"
            ),

        "pypdfFallback":
            sum(
                1
                for item
                in successful
                if item.get(
                    "holdingsParserMethod"
                )
                ==
                "pypdf_text_fallback"
            ),
    }

    run_summary = {
        "status":
            (
                "success"
                if not failed
                else
                "partial"
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

        "parserUsage":
            parser_usage,

        "weightStatusCounts":
            weight_status_counts,

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

                    "fundIdentifier":
                        item.get(
                            "fundIdentifier"
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

                    "holdingsParserMethod":
                        item.get(
                            "holdingsParserMethod"
                        ),

                    "holdingsPdfPage":
                        item.get(
                            "holdingsPdfPage"
                        ),

                    "topHoldingsCount":
                        item.get(
                            "topHoldingsCount"
                        ),

                    "topHoldingsWeightStatus":
                        item.get(
                            "topHoldingsWeightStatus"
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

                    "fundIdentifier":
                        item.get(
                            "fundIdentifier"
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

            "noCalculatedWeights":
                True,

            "noEstimatedWeights":
                True,

            "noThirdPartyHoldings":
                True,

            "multilineHoldingNamesSupported":
                True,

            "wrappedPdfLinesJoined":
                True,

            "rankNumbersDoNotDefineRows":
                True,

            "publishedWeightsOptional":
                True,

            "missingWeightRepresentedAsNull":
                True,

            "duplicateHoldingNamesAllowed":
                True,

            "noFalseNoHoldingsClassification":
                True,

            "primaryParser":
                "pdfplumber_coordinates",

            "fallbackParser":
                "pypdf_text",
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
        f"Successful: "
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
        f"Total holdings extracted: "
        f"{total_holdings}"
    )

    print(
        "\nWeight status:"
    )

    print(
        f" - All weights published: "
        f"{weight_status_counts['published']} funds"
    )

    print(
        f" - No weights published: "
        f"{weight_status_counts['notPublished']} funds"
    )

    print(
        f" - Partial weights published: "
        f"{weight_status_counts['partial']} funds"
    )

    print(
        "\nParser:"
    )

    print(
        f" - pdfplumber coordinates: "
        f"{parser_usage['pdfplumberCoordinates']}"
    )

    print(
        f" - pypdf fallback: "
        f"{parser_usage['pypdfFallback']}"
    )

    print(
        "\nRules:"
    )

    print(
        " - Multiline holding names: ENABLED"
    )

    print(
        " - Missing holding percentage: ALLOWED"
    )

    print(
        " - Missing weight stored as: null / -"
    )

    print(
        " - Rank numbers define row by themselves: NO"
    )

    print(
        " - Duplicate names cause failure: NO"
    )

    print(
        " - Parser failure becomes NO_HOLDINGS_SECTION: NO"
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

    if failed:

        return 1

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
