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

Workflow:

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
    Extract every published holding and percentage
          |
          v
    If primary parsing fails, automatically retry the same official PDF
    with the fallback parser
          |
          v
    Save one result per fund


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
- Holdings are stored in the order published by Prudential.
- Multi-line holding names are joined into one holding name.
- Wrapped PDF text lines are treated as part of the same holding until the
  published percentage is encountered.
- Duplicate holding names are allowed.
- Duplicate holding percentages are allowed.
- If the primary parser fails, the same official factsheet section is retried
  automatically with the fallback parser.
- The fallback parser uses the LAST percentage on a logical line as the
  portfolio weight, preserving earlier coupon/rate percentages inside the
  security name.
- If both parsers fail, the fund is marked FAILED.
- If a factsheet has no Top 10 holdings section at all, the fund is marked
  as NO_HOLDINGS_SECTION.
- Failed funds do not get blank fabricated holdings.
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
        <excelRow>_<citicode>/
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
from io import BytesIO
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
    """
    Normalize whitespace while preserving meaningful text.

    Common PDF extraction artifacts:
        - non-breaking spaces
        - repeated whitespace
        - pipe separators
        - the literal word "None"
    are handled elsewhere where necessary.
    """

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
    """
    Case-insensitive comparison helper.
    """

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
    """
    Locate the official Fund Factsheet link from the official Prudential page.

    Ranking:

        1. Anchor text contains "Fund Factsheet"
        2. Anchor text contains "Factsheet"
        3. Href contains "factsheet"
        4. PDF extension
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
# PDF TEXT
# =============================================================================

def extract_pdf_text(
    pdf_bytes: bytes,
) -> tuple[str, int]:
    """
    Extract all available text from the official PDF.
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
# SPATIAL PDF FALLBACK
# =============================================================================

def _positioned_pdf_lines(
    pdf_bytes: bytes,
) -> list[dict]:
    """
    Extract PDF text together with its physical x/y position.

    This is a FALLBACK ONLY. The normal text parser remains unchanged and is
    always attempted first.

    The purpose is to recover Prudential factsheets where the PDF's visual
    columns are interleaved by normal text extraction, such as:

        Top 10 Holdings       Dividend History
        Company A      8.9%   Date
        Company B      5.8%   1.50%

    The returned records are page-local visual lines. No holding is inferred
    here; this function only reconstructs the visible PDF layout.
    """

    reader = PdfReader(BytesIO(pdf_bytes))

    all_lines = []

    for page_number, pdf_page in enumerate(reader.pages, start=1):
        fragments = []

        def visitor_text(text, cm, tm, font_dict, font_size):
            if text is None:
                return

            raw = str(text)
            if not raw.strip():
                return

            try:
                x = float(tm[4])
                y = float(tm[5])
            except Exception:
                return

            # pypdf normally gives one visual line plus a trailing newline.
            # Split embedded newlines so that each visible text unit can be
            # grouped independently.
            parts = raw.splitlines()
            if not parts:
                parts = [raw]

            for part_index, part in enumerate(parts):
                part = clean_text(part)
                if not part:
                    continue

                adjusted_y = y
                if part_index:
                    try:
                        adjusted_y = y - (float(font_size or 8) * part_index * 1.15)
                    except Exception:
                        adjusted_y = y

                fragments.append(
                    {
                        "x0": x,
                        "x1": x,
                        "y": adjusted_y,
                        "text": part,
                    }
                )

        try:
            pdf_page.extract_text(visitor_text=visitor_text)
        except Exception:
            # Some PDFs do not support visitor extraction cleanly. The normal
            # parser has already failed before this fallback is called, so a
            # clear fallback failure is preferable to fabricated data.
            continue

        # Group fragments that share the same visual baseline.
        groups = []
        y_tolerance = 3.0

        for fragment in sorted(
            fragments,
            key=lambda item: (-item["y"], item["x0"]),
        ):
            target = None

            for group in groups:
                if abs(group["y"] - fragment["y"]) <= y_tolerance:
                    target = group
                    break

            if target is None:
                target = {
                    "page": page_number,
                    "y": fragment["y"],
                    "fragments": [],
                }
                groups.append(target)

            target["fragments"].append(fragment)
            target["y"] = sum(
                item["y"] for item in target["fragments"]
            ) / len(target["fragments"])

        for group in groups:
            fragments_sorted = sorted(
                group["fragments"],
                key=lambda item: item["x0"],
            )

            text_parts = []
            x0 = None
            x1 = None

            for fragment in fragments_sorted:
                text_parts.append(fragment["text"])
                x0 = fragment["x0"] if x0 is None else min(x0, fragment["x0"])
                x1 = fragment["x1"] if x1 is None else max(x1, fragment["x1"])

            line_text = clean_text(" ".join(text_parts))
            if not line_text:
                continue

            all_lines.append(
                {
                    "page": page_number,
                    "y": group["y"],
                    "x0": x0 or 0.0,
                    "x1": x1 or 0.0,
                    "text": line_text,
                    "fragments": fragments_sorted,
                }
            )

    return all_lines


def _is_spatial_top_holdings_heading(text: str) -> bool:
    normalized = normalize_text(text)
    return bool(
        re.search(
            r"\btop\s+(?:10|ten)\s+holdings",
            normalized,
            re.IGNORECASE,
        )
    )


def _spatial_column_boundary(
    page_lines: list[dict],
    heading: dict,
    page_width: float,
) -> float:
    """
    Find the right edge of the Top Holdings visual column.

    Prudential often places Dividend History immediately beside Top Holdings.
    We use the physical position of that adjacent heading/header when it is
    available. Otherwise we use a conservative page-relative boundary.
    """

    heading_x = heading["x0"]
    heading_y = heading["y"]

    explicit_right_headers = []

    for line in page_lines:
        if abs(line["y"] - heading_y) > 45:
            continue

        for fragment in line.get("fragments", []):
            if fragment["x0"] <= heading_x + 20:
                continue

            normalized = normalize_text(fragment["text"])

            if any(
                marker in normalized
                for marker in (
                    "dividend history",
                    "distribution history",
                    "date",
                    "frequency",
                    "performance history",
                )
            ):
                explicit_right_headers.append(fragment["x0"])

    if explicit_right_headers:
        return min(explicit_right_headers) - 8.0

    # Look for a strong horizontal gap in the lines immediately below the
    # heading. This catches two-column layouts whose right header is extracted
    # without a useful label.
    nearby_x = []
    for line in page_lines:
        if heading_y - 70 <= line["y"] <= heading_y - 5:
            for fragment in line.get("fragments", []):
                if fragment["x0"] > heading_x + 20:
                    nearby_x.append(fragment["x0"])

    if nearby_x:
        candidate = min(nearby_x)
        if candidate > heading_x + page_width * 0.18:
            return candidate - 8.0

    # Final conservative fallback: keep the left 68% of the page. This is
    # deliberately only used when the PDF exposes no usable second-column
    # marker.
    return max(
        heading_x + 180.0,
        page_width * 0.68,
    )


def _build_spatial_holdings_section(
    page_lines: list[dict],
    heading: dict,
    page_width: float,
) -> str:
    """
    Build a visual-order Top Holdings section from one page.
    """

    left_edge = max(0.0, heading["x0"] - 15.0)
    right_edge = _spatial_column_boundary(
        page_lines,
        heading,
        page_width,
    )

    selected = []

    # Text below the heading in PDF coordinates has a lower y value.
    for line in page_lines:
        if line["y"] >= heading["y"] - 2.0:
            continue

        if line["y"] < heading["y"] - 420.0:
            continue

        if line["x1"] < left_edge:
            continue

        if line["x0"] >= right_edge:
            continue

        cropped_fragments = []
        for fragment in line.get("fragments", []):
            if fragment["x1"] < left_edge:
                continue
            if fragment["x0"] >= right_edge:
                continue
            cropped_fragments.append(fragment)

        if not cropped_fragments:
            continue

        cropped_fragments.sort(key=lambda item: item["x0"])
        cropped_text = clean_text(
            " ".join(fragment["text"] for fragment in cropped_fragments)
        )

        if not cropped_text:
            continue

        normalized = normalize_text(cropped_text)

        # A second Top Holdings heading begins a separate visual candidate.
        # This matters for factsheets where the normal PDF text order places
        # the real holdings table before/after an unrelated chart section.
        if normalized in {
            "dividend history",
            "distribution history",
            "date",
            "frequency",
        }:
            continue

        if _is_spatial_top_holdings_heading(cropped_text):
            break

        if is_holdings_end(cropped_text):
            break

        cropped_line = dict(line)
        cropped_line["text"] = cropped_text
        cropped_line["fragments"] = cropped_fragments
        cropped_line["x0"] = min(fragment["x0"] for fragment in cropped_fragments)
        cropped_line["x1"] = max(fragment["x1"] for fragment in cropped_fragments)
        selected.append(cropped_line)

    selected.sort(key=lambda item: -item["y"])

    # Keep the section compact. Once ten visible weight-bearing rows have been
    # encountered, subsequent chart/allocation material is not part of Top
    # Holdings.
    output = []
    weight_rows = 0

    for line in selected:
        text = clean_text(line["text"])
        if not text:
            continue

        output.append(text)

        if find_last_percentage_in_line(text) is not None:
            weight_rows += 1

        if weight_rows >= MAX_HOLDINGS:
            break

    return "\n".join(output)


def extract_holdings_spatial_fallback(
    pdf_bytes: bytes,
) -> tuple[list[dict], str]:
    """
    Recover Top Holdings from the physical PDF layout.

    This fallback scans every page containing a visible "Top 10 Holdings"
    heading, reconstructs the left visual column, and then applies the
    existing fallback parser to that reconstructed section.

    A candidate is accepted only if the resulting holdings pass the same
    strict validation as the normal parser. No names or weights are guessed.
    """

    positioned = _positioned_pdf_lines(pdf_bytes)

    if not positioned:
        raise RuntimeError(
            "Spatial PDF fallback could not extract positioned text."
        )

    candidates = []

    pages = sorted({line["page"] for line in positioned})

    for page_number in pages:
        page_lines = [
            line
            for line in positioned
            if line["page"] == page_number
        ]

        if not page_lines:
            continue

        max_x = max(line["x1"] for line in page_lines)
        min_x = min(line["x0"] for line in page_lines)
        page_width = max(595.0, max_x + 20.0, min_x + 595.0)

        headings = [
            line
            for line in page_lines
            if _is_spatial_top_holdings_heading(line["text"])
        ]

        for heading in headings:
            section = _build_spatial_holdings_section(
                page_lines,
                heading,
                page_width,
            )

            if not clean_text(section):
                continue

            try:
                holdings = parse_holdings_fallback(section)
            except Exception as error:
                candidates.append(
                    {
                        "page": page_number,
                        "section": section,
                        "holdings": None,
                        "error": clean_text(str(error)),
                    }
                )
                continue

            # A valid candidate is scored by published holding count and the
            # presence of names/weights. We do not rank by fund-specific
            # expectations.
            candidates.append(
                {
                    "page": page_number,
                    "section": section,
                    "holdings": holdings,
                    "error": None,
                }
            )

    valid = [
        candidate
        for candidate in candidates
        if candidate.get("holdings")
    ]

    if not valid:
        errors = [
            candidate.get("error")
            for candidate in candidates
            if candidate.get("error")
        ]

        detail = (
            "; ".join(errors[:3])
            if errors
            else "No usable Top Holdings visual candidate was found."
        )

        raise RuntimeError(
            "Spatial PDF fallback failed: "
            f"{detail}"
        )

    # Prefer the candidate with the most published holdings. If two candidates
    # contain the same count, prefer the one with more substantial name text.
    valid.sort(
        key=lambda candidate: (
            len(candidate["holdings"]),
            sum(len(item["name"]) for item in candidate["holdings"]),
        ),
        reverse=True,
    )

    selected = valid[0]

    print(
        "Spatial PDF fallback recovered Top Holdings "
        f"from page {selected['page']} with "
        f"{len(selected['holdings'])} holdings."
    )

    return (
        selected["holdings"],
        selected["section"],
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
    """
    Extract explicit factsheet data date when present.

    Examples:

        Data as at 30 Sep 2025
        All data as at 28 Feb 2026
    """

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
    """
    Attempt to capture a visible month/year document date.

    This is an extraction only.

    No date is fabricated if unavailable.
    """

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

    # Search beginning of document first.
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
) -> tuple[
    str,
    str,
]:
    """
    Return:

        section_text
        section_status

    Status:

        published
        not_published
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

    # Remove common bullet characters.
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

    # Remove pipe separators introduced by PDF table extraction.
    name = name.replace(
        "|",
        " ",
    )

    # Remove the literal PDF artifact "None".
    name = re.sub(
        r"\bNone\b",
        " ",
        name,
        flags=re.IGNORECASE,
    )

    name = clean_text(
        name
    )

    # Remove trailing separators.
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


def find_percentage_in_line(
    line: str,
) -> tuple[
    float,
    str,
    int,
    int,
] | None:
    """
    Find the FIRST published percentage candidate in a line.

    This is retained as the PRIMARY parser behavior.
    The fallback parser below uses the LAST percentage on a logical line,
    which is required for factsheets where security coupon rates appear
    before the portfolio holding weight.
    """

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
        percentage = float(match.group(1))
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
    """
    Find the LAST percentage in a line.

    Prudential fixed-income factsheets can contain security coupon/rate
    percentages inside the holding description, followed by the actual
    portfolio weight. Example:

        INDIA (REPUBLIC OF) 7.09% 5-AUG-2054 2.7%

    The final percentage is the portfolio holding weight. Earlier
    percentages remain part of the security name.
    """

    if not line:
        return None

    matches = list(re.finditer(
        r"(?<![\d.])([0-9]+(?:\.[0-9]+)?)\s*%",
        line,
    ))

    if not matches:
        return None

    match = matches[-1]

    try:
        percentage = float(match.group(1))
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
# HOLDING RANK
# =============================================================================

def extract_leading_rank(
    line: str,
) -> tuple[
    int | None,
    str,
]:
    """
    Extract a leading holding rank when a real rank marker is present.

    Supported forms:

        1 NAME
        1. NAME
        1) NAME
        1 - NAME
        1: NAME
        1

    IMPORTANT:
    ---------
    A percentage line such as ``5.5% 31-DEC-2079 2.5%`` must NOT be
    interpreted as holding rank 5.  Therefore punctuation-based ranks
    require whitespace after the punctuation, and whitespace-based ranks
    require whitespace before the holding text.
    """

    text = clean_text(line)

    if not text:
        return None, ""

    # A standalone integer is a possible rank line.
    standalone = re.fullmatch(r"(\d{1,2})", text)
    if standalone:
        rank = int(standalone.group(1))
        if 1 <= rank <= 99:
            return rank, ""
        return None, text

    # Explicit punctuation rank markers.  The required whitespace after the
    # marker prevents decimal percentages such as ``5.5%`` from matching.
    match = re.match(
        r"^\s*(\d{1,2})[.)\-:]\s+(.+)$",
        text,
    )
    if match:
        rank = int(match.group(1))
        if 1 <= rank <= 99:
            return rank, clean_text(match.group(2))
        return None, text

    # Space-separated rank, e.g. ``1 INDIA ...``.
    match = re.match(
        r"^\s*(\d{1,2})\s+(.+)$",
        text,
    )
    if match:
        rank = int(match.group(1))
        if 1 <= rank <= 99:
            return rank, clean_text(match.group(2))

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
    """
    Clean one fragment of a potentially wrapped holding name.
    """

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
    """
    Join wrapped PDF lines into one clean holding name.
    """

    cleaned_fragments = []

    for fragment in fragments:

        fragment = clean_holding_fragment(
            fragment
        )

        if not fragment:

            continue

        # Do not accidentally preserve standalone rank lines.
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
# HOLDINGS PARSER
# =============================================================================

def parse_holdings(
    section_text: str,
) -> list[dict]:
    """
    Robust parser for Prudential Top Holdings sections.

    IMPORTANT:

    The parser DOES NOT assume that a holding fits on one PDF text line.

    Example:

        1
        JPMORGAN FUNDS - EMERGING
        MARKETS EQUITY FUND
        14.7%

    becomes:

        rank = 1
        name = "JPMORGAN FUNDS - EMERGING MARKETS EQUITY FUND"
        weight = 14.7%

    Also supports:

        1 JPMORGAN FUNDS - EMERGING
        MARKETS EQUITY FUND 14.7%

    And:

        JPMORGAN FUNDS - EMERGING
        MARKETS EQUITY FUND | 14.7%

    The percentage is used as the definitive boundary for the holding.

    No holding is created until a published percentage is found.

    Duplicate holding names are explicitly allowed.

    Duplicate holding percentages are explicitly allowed.
    """

    if not clean_text(
        section_text
    ):

        raise RuntimeError(
            "Top 10 holdings section is empty."
        )

    lines = pdf_lines(
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

        name = (
            combine_holding_name_fragments(
                pending_fragments
            )
        )

        if not name:

            raise RuntimeError(
                "A published holding percentage was found "
                "but no holding name could be extracted."
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

        pending_fragments = []

        pending_rank = None

    # -------------------------------------------------------------------------
    # Process sequentially.
    #
    # Wrapped lines are accumulated until a percentage is encountered.
    #
    # IMPORTANT:
    # Duplicate holding names are allowed.
    # Duplicate percentages are allowed.
    # -------------------------------------------------------------------------

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

        # -------------------------------------------------------------
        # Check if the line begins with a rank.
        # -------------------------------------------------------------

        detected_rank, remainder = (
            extract_leading_rank(
                line
            )
        )

        if detected_rank is not None:

            # A new rank appearing while the current holding has not yet
            # received a percentage indicates a structure we cannot safely
            # interpret.
            if (
                pending_fragments
                and pending_rank is not None
            ):

                raise RuntimeError(
                    "A new holding rank appeared before "
                    "the previous holding received a published percentage. "
                    f"Previous rank={pending_rank}, "
                    f"new rank={detected_rank}."
                )

            if (
                pending_fragments
                and pending_rank is None
            ):

                raise RuntimeError(
                    "A new holding rank appeared before "
                    "the previous holding received a published percentage."
                )

            pending_rank = (
                detected_rank
            )

            line = remainder

            # A standalone rank such as "1" carries no name.
            if not line:

                continue

        # -------------------------------------------------------------
        # Look for a percentage anywhere on the line.
        # -------------------------------------------------------------

        percentage_info = (
            find_percentage_in_line(
                line
            )
        )

        # If a logical line contains multiple percentages, the primary
        # parser deliberately refuses to guess which one is the portfolio
        # weight. This sends the fund to the automatic fallback parser,
        # which uses the final percentage as the published holding weight.
        percentage_candidates = re.findall(
            r"(?<![\d.])([0-9]+(?:\.[0-9]+)?)\s*%",
            line,
        )

        if len(percentage_candidates) > 1:
            raise RuntimeError(
                "Primary parser found multiple percentages on one logical "
                "line and will not guess the holding weight. "
                "Fallback parser required."
            )

        if percentage_info is not None:

            (
                percentage,
                percentage_text,
                start_index,
                _end_index,
            ) = percentage_info

            # Everything before the percentage belongs to the holding
            # name. This may be only the final wrapped line, or the entire
            # same-line name.
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

        # -------------------------------------------------------------
        # No percentage on this line.
        #
        # Therefore this is either:
        #
        #   - a wrapped holding-name line
        #   - a PDF table artifact
        # -------------------------------------------------------------

        fragment = clean_holding_fragment(
            line
        )

        if not fragment:

            continue

        # Ignore table column labels that survived PDF extraction.
        if is_holding_header_or_noise(
            fragment
        ):

            continue

        pending_fragments.append(
            fragment
        )

    # -------------------------------------------------------------------------
    # Any unfinished text is not a valid published holding because no
    # published percentage was found.
    #
    # We deliberately do NOT fabricate or guess its weight.
    # -------------------------------------------------------------------------

    if not holdings:

        raise RuntimeError(
            "Top 10 holdings section was found, "
            "but no holding/percentage pairs could be parsed."
        )

    if len(
        holdings
    ) > MAX_HOLDINGS:

        raise RuntimeError(
            "Parser produced more than 10 holdings."
        )

    # -------------------------------------------------------------------------
    # Validate published order.
    # -------------------------------------------------------------------------

    ranks = [
        holding[
            "rank"
        ]
        for holding in holdings
    ]

    # If Prudential provides explicit ranks, they must be sequential.
    explicit_ranks_present = any(
        holding[
            "rank"
        ] != position
        for position, holding
        in enumerate(
            holdings,
            start=1,
        )
    )

    if explicit_ranks_present:

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
    # Validate names and weights.
    #
    # IMPORTANT:
    # Duplicate names are allowed.
    # Duplicate percentages are allowed.
    # Only validity of each individual name and weight is checked.
    # -------------------------------------------------------------------------

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

        weight = holding.get(
            "weightPercent"
        )

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

        # Write back cleaned name only.
        holding[
            "name"
        ] = name

    return holdings


# =============================================================================
# FALLBACK HOLDINGS PARSER
# =============================================================================

def parse_holdings_fallback(
    section_text: str,
) -> list[dict]:
    """
    Alternative parser used after the primary parser fails.

    Key difference from the primary parser:
    - uses the LAST percentage on each logical line;
    - therefore preserves coupon/rate percentages embedded in fixed-income
      security names and uses the final percentage as the portfolio weight;
    - supports wrapped names and standalone weight lines;
    - never fabricates a weight;
    - duplicate names and duplicate percentages are allowed.
    """

    if not clean_text(section_text):
        raise RuntimeError("Top 10 holdings section is empty.")

    lines = pdf_lines(section_text)
    holdings = []
    pending_fragments: list[str] = []
    pending_rank: int | None = None

    def commit_pending(percentage: float, percentage_text: str) -> None:
        nonlocal pending_fragments, pending_rank

        name = combine_holding_name_fragments(pending_fragments)
        if not name:
            raise RuntimeError(
                "Fallback parser found a published holding percentage "
                "but no holding name could be extracted."
            )

        if len(holdings) >= MAX_HOLDINGS:
            raise RuntimeError("Fallback parser produced more than 10 holdings.")

        rank = pending_rank if pending_rank is not None else len(holdings) + 1
        holdings.append({
            "rank": rank,
            "name": name,
            "weightPercent": percentage,
            "weightText": percentage_text,
        })
        pending_fragments = []
        pending_rank = None

    for raw_line in lines:
        line = clean_text(raw_line)
        if not line or is_holding_header_or_noise(line):
            continue

        detected_rank, remainder = extract_leading_rank(line)
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

        percentage_info = find_last_percentage_in_line(line)
        if percentage_info is not None:
            percentage, percentage_text, start_index, _end_index = percentage_info
            name_fragment = clean_text(line[:start_index])
            if name_fragment:
                pending_fragments.append(name_fragment)
            commit_pending(percentage, percentage_text)
            if len(holdings) >= MAX_HOLDINGS:
                break
            continue

        fragment = clean_holding_fragment(line)
        if fragment and not is_holding_header_or_noise(fragment):
            pending_fragments.append(fragment)

    if not holdings:
        raise RuntimeError(
            "Fallback parser found the Top 10 holdings section but could not "
            "extract any holding/percentage pairs."
        )

    ranks = [holding["rank"] for holding in holdings]
    expected_ranks = list(range(1, len(holdings) + 1))
    if ranks != expected_ranks:
        raise RuntimeError(
            "Fallback holding ranks are not sequential. "
            f"Parsed={ranks}; Expected={expected_ranks}"
        )

    for holding in holdings:
        name = clean_holding_name(holding.get("name"))
        weight = holding.get("weightPercent")
        if not name:
            raise RuntimeError("Fallback holding name is empty.")
        if not isinstance(weight, (int, float)) or weight < 0 or weight > 100:
            raise RuntimeError("Fallback holding weight is invalid.")
        holding["name"] = name

    return holdings


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
            "Factsheet response contained "
            "zero bytes."
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
    """
    Extract one complete Prudential factsheet holdings result.
    """

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

        # Some pages never become fully idle because of analytics.
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

        # Conservative fallback for fund names.
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
            },
        }

        return result

    # -------------------------------------------------------------------------
    # A Top 10 holdings section exists.
    #
    # PRIMARY parser runs first. If it fails, automatically retry the SAME
    # official factsheet section with the fallback parser.
    # -------------------------------------------------------------------------

    parser_used = "primary"
    primary_parser_error = None
    fallback_parser_error = None
    selected_section_text = section_text

    try:
        holdings = parse_holdings(section_text)
    except Exception as primary_error:
        primary_parser_error = clean_text(str(primary_error))
        print("Primary holdings parser failed; running fallback parser...")
        print(f"Primary parser error: {primary_parser_error}")

        try:
            holdings = parse_holdings_fallback(section_text)
            parser_used = "fallback"
        except Exception as fallback_error:
            fallback_parser_error = clean_text(str(fallback_error))
            print("Text fallback parser failed; running spatial PDF fallback...")
            print(f"Text fallback parser error: {fallback_parser_error}")

            holdings, selected_section_text = extract_holdings_spatial_fallback(
                factsheet_bytes
            )
            parser_used = "spatial_fallback"

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

        "fallbackParserError":
            fallback_parser_error,

        "holdingsSectionExtraction":
            (
                "pdf_coordinates"
                if parser_used == "spatial_fallback"
                else "pdf_text"
            ),

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

            "percentageDefinesHoldingBoundary":
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

        # Use the final URL if identifier is unavailable.
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

    # Official PDF copy.
    (
        directory
        / "factsheet.pdf"
    ).write_bytes(
        factsheet_bytes
    )

    # Complete extracted text for audit/debugging.
    (
        directory
        / "factsheet_text.txt"
    ).write_text(
        full_text,
        encoding="utf-8",
    )

    # Exact Top Holdings section extracted by the parser.
    (
        directory
        / "top_holdings_section.txt"
    ).write_text(
        section_text,
        encoding="utf-8",
    )

    # Parsed holdings.
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

    # -------------------------------------------------------------------------
    # Excel
    # -------------------------------------------------------------------------

    funds = read_excel_funds()

    successful = []

    no_holdings_section = []

    failed = []

    fallback_recovered = []

    # -------------------------------------------------------------------------
    # Browser
    # -------------------------------------------------------------------------

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

            # =================================================================
            # ALL EXCEL FUNDS
            # =================================================================

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

                # -------------------------------------------------------------
                # Retry whole fund extraction.
                # -------------------------------------------------------------

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

                # -------------------------------------------------------------
                # Total failure.
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
                # No Top Holdings section.
                # -------------------------------------------------------------

                if (
                    result[
                        "status"
                    ]
                    ==
                    "no_holdings_section"
                ):

                    # There is intentionally no fake PDF/text output in this
                    # branch. The result itself records that Prudential did
                    # not publish a Top Holdings section.

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

                # -------------------------------------------------------------
                # SUCCESS
                #
                # Re-open the official PDF so we can save the exact source
                # document and extracted text.
                # -------------------------------------------------------------

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

                    # Re-parse during final verification using the parser
                    # selected during extraction. Spatial fallback must use the
                    # physical PDF layout again; it must never silently fall
                    # back to the interleaved plain-text section.
                    verification_parser = result.get("holdingsParser") or "primary"
                    verification_section_text = section_text

                    if verification_parser == "spatial_fallback":
                        (
                            verified_holdings,
                            verification_section_text,
                        ) = extract_holdings_spatial_fallback(
                            factsheet_bytes
                        )

                    elif verification_parser == "fallback":
                        verified_holdings = parse_holdings_fallback(section_text)

                    else:
                        try:
                            verified_holdings = parse_holdings(section_text)
                        except Exception as verification_primary_error:
                            print(
                                "Primary verification parser failed; "
                                "running fallback parser..."
                            )

                            try:
                                verified_holdings = parse_holdings_fallback(section_text)
                                result["holdingsParser"] = "fallback"
                                result["primaryParserError"] = clean_text(
                                    str(verification_primary_error)
                                )
                            except Exception as verification_fallback_error:
                                print(
                                    "Text fallback verification failed; "
                                    "running spatial PDF fallback..."
                                )
                                (
                                    verified_holdings,
                                    verification_section_text,
                                ) = extract_holdings_spatial_fallback(
                                    factsheet_bytes
                                )
                                result["holdingsParser"] = "spatial_fallback"
                                result["primaryParserError"] = clean_text(
                                    str(verification_primary_error)
                                )
                                result["fallbackParserError"] = clean_text(
                                    str(verification_fallback_error)
                                )

                    section_text = verification_section_text

                    if len(
                        verified_holdings
                    ) != result[
                        "topHoldingsCount"
                    ]:

                        raise RuntimeError(
                            "Holding count changed during "
                            "final verification."
                        )

                    def holding_signature(items):
                        return [
                            (
                                int(item.get("rank", 0)),
                                clean_text(item.get("name")),
                                float(item.get("weightPercent")),
                            )
                            for item in items
                        ]

                    if holding_signature(verified_holdings) != holding_signature(
                        result.get("topHoldings", [])
                    ):

                        raise RuntimeError(
                            "Holding names or published weights changed during "
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

                if result.get("holdingsParser") in {
                    "fallback",
                    "spatial_fallback",
                }:
                    fallback_recovered.append(result)

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
                    f"Holdings parser: "
                    f"{result.get('holdingsParser') or '-'}"
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
    # ALL-HOLDINGS CONSOLIDATED FILE
    # =========================================================================

    completed_at = utc_now_iso()

    all_results = (
        successful
        + no_holdings_section
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

        "failedFunds":
            len(
                failed
            ),

        "fallbackRecoveredFunds":
            len(
                fallback_recovered
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

            "percentageDefinesHoldingBoundary":
                True,

            "automaticFallbackParser":
                True,

            "fallbackUsesLastPercentageOnLogicalLine":
                True,

            "spatialPdfFallback":
                True,

            "spatialFallbackUsesOfficialPdfCoordinates":
                True,

            "spatialFallbackOnlyAfterTextParsersFail":
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

        "fallbackRecoveredFunds":
            len(
                fallback_recovered
            ),

        "fallbackRecoveredFundsDetail":
            [
                {
                    "excelRow": item.get("excelRow"),
                    "fundName": item.get("fundName"),
                    "holdingsParser": item.get("holdingsParser"),
                    "holdingsSectionExtraction": item.get("holdingsSectionExtraction"),
                    "primaryParserError": item.get("primaryParserError"),
                    "fallbackParserError": item.get("fallbackParserError"),
                    "factsheetUrl": item.get("factsheetUrl"),
                    "topHoldingsCount": item.get("topHoldingsCount"),
                }
                for item in fallback_recovered
            ],

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

            "percentageDefinesHoldingBoundary":
                True,

            "automaticFallbackParser":
                True,

            "fallbackUsesLastPercentageOnLogicalLine":
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
        f"Recovered by fallback parser: "
        f"{len(fallback_recovered)}"
    )

    print(
        f"Total published holdings extracted: "
        f"{total_holdings}"
    )

    print(
        "\nMultiline holding names: ENABLED"
    )

    print(
        "Wrapped PDF lines: JOINED"
    )

    print(
        "Duplicate holding names: ALLOWED"
    )

    print(
        "Duplicate holding percentages: ALLOWED"
    )

    print(
        "Holding boundary: PUBLISHED PERCENTAGE"
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
