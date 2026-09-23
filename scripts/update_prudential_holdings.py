#!/usr/bin/env python3

"""
VGrat FMS - Prudential Holdings FAILED-FUND RECOVERY TEST

PURPOSE
=======

This script is a SEPARATE recovery/test engine for the frozen
test_prudential_holdings.py baseline.

It does NOT modify:

    scripts/test_prudential_holdings.py
    scripts/test_pruaccess.py

It does NOT overwrite:

    output_holdings/all_holdings.json
    output_holdings/run_summary.json

It does NOT create:

    data.json
    index.html
    css/*
    js/*

Instead it reads the existing baseline results:

    output_holdings/run_summary.json

and dynamically identifies the funds whose baseline extraction failed.

For each failed fund:

    run_summary.json
            |
            v
    failed Excel row
            |
            v
    Funds Links.xlsm
            |
            v
    official Prudential URL
            |
            v
    official Prudential fund page
            |
            v
    official factsheet PDF
            |
            v
    physical PDF coordinate extraction
            |
            v
    visual Top Holdings reconstruction
            |
            v
    strict validation
            |
            v
    second official PDF verification
            |
            v
    output_holdings_recovery/

IMPORTANT
=========

This is a TEST / RECOVERY collector only.

Successful recovery results are NOT merged into the baseline yet.

No GitHub commit/push is performed by this script.

No baseline output is overwritten.
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

EXCEL_FILE = Path("Funds Links.xlsm")

BASELINE_OUTPUT_DIR = Path("output_holdings")

BASELINE_RUN_SUMMARY = (
    BASELINE_OUTPUT_DIR / "run_summary.json"
)

# IMPORTANT:
# Recovery output is completely separate from the frozen baseline.
RECOVERY_OUTPUT_DIR = Path(
    "output_holdings_recovery"
)

RECOVERY_FUNDS_DIR = (
    RECOVERY_OUTPUT_DIR / "funds"
)

RECOVERY_SUMMARY_FILE = (
    RECOVERY_OUTPUT_DIR / "run_summary.json"
)

RECOVERY_ALL_HOLDINGS_FILE = (
    RECOVERY_OUTPUT_DIR / "recovered_holdings.json"
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
# Official Prudential Singapore only
# -----------------------------------------------------------------------------

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

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def normalize_text(value) -> str:
    text = clean_text(value).casefold()

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
    text = clean_text(value)

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


def is_prudential_url(url: str) -> bool:
    try:
        parsed = urlparse(url)

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

    url = clean_text(url)

    if not url:
        raise RuntimeError(
            "URL is empty."
        )

    if not is_prudential_url(url):
        raise RuntimeError(
            "Non-Prudential URL rejected: "
            f"{url}"
        )

    return url


# =============================================================================
# BASELINE FAILURE READING
# =============================================================================

def load_baseline_summary() -> dict:
    """
    Read the frozen baseline summary.

    The recovery script never imports or executes the frozen baseline Python
    script. It only consumes the JSON result it already produced.
    """

    if not BASELINE_RUN_SUMMARY.exists():
        raise FileNotFoundError(
            "Frozen baseline run_summary.json was not found: "
            f"{BASELINE_RUN_SUMMARY}"
        )

    try:
        data = json.loads(
            BASELINE_RUN_SUMMARY.read_text(
                encoding="utf-8"
            )
        )
    except Exception as error:
        raise RuntimeError(
            "Could not read baseline run_summary.json: "
            f"{error}"
        ) from error

    if not isinstance(data, dict):
        raise RuntimeError(
            "Baseline run_summary.json is not a JSON object."
        )

    return data


def extract_failed_rows(
    summary: dict,
) -> list[dict]:
    """
    Dynamically obtain failed funds from the baseline summary.

    Supports the current structure:

        "failed": [
            {
                "excelRow": 13,
                ...
            }
        ]

    Also tolerates a future failedFunds representation if necessary.
    """

    failed_items = summary.get(
        "failed"
    )

    if isinstance(
        failed_items,
        list,
    ):
        return [
            item
            for item in failed_items
            if isinstance(item, dict)
            and item.get("excelRow") is not None
        ]

    # Defensive fallback.
    failed_items = summary.get(
        "failedFunds"
    )

    if isinstance(
        failed_items,
        list,
    ):
        converted = []

        for item in failed_items:

            if isinstance(
                item,
                dict,
            ) and item.get("excelRow") is not None:
                converted.append(item)

            elif isinstance(
                item,
                int,
            ):
                converted.append(
                    {
                        "excelRow": item
                    }
                )

        return converted

    return []


# =============================================================================
# EXCEL MASTER UNIVERSE
# =============================================================================

def read_excel_funds() -> dict[int, dict]:
    """
    Read every populated URL from Column A.

    Column A remains the master universe.

    Column B is retained as the exact PruAccess reference name.
    """

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

    funds_by_row = {}

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

        funds_by_row[row_number] = {
            "excelRow": row_number,
            "prudentialUrl": prudential_url,
            "pruAccessName": pruaccess_name,
        }

    workbook.close()

    if not funds_by_row:
        raise RuntimeError(
            "No populated Prudential URLs were found "
            "in Excel Column A."
        )

    return funds_by_row


# =============================================================================
# PLAYWRIGHT / PAGE LOADING
# =============================================================================

def open_prudential_page(
    browser,
    prudential_url: str,
):
    """
    Open an official Prudential Singapore fund page.

    Retries are used for network instability only.
    """

    prudential_url = ensure_prudential_url(
        prudential_url
    )

    last_error = None

    for attempt in range(
        1,
        RETRY_COUNT + 1,
    ):

        page = browser.new_page()

        page.set_default_timeout(
            PAGE_TIMEOUT_MS
        )

        try:

            print(
                f"Opening Prudential page "
                f"(attempt {attempt}/{RETRY_COUNT})"
            )

            response = page.goto(
                prudential_url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT_MS,
            )

            if response is not None:

                status = response.status

                if status >= 400:
                    raise RuntimeError(
                        "Prudential page returned "
                        f"HTTP {status}."
                    )

            page.wait_for_timeout(
                POST_PAGE_WAIT_MS
            )

            final_url = clean_text(
                page.url
            )

            if not is_prudential_url(
                final_url
            ):
                raise RuntimeError(
                    "Prudential page redirected to "
                    "a non-Prudential host: "
                    f"{final_url}"
                )

            return page

        except Exception as error:

            last_error = error

            try:
                page.close()
            except Exception:
                pass

            if attempt < RETRY_COUNT:
                time.sleep(
                    RETRY_DELAY_SECONDS
                )

    raise RuntimeError(
        "Failed to open official Prudential page: "
        f"{last_error}"
    )


# =============================================================================
# FACTSHEET DISCOVERY
# =============================================================================

def find_factsheet_url(
    page,
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
                    "score": score,
                    "url": absolute_url,
                    "text": anchor_text,
                }
            )

    if not candidates:
        raise RuntimeError(
            "Could not find an official Prudential "
            "factsheet link."
        )

    candidates.sort(
        key=lambda item: (
            -item["score"],
            item["url"],
        )
    )

    selected = candidates[0]

    return ensure_prudential_url(
        selected["url"]
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

    last_error = None

    for attempt in range(
        1,
        RETRY_COUNT + 1,
    ):

        try:

            print(
                "Downloading official factsheet "
                f"(attempt {attempt}/{RETRY_COUNT})"
            )

            response = page.request.get(
                factsheet_url,
                timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
            )

            if response.status != 200:
                raise RuntimeError(
                    "Factsheet download returned "
                    f"HTTP {response.status}."
                )

            body = response.body()

            if not body.startswith(
                b"%PDF"
            ):
                raise RuntimeError(
                    "Factsheet response is not a valid PDF."
                )

            if len(body) < 1000:
                raise RuntimeError(
                    "Factsheet PDF is unexpectedly small."
                )

            return body

        except Exception as error:

            last_error = error

            if attempt < RETRY_COUNT:
                time.sleep(
                    RETRY_DELAY_SECONDS
                )

    raise RuntimeError(
        "Failed to download official factsheet: "
        f"{last_error}"
    )


# =============================================================================
# PDF BASIC VALIDATION
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

    return (
        full_text,
        page_count,
    )


# =============================================================================
# PDF COORDINATE EXTRACTION
# =============================================================================

def positioned_pdf_lines(
    pdf_bytes: bytes,
) -> list[dict]:
    """
    Extract visible PDF text together with physical x/y positions.

    This is the main recovery mechanism.

    We deliberately do not rely on pypdf's normal reading order because that
    can interleave two visual columns.
    """

    reader = PdfReader(
        BytesIO(pdf_bytes)
    )

    all_lines = []

    for page_number, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):

        fragments = []

        def visitor_text(
            text,
            cm,
            tm,
            font_dict,
            font_size,
        ):

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

            parts = raw.splitlines()

            if not parts:
                parts = [raw]

            for part_index, part in enumerate(
                parts
            ):

                part = clean_text(
                    part
                )

                if not part:
                    continue

                adjusted_y = y

                if part_index:
                    try:
                        adjusted_y = (
                            y
                            -
                            (
                                float(
                                    font_size
                                    or 8
                                )
                                *
                                part_index
                                *
                                1.15
                            )
                        )
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

            pdf_page.extract_text(
                visitor_text=visitor_text
            )

        except Exception:

            continue

        # Group text fragments that share a visual baseline.
        groups = []

        y_tolerance = 3.0

        for fragment in sorted(
            fragments,
            key=lambda item: (
                -item["y"],
                item["x0"],
            ),
        ):

            target = None

            for group in groups:

                if (
                    abs(
                        group["y"]
                        -
                        fragment["y"]
                    )
                    <= y_tolerance
                ):
                    target = group
                    break

            if target is None:

                target = {
                    "page": page_number,
                    "y": fragment["y"],
                    "fragments": [],
                }

                groups.append(
                    target
                )

            target["fragments"].append(
                fragment
            )

            target["y"] = (
                sum(
                    item["y"]
                    for item
                    in target["fragments"]
                )
                /
                len(
                    target["fragments"]
                )
            )

        for group in groups:

            fragments_sorted = sorted(
                group["fragments"],
                key=lambda item: item["x0"],
            )

            text_parts = []

            x0 = None
            x1 = None

            for fragment in fragments_sorted:

                text_parts.append(
                    fragment["text"]
                )

                x0 = (
                    fragment["x0"]
                    if x0 is None
                    else min(
                        x0,
                        fragment["x0"],
                    )
                )

                x1 = (
                    fragment["x1"]
                    if x1 is None
                    else max(
                        x1,
                        fragment["x1"],
                    )
                )

            line_text = clean_text(
                " ".join(
                    text_parts
                )
            )

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


# =============================================================================
# VISUAL HOLDINGS HELPERS
# =============================================================================

def is_top_holdings_heading(
    text: str,
) -> bool:

    normalized = normalize_text(
        text
    )

    return bool(
        re.search(
            r"\btop\s+(?:10|ten)\s+holdings\b",
            normalized,
            re.IGNORECASE,
        )
    )


def is_percentage_token(
    text: str,
) -> bool:

    text = clean_text(
        text
    )

    return bool(
        re.fullmatch(
            r"(?:\(?\s*)?\d+(?:\.\d+)?\s*%(?:\s*\)?)?",
            text,
        )
    )


def extract_percentages(
    text: str,
) -> list[str]:

    return re.findall(
        r"\(?\s*\d+(?:\.\d+)?\s*%\s*\)?",
        text,
    )


def clean_percentage(
    value: str,
) -> str:

    return clean_text(
        value
    ).replace(
        " ",
        "",
    )


def find_last_percentage(
    text: str,
) -> str | None:

    matches = extract_percentages(
        text
    )

    if not matches:
        return None

    return clean_percentage(
        matches[-1]
    )


def strip_last_percentage(
    text: str,
) -> tuple[str, str | None]:

    matches = list(
        re.finditer(
            r"\(?\s*\d+(?:\.\d+)?\s*%\s*\)?",
            text,
        )
    )

    if not matches:
        return (
            clean_text(text),
            None,
        )

    match = matches[-1]

    percentage = clean_percentage(
        match.group(0)
    )

    name = clean_text(
        (
            text[:match.start()]
            +
            " "
            +
            text[match.end():]
        )
    )

    return (
        name,
        percentage,
    )


def is_obvious_nonholding_row(
    text: str,
) -> bool:

    normalized = normalize_text(
        text
    )

    if not normalized:
        return True

    exact_rejections = {
        "date",
        "frequency",
        "dividend history",
        "distribution history",
        "performance",
        "calendar year performance",
        "year",
        "fund",
        "benchmark",
        "holding",
        "holdings",
        "weight",
        "weights",
        "source",
        "source:",
        "important information",
    }

    if normalized in exact_rejections:
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
        "top 10 holdings"
    ):
        return True

    return False


def determine_column_boundary(
    page_lines: list[dict],
    heading: dict,
    page_width: float,
) -> float:

    heading_x = heading["x0"]
    heading_y = heading["y"]

    # First preference: an explicit adjacent right-hand heading.
    right_headers = []

    for line in page_lines:

        if (
            abs(
                line["y"]
                -
                heading_y
            )
            > 45
        ):
            continue

        for fragment in line.get(
            "fragments",
            [],
        ):

            if (
                fragment["x0"]
                <=
                heading_x + 20
            ):
                continue

            normalized = normalize_text(
                fragment["text"]
            )

            if any(
                marker in normalized
                for marker in (
                    "dividend history",
                    "distribution history",
                    "performance history",
                )
            ):

                right_headers.append(
                    fragment["x0"]
                )

    if right_headers:

        return (
            min(
                right_headers
            )
            -
            8.0
        )

    # Second preference: detect a strong right-hand x position immediately
    # below the heading.
    nearby_x = []

    for line in page_lines:

        if (
            heading_y - 80
            <= line["y"]
            <= heading_y - 5
        ):

            for fragment in line.get(
                "fragments",
                [],
            ):

                if (
                    fragment["x0"]
                    >
                    heading_x + 25
                ):

                    nearby_x.append(
                        fragment["x0"]
                    )

    if nearby_x:

        candidate = min(
            nearby_x
        )

        if (
            candidate
            >
            heading_x
            +
            page_width * 0.18
        ):

            return (
                candidate
                -
                8.0
            )

    # Conservative fallback.
    return max(
        heading_x + 180.0,
        page_width * 0.68,
    )


def visual_row_text(
    line: dict,
) -> str:

    fragments = list(
        line.get(
            "fragments",
            [],
        )
    )

    fragments.sort(
        key=lambda item: item["x0"]
    )

    return clean_text(
        " ".join(
            fragment["text"]
            for fragment in fragments
        )
    )


# =============================================================================
# VISUAL SECTION RECONSTRUCTION
# =============================================================================

def build_visual_section(
    page_lines: list[dict],
    heading: dict,
    page_width: float,
) -> list[dict]:

    left_edge = max(
        0.0,
        heading["x0"] - 15.0,
    )

    right_edge = determine_column_boundary(
        page_lines,
        heading,
        page_width,
    )

    selected = []

    # Coordinates generally decrease down the page.
    for line in page_lines:

        if line["y"] >= heading["y"] - 2.0:
            continue

        if line["y"] < heading["y"] - 520.0:
            continue

        if (
            line["x1"]
            <
            left_edge
        ):
            continue

        if (
            line["x0"]
            >
            right_edge
        ):
            continue

        text = visual_row_text(
            line
        )

        if not text:
            continue

        if is_obvious_nonholding_row(
            text
        ):
            continue

        selected.append(
            line
        )

    selected.sort(
        key=lambda item: (
            -item["y"],
            item["x0"],
        )
    )

    return selected


# =============================================================================
# VISUAL HOLDING PARSER
# =============================================================================

def parse_visual_holdings(
    visual_lines: list[dict],
) -> list[dict]:
    """
    Parse the already-reconstructed visual holding column.

    HARD RULES:

    - A holding is accepted only when a published percentage exists.
    - The last percentage on a logical line is the portfolio weight.
    - Earlier percentages remain part of the security name.
    - No holding name is invented.
    - No percentage is invented.
    - No synthetic values are created.
    - Published order is preserved.
    """

    holdings = []

    pending_name_parts = []

    for line in visual_lines:

        text = visual_row_text(
            line
        )

        if not text:
            continue

        if is_obvious_nonholding_row(
            text
        ):
            continue

        percentage_matches = list(
            re.finditer(
                r"\(?\s*\d+(?:\.\d+)?\s*%\s*\)?",
                text,
            )
        )

        if percentage_matches:

            last_match = (
                percentage_matches[-1]
            )

            percentage = clean_percentage(
                last_match.group(0)
            )

            current_name = clean_text(
                (
                    text[:last_match.start()]
                    +
                    " "
                    +
                    text[last_match.end():]
                )
            )

            if pending_name_parts:
                name = clean_text(
                    " ".join(
                        pending_name_parts
                        +
                        [current_name]
                    )
                )
            else:
                name = current_name

            name = clean_text(
                name
            )

            if not name:
                raise RuntimeError(
                    "Published holding percentage found "
                    "but no holding name could be recovered "
                    "from the physical PDF row."
                )

            holdings.append(
                {
                    "rank": len(holdings) + 1,
                    "name": name,
                    "weight": percentage,
                }
            )

            pending_name_parts = []

            if len(holdings) >= MAX_HOLDINGS:
                break

        else:

            # A text-only row can be a wrapped holding name.
            # It is held until a subsequent row contains the published
            # portfolio percentage.
            pending_name_parts.append(
                text
            )

    if not holdings:
        raise RuntimeError(
            "No published holdings with percentages "
            "were recovered from the visual PDF layout."
        )

    if pending_name_parts:
        # Do not silently convert a dangling name into a holding.
        # It is preserved for diagnostics only.
        pass

    return holdings


# =============================================================================
# CANDIDATE DISCOVERY
# =============================================================================

def find_visual_candidates(
    positioned_lines: list[dict],
) -> list[dict]:

    candidates = []

    pages = sorted(
        {
            line["page"]
            for line in positioned_lines
        }
    )

    for page_number in pages:

        page_lines = [
            line
            for line in positioned_lines
            if line["page"] == page_number
        ]

        if not page_lines:
            continue

        max_x = max(
            line["x1"]
            for line in page_lines
        )

        min_x = min(
            line["x0"]
            for line in page_lines
        )

        page_width = max(
            595.0,
            max_x + 20.0,
            min_x + 595.0,
        )

        headings = [
            line
            for line in page_lines
            if is_top_holdings_heading(
                line["text"]
            )
        ]

        for heading in headings:

            visual_lines = build_visual_section(
                page_lines,
                heading,
                page_width,
            )

            if not visual_lines:
                continue

            try:

                holdings = parse_visual_holdings(
                    visual_lines
                )

            except Exception as error:

                candidates.append(
                    {
                        "page": page_number,
                        "heading": heading["text"],
                        "holdings": None,
                        "error": clean_text(
                            str(error)
                        ),
                        "visualLines": visual_lines,
                    }
                )

                continue

            candidates.append(
                {
                    "page": page_number,
                    "heading": heading["text"],
                    "holdings": holdings,
                    "error": None,
                    "visualLines": visual_lines,
                }
            )

    return candidates


# =============================================================================
# STRICT HOLDINGS VALIDATION
# =============================================================================

def validate_holdings(
    holdings: list[dict],
) -> None:

    if not holdings:
        raise RuntimeError(
            "Recovered holdings list is empty."
        )

    if len(holdings) > MAX_HOLDINGS:
        raise RuntimeError(
            "Recovered more than the permitted "
            f"{MAX_HOLDINGS} holdings."
        )

    for expected_rank, holding in enumerate(
        holdings,
        start=1,
    ):

        if holding.get(
            "rank"
        ) != expected_rank:
            raise RuntimeError(
                "Published holding order was not preserved."
            )

        name = clean_text(
            holding.get(
                "name"
            )
        )

        weight = clean_text(
            holding.get(
                "weight"
            )
        )

        if not name:
            raise RuntimeError(
                "Recovered holding has an empty name."
            )

        if not weight:
            raise RuntimeError(
                "Recovered holding has an empty weight."
            )

        if not re.fullmatch(
            r"\(?\d+(?:\.\d+)?%\)?",
            weight,
        ):
            raise RuntimeError(
                "Recovered holding weight is not a "
                "published percentage: "
                f"{weight}"
            )


def validate_against_pdf_text(
    holdings: list[dict],
    full_text: str,
) -> None:
    """
    Final source verification.

    Every recovered name and exact percentage must appear in the official
    PDF text.

    This is intentionally strict.

    We do NOT accept a recovered value merely because it was generated from
    spatial coordinates.
    """

    normalized_pdf = normalize_text(
        full_text
    )

    for holding in holdings:

        name = normalize_text(
            holding["name"]
        )

        weight = normalize_text(
            holding["weight"]
        )

        if not name:
            raise RuntimeError(
                "Cannot verify an empty holding name."
            )

        if name not in normalized_pdf:
            raise RuntimeError(
                "Recovered holding name does not appear "
                "in the official PDF text: "
                f"{holding['name']}"
            )

        # Percentage must also appear in the official source.
        if weight not in normalized_pdf:
            raise RuntimeError(
                "Recovered holding percentage does not "
                "appear in the official PDF text: "
                f"{holding['weight']}"
            )


# =============================================================================
# FUND NAME EXTRACTION
# =============================================================================

def extract_fund_name(
    page,
    fallback_name: str,
) -> str:

    try:

        title = clean_text(
            page.title()
        )

        if title:
            return title

    except Exception:
        pass

    try:

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
            return clean_text(
                match.group(0)
            )

    except Exception:
        pass

    return clean_text(
        fallback_name
    )


# =============================================================================
# SAVE RECOVERY RESULT
# =============================================================================

def save_fund_result(
    result: dict,
    factsheet_bytes: bytes | None = None,
    full_text: str | None = None,
    visual_section_text: str | None = None,
) -> None:

    excel_row = result.get(
        "excelRow"
    )

    fund_name = (
        result.get("fundName")
        or
        result.get("pruAccessName")
        or
        "fund"
    )

    directory = (
        RECOVERY_FUNDS_DIR
        /
        (
            f"{excel_row}_"
            f"{safe_filename(fund_name)}"
        )
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    if factsheet_bytes is not None:

        (
            directory
            /
            "factsheet.pdf"
        ).write_bytes(
            factsheet_bytes
        )

    if full_text is not None:

        (
            directory
            /
            "factsheet_text.txt"
        ).write_text(
            full_text,
            encoding="utf-8",
        )

    if visual_section_text is not None:

        (
            directory
            /
            "top_holdings_visual_section.txt"
        ).write_text(
            visual_section_text,
            encoding="utf-8",
        )

    save_json(
        directory
        /
        "top_holdings.json",
        result,
    )


# =============================================================================
# SINGLE FUND RECOVERY
# =============================================================================

def recover_fund(
    browser,
    fund: dict,
    baseline_failure: dict,
) -> dict:

    excel_row = fund["excelRow"]

    prudential_url = ensure_prudential_url(
        fund["prudentialUrl"]
    )

    pruaccess_name = fund.get(
        "pruAccessName",
        "",
    )

    print(
        "\n"
        "============================================================"
    )

    print(
        f"RECOVERING EXCEL ROW {excel_row}"
    )

    print(
        "============================================================"
    )

    print(
        f"Prudential URL: {prudential_url}"
    )

    print(
        f"PruAccess name: {pruaccess_name}"
    )

    print(
        "Baseline failure:"
    )

    print(
        clean_text(
            baseline_failure.get(
                "error",
                "Unknown baseline failure",
            )
        )
    )

    page = None

    try:

        page = open_prudential_page(
            browser,
            prudential_url,
        )

        final_url = clean_text(
            page.url
        )

        page_title = clean_text(
            page.title()
        )

        fund_name = extract_fund_name(
            page,
            pruaccess_name,
        )

        factsheet_url = find_factsheet_url(
            page
        )

        print(
            f"Factsheet URL: {factsheet_url}"
        )

        factsheet_bytes = download_factsheet(
            page,
            factsheet_url,
        )

        # ---------------------------------------------------------------------
        # First PDF read.
        # ---------------------------------------------------------------------

        full_text, page_count = extract_pdf_text(
            factsheet_bytes
        )

        if not re.search(
            r"\btop\s+(?:10|ten)\s+holdings\b",
            full_text,
            re.IGNORECASE,
        ):
            raise RuntimeError(
                "Official factsheet does not expose a "
                "Top 10 Holdings heading in extractable PDF text."
            )

        # ---------------------------------------------------------------------
        # Spatial recovery.
        # ---------------------------------------------------------------------

        positioned = positioned_pdf_lines(
            factsheet_bytes
        )

        candidates = find_visual_candidates(
            positioned
        )

        valid_candidates = [
            candidate
            for candidate in candidates
            if candidate.get(
                "holdings"
            )
        ]

        if not valid_candidates:

            diagnostics = [
                candidate.get(
                    "error"
                )
                for candidate in candidates
                if candidate.get(
                    "error"
                )
            ]

            detail = (
                "; ".join(
                    diagnostics[:5]
                )
                if diagnostics
                else
                "No usable visual Top Holdings candidate found."
            )

            raise RuntimeError(
                "Spatial recovery failed: "
                f"{detail}"
            )

        # Prefer the candidate with the most published holdings.
        # If counts tie, prefer the candidate with more substantial names.
        valid_candidates.sort(
            key=lambda candidate: (
                len(
                    candidate["holdings"]
                ),
                sum(
                    len(
                        item["name"]
                    )
                    for item
                    in candidate["holdings"]
                ),
            ),
            reverse=True,
        )

        selected = valid_candidates[0]

        holdings = selected[
            "holdings"
        ]

        validate_holdings(
            holdings
        )

        validate_against_pdf_text(
            holdings,
            full_text,
        )

        # ---------------------------------------------------------------------
        # FINAL VERIFICATION
        #
        # Re-download the official PDF and run the spatial extraction again.
        # This prevents a transient or partial first download from becoming
        # an accepted recovery.
        # ---------------------------------------------------------------------

        print(
            "Performing final official PDF verification..."
        )

        verification_bytes = download_factsheet(
            page,
            factsheet_url,
        )

        verification_text, verification_page_count = (
            extract_pdf_text(
                verification_bytes
            )
        )

        verification_positioned = (
            positioned_pdf_lines(
                verification_bytes
            )
        )

        verification_candidates = (
            find_visual_candidates(
                verification_positioned
            )
        )

        verification_valid = [
            candidate
            for candidate in verification_candidates
            if candidate.get(
                "holdings"
            )
        ]

        if not verification_valid:
            raise RuntimeError(
                "Final verification could not recover "
                "any visual Top Holdings candidate."
            )

        verification_valid.sort(
            key=lambda candidate: (
                len(
                    candidate["holdings"]
                ),
                sum(
                    len(
                        item["name"]
                    )
                    for item
                    in candidate["holdings"]
                ),
            ),
            reverse=True,
        )

        verified_holdings = (
            verification_valid[0][
                "holdings"
            ]
        )

        validate_holdings(
            verified_holdings
        )

        validate_against_pdf_text(
            verified_holdings,
            verification_text,
        )

        if holdings != verified_holdings:
            raise RuntimeError(
                "Final verification produced a different "
                "holding list from the initial recovery."
            )

        # ---------------------------------------------------------------------
        # Visual section diagnostic text.
        # ---------------------------------------------------------------------

        visual_lines = selected.get(
            "visualLines",
            [],
        )

        visual_section_text = "\n".join(
            visual_row_text(
                line
            )
            for line
            in visual_lines
        )

        result = {
            "status": "recovered",
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "finalUrl": final_url,
            "pageTitle": page_title,
            "fundName": fund_name,
            "pruAccessName": pruaccess_name,
            "factsheetUrl": factsheet_url,
            "factsheetPageCount": page_count,
            "verificationFactsheetPageCount": (
                verification_page_count
            ),
            "holdingsParser": "spatial_recovery",
            "recoveredFromPage": selected[
                "page"
            ],
            "verifiedFromPage": verification_valid[0][
                "page"
            ],
            "topHoldingsCount": len(
                holdings
            ),
            "topHoldings": holdings,
            "baselineFailure": baseline_failure,
            "verification": {
                "performed": True,
                "secondOfficialPdfDownload": True,
                "holdingsMatched": True,
                "namesPresentInOfficialPdf": True,
                "weightsPresentInOfficialPdf": True,
                "noInference": True,
                "noFabrication": True,
                "noSyntheticData": True,
                "publishedOrderPreserved": True,
                "fewerThanTenAllowed": True,
                "maximumTenHoldings": True,
            },
            "rules": {
                "officialPrudentialSourceOnly": True,
                "officialFactsheetOnly": True,
                "physicalPdfLayoutUsed": True,
                "noInferredHoldings": True,
                "noFabricatedPercentages": True,
                "noSyntheticData": True,
                "noForcedTenEntries": True,
                "publishedOrderPreserved": True,
                "duplicateNamesAllowed": True,
                "duplicateWeightsAllowed": True,
                "fixedIncomeLastPercentageUsed": True,
                "baselineUntouched": True,
                "pruaccessUntouched": True,
            },
            "recoveredAtUtc": utc_now_iso(),
        }

        save_fund_result(
            result,
            factsheet_bytes=verification_bytes,
            full_text=verification_text,
            visual_section_text=visual_section_text,
        )

        print(
            "\nRECOVERY SUCCESS"
        )

        print(
            f"Row: {excel_row}"
        )

        print(
            f"Holdings recovered: {len(holdings)}"
        )

        for holding in holdings:

            print(
                f"  {holding['rank']}. "
                f"{holding['name']} "
                f"— {holding['weight']}"
            )

        return result

    except Exception as error:

        error_text = clean_text(
            str(error)
        )

        result = {
            "status": "failed",
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "pruAccessName": pruaccess_name,
            "baselineFailure": baseline_failure,
            "error": error_text,
            "recoveredAtUtc": utc_now_iso(),
        }

        save_fund_result(
            result
        )

        print(
            "\nRECOVERY FAILED"
        )

        print(
            f"Row: {excel_row}"
        )

        print(
            f"Error: {error_text}"
        )

        return result

    finally:

        if page is not None:

            try:
                page.close()
            except Exception:
                pass


# =============================================================================
# MAIN
# =============================================================================

def main() -> int:

    print(
        "\n"
        "============================================================"
    )

    print(
        "VGrat FMS - PRUDENTIAL FAILED-FUND RECOVERY TEST"
    )

    print(
        "============================================================"
    )

    print(
        "IMPORTANT:"
    )

    print(
        "Frozen baseline will NOT be modified."
    )

    print(
        "PruAccess script will NOT be modified."
    )

    print(
        "Baseline output will NOT be overwritten."
    )

    print(
        "Recovery output:"
    )

    print(
        f"  {RECOVERY_OUTPUT_DIR}"
    )

    print(
        "============================================================\n"
    )

    # -------------------------------------------------------------------------
    # Read frozen baseline results.
    # -------------------------------------------------------------------------

    baseline_summary = (
        load_baseline_summary()
    )

    failed_items = extract_failed_rows(
        baseline_summary
    )

    if not failed_items:

        print(
            "No failed funds were reported by the "
            "frozen baseline."
        )

        summary = {
            "status": "nothing_to_recover",
            "generatedAtUtc": utc_now_iso(),
            "baselineRunSummary": str(
                BASELINE_RUN_SUMMARY
            ),
            "failedFundsFromBaseline": 0,
            "successfulFunds": 0,
            "failedFunds": 0,
            "recovered": [],
            "failed": [],
            "rules": {
                "baselineUntouched": True,
                "testOnly": True,
                "noGitCommit": True,
            },
        }

        save_json(
            RECOVERY_SUMMARY_FILE,
            summary,
        )

        save_json(
            RECOVERY_ALL_HOLDINGS_FILE,
            [],
        )

        return 0

    # -------------------------------------------------------------------------
    # Read Excel master universe.
    # -------------------------------------------------------------------------

    funds_by_row = read_excel_funds()

    # -------------------------------------------------------------------------
    # Match baseline failures to Excel rows.
    # -------------------------------------------------------------------------

    recovery_targets = []

    missing_excel_rows = []

    for failed_item in failed_items:

        try:
            excel_row = int(
                failed_item[
                    "excelRow"
                ]
            )
        except Exception:

            continue

        fund = funds_by_row.get(
            excel_row
        )

        if fund is None:

            missing_excel_rows.append(
                excel_row
            )

            continue

        recovery_targets.append(
            (
                fund,
                failed_item,
            )
        )

    print(
        "Baseline failed funds:"
        f" {len(failed_items)}"
    )

    print(
        "Recovery targets matched to Excel:"
        f" {len(recovery_targets)}"
    )

    if missing_excel_rows:

        print(
            "WARNING - failed rows not found in Excel:"
        )

        for row in missing_excel_rows:
            print(
                f"  Row {row}"
            )

    print(
        "\nRecovery rows:"
    )

    for fund, failed_item in recovery_targets:

        print(
            f"  Row {fund['excelRow']} "
            f"→ {fund['pruAccessName']}"
        )

    # -------------------------------------------------------------------------
    # Prepare recovery output.
    # -------------------------------------------------------------------------

    RECOVERY_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    recovered_results = []
    failed_results = []

    # -------------------------------------------------------------------------
    # Browser.
    # -------------------------------------------------------------------------

    with sync_playwright() as playwright:

        browser = playwright.chromium.launch(
            headless=BROWSER_HEADLESS
        )

        try:

            for fund, failed_item in recovery_targets:

                result = recover_fund(
                    browser,
                    fund,
                    failed_item,
                )

                if (
                    result.get(
                        "status"
                    )
                    ==
                    "recovered"
                ):

                    recovered_results.append(
                        result
                    )

                else:

                    failed_results.append(
                        result
                    )

        finally:

            browser.close()

    # -------------------------------------------------------------------------
    # Consolidated recovery-only output.
    # -------------------------------------------------------------------------

    recovery_status = (
        "success"
        if not failed_results
        else
        "partial_failure"
    )

    summary = {
        "status": recovery_status,
        "generatedAtUtc": utc_now_iso(),
        "baselineRunSummary": str(
            BASELINE_RUN_SUMMARY
        ),
        "excelFile": str(
            EXCEL_FILE
        ),
        "failedFundsFromBaseline": len(
            failed_items
        ),
        "matchedRecoveryTargets": len(
            recovery_targets
        ),
        "successfulFunds": len(
            recovered_results
        ),
        "failedFunds": len(
            failed_results
        ),
        "missingExcelRows": missing_excel_rows,
        "recovered": [
            {
                "excelRow": item.get(
                    "excelRow"
                ),
                "fundName": item.get(
                    "fundName"
                ),
                "topHoldingsCount": item.get(
                    "topHoldingsCount"
                ),
                "holdingsParser": item.get(
                    "holdingsParser"
                ),
            }
            for item in recovered_results
        ],
        "failed": [
            {
                "excelRow": item.get(
                    "excelRow"
                ),
                "error": item.get(
                    "error"
                ),
            }
            for item in failed_results
        ],
        "rules": {
            "baselineUntouched": True,
            "baselineRunSummaryReadOnly": True,
            "baselineAllHoldingsUntouched": True,
            "testPrudentialHoldingsUntouched": True,
            "testPruaccessUntouched": True,
            "recoveryOnlyFailedFunds": True,
            "officialPrudentialOnly": True,
            "officialFactsheetOnly": True,
            "physicalPdfLayoutRecovery": True,
            "secondPdfVerification": True,
            "noInferredHoldings": True,
            "noFabricatedData": True,
            "noSyntheticData": True,
            "noForcedTenHoldings": True,
            "publishedOrderPreserved": True,
            "noBaselineReplacement": True,
            "noGitCommit": True,
            "noGitPush": True,
        },
    }

    save_json(
        RECOVERY_SUMMARY_FILE,
        summary,
    )

    # Only recovered records are placed here.
    save_json(
        RECOVERY_ALL_HOLDINGS_FILE,
        recovered_results,
    )

    # -------------------------------------------------------------------------
    # Final console summary.
    # -------------------------------------------------------------------------

    print(
        "\n"
        "============================================================"
    )

    print(
        "PRUDENTIAL HOLDINGS RECOVERY TEST COMPLETE"
    )

    print(
        "============================================================"
    )

    print(
        f"Baseline failures: {len(failed_items)}"
    )

    print(
        f"Recovery successful: {len(recovered_results)}"
    )

    print(
        f"Recovery failed: {len(failed_results)}"
    )

    print(
        f"Recovery output: {RECOVERY_OUTPUT_DIR}"
    )

    print(
        "\nRecovered rows:"
    )

    for item in recovered_results:

        print(
            f"  Row {item.get('excelRow')}: "
            f"{item.get('fundName')} "
            f"({item.get('topHoldingsCount')} holdings)"
        )

    print(
        "\nStill failed:"
    )

    for item in failed_results:

        print(
            f"  Row {item.get('excelRow')}: "
            f"{item.get('error')}"
        )

    print(
        "\nNO BASELINE FILES WERE REPLACED."
    )

    print(
        "NO GIT COMMIT/PUSH WAS PERFORMED."
    )

    print(
        "============================================================"
    )

    return (
        0
        if not failed_results
        else 1
    )


if __name__ == "__main__":

    try:

        sys.exit(
            main()
        )

    except KeyboardInterrupt:

        print(
            "\nInterrupted by user."
        )

        sys.exit(130)

    except Exception as error:

        print(
            "\nFATAL RECOVERY ERROR:"
        )

        print(
            clean_text(
                str(error)
            )
        )

        sys.exit(1)
