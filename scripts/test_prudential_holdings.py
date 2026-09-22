#!/usr/bin/env python3

"""
VGrat FMS - Prudential ALL-FUND HOLDINGS TEST EXTRACTOR

CURRENT STAGE
=============

Stage 1:
    Official Prudential factsheet discovery/download       COMPLETE

Stage 2:
    Top Holdings section detection                         CURRENT

Stage 3:
    Holdings name/weight extraction                        NEXT

Stage 4:
    Fixed-income percentage handling                       NEXT

Stage 5:
    Wrapped holdings handling                              NEXT

Stage 6:
    Full 67-fund regression                                NEXT


PURPOSE
=======

This is the permanent Prudential holdings testing script.

It:

1. Reads every populated Prudential fund URL from:

       Funds Links.xlsm

2. Opens the official Prudential Singapore product page.

3. Finds the official Prudential Fund Factsheet PDF.

4. Downloads the PDF.

5. Extracts ALL PDF text.

6. Detects Top Holdings sections.

7. For funds where the current Top Holdings detector does not
   find a section, searches the PDF text for candidate holdings-
   related headings.

8. Saves the raw PDF, complete PDF text, detected Top Holdings
   section, and diagnostic candidate headings.

IMPORTANT
=========

This stage DOES NOT yet interpret portfolio percentages.

It does NOT decide whether a percentage is:

    - portfolio weight
    - bond coupon
    - security rate
    - maturity-related percentage
    - another security attribute

Those rules will be added only after the actual Prudential
factsheet layouts have been identified and tested.

PERMANENT DEVELOPMENT RULE
==========================

When a parsing rule is tested against the official Prudential
factsheets and proven correct, that rule becomes a permanent
part of this SAME script.

We do not maintain disposable parser versions.

UNIVERSE RULES
==============

1. Excel Column A controls the universe.

2. Every populated URL in Column A is processed.

3. There is NO hardcoded fund count.

4. Duplicate URLs are NOT removed.

5. Each Excel row is a separate fund entry.

6. Only official Prudential Singapore URLs are accepted.

7. Only official Prudential Singapore factsheets are accepted.

8. No third-party holdings sources.

9. No inferred holdings.

10. No fabricated holdings.

11. No fabricated percentages.

12. No forced 10 holdings.

13. If Prudential publishes fewer than 10 holdings, the final
    parser must preserve the actual published count.

14. Published holding order must be preserved.

15. Duplicate holding names are allowed if Prudential publishes
    them.

16. Duplicate percentages are allowed if Prudential publishes
    them.

17. This script does not modify:

       test_pruaccess.py
       data.json
       index.html
       css/style.css
       js/app.js


OUTPUT
======

output_factsheets/
    all_factsheets.json
    run_summary.json

    diagnostic_candidate_headings.json

    funds/
        <excelRow>_<identifier>/
            factsheet.pdf
            factsheet_text.txt
            top_holdings_section.txt
            candidate_holdings_headings.txt
            metadata.json

    failed/
        <excelRow>_failed/
            failure.json


REQUIREMENTS
============

    pip install openpyxl pypdf playwright

    playwright install chromium
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
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)


# ============================================================================
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_factsheets")

FUNDS_OUTPUT_DIR = (
    OUTPUT_DIR / "funds"
)

FAILED_OUTPUT_DIR = (
    OUTPUT_DIR / "failed"
)

ALL_FACTSHEETS_FILE = (
    OUTPUT_DIR / "all_factsheets.json"
)

RUN_SUMMARY_FILE = (
    OUTPUT_DIR / "run_summary.json"
)

CANDIDATE_HEADINGS_FILE = (
    OUTPUT_DIR / "diagnostic_candidate_headings.json"
)

HEADLESS = True

PAGE_TIMEOUT_MS = 120000

FACTSHEET_DOWNLOAD_TIMEOUT_MS = 120000

POST_PAGE_WAIT_MS = 1500

MAX_RETRIES = 3

RETRY_DELAY_SECONDS = 3

OFFICIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}


# ============================================================================
# GENERAL HELPERS
# ============================================================================

def utc_now_iso() -> str:
    return datetime.now(
        timezone.utc
    ).isoformat()


def clean_text(value) -> str:

    if value is None:
        return ""

    text = str(value)

    text = text.replace(
        "\xa0",
        " ",
    )

    text = text.replace(
        "\r\n",
        "\n",
    )

    text = text.replace(
        "\r",
        "\n",
    )

    return text.strip()


def normalize_text(value) -> str:

    text = clean_text(
        value
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip().lower()


def save_json(
    path: Path,
    data,
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
            data,
            handle,
            ensure_ascii=False,
            indent=2,
        )


def safe_filename(
    value: str,
    fallback: str = "fund",
) -> str:

    text = clean_text(
        value
    )

    if not text:
        text = fallback

    text = re.sub(
        r'[<>:"/\\|?*\x00-\x1F]',
        "_",
        text,
    )

    text = re.sub(
        r"\s+",
        "_",
        text,
    )

    text = text.strip(
        " ._"
    )

    if not text:
        text = fallback

    return text[:180]


# ============================================================================
# URL VALIDATION
# ============================================================================

def is_prudential_url(
    url: str,
) -> bool:

    try:

        parsed = urlparse(
            url
        )

        hostname = (
            parsed.hostname or ""
        ).lower()

        if parsed.scheme not in {
            "http",
            "https",
        }:
            return False

        return (
            hostname in OFFICIAL_HOSTS
            or hostname.endswith(
                ".prudential.com.sg"
            )
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

        raise ValueError(
            "Prudential URL is blank."
        )

    if not is_prudential_url(
        url
    ):

        raise ValueError(
            "URL is not an official "
            f"Prudential Singapore URL: {url}"
        )

    return url


# ============================================================================
# EXCEL
# ============================================================================

def read_excel_funds() -> list[dict]:
    """
    Read every populated URL from Column A.

    Column A:
        Prudential fund URL

    Column B:
        PruAccess fund name

    No hardcoded row limit.

    Duplicate URLs are preserved.
    """

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel file not found: "
            f"{EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    try:

        worksheet = (
            workbook.active
        )

        funds = []

        for row_number in range(
            2,
            worksheet.max_row + 1,
        ):

            prudential_url = clean_text(
                worksheet.cell(
                    row=row_number,
                    column=1,
                ).value
            )

            pruaccess_name = clean_text(
                worksheet.cell(
                    row=row_number,
                    column=2,
                ).value
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

        return funds

    finally:

        workbook.close()


# ============================================================================
# FACTSHEET LINK DISCOVERY
# ============================================================================

def factsheet_link_score(
    anchor_text: str,
    href: str,
) -> int:

    text = normalize_text(
        anchor_text
    )

    href_normalized = normalize_text(
        href
    )

    score = 0

    if "fund factsheet" in text:
        score += 100

    if "fund fact sheet" in text:
        score += 100

    if "factsheet" in text:
        score += 80

    if "fact sheet" in text:
        score += 80

    if "factsheet" in href_normalized:
        score += 60

    if "fact-sheet" in href_normalized:
        score += 60

    if "fund" in text:
        score += 10

    if href_normalized.endswith(
        ".pdf"
    ):
        score += 20

    if ".pdf?" in href_normalized:
        score += 20

    return score


def find_factsheet_url(
    page,
    product_url: str,
) -> dict:

    candidates = []

    anchors = page.locator(
        "a"
    )

    count = anchors.count()

    for index in range(
        count
    ):

        anchor = anchors.nth(
            index
        )

        try:

            href = anchor.get_attribute(
                "href"
            )

        except Exception:

            href = None

        if not href:
            continue

        href = clean_text(
            href
        )

        if not href:
            continue

        absolute_url = urljoin(
            product_url,
            href,
        )

        if not is_prudential_url(
            absolute_url
        ):
            continue

        try:

            anchor_text = clean_text(
                anchor.inner_text()
            )

        except Exception:

            anchor_text = ""

        score = factsheet_link_score(
            anchor_text,
            absolute_url,
        )

        if score <= 0:
            continue

        candidates.append(
            {
                "anchorText": anchor_text,
                "url": absolute_url,
                "score": score,
            }
        )

    if not candidates:

        raise RuntimeError(
            "No official Prudential Fund Factsheet "
            "link was found on the product page."
        )

    candidates.sort(
        key=lambda item: (
            item["score"],
            len(item["url"]),
        ),
        reverse=True,
    )

    selected = candidates[0]

    return {
        "selected": selected,
        "candidates": candidates,
    }


# ============================================================================
# PDF DOWNLOAD
# ============================================================================

def download_factsheet(
    page,
    factsheet_url: str,
) -> bytes:

    response = page.request.get(
        factsheet_url,
        timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
    )

    status = response.status

    if status != 200:

        raise RuntimeError(
            f"Factsheet download returned "
            f"HTTP {status}."
        )

    pdf_bytes = response.body()

    if not pdf_bytes:

        raise RuntimeError(
            "Factsheet download returned "
            "an empty response."
        )

    if not pdf_bytes.startswith(
        b"%PDF"
    ):

        preview = pdf_bytes[:100]

        raise RuntimeError(
            "Downloaded file does not appear "
            "to be a PDF. "
            f"First bytes: {preview!r}"
        )

    return pdf_bytes


# ============================================================================
# PDF TEXT EXTRACTION
# ============================================================================

def extract_pdf_text(
    pdf_bytes: bytes,
) -> tuple[str, int]:

    reader = PdfReader(
        BytesIO(pdf_bytes)
    )

    page_count = len(
        reader.pages
    )

    if page_count == 0:

        raise RuntimeError(
            "PDF contains zero pages."
        )

    extracted_pages = []

    for page_number, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):

        try:

            text = pdf_page.extract_text()

        except Exception as exc:

            raise RuntimeError(
                "PDF text extraction failed "
                f"on page {page_number}: {exc}"
            ) from exc

        if text is None:
            text = ""

        text = text.replace(
            "\r\n",
            "\n",
        )

        text = text.replace(
            "\r",
            "\n",
        )

        extracted_pages.append(
            text
        )

    full_text = "\n\n".join(
        extracted_pages
    )

    full_text = full_text.strip()

    if not full_text:

        raise RuntimeError(
            "PDF contains no extractable text."
        )

    return (
        full_text,
        page_count,
    )


# ============================================================================
# FUND NAME EXTRACTION
# ============================================================================

def extract_page_fund_name(
    page,
) -> str:

    try:

        headings = page.locator(
            "h1"
        )

        if headings.count() > 0:

            for index in range(
                headings.count()
            ):

                try:

                    value = clean_text(
                        headings.nth(
                            index
                        ).inner_text()
                    )

                    if value:
                        return value

                except Exception:

                    continue

    except Exception:

        pass

    try:

        body_text = clean_text(
            page.locator(
                "body"
            ).inner_text()
        )

        match = re.search(
            r"\bPRU(?:Link|Prime)\b[^\n]{3,200}",
            body_text,
            flags=re.IGNORECASE,
        )

        if match:

            return clean_text(
                match.group(0)
            )

    except Exception:

        pass

    return ""


# ============================================================================
# TOP HOLDINGS SECTION DETECTION
# ============================================================================

def is_top_holdings_header(
    line: str,
) -> bool:
    """
    CURRENT PERMANENT RULES

    These are the rules already proven by the first test:

        Top 10 Holdings
        Top Ten Holdings
        Top 10 Holding
        Top Ten Holding

    This function will be expanded only when a new heading
    pattern is confirmed from actual Prudential factsheets.
    """

    normalized = normalize_text(
        line
    )

    normalized = re.sub(
        r"[:\-–—]+$",
        "",
        normalized,
    ).strip()

    accepted_headers = {
        "top 10 holdings",
        "top ten holdings",
        "top 10 holding",
        "top ten holding",
    }

    if normalized in accepted_headers:
        return True

    if re.fullmatch(
        r"top\s+(10|ten)\s+holdings?",
        normalized,
        flags=re.IGNORECASE,
    ):
        return True

    return False


def is_top_holdings_end(
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
        "inception date",
        "inception date:",
        "important information",
        "important information:",
        "disclaimer",
        "disclaimer:",
        "past performance",
        "past performance:",
        "portfolio characteristics",
        "portfolio characteristics:",
        "asset allocation",
        "asset allocation:",
    }

    if normalized in exact_endings:
        return True

    if re.fullmatch(
        r"page\s+\d+",
        normalized,
        flags=re.IGNORECASE,
    ):
        return True

    if re.fullmatch(
        r"page\s+\d+\s+of\s+\d+",
        normalized,
        flags=re.IGNORECASE,
    ):
        return True

    return False


def find_top_holdings_sections(
    pdf_text: str,
) -> list[dict]:

    lines = pdf_text.splitlines()

    sections = []

    for index, raw_line in enumerate(
        lines
    ):

        line = clean_text(
            raw_line
        )

        if not is_top_holdings_header(
            line
        ):
            continue

        section_lines = []

        start_line_number = (
            index + 1
        )

        for following_index in range(
            index + 1,
            len(lines),
        ):

            following_line = clean_text(
                lines[following_index]
            )

            if is_top_holdings_end(
                following_line
            ):
                break

            section_lines.append(
                following_line
            )

        while (
            section_lines
            and not section_lines[-1]
        ):

            section_lines.pop()

        sections.append(
            {
                "header": line,
                "startPdfTextLine": (
                    start_line_number
                ),
                "lineCount": len(
                    section_lines
                ),
                "lines": section_lines,
            }
        )

    return sections


def select_top_holdings_section(
    pdf_text: str,
) -> dict:

    sections = (
        find_top_holdings_sections(
            pdf_text
        )
    )

    if not sections:

        return {
            "status": "no_holdings_section",
            "sectionCount": 0,
            "selectedSection": None,
            "sections": [],
        }

    selected = max(
        sections,
        key=lambda section: (
            section["lineCount"],
            section["startPdfTextLine"],
        ),
    )

    raw_section = "\n".join(
        selected["lines"]
    ).strip()

    return {
        "status": "section_found",
        "sectionCount": len(
            sections
        ),
        "selectedSection": selected,
        "rawSectionText": raw_section,
        "sections": sections,
    }


# ============================================================================
# CANDIDATE HOLDINGS HEADING DIAGNOSTICS
# ============================================================================

CANDIDATE_HEADING_PATTERNS = [

    (
        "top_holdings",
        re.compile(
            r"\btop\s+(?:10|ten)\s+holdings?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "top_holdings_without_number",
        re.compile(
            r"\btop\s+holdings?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "largest_holdings",
        re.compile(
            r"\blargest\s+holdings?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "largest_investments",
        re.compile(
            r"\blargest\s+investments?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "top_investments",
        re.compile(
            r"\btop\s+investments?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "portfolio_holdings",
        re.compile(
            r"\bportfolio\s+holdings?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "portfolio_investments",
        re.compile(
            r"\bportfolio\s+investments?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "holdings",
        re.compile(
            r"\bholdings?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "investments",
        re.compile(
            r"\binvestments?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "security_holdings",
        re.compile(
            r"\bsecurity\s+holdings?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "equity_holdings",
        re.compile(
            r"\bequity\s+holdings?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "bond_holdings",
        re.compile(
            r"\bbond\s+holdings?\b",
            re.IGNORECASE,
        ),
    ),

    (
        "fund_holdings",
        re.compile(
            r"\bfund\s+holdings?\b",
            re.IGNORECASE,
        ),
    ),
]


def is_probably_noise_heading(
    line: str,
) -> bool:
    """
    Reject obvious narrative sentences that merely contain
    the word "holdings" or "investments".

    This is intentionally conservative.

    We are diagnosing candidate headings, not yet accepting
    them as real Top Holdings sections.
    """

    normalized = normalize_text(
        line
    )

    if not normalized:
        return True

    if len(normalized) > 100:
        return True

    # Sentences are less likely to be headings.
    if normalized.endswith(
        "."
    ):
        return True

    # Common narrative language.
    narrative_phrases = (
        "the fund",
        "the portfolio",
        "may invest",
        "will invest",
        "invests in",
        "invested in",
        "holdings may",
        "holdings are",
        "holdings were",
        "holdings can",
        "investment objective",
        "investment strategy",
        "investment approach",
        "investment manager",
        "investment management",
    )

    for phrase in narrative_phrases:

        if phrase in normalized:

            return True

    return False


def heading_candidate_score(
    line: str,
    category: str,
) -> int:
    """
    Score a possible holdings heading.

    This is diagnostic only.

    A score does NOT automatically make the line a permanent
    holdings heading.
    """

    normalized = normalize_text(
        line
    )

    score = 0

    if category == "top_holdings":
        score += 100

    elif category == (
        "top_holdings_without_number"
    ):
        score += 90

    elif category == "largest_holdings":
        score += 85

    elif category == "largest_investments":
        score += 80

    elif category == "top_investments":
        score += 80

    elif category == "portfolio_holdings":
        score += 70

    elif category == "portfolio_investments":
        score += 65

    elif category == "security_holdings":
        score += 60

    elif category == "equity_holdings":
        score += 55

    elif category == "bond_holdings":
        score += 55

    elif category == "fund_holdings":
        score += 50

    elif category == "holdings":
        score += 40

    elif category == "investments":
        score += 30

    # Short heading-like text receives a small bonus.
    if len(normalized) <= 50:
        score += 10

    # A line containing only heading words is stronger than
    # a long mixed sentence.
    word_count = len(
        normalized.split()
    )

    if word_count <= 6:
        score += 10

    return score


def find_candidate_holdings_headings(
    pdf_text: str,
) -> list[dict]:
    """
    Find possible holdings-related headings in the entire
    PDF text.

    This is diagnostic only.

    No candidate found here is automatically accepted as a
    permanent Top Holdings section.
    """

    lines = pdf_text.splitlines()

    candidates = []

    seen = set()

    for index, raw_line in enumerate(
        lines
    ):

        line = clean_text(
            raw_line
        )

        if not line:
            continue

        if is_probably_noise_heading(
            line
        ):
            continue

        normalized = normalize_text(
            line
        )

        matched_categories = []

        for category, pattern in (
            CANDIDATE_HEADING_PATTERNS
        ):

            if pattern.search(
                normalized
            ):

                matched_categories.append(
                    category
                )

        if not matched_categories:
            continue

        # Prevent the same line from being repeated when
        # multiple broad patterns match it.
        key = (
            index,
            normalized,
        )

        if key in seen:
            continue

        seen.add(key)

        best_category = max(
            matched_categories,
            key=lambda category:
                heading_candidate_score(
                    line,
                    category,
                ),
        )

        previous_lines = []

        for previous_index in range(
            max(0, index - 3),
            index,
        ):

            previous = clean_text(
                lines[previous_index]
            )

            if previous:
                previous_lines.append(
                    previous
                )

        next_lines = []

        for next_index in range(
            index + 1,
            min(
                len(lines),
                index + 5,
            ),
        ):

            following = clean_text(
                lines[next_index]
            )

            if following:
                next_lines.append(
                    following
                )

        candidates.append(
            {
                "pdfTextLineNumber": (
                    index + 1
                ),
                "line": line,
                "matchedCategories": (
                    matched_categories
                ),
                "bestCategory": (
                    best_category
                ),
                "score": (
                    heading_candidate_score(
                        line,
                        best_category,
                    )
                ),
                "previousLines": (
                    previous_lines
                ),
                "nextLines": (
                    next_lines
                ),
            }
        )

    candidates.sort(
        key=lambda item: (
            item["score"],
            -item["pdfTextLineNumber"],
        ),
        reverse=True,
    )

    return candidates


def create_candidate_heading_text(
    candidates: list[dict],
) -> str:

    if not candidates:

        return (
            "No holdings-related candidate "
            "headings were detected.\n"
        )

    output = []

    output.append(
        "DIAGNOSTIC CANDIDATE HOLDINGS HEADINGS"
    )

    output.append(
        "=" * 80
    )

    output.append(
        "IMPORTANT:"
    )

    output.append(
        "These are diagnostic candidates only."
    )

    output.append(
        "They are NOT automatically accepted "
        "as Top Holdings sections."
    )

    output.append("")

    for position, candidate in enumerate(
        candidates,
        start=1,
    ):

        output.append(
            f"[Candidate {position}]"
        )

        output.append(
            f"PDF text line: "
            f"{candidate['pdfTextLineNumber']}"
        )

        output.append(
            f"Category: "
            f"{candidate['bestCategory']}"
        )

        output.append(
            f"Score: "
            f"{candidate['score']}"
        )

        output.append(
            f"Line: "
            f"{candidate['line']}"
        )

        if candidate[
            "matchedCategories"
        ]:

            output.append(
                "Matched categories: "
                + ", ".join(
                    candidate[
                        "matchedCategories"
                    ]
                )
            )

        output.append(
            "Previous lines:"
        )

        for previous in candidate[
            "previousLines"
        ]:

            output.append(
                f"  {previous}"
            )

        output.append(
            "Next lines:"
        )

        for following in candidate[
            "nextLines"
        ]:

            output.append(
                f"  {following}"
            )

        output.append(
            "-" * 80
        )

    return "\n".join(
        output
    ) + "\n"


# ============================================================================
# TOP HOLDINGS DIAGNOSTIC ANALYSIS
# ============================================================================

PERCENTAGE_PATTERN = re.compile(
    r"(?<![\d.])"
    r"([0-9]+(?:\.[0-9]+)?)"
    r"\s*%"
)


def analyze_top_holdings_section(
    holdings_info: dict,
) -> dict:

    if holdings_info.get(
        "status"
    ) != "section_found":

        return {
            "status": holdings_info.get(
                "status"
            ),
            "lineCount": 0,
            "percentageLineCount": 0,
            "percentageOccurrences": 0,
            "lines": [],
        }

    section_lines = (
        holdings_info[
            "selectedSection"
        ]["lines"]
    )

    analyzed_lines = []

    percentage_line_count = 0

    percentage_occurrences = 0

    for line_number, line in enumerate(
        section_lines,
        start=1,
    ):

        matches = list(
            PERCENTAGE_PATTERN.finditer(
                line
            )
        )

        percentages = [
            float(
                match.group(1)
            )
            for match in matches
        ]

        if percentages:

            percentage_line_count += 1

            percentage_occurrences += (
                len(percentages)
            )

        analyzed_lines.append(
            {
                "lineNumber": line_number,
                "text": line,
                "percentageValues": (
                    percentages
                ),
                "percentageCount": len(
                    percentages
                ),
            }
        )

    return {
        "status": "section_found",
        "lineCount": len(
            section_lines
        ),
        "percentageLineCount": (
            percentage_line_count
        ),
        "percentageOccurrences": (
            percentage_occurrences
        ),
        "multiplePercentageLines": [
            item
            for item in analyzed_lines
            if item[
                "percentageCount"
            ] > 1
        ],
        "lines": analyzed_lines,
    }


# ============================================================================
# SINGLE FUND
# ============================================================================

def process_single_fund(
    browser_page,
    fund: dict,
) -> dict:

    excel_row = fund[
        "excelRow"
    ]

    product_url = fund[
        "prudentialUrl"
    ]

    pruaccess_name = fund[
        "pruAccessName"
    ]

    product_url = ensure_prudential_url(
        product_url
    )

    print()
    print("=" * 80)

    print(
        f"EXCEL ROW {excel_row}"
    )

    print(
        f"Product URL: "
        f"{product_url}"
    )

    browser_page.goto(
        product_url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

    try:

        browser_page.wait_for_load_state(
            "networkidle",
            timeout=30000,
        )

    except PlaywrightTimeoutError:

        pass

    if POST_PAGE_WAIT_MS > 0:

        browser_page.wait_for_timeout(
            POST_PAGE_WAIT_MS
        )

    final_url = clean_text(
        browser_page.url
    )

    if not is_prudential_url(
        final_url
    ):

        raise RuntimeError(
            "Prudential product page redirected "
            f"to a non-Prudential URL: {final_url}"
        )

    page_fund_name = (
        extract_page_fund_name(
            browser_page
        )
    )

    print(
        f"Fund name: "
        f"{page_fund_name or '-'}"
    )

    discovery = find_factsheet_url(
        browser_page,
        final_url,
    )

    selected = discovery[
        "selected"
    ]

    factsheet_url = selected[
        "url"
    ]

    print(
        f"Factsheet: "
        f"{factsheet_url}"
    )

    print(
        f"Factsheet score: "
        f"{selected['score']}"
    )

    print(
        "Candidate factsheets found: "
        f"{len(discovery['candidates'])}"
    )

    pdf_bytes = download_factsheet(
        browser_page,
        factsheet_url,
    )

    pdf_text, page_count = (
        extract_pdf_text(
            pdf_bytes
        )
    )

    print(
        f"PDF pages: "
        f"{page_count}"
    )

    print(
        f"Extracted characters: "
        f"{len(pdf_text)}"
    )

    # ------------------------------------------------------------------------
    # CURRENT TOP HOLDINGS DETECTOR
    # ------------------------------------------------------------------------

    holdings_info = (
        select_top_holdings_section(
            pdf_text
        )
    )

    holdings_analysis = (
        analyze_top_holdings_section(
            holdings_info
        )
    )

    # ------------------------------------------------------------------------
    # NEW DIAGNOSTIC CANDIDATE SEARCH
    # ------------------------------------------------------------------------

    candidate_headings = (
        find_candidate_holdings_headings(
            pdf_text
        )
    )

    # ------------------------------------------------------------------------
    # Console diagnostics
    # ------------------------------------------------------------------------

    if holdings_info[
        "status"
    ] == "section_found":

        print()
        print(
            "TOP HOLDINGS: FOUND "
            "USING CONFIRMED RULE"
        )

        print(
            "Top Holdings sections found: "
            f"{holdings_info['sectionCount']}"
        )

        print(
            "Selected section lines: "
            f"{holdings_info['selectedSection']['lineCount']}"
        )

        print(
            "Percentage occurrences: "
            f"{holdings_analysis['percentageOccurrences']}"
        )

        print()
        print(
            "RAW TOP HOLDINGS SECTION"
        )

        print(
            "-" * 80
        )

        print(
            holdings_info[
                "rawSectionText"
            ]
        )

        print(
            "-" * 80
        )

    else:

        print()
        print(
            "TOP HOLDINGS: NOT FOUND "
            "USING CURRENT CONFIRMED RULE"
        )

        print(
            "Diagnostic candidate headings: "
            f"{len(candidate_headings)}"
        )

        if candidate_headings:

            print()

            print(
                "TOP CANDIDATE HEADINGS:"
            )

            for candidate in (
                candidate_headings[:10]
            ):

                print(
                    f"  "
                    f"Line "
                    f"{candidate['pdfTextLineNumber']}: "
                    f"{candidate['line']}"
                )

                print(
                    f"    Category: "
                    f"{candidate['bestCategory']}"
                )

                print(
                    f"    Score: "
                    f"{candidate['score']}"
                )

        else:

            print(
                "  No candidate headings found."
            )

    # ------------------------------------------------------------------------
    # OUTPUT DIRECTORY
    # ------------------------------------------------------------------------

    identifier_source = (
        page_fund_name
        or pruaccess_name
        or f"row_{excel_row}"
    )

    identifier = safe_filename(
        identifier_source,
        fallback=f"row_{excel_row}",
    )

    fund_output_dir = (
        FUNDS_OUTPUT_DIR
        / f"{excel_row}_{identifier}"
    )

    fund_output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    pdf_path = (
        fund_output_dir
        / "factsheet.pdf"
    )

    text_path = (
        fund_output_dir
        / "factsheet_text.txt"
    )

    top_holdings_path = (
        fund_output_dir
        / "top_holdings_section.txt"
    )

    candidate_headings_path = (
        fund_output_dir
        / "candidate_holdings_headings.txt"
    )

    metadata_path = (
        fund_output_dir
        / "metadata.json"
    )

    # ------------------------------------------------------------------------
    # SAVE PDF
    # ------------------------------------------------------------------------

    with pdf_path.open(
        "wb"
    ) as handle:

        handle.write(
            pdf_bytes
        )

    # ------------------------------------------------------------------------
    # SAVE COMPLETE PDF TEXT
    # ------------------------------------------------------------------------

    with text_path.open(
        "w",
        encoding="utf-8",
    ) as handle:

        handle.write(
            pdf_text
        )

    # ------------------------------------------------------------------------
    # SAVE CONFIRMED TOP HOLDINGS SECTION
    # ------------------------------------------------------------------------

    if holdings_info[
        "status"
    ] == "section_found":

        with top_holdings_path.open(
            "w",
            encoding="utf-8",
        ) as handle:

            handle.write(
                holdings_info[
                    "rawSectionText"
                ]
            )

    else:

        if top_holdings_path.exists():

            top_holdings_path.unlink()

    # ------------------------------------------------------------------------
    # SAVE CANDIDATE HEADING DIAGNOSTICS
    # ------------------------------------------------------------------------

    candidate_text = (
        create_candidate_heading_text(
            candidate_headings
        )
    )

    with candidate_headings_path.open(
        "w",
        encoding="utf-8",
    ) as handle:

        handle.write(
            candidate_text
        )

    # ------------------------------------------------------------------------
    # RESULT
    # ------------------------------------------------------------------------

    result = {
        "status": "success",

        "excelRow": excel_row,

        "prudentialUrl": product_url,

        "finalProductUrl": final_url,

        "excelPruAccessName": (
            pruaccess_name
        ),

        "fundName": page_fund_name,

        "factsheetUrl": factsheet_url,

        "factsheetAnchorText": (
            selected[
                "anchorText"
            ]
        ),

        "factsheetCandidateScore": (
            selected[
                "score"
            ]
        ),

        "factsheetCandidates": (
            discovery[
                "candidates"
            ]
        ),

        "pdfPageCount": page_count,

        "pdfByteCount": len(
            pdf_bytes
        ),

        "extractedCharacterCount": len(
            pdf_text
        ),

        # --------------------------------------------------------------------
        # Confirmed Top Holdings detector
        # --------------------------------------------------------------------

        "topHoldingsStatus": (
            holdings_info[
                "status"
            ]
        ),

        "topHoldingsSectionCount": (
            holdings_info[
                "sectionCount"
            ]
        ),

        "topHoldingsSection": (
            holdings_info.get(
                "selectedSection"
            )
        ),

        "topHoldingsRawText": (
            holdings_info.get(
                "rawSectionText",
                "",
            )
        ),

        "topHoldingsAnalysis": (
            holdings_analysis
        ),

        # --------------------------------------------------------------------
        # Candidate diagnostic detector
        # --------------------------------------------------------------------

        "candidateHoldingsHeadings": (
            candidate_headings
        ),

        "candidateHoldingsHeadingCount": (
            len(candidate_headings)
        ),

        "outputDirectory": str(
            fund_output_dir
        ),

        "pdfFile": str(
            pdf_path
        ),

        "textFile": str(
            text_path
        ),

        "topHoldingsFile": (
            str(
                top_holdings_path
            )
            if holdings_info[
                "status"
            ] == "section_found"
            else None
        ),

        "candidateHoldingsHeadingsFile": str(
            candidate_headings_path
        ),

        "generatedAtUtc": (
            utc_now_iso()
        ),
    }

    save_json(
        metadata_path,
        result,
    )

    print()

    print(
        f"Saved PDF: "
        f"{pdf_path}"
    )

    print(
        f"Saved text: "
        f"{text_path}"
    )

    if holdings_info[
        "status"
    ] == "section_found":

        print(
            "Saved confirmed Top Holdings: "
            f"{top_holdings_path}"
        )

    print(
        "Saved candidate-heading diagnostics: "
        f"{candidate_headings_path}"
    )

    return result


# ============================================================================
# FAILURE HANDLING
# ============================================================================

def save_failure_result(
    fund: dict,
    error: Exception,
) -> dict:

    excel_row = fund[
        "excelRow"
    ]

    failed_dir = (
        FAILED_OUTPUT_DIR
        / f"{excel_row}_failed"
    )

    failed_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    result = {
        "status": "failed",

        "excelRow": excel_row,

        "prudentialUrl": fund[
            "prudentialUrl"
        ],

        "excelPruAccessName": fund[
            "pruAccessName"
        ],

        "error": str(
            error
        ),

        "errorType": type(
            error
        ).__name__,

        "generatedAtUtc": (
            utc_now_iso()
        ),
    }

    save_json(
        failed_dir / "failure.json",
        result,
    )

    return result


# ============================================================================
# MAIN
# ============================================================================

def main() -> int:

    print("=" * 80)

    print(
        "VGrat FMS - PRUDENTIAL "
        "ALL-FUND HOLDINGS TEST EXTRACTOR"
    )

    print("=" * 80)

    print()

    print(
        f"Excel file: "
        f"{EXCEL_FILE}"
    )

    print(
        f"Output directory: "
        f"{OUTPUT_DIR}"
    )

    print()

    # ------------------------------------------------------------------------
    # Prepare output directories
    # ------------------------------------------------------------------------

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FAILED_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ------------------------------------------------------------------------
    # Read Excel
    # ------------------------------------------------------------------------

    try:

        funds = read_excel_funds()

    except Exception as exc:

        print()

        print(
            "ERROR: Unable to read Excel."
        )

        print(
            str(exc)
        )

        return 1

    print(
        f"Excel fund universe: "
        f"{len(funds)}"
    )

    print(
        "Duplicate URLs are preserved."
    )

    print(
        "Every populated Column A row "
        "will be processed."
    )

    print()

    if not funds:

        print(
            "ERROR: No populated URLs "
            "found in Column A."
        )

        return 1

    successful = []

    failed = []

    # ------------------------------------------------------------------------
    # Start browser
    # ------------------------------------------------------------------------

    with sync_playwright() as playwright:

        browser = playwright.chromium.launch(
            headless=HEADLESS
        )

        page = browser.new_page()

        page.set_default_timeout(
            PAGE_TIMEOUT_MS
        )

        try:

            for position, fund in enumerate(
                funds,
                start=1,
            ):

                print()

                print(
                    f"[{position}/{len(funds)}]"
                )

                last_error = None

                for attempt in range(
                    1,
                    MAX_RETRIES + 1,
                ):

                    try:

                        print(
                            f"Attempt "
                            f"{attempt}/"
                            f"{MAX_RETRIES}"
                        )

                        result = (
                            process_single_fund(
                                page,
                                fund,
                            )
                        )

                        successful.append(
                            result
                        )

                        print(
                            "STATUS: SUCCESS"
                        )

                        last_error = None

                        break

                    except Exception as exc:

                        last_error = exc

                        print(
                            "ERROR:"
                        )

                        print(
                            f"  "
                            f"{type(exc).__name__}: "
                            f"{exc}"
                        )

                        if (
                            attempt
                            < MAX_RETRIES
                        ):

                            print(
                                f"Retrying in "
                                f"{RETRY_DELAY_SECONDS} "
                                f"seconds..."
                            )

                            time.sleep(
                                RETRY_DELAY_SECONDS
                            )

                if last_error is not None:

                    failure = (
                        save_failure_result(
                            fund,
                            last_error,
                        )
                    )

                    failed.append(
                        failure
                    )

                    print(
                        "STATUS: FAILED"
                    )

        finally:

            browser.close()

    # =========================================================================
    # CONSOLIDATED OUTPUT
    # =========================================================================

    all_results = (
        successful
        + failed
    )

    all_results.sort(
        key=lambda item:
            item.get(
                "excelRow",
                0,
            )
    )

    top_holdings_found = sum(
        1
        for result in successful
        if result.get(
            "topHoldingsStatus"
        ) == "section_found"
    )

    top_holdings_missing = sum(
        1
        for result in successful
        if result.get(
            "topHoldingsStatus"
        ) == "no_holdings_section"
    )

    candidate_results = []

    for result in successful:

        if result.get(
            "topHoldingsStatus"
        ) != "section_found":

            candidate_results.append(
                {
                    "excelRow": (
                        result[
                            "excelRow"
                        ]
                    ),
                    "fundName": (
                        result.get(
                            "fundName",
                            "",
                        )
                    ),
                    "prudentialUrl": (
                        result[
                            "prudentialUrl"
                        ]
                    ),
                    "factsheetUrl": (
                        result[
                            "factsheetUrl"
                        ]
                    ),
                    "candidateCount": (
                        result.get(
                            "candidateHoldingsHeadingCount",
                            0,
                        )
                    ),
                    "candidates": (
                        result.get(
                            "candidateHoldingsHeadings",
                            [],
                        )
                    ),
                }
            )

    all_factsheets = {

        "generatedAtUtc": (
            utc_now_iso()
        ),

        "excelFile": str(
            EXCEL_FILE
        ),

        "excelFundUniverse": (
            len(funds)
        ),

        "successfulFunds": (
            len(successful)
        ),

        "failedFunds": (
            len(failed)
        ),

        "topHoldingsFound": (
            top_holdings_found
        ),

        "topHoldingsMissing": (
            top_holdings_missing
        ),

        "candidateHeadingDiagnosticsForMissingFunds": (
            len(candidate_results)
        ),

        "results": all_results,
    }

    save_json(
        ALL_FACTSHEETS_FILE,
        all_factsheets,
    )

    # ------------------------------------------------------------------------
    # Candidate diagnostic master file
    # ------------------------------------------------------------------------

    save_json(
        CANDIDATE_HEADINGS_FILE,
        {
            "generatedAtUtc": (
                utc_now_iso()
            ),
            "excelFundUniverse": (
                len(funds)
            ),
            "fundsWithConfirmedTopHoldings": (
                top_holdings_found
            ),
            "fundsWithoutConfirmedTopHoldings": (
                top_holdings_missing
            ),
            "candidateResults": (
                candidate_results
            ),
        },
    )

    # =========================================================================
    # RUN SUMMARY
    # =========================================================================

    summary = {

        "status": (
            "success"
            if not failed
            else "partial"
        ),

        "generatedAtUtc": (
            utc_now_iso()
        ),

        "excelFile": str(
            EXCEL_FILE
        ),

        "excelFundUniverse": (
            len(funds)
        ),

        "successfulFunds": (
            len(successful)
        ),

        "failedFunds": (
            len(failed)
        ),

        "topHoldingsFound": (
            top_holdings_found
        ),

        "topHoldingsMissing": (
            top_holdings_missing
        ),

        "candidateHeadingDiagnostics": (
            len(candidate_results)
        ),

        "successfulRows": [
            result[
                "excelRow"
            ]
            for result in successful
        ],

        "failedRows": [
            result[
                "excelRow"
            ]
            for result in failed
        ],

        "purpose": (
            "Download official Prudential "
            "Fund Factsheets, extract raw PDF "
            "text, detect confirmed Top Holdings "
            "sections, and diagnose alternative "
            "holdings-related headings."
        ),

        "rules": {

            "excelColumnAControlsUniverse": True,

            "processEveryPopulatedColumnAUrl": True,

            "hardcodedFundLimit": False,

            "duplicateUrlsPreserved": True,

            "officialPrudentialSingaporeOnly": True,

            "officialFactsheetsOnly": True,

            "thirdPartySources": False,

            "confirmedTopHoldingsDetection": True,

            "candidateHeadingDiagnostics": True,

            "holdingsWeightInterpretation": False,

            "percentageInterpretationPerformed": False,

            "holdingsNamesInferred": False,

            "fabricatedData": False,

            "syntheticData": False,

            "pypdfTextExtraction": True,

            "rawTopHoldingsPreserved": True,

            "fullPdfPreserved": True,

            "candidateHeadingsAutomaticallyAccepted": False,
        },
    }

    save_json(
        RUN_SUMMARY_FILE,
        summary,
    )

    # =========================================================================
    # FINAL CONSOLE SUMMARY
    # =========================================================================

    print()
    print()
    print("=" * 80)

    print(
        "FINAL SUMMARY"
    )

    print("=" * 80)

    print(
        f"Excel fund universe : "
        f"{len(funds)}"
    )

    print(
        f"Successful funds    : "
        f"{len(successful)}"
    )

    print(
        f"Failed funds        : "
        f"{len(failed)}"
    )

    print(
        f"Top Holdings found  : "
        f"{top_holdings_found}"
    )

    print(
        f"Top Holdings missing: "
        f"{top_holdings_missing}"
    )

    print()

    print(
        "CANDIDATE HEADING DIAGNOSTICS"
    )

    print(
        f"Funds requiring diagnosis: "
        f"{len(candidate_results)}"
    )

    print()

    # ------------------------------------------------------------------------
    # Aggregate candidate categories
    # ------------------------------------------------------------------------

    category_counts = {}

    for result in candidate_results:

        for candidate in result.get(
            "candidates",
            [],
        ):

            category = candidate.get(
                "bestCategory",
                "unknown",
            )

            category_counts[
                category
            ] = (
                category_counts.get(
                    category,
                    0,
                )
                + 1
            )

    if category_counts:

        print(
            "Candidate categories:"
        )

        for category, count in sorted(
            category_counts.items(),
            key=lambda item: (
                -item[1],
                item[0],
            ),
        ):

            print(
                f"  {category}: "
                f"{count}"
            )

    print()

    if failed:

        print(
            "FAILED ROWS:"
        )

        for result in failed:

            print(
                f"  Row "
                f"{result['excelRow']}: "
                f"{result['error']}"
            )

        print()

    print(
        "IMPORTANT:"
    )

    print(
        "Candidate headings have NOT been "
        "automatically promoted to permanent "
        "Top Holdings rules."
    )

    print(
        "They must be reviewed against the "
        "actual Prudential factsheet layouts "
        "before becoming permanent rules."
    )

    print()

    print(
        f"All results: "
        f"{ALL_FACTSHEETS_FILE}"
    )

    print(
        f"Run summary: "
        f"{RUN_SUMMARY_FILE}"
    )

    print(
        f"Candidate diagnostics: "
        f"{CANDIDATE_HEADINGS_FILE}"
    )

    print()

    print("=" * 80)

    # Preserve baseline behavior:
    # individual failures do not terminate the workflow.
    return 0


if __name__ == "__main__":

    sys.exit(
        main()
    )
