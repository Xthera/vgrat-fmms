#!/usr/bin/env python3

"""
VGrat FMS - Prudential ALL-FUND HOLDINGS TEST EXTRACTOR

PURPOSE
=======

This is the first permanent Top Holdings testing stage.

It:

1. Reads every populated Prudential fund URL from:

       Funds Links.xlsm

2. Opens the official Prudential Singapore product page.

3. Finds the official Prudential Fund Factsheet PDF.

4. Downloads the PDF.

5. Extracts ALL PDF text.

6. Locates the official "Top 10 Holdings" section.

7. Saves the RAW Top Holdings section exactly as extracted
   from the PDF.

IMPORTANT:

This stage does NOT yet interpret portfolio percentages.

It does NOT yet attempt to determine whether a percentage
is:

    - a portfolio weight
    - a bond coupon
    - a maturity-related percentage
    - another security attribute

That interpretation will be added only after the actual
Prudential factsheet layouts have been tested.

This script is now the permanent testing script.

Whenever a holdings rule is proven correct, that rule should
remain permanently in this same script.

MASTER SOURCE
=============

Funds Links.xlsm

Column A:
    Prudential fund URL

Column B:
    PruAccess fund name

UNIVERSE RULES
==============

1. Excel Column A controls the universe.

2. Every populated URL in Column A is processed.

3. There is NO hardcoded fund count.

4. Duplicate URLs are NOT removed.

5. Each Excel row is treated as a separate fund entry.

6. Only official Prudential Singapore URLs are accepted.

7. Only official Prudential Singapore factsheets are accepted.

8. No third-party holdings sources.

9. No inferred holdings.

10. No fabricated holdings.

11. No fabricated percentages.

12. If the Top Holdings section cannot be located,
    the fund is reported as "no_holdings_section".

13. The raw PDF is retained.

14. The complete extracted PDF text is retained.

15. The raw extracted Top Holdings section is retained.

16. This script does not modify:

       test_pruaccess.py
       data.json
       index.html
       css/style.css
       js/app.js


CURRENT TEST STAGE
==================

Stage 1:
    Factsheet discovery/download/text extraction       COMPLETE

Stage 2:
    Top Holdings section detection                     CURRENT

Stage 3:
    Holdings name/weight interpretation                NEXT

Stage 4:
    Fixed-income percentage handling                   NEXT

Stage 5:
    Wrapped holdings handling                           NEXT

Stage 6:
    Full 67-fund regression                            NEXT

OUTPUT
======

output_factsheets/
    all_factsheets.json
    run_summary.json

    funds/
        <excelRow>_<identifier>/
            factsheet.pdf
            factsheet_text.txt
            top_holdings_section.txt
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
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"
FAILED_OUTPUT_DIR = OUTPUT_DIR / "failed"

ALL_FACTSHEETS_FILE = OUTPUT_DIR / "all_factsheets.json"
RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"

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
    return datetime.now(timezone.utc).isoformat()


def clean_text(value) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace("\xa0", " ")
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    return text.strip()


def normalize_text(value) -> str:
    text = clean_text(value)

    text = re.sub(r"\s+", " ", text)

    return text.strip().lower()


def save_json(path: Path, data) -> None:
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

    text = clean_text(value)

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

    text = text.strip(" ._")

    if not text:
        text = fallback

    return text[:180]


# ============================================================================
# URL VALIDATION
# ============================================================================

def is_prudential_url(url: str) -> bool:

    try:

        parsed = urlparse(url)

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


def ensure_prudential_url(url: str) -> str:

    url = clean_text(url)

    if not url:

        raise ValueError(
            "Prudential URL is blank."
        )

    if not is_prudential_url(url):

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

    IMPORTANT:

    - No hardcoded row limit.
    - Duplicate URLs are preserved.
    - Every populated Column A row becomes a separate fund entry.
    """

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    try:

        worksheet = workbook.active

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
    """
    Score an anchor to identify the most likely official
    Fund Factsheet.

    Higher score = stronger candidate.
    """

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

    if href_normalized.endswith(".pdf"):
        score += 20

    if ".pdf?" in href_normalized:
        score += 20

    return score


def find_factsheet_url(
    page,
    product_url: str,
) -> dict:
    """
    Find an official Prudential Fund Factsheet link.
    """

    candidates = []

    anchors = page.locator("a")

    count = anchors.count()

    for index in range(count):

        anchor = anchors.nth(index)

        try:

            href = anchor.get_attribute(
                "href"
            )

        except Exception:

            href = None

        if not href:
            continue

        href = clean_text(href)

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
    """
    Download the official factsheet.
    """

    response = page.request.get(
        factsheet_url,
        timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
    )

    status = response.status

    if status != 200:

        raise RuntimeError(
            f"Factsheet download returned HTTP {status}."
        )

    pdf_bytes = response.body()

    if not pdf_bytes:

        raise RuntimeError(
            "Factsheet download returned an empty response."
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
    """
    Extract ALL text from ALL pages.

    No holdings interpretation occurs here.
    """

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

    return full_text, page_count


# ============================================================================
# FUND NAME EXTRACTION
# ============================================================================

def extract_page_fund_name(page) -> str:
    """
    Try to obtain the fund name from the Prudential
    product page.
    """

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
                        headings.nth(index).inner_text()
                    )

                    if value:
                        return value

                except Exception:

                    continue

    except Exception:

        pass

    try:

        body_text = clean_text(
            page.locator("body").inner_text()
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
# TOP HOLDINGS DETECTION
# ============================================================================

def is_top_holdings_header(
    line: str,
) -> bool:
    """
    Determine whether a PDF text line represents the
    beginning of a Top Holdings section.

    This stage deliberately uses conservative detection.

    Accepted examples include:

        Top 10 Holdings
        Top Ten Holdings
        Top 10 Holdings:
        TOP 10 HOLDINGS

    We do NOT yet interpret holdings or percentages.
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

    # Allow simple PDF extraction variations such as
    # "Top 10 Holdings :" or similar whitespace.
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
    """
    Conservative detection of lines that commonly indicate
    the end of the Top Holdings section.

    This function does NOT interpret portfolio percentages.
    """

    normalized = normalize_text(
        line
    )

    if not normalized:
        return False

    # Common section boundaries.
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

    # Page markers commonly inserted by PDF extraction.
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
    """
    Find ALL Top Holdings sections in the extracted PDF text.

    We intentionally collect every occurrence rather than
    assuming there is only one.

    This is important during testing because PDF extraction
    can repeat headers across pages or contain multiple
    relevant sections.
    """

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

        start_line_number = index + 1

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

        # Remove trailing empty lines.
        while (
            section_lines
            and not section_lines[-1]
        ):
            section_lines.pop()

        sections.append(
            {
                "header": line,
                "startPdfTextLine": start_line_number,
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
    """
    Select the most useful Top Holdings section.

    Current rule:

    - If none exists -> no_holdings_section.
    - If one exists -> use it.
    - If multiple exist -> use the section containing
      the most non-empty lines.

    This selection rule is deliberately simple and will
    remain subject to regression testing.
    """

    sections = find_top_holdings_sections(
        pdf_text
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
    """
    Diagnostic analysis only.

    This does NOT decide which percentage is a portfolio
    weight.

    It simply reports what the raw section contains.

    This allows the next permanent parser rules to be built
    from actual Prudential examples.
    """

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

    section_lines = holdings_info[
        "selectedSection"
    ]["lines"]

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
            float(match.group(1))
            for match in matches
        ]

        if percentages:
            percentage_line_count += 1
            percentage_occurrences += len(
                percentages
            )

        analyzed_lines.append(
            {
                "lineNumber": line_number,
                "text": line,
                "percentageValues": percentages,
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
            if item["percentageCount"] > 1
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
        f"Product URL: {product_url}"
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

    page_fund_name = extract_page_fund_name(
        browser_page
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

    pdf_text, page_count = extract_pdf_text(
        pdf_bytes
    )

    print(
        f"PDF pages: {page_count}"
    )

    print(
        f"Extracted characters: "
        f"{len(pdf_text)}"
    )

    # ------------------------------------------------------------------------
    # TOP HOLDINGS DETECTION
    # ------------------------------------------------------------------------

    holdings_info = select_top_holdings_section(
        pdf_text
    )

    holdings_analysis = analyze_top_holdings_section(
        holdings_info
    )

    if holdings_info[
        "status"
    ] == "section_found":

        print()
        print(
            "TOP HOLDINGS: FOUND"
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

        if holdings_analysis[
            "multiplePercentageLines"
        ]:

            print(
                "Lines containing multiple percentages: "
                f"{len(holdings_analysis['multiplePercentageLines'])}"
            )

        print()
        print(
            "RAW TOP HOLDINGS SECTION"
        )
        print("-" * 80)

        raw_section = holdings_info[
            "rawSectionText"
        ]

        print(
            raw_section
        )

        print("-" * 80)

    else:

        print()
        print(
            "TOP HOLDINGS: NOT FOUND"
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

    metadata_path = (
        fund_output_dir
        / "metadata.json"
    )

    # ------------------------------------------------------------------------
    # SAVE ORIGINAL PDF
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
    # SAVE RAW TOP HOLDINGS SECTION
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

        # Remove an old section file if a previous run
        # happened to find one.
        if top_holdings_path.exists():

            top_holdings_path.unlink()

    # ------------------------------------------------------------------------
    # RESULT
    # ------------------------------------------------------------------------

    result = {
        "status": "success",
        "excelRow": excel_row,
        "prudentialUrl": product_url,
        "finalProductUrl": final_url,
        "excelPruAccessName": pruaccess_name,
        "fundName": page_fund_name,

        "factsheetUrl": factsheet_url,
        "factsheetAnchorText": selected[
            "anchorText"
        ],
        "factsheetCandidateScore": selected[
            "score"
        ],
        "factsheetCandidates": discovery[
            "candidates"
        ],

        "pdfPageCount": page_count,
        "pdfByteCount": len(
            pdf_bytes
        ),
        "extractedCharacterCount": len(
            pdf_text
        ),

        "topHoldingsStatus": holdings_info[
            "status"
        ],
        "topHoldingsSectionCount": holdings_info[
            "sectionCount"
        ],

        "topHoldingsSection": (
            holdings_info.get(
                "selectedSection"
            )
        ),

        "topHoldingsRawText": holdings_info.get(
            "rawSectionText",
            "",
        ),

        "topHoldingsAnalysis": holdings_analysis,

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
            str(top_holdings_path)
            if holdings_info[
                "status"
            ] == "section_found"
            else None
        ),

        "generatedAtUtc": utc_now_iso(),
    }

    save_json(
        metadata_path,
        result,
    )

    print()
    print(
        f"Saved PDF: {pdf_path}"
    )

    print(
        f"Saved text: {text_path}"
    )

    if holdings_info[
        "status"
    ] == "section_found":

        print(
            "Saved Top Holdings: "
            f"{top_holdings_path}"
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
        "error": str(error),
        "errorType": type(error).__name__,
        "generatedAtUtc": utc_now_iso(),
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
        f"Excel file: {EXCEL_FILE}"
    )

    print(
        f"Output directory: {OUTPUT_DIR}"
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

            # ---------------------------------------------------------------
            # Process every Excel row
            # ---------------------------------------------------------------

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
                            f"{attempt}/{MAX_RETRIES}"
                        )

                        result = process_single_fund(
                            page,
                            fund,
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

                        if attempt < MAX_RETRIES:

                            print(
                                f"Retrying in "
                                f"{RETRY_DELAY_SECONDS} "
                                f"seconds..."
                            )

                            time.sleep(
                                RETRY_DELAY_SECONDS
                            )

                if last_error is not None:

                    failure = save_failure_result(
                        fund,
                        last_error,
                    )

                    failed.append(
                        failure
                    )

                    print(
                        "STATUS: FAILED"
                    )

        finally:

            browser.close()

    # ------------------------------------------------------------------------
    # Consolidated output
    # ------------------------------------------------------------------------

    all_results = (
        successful
        + failed
    )

    all_results.sort(
        key=lambda item: item.get(
            "excelRow",
            0,
        )
    )

    all_factsheets = {
        "generatedAtUtc": utc_now_iso(),

        "excelFile": str(
            EXCEL_FILE
        ),

        "excelFundUniverse": len(
            funds
        ),

        "successfulFunds": len(
            successful
        ),

        "failedFunds": len(
            failed
        ),

        "topHoldingsFound": sum(
            1
            for result in successful
            if result.get(
                "topHoldingsStatus"
            ) == "section_found"
        ),

        "topHoldingsNotFound": sum(
            1
            for result in successful
            if result.get(
                "topHoldingsStatus"
            ) == "no_holdings_section"
        ),

        "results": all_results,
    }

    save_json(
        ALL_FACTSHEETS_FILE,
        all_factsheets,
    )

    # ------------------------------------------------------------------------
    # Run summary
    # ------------------------------------------------------------------------

    summary = {
        "status": (
            "success"
            if not failed
            else "partial"
        ),

        "generatedAtUtc": utc_now_iso(),

        "excelFile": str(
            EXCEL_FILE
        ),

        "excelFundUniverse": len(
            funds
        ),

        "successfulFunds": len(
            successful
        ),

        "failedFunds": len(
            failed
        ),

        "topHoldingsFound": sum(
            1
            for result in successful
            if result.get(
                "topHoldingsStatus"
            ) == "section_found"
        ),

        "topHoldingsNotFound": sum(
            1
            for result in successful
            if result.get(
                "topHoldingsStatus"
            ) == "no_holdings_section"
        ),

        "successfulRows": [
            result["excelRow"]
            for result in successful
        ],

        "failedRows": [
            result["excelRow"]
            for result in failed
        ],

        "purpose": (
            "Download official Prudential "
            "Fund Factsheets, extract raw PDF "
            "text, and locate the raw Top "
            "Holdings section."
        ),

        "rules": {

            "excelColumnAControlsUniverse": True,

            "processEveryPopulatedColumnAUrl": True,

            "hardcodedFundLimit": False,

            "duplicateUrlsPreserved": True,

            "officialPrudentialSingaporeOnly": True,

            "officialFactsheetsOnly": True,

            "thirdPartySources": False,

            "holdingsSectionDetection": True,

            "holdingsWeightInterpretation": False,

            "percentageInterpretationPerformed": False,

            "holdingsNamesInferred": False,

            "fabricatedData": False,

            "syntheticData": False,

            "pypdfTextExtraction": True,

            "rawTopHoldingsPreserved": True,

            "fullPdfPreserved": True,
        },
    }

    save_json(
        RUN_SUMMARY_FILE,
        summary,
    )

    # ------------------------------------------------------------------------
    # Console summary
    # ------------------------------------------------------------------------

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
        f"{sum(1 for result in successful if result.get('topHoldingsStatus') == 'section_found')}"
    )

    print(
        f"Top Holdings missing: "
        f"{sum(1 for result in successful if result.get('topHoldingsStatus') == 'no_holdings_section')}"
    )

    print()

    if successful:

        print(
            "FUND RESULTS:"
        )

        for result in successful:

            holdings_status = result.get(
                "topHoldingsStatus",
                "-",
            )

            section_count = result.get(
                "topHoldingsSectionCount",
                0,
            )

            print(
                f"  Row "
                f"{result['excelRow']}: "
                f"{result.get('fundName') or '-'}"
            )

            print(
                f"    Factsheet: "
                f"{result['factsheetUrl']}"
            )

            print(
                f"    Pages: "
                f"{result['pdfPageCount']}"
            )

            print(
                f"    Characters: "
                f"{result['extractedCharacterCount']}"
            )

            print(
                f"    Top Holdings: "
                f"{holdings_status}"
            )

            print(
                f"    Sections found: "
                f"{section_count}"
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
        f"All results: "
        f"{ALL_FACTSHEETS_FILE}"
    )

    print(
        f"Run summary: "
        f"{RUN_SUMMARY_FILE}"
    )

    print()
    print("=" * 80)

    # Keep baseline behavior:
    # workflow completes even if individual funds fail.
    return 0


if __name__ == "__main__":

    sys.exit(
        main()
    )
