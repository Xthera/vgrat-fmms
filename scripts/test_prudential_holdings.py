#!/usr/bin/env python3

"""
VGrat FMS - Prudential Top Holdings Extraction Test
====================================================

MASTER FUND UNIVERSE
--------------------
Funds Links.xlsm

Column A:
    Prudential fund URL

Column B:
    PruAccess fund name

IMPORTANT
---------
- Excel Column A controls the fund universe dynamically.
- Every populated Prudential URL is processed.
- No hardcoded fund count.
- No additional funds are discovered from Prudential.
- This script does NOT modify test_pruaccess.py.
- This script extracts the official Prudential factsheet and
  the published Top 10 Holdings section.

HOLDINGS RULES
--------------
- Source must be the official Prudential Singapore fund page/factsheet.
- Holdings come only from the official published factsheet.
- Maximum 10 holdings.
- Fewer than 10 published holdings is valid.
- Do not invent missing holdings.
- Do not calculate holdings.
- Do not estimate holdings.
- Do not fill blanks.
- Do not force exactly 10 holdings.
- Multi-line holding names are joined into one name.
- The holding percentage is used to determine the end of the
  holding-name block.
- PDF text extraction artifacts such as "None", "|" and repeated
  whitespace are cleaned.
- Raw section text is preserved for audit.

OUTPUT
------
output_holdings/
    all_holdings.json
    run_summary.json

    funds/
        <excelRow>_<fundIdentifier>/
            factsheet.pdf
            factsheet_text.txt
            top_holdings_section.txt
            top_holdings.json
            metadata.json

    failures/
        <excelRow>_<fundIdentifier>/
            error.json
"""

from __future__ import annotations

import json
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

from bs4 import BeautifulSoup
from openpyxl import load_workbook
from playwright.sync_api import (
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)
from pypdf import PdfReader


# ============================================================================
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_holdings")
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"
FAILURES_OUTPUT_DIR = OUTPUT_DIR / "failures"

ALL_HOLDINGS_FILE = OUTPUT_DIR / "all_holdings.json"
RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"

MAX_RETRIES = 3

RETRY_DELAY_SECONDS = 2

PAGE_TIMEOUT_MS = 90000

POST_LOAD_WAIT_MS = 2500

FACTSHEET_DOWNLOAD_TIMEOUT_SECONDS = 90

MAX_HOLDINGS = 10

PRUDENTIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/153.0.0.0 Safari/537.36"
)


# ============================================================================
# GENERAL HELPERS
# ============================================================================

def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(
        microsecond=0
    ).isoformat().replace("+00:00", "Z")


def clean_text(value: Any) -> str:
    """
    Normalize whitespace and common PDF extraction artifacts.
    """
    if value is None:
        return ""

    text = str(value)

    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")

    # Pipe characters often appear because PDF table extraction
    # interpreted table separators as text.
    text = text.replace("|", " ")

    # Common PDF extraction artifact.
    text = re.sub(
        r"\bNone\b",
        " ",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def normalize_name(value: Any) -> str:
    """
    Normalize a holding name while preserving meaningful punctuation.
    """
    text = clean_text(value)

    text = re.sub(
        r"\s*-\s*",
        " - ",
        text,
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip(" -|")


def normalize_match_text(value: Any) -> str:
    return re.sub(
        r"\s+",
        " ",
        clean_text(value),
    ).casefold().strip()


def safe_filename(value: str) -> str:
    value = clean_text(value)

    value = re.sub(
        r"[<>:\"/\\|?*]",
        "_",
        value,
    )

    value = re.sub(
        r"\s+",
        "_",
        value,
    )

    value = value.strip("._ ")

    if not value:
        value = "unknown"

    return value[:180]


def save_json(
    path: Path,
    data: Any,
) -> None:

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with path.open(
        "w",
        encoding="utf-8",
    ) as file:

        json.dump(
            data,
            file,
            indent=2,
            ensure_ascii=False,
        )

        file.write("\n")


def load_json(path: Path) -> Any:

    with path.open(
        "r",
        encoding="utf-8",
    ) as file:

        return json.load(file)


def is_prudential_url(url: str) -> bool:

    try:
        parsed = urlparse(url)

        hostname = (
            parsed.hostname or ""
        ).lower()

        return hostname in PRUDENTIAL_HOSTS

    except Exception:
        return False


def parse_percent(value: str) -> Optional[float]:

    if not value:
        return None

    match = re.search(
        r"(?<![\d.])"
        r"(\d+(?:\.\d+)?|\.\d+)"
        r"\s*%",
        value,
    )

    if not match:
        return None

    try:

        number = float(
            match.group(1)
        )

        if number < 0:
            return None

        if number > 100:
            return None

        return number

    except ValueError:
        return None


def extract_percent_tokens(
    line: str,
) -> List[Tuple[float, int, int]]:

    results: List[
        Tuple[float, int, int]
    ] = []

    pattern = re.compile(
        r"(?<![\d.])"
        r"(\d+(?:\.\d+)?|\.\d+)"
        r"\s*%"
    )

    for match in pattern.finditer(line):

        try:

            number = float(
                match.group(1)
            )

            if (
                number >= 0
                and number <= 100
            ):
                results.append(
                    (
                        number,
                        match.start(),
                        match.end(),
                    )
                )

        except ValueError:
            continue

    return results


# ============================================================================
# EXCEL
# ============================================================================

def read_excel_funds() -> List[Dict[str, Any]]:
    """
    Read every populated URL from Column A.

    Column B is kept as the corresponding PruAccess fund name.
    """

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        filename=EXCEL_FILE,
        read_only=True,
        data_only=False,
    )

    worksheet = workbook.active

    funds: List[
        Dict[str, Any]
    ] = []

    seen_urls = set()

    for row_number in range(
        2,
        worksheet.max_row + 1,
    ):

        url_cell = worksheet.cell(
            row=row_number,
            column=1,
        )

        name_cell = worksheet.cell(
            row=row_number,
            column=2,
        )

        url = ""

        if url_cell.value:
            url = clean_text(
                url_cell.value
            )

        # Support actual Excel hyperlinks.
        if (
            not url
            and url_cell.hyperlink
            and url_cell.hyperlink.target
        ):
            url = clean_text(
                url_cell.hyperlink.target
            )

        if not url:
            continue

        if not is_prudential_url(
            url
        ):
            raise ValueError(
                f"Excel row {row_number} "
                f"contains a non-Prudential URL: "
                f"{url}"
            )

        normalized_url = url.rstrip("/")

        if normalized_url in seen_urls:
            continue

        seen_urls.add(
            normalized_url
        )

        pruaccess_name = ""

        if name_cell.value:
            pruaccess_name = clean_text(
                name_cell.value
            )

        funds.append(
            {
                "excelRow": row_number,
                "prudentialUrl": url,
                "excelPruAccessName": pruaccess_name,
            }
        )

    workbook.close()

    return funds


# ============================================================================
# WEB PAGE
# ============================================================================

def get_page_html(
    page,
) -> str:

    return page.content()


def extract_page_fund_name(
    page,
) -> str:

    selectors = [
        "h1",
        "[class*='fund-name']",
        "[class*='fundName']",
        "[class*='product-title']",
        "[class*='productTitle']",
    ]

    for selector in selectors:

        try:

            elements = page.locator(
                selector
            )

            count = elements.count()

            for index in range(count):

                value = clean_text(
                    elements.nth(
                        index
                    ).inner_text(
                        timeout=3000
                    )
                )

                if value:

                    # Avoid obvious generic headings.
                    lowered = value.casefold()

                    if lowered not in {
                        "fund factsheet",
                        "factsheet",
                        "prulink funds",
                    }:

                        return value

        except Exception:
            continue

    try:

        title = clean_text(
            page.title()
        )

        if title:
            return title

    except Exception:
        pass

    return ""


def find_factsheet_url(
    page,
) -> Optional[str]:
    """
    Find an official Prudential factsheet PDF.

    Checks:
        1. visible Factsheet text
        2. PDF href
        3. href containing 'factsheet'
    """

    html = get_page_html(
        page
    )

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    candidates: List[
        Tuple[int, str, str]
    ] = []

    for anchor in soup.find_all(
        "a"
    ):

        href = clean_text(
            anchor.get(
                "href",
                "",
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

        text = clean_text(
            anchor.get_text(
                " ",
                strip=True,
            )
        )

        lowered_href = (
            absolute_url.casefold()
        )

        lowered_text = (
            text.casefold()
        )

        score = 0

        if "factsheet" in lowered_text:
            score += 100

        if "fact sheet" in lowered_text:
            score += 100

        if "factsheet" in lowered_href:
            score += 80

        if ".pdf" in lowered_href:
            score += 40

        if (
            "view" in lowered_text
            and "fact" in lowered_text
        ):
            score += 20

        if score > 0:

            candidates.append(
                (
                    score,
                    absolute_url,
                    text,
                )
            )

    # Additional direct PDF links.
    for anchor in soup.select(
        "a[href$='.pdf'], a[href*='.pdf?']"
    ):

        href = clean_text(
            anchor.get(
                "href",
                "",
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

        if not any(
            candidate[1]
            == absolute_url
            for candidate in candidates
        ):

            candidates.append(
                (
                    10,
                    absolute_url,
                    clean_text(
                        anchor.get_text(
                            " ",
                            strip=True,
                        )
                    ),
                )
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


# ============================================================================
# PDF
# ============================================================================

def download_pdf(
    url: str,
) -> bytes:

    request = Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": (
                "application/pdf,"
                "application/octet-stream,"
                "*/*"
            ),
        },
    )

    with urlopen(
        request,
        timeout=FACTSHEET_DOWNLOAD_TIMEOUT_SECONDS,
    ) as response:

        content_type = (
            response.headers.get(
                "Content-Type",
                "",
            )
        )

        data = response.read()

        if not data:
            raise ValueError(
                "Downloaded factsheet is empty."
            )

        # Do not require the server to provide a perfect PDF content type.
        # Check the actual PDF signature.
        if not data.startswith(
            b"%PDF"
        ):

            raise ValueError(
                "Downloaded factsheet is not a valid PDF."
                f" Content-Type={content_type}"
            )

        return data


def extract_pdf_text(
    pdf_bytes: bytes,
) -> Tuple[str, int]:

    temp_path = (
        OUTPUT_DIR
        / "_temp_factsheet.pdf"
    )

    temp_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temp_path.write_bytes(
        pdf_bytes
    )

    try:

        reader = PdfReader(
            str(temp_path)
        )

        pages: List[str] = []

        for page in reader.pages:

            try:

                text = page.extract_text()

                if text:
                    pages.append(
                        text
                    )
                else:
                    pages.append("")

            except Exception as exc:

                pages.append(
                    f"[PDF_TEXT_EXTRACTION_ERROR: {exc}]"
                )

        full_text = "\n".join(
            pages
        )

        return (
            full_text,
            len(reader.pages),
        )

    finally:

        try:
            temp_path.unlink()
        except FileNotFoundError:
            pass


# ============================================================================
# FACTSHEET METADATA
# ============================================================================

def extract_first_matching_date(
    text: str,
    patterns: List[str],
) -> Optional[str]:

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match:

            value = clean_text(
                match.group(1)
            )

            if value:
                return value

    return None


def extract_factsheet_metadata(
    pdf_text: str,
) -> Dict[str, Any]:

    factsheet_document_date = (
        extract_first_matching_date(
            pdf_text,
            [
                r"(?:factsheet|document)"
                r"\s*(?:date|dated)"
                r"\s*[:\-]?\s*"
                r"([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})",

                r"(?:factsheet|document)"
                r"\s*(?:date|dated)"
                r"\s*[:\-]?\s*"
                r"([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{4})",
            ],
        )
    )

    factsheet_data_as_at = (
        extract_first_matching_date(
            pdf_text,
            [
                r"(?:data|price|information)"
                r"\s+(?:as|as of)"
                r"\s+at?\s*"
                r"[:\-]?\s*"
                r"([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})",

                r"(?:as at|as of)"
                r"\s*"
                r"([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})",

                r"(?:as at|as of)"
                r"\s*"
                r"([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{4})",
            ],
        )
    )

    return {
        "factsheetDocumentDate": factsheet_document_date,
        "factsheetDataAsAt": factsheet_data_as_at,
    }


# ============================================================================
# TOP HOLDINGS SECTION
# ============================================================================

SECTION_STOP_PATTERNS = [
    r"^\s*asset allocation\s*$",
    r"^\s*geographical allocation\s*$",
    r"^\s*geographic allocation\s*$",
    r"^\s*country allocation\s*$",
    r"^\s*regional allocation\s*$",
    r"^\s*sector allocation\s*$",
    r"^\s*currency allocation\s*$",
    r"^\s*portfolio allocation\s*$",
    r"^\s*fund allocation\s*$",
    r"^\s*risk allocation\s*$",
    r"^\s*performance\s*$",
    r"^\s*fund performance\s*$",
    r"^\s*investment strategy\s*$",
    r"^\s*investment objective\s*$",
    r"^\s*important information\s*$",
    r"^\s*important notes\s*$",
    r"^\s*disclaimer\s*$",
    r"^\s*contact us\s*$",
]


def is_stop_heading(
    line: str,
) -> bool:

    normalized = normalize_match_text(
        line
    )

    if not normalized:
        return False

    for pattern in SECTION_STOP_PATTERNS:

        if re.match(
            pattern,
            normalized,
            flags=re.IGNORECASE,
        ):
            return True

    return False


def find_top_holdings_section(
    pdf_text: str,
) -> Optional[str]:

    lines = pdf_text.splitlines()

    start_index: Optional[int] = None

    for index, line in enumerate(
        lines
    ):

        normalized = normalize_match_text(
            line
        )

        if (
            normalized == "top 10 holdings"
            or normalized == "top holdings"
            or "top 10 holdings" in normalized
        ):

            start_index = index
            break

    if start_index is None:
        return None

    section_lines: List[str] = []

    for index in range(
        start_index,
        len(lines),
    ):

        raw_line = lines[index]

        line = clean_text(
            raw_line
        )

        if (
            index > start_index
            and is_stop_heading(line)
        ):
            break

        section_lines.append(
            raw_line
        )

        # Enough protection against accidental extraction
        # of huge parts of the PDF.
        if len(section_lines) > 250:
            break

    section = "\n".join(
        section_lines
    ).strip()

    if not section:
        return None

    return section


# ============================================================================
# TOP HOLDINGS PARSER
# ============================================================================

HEADER_WORDS = {
    "holding",
    "holdings",
    "name",
    "names",
    "weight",
    "weights",
    "%",
    "allocation",
    "allocations",
}


def is_probable_header_line(
    line: str,
) -> bool:

    normalized = normalize_match_text(
        line
    )

    if not normalized:
        return True

    if normalized.startswith(
        "top 10 holdings"
    ):
        return True

    tokens = normalized.split()

    if (
        tokens
        and all(
            token in HEADER_WORDS
            for token in tokens
        )
    ):
        return True

    if normalized in {
        "name weight",
        "holding weight",
        "holdings weight",
        "investment weight",
    }:
        return True

    return False


def strip_leading_rank(
    line: str,
) -> Tuple[Optional[int], str]:

    text = clean_text(
        line
    )

    match = re.match(
        r"^\s*(\d{1,2})[\s.)\-:]+(.+)$",
        text,
    )

    if match:

        try:

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

        except ValueError:
            pass

    match = re.match(
        r"^\s*(\d{1,2})\s*$",
        text,
    )

    if match:

        try:

            rank = int(
                match.group(1)
            )

            if 1 <= rank <= 99:

                return (
                    rank,
                    "",
                )

        except ValueError:
            pass

    return (
        None,
        text,
    )


def looks_like_noise(
    line: str,
) -> bool:

    normalized = normalize_match_text(
        line
    )

    if not normalized:
        return True

    if is_probable_header_line(
        line
    ):
        return True

    noise_patterns = [
        r"^top\s+10\s+holdings",
        r"^holdings$",
        r"^name\s+weight$",
        r"^weight$",
        r"^investment$",
        r"^none$",
        r"^[\-–—]+$",
    ]

    for pattern in noise_patterns:

        if re.match(
            pattern,
            normalized,
        ):
            return True

    return False


def parse_top_holdings(
    section_text: str,
) -> List[Dict[str, Any]]:
    """
    Parse Top 10 Holdings while supporting names that wrap across
    multiple PDF text lines.

    Example PDF text:

        1
        JPMORGAN FUNDS - EMERGING
        MARKETS EQUITY FUND
        14.7%

    becomes:

        {
            "rank": 1,
            "name": "JPMORGAN FUNDS - EMERGING MARKETS EQUITY FUND",
            "weightPercent": 14.7,
            "weightText": "14.7%"
        }

    The percentage acts as the row boundary.
    """

    raw_lines = (
        section_text.splitlines()
    )

    holdings: List[
        Dict[str, Any]
    ] = []

    pending_name_lines: List[str] = []

    pending_rank: Optional[int] = None

    def flush_holding(
        rank: Optional[int],
        extra_name: str,
        weight: float,
        weight_text: str,
    ) -> None:

        nonlocal pending_name_lines

        name_parts = []

        for line in pending_name_lines:

            cleaned = clean_text(
                line
            )

            if cleaned:
                name_parts.append(
                    cleaned
                )

        cleaned_extra = clean_text(
            extra_name
        )

        if cleaned_extra:
            name_parts.append(
                cleaned_extra
            )

        # Remove accidental duplicate rank text.
        normalized_parts: List[str] = []

        for part in name_parts:

            part_rank, part_text = (
                strip_leading_rank(
                    part
                )
            )

            if (
                part_rank is not None
                and not part_text
            ):
                continue

            if part_text:
                part = part_text

            normalized_parts.append(
                clean_text(part)
            )

        # Remove common PDF table artifacts.
        filtered_parts = []

        for part in normalized_parts:

            if normalize_match_text(
                part
            ) in {
                "none",
                "-",
                "—",
                "–",
            }:
                continue

            filtered_parts.append(
                part
            )

        name = normalize_name(
            " ".join(
                filtered_parts
            )
        )

        if not name:
            # Do not invent a name.
            # The caller will not add an empty holding.
            pending_name_lines = []
            return

        actual_rank = (
            rank
            if rank is not None
            else len(holdings) + 1
        )

        holdings.append(
            {
                "rank": actual_rank,
                "name": name,
                "weightPercent": weight,
                "weightText": weight_text,
            }
        )

        pending_name_lines = []

    for raw_line in raw_lines:

        line = clean_text(
            raw_line
        )

        if not line:
            continue

        if looks_like_noise(
            line
        ):
            continue

        detected_rank, rank_text = (
            strip_leading_rank(
                line
            )
        )

        if (
            detected_rank is not None
            and not rank_text
        ):

            pending_rank = (
                detected_rank
            )

            continue

        if (
            detected_rank is not None
            and rank_text
        ):

            if (
                not pending_name_lines
                and pending_rank is None
            ):
                pending_rank = (
                    detected_rank
                )
                line = rank_text

            else:
                line = rank_text

        percent_tokens = (
            extract_percent_tokens(
                line
            )
        )

        if percent_tokens:

            # Use the final percentage on the line.
            # This is safer with PDF table extraction where
            # another percentage may appear earlier in the line.
            weight, start_pos, end_pos = (
                percent_tokens[-1]
            )

            name_before_weight = (
                line[:start_pos]
            )

            weight_text = clean_text(
                line[
                    start_pos:end_pos
                ]
            )

            flush_holding(
                rank=pending_rank,
                extra_name=name_before_weight,
                weight=weight,
                weight_text=weight_text,
            )

            pending_rank = None

            if len(holdings) >= MAX_HOLDINGS:
                break

            continue

        # No percentage yet.
        # This is either a wrapped part of the holding name,
        # or a rank/name line.
        cleaned_line = clean_text(
            line
        )

        if cleaned_line:

            pending_name_lines.append(
                cleaned_line
            )

    return holdings[:MAX_HOLDINGS]


# ============================================================================
# FUND EXTRACTION
# ============================================================================

def make_fund_directory(
    excel_row: int,
    fund_identifier: Optional[str],
) -> Path:

    identifier = (
        safe_filename(
            fund_identifier
            or "unknown"
        )
    )

    return (
        FUNDS_OUTPUT_DIR
        / f"{excel_row}_{identifier}"
    )


def make_failure_directory(
    excel_row: int,
    fund_identifier: Optional[str],
) -> Path:

    identifier = (
        safe_filename(
            fund_identifier
            or "unknown"
        )
    )

    return (
        FAILURES_OUTPUT_DIR
        / f"{excel_row}_{identifier}"
    )


def extract_identifier_from_url(
    url: str,
) -> Optional[str]:

    parsed = urlparse(
        url
    )

    query = parsed.query

    match = re.search(
        r"(?:^|&)citicode=([^&]+)",
        query,
        flags=re.IGNORECASE,
    )

    if match:
        return clean_text(
            match.group(1)
        )

    return None


def process_single_fund(
    page,
    fund: Dict[str, Any],
) -> Dict[str, Any]:

    excel_row = fund[
        "excelRow"
    ]

    prudential_url = fund[
        "prudentialUrl"
    ]

    excel_pruaccess_name = fund[
        "excelPruAccessName"
    ]

    print(
        "\n============================================================"
    )

    print(
        f"Excel row: {excel_row}"
    )

    print(
        f"URL: {prudential_url}"
    )

    if excel_pruaccess_name:

        print(
            f"PruAccess name: "
            f"{excel_pruaccess_name}"
        )

    print(
        "============================================================"
    )

    last_error: Optional[str] = None

    for attempt in range(
        1,
        MAX_RETRIES + 1,
    ):

        try:

            print(
                f"[Row {excel_row}] "
                f"Opening Prudential page "
                f"(attempt {attempt}/{MAX_RETRIES})..."
            )

            page.goto(
                prudential_url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT_MS,
            )

            page.wait_for_timeout(
                POST_LOAD_WAIT_MS
            )

            final_url = clean_text(
                page.url
            )

            if not is_prudential_url(
                final_url
            ):
                raise ValueError(
                    "Final URL is not an "
                    f"official Prudential Singapore URL: "
                    f"{final_url}"
                )

            page_title = clean_text(
                page.title()
            )

            fund_name = (
                extract_page_fund_name(
                    page
                )
            )

            print(
                f"[Row {excel_row}] "
                f"Prudential fund name: "
                f"{fund_name or '-'}"
            )

            factsheet_url = (
                find_factsheet_url(
                    page
                )
            )

            if not factsheet_url:

                raise ValueError(
                    "Could not locate an "
                    "official Prudential factsheet PDF "
                    "on the fund page."
                )

            if not is_prudential_url(
                factsheet_url
            ):

                raise ValueError(
                    "Factsheet URL is not an official "
                    f"Prudential Singapore URL: "
                    f"{factsheet_url}"
                )

            print(
                f"[Row {excel_row}] "
                f"Factsheet: "
                f"{factsheet_url}"
            )

            pdf_bytes = download_pdf(
                factsheet_url
            )

            (
                pdf_text,
                page_count,
            ) = extract_pdf_text(
                pdf_bytes
            )

            if not pdf_text.strip():

                raise ValueError(
                    "Factsheet PDF produced no "
                    "extractable text."
                )

            holdings_section = (
                find_top_holdings_section(
                    pdf_text
                )
            )

            fund_identifier = (
                extract_identifier_from_url(
                    prudential_url
                )
            )

            fund_dir = (
                make_fund_directory(
                    excel_row,
                    fund_identifier,
                )
            )

            fund_dir.mkdir(
                parents=True,
                exist_ok=True,
            )

            # Save the official PDF for audit.
            (
                fund_dir
                / "factsheet.pdf"
            ).write_bytes(
                pdf_bytes
            )

            (
                fund_dir
                / "factsheet_text.txt"
            ).write_text(
                pdf_text,
                encoding="utf-8",
            )

            if holdings_section is None:

                metadata = {
                    "status": "no_holdings_section",
                    "excelRow": excel_row,
                    "prudentialUrl": prudential_url,
                    "finalUrl": final_url,
                    "pageTitle": page_title,
                    "fundName": fund_name,
                    "fundIdentifier": fund_identifier,
                    "excelPruAccessName": (
                        excel_pruaccess_name
                    ),
                    "factsheetUrl": factsheet_url,
                    "factsheetPageCount": page_count,
                    "holdingsSectionStatus": (
                        "not_found"
                    ),
                    "topHoldingsCount": 0,
                    "topHoldings": [],
                    "rules": {
                        "source": (
                            "Official Prudential "
                            "Singapore factsheet"
                        ),
                        "maximumHoldings": MAX_HOLDINGS,
                        "multilineHoldingNames": True,
                        "syntheticData": False,
                        "calculatedData": False,
                        "inferredData": False,
                    },
                }

                save_json(
                    fund_dir
                    / "top_holdings.json",
                    [],
                )

                save_json(
                    fund_dir
                    / "metadata.json",
                    metadata,
                )

                return metadata

            (
                fund_dir
                / "top_holdings_section.txt"
            ).write_text(
                holdings_section,
                encoding="utf-8",
            )

            holdings = (
                parse_top_holdings(
                    holdings_section
                )
            )

            if not holdings:

                raise ValueError(
                    "Top 10 holdings section was found, "
                    "but no valid holding rows with "
                    "published percentages could be parsed."
                )

            # Validate rank ordering without changing the published
            # information. This is validation only.
            for position, holding in enumerate(
                holdings,
                start=1,
            ):

                rank = holding.get(
                    "rank"
                )

                if not isinstance(
                    rank,
                    int,
                ):

                    raise ValueError(
                        "Invalid holding rank "
                        f"at position {position}."
                    )

                if rank < 1:
                    raise ValueError(
                        "Holding rank must be >= 1."
                    )

                name = clean_text(
                    holding.get(
                        "name"
                    )
                )

                if not name:

                    raise ValueError(
                        "Holding name is empty "
                        f"at position {position}."
                    )

                weight = holding.get(
                    "weightPercent"
                )

                if not isinstance(
                    weight,
                    (int, float),
                ):

                    raise ValueError(
                        "Holding weight is invalid "
                        f"at position {position}."
                    )

                if (
                    weight < 0
                    or weight > 100
                ):

                    raise ValueError(
                        "Holding weight is outside "
                        f"0-100% at position {position}."
                    )

            metadata_dates = (
                extract_factsheet_metadata(
                    pdf_text
                )
            )

            metadata = {
                "status": "success",
                "excelRow": excel_row,
                "prudentialUrl": prudential_url,
                "finalUrl": final_url,
                "pageTitle": page_title,
                "fundName": fund_name,
                "fundIdentifier": fund_identifier,
                "excelPruAccessName": (
                    excel_pruaccess_name
                ),
                "factsheetUrl": factsheet_url,
                "factsheetDocumentDate": (
                    metadata_dates[
                        "factsheetDocumentDate"
                    ]
                ),
                "factsheetDataAsAt": (
                    metadata_dates[
                        "factsheetDataAsAt"
                    ]
                ),
                "factsheetPageCount": page_count,
                "holdingsSectionStatus": (
                    "found_and_parsed"
                ),
                "topHoldingsCount": len(
                    holdings
                ),
                "topHoldings": holdings,
                "rules": {
                    "source": (
                        "Official Prudential "
                        "Singapore factsheet"
                    ),
                    "maximumHoldings": MAX_HOLDINGS,
                    "multilineHoldingNames": True,
                    "joinedWrappedPdfLines": True,
                    "nameBoundary": (
                        "Published percentage"
                    ),
                    "fewerThanTenAllowed": True,
                    "syntheticData": False,
                    "calculatedData": False,
                    "inferredData": False,
                },
            }

            save_json(
                fund_dir
                / "top_holdings.json",
                holdings,
            )

            save_json(
                fund_dir
                / "metadata.json",
                metadata,
            )

            print(
                f"[Row {excel_row}] "
                f"SUCCESS - extracted "
                f"{len(holdings)} holdings"
            )

            for holding in holdings:

                print(
                    f"  {holding['rank']}. "
                    f"{holding['name']} "
                    f"| {holding['weightText']}"
                )

            return metadata

        except (
            PlaywrightTimeoutError,
            TimeoutError,
            ValueError,
            OSError,
        ) as exc:

            last_error = (
                f"{type(exc).__name__}: {exc}"
            )

            print(
                f"[Row {excel_row}] "
                f"Attempt {attempt} failed: "
                f"{last_error}"
            )

            if attempt < MAX_RETRIES:

                time.sleep(
                    RETRY_DELAY_SECONDS
                    * attempt
                )

        except Exception as exc:

            last_error = (
                f"{type(exc).__name__}: {exc}"
            )

            print(
                f"[Row {excel_row}] "
                f"Unexpected failure on attempt "
                f"{attempt}: "
                f"{last_error}"
            )

            if attempt < MAX_RETRIES:

                time.sleep(
                    RETRY_DELAY_SECONDS
                    * attempt
                )

    fund_identifier = (
        extract_identifier_from_url(
            prudential_url
        )
    )

    failure_dir = (
        make_failure_directory(
            excel_row,
            fund_identifier,
        )
    )

    failure_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    failure_record = {
        "status": "failed",
        "excelRow": excel_row,
        "prudentialUrl": prudential_url,
        "excelPruAccessName": (
            excel_pruaccess_name
        ),
        "fundIdentifier": fund_identifier,
        "error": last_error
        or "Unknown extraction error",
    }

    save_json(
        failure_dir
        / "error.json",
        failure_record,
    )

    return failure_record


# ============================================================================
# MAIN
# ============================================================================

def main() -> int:

    run_started = (
        datetime.now(
            timezone.utc
        )
    )

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FAILURES_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        "============================================================"
    )

    print(
        "PRUDENTIAL TOP 10 HOLDINGS"
    )

    print(
        "ALL-FUND EXTRACTION TEST"
    )

    print(
        "============================================================"
    )

    # ------------------------------------------------------------------------
    # Read Excel universe
    # ------------------------------------------------------------------------

    try:

        funds = read_excel_funds()

    except Exception as exc:

        print(
            f"ERROR reading Excel universe: {exc}"
        )

        return 1

    print(
        f"Excel fund universe: {len(funds)}"
    )

    if not funds:

        print(
            "ERROR: No populated Prudential URLs "
            "were found in Column A."
        )

        return 1

    successful_funds: List[
        Dict[str, Any]
    ] = []

    no_holdings_section_funds: List[
        Dict[str, Any]
    ] = []

    failed_funds: List[
        Dict[str, Any]
    ] = []

    all_holdings: List[
        Dict[str, Any]
    ] = []

    # ------------------------------------------------------------------------
    # Playwright
    # ------------------------------------------------------------------------

    with sync_playwright() as playwright:

        browser = playwright.chromium.launch(
            headless=True,
        )

        context = browser.new_context(
            viewport={
                "width": 1440,
                "height": 1000,
            },
            user_agent=USER_AGENT,
        )

        page = context.new_page()

        total = len(
            funds
        )

        for index, fund in enumerate(
            funds,
            start=1,
        ):

            print(
                f"\n\nPROCESSING FUND "
                f"{index}/{total}"
            )

            result = process_single_fund(
                page,
                fund,
            )

            status = result.get(
                "status"
            )

            if status == "success":

                successful_funds.append(
                    result
                )

                all_holdings.append(
                    {
                        "excelRow": result[
                            "excelRow"
                        ],
                        "fundName": result.get(
                            "fundName"
                        ),
                        "fundIdentifier": result.get(
                            "fundIdentifier"
                        ),
                        "prudentialUrl": result.get(
                            "prudentialUrl"
                        ),
                        "factsheetUrl": result.get(
                            "factsheetUrl"
                        ),
                        "factsheetDocumentDate": (
                            result.get(
                                "factsheetDocumentDate"
                            )
                        ),
                        "factsheetDataAsAt": (
                            result.get(
                                "factsheetDataAsAt"
                            )
                        ),
                        "topHoldingsCount": result[
                            "topHoldingsCount"
                        ],
                        "topHoldings": result[
                            "topHoldings"
                        ],
                    }
                )

            elif (
                status
                == "no_holdings_section"
            ):

                no_holdings_section_funds.append(
                    result
                )

                all_holdings.append(
                    {
                        "excelRow": result[
                            "excelRow"
                        ],
                        "fundName": result.get(
                            "fundName"
                        ),
                        "fundIdentifier": result.get(
                            "fundIdentifier"
                        ),
                        "prudentialUrl": result.get(
                            "prudentialUrl"
                        ),
                        "factsheetUrl": result.get(
                            "factsheetUrl"
                        ),
                        "factsheetDocumentDate": (
                            result.get(
                                "factsheetDocumentDate"
                            )
                        ),
                        "factsheetDataAsAt": (
                            result.get(
                                "factsheetDataAsAt"
                            )
                        ),
                        "topHoldingsCount": 0,
                        "topHoldings": [],
                    }
                )

            else:

                failed_funds.append(
                    result
                )

        browser.close()

    # ------------------------------------------------------------------------
    # Save consolidated holdings
    # ------------------------------------------------------------------------

    consolidated = {
        "status": (
            "success"
            if not failed_funds
            else "partial"
        ),
        "generatedAtUtc": utc_now_iso(),
        "excelFile": str(
            EXCEL_FILE
        ),
        "fundUniverseCount": len(
            funds
        ),
        "successfulFundCount": len(
            successful_funds
        ),
        "noHoldingsSectionFundCount": len(
            no_holdings_section_funds
        ),
        "failedFundCount": len(
            failed_funds
        ),
        "maximumPublishedHoldings": MAX_HOLDINGS,
        "rules": {
            "excelColumnAControlsUniverse": True,
            "officialPrudentialFactsheetOnly": True,
            "maximumHoldings": MAX_HOLDINGS,
            "fewerThanTenAllowed": True,
            "multilineHoldingNamesSupported": True,
            "wrappedPdfLinesJoined": True,
            "noSyntheticData": True,
            "noCalculatedHoldings": True,
            "noInferredHoldings": True,
        },
        "funds": all_holdings,
    }

    save_json(
        ALL_HOLDINGS_FILE,
        consolidated,
    )

    # ------------------------------------------------------------------------
    # Run summary
    # ------------------------------------------------------------------------

    run_finished = (
        datetime.now(
            timezone.utc
        )
    )

    total_holdings = sum(
        fund.get(
            "topHoldingsCount",
            0,
        )
        or 0
        for fund in all_holdings
    )

    run_summary = {
        "status": (
            "success"
            if not failed_funds
            else "partial"
        ),
        "startedAtUtc": (
            run_started.isoformat()
            .replace("+00:00", "Z")
        ),
        "completedAtUtc": (
            run_finished.isoformat()
            .replace("+00:00", "Z")
        ),
        "excelFile": str(
            EXCEL_FILE
        ),
        "fundUniverseCount": len(
            funds
        ),
        "successfulFundCount": len(
            successful_funds
        ),
        "noHoldingsSectionFundCount": len(
            no_holdings_section_funds
        ),
        "failedFundCount": len(
            failed_funds
        ),
        "totalPublishedHoldingsExtracted": (
            total_holdings
        ),
        "maximumHoldingsPerFund": MAX_HOLDINGS,
        "multilineHoldingNamesSupported": True,
        "successfulFunds": [
            {
                "excelRow": fund.get(
                    "excelRow"
                ),
                "fundName": fund.get(
                    "fundName"
                ),
                "fundIdentifier": fund.get(
                    "fundIdentifier"
                ),
                "topHoldingsCount": fund.get(
                    "topHoldingsCount"
                ),
            }
            for fund in successful_funds
        ],
        "noHoldingsSectionFunds": [
            {
                "excelRow": fund.get(
                    "excelRow"
                ),
                "fundName": fund.get(
                    "fundName"
                ),
                "fundIdentifier": fund.get(
                    "fundIdentifier"
                ),
            }
            for fund in no_holdings_section_funds
        ],
        "failedFunds": failed_funds,
    }

    save_json(
        RUN_SUMMARY_FILE,
        run_summary,
    )

    # ------------------------------------------------------------------------
    # Console summary
    # ------------------------------------------------------------------------

    print(
        "\n\n============================================================"
    )

    print(
        "ALL-FUND HOLDINGS EXTRACTION COMPLETE"
    )

    print(
        "============================================================"
    )

    print(
        f"Excel fund universe: "
        f"{len(funds)}"
    )

    print(
        f"Successful funds: "
        f"{len(successful_funds)}"
    )

    print(
        f"No Top 10 Holdings section: "
        f"{len(no_holdings_section_funds)}"
    )

    print(
        f"Failed funds: "
        f"{len(failed_funds)}"
    )

    print(
        f"Total published holdings extracted: "
        f"{total_holdings}"
    )

    print(
        "\nMultiline holding names: ENABLED"
    )

    print(
        "Holding name boundary: published percentage"
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

    print(
        f" - {FAILURES_OUTPUT_DIR}"
    )

    if failed_funds:

        print(
            "\nFAILED FUNDS:"
        )

        for failure in failed_funds:

            print(
                f" - Row "
                f"{failure.get('excelRow')}: "
                f"{failure.get('error')}"
            )

        print(
            "\nResult status: PARTIAL"
        )

        return 1

    print(
        "\nResult status: SUCCESS"
    )

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )
