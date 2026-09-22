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

Test official Prudential Top Holdings extraction for EVERY fund listed
in Excel Column A.

Workflow:

    Funds Links.xlsm
          |
          v
    Prudential fund page
          |
          v
    Official Prudential Fund Factsheet PDF
          |
          v
    Raw PDF text extraction
          |
          v
    Confirmed Top Holdings section
          |
          v
    Logical holding-line reconstruction
          |
          v
    Primary holdings parser
          |
          +---- failure ----> fallback holdings parser
          |
          v
    Validation
          |
          v
    output_holdings/


IMPORTANT HARD RULES
====================

1. Excel Column A controls the universe.

2. Every populated URL in Column A is processed.

3. There is NO hardcoded 67-fund limit.

4. Duplicate URLs are preserved and processed independently.

5. Only official Prudential Singapore pages are accepted.

6. Only official Prudential Singapore factsheets are accepted.

7. No third-party holdings sources are permitted.

8. No holdings may be inferred.

9. No holdings may be fabricated.

10. No percentages may be fabricated.

11. No synthetic holdings data is permitted.

12. If Prudential publishes fewer than 10 holdings, store exactly
    the published number.

13. Never force exactly 10 holdings.

14. Preserve Prudential's published holding order.

15. Multi-line holding names are joined.

16. Wrapped PDF lines are treated as one logical holding until the
    published portfolio percentage is encountered.

17. Fixed-income coupon/rate percentages inside security descriptions
    are NOT automatically treated as portfolio weights.

18. If a fixed-income logical line contains multiple percentages,
    the LAST percentage is treated as the portfolio weight by the
    fallback parser.

19. Fixed-income maturity dates may occur between a coupon/rate and
    the portfolio weight. Such lines are reconstructed before
    percentage parsing.

20. If the primary parser fails, the SAME official factsheet section
    is retried using the fallback parser.

21. If both parsers fail, the fund is FAILED.

22. If no confirmed Top Holdings section exists, the fund receives
    status "no_holdings_section".

23. Generic occurrences of "holdings" are NOT treated as a Top
    Holdings section.

24. Generic occurrences of "investments" are NOT treated as a Top
    Holdings section.

25. A confirmed heading may contain an attached numeric PDF footnote
    marker, for example:

        Top 10 Holdings3
        Top 10 Holdings4
        Top Ten Holdings3

26. The numeric footnote marker is part of the heading and is NOT
    part of the first holding.

27. No candidate heading is automatically accepted merely because
    it contains the words "holdings" or "investments".

28. Holdings weights are not interpreted beyond the published
    percentage extraction rules.

29. The full original PDF is preserved.

30. Raw extracted PDF text is preserved.

31. Raw Top Holdings section text is preserved.

32. This is a TEST collector only.

33. This script does not modify:
       - test_pruaccess.py
       - data.json
       - index.html
       - css/style.css
       - js/app.js

34. Whenever parser rules are changed, this entire script is intended
    to be replaced and rerun against the complete Excel universe.
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
    PlaywrightTimeoutError,
    sync_playwright,
)


# ============================================================
# CONFIGURATION
# ============================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_holdings")
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"

ALL_HOLDINGS_FILE = OUTPUT_DIR / "all_holdings.json"
RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"

BROWSER_HEADLESS = True

PAGE_TIMEOUT_MS = 120000
FACTSHEET_TIMEOUT_MS = 120000

POST_PAGE_WAIT_MS = 1500

MAX_RETRIES = 3
RETRY_DELAY_SECONDS = 3

MAX_HOLDINGS = 10

OFFICIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}


# ============================================================
# REGEX
# ============================================================

PERCENTAGE_PATTERN = re.compile(
    r"(?<![\d.])([0-9]+(?:\.[0-9]+)?)\s*%"
)

RANK_PATTERN = re.compile(
    r"^\s*(\d{1,2})(?:[.)\-:]|\s+)(.*)$"
)

STANDALONE_RANK_PATTERN = re.compile(
    r"^\s*(\d{1,2})[.)\-:]?\s*$"
)

PAGE_PATTERN = re.compile(
    r"^\s*page\s+\d+(?:\s+of\s+\d+)?\s*$",
    re.IGNORECASE,
)

MONTH_PATTERN = (
    r"(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)"
)

MATURITY_DATE_PATTERNS = [
    re.compile(
        rf"\b\d{{1,2}}[-\s]{{0,1}}{MONTH_PATTERN}"
        rf"[-\s]{{0,1}}\d{{2,4}}\b",
        re.IGNORECASE,
    ),
    re.compile(
        rf"\b\d{{1,2}}/{MONTH_PATTERN}/\d{{2,4}}\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"\b\d{1,2}/\d{1,2}/\d{2,4}\b",
        re.IGNORECASE,
    ),
    re.compile(
        rf"\b\d{{1,2}}\s+{MONTH_PATTERN}\s+\d{{2,4}}\b",
        re.IGNORECASE,
    ),
]

# ------------------------------------------------------------
# IMPORTANT:
#
# The heading may have a PDF footnote marker attached directly
# to it, e.g.
#
#     Top 10 Holdings3
#
# Therefore the numeric suffix is deliberately permitted.
#
# We do NOT permit arbitrary text after "Holdings".
# ------------------------------------------------------------

TOP_HOLDINGS_HEADING_PATTERN = re.compile(
    r"^\s*top\s+(?:10|ten)\s+holdings\s*\d*\s*$",
    re.IGNORECASE,
)


# ============================================================
# TEXT HELPERS
# ============================================================

def clean_text(value: str | None) -> str:
    if value is None:
        return ""

    value = value.replace("\u00a0", " ")
    value = value.replace("\u00ad", "")
    value = value.replace("\u200b", "")
    value = value.replace("\ufeff", "")

    value = re.sub(r"[ \t]+", " ", value)

    return value.strip()


def normalize_text(value: str | None) -> str:
    value = clean_text(value)

    value = value.replace("-", "-")
    value = value.replace("–", "-")
    value = value.replace("—", "-")
    value = value.replace("−", "-")

    return value.lower().strip()


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def save_json(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            payload,
            handle,
            ensure_ascii=False,
            indent=2,
        )


def safe_filename(value: str, max_length: int = 150) -> str:
    value = clean_text(value)

    value = re.sub(
        r"[<>:\"/\\|?*\x00-\x1F]",
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
        value = "fund"

    return value[:max_length]


# ============================================================
# URL VALIDATION
# ============================================================

def is_prudential_url(url: str) -> bool:
    if not url:
        return False

    try:
        parsed = urlparse(url)
    except Exception:
        return False

    if parsed.scheme.lower() not in {"http", "https"}:
        return False

    host = (parsed.hostname or "").lower()

    return host in OFFICIAL_HOSTS


def ensure_prudential_url(url: str) -> str:
    url = clean_text(url)

    if not url:
        raise ValueError("Empty Prudential URL.")

    if not is_prudential_url(url):
        raise ValueError(
            f"Non-official Prudential URL rejected: {url}"
        )

    return url


# ============================================================
# EXCEL
# ============================================================

def read_excel_funds() -> list[dict]:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        filename=EXCEL_FILE,
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
            url_value = worksheet.cell(
                row=row_number,
                column=1,
            ).value

            pruaccess_name = worksheet.cell(
                row=row_number,
                column=2,
            ).value

            url = clean_text(
                str(url_value)
                if url_value is not None
                else ""
            )

            name = clean_text(
                str(pruaccess_name)
                if pruaccess_name is not None
                else ""
            )

            if not url:
                continue

            funds.append(
                {
                    "excelRow": row_number,
                    "prudentialUrl": url,
                    "pruAccessName": name,
                }
            )

        return funds

    finally:
        workbook.close()


# ============================================================
# FACTSHEET DISCOVERY
# ============================================================

def find_factsheet_url(page) -> str:
    anchors = page.locator("a")

    candidates = []

    count = anchors.count()

    for index in range(count):
        anchor = anchors.nth(index)

        try:
            href = anchor.get_attribute("href")
            text = clean_text(anchor.inner_text())
        except Exception:
            continue

        if not href:
            continue

        absolute_url = urljoin(
            page.url,
            href,
        )

        if not is_prudential_url(absolute_url):
            continue

        lower_text = text.lower()
        lower_href = absolute_url.lower()

        score = 0

        if "fund factsheet" in lower_text:
            score += 100

        if "factsheet" in lower_text:
            score += 70

        if "fund factsheet" in lower_href:
            score += 80

        if "factsheet" in lower_href:
            score += 50

        if lower_href.endswith(".pdf"):
            score += 40

        if ".pdf?" in lower_href:
            score += 40

        if "fund-document" in lower_href:
            score += 30

        if score > 0:
            candidates.append(
                (
                    score,
                    absolute_url,
                    text,
                )
            )

    if not candidates:
        raise RuntimeError(
            "No official Prudential Fund Factsheet PDF link found."
        )

    candidates.sort(
        key=lambda item: (
            item[0],
            len(item[1]),
        ),
        reverse=True,
    )

    return candidates[0][1]


def download_factsheet(
    page,
    factsheet_url: str,
) -> bytes:
    last_error = None

    for attempt in range(
        1,
        MAX_RETRIES + 1,
    ):
        try:
            response = page.request.get(
                factsheet_url,
                timeout=FACTSHEET_TIMEOUT_MS,
            )

            if response.status != 200:
                raise RuntimeError(
                    f"Factsheet HTTP status: {response.status}"
                )

            pdf_bytes = response.body()

            if not pdf_bytes:
                raise RuntimeError(
                    "Factsheet response was empty."
                )

            if not pdf_bytes.startswith(b"%PDF"):
                raise RuntimeError(
                    "Factsheet response is not a PDF."
                )

            return pdf_bytes

        except Exception as exc:
            last_error = exc

            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY_SECONDS)

    raise RuntimeError(
        f"Unable to download factsheet after "
        f"{MAX_RETRIES} attempts: {last_error}"
    )


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf_text(
    pdf_bytes: bytes,
) -> tuple[str, int]:
    reader = PdfReader(
        BytesIO(pdf_bytes)
    )

    page_count = len(reader.pages)

    if page_count == 0:
        raise RuntimeError(
            "PDF contains zero pages."
        )

    page_texts = []

    for page_index, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):
        try:
            text = pdf_page.extract_text()
        except Exception as exc:
            raise RuntimeError(
                f"PDF text extraction failed on page "
                f"{page_index}: {exc}"
            ) from exc

        if text is None:
            text = ""

        page_texts.append(text)

    combined = "\n".join(
        page_texts
    )

    if not clean_text(combined):
        raise RuntimeError(
            "PDF contains no extractable text."
        )

    return combined, page_count


def pdf_lines(text: str) -> list[str]:
    return [
        clean_text(line)
        for line in text.splitlines()
        if clean_text(line)
    ]


# ============================================================
# PDF METADATA
# ============================================================

def extract_data_as_at(
    text: str,
) -> str | None:
    patterns = [
        re.compile(
            r"data\s+as\s+at\s*[:\-]?\s*(.+)",
            re.IGNORECASE,
        ),
        re.compile(
            r"as\s+at\s*[:\-]?\s*(.+)",
            re.IGNORECASE,
        ),
    ]

    lines = pdf_lines(text)

    for line in lines:
        for pattern in patterns:
            match = pattern.search(line)

            if match:
                value = clean_text(
                    match.group(1)
                )

                if value:
                    return value

    return None


def extract_document_date(
    text: str,
) -> str | None:
    lines = pdf_lines(text)

    date_patterns = [
        re.compile(
            r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b"
        ),
        re.compile(
            rf"\b\d{{1,2}}\s+{MONTH_PATTERN}"
            rf"\s+\d{{2,4}}\b",
            re.IGNORECASE,
        ),
        re.compile(
            rf"\b{MONTH_PATTERN}\s+\d{{4}}\b",
            re.IGNORECASE,
        ),
    ]

    for line in lines:
        lower = line.lower()

        if (
            "factsheet" not in lower
            and "fund factsheet" not in lower
            and "dated" not in lower
            and "date" not in lower
        ):
            continue

        for pattern in date_patterns:
            match = pattern.search(line)

            if match:
                return clean_text(
                    match.group(0)
                )

    return None


# ============================================================
# HOLDINGS HEADING DETECTION
# ============================================================

def is_confirmed_top_holdings_heading(
    line: str,
) -> bool:
    """
    Confirmed headings only.

    Valid:
        Top 10 Holdings
        Top 10 Holdings3
        Top 10 Holdings4
        Top Ten Holdings
        Top Ten Holdings3

    Invalid:
        Holdings
        Top Holdings
        Investment Holdings
        DBS GROUP HOLDINGS LTD
        Portfolio Holdings

    The optional numeric suffix represents a PDF footnote
    marker attached directly to the heading.
    """

    normalized = normalize_text(line)

    return bool(
        TOP_HOLDINGS_HEADING_PATTERN.fullmatch(
            normalized
        )
    )


def find_holdings_start(
    lines: list[str],
) -> int | None:
    for index, line in enumerate(lines):
        if is_confirmed_top_holdings_heading(line):
            return index

    return None


def is_holdings_end(line: str) -> bool:
    normalized = normalize_text(line)

    if not normalized:
        return False

    exact_end_markers = {
        "source",
        "source:",
        "inception date",
        "important information",
        "disclaimer",
        "past performance",
        "portfolio characteristics",
        "asset allocation",
    }

    if normalized in exact_end_markers:
        return True

    if normalized.startswith(
        "important information"
    ):
        return True

    if normalized.startswith(
        "disclaimer"
    ):
        return True

    if normalized.startswith(
        "past performance"
    ):
        return True

    if normalized.startswith(
        "portfolio characteristics"
    ):
        return True

    if normalized.startswith(
        "asset allocation"
    ):
        return True

    if PAGE_PATTERN.fullmatch(normalized):
        return True

    return False


def extract_holdings_section(
    text: str,
) -> tuple[str, str]:
    lines = pdf_lines(text)

    start_index = find_holdings_start(lines)

    if start_index is None:
        return "", "not_published"

    section_lines = []

    for line in lines[
        start_index + 1:
    ]:
        if is_holdings_end(line):
            break

        section_lines.append(line)

    section_text = "\n".join(
        section_lines
    ).strip()

    if not section_text:
        return "", "empty"

    return (
        section_text,
        "confirmed",
    )


# ============================================================
# HOLDING NAME / PERCENTAGE HELPERS
# ============================================================

def clean_holding_name(
    name: str,
) -> str:
    name = clean_text(name)

    name = re.sub(
        r"^[•·▪●◦‣]\s*",
        "",
        name,
    )

    name = re.sub(
        r"^\s*\d{1,2}[.)\-:]\s*",
        "",
        name,
    )

    name = name.replace("|", " ")

    name = re.sub(
        r"\bNone\b",
        "",
        name,
        flags=re.IGNORECASE,
    )

    name = re.sub(
        r"\s+",
        " ",
        name,
    )

    return name.strip()


def parse_percentage(
    value: str,
) -> float:
    number = float(value)

    if number < 0 or number > 100:
        raise ValueError(
            f"Invalid percentage: {number}%"
        )

    return number


def percentage_matches(
    line: str,
) -> list[re.Match]:
    return list(
        PERCENTAGE_PATTERN.finditer(line)
    )


def find_percentage_in_line(
    line: str,
) -> tuple[float, int, int] | None:
    matches = percentage_matches(line)

    if not matches:
        return None

    match = matches[0]

    return (
        parse_percentage(
            match.group(1)
        ),
        match.start(),
        match.end(),
    )


def find_last_percentage_in_line(
    line: str,
) -> tuple[float, int, int] | None:
    matches = percentage_matches(line)

    if not matches:
        return None

    match = matches[-1]

    return (
        parse_percentage(
            match.group(1)
        ),
        match.start(),
        match.end(),
    )


# ============================================================
# MATURITY DATE DETECTION
# ============================================================

def contains_maturity_date(
    line: str,
) -> bool:
    for pattern in MATURITY_DATE_PATTERNS:
        if pattern.search(line):
            return True

    return False


def line_ends_with_percentage(
    line: str,
) -> bool:
    return bool(
        re.search(
            r"\d+(?:\.\d+)?\s*%\s*$",
            line,
        )
    )


def line_has_percentage(
    line: str,
) -> bool:
    return bool(
        percentage_matches(line)
    )


# ============================================================
# LOGICAL HOLDING LINE RECONSTRUCTION
# ============================================================

def build_logical_holding_lines(
    section_text: str,
) -> list[str]:
    """
    Reconstruct PDF-wrapped fixed-income holdings.

    Important example:

        CORPORACION ANDINA DE FOMENTO 7.7%
        6-MAR2029 1.7%

    becomes:

        CORPORACION ANDINA DE FOMENTO 7.7% 6-MAR2029 1.7%

    Also supports a three-line reconstruction:

        SECURITY 7.7%
        6-MAR2029
        1.7%

    This prevents the coupon/rate from being mistaken for the
    portfolio weight.
    """

    lines = [
        clean_text(line)
        for line in section_text.splitlines()
        if clean_text(line)
    ]

    result = []

    index = 0

    while index < len(lines):
        current = lines[index]

        # ----------------------------------------------------
        # Pattern:
        #
        # current ends in %
        # next line contains maturity date + percentage
        # ----------------------------------------------------
        if (
            line_ends_with_percentage(current)
            and index + 1 < len(lines)
            and contains_maturity_date(lines[index + 1])
            and line_has_percentage(lines[index + 1])
        ):
            combined = (
                current
                + " "
                + lines[index + 1]
            )

            result.append(
                clean_text(combined)
            )

            index += 2
            continue

        # ----------------------------------------------------
        # Pattern:
        #
        # current ends in %
        # next contains maturity date
        # following contains percentage
        # ----------------------------------------------------
        if (
            line_ends_with_percentage(current)
            and index + 2 < len(lines)
            and contains_maturity_date(lines[index + 1])
            and line_has_percentage(lines[index + 2])
        ):
            combined = (
                current
                + " "
                + lines[index + 1]
                + " "
                + lines[index + 2]
            )

            result.append(
                clean_text(combined)
            )

            index += 3
            continue

        result.append(current)

        index += 1

    return result


# ============================================================
# RANK HANDLING
# ============================================================

def extract_leading_rank(
    line: str,
) -> tuple[int | None, str]:
    match = RANK_PATTERN.match(line)

    if not match:
        return None, line

    rank = int(
        match.group(1)
    )

    remainder = clean_text(
        match.group(2)
    )

    if rank < 1 or rank > 99:
        return None, line

    return rank, remainder


def is_standalone_rank(
    line: str,
) -> int | None:
    match = STANDALONE_RANK_PATTERN.fullmatch(
        line
    )

    if not match:
        return None

    rank = int(
        match.group(1)
    )

    if rank < 1 or rank > 99:
        return None

    return rank


# ============================================================
# HOLDING NOISE FILTER
# ============================================================

def is_holding_header_or_noise(
    line: str,
) -> bool:
    normalized = normalize_text(line)

    if not normalized:
        return True

    if is_confirmed_top_holdings_heading(
        normalized
    ):
        return True

    noise_exact = {
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

    if normalized in noise_exact:
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


def clean_holding_fragment(
    line: str,
) -> str:
    line = clean_text(line)

    if not line:
        return ""

    if is_holding_header_or_noise(line):
        return ""

    rank, remainder = extract_leading_rank(
        line
    )

    if rank is not None:
        line = remainder

    return clean_holding_name(line)


def combine_holding_name_fragments(
    fragments: list[str],
) -> str:
    cleaned = []

    for fragment in fragments:
        fragment = clean_holding_fragment(
            fragment
        )

        if fragment:
            cleaned.append(fragment)

    return clean_holding_name(
        " ".join(cleaned)
    )


# ============================================================
# VALIDATION
# ============================================================

def validate_holdings(
    holdings: list[dict],
) -> None:
    if not holdings:
        raise ValueError(
            "No holdings were extracted."
        )

    if len(holdings) > MAX_HOLDINGS:
        raise ValueError(
            f"Extracted {len(holdings)} holdings; "
            f"maximum permitted is {MAX_HOLDINGS}."
        )

    expected_ranks = list(
        range(
            1,
            len(holdings) + 1,
        )
    )

    actual_ranks = [
        holding["rank"]
        for holding in holdings
    ]

    if actual_ranks != expected_ranks:
        raise ValueError(
            "Holding ranks are not sequential. "
            f"Expected {expected_ranks}, "
            f"received {actual_ranks}."
        )

    for holding in holdings:
        name = clean_text(
            holding.get("name", "")
        )

        if not name:
            raise ValueError(
                f"Holding {holding['rank']} "
                "has an empty name."
            )

        weight = holding.get(
            "weightPercent"
        )

        if not isinstance(
            weight,
            (int, float),
        ):
            raise ValueError(
                f"Holding {holding['rank']} "
                "has invalid portfolio weight."
            )

        if weight < 0 or weight > 100:
            raise ValueError(
                f"Holding {holding['rank']} "
                f"has invalid portfolio weight: "
                f"{weight}%"
            )


# ============================================================
# PRIMARY PARSER
# ============================================================

def parse_holdings(
    section_text: str,
) -> list[dict]:
    """
    Primary parser.

    Uses the FIRST percentage in a logical holding line.

    If multiple percentages appear in the same logical line,
    the primary parser deliberately fails and the fallback
    parser is used.

    This is intentional. It prevents the parser from silently
    guessing which percentage is the portfolio weight.
    """

    logical_lines = build_logical_holding_lines(
        section_text
    )

    holdings = []

    pending_fragments = []
    pending_rank = None

    for line in logical_lines:
        line = clean_text(line)

        if not line:
            continue

        if is_holding_header_or_noise(line):
            continue

        rank, remainder = extract_leading_rank(
            line
        )

        if rank is not None:
            if (
                pending_fragments
                and pending_rank is not None
            ):
                raise ValueError(
                    "Encountered a new holding rank "
                    "before the previous holding was "
                    "committed."
                )

            if rank != len(holdings) + 1:
                raise ValueError(
                    f"Unexpected holding rank {rank}; "
                    f"expected "
                    f"{len(holdings) + 1}."
                )

            pending_rank = rank
            line = remainder

        else:
            standalone_rank = is_standalone_rank(
                line
            )

            if standalone_rank is not None:
                if pending_fragments:
                    raise ValueError(
                        "Standalone rank encountered "
                        "while previous holding remains "
                        "pending."
                    )

                if (
                    standalone_rank
                    != len(holdings) + 1
                ):
                    raise ValueError(
                        f"Unexpected holding rank "
                        f"{standalone_rank}; expected "
                        f"{len(holdings) + 1}."
                    )

                pending_rank = standalone_rank
                continue

        matches = percentage_matches(line)

        if not matches:
            fragment = clean_holding_fragment(
                line
            )

            if fragment:
                pending_fragments.append(
                    fragment
                )

            continue

        if len(matches) > 1:
            raise ValueError(
                "Primary parser encountered multiple "
                f"percentages in one logical line: {line}"
            )

        match = matches[0]

        weight = parse_percentage(
            match.group(1)
        )

        name_fragment = clean_holding_fragment(
            line[:match.start()]
        )

        if name_fragment:
            pending_fragments.append(
                name_fragment
            )

        name = combine_holding_name_fragments(
            pending_fragments
        )

        if not name:
            raise ValueError(
                f"Unable to determine holding name "
                f"for portfolio weight {weight}%."
            )

        rank_to_use = (
            pending_rank
            if pending_rank is not None
            else len(holdings) + 1
        )

        holdings.append(
            {
                "rank": rank_to_use,
                "name": name,
                "weightPercent": weight,
            }
        )

        pending_fragments = []
        pending_rank = None

        if len(holdings) >= MAX_HOLDINGS:
            break

    validate_holdings(
        holdings
    )

    return holdings


# ============================================================
# FALLBACK PARSER
# ============================================================

def parse_holdings_fallback(
    section_text: str,
) -> list[dict]:
    """
    Fallback parser.

    Uses the LAST percentage in a logical holding line.

    This is required for fixed-income securities where a
    security description may contain:

        coupon/rate %
        maturity date
        portfolio weight %

    Example:

        CORPORACION ANDINA DE FOMENTO 7.7%
        6-MAR2029 1.7%

    After logical reconstruction:

        CORPORACION ANDINA DE FOMENTO 7.7% 6-MAR2029 1.7%

    LAST percentage = 1.7%, which is the portfolio weight.
    """

    logical_lines = build_logical_holding_lines(
        section_text
    )

    holdings = []

    pending_fragments = []
    pending_rank = None

    for line in logical_lines:
        line = clean_text(line)

        if not line:
            continue

        if is_holding_header_or_noise(line):
            continue

        rank, remainder = extract_leading_rank(
            line
        )

        if rank is not None:
            if (
                pending_fragments
                and pending_rank is not None
            ):
                raise ValueError(
                    "Encountered a new holding rank "
                    "before the previous holding was "
                    "committed."
                )

            if rank != len(holdings) + 1:
                raise ValueError(
                    f"Unexpected holding rank {rank}; "
                    f"expected "
                    f"{len(holdings) + 1}."
                )

            pending_rank = rank
            line = remainder

        else:
            standalone_rank = is_standalone_rank(
                line
            )

            if standalone_rank is not None:
                if pending_fragments:
                    raise ValueError(
                        "Standalone rank encountered "
                        "while previous holding remains "
                        "pending."
                    )

                if (
                    standalone_rank
                    != len(holdings) + 1
                ):
                    raise ValueError(
                        f"Unexpected holding rank "
                        f"{standalone_rank}; expected "
                        f"{len(holdings) + 1}."
                    )

                pending_rank = standalone_rank
                continue

        matches = percentage_matches(line)

        if not matches:
            fragment = clean_holding_fragment(
                line
            )

            if fragment:
                pending_fragments.append(
                    fragment
                )

            continue

        match = matches[-1]

        weight = parse_percentage(
            match.group(1)
        )

        name_fragment = clean_holding_fragment(
            line[:match.start()]
        )

        if name_fragment:
            pending_fragments.append(
                name_fragment
            )

        name = combine_holding_name_fragments(
            pending_fragments
        )

        if not name:
            raise ValueError(
                f"Unable to determine holding name "
                f"for portfolio weight {weight}%."
            )

        rank_to_use = (
            pending_rank
            if pending_rank is not None
            else len(holdings) + 1
        )

        holdings.append(
            {
                "rank": rank_to_use,
                "name": name,
                "weightPercent": weight,
            }
        )

        pending_fragments = []
        pending_rank = None

        if len(holdings) >= MAX_HOLDINGS:
            break

    validate_holdings(
        holdings
    )

    return holdings


# ============================================================
# FUND EXTRACTION
# ============================================================

def extract_single_fund(
    page,
    fund: dict,
) -> dict:
    excel_row = fund["excelRow"]
    prudential_url = ensure_prudential_url(
        fund["prudentialUrl"]
    )

    page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

    try:
        page.wait_for_load_state(
            "networkidle",
            timeout=30000,
        )
    except PlaywrightTimeoutError:
        pass

    if POST_PAGE_WAIT_MS > 0:
        page.wait_for_timeout(
            POST_PAGE_WAIT_MS
        )

    final_url = page.url

    if not is_prudential_url(final_url):
        raise RuntimeError(
            "Prudential product page redirected "
            f"outside official Prudential Singapore: "
            f"{final_url}"
        )

    # --------------------------------------------------------
    # Fund name
    # --------------------------------------------------------

    fund_name = ""

    try:
        h1 = page.locator("h1").first

        if h1.count() > 0:
            fund_name = clean_text(
                h1.inner_text()
            )
    except Exception:
        fund_name = ""

    if not fund_name:
        try:
            body_text = clean_text(
                page.locator("body").inner_text()
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
        except Exception:
            fund_name = ""

    if not fund_name:
        fund_name = fund.get(
            "pruAccessName"
        ) or f"Excel Row {excel_row}"

    # --------------------------------------------------------
    # Factsheet
    # --------------------------------------------------------

    factsheet_url = find_factsheet_url(
        page
    )

    if not is_prudential_url(
        factsheet_url
    ):
        raise RuntimeError(
            "Discovered factsheet is not hosted "
            "on official Prudential Singapore."
        )

    pdf_bytes = download_factsheet(
        page,
        factsheet_url,
    )

    pdf_text, pdf_page_count = extract_pdf_text(
        pdf_bytes
    )

    # --------------------------------------------------------
    # Metadata
    # --------------------------------------------------------

    data_as_at = extract_data_as_at(
        pdf_text
    )

    document_date = extract_document_date(
        pdf_text
    )

    # --------------------------------------------------------
    # Holdings section
    # --------------------------------------------------------

    holdings_section, section_status = (
        extract_holdings_section(
            pdf_text
        )
    )

    result = {
        "status": "success",
        "excelRow": excel_row,
        "fundName": fund_name,
        "prudentialUrl": prudential_url,
        "finalPrudentialUrl": final_url,
        "pruAccessName": fund.get(
            "pruAccessName",
            "",
        ),
        "factsheetUrl": factsheet_url,
        "pdfPageCount": pdf_page_count,
        "dataAsAt": data_as_at,
        "documentDate": document_date,
        "holdingsSectionStatus": section_status,
        "topHoldingsCount": 0,
        "holdings": [],
        "parser": None,
        "rules": {
            "excelColumnAControlsUniverse": True,
            "processEveryPopulatedColumnAUrl": True,
            "hardcodedFundLimit": False,
            "duplicateUrlsPreserved": True,
            "officialPrudentialSingaporeOnly": True,
            "officialFactsheetsOnly": True,
            "thirdPartySources": False,
            "confirmedTopHoldingsDetection": True,
            "topHoldingsFootnoteSuffixAccepted": True,
            "genericHoldingsHeadingRejected": True,
            "genericInvestmentsHeadingRejected": True,
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

    if section_status != "confirmed":
        result["status"] = (
            "no_holdings_section"
        )

        return result, pdf_bytes, pdf_text, ""

    # --------------------------------------------------------
    # Primary parser
    # --------------------------------------------------------

    primary_error = None

    try:
        holdings = parse_holdings(
            holdings_section
        )

        result["holdings"] = holdings
        result["topHoldingsCount"] = len(
            holdings
        )
        result["parser"] = "primary"

        return (
            result,
            pdf_bytes,
            pdf_text,
            holdings_section,
        )

    except Exception as exc:
        primary_error = str(exc)

    # --------------------------------------------------------
    # Fallback parser
    # --------------------------------------------------------

    try:
        holdings = parse_holdings_fallback(
            holdings_section
        )

        result["holdings"] = holdings
        result["topHoldingsCount"] = len(
            holdings
        )
        result["parser"] = "fallback"
        result["primaryParserError"] = (
            primary_error
        )

        return (
            result,
            pdf_bytes,
            pdf_text,
            holdings_section,
        )

    except Exception as fallback_exc:
        raise RuntimeError(
            "Both holdings parsers failed. "
            f"Primary error: {primary_error}; "
            f"Fallback error: {fallback_exc}"
        ) from fallback_exc


# ============================================================
# OUTPUT HELPERS
# ============================================================

def fund_output_directory(
    excel_row: int,
    fund_name: str,
) -> Path:
    identifier = safe_filename(
        fund_name
    )

    return (
        FUNDS_OUTPUT_DIR
        / f"{excel_row}_{identifier}"
    )


def save_success_result(
    result: dict,
    pdf_bytes: bytes,
    pdf_text: str,
    holdings_section: str,
) -> None:
    output_dir = fund_output_directory(
        result["excelRow"],
        result["fundName"],
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    factsheet_path = (
        output_dir
        / "factsheet.pdf"
    )

    factsheet_text_path = (
        output_dir
        / "factsheet_text.txt"
    )

    holdings_section_path = (
        output_dir
        / "top_holdings_section.txt"
    )

    holdings_json_path = (
        output_dir
        / "top_holdings.json"
    )

    metadata_path = (
        output_dir
        / "metadata.json"
    )

    factsheet_path.write_bytes(
        pdf_bytes
    )

    factsheet_text_path.write_text(
        pdf_text,
        encoding="utf-8",
    )

    if holdings_section:
        holdings_section_path.write_text(
            holdings_section,
            encoding="utf-8",
        )

    save_json(
        holdings_json_path,
        {
            "excelRow": result["excelRow"],
            "fundName": result["fundName"],
            "factsheetUrl": result[
                "factsheetUrl"
            ],
            "topHoldingsCount": result[
                "topHoldingsCount"
            ],
            "parser": result.get(
                "parser"
            ),
            "holdings": result[
                "holdings"
            ],
        },
    )

    save_json(
        metadata_path,
        result,
    )


def save_no_holdings_result(
    result: dict,
    pdf_bytes: bytes,
    pdf_text: str,
) -> None:
    output_dir = fund_output_directory(
        result["excelRow"],
        result["fundName"],
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    (
        output_dir
        / "factsheet.pdf"
    ).write_bytes(
        pdf_bytes
    )

    (
        output_dir
        / "factsheet_text.txt"
    ).write_text(
        pdf_text,
        encoding="utf-8",
    )

    save_json(
        output_dir
        / "top_holdings.json",
        {
            "excelRow": result["excelRow"],
            "fundName": result["fundName"],
            "status": "no_holdings_section",
            "topHoldingsCount": 0,
            "holdings": [],
        },
    )

    save_json(
        output_dir
        / "metadata.json",
        result,
    )


def save_failure_result(
    fund: dict,
    error: str,
) -> None:
    excel_row = fund["excelRow"]

    identifier = safe_filename(
        fund.get(
            "pruAccessName"
        )
        or f"excel_row_{excel_row}"
    )

    output_dir = (
        OUTPUT_DIR
        / "failed"
        / f"{excel_row}_{identifier}"
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    save_json(
        output_dir
        / "failure.json",
        {
            "status": "failed",
            "excelRow": excel_row,
            "prudentialUrl": fund[
                "prudentialUrl"
            ],
            "pruAccessName": fund.get(
                "pruAccessName",
                "",
            ),
            "error": error,
        },
    )


# ============================================================
# MAIN
# ============================================================

def main() -> int:
    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    funds = read_excel_funds()

    if not funds:
        raise RuntimeError(
            "No populated Prudential URLs found "
            "in Excel Column A."
        )

    print(
        "=" * 70
    )
    print(
        "VGrat FMS - Prudential Top Holdings "
        "ALL-FUND TEST EXTRACTOR"
    )
    print(
        "=" * 70
    )
    print(
        f"Excel fund universe: {len(funds)}"
    )
    print(
        "Universe source: Funds Links.xlsm "
        "Column A"
    )
    print(
        "Confirmed heading rule:"
    )
    print(
        "  Top 10 Holdings + optional numeric "
        "footnote marker"
    )
    print(
        "  Top Ten Holdings + optional numeric "
        "footnote marker"
    )
    print(
        "=" * 70
    )

    successful_results = []
    failed_results = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=BROWSER_HEADLESS
        )

        context = browser.new_context(
            accept_downloads=True
        )

        page = context.new_page()

        page.set_default_timeout(
            PAGE_TIMEOUT_MS
        )

        for position, fund in enumerate(
            funds,
            start=1,
        ):
            excel_row = fund[
                "excelRow"
            ]

            print()
            print(
                "-" * 70
            )
            print(
                f"[{position}/{len(funds)}] "
                f"Excel Row {excel_row}"
            )
            print(
                fund[
                    "prudentialUrl"
                ]
            )

            try:
                (
                    result,
                    pdf_bytes,
                    pdf_text,
                    holdings_section,
                ) = extract_single_fund(
                    page,
                    fund,
                )

                if result[
                    "status"
                ] == "success":
                    save_success_result(
                        result,
                        pdf_bytes,
                        pdf_text,
                        holdings_section,
                    )

                    successful_results.append(
                        result
                    )

                    print(
                        "STATUS: SUCCESS"
                    )

                    print(
                        "Fund: "
                        f"{result['fundName']}"
                    )

                    print(
                        "Top Holdings: "
                        f"{result['topHoldingsCount']}"
                    )

                    print(
                        "Parser: "
                        f"{result.get('parser')}"
                    )

                    for holding in result[
                        "holdings"
                    ]:
                        print(
                            f"  "
                            f"{holding['rank']}. "
                            f"{holding['name']} - "
                            f"{holding['weightPercent']}%"
                        )

                elif result[
                    "status"
                ] == "no_holdings_section":
                    save_no_holdings_result(
                        result,
                        pdf_bytes,
                        pdf_text,
                    )

                    successful_results.append(
                        result
                    )

                    print(
                        "STATUS: "
                        "NO HOLDINGS SECTION"
                    )

                    print(
                        "Fund: "
                        f"{result['fundName']}"
                    )

                else:
                    raise RuntimeError(
                        "Unexpected extraction "
                        f"status: "
                        f"{result.get('status')}"
                    )

            except Exception as exc:
                error_message = str(exc)

                failed_results.append(
                    {
                        "status": "failed",
                        "excelRow": excel_row,
                        "fundName": fund.get(
                            "pruAccessName",
                            "",
                        ),
                        "prudentialUrl": fund[
                            "prudentialUrl"
                        ],
                        "error": error_message,
                    }
                )

                save_failure_result(
                    fund,
                    error_message,
                )

                print(
                    "STATUS: FAILED"
                )

                print(
                    f"ERROR: {error_message}"
                )

        context.close()
        browser.close()

    # ========================================================
    # FINAL VERIFICATION
    # ========================================================

    verified_successful = []
    verification_failed = []

    for result in successful_results:
        if result[
            "status"
        ] != "success":
            verified_successful.append(
                result
            )
            continue

        try:
            page_url = result[
                "factsheetUrl"
            ]

            if not is_prudential_url(
                page_url
            ):
                raise RuntimeError(
                    "Factsheet URL failed "
                    "official Prudential host "
                    "verification during "
                    "final validation."
                )

            verified_successful.append(
                result
            )

        except Exception as exc:
            verification_failed.append(
                {
                    "status": "failed",
                    "excelRow": result[
                        "excelRow"
                    ],
                    "fundName": result[
                        "fundName"
                    ],
                    "prudentialUrl": result[
                        "prudentialUrl"
                    ],
                    "error": (
                        "Final verification failed: "
                        + str(exc)
                    ),
                }
            )

    all_results = (
        verified_successful
        + failed_results
        + verification_failed
    )

    all_results.sort(
        key=lambda item: item.get(
            "excelRow",
            999999,
        )
    )

    successful_count = len(
        verified_successful
    )

    failed_count = (
        len(failed_results)
        + len(verification_failed)
    )

    holdings_found = sum(
        1
        for result in verified_successful
        if result.get(
            "status"
        ) == "success"
        and result.get(
            "topHoldingsCount",
            0,
        ) > 0
    )

    holdings_missing = sum(
        1
        for result in verified_successful
        if result.get(
            "status"
        ) == "no_holdings_section"
    )

    total_holdings = sum(
        result.get(
            "topHoldingsCount",
            0,
        )
        for result in verified_successful
        if result.get(
            "status"
        ) == "success"
    )

    total_rows_with_results = len(
        all_results
    )

    run_summary = {
        "status": (
            "success"
            if failed_count == 0
            else "failed"
        ),
        "generatedAtUtc": utc_now_iso(),
        "excelFile": str(
            EXCEL_FILE
        ),
        "excelFundUniverse": len(
            funds
        ),
        "successfulFunds": successful_count,
        "failedFunds": failed_count,
        "topHoldingsFound": holdings_found,
        "topHoldingsMissing": holdings_missing,
        "totalHoldings": total_holdings,
        "totalRowsWithResults": (
            total_rows_with_results
        ),
        "failedRows": [
            item["excelRow"]
            for item in (
                failed_results
                + verification_failed
            )
        ],
        "successfulRows": [
            item["excelRow"]
            for item in verified_successful
        ],
        "purpose": (
            "Test official Prudential "
            "Top Holdings extraction "
            "for every populated Excel "
            "Column A URL."
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
            "topHoldingsFootnoteSuffixAccepted": True,
            "genericHoldingsHeadingRejected": True,
            "genericInvestmentsHeadingRejected": True,
            "primaryParser": True,
            "fallbackParser": True,
            "fixedIncomeMaturityReconstruction": True,
            "lastPercentageFallback": True,
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
        ALL_HOLDINGS_FILE,
        {
            "generatedAtUtc": utc_now_iso(),
            "excelFundUniverse": len(
                funds
            ),
            "results": all_results,
        },
    )

    save_json(
        RUN_SUMMARY_FILE,
        run_summary,
    )

    # ========================================================
    # CONSOLE SUMMARY
    # ========================================================

    print()
    print(
        "=" * 70
    )
    print(
        "FINAL SUMMARY"
    )
    print(
        "=" * 70
    )
    print(
        f"Excel fund universe : {len(funds)}"
    )
    print(
        f"Successful funds    : "
        f"{successful_count}"
    )
    print(
        f"Failed funds        : "
        f"{failed_count}"
    )
    print(
        f"Top Holdings found  : "
        f"{holdings_found}"
    )
    print(
        f"Top Holdings missing: "
        f"{holdings_missing}"
    )
    print(
        f"Total holdings      : "
        f"{total_holdings}"
    )

    if failed_results or verification_failed:
        print()
        print(
            "FAILED ROWS:"
        )

        for failure in (
            failed_results
            + verification_failed
        ):
            print(
                f"  Row "
                f"{failure['excelRow']}: "
                f"{failure['error']}"
            )

    print()
    print(
        "Output:"
    )
    print(
        f"  {ALL_HOLDINGS_FILE}"
    )
    print(
        f"  {RUN_SUMMARY_FILE}"
    )
    print(
        f"  {FUNDS_OUTPUT_DIR}"
    )

    print(
        "=" * 70
    )

    return (
        0
        if failed_count == 0
        else 1
    )


if __name__ == "__main__":
    try:
        sys.exit(
            main()
        )
    except KeyboardInterrupt:
        print(
            "\nInterrupted by user.",
            file=sys.stderr,
        )
        sys.exit(130)
    except Exception as exc:
        print(
            f"\nFATAL ERROR: {exc}",
            file=sys.stderr,
        )
        sys.exit(1)
