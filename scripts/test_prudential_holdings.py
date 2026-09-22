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
    Prudential fund page
          |
          v
    Official Prudential factsheet PDF
          |
          v
    Extract PDF text
          |
          v
    Locate "Top 10 Holdings"
          |
          v
    Primary parser
          |
          +---- if ambiguous ----> rebuilt fallback parser
          |
          v
    Validate
          |
          v
    Save individual fund result
          |
          v
    all_holdings.json
    run_summary.json


HARD RULES
==========

1. Funds Links.xlsm is the MASTER SOURCE.

2. Column A controls the entire fund universe.

3. Process EVERY populated URL in Column A.
   Do not hardcode a fund count.

4. Column B is the exact PruAccess fund name/reference name.

5. Only official Prudential Singapore sources may be used.

6. No third-party holdings data.

7. No inferred holdings.

8. No fabricated holdings.

9. No fabricated percentages.

10. Do not force exactly 10 holdings.

11. If Prudential publishes fewer than 10 holdings, preserve the exact
    published number.

12. Preserve published holding order.

13. Duplicate holding names are allowed.

14. Duplicate percentages are allowed.

15. Fixed-income coupon/rate percentages are part of the security name
    and MUST NOT automatically be treated as the portfolio weight.

16. The actual portfolio weight is the percentage associated with the
    holding's portfolio position, normally the LAST percentage before
    the next holding begins.

17. A holding may span multiple PDF-extracted lines.

18. Multiple holdings may appear on one PDF-extracted line.

19. A maturity date may be split from the rest of the security name.

20. Wrapped PDF lines must be joined when they belong to the same
    logical holding.

21. No holdings may be created merely because a percentage appears.

22. If a Top 10 Holdings section does not exist:
       status = "no_holdings_section"

23. If holdings cannot be safely parsed:
       the fund must fail.
       Never guess.

24. Primary parser is intentionally conservative.

25. If primary parser encounters multiple percentages on a logical line,
    fallback parser is used.

26. Fallback parser is specifically designed for fixed-income holdings
    where coupon/rate percentages and portfolio-weight percentages
    coexist.

27. Fallback parser must be able to separate multiple holdings appearing
    on one physical PDF line.

28. No synthetic holdings.

29. No synthetic weights.

30. No interpolation.

31. No carry-forward.

32. This script does NOT modify:
       test_pruaccess.py
       data.json
       index.html
       css
       js

33. Output is written under:

       output_holdings/

    including:

       output_holdings/all_holdings.json
       output_holdings/run_summary.json
       output_holdings/funds/
       output_holdings/<row>_failed/

34. Whenever this script is changed, replace the ENTIRE script with the
    complete version provided.


OUTPUT STRUCTURE
================

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


DEPENDENCIES
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
from pathlib import Path
from urllib.parse import urljoin, urlparse

from openpyxl import load_workbook
from pypdf import PdfReader
from playwright.sync_api import sync_playwright


# ============================================================================
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_holdings")
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"

ALL_HOLDINGS_FILE = OUTPUT_DIR / "all_holdings.json"
RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"

HEADLESS = True

PAGE_TIMEOUT_MS = 120_000
PDF_DOWNLOAD_TIMEOUT_MS = 120_000

WAIT_AFTER_PAGE_LOAD_MS = 1_500

MAX_RETRIES = 3
RETRY_DELAY_SECONDS = 3

MAX_HOLDINGS = 10

OFFICIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}


# ============================================================================
# REGEX
# ============================================================================

PERCENT_RE = re.compile(
    r"(?<![\d.])"
    r"(\d+(?:\.\d+)?)"
    r"\s*%"
)

RANK_RE = re.compile(
    r"""
    ^
    \s*
    (?:
        (?P<number>\d{1,2})
        \s*
        (?:[.)]|[-:]\s*|(?=\s))
    )
    \s*
    """,
    re.IGNORECASE | re.VERBOSE,
)

RANK_ONLY_RE = re.compile(
    r"^\s*\d{1,2}\s*[.)]?\s*$"
)

DATE_TOKEN_RE = re.compile(
    r"""
    \b
    \d{1,2}
    [-/]
    (?:
        JAN|FEB|MAR|APR|MAY|JUN|
        JUL|AUG|SEP|OCT|NOV|DEC
    )
    -?
    \d{4}
    \b
    """,
    re.IGNORECASE | re.VERBOSE,
)

YEAR_ONLY_RE = re.compile(
    r"\b(?:19|20)\d{2}\b"
)

MONTH_RE = re.compile(
    r"\b(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b",
    re.IGNORECASE,
)

TOP_HOLDINGS_RE = re.compile(
    r"\btop\s+(?:10|ten)\s+holdings?\b",
    re.IGNORECASE,
)

PAGE_NUMBER_RE = re.compile(
    r"^\s*page\s+\d+(?:\s+of\s+\d+)?\s*$",
    re.IGNORECASE,
)


# ============================================================================
# GENERAL HELPERS
# ============================================================================

def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def safe_filename(value: str, max_length: int = 160) -> str:
    value = value or "unknown"
    value = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", value)
    value = re.sub(r"\s+", "_", value)
    value = value.strip("._ ")

    if not value:
        value = "unknown"

    return value[:max_length]


def normalize_whitespace(value: str) -> str:
    value = value.replace("\u00a0", " ")
    value = value.replace("\u200b", "")
    value = value.replace("\ufeff", "")
    value = re.sub(r"[ \t]+", " ", value)
    return value.strip()


def normalize_pdf_line(line: str) -> str:
    line = line.replace("\u00a0", " ")
    line = line.replace("\u200b", "")
    line = line.replace("\ufeff", "")

    # Replace common PDF extraction separators.
    line = line.replace("│", " ")
    line = line.replace("|", " ")

    # Normalize repeated whitespace.
    line = re.sub(r"\s+", " ", line)

    return line.strip()


def is_official_prudential_url(url: str) -> bool:
    try:
        parsed = urlparse(url)
        host = (parsed.hostname or "").lower()
        return host in OFFICIAL_HOSTS
    except Exception:
        return False


def ensure_official_url(url: str, description: str = "URL") -> None:
    if not is_official_prudential_url(url):
        raise RuntimeError(
            f"{description} is not an official Prudential Singapore URL: {url}"
        )


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", encoding="utf-8") as handle:
        json.dump(
            data,
            handle,
            ensure_ascii=False,
            indent=2,
        )


def clean_holding_name(name: str) -> str:
    if not name:
        return ""

    name = name.replace("\u00a0", " ")
    name = name.replace("\u200b", "")
    name = name.replace("\ufeff", "")

    # Remove PDF bullets.
    name = re.sub(r"^[•·▪◦●■□◆◇►]+\s*", "", name)

    # Remove obvious leading rank.
    name = re.sub(
        r"^\s*\d{1,2}\s*(?:[.)]|[-:])\s*",
        "",
        name,
    )

    # Remove standalone "None" occasionally produced by PDF extraction.
    name = re.sub(
        r"^\s*None\s+",
        "",
        name,
        flags=re.IGNORECASE,
    )

    # Remove trailing separators.
    name = re.sub(r"\s*[-|:]+\s*$", "", name)

    # Normalize whitespace.
    name = re.sub(r"\s+", " ", name)

    return name.strip()


def extract_leading_rank(line: str) -> int | None:
    match = RANK_RE.match(line)

    if not match:
        return None

    try:
        return int(match.group("number"))
    except Exception:
        return None


def count_percentages(line: str) -> int:
    return len(PERCENT_RE.findall(line))


def find_percentage_matches(line: str):
    return list(PERCENT_RE.finditer(line))


def find_percentage_in_line(line: str):
    matches = find_percentage_matches(line)

    if not matches:
        return None

    match = matches[0]

    try:
        value = float(match.group(1))
    except Exception:
        return None

    return {
        "value": value,
        "text": match.group(0),
        "start": match.start(),
        "end": match.end(),
    }


def find_last_percentage_in_line(line: str):
    matches = find_percentage_matches(line)

    if not matches:
        return None

    match = matches[-1]

    try:
        value = float(match.group(1))
    except Exception:
        return None

    return {
        "value": value,
        "text": match.group(0),
        "start": match.start(),
        "end": match.end(),
    }


def contains_maturity_date(line: str) -> bool:
    return bool(DATE_TOKEN_RE.search(line))


def contains_year(line: str) -> bool:
    return bool(YEAR_ONLY_RE.search(line))


def contains_month(line: str) -> bool:
    return bool(MONTH_RE.search(line))


def has_date_like_fragment(line: str) -> bool:
    """
    Detect fixed-income maturity/date fragments.

    Examples:

        5-AUG-2054
        31-DEC-2079
        6-MAR2029
        31-DEC2079

    Also recognizes a year-only continuation line such as:

        2029

    because PDF extraction can split:

        CORPORACION ANDINA DE FOMENTO 7.7%
        6-MAR
        2029 1.7%

    """
    if contains_maturity_date(line):
        return True

    stripped = normalize_pdf_line(line)

    if YEAR_ONLY_RE.fullmatch(stripped):
        return True

    if re.search(
        r"\b\d{1,2}[-/](?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b",
        line,
        flags=re.IGNORECASE,
    ):
        return True

    return False


def looks_like_portfolio_weight(value: float) -> bool:
    """
    Portfolio weights in the top holdings table should be positive and
    normally <= 100.

    This function is deliberately permissive because the script must not
    impose an arbitrary maximum such as 10%.

    We only reject impossible percentages.
    """
    return value > 0 and value <= 100


# ============================================================================
# EXCEL
# ============================================================================

def read_excel_funds() -> list[dict]:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel master file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    try:
        worksheet = workbook.active

        funds = []

        for row_number in range(2, worksheet.max_row + 1):
            url_value = worksheet.cell(row=row_number, column=1).value
            pruaccess_value = worksheet.cell(row=row_number, column=2).value

            if url_value is None:
                continue

            url = str(url_value).strip()

            if not url:
                continue

            if not is_official_prudential_url(url):
                raise RuntimeError(
                    f"Excel row {row_number} contains a non-official "
                    f"Prudential URL: {url}"
                )

            funds.append(
                {
                    "excelRow": row_number,
                    "prudentialUrl": url,
                    "pruAccessName": (
                        str(pruaccess_value).strip()
                        if pruaccess_value is not None
                        else ""
                    ),
                }
            )

        return funds

    finally:
        workbook.close()


# ============================================================================
# PRUDENTIAL PAGE / FACTSHEET
# ============================================================================

def score_factsheet_anchor(anchor) -> int:
    try:
        text = normalize_whitespace(anchor.inner_text())
    except Exception:
        text = ""

    try:
        href = anchor.get_attribute("href") or ""
    except Exception:
        href = ""

    combined = f"{text} {href}".lower()

    score = 0

    if "fund factsheet" in combined:
        score += 100

    if "factsheet" in combined:
        score += 50

    if href.lower().endswith(".pdf"):
        score += 25

    if "fund-documents" in href.lower():
        score += 20

    return score


def find_factsheet_url(page, fund_url: str) -> str:
    anchors = page.locator("a")

    candidates = []

    try:
        count = anchors.count()
    except Exception:
        count = 0

    for index in range(count):
        anchor = anchors.nth(index)

        try:
            href = anchor.get_attribute("href")
        except Exception:
            continue

        if not href:
            continue

        absolute_url = urljoin(fund_url, href)

        if not is_official_prudential_url(absolute_url):
            continue

        score = score_factsheet_anchor(anchor)

        if score <= 0:
            continue

        candidates.append(
            (
                score,
                absolute_url,
            )
        )

    if not candidates:
        raise RuntimeError(
            "Could not find an official Prudential factsheet PDF link."
        )

    candidates.sort(
        key=lambda item: (
            item[0],
            item[1].lower(),
        ),
        reverse=True,
    )

    return candidates[0][1]


def extract_pdf_text(pdf_bytes: bytes) -> str:
    if not pdf_bytes:
        raise RuntimeError("Factsheet PDF is empty.")

    reader = PdfReader(
        __import__("io").BytesIO(pdf_bytes)
    )

    pages = []

    for page_number, page in enumerate(reader.pages, start=1):
        try:
            text = page.extract_text() or ""
        except Exception as exc:
            raise RuntimeError(
                f"Failed extracting PDF page {page_number}: {exc}"
            ) from exc

        pages.append(text)

    full_text = "\n".join(pages)

    if not full_text.strip():
        raise RuntimeError(
            "Factsheet PDF produced no extractable text."
        )

    return full_text


# ============================================================================
# FACTSHEET DATES
# ============================================================================

def extract_data_as_at(text: str) -> str | None:
    patterns = [
        r"data\s+as\s+at\s*[:\-]?\s*"
        r"(\d{1,2}\s+"
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"\s+\d{4})",

        r"as\s+at\s*[:\-]?\s*"
        r"(\d{1,2}\s+"
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"\s+\d{4})",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match:
            return normalize_whitespace(match.group(1))

    return None


def extract_document_date(text: str) -> str | None:
    patterns = [
        r"\b(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+\d{4}\b",

        r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"\s+\d{4}\b",
    ]

    matches = re.findall(
        "|".join(f"(?:{pattern})" for pattern in patterns),
        text,
        flags=re.IGNORECASE,
    )

    if not matches:
        return None

    # The factsheet normally contains the document date near the end.
    # Use the last date-like month/year occurrence.
    month_year_re = re.compile(
        r"\b(?:January|February|March|April|May|June|July|August|"
        r"September|October|November|December|"
        r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"\s+\d{4}\b",
        re.IGNORECASE,
    )

    date_matches = month_year_re.findall(text)

    if not date_matches:
        return None

    return normalize_whitespace(date_matches[-1])


# ============================================================================
# TOP HOLDINGS SECTION
# ============================================================================

def find_holdings_start(lines: list[str]) -> int | None:
    for index, line in enumerate(lines):
        if TOP_HOLDINGS_RE.search(line):
            return index

    return None


def is_holdings_end(line: str) -> bool:
    normalized = normalize_pdf_line(line)
    lowered = normalized.lower()

    if not normalized:
        return False

    if PAGE_NUMBER_RE.match(normalized):
        return True

    endings = [
        "source:",
        "source -",
        "inception date",
        "important information",
        "important information:",
        "disclaimer",
        "past performance",
        "portfolio characteristics",
        "asset allocation",
        "geographical allocation",
        "sector allocation",
        "fund characteristics",
        "fund performance",
        "investment objective",
        "investment strategy",
    ]

    for ending in endings:
        if lowered.startswith(ending):
            return True

    return False


def extract_holdings_section(text: str) -> tuple[str, list[str]]:
    raw_lines = text.splitlines()

    lines = [
        normalize_pdf_line(line)
        for line in raw_lines
    ]

    start_index = find_holdings_start(lines)

    if start_index is None:
        raise LookupError(
            "no_holdings_section"
        )

    section_lines = []

    for index in range(start_index + 1, len(lines)):
        line = lines[index]

        if is_holdings_end(line):
            break

        section_lines.append(line)

    # Remove leading/trailing blank lines.
    while section_lines and not section_lines[0]:
        section_lines.pop(0)

    while section_lines and not section_lines[-1]:
        section_lines.pop()

    section_text = "\n".join(section_lines)

    if not section_text.strip():
        raise RuntimeError(
            "Top 10 Holdings heading was found but its section was empty."
        )

    return section_text, section_lines


# ============================================================================
# PRIMARY PARSER
# ============================================================================

def build_primary_logical_lines(lines: list[str]) -> list[str]:
    """
    Conservative logical-line builder.

    The primary parser is intentionally not aggressive. It only joins
    obvious continuation lines that do not contain a new percentage.

    The fallback parser handles the difficult fixed-income cases.
    """
    logical_lines = []

    pending = ""

    for raw_line in lines:
        line = normalize_pdf_line(raw_line)

        if not line:
            continue

        if RANK_ONLY_RE.match(line):
            if pending:
                logical_lines.append(pending)
                pending = ""

            logical_lines.append(line)
            continue

        if not pending:
            pending = line
            continue

        previous_has_percent = bool(PERCENT_RE.search(pending))
        current_has_percent = bool(PERCENT_RE.search(line))

        if previous_has_percent and current_has_percent:
            logical_lines.append(pending)
            pending = line
            continue

        if previous_has_percent and not current_has_percent:
            pending = f"{pending} {line}"
            continue

        pending = f"{pending} {line}"

    if pending:
        logical_lines.append(pending)

    return logical_lines


def parse_holdings_primary(lines: list[str]) -> list[dict]:
    """
    Conservative parser.

    If a logical line contains multiple percentages, fail intentionally
    rather than guessing.

    This triggers the rebuilt fallback parser.
    """
    logical_lines = build_primary_logical_lines(lines)

    holdings = []

    pending_fragments = []
    pending_rank = None

    for line in logical_lines:
        line = normalize_pdf_line(line)

        if not line:
            continue

        rank = extract_leading_rank(line)

        if rank is not None:
            line_without_rank = RANK_RE.sub("", line).strip()

            if pending_fragments:
                pending_name = clean_holding_name(
                    " ".join(pending_fragments)
                )

                if pending_name:
                    raise RuntimeError(
                        "Primary parser encountered a new ranked holding "
                        "before the previous holding had a portfolio weight."
                    )

                pending_fragments = []

            pending_rank = rank

            if not line_without_rank:
                continue

            line = line_without_rank

        percentage_count = count_percentages(line)

        if percentage_count > 1:
            raise RuntimeError(
                "Primary parser found multiple percentages on one logical "
                "line and will not guess the holding weight. "
                "Fallback parser required."
            )

        percentage = find_percentage_in_line(line)

        if percentage is None:
            pending_fragments.append(line)
            continue

        name_fragment = line[:percentage["start"]].strip()

        if name_fragment:
            pending_fragments.append(name_fragment)

        holding_name = clean_holding_name(
            " ".join(pending_fragments)
        )

        if not holding_name:
            raise RuntimeError(
                "Primary parser found a percentage without a holding name."
            )

        weight = percentage["value"]

        if not looks_like_portfolio_weight(weight):
            raise RuntimeError(
                f"Primary parser found an invalid percentage: {weight}%"
            )

        holdings.append(
            {
                "rank": pending_rank
                if pending_rank is not None
                else len(holdings) + 1,
                "name": holding_name,
                "weightPercent": weight,
            }
        )

        pending_fragments = []
        pending_rank = None

        if len(holdings) >= MAX_HOLDINGS:
            break

    if pending_fragments:
        leftover = clean_holding_name(
            " ".join(pending_fragments)
        )

        if leftover:
            raise RuntimeError(
                "Primary parser ended with an unweighted holding fragment: "
                f"{leftover}"
            )

    if not holdings:
        raise RuntimeError(
            "Primary parser extracted zero holdings."
        )

    return holdings[:MAX_HOLDINGS]


# ============================================================================
# REBUILT FALLBACK PARSER
# ============================================================================

def split_percentage_segments(line: str) -> list[dict]:
    """
    Split a line around EVERY percentage.

    Example:

        KEPPEL LTD 2.9% 31-DEC-2079 2.3%

    becomes:

        prefix = "KEPPEL LTD "
        percentage = 2.9%
        text_after = " 31-DEC-2079 "
        percentage = 2.3%

    This is essential because the first percentage may be a coupon.
    """
    matches = find_percentage_matches(line)

    if not matches:
        return []

    segments = []

    for index, match in enumerate(matches):
        start = match.start()
        end = match.end()

        prefix_start = (
            matches[index - 1].end()
            if index > 0
            else 0
        )

        prefix = line[prefix_start:start].strip()

        try:
            value = float(match.group(1))
        except Exception:
            continue

        segments.append(
            {
                "prefix": prefix,
                "value": value,
                "text": match.group(0),
                "start": start,
                "end": end,
            }
        )

    return segments


def looks_like_new_holding_start(fragment: str) -> bool:
    """
    Determine whether text following a portfolio weight appears to begin
    another security.

    This is deliberately conservative.

    A new holding is strongly indicated by:

    - a leading rank
    - a known security/company-style name
    - an uppercase-heavy fragment
    - a line containing no date continuation

    We do NOT require a particular company name.
    """
    fragment = normalize_pdf_line(fragment)

    if not fragment:
        return False

    if extract_leading_rank(fragment) is not None:
        return True

    # If the fragment begins with a clear uppercase security/company name,
    # it is likely the next holding.
    first_word_match = re.match(
        r"^([A-Z][A-Z0-9&.,'()/\-]*)\b",
        fragment,
    )

    if first_word_match:
        first_word = first_word_match.group(1)

        # Date-like fragments should NOT be treated as a new holding.
        if re.fullmatch(
            r"\d{1,2}[-/]?[A-Z]{3}[-/]?\d{0,4}",
            first_word,
            flags=re.IGNORECASE,
        ):
            return False

        if first_word.isupper():
            return True

    return False


def merge_fixed_income_continuations(lines: list[str]) -> list[str]:
    """
    Rebuild physical PDF lines into logical lines.

    Key problem solved here:

        CORPORACION ANDINA DE FOMENTO 7.7%
        6-MAR2029 1.7%

    becomes:

        CORPORACION ANDINA DE FOMENTO 7.7% 6-MAR2029 1.7%

    Also handles:

        MAPLETREE TREASURY SERVICES LTD 3.95%
        31-DEC2079 1.7%

    and split date fragments such as:

        CORPORACION ANDINA DE FOMENTO 7.7%
        6-MAR
        2029 1.7%

    The merge occurs only when the continuation strongly resembles a
    maturity/date fragment and contains a subsequent percentage or is
    itself clearly part of a date.
    """
    cleaned = []

    for line in lines:
        normalized = normalize_pdf_line(line)

        if normalized:
            cleaned.append(normalized)

    logical = []

    index = 0

    while index < len(cleaned):
        current = cleaned[index]

        # Rank-only lines are retained.
        if RANK_ONLY_RE.match(current):
            logical.append(current)
            index += 1
            continue

        combined = current
        j = index + 1

        while j < len(cleaned):
            next_line = cleaned[j]

            if not next_line:
                j += 1
                continue

            # Never absorb an explicit new rank.
            if extract_leading_rank(next_line) is not None:
                break

            current_percentages = find_percentage_matches(combined)
            next_percentages = find_percentage_matches(next_line)

            if not current_percentages:
                # Normal text continuation.
                combined = f"{combined} {next_line}"

                # Stop if the next line itself now completes a weighted
                # holding and the following line is clearly a new holding.
                if next_percentages:
                    j += 1
                    break

                j += 1
                continue

            last_current = current_percentages[-1]

            # If current line already has a percentage and the next line
            # contains a date-like fragment, it may be the continuation
            # of a fixed-income security.
            next_has_date = has_date_like_fragment(next_line)

            # Strongest case:
            #
            # current: COMPANY 7.7%
            # next:    6-MAR2029 1.7%
            #
            # Merge.
            if next_has_date and next_percentages:
                combined = f"{combined} {next_line}"
                j += 1
                break

            # Date fragment without the final portfolio weight:
            #
            # current: COMPANY 7.7%
            # next:    6-MAR
            # next:    2029 1.7%
            #
            # Merge one or more date continuation lines.
            if next_has_date and not next_percentages:
                combined = f"{combined} {next_line}"

                j += 1

                # Look ahead for a final weight.
                while j < len(cleaned):
                    continuation = cleaned[j]

                    if extract_leading_rank(continuation) is not None:
                        break

                    continuation_has_percent = bool(
                        find_percentage_matches(continuation)
                    )

                    continuation_has_date = has_date_like_fragment(
                        continuation
                    )

                    if continuation_has_percent or continuation_has_date:
                        combined = (
                            f"{combined} {continuation}"
                        )
                        j += 1

                        if continuation_has_percent:
                            break

                        continue

                    break

                break

            # If current line ends in a percentage and the next line is
            # another holding, do NOT merge.
            break

        logical.append(combined)
        index = max(j, index + 1)

    return logical


def parse_fallback_logical_line(
    line: str,
    pending_fragments: list[str],
    pending_rank: int | None,
    holdings: list[dict],
):
    """
    Parse a rebuilt logical line.

    Important behavior:

    For:

        COMPANY 5.5% DATE 2.5% NEXT COMPANY 6.0% DATE 2.4%

    the parser identifies two holdings:

        COMPANY 5.5% DATE    2.5%
        NEXT COMPANY 6.0% DATE    2.4%

    The first percentage in each security can therefore remain in the
    security name.
    """
    line = normalize_pdf_line(line)

    if not line:
        return pending_fragments, pending_rank

    # Remove explicit leading rank, preserving it.
    rank = extract_leading_rank(line)

    if rank is not None:
        if pending_fragments:
            pending_name = clean_holding_name(
                " ".join(pending_fragments)
            )

            if pending_name:
                # A new ranked holding appeared before a portfolio weight.
                raise RuntimeError(
                    "Fallback parser found a new ranked holding before "
                    f"the previous holding had a weight: {pending_name}"
                )

            pending_fragments = []

        pending_rank = rank
        line = RANK_RE.sub("", line).strip()

        if not line:
            return pending_fragments, pending_rank

    matches = find_percentage_matches(line)

    if not matches:
        pending_fragments.append(line)
        return pending_fragments, pending_rank

    # ----------------------------------------------------------------------
    # IMPORTANT:
    #
    # Multiple percentages may represent:
    #
    #     coupon + maturity + portfolio weight
    #
    # OR:
    #
    #     previous holding weight + next holding coupon + next holding weight
    #
    # We therefore process the line segment-by-segment.
    # ----------------------------------------------------------------------

    if len(matches) == 1:
        match = matches[0]

        before = line[:match.start()].strip()
        after = line[match.end():].strip()

        if before:
            pending_fragments.append(before)

        # If text follows this percentage and looks like a maturity/date,
        # the percentage is probably a coupon and the actual weight is
        # on the continuation.
        if after and (
            has_date_like_fragment(after)
            or contains_year(after)
        ):
            pending_fragments.append(
                f"{match.group(0)} {after}"
            )

            # There is no portfolio weight yet.
            return pending_fragments, pending_rank

        # If we already have a fixed-income security name containing a
        # coupon percentage and this percentage is at the end, it is the
        # portfolio weight.
        if pending_fragments:
            holding_name = clean_holding_name(
                " ".join(pending_fragments)
            )

            if not holding_name:
                raise RuntimeError(
                    "Fallback parser found a weight without a name."
                )

            weight = float(match.group(1))

            if not looks_like_portfolio_weight(weight):
                raise RuntimeError(
                    f"Fallback parser found invalid weight: {weight}%"
                )

            holdings.append(
                {
                    "rank": pending_rank
                    if pending_rank is not None
                    else len(holdings) + 1,
                    "name": holding_name,
                    "weightPercent": weight,
                }
            )

            pending_fragments = []
            pending_rank = None

            return pending_fragments, pending_rank

        # No previous fragments.
        #
        # This can be a simple holding:
        #
        #     APPLE INC 2.5%
        #
        # so everything before the percentage is the name.
        holding_name = clean_holding_name(before)

        if not holding_name:
            raise RuntimeError(
                "Fallback parser found a percentage without a name."
            )

        weight = float(match.group(1))

        if not looks_like_portfolio_weight(weight):
            raise RuntimeError(
                f"Fallback parser found invalid weight: {weight}%"
            )

        holdings.append(
            {
                "rank": pending_rank
                if pending_rank is not None
                else len(holdings) + 1,
                "name": holding_name,
                "weightPercent": weight,
            }
        )

        pending_fragments = []
        pending_rank = None

        return pending_fragments, pending_rank

    # ----------------------------------------------------------------------
    # MULTIPLE PERCENTAGES
    #
    # We now identify which percentages are likely security-name coupon
    # percentages and which one is the actual portfolio weight.
    # ----------------------------------------------------------------------

    cursor = 0

    local_pending = list(pending_fragments)
    local_rank = pending_rank

    for match_index, match in enumerate(matches):
        before = line[cursor:match.start()].strip()

        if before:
            local_pending.append(before)

        percentage_text = match.group(0)
        percentage_value = float(match.group(1))

        after_start = match.end()

        if match_index + 1 < len(matches):
            next_match = matches[match_index + 1]

            between = line[after_start:next_match.start()].strip()

            # If text between two percentages contains a maturity/date,
            # then the current percentage is a coupon and the NEXT
            # percentage is more likely to be the portfolio weight.
            if between and (
                has_date_like_fragment(between)
                or contains_year(between)
            ):
                local_pending.append(
                    f"{percentage_text}"
                )

                if between:
                    local_pending.append(between)

                cursor = next_match.start()
                continue

            # No maturity/date between the percentages.
            #
            # It may be:
            #
            #     previous holding weight
            #     next holding coupon
            #
            # In that case, if we already have a complete holding name,
            # commit the current percentage as the weight.
            if local_pending:
                candidate_name = clean_holding_name(
                    " ".join(local_pending)
                )

                if candidate_name:
                    # Look at the text immediately after this percentage.
                    trailing = line[
                        match.end():next_match.start()
                    ].strip()

                    # If the next segment begins like a new security name,
                    # the current percentage belongs to the previous
                    # holding.
                    if trailing and looks_like_new_holding_start(
                        trailing
                    ):
                        holdings.append(
                            {
                                "rank": local_rank
                                if local_rank is not None
                                else len(holdings) + 1,
                                "name": candidate_name,
                                "weightPercent": percentage_value,
                            }
                        )

                        local_pending = []
                        local_rank = None

                        cursor = match.end()
                        continue

                # Otherwise conservatively keep this percentage as part
                # of the current security name.
                local_pending.append(percentage_text)

                cursor = match.end()
                continue

        # Last percentage on this logical line.
        #
        # By design, this is the actual portfolio weight.
        local_pending_name = clean_holding_name(
            " ".join(local_pending)
        )

        if local_pending_name:
            holdings.append(
                {
                    "rank": local_rank
                    if local_rank is not None
                    else len(holdings) + 1,
                    "name": local_pending_name,
                    "weightPercent": percentage_value,
                }
            )

            local_pending = []
            local_rank = None

        cursor = match.end()

    trailing_text = line[cursor:].strip()

    if trailing_text:
        local_pending.append(trailing_text)

    return local_pending, local_rank


def parse_holdings_fallback(lines: list[str]) -> list[dict]:
    """
    Rebuilt fallback parser.

    This parser is designed specifically for the difficult cases caused
    by PDF text extraction.

    It first reconstructs logical fixed-income lines and then uses the
    LAST percentage as the portfolio weight whenever a logical holding
    contains multiple percentages.

    It also separates multiple holdings appearing on one physical line.
    """
    logical_lines = merge_fixed_income_continuations(lines)

    holdings = []

    pending_fragments = []
    pending_rank = None

    for line in logical_lines:
        pending_fragments, pending_rank = parse_fallback_logical_line(
            line=line,
            pending_fragments=pending_fragments,
            pending_rank=pending_rank,
            holdings=holdings,
        )

        if len(holdings) >= MAX_HOLDINGS:
            break

    # ----------------------------------------------------------------------
    # If we have an unweighted fragment after the final holding, determine
    # whether it is a date continuation or an actual malformed holding.
    # ----------------------------------------------------------------------

    if pending_fragments:
        leftover = clean_holding_name(
            " ".join(pending_fragments)
        )

        if leftover:
            # A pure date/year continuation after a completed holding is
            # not enough to create another holding.
            if not (
                has_date_like_fragment(leftover)
                or YEAR_ONLY_RE.fullmatch(leftover)
            ):
                raise RuntimeError(
                    "Fallback parser ended with an unweighted holding "
                    f"fragment: {leftover}"
                )

    if not holdings:
        raise RuntimeError(
            "Fallback parser extracted zero holdings."
        )

    return holdings[:MAX_HOLDINGS]


# ============================================================================
# HOLDINGS VALIDATION
# ============================================================================

def validate_holdings(holdings: list[dict]) -> None:
    if not holdings:
        raise RuntimeError(
            "No holdings extracted."
        )

    if len(holdings) > MAX_HOLDINGS:
        raise RuntimeError(
            f"Extractor returned more than {MAX_HOLDINGS} holdings."
        )

    expected_rank = 1

    for holding in holdings:
        if not isinstance(holding, dict):
            raise RuntimeError(
                "Holding is not an object."
            )

        name = holding.get("name")
        weight = holding.get("weightPercent")

        if not name or not str(name).strip():
            raise RuntimeError(
                "Holding has an empty name."
            )

        if weight is None:
            raise RuntimeError(
                f"Holding has no weight: {name}"
            )

        try:
            numeric_weight = float(weight)
        except Exception:
            raise RuntimeError(
                f"Holding has invalid weight: {name} -> {weight}"
            )

        if not looks_like_portfolio_weight(numeric_weight):
            raise RuntimeError(
                f"Holding has invalid percentage: "
                f"{name} -> {numeric_weight}%"
            )

        # Rank is informational and should remain sequential in extracted
        # order. We do not reject duplicate names or duplicate percentages.
        holding["rank"] = expected_rank
        holding["name"] = clean_holding_name(str(name))
        holding["weightPercent"] = numeric_weight

        expected_rank += 1


# ============================================================================
# FUND EXTRACTION
# ============================================================================

def extract_single_fund(
    browser,
    fund: dict,
) -> dict:
    excel_row = fund["excelRow"]
    prudential_url = fund["prudentialUrl"]
    pruaccess_name = fund["pruAccessName"]

    print("")
    print("=" * 60)
    print(f"PROCESSING FUND ROW {excel_row}")
    print("=" * 60)
    print("")
    print(f"Prudential URL: {prudential_url}")
    print(f"Excel PruAccess name: {pruaccess_name}")

    ensure_official_url(
        prudential_url,
        "Prudential fund URL",
    )

    page = browser.new_page()

    try:
        page.set_default_timeout(PAGE_TIMEOUT_MS)

        page.goto(
            prudential_url,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )

        page.wait_for_timeout(
            WAIT_AFTER_PAGE_LOAD_MS
        )

        factsheet_url = find_factsheet_url(
            page,
            prudential_url,
        )

        ensure_official_url(
            factsheet_url,
            "Factsheet URL",
        )

        print("")
        print(f"Factsheet link found:")
        print(f"{factsheet_url}")

        print("")
        print("Downloading factsheet...")

        response = page.request.get(
            factsheet_url,
            timeout=PDF_DOWNLOAD_TIMEOUT_MS,
        )

        if not response.ok:
            raise RuntimeError(
                f"Factsheet download failed: "
                f"HTTP {response.status}"
            )

        pdf_bytes = response.body()

        if not pdf_bytes:
            raise RuntimeError(
                "Factsheet download returned zero bytes."
            )

        print(
            f"Factsheet bytes: {len(pdf_bytes):,}"
        )

        pdf_text = extract_pdf_text(
            pdf_bytes
        )

        data_as_at = extract_data_as_at(
            pdf_text
        )

        document_date = extract_document_date(
            pdf_text
        )

        section_text, section_lines = extract_holdings_section(
            pdf_text
        )

        # --------------------------------------------------------------
        # PRIMARY PARSER
        # --------------------------------------------------------------

        parser_used = "primary"

        try:
            holdings = parse_holdings_primary(
                section_lines
            )

            validate_holdings(
                holdings
            )

        except Exception as primary_error:
            print("")
            print(
                "Primary holdings parser failed; "
                "running fallback parser..."
            )

            print(
                f"Primary parser error: {primary_error}"
            )

            parser_used = "fallback"

            holdings = parse_holdings_fallback(
                section_lines
            )

            validate_holdings(
                holdings
            )

        # --------------------------------------------------------------
        # SECOND VALIDATION
        # --------------------------------------------------------------

        if len(holdings) == 0:
            raise RuntimeError(
                "No holdings were extracted."
            )

        if len(holdings) > MAX_HOLDINGS:
            raise RuntimeError(
                f"More than {MAX_HOLDINGS} holdings extracted."
            )

        result = {
            "status": "success",
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "excelPruAccessName": pruaccess_name,
            "fundName": None,
            "factsheetUrl": factsheet_url,
            "factsheetDataAsAt": data_as_at,
            "factsheetDocumentDate": document_date,
            "parser": parser_used,
            "holdings": holdings,
        }

        # Attempt to identify fund name from PDF.
        fund_name = extract_fund_name_from_text(
            pdf_text
        )

        if fund_name:
            result["fundName"] = fund_name

        return {
            "result": result,
            "pdfBytes": pdf_bytes,
            "pdfText": pdf_text,
            "sectionText": section_text,
            "factsheetUrl": factsheet_url,
        }

    finally:
        try:
            page.close()
        except Exception:
            pass


# ============================================================================
# FUND NAME
# ============================================================================

def extract_fund_name_from_text(text: str) -> str | None:
    lines = [
        normalize_pdf_line(line)
        for line in text.splitlines()
    ]

    candidates = []

    for index, line in enumerate(lines[:80]):
        if not line:
            continue

        lowered = line.lower()

        if "prulink" in lowered:
            candidates.append(line)

            # Sometimes the fund name wraps onto the next line.
            if index + 1 < len(lines):
                next_line = lines[index + 1]

                if (
                    next_line
                    and "fund" in next_line.lower()
                    and len(next_line) < 160
                ):
                    combined = normalize_whitespace(
                        f"{line} {next_line}"
                    )
                    candidates.append(combined)

    if not candidates:
        return None

    # Prefer a candidate containing both PRULink and Fund.
    for candidate in candidates:
        lowered = candidate.lower()

        if (
            "prulink" in lowered
            and "fund" in lowered
        ):
            return candidate

    return candidates[0]


# ============================================================================
# OUTPUT
# ============================================================================

def get_output_identifier(
    result: dict,
) -> str:
    fund_name = result.get("fundName")

    if fund_name:
        return safe_filename(
            fund_name
        )

    url = result.get(
        "prudentialUrl",
        "",
    )

    parsed = urlparse(url)

    path = parsed.path.rstrip("/")

    if path:
        slug = path.split("/")[-1]

        if slug.lower().endswith(".html"):
            slug = slug[:-5]

        if slug:
            return safe_filename(slug)

    return "fund"


def save_success_result(
    result: dict,
    pdf_bytes: bytes,
    pdf_text: str,
    section_text: str,
) -> Path:
    excel_row = result["excelRow"]

    identifier = get_output_identifier(
        result
    )

    fund_dir = (
        FUNDS_OUTPUT_DIR
        / f"{excel_row}_{identifier}"
    )

    fund_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    pdf_path = fund_dir / "factsheet.pdf"

    with pdf_path.open("wb") as handle:
        handle.write(pdf_bytes)

    text_path = fund_dir / "factsheet_text.txt"

    text_path.write_text(
        pdf_text,
        encoding="utf-8",
    )

    section_path = fund_dir / "top_holdings_section.txt"

    section_path.write_text(
        section_text,
        encoding="utf-8",
    )

    holdings_path = fund_dir / "top_holdings.json"

    write_json(
        holdings_path,
        {
            "status": result["status"],
            "excelRow": result["excelRow"],
            "fundName": result.get("fundName"),
            "prudentialUrl": result["prudentialUrl"],
            "excelPruAccessName": result[
                "excelPruAccessName"
            ],
            "factsheetUrl": result["factsheetUrl"],
            "factsheetDataAsAt": result[
                "factsheetDataAsAt"
            ],
            "factsheetDocumentDate": result[
                "factsheetDocumentDate"
            ],
            "parser": result["parser"],
            "holdings": result["holdings"],
        },
    )

    metadata_path = fund_dir / "metadata.json"

    write_json(
        metadata_path,
        {
            "excelRow": result["excelRow"],
            "fundName": result.get("fundName"),
            "prudentialUrl": result["prudentialUrl"],
            "excelPruAccessName": result[
                "excelPruAccessName"
            ],
            "factsheetUrl": result["factsheetUrl"],
            "factsheetDataAsAt": result[
                "factsheetDataAsAt"
            ],
            "factsheetDocumentDate": result[
                "factsheetDocumentDate"
            ],
            "parser": result["parser"],
            "holdingsCount": len(
                result["holdings"]
            ),
            "savedAtUtc": utc_now_iso(),
        },
    )

    return fund_dir


def save_failure(
    fund: dict,
    error: Exception | str,
    attempt: int | None = None,
) -> Path:
    excel_row = fund["excelRow"]

    failure_dir = (
        OUTPUT_DIR
        / f"{excel_row}_failed"
    )

    failure_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    failure = {
        "status": "failed",
        "excelRow": excel_row,
        "prudentialUrl": fund[
            "prudentialUrl"
        ],
        "excelPruAccessName": fund[
            "pruAccessName"
        ],
        "error": str(error),
        "failedAtUtc": utc_now_iso(),
    }

    if attempt is not None:
        failure["attempt"] = attempt

    failure_path = failure_dir / "failure.json"

    write_json(
        failure_path,
        failure,
    )

    return failure_path


# ============================================================================
# CONSOLE OUTPUT
# ============================================================================

def print_holdings(result: dict) -> None:
    print("")
    print("SUCCESS")
    print(
        f"Fund: {result.get('fundName') or '-'}"
    )
    print(
        "Factsheet data as at: "
        f"{result.get('factsheetDataAsAt') or '-'}"
    )
    print(
        "Factsheet document date: "
        f"{result.get('factsheetDocumentDate') or '-'}"
    )

    print(
        "Top holdings extracted: "
        f"{len(result['holdings'])}"
    )

    for holding in result["holdings"]:
        print(
            f"  {holding['rank']}. "
            f"{holding['name']} - "
            f"{holding['weightPercent']}%"
        )


# ============================================================================
# MAIN
# ============================================================================

def main() -> int:
    print("")
    print("=" * 70)
    print("VGrat FMS - PRUDENTIAL TOP HOLDINGS ALL-FUND TEST EXTRACTOR")
    print("=" * 70)
    print("")

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        funds = read_excel_funds()
    except Exception as exc:
        print(
            f"FATAL: Failed reading Excel master: {exc}"
        )
        return 1

    print(
        f"Excel fund universe: {len(funds)}"
    )

    if not funds:
        print(
            "FATAL: No populated Prudential URLs found in Column A."
        )
        return 1

    successful_results = []
    failed_results = []

    total = len(funds)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=HEADLESS
        )

        try:
            for position, fund in enumerate(
                funds,
                start=1,
            ):
                print("")
                print(
                    f"[{position}/{total}] "
                    f"Excel row {fund['excelRow']}"
                )

                success = False
                last_error = None

                for attempt in range(
                    1,
                    MAX_RETRIES + 1,
                ):
                    try:
                        extracted = extract_single_fund(
                            browser,
                            fund,
                        )

                        result = extracted["result"]

                        # --------------------------------------------------
                        # Re-validate before saving.
                        # --------------------------------------------------

                        validate_holdings(
                            result["holdings"]
                        )

                        # --------------------------------------------------
                        # Save only after successful validation.
                        # --------------------------------------------------

                        output_dir = save_success_result(
                            result=result,
                            pdf_bytes=extracted[
                                "pdfBytes"
                            ],
                            pdf_text=extracted[
                                "pdfText"
                            ],
                            section_text=extracted[
                                "sectionText"
                            ],
                        )

                        print(
                            f"Saved: {output_dir}"
                        )

                        print_holdings(
                            result
                        )

                        successful_results.append(
                            result
                        )

                        success = True
                        break

                    except LookupError as exc:
                        # no_holdings_section is not a parser guess.
                        # It is a legitimate status indicating that the
                        # official factsheet does not expose the requested
                        # section.
                        if str(exc) == "no_holdings_section":
                            result = {
                                "status": "no_holdings_section",
                                "excelRow": fund[
                                    "excelRow"
                                ],
                                "prudentialUrl": fund[
                                    "prudentialUrl"
                                ],
                                "excelPruAccessName": fund[
                                    "pruAccessName"
                                ],
                                "error": (
                                    "Official Prudential factsheet "
                                    "does not contain a Top 10 Holdings "
                                    "section."
                                ),
                            }

                            successful_results.append(
                                result
                            )

                            print("")
                            print(
                                "NO HOLDINGS SECTION"
                            )

                            success = True
                            break

                        last_error = exc

                        print(
                            f"Attempt {attempt}/"
                            f"{MAX_RETRIES} failed: "
                            f"{exc}"
                        )

                    except Exception as exc:
                        last_error = exc

                        print(
                            f"Attempt {attempt}/"
                            f"{MAX_RETRIES} failed: "
                            f"{exc}"
                        )

                    if attempt < MAX_RETRIES:
                        time.sleep(
                            RETRY_DELAY_SECONDS
                        )

                if not success:
                    print("")
                    print("FAILED")

                    failure_path = save_failure(
                        fund,
                        last_error
                        if last_error is not None
                        else "Unknown extraction failure.",
                    )

                    print(
                        f"Failure saved: {failure_path}"
                    )

                    failed_results.append(
                        {
                            "status": "failed",
                            "excelRow": fund[
                                "excelRow"
                            ],
                            "prudentialUrl": fund[
                                "prudentialUrl"
                            ],
                            "excelPruAccessName": fund[
                                "pruAccessName"
                            ],
                            "error": str(
                                last_error
                                if last_error is not None
                                else "Unknown extraction failure."
                            ),
                        }
                    )

        finally:
            browser.close()

    # =========================================================================
    # ALL HOLDINGS OUTPUT
    # =========================================================================

    generated_at = utc_now_iso()

    all_holdings_payload = {
        "status": (
            "success"
            if not failed_results
            else "failed"
        ),
        "generatedAtUtc": generated_at,
        "excelFundUniverse": len(funds),
        "successfulFunds": len(
            successful_results
        ),
        "failedFunds": len(
            failed_results
        ),
        "funds": successful_results,
        "failures": failed_results,
        "rules": {
            "masterSource": "Funds Links.xlsm",
            "masterSourceColumnA": (
                "Prudential fund URL"
            ),
            "masterSourceColumnB": (
                "Exact PruAccess fund name"
            ),
            "processEveryPopulatedColumnARow": True,
            "officialPrudentialSourcesOnly": True,
            "thirdPartyHoldingsForbidden": True,
            "fabricatedHoldingsForbidden": True,
            "fabricatedWeightsForbidden": True,
            "maximumHoldings": MAX_HOLDINGS,
            "forceTenHoldings": False,
            "preservePublishedOrder": True,
            "allowDuplicateNames": True,
            "allowDuplicatePercentages": True,
            "fixedIncomeCouponPercentagesAllowedInNames": True,
            "portfolioWeightUsesFinalRelevantPercentage": True,
            "multiLineHoldingsSupported": True,
            "multipleHoldingsPerPdfLineSupported": True,
            "maturityDateContinuationSupported": True,
            "fallbackParserRebuilt": True,
        },
    }

    write_json(
        ALL_HOLDINGS_FILE,
        all_holdings_payload,
    )

    # =========================================================================
    # RUN SUMMARY
    # =========================================================================

    run_summary = {
        "status": (
            "success"
            if not failed_results
            else "failed"
        ),
        "generatedAtUtc": generated_at,
        "excelFundUniverse": len(funds),
        "successfulFunds": len(
            successful_results
        ),
        "failedFunds": [
            failure["excelRow"]
            for failure in failed_results
        ],
        "successfulFundRows": [
            result["excelRow"]
            for result in successful_results
        ],
        "totalHoldingsExtracted": sum(
            len(result.get("holdings", []))
            for result in successful_results
        ),
        "rules": {
            "automaticUniverseFromExcelColumnA": True,
            "noHardcodedFundCount": True,
            "officialPrudentialFactsheetOnly": True,
            "noThirdPartyHoldings": True,
            "noSyntheticHoldings": True,
            "noSyntheticWeights": True,
            "preservePublishedOrder": True,
            "preserveDuplicateHoldings": True,
            "preserveDuplicatePercentages": True,
            "doNotForceTenHoldings": True,
            "fixedIncomeCouponHandling": True,
            "multipleHoldingsOnSamePdfLine": True,
            "multiLineSecurityNames": True,
            "maturityDateContinuation": True,
            "fallbackUsesLastRelevantPercentage": True,
        },
    }

    write_json(
        RUN_SUMMARY_FILE,
        run_summary,
    )

    # =========================================================================
    # FINAL CONSOLE SUMMARY
    # =========================================================================

    print("")
    print("=" * 70)
    print("RUN COMPLETE")
    print("=" * 70)
    print("")
    print(
        f"Excel fund universe : {len(funds)}"
    )
    print(
        f"Successful funds    : {len(successful_results)}"
    )
    print(
        f"Failed funds        : {len(failed_results)}"
    )
    print(
        "Total holdings      : "
        f"{run_summary['totalHoldingsExtracted']}"
    )
    print("")
    print(
        f"All holdings file   : {ALL_HOLDINGS_FILE}"
    )
    print(
        f"Run summary file    : {RUN_SUMMARY_FILE}"
    )
    print("")

    if failed_results:
        print("FAILED ROWS:")

        for failure in failed_results:
            print(
                f"  Row {failure['excelRow']}: "
                f"{failure['error']}"
            )

        print("")

        return 1

    print("ALL FUNDS PASSED")
    print("")

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )
