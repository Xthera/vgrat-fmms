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
    PDF text extraction
          |
          v
    Confirmed Top Holdings section
          |
          v
    Primary holdings parser
          |
          v
    Fallback holdings parser
          |
          v
    Validation
          |
          v
    output_holdings/


IMPORTANT RULES
===============

1. Excel Column A controls the fund universe.

2. Every populated URL in Column A is processed.

3. No hardcoded 67-fund limit.

4. Duplicate URLs are preserved and processed as separate Excel rows.

5. Only official Prudential Singapore pages are accepted.

6. Only official Prudential Singapore factsheets are accepted.

7. No third-party holdings sources.

8. No inferred holdings.

9. No fabricated holdings.

10. No fabricated percentages.

11. No forced ten holdings.

12. If Prudential publishes fewer than ten holdings, exactly that
    published number is retained.

13. Prudential published order is preserved.

14. Wrapped PDF lines belonging to one security are reconstructed.

15. Fixed-income coupon/rate percentages are not automatically treated
    as portfolio weights.

16. For fixed-income logical lines containing multiple percentages,
    the final percentage is treated as the portfolio weight.

17. Primary parser uses the first percentage when the logical line
    contains exactly one percentage.

18. If the primary parser encounters multiple percentages, the
    fallback parser is used.

19. Fallback parser uses the final percentage on the logical holding
    line.

20. Confirmed holdings headings include:

        Top 10 Holdings
        Top Ten Holdings
        Top 10 Holdings3
        Top Ten Holdings3

    where the trailing digits are treated as PDF footnote markers.

21. Generic words such as "holdings" or "investments" are NOT accepted
    as holdings section headings.

22. Security names containing the word HOLDINGS are not headings.

23. No candidate diagnostic heading is automatically accepted merely
    because it contains the word "holdings" or "investments".

24. If no confirmed Top Holdings heading is found, the fund is recorded
    as no_holdings_section.

25. Failed funds do not receive fabricated blank holdings.

26. The script does not modify PruAccess data.

27. The script does not modify test_pruaccess.py.

28. The script does not modify data.json, index.html, CSS, or JS.

29. This is a TEST collector only.

30. Whenever parser rules are changed, the complete script should be
    replaced so the full implementation remains reproducible.
"""

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

from playwright.sync_api import sync_playwright

# IMPORTANT:
# PlaywrightTimeoutError is not exported from playwright.sync_api in
# the installed Playwright version used by GitHub Actions.
from playwright._impl._errors import TimeoutError as PlaywrightTimeoutError


# ============================================================================
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_holdings")
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"
FAILED_OUTPUT_DIR = OUTPUT_DIR / "failed"

ALL_HOLDINGS_FILE = OUTPUT_DIR / "all_holdings.json"
RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"

HEADLESS = True

PAGE_TIMEOUT_MS = 120_000
FACTSHEET_TIMEOUT_MS = 120_000
POST_PAGE_WAIT_MS = 1_500

MAX_RETRIES = 3
RETRY_DELAY_SECONDS = 3

MAX_HOLDINGS = 10

OFFICIAL_HOSTS = {
    "prudential.com.sg",
    "www.prudential.com.sg",
}


# ============================================================================
# TEXT / GENERAL HELPERS
# ============================================================================

def clean_text(value):
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\x00", " ")
    text = text.replace("\r", "\n")
    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\u200c", "")
    text = text.replace("\u200d", "")
    text = text.replace("\ufeff", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def normalize_text(value):
    text = clean_text(value)

    replacements = {
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2013": "-",
        "\u2014": "-",
        "\u2212": "-",
        "\u00b7": " ",
        "\u2022": " ",
        "\u00a0": " ",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"\s+", " ", text)

    return text.strip().lower()


def utc_now_iso():
    return datetime.now(timezone.utc).isoformat()


def save_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", encoding="utf-8") as handle:
        json.dump(
            data,
            handle,
            indent=2,
            ensure_ascii=False,
        )


def safe_filename(value, fallback="item"):
    text = clean_text(value)

    if not text:
        text = fallback

    text = re.sub(r"[^\w.\-]+", "_", text, flags=re.UNICODE)
    text = re.sub(r"_+", "_", text)
    text = text.strip("._")

    if not text:
        text = fallback

    return text[:180]


def is_prudential_url(url):
    if not url:
        return False

    try:
        parsed = urlparse(url)
        hostname = (parsed.hostname or "").lower()

        return hostname in OFFICIAL_HOSTS
    except Exception:
        return False


def ensure_prudential_url(url):
    if not url:
        raise ValueError("Empty Prudential URL")

    if not is_prudential_url(url):
        raise ValueError(
            f"URL is not an official Prudential Singapore URL: {url}"
        )

    return url


# ============================================================================
# EXCEL UNIVERSE
# ============================================================================

def read_excel_funds():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Required Excel file not found: {EXCEL_FILE}"
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
            pruaccess_name = worksheet.cell(row=row_number, column=2).value

            if url_value is None:
                continue

            prudential_url = clean_text(url_value)

            if not prudential_url:
                continue

            funds.append(
                {
                    "excelRow": row_number,
                    "prudentialUrl": prudential_url,
                    "pruAccessName": clean_text(pruaccess_name),
                }
            )

        return funds

    finally:
        workbook.close()


# ============================================================================
# FACTSHEET DISCOVERY
# ============================================================================

def find_factsheet_url(page):
    anchors = page.locator("a")

    try:
        count = anchors.count()
    except Exception:
        count = 0

    candidates = []

    for index in range(count):
        try:
            anchor = anchors.nth(index)

            href = anchor.get_attribute("href")
            text = clean_text(anchor.inner_text())

            if not href:
                continue

            absolute_url = urljoin(page.url, href)

            if not is_prudential_url(absolute_url):
                continue

            normalized_href = normalize_text(absolute_url)
            normalized_text = normalize_text(text)

            score = 0

            if "fund factsheet" in normalized_text:
                score += 100

            if "fund factsheet" in normalized_href:
                score += 100

            if "factsheet" in normalized_text:
                score += 50

            if "factsheet" in normalized_href:
                score += 50

            if normalized_href.endswith(".pdf"):
                score += 40

            if ".pdf?" in normalized_href:
                score += 40

            if "download" in normalized_text:
                score += 10

            if score <= 0:
                continue

            candidates.append(
                {
                    "url": absolute_url,
                    "text": text,
                    "score": score,
                }
            )

        except Exception:
            continue

    if not candidates:
        raise RuntimeError(
            "Could not locate an official Prudential factsheet link"
        )

    candidates.sort(
        key=lambda item: (
            item["score"],
            len(item["url"]),
        ),
        reverse=True,
    )

    return candidates[0]["url"]


def download_factsheet(page, factsheet_url):
    response = page.request.get(
        factsheet_url,
        timeout=FACTSHEET_TIMEOUT_MS,
    )

    if response.status != 200:
        raise RuntimeError(
            f"Factsheet HTTP status {response.status}: {factsheet_url}"
        )

    body = response.body()

    if not body:
        raise RuntimeError(
            f"Factsheet response was empty: {factsheet_url}"
        )

    if not body.startswith(b"%PDF"):
        content_type = response.headers.get("content-type", "")

        raise RuntimeError(
            "Factsheet response was not a PDF: "
            f"content-type={content_type!r}, "
            f"url={factsheet_url}"
        )

    return body


# ============================================================================
# PDF EXTRACTION
# ============================================================================

def extract_pdf_text(pdf_bytes):
    reader = PdfReader(BytesIO(pdf_bytes))

    if not reader.pages:
        raise RuntimeError("PDF contains no pages")

    page_texts = []

    for page_number, pdf_page in enumerate(reader.pages, start=1):
        try:
            page_text = pdf_page.extract_text()
        except Exception as exc:
            raise RuntimeError(
                f"PDF text extraction failed on page {page_number}: {exc}"
            ) from exc

        if page_text:
            page_texts.append(page_text)

    full_text = "\n".join(page_texts)

    if not clean_text(full_text):
        raise RuntimeError("PDF contains no extractable text")

    return full_text


def pdf_lines(text):
    lines = []

    for raw_line in text.splitlines():
        line = clean_text(raw_line)

        if line:
            lines.append(line)

    return lines


# ============================================================================
# PDF METADATA EXTRACTION
# ============================================================================

def extract_data_as_at(text):
    patterns = [
        r"data\s+as\s+at\s*[:\-]?\s*([^\n]+)",
        r"data\s+as\s+of\s*[:\-]?\s*([^\n]+)",
        r"data\s+as\s+at\s+([^\n]+)",
        r"as\s+at\s*[:\-]?\s*([0-9]{1,2}[\-/][A-Za-z0-9]{3,9}[\-/][0-9]{2,4})",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match:
            return clean_text(match.group(1))

    return ""


def extract_document_date(text):
    patterns = [
        r"document\s+date\s*[:\-]?\s*([^\n]+)",
        r"date\s+of\s+document\s*[:\-]?\s*([^\n]+)",
        r"dated\s*[:\-]?\s*([^\n]+)",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match:
            return clean_text(match.group(1))

    return ""


# ============================================================================
# CONFIRMED TOP HOLDINGS HEADING
# ============================================================================

# We deliberately do NOT accept a generic word such as:
#
#     Holdings
#     Portfolio Holdings
#     Investments
#
# because the diagnostics showed that these frequently occur in ordinary
# fund/security names and investment-related metadata.
#
# The confirmed heading is specifically:
#
#     Top 10 Holdings
#     Top Ten Holdings
#
# optionally followed by one or more PDF footnote digits.
#
# Examples:
#
#     Top 10 Holdings
#     Top 10 Holdings3
#     Top Ten Holdings3
#
# Unicode superscript digits are also accepted because PDF text extraction
# can represent footnotes differently depending on the source PDF.

FOOTNOTE_DIGITS = "0123456789⁰¹²³⁴⁵⁶⁷⁸⁹"

CONFIRMED_HOLDINGS_HEADING_RE = re.compile(
    rf"^top\s+(?:10|ten)\s+holdings\s*[{re.escape(FOOTNOTE_DIGITS)}]*$",
    flags=re.IGNORECASE,
)


def normalize_footnote_digits(text):
    replacements = {
        "⁰": "0",
        "¹": "1",
        "²": "2",
        "³": "3",
        "⁴": "4",
        "⁵": "5",
        "⁶": "6",
        "⁷": "7",
        "⁸": "8",
        "⁹": "9",
    }

    result = text

    for old, new in replacements.items():
        result = result.replace(old, new)

    return result


def is_confirmed_holdings_heading(line):
    normalized = normalize_text(line)
    normalized = normalize_footnote_digits(normalized)

    return bool(
        CONFIRMED_HOLDINGS_HEADING_RE.fullmatch(normalized)
    )


def find_holdings_start(lines):
    for index, line in enumerate(lines):
        if is_confirmed_holdings_heading(line):
            return index

    return None


# ============================================================================
# HOLDINGS SECTION END DETECTION
# ============================================================================

def is_holdings_end(line):
    normalized = normalize_text(line)

    exact_markers = {
        "source",
        "source:",
        "inception date",
        "important information",
        "disclaimer",
        "past performance",
        "portfolio characteristics",
        "asset allocation",
    }

    if normalized in exact_markers:
        return True

    if normalized.startswith("source:"):
        return True

    if normalized.startswith("important information"):
        return True

    if normalized.startswith("disclaimer"):
        return True

    if normalized.startswith("past performance"):
        return True

    if re.fullmatch(
        r"page\s+\d+(?:\s+of\s+\d+)?",
        normalized,
    ):
        return True

    return False


def extract_holdings_section(text):
    lines = pdf_lines(text)

    start_index = find_holdings_start(lines)

    if start_index is None:
        return "", "not_published"

    section_lines = []

    for line in lines[start_index + 1:]:
        if is_holdings_end(line):
            break

        section_lines.append(line)

    section = "\n".join(section_lines).strip()

    if not section:
        return "", "empty"

    return section, "confirmed"


# ============================================================================
# HOLDING NAME / PERCENTAGE HELPERS
# ============================================================================

PERCENTAGE_RE = re.compile(
    r"(?<![\d.])"
    r"([0-9]+(?:\.[0-9]+)?)"
    r"\s*%"
)


MONTH_PATTERN = (
    r"(?:"
    r"JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC"
    r")"
)

MATURITY_PATTERNS = [
    re.compile(
        rf"\b\d{{1,2}}[-/\s]{MONTH_PATTERN}[-/\s]\d{{2,4}}\b",
        flags=re.IGNORECASE,
    ),
    re.compile(
        rf"\b\d{{1,2}}[-/\s]{MONTH_PATTERN}\d{{2,4}}\b",
        flags=re.IGNORECASE,
    ),
    re.compile(
        r"\b\d{1,2}/\d{1,2}/\d{2,4}\b",
        flags=re.IGNORECASE,
    ),
    re.compile(
        r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b",
        flags=re.IGNORECASE,
    ),
]


def parse_percentage(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None

    if number < 0 or number > 100:
        return None

    return number


def percentage_matches(line):
    matches = list(PERCENTAGE_RE.finditer(line))

    values = []

    for match in matches:
        value = parse_percentage(match.group(1))

        if value is not None:
            values.append(
                {
                    "match": match,
                    "value": value,
                }
            )

    return values


def find_percentage_in_line(line):
    matches = percentage_matches(line)

    if not matches:
        return None

    return matches[0]


def find_last_percentage_in_line(line):
    matches = percentage_matches(line)

    if not matches:
        return None

    return matches[-1]


def contains_maturity_date(line):
    for pattern in MATURITY_PATTERNS:
        if pattern.search(line):
            return True

    return False


def line_ends_with_percentage(line):
    return bool(
        re.search(
            r"[0-9]+(?:\.[0-9]+)?\s*%\s*$",
            line,
        )
    )


def line_has_percentage(line):
    return bool(PERCENTAGE_RE.search(line))


# ============================================================================
# LOGICAL HOLDING LINE RECONSTRUCTION
# ============================================================================

def build_logical_holding_lines(section_text):
    """
    Reconstruct wrapped fixed-income holdings.

    Example PDF extraction:

        CORPORACION ANDINA DE FOMENTO 7.7%
        6-MAR2029 1.7%

    becomes:

        CORPORACION ANDINA DE FOMENTO 7.7% 6-MAR2029 1.7%

    Another common extraction:

        CORPORACION ANDINA DE FOMENTO 7.7%
        6-MAR2029
        1.7%

    becomes:

        CORPORACION ANDINA DE FOMENTO 7.7% 6-MAR2029 1.7%

    The purpose is to ensure the final portfolio weight is available
    to the fallback parser without treating the coupon/rate as the
    portfolio weight.
    """

    lines = pdf_lines(section_text)

    result = []

    index = 0

    while index < len(lines):
        current = lines[index]

        if (
            line_ends_with_percentage(current)
            and index + 1 < len(lines)
            and contains_maturity_date(lines[index + 1])
            and line_has_percentage(lines[index + 1])
        ):
            result.append(
                clean_text(
                    current + " " + lines[index + 1]
                )
            )

            index += 2
            continue

        if (
            line_ends_with_percentage(current)
            and index + 2 < len(lines)
            and contains_maturity_date(lines[index + 1])
            and line_has_percentage(lines[index + 2])
        ):
            result.append(
                clean_text(
                    current
                    + " "
                    + lines[index + 1]
                    + " "
                    + lines[index + 2]
                )
            )

            index += 3
            continue

        result.append(current)
        index += 1

    return result


# ============================================================================
# RANK DETECTION
# ============================================================================

RANK_ONLY_RE = re.compile(
    r"^\s*(\d{1,2})\s*[\.\)\-:]?\s*$"
)

RANK_PREFIX_RE = re.compile(
    r"^\s*(\d{1,2})\s*[\.\)\-:]\s+(.+)$"
)

RANK_SPACE_PREFIX_RE = re.compile(
    r"^\s*(\d{1,2})\s+(.+)$"
)


def extract_leading_rank(line):
    match = RANK_ONLY_RE.match(line)

    if match:
        return int(match.group(1)), ""

    match = RANK_PREFIX_RE.match(line)

    if match:
        return int(match.group(1)), clean_text(match.group(2))

    match = RANK_SPACE_PREFIX_RE.match(line)

    if match:
        number = int(match.group(1))

        if 1 <= number <= 99:
            return number, clean_text(match.group(2))

    return None, line


# ============================================================================
# HOLDINGS NOISE
# ============================================================================

def is_holding_header_or_noise(line):
    normalized = normalize_text(line)

    if not normalized:
        return True

    exact_noise = {
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

    if normalized in exact_noise:
        return True

    if is_confirmed_holdings_heading(line):
        return True

    if normalized.startswith("top 10 holdings"):
        return True

    if normalized.startswith("top ten holdings"):
        return True

    return False


def clean_holding_fragment(fragment):
    text = clean_text(fragment)

    text = re.sub(
        r"^[•·▪◦●○■□]+\s*",
        "",
        text,
    )

    text = re.sub(
        r"^[|]+\s*",
        "",
        text,
    )

    text = text.replace("|", " ")

    text = re.sub(
        r"\bNone\b",
        "",
        text,
        flags=re.IGNORECASE,
    )

    text = clean_text(text)

    return text


def clean_holding_name(name):
    text = clean_holding_fragment(name)

    rank, remainder = extract_leading_rank(text)

    if rank is not None:
        text = remainder

    text = clean_text(text)

    return text


def combine_holding_name_fragments(fragments):
    cleaned = []

    for fragment in fragments:
        fragment = clean_holding_fragment(fragment)

        if not fragment:
            continue

        rank, remainder = extract_leading_rank(fragment)

        if rank is not None:
            fragment = remainder

        fragment = clean_holding_fragment(fragment)

        if fragment:
            cleaned.append(fragment)

    return clean_holding_name(" ".join(cleaned))


# ============================================================================
# HOLDINGS VALIDATION
# ============================================================================

def validate_holdings(holdings):
    if not holdings:
        raise ValueError(
            "No holdings were extracted"
        )

    if len(holdings) > MAX_HOLDINGS:
        raise ValueError(
            f"More than {MAX_HOLDINGS} holdings extracted"
        )

    expected_ranks = list(range(1, len(holdings) + 1))
    actual_ranks = [
        holding["rank"]
        for holding in holdings
    ]

    if actual_ranks != expected_ranks:
        raise ValueError(
            "Holding ranks are not sequential: "
            f"expected={expected_ranks}, "
            f"actual={actual_ranks}"
        )

    for holding in holdings:
        name = clean_text(holding.get("name"))
        weight = holding.get("weight")

        if not name:
            raise ValueError(
                f"Holding {holding['rank']} has no name"
            )

        if not isinstance(weight, (int, float)):
            raise ValueError(
                f"Holding {holding['rank']} has invalid "
                f"portfolio weight: {weight!r}"
            )

        if weight < 0 or weight > 100:
            raise ValueError(
                f"Holding {holding['rank']} has invalid "
                f"portfolio weight: {weight}%"
            )

    return True


# ============================================================================
# PRIMARY HOLDINGS PARSER
# ============================================================================

def parse_holdings(section_text):
    """
    Primary parser.

    A logical holding line with exactly one percentage is straightforward.

    If a line contains multiple percentages, the primary parser deliberately
    does not guess which one is the portfolio weight. It raises and allows
    the fallback parser to retry the same official factsheet section.
    """

    logical_lines = build_logical_holding_lines(section_text)

    holdings = []

    pending_fragments = []
    pending_rank = None

    for line in logical_lines:
        line = clean_text(line)

        if not line:
            continue

        if is_holding_header_or_noise(line):
            continue

        rank, rank_text = extract_leading_rank(line)

        if rank is not None:
            if rank_text:
                line = rank_text
            else:
                if pending_fragments:
                    raise ValueError(
                        "Encountered a new holding rank while "
                        "previous holding was incomplete"
                    )

                pending_rank = rank
                continue

        matches = percentage_matches(line)

        if len(matches) > 1:
            raise ValueError(
                "Primary parser encountered multiple percentages "
                f"in holding line: {line}"
            )

        if len(matches) == 1:
            match_info = matches[0]
            match = match_info["match"]
            weight = match_info["value"]

            name_part = line[:match.start()]

            fragments = list(pending_fragments)

            if name_part.strip():
                fragments.append(name_part)

            name = combine_holding_name_fragments(fragments)

            if not name:
                raise ValueError(
                    "Holding percentage found without a holding name: "
                    f"{line}"
                )

            if pending_rank is not None:
                holding_rank = pending_rank
            else:
                holding_rank = len(holdings) + 1

            holdings.append(
                {
                    "rank": holding_rank,
                    "name": name,
                    "weight": weight,
                    "weightText": match.group(0),
                    "rawLine": line,
                    "parser": "primary",
                }
            )

            pending_fragments = []
            pending_rank = None

            if len(holdings) >= MAX_HOLDINGS:
                break

            continue

        pending_fragments.append(line)

    validate_holdings(holdings)

    return holdings


# ============================================================================
# FALLBACK HOLDINGS PARSER
# ============================================================================

def parse_holdings_fallback(section_text):
    """
    Fallback parser.

    Uses the LAST percentage in each logical holding line.

    This is important for fixed-income securities such as:

        USA DL-Treasury Bills 2026(26) 0.00% 5.8%

    where:

        0.00%

    is part of the security description while:

        5.8%

    is the portfolio weight.

    It also handles reconstructed fixed-income lines such as:

        CORPORACION ANDINA DE FOMENTO 7.7% 6-MAR2029 1.7%

    where 7.7% is the coupon and 1.7% is the portfolio weight.
    """

    logical_lines = build_logical_holding_lines(section_text)

    holdings = []

    pending_fragments = []
    pending_rank = None

    for line in logical_lines:
        line = clean_text(line)

        if not line:
            continue

        if is_holding_header_or_noise(line):
            continue

        rank, rank_text = extract_leading_rank(line)

        if rank is not None:
            if rank_text:
                line = rank_text
            else:
                if pending_fragments:
                    raise ValueError(
                        "Encountered a new holding rank while "
                        "previous holding was incomplete"
                    )

                pending_rank = rank
                continue

        match_info = find_last_percentage_in_line(line)

        if match_info is not None:
            match = match_info["match"]
            weight = match_info["value"]

            name_part = line[:match.start()]

            fragments = list(pending_fragments)

            if name_part.strip():
                fragments.append(name_part)

            name = combine_holding_name_fragments(fragments)

            if not name:
                raise ValueError(
                    "Fallback parser found a percentage without "
                    f"a holding name: {line}"
                )

            if pending_rank is not None:
                holding_rank = pending_rank
            else:
                holding_rank = len(holdings) + 1

            holdings.append(
                {
                    "rank": holding_rank,
                    "name": name,
                    "weight": weight,
                    "weightText": match.group(0),
                    "rawLine": line,
                    "parser": "fallback",
                }
            )

            pending_fragments = []
            pending_rank = None

            if len(holdings) >= MAX_HOLDINGS:
                break

            continue

        pending_fragments.append(line)

    validate_holdings(holdings)

    return holdings


# ============================================================================
# FUND EXTRACTION
# ============================================================================

def extract_single_fund(page, fund):
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
            timeout=PAGE_TIMEOUT_MS,
        )
    except PlaywrightTimeoutError:
        pass

    page.wait_for_timeout(POST_PAGE_WAIT_MS)

    final_url = page.url

    if not is_prudential_url(final_url):
        raise RuntimeError(
            "Prudential URL redirected to a non-Prudential host: "
            f"{final_url}"
        )

    fund_name = ""

    try:
        h1 = page.locator("h1").first

        if h1.count() > 0:
            fund_name = clean_text(h1.inner_text())
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
                flags=re.IGNORECASE,
            )

            if match:
                fund_name = clean_text(match.group(0))
        except Exception:
            pass

    if not fund_name:
        fund_name = fund["pruAccessName"] or (
            f"Excel Row {excel_row}"
        )

    factsheet_url = find_factsheet_url(page)

    pdf_bytes = download_factsheet(
        page,
        factsheet_url,
    )

    pdf_text = extract_pdf_text(pdf_bytes)

    data_as_at = extract_data_as_at(pdf_text)
    document_date = extract_document_date(pdf_text)

    holdings_section, section_status = extract_holdings_section(
        pdf_text
    )

    if section_status != "confirmed":
        return {
            "status": "no_holdings_section",
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "finalUrl": final_url,
            "pruAccessName": fund["pruAccessName"],
            "fundName": fund_name,
            "factsheetUrl": factsheet_url,
            "dataAsAt": data_as_at,
            "documentDate": document_date,
            "holdingsSectionStatus": section_status,
            "topHoldings": [],
            "topHoldingsCount": 0,
            "parser": None,
            "rules": {
                "confirmedTopHoldingsHeading": True,
                "genericHoldingsHeadingRejected": True,
                "footnoteHeadingAccepted": True,
            },
        }

    parser_used = None
    holdings = None
    primary_error = None
    fallback_error = None

    try:
        holdings = parse_holdings(
            holdings_section
        )

        parser_used = "primary"

    except Exception as exc:
        primary_error = str(exc)

        try:
            holdings = parse_holdings_fallback(
                holdings_section
            )

            parser_used = "fallback"

        except Exception as fallback_exc:
            fallback_error = str(fallback_exc)

            raise RuntimeError(
                "Both holdings parsers failed. "
                f"Primary: {primary_error}; "
                f"Fallback: {fallback_error}"
            )

    return {
        "status": "success",
        "excelRow": excel_row,
        "prudentialUrl": prudential_url,
        "finalUrl": final_url,
        "pruAccessName": fund["pruAccessName"],
        "fundName": fund_name,
        "factsheetUrl": factsheet_url,
        "dataAsAt": data_as_at,
        "documentDate": document_date,
        "holdingsSectionStatus": section_status,
        "topHoldings": holdings,
        "topHoldingsCount": len(holdings),
        "parser": parser_used,
        "parserErrors": {
            "primary": primary_error,
            "fallback": fallback_error,
        },
        "rules": {
            "confirmedTopHoldingsHeading": True,
            "genericHoldingsHeadingRejected": True,
            "footnoteHeadingAccepted": True,
            "top10HoldingsFootnoteDigitsAccepted": True,
            "topTenHoldingsFootnoteDigitsAccepted": True,
            "holdingsNamesInferred": False,
            "fabricatedData": False,
            "syntheticData": False,
        },
    }


# ============================================================================
# SAVE SUCCESS
# ============================================================================

def save_success_result(
    fund,
    result,
    pdf_bytes,
    pdf_text,
    holdings_section,
):
    excel_row = fund["excelRow"]

    identifier = (
        result.get("fundName")
        or fund.get("pruAccessName")
        or f"row_{excel_row}"
    )

    directory_name = (
        f"{excel_row}_"
        f"{safe_filename(identifier, f'row_{excel_row}')}"
    )

    output_dir = FUNDS_OUTPUT_DIR / directory_name

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    pdf_path = output_dir / "factsheet.pdf"
    text_path = output_dir / "factsheet_text.txt"
    section_path = output_dir / "top_holdings_section.txt"
    holdings_path = output_dir / "top_holdings.json"
    metadata_path = output_dir / "metadata.json"

    pdf_path.write_bytes(pdf_bytes)

    text_path.write_text(
        pdf_text,
        encoding="utf-8",
    )

    section_path.write_text(
        holdings_section,
        encoding="utf-8",
    )

    save_json(
        holdings_path,
        result["topHoldings"],
    )

    metadata = {
        "generatedAtUtc": utc_now_iso(),
        "excelRow": fund["excelRow"],
        "prudentialUrl": fund["prudentialUrl"],
        "pruAccessName": fund["pruAccessName"],
        "fundName": result.get("fundName"),
        "finalUrl": result.get("finalUrl"),
        "factsheetUrl": result.get("factsheetUrl"),
        "dataAsAt": result.get("dataAsAt"),
        "documentDate": result.get("documentDate"),
        "holdingsSectionStatus": result.get(
            "holdingsSectionStatus"
        ),
        "topHoldingsCount": result.get(
            "topHoldingsCount"
        ),
        "parser": result.get("parser"),
        "files": {
            "factsheetPdf": str(pdf_path),
            "factsheetText": str(text_path),
            "topHoldingsSection": str(section_path),
            "topHoldingsJson": str(holdings_path),
        },
    }

    save_json(
        metadata_path,
        metadata,
    )

    return output_dir


# ============================================================================
# SAVE NO-HOLDINGS RESULT
# ============================================================================

def save_no_holdings_result(
    fund,
    result,
    pdf_bytes,
    pdf_text,
):
    excel_row = fund["excelRow"]

    identifier = (
        result.get("fundName")
        or fund.get("pruAccessName")
        or f"row_{excel_row}"
    )

    directory_name = (
        f"{excel_row}_"
        f"{safe_filename(identifier, f'row_{excel_row}')}"
    )

    output_dir = FUNDS_OUTPUT_DIR / directory_name

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    pdf_path = output_dir / "factsheet.pdf"
    text_path = output_dir / "factsheet_text.txt"
    metadata_path = output_dir / "metadata.json"
    section_path = output_dir / "top_holdings_section.txt"

    pdf_path.write_bytes(pdf_bytes)

    text_path.write_text(
        pdf_text,
        encoding="utf-8",
    )

    section_path.write_text(
        "",
        encoding="utf-8",
    )

    metadata = {
        "generatedAtUtc": utc_now_iso(),
        "excelRow": fund["excelRow"],
        "prudentialUrl": fund["prudentialUrl"],
        "pruAccessName": fund["pruAccessName"],
        "fundName": result.get("fundName"),
        "finalUrl": result.get("finalUrl"),
        "factsheetUrl": result.get("factsheetUrl"),
        "dataAsAt": result.get("dataAsAt"),
        "documentDate": result.get("documentDate"),
        "status": result.get("status"),
        "holdingsSectionStatus": result.get(
            "holdingsSectionStatus"
        ),
        "topHoldingsCount": 0,
        "files": {
            "factsheetPdf": str(pdf_path),
            "factsheetText": str(text_path),
            "topHoldingsSection": str(section_path),
        },
    }

    save_json(
        metadata_path,
        metadata,
    )

    return output_dir


# ============================================================================
# SAVE FAILURE
# ============================================================================

def save_failure_result(fund, error):
    excel_row = fund["excelRow"]

    directory_name = (
        f"{excel_row}_failed"
    )

    output_dir = FAILED_OUTPUT_DIR / directory_name

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    failure = {
        "status": "failed",
        "excelRow": fund["excelRow"],
        "prudentialUrl": fund["prudentialUrl"],
        "pruAccessName": fund["pruAccessName"],
        "error": str(error),
        "generatedAtUtc": utc_now_iso(),
    }

    save_json(
        output_dir / "failure.json",
        failure,
    )


# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 68)
    print("VGrat FMS - PRUDENTIAL TOP HOLDINGS ALL-FUND TEST")
    print("=" * 68)
    print()

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

    funds = read_excel_funds()

    print(
        f"Excel fund universe : {len(funds)}"
    )

    if not funds:
        raise RuntimeError(
            "No populated Prudential URLs found in Excel Column A"
        )

    all_results = []

    successful_results = []
    no_holdings_results = []
    failed_results = []

    total_holdings = 0

    started_at = utc_now_iso()

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=HEADLESS
        )

        context = browser.new_context(
            accept_downloads=True,
        )

        page = context.new_page()

        page.set_default_timeout(
            PAGE_TIMEOUT_MS
        )

        for position, fund in enumerate(
            funds,
            start=1,
        ):
            excel_row = fund["excelRow"]

            print()
            print("-" * 68)
            print(
                f"[{position}/{len(funds)}] "
                f"Excel Row {excel_row}"
            )
            print(
                f"URL: {fund['prudentialUrl']}"
            )

            success = False
            last_error = None

            for attempt in range(
                1,
                MAX_RETRIES + 1,
            ):
                try:
                    result = extract_single_fund(
                        page,
                        fund,
                    )

                    # Re-download the official factsheet once the section
                    # has been confirmed. This verifies that the extraction
                    # result is based on the same official PDF source.
                    factsheet_url = result.get(
                        "factsheetUrl"
                    )

                    pdf_bytes = download_factsheet(
                        page,
                        factsheet_url,
                    )

                    pdf_text = extract_pdf_text(
                        pdf_bytes
                    )

                    verified_section, verified_status = (
                        extract_holdings_section(
                            pdf_text
                        )
                    )

                    if result["status"] == "success":
                        if verified_status != "confirmed":
                            raise RuntimeError(
                                "Holdings section disappeared during "
                                "verification"
                            )

                        if result["parser"] == "fallback":
                            verified_holdings = (
                                parse_holdings_fallback(
                                    verified_section
                                )
                            )
                        else:
                            try:
                                verified_holdings = (
                                    parse_holdings(
                                        verified_section
                                    )
                                )
                            except Exception:
                                verified_holdings = (
                                    parse_holdings_fallback(
                                        verified_section
                                    )
                                )

                        if len(verified_holdings) != result[
                            "topHoldingsCount"
                        ]:
                            raise RuntimeError(
                                "Verification holding count differs: "
                                f"initial={result['topHoldingsCount']}, "
                                f"verified={len(verified_holdings)}"
                            )

                        result["topHoldings"] = (
                            verified_holdings
                        )

                        result["topHoldingsCount"] = (
                            len(verified_holdings)
                        )

                        save_success_result(
                            fund,
                            result,
                            pdf_bytes,
                            pdf_text,
                            verified_section,
                        )

                        successful_results.append(
                            result
                        )

                        total_holdings += len(
                            verified_holdings
                        )

                        print(
                            "STATUS: SUCCESS"
                        )
                        print(
                            f"Fund: {result.get('fundName', '')}"
                        )
                        print(
                            f"Top Holdings: "
                            f"{len(verified_holdings)}"
                        )
                        print(
                            f"Parser: "
                            f"{result.get('parser')}"
                        )

                    else:
                        save_no_holdings_result(
                            fund,
                            result,
                            pdf_bytes,
                            pdf_text,
                        )

                        no_holdings_results.append(
                            result
                        )

                        print(
                            "STATUS: NO CONFIRMED "
                            "TOP HOLDINGS SECTION"
                        )
                        print(
                            f"Fund: "
                            f"{result.get('fundName', '')}"
                        )

                    all_results.append(result)

                    success = True
                    break

                except Exception as exc:
                    last_error = exc

                    print(
                        f"Attempt {attempt}/{MAX_RETRIES} failed: "
                        f"{exc}"
                    )

                    if attempt < MAX_RETRIES:
                        time.sleep(
                            RETRY_DELAY_SECONDS
                        )

            if not success:
                failure = {
                    "status": "failed",
                    "excelRow": excel_row,
                    "prudentialUrl": fund[
                        "prudentialUrl"
                    ],
                    "pruAccessName": fund[
                        "pruAccessName"
                    ],
                    "error": str(last_error),
                }

                failed_results.append(
                    failure
                )

                all_results.append(
                    failure
                )

                save_failure_result(
                    fund,
                    last_error,
                )

                print(
                    "STATUS: FAILED"
                )
                print(
                    f"Error: {last_error}"
                )

        context.close()
        browser.close()

    # ========================================================================
    # FINAL JSON OUTPUT
    # ========================================================================

    finished_at = utc_now_iso()

    all_holdings_payload = {
        "status": "success",
        "generatedAtUtc": finished_at,
        "startedAtUtc": started_at,
        "excelFundUniverse": len(funds),
        "successfulFunds": len(
            successful_results
        ),
        "noHoldingsFunds": len(
            no_holdings_results
        ),
        "failedFunds": len(
            failed_results
        ),
        "totalHoldings": total_holdings,
        "results": successful_results,
    }

    save_json(
        ALL_HOLDINGS_FILE,
        all_holdings_payload,
    )

    run_summary = {
        "status": (
            "success"
            if not failed_results
            else "completed_with_failures"
        ),
        "generatedAtUtc": finished_at,
        "excelFile": str(EXCEL_FILE),
        "excelFundUniverse": len(funds),
        "successfulFunds": len(
            successful_results
        ),
        "noHoldingsFunds": len(
            no_holdings_results
        ),
        "failedFunds": len(
            failed_results
        ),
        "totalHoldings": total_holdings,
        "failedRows": [
            item["excelRow"]
            for item in failed_results
        ],
        "successfulRows": [
            item["excelRow"]
            for item in successful_results
        ],
        "noHoldingsRows": [
            item["excelRow"]
            for item in no_holdings_results
        ],
        "rules": {
            "excelColumnAControlsUniverse": True,
            "processEveryPopulatedColumnAUrl": True,
            "hardcodedFundLimit": False,
            "duplicateUrlsPreserved": True,
            "officialPrudentialSingaporeOnly": True,
            "officialFactsheetsOnly": True,
            "thirdPartySources": False,
            "confirmedTopHoldingsDetection": True,
            "confirmedHeadingPatterns": [
                "Top 10 Holdings",
                "Top Ten Holdings",
                "Top 10 Holdings<footnote digits>",
                "Top Ten Holdings<footnote digits>",
            ],
            "unicodeFootnoteDigitsAccepted": True,
            "genericHoldingsHeadingAccepted": False,
            "genericInvestmentsHeadingAccepted": False,
            "securityNamesContainingHoldingsNotHeadings": True,
            "candidateHeadingsAutomaticallyAccepted": False,
            "holdingsWeightInterpretation": True,
            "percentageInterpretationPerformed": True,
            "holdingsNamesInferred": False,
            "fabricatedData": False,
            "syntheticData": False,
            "pypdfTextExtraction": True,
            "rawTopHoldingsPreserved": True,
            "fullPdfPreserved": True,
            "noForcedTenHoldings": True,
            "publishedOrderPreserved": True,
            "multiLineHoldingsReconstructed": True,
            "fixedIncomeLastPercentageFallback": True,
            "verificationPass": True,
        },
        "failed": failed_results,
    }

    save_json(
        RUN_SUMMARY_FILE,
        run_summary,
    )

    # ========================================================================
    # CONSOLE SUMMARY
    # ========================================================================

    print()
    print()
    print("=" * 68)
    print("FINAL SUMMARY")
    print("=" * 68)
    print(
        f"Excel fund universe : {len(funds)}"
    )
    print(
        f"Successful funds    : "
        f"{len(successful_results)}"
    )
    print(
        f"No holdings section : "
        f"{len(no_holdings_results)}"
    )
    print(
        f"Failed funds        : "
        f"{len(failed_results)}"
    )
    print(
        f"Total holdings      : "
        f"{total_holdings}"
    )

    if failed_results:
        print()
        print("FAILED ROWS:")

        for failure in failed_results:
            print(
                f"  Row {failure['excelRow']}: "
                f"{failure['error']}"
            )

    print()
    print("CONFIRMED HEADING RULE:")
    print("  Top 10 Holdings")
    print("  Top Ten Holdings")
    print("  Top 10 Holdings3")
    print("  Top Ten Holdings3")
    print()
    print(
        "Generic 'holdings' / 'investments' headings "
        "are NOT accepted."
    )
    print()
    print(
        f"Output: {OUTPUT_DIR}"
    )
    print(
        f"All holdings: {ALL_HOLDINGS_FILE}"
    )
    print(
        f"Run summary: {RUN_SUMMARY_FILE}"
    )
    print("=" * 68)


if __name__ == "__main__":
    try:
        main()

    except KeyboardInterrupt:
        print(
            "\nExecution interrupted by user.",
            file=sys.stderr,
        )
        sys.exit(130)

    except Exception as exc:
        print(
            "\nFATAL ERROR:",
            file=sys.stderr,
        )
        print(
            str(exc),
            file=sys.stderr,
        )
        sys.exit(1)
