```python
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
          +---- if ambiguous ----> robust fallback parser
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

26. Fallback parser handles fixed-income holdings where coupon/rate
    percentages and portfolio-weight percentages coexist.

27. Fallback parser handles multiple holdings appearing on one physical
    PDF line.

28. Fallback parser handles maturity/date fragments split across
    physical PDF lines.

29. A zero percentage is numerically valid.
    0.00% must NOT automatically be rejected.

30. When two percentages are directly adjacent with only whitespace or
    table separators between them, the earlier percentage may be an
    embedded security rate/coupon and the later percentage is the
    portfolio weight.

31. No synthetic holdings.

32. No synthetic weights.

33. No interpolation.

34. No carry-forward.

35. This script does NOT modify:
       test_pruaccess.py
       data.json
       index.html
       css
       js

36. Output is written under:

       output_holdings/

    including:

       output_holdings/all_holdings.json
       output_holdings/run_summary.json
       output_holdings/funds/
       output_holdings/<row>_failed/

37. Whenever this script is changed, replace the ENTIRE script with the
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

import io
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

DATE_FRAGMENT_RE = re.compile(
    r"""
    \b
    \d{1,2}
    [-/]
    (?:
        JAN|FEB|MAR|APR|MAY|JUN|
        JUL|AUG|SEP|OCT|NOV|DEC
    )
    (?:[-/]?\d{4})?
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

    value = re.sub(
        r"[<>:\"/\\|?*\x00-\x1f]",
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

    return value[:max_length]


def normalize_whitespace(value: str) -> str:
    value = value.replace("\u00a0", " ")
    value = value.replace("\u200b", "")
    value = value.replace("\ufeff", "")

    value = re.sub(
        r"[ \t]+",
        " ",
        value,
    )

    return value.strip()


def normalize_pdf_line(line: str) -> str:
    line = line.replace("\u00a0", " ")
    line = line.replace("\u200b", "")
    line = line.replace("\ufeff", "")

    line = line.replace("│", " ")
    line = line.replace("|", " ")

    line = re.sub(
        r"\s+",
        " ",
        line,
    )

    return line.strip()


def is_official_prudential_url(url: str) -> bool:
    try:
        parsed = urlparse(url)

        host = (
            parsed.hostname or ""
        ).lower()

        return host in OFFICIAL_HOSTS

    except Exception:
        return False


def ensure_official_url(
    url: str,
    description: str = "URL",
) -> None:

    if not is_official_prudential_url(url):
        raise RuntimeError(
            f"{description} is not an official Prudential Singapore URL: "
            f"{url}"
        )


def write_json(
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


def clean_holding_name(
    name: str,
) -> str:

    if not name:
        return ""

    name = name.replace(
        "\u00a0",
        " ",
    )

    name = name.replace(
        "\u200b",
        "",
    )

    name = name.replace(
        "\ufeff",
        "",
    )

    # Remove PDF bullets.
    name = re.sub(
        r"^[•·▪◦●■□◆◇►]+\s*",
        "",
        name,
    )

    # Remove a leading rank.
    name = re.sub(
        r"^\s*\d{1,2}\s*(?:[.)]|[-:])\s*",
        "",
        name,
    )

    # Also handle a rank followed simply by whitespace.
    name = re.sub(
        r"^\s*\d{1,2}\s+(?=[A-Z])",
        "",
        name,
    )

    # Remove common table headings that can occur before the first
    # actual holding.
    name = re.sub(
        r"^\s*(?:holding|holdings|security|securities)"
        r"\s*(?:name)?\s*(?:weight|weights|portfolio\s+weight)?"
        r"\s*",
        "",
        name,
        flags=re.IGNORECASE,
    )

    # Remove a leading "Name Weight" style header.
    name = re.sub(
        r"^\s*name\s+weight\s+",
        "",
        name,
        flags=re.IGNORECASE,
    )

    # Remove standalone "None".
    name = re.sub(
        r"^\s*None\s+",
        "",
        name,
        flags=re.IGNORECASE,
    )

    # Remove separators immediately before the portfolio weight.
    name = re.sub(
        r"\s+[-–—:|]+\s*$",
        "",
        name,
    )

    # Remove trailing punctuation left by PDF extraction.
    name = re.sub(
        r"\s+[,;]+\s*$",
        "",
        name,
    )

    name = re.sub(
        r"\s+",
        " ",
        name,
    )

    return name.strip()


def extract_leading_rank(
    line: str,
) -> int | None:

    match = RANK_RE.match(line)

    if not match:
        return None

    try:
        return int(
            match.group("number")
        )
    except Exception:
        return None


def count_percentages(
    line: str,
) -> int:

    return len(
        PERCENT_RE.findall(line)
    )


def find_percentage_matches(
    line: str,
):
    return list(
        PERCENT_RE.finditer(line)
    )


def find_percentage_in_line(
    line: str,
):
    matches = find_percentage_matches(
        line
    )

    if not matches:
        return None

    match = matches[0]

    try:
        value = float(
            match.group(1)
        )
    except Exception:
        return None

    return {
        "value": value,
        "text": match.group(0),
        "start": match.start(),
        "end": match.end(),
    }


def contains_maturity_date(
    text: str,
) -> bool:

    if not text:
        return False

    return bool(
        DATE_TOKEN_RE.search(text)
    )


def contains_date_fragment(
    text: str,
) -> bool:

    if not text:
        return False

    return bool(
        DATE_FRAGMENT_RE.search(text)
    )


def contains_year(
    text: str,
) -> bool:

    if not text:
        return False

    return bool(
        YEAR_ONLY_RE.search(text)
    )


def has_date_like_fragment(
    text: str,
) -> bool:

    if not text:
        return False

    if contains_maturity_date(text):
        return True

    if contains_date_fragment(text):
        return True

    stripped = normalize_pdf_line(
        text
    )

    if YEAR_ONLY_RE.fullmatch(
        stripped
    ):
        return True

    return False


def looks_like_portfolio_weight(
    value: float,
) -> bool:
    """
    Validate only the mathematical range.

    IMPORTANT:

    0.0% is valid.

    The parser must not reject zero merely because it is zero.
    """

    return (
        value >= 0
        and value <= 100
    )


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

        for row_number in range(
            2,
            worksheet.max_row + 1,
        ):

            url_value = worksheet.cell(
                row=row_number,
                column=1,
            ).value

            pruaccess_value = worksheet.cell(
                row=row_number,
                column=2,
            ).value

            if url_value is None:
                continue

            url = str(
                url_value
            ).strip()

            if not url:
                continue

            if not is_official_prudential_url(
                url
            ):
                raise RuntimeError(
                    f"Excel row {row_number} contains a non-official "
                    f"Prudential URL: {url}"
                )

            funds.append(
                {
                    "excelRow": row_number,
                    "prudentialUrl": url,
                    "pruaccessName": (
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

def score_factsheet_anchor(
    anchor,
) -> int:

    try:
        text = normalize_whitespace(
            anchor.inner_text()
        )
    except Exception:
        text = ""

    try:
        href = (
            anchor.get_attribute("href")
            or ""
        )
    except Exception:
        href = ""

    combined = (
        f"{text} {href}"
    ).lower()

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


def find_factsheet_url(
    page,
    fund_url: str,
) -> str:

    anchors = page.locator(
        "a"
    )

    candidates = []

    try:
        count = anchors.count()
    except Exception:
        count = 0

    for index in range(count):

        anchor = anchors.nth(
            index
        )

        try:
            href = anchor.get_attribute(
                "href"
            )
        except Exception:
            continue

        if not href:
            continue

        absolute_url = urljoin(
            fund_url,
            href,
        )

        if not is_official_prudential_url(
            absolute_url
        ):
            continue

        score = score_factsheet_anchor(
            anchor
        )

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


def extract_pdf_text(
    pdf_bytes: bytes,
) -> str:

    if not pdf_bytes:
        raise RuntimeError(
            "Factsheet PDF is empty."
        )

    reader = PdfReader(
        io.BytesIO(
            pdf_bytes
        )
    )

    pages = []

    for page_number, page in enumerate(
        reader.pages,
        start=1,
    ):

        try:
            text = (
                page.extract_text()
                or ""
            )

        except Exception as exc:

            raise RuntimeError(
                f"Failed extracting PDF page "
                f"{page_number}: {exc}"
            ) from exc

        pages.append(
            text
        )

    full_text = "\n".join(
        pages
    )

    if not full_text.strip():
        raise RuntimeError(
            "Factsheet PDF produced no extractable text."
        )

    return full_text


# ============================================================================
# FACTSHEET DATES
# ============================================================================

def extract_data_as_at(
    text: str,
) -> str | None:

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
            return normalize_whitespace(
                match.group(1)
            )

    return None


def extract_document_date(
    text: str,
) -> str | None:

    month_year_re = re.compile(
        r"\b(?:January|February|March|April|May|June|July|August|"
        r"September|October|November|December|"
        r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"\s+\d{4}\b",
        re.IGNORECASE,
    )

    matches = month_year_re.findall(
        text
    )

    if not matches:
        return None

    return normalize_whitespace(
        matches[-1]
    )


# ============================================================================
# TOP HOLDINGS SECTION
# ============================================================================

def find_holdings_start(
    lines: list[str],
) -> int | None:

    for index, line in enumerate(
        lines
    ):

        if TOP_HOLDINGS_RE.search(
            line
        ):
            return index

    return None


def is_holdings_end(
    line: str,
) -> bool:

    normalized = normalize_pdf_line(
        line
    )

    lowered = normalized.lower()

    if not normalized:
        return False

    if PAGE_NUMBER_RE.match(
        normalized
    ):
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
        "dividend history",
    ]

    for ending in endings:

        if lowered.startswith(
            ending
        ):
            return True

    return False


def extract_holdings_section(
    text: str,
) -> tuple[str, list[str]]:

    raw_lines = text.splitlines()

    lines = [
        normalize_pdf_line(line)
        for line in raw_lines
    ]

    start_index = find_holdings_start(
        lines
    )

    if start_index is None:
        raise LookupError(
            "no_holdings_section"
        )

    section_lines = []

    for index in range(
        start_index + 1,
        len(lines),
    ):

        line = lines[index]

        if is_holdings_end(
            line
        ):
            break

        section_lines.append(
            line
        )

    while (
        section_lines
        and not section_lines[0]
    ):
        section_lines.pop(0)

    while (
        section_lines
        and not section_lines[-1]
    ):
        section_lines.pop()

    section_text = "\n".join(
        section_lines
    )

    if not section_text.strip():
        raise RuntimeError(
            "Top 10 Holdings heading was found but its section was empty."
        )

    return (
        section_text,
        section_lines,
    )


# ============================================================================
# PRIMARY PARSER
# ============================================================================

def build_primary_logical_lines(
    lines: list[str],
) -> list[str]:

    logical_lines = []

    pending = ""

    for raw_line in lines:

        line = normalize_pdf_line(
            raw_line
        )

        if not line:
            continue

        if RANK_ONLY_RE.match(
            line
        ):

            if pending:
                logical_lines.append(
                    pending
                )
                pending = ""

            logical_lines.append(
                line
            )

            continue

        if not pending:
            pending = line
            continue

        previous_has_percent = bool(
            PERCENT_RE.search(
                pending
            )
        )

        current_has_percent = bool(
            PERCENT_RE.search(
                line
            )
        )

        if (
            previous_has_percent
            and current_has_percent
        ):

            logical_lines.append(
                pending
            )

            pending = line

            continue

        if (
            previous_has_percent
            and not current_has_percent
        ):

            pending = (
                f"{pending} {line}"
            )

            continue

        pending = (
            f"{pending} {line}"
        )

    if pending:
        logical_lines.append(
            pending
        )

    return logical_lines


def parse_holdings_primary(
    lines: list[str],
) -> list[dict]:

    logical_lines = (
        build_primary_logical_lines(
            lines
        )
    )

    holdings = []

    pending_fragments = []
    pending_rank = None

    for line in logical_lines:

        line = normalize_pdf_line(
            line
        )

        if not line:
            continue

        rank = extract_leading_rank(
            line
        )

        if rank is not None:

            line_without_rank = (
                RANK_RE.sub(
                    "",
                    line,
                ).strip()
            )

            if pending_fragments:

                pending_name = clean_holding_name(
                    " ".join(
                        pending_fragments
                    )
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

        percentage_count = count_percentages(
            line
        )

        if percentage_count > 1:

            raise RuntimeError(
                "Primary parser found multiple percentages on one logical "
                "line and will not guess the holding weight. "
                "Fallback parser required."
            )

        percentage = find_percentage_in_line(
            line
        )

        if percentage is None:

            pending_fragments.append(
                line
            )

            continue

        name_fragment = (
            line[
                :percentage["start"]
            ].strip()
        )

        if name_fragment:
            pending_fragments.append(
                name_fragment
            )

        holding_name = clean_holding_name(
            " ".join(
                pending_fragments
            )
        )

        if not holding_name:

            raise RuntimeError(
                "Primary parser found a percentage without a holding name."
            )

        weight = percentage[
            "value"
        ]

        if not looks_like_portfolio_weight(
            weight
        ):

            raise RuntimeError(
                f"Primary parser found an invalid percentage: "
                f"{weight}%"
            )

        holdings.append(
            {
                "rank": (
                    pending_rank
                    if pending_rank is not None
                    else len(holdings) + 1
                ),
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
            " ".join(
                pending_fragments
            )
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
# FALLBACK PARSER
# ============================================================================

def flatten_holdings_section(
    lines: list[str],
) -> str:

    cleaned = []

    for raw_line in lines:

        line = normalize_pdf_line(
            raw_line
        )

        if not line:
            continue

        cleaned.append(
            line
        )

    return normalize_whitespace(
        " ".join(
            cleaned
        )
    )


def percentage_has_maturity_before_next_percentage(
    text: str,
    match,
    next_match,
) -> bool:

    if next_match is None:

        following = text[
            match.end():
        ]

    else:

        following = text[
            match.end():
            next_match.start()
        ]

    following = normalize_pdf_line(
        following
    )

    if not following:
        return False

    if contains_maturity_date(
        following
    ):
        return True

    if contains_date_fragment(
        following
    ):
        return True

    if YEAR_ONLY_RE.search(
        following
    ):
        return True

    return False


def percentage_is_immediately_before_next_percentage(
    text: str,
    match,
    next_match,
) -> bool:
    """
    Detect percentages such as:

        USA DL-Treasury Bills 2026(26) 0.00% 5.8%

    Here the text between the two percentages is only whitespace/table
    separation.

    Therefore the first percentage is treated as an embedded security
    rate and the second percentage is treated as the portfolio weight.

    This is deliberately narrow so that normal same-line holdings such as:

        SECURITY A 5.8% SECURITY B 4.2%

    are NOT incorrectly merged.
    """

    if next_match is None:
        return False

    between = text[
        match.end():
        next_match.start()
    ]

    between = between.strip()

    if not between:
        return True

    # Allow common PDF/table separators only.
    if re.fullmatch(
        r"[-–—:|;,/]+",
        between,
    ):
        return True

    return False


def strip_leading_table_rank(
    text: str,
) -> str:

    text = normalize_pdf_line(
        text
    )

    if not text:
        return ""

    text = re.sub(
        r"^\s*\d{1,2}\s*[.)]\s*",
        "",
        text,
    )

    text = re.sub(
        r"^\s*\d{1,2}\s+(?=[A-Z])",
        "",
        text,
    )

    return text.strip()


def remove_leading_holdings_headers(
    text: str,
) -> str:

    text = normalize_pdf_line(
        text
    )

    if not text:
        return ""

    header_patterns = [
        r"^\s*holdings?\s+weight\s+",
        r"^\s*holdings?\s+portfolio\s+weight\s+",
        r"^\s*security\s+weight\s+",
        r"^\s*name\s+weight\s+",
        r"^\s*holding\s+name\s+weight\s+",
    ]

    changed = True

    while changed:

        changed = False

        for pattern in header_patterns:

            updated = re.sub(
                pattern,
                "",
                text,
                count=1,
                flags=re.IGNORECASE,
            )

            if updated != text:
                text = updated
                changed = True

    return text.strip()


def commit_fallback_holding(
    holdings: list[dict],
    candidate_name: str,
    weight: float,
) -> None:

    candidate_name = clean_holding_name(
        candidate_name
    )

    candidate_name = strip_leading_table_rank(
        candidate_name
    )

    candidate_name = clean_holding_name(
        candidate_name
    )

    if not candidate_name:
        raise RuntimeError(
            "Fallback parser identified a portfolio weight but could not "
            "identify a holding name."
        )

    if not looks_like_portfolio_weight(
        weight
    ):
        raise RuntimeError(
            f"Fallback parser found invalid portfolio weight: "
            f"{weight}%"
        )

    holdings.append(
        {
            "rank": len(holdings) + 1,
            "name": candidate_name,
            "weightPercent": float(weight),
        }
    )


def parse_holdings_fallback(
    lines: list[str],
) -> list[dict]:
    """
    Robust fallback parser.

    Classification rules:

    1. percentage + maturity/date + percentage
       ->
       first percentage = security coupon/rate
       second percentage = portfolio weight

    2. percentage + percentage
       with only whitespace/separators between them
       ->
       first percentage = embedded security rate
       second percentage = portfolio weight

    Example:

        USA DL-Treasury Bills 2026(26) 0.00% 5.8%

    becomes:

        name:
            USA DL-Treasury Bills 2026(26) 0.00%

        weight:
            5.8%

    Example:

        INDIA (REPUBLIC OF) 7.09% 5-AUG-2054 - 2.7%

    becomes:

        name:
            INDIA (REPUBLIC OF) 7.09% 5-AUG-2054

        weight:
            2.7%

    Multiple holdings on one physical PDF line are also handled.
    """

    text = flatten_holdings_section(
        lines
    )

    if not text:
        raise RuntimeError(
            "Fallback parser received an empty Top Holdings section."
        )

    text = remove_leading_holdings_headers(
        text
    )

    matches = find_percentage_matches(
        text
    )

    if not matches:
        raise RuntimeError(
            "Fallback parser found no percentages in the Top Holdings section."
        )

    holdings = []

    candidate_parts = []

    cursor = 0

    for index, match in enumerate(
        matches
    ):

        next_match = (
            matches[index + 1]
            if index + 1 < len(matches)
            else None
        )

        before = text[
            cursor:
            match.start()
        ].strip()

        if before:
            candidate_parts.append(
                before
            )

        percentage_text = match.group(
            0
        )

        try:
            percentage_value = float(
                match.group(1)
            )
        except Exception as exc:

            raise RuntimeError(
                f"Fallback parser could not read percentage: "
                f"{percentage_text}"
            ) from exc

        # --------------------------------------------------------------
        # RULE 1:
        #
        # Current percentage followed by a maturity/date and then
        # another percentage.
        #
        # Example:
        #
        #     7.09% 5-AUG-2054 - 2.7%
        #
        # 7.09% is part of the security name.
        # 2.7% is the portfolio weight.
        # --------------------------------------------------------------

        is_coupon = (
            percentage_has_maturity_before_next_percentage(
                text=text,
                match=match,
                next_match=next_match,
            )
        )

        # --------------------------------------------------------------
        # RULE 2:
        #
        # Current percentage is immediately followed by another
        # percentage.
        #
        # Example:
        #
        #     0.00% 5.8%
        #
        # The first percentage is embedded in the security name.
        # The second is the portfolio weight.
        #
        # This specifically fixes fixed-income securities such as
        # Treasury Bills where a coupon/rate can be 0.00%.
        # --------------------------------------------------------------

        if not is_coupon:

            is_coupon = (
                percentage_is_immediately_before_next_percentage(
                    text=text,
                    match=match,
                    next_match=next_match,
                )
            )

        if is_coupon:

            candidate_parts.append(
                percentage_text
            )

            cursor = match.end()

            continue

        # ------------------------------------------------------------------
        # This percentage is a portfolio weight.
        # ------------------------------------------------------------------

        candidate_name = normalize_whitespace(
            " ".join(
                candidate_parts
            )
        )

        candidate_name = clean_holding_name(
            candidate_name
        )

        if not candidate_name:

            raise RuntimeError(
                "Fallback parser found a portfolio weight without a "
                "holding name."
            )

        commit_fallback_holding(
            holdings=holdings,
            candidate_name=candidate_name,
            weight=percentage_value,
        )

        candidate_parts = []

        cursor = match.end()

        if len(holdings) >= MAX_HOLDINGS:
            break

    if not holdings:
        raise RuntimeError(
            "Fallback parser extracted zero holdings."
        )

    return holdings[:MAX_HOLDINGS]


# ============================================================================
# HOLDINGS VALIDATION
# ============================================================================

def validate_holdings(
    holdings: list[dict],
) -> None:

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

        if not isinstance(
            holding,
            dict,
        ):
            raise RuntimeError(
                "Holding is not an object."
            )

        name = holding.get(
            "name"
        )

        weight = holding.get(
            "weightPercent"
        )

        if not name or not str(name).strip():
            raise RuntimeError(
                "Holding has an empty name."
            )

        if weight is None:
            raise RuntimeError(
                f"Holding has no weight: {name}"
            )

        try:
            numeric_weight = float(
                weight
            )
        except Exception as exc:

            raise RuntimeError(
                f"Holding has invalid weight: "
                f"{name} -> {weight}"
            ) from exc

        if not looks_like_portfolio_weight(
            numeric_weight
        ):
            raise RuntimeError(
                f"Holding has invalid percentage: "
                f"{name} -> {numeric_weight}%"
            )

        cleaned_name = clean_holding_name(
            str(name)
        )

        if not cleaned_name:
            raise RuntimeError(
                "Holding name became empty after cleaning."
            )

        lowered = cleaned_name.lower()

        if lowered in {
            "holdings",
            "holding",
            "security",
            "securities",
            "weight",
            "weights",
            "name",
        }:
            raise RuntimeError(
                f"Parser produced a table heading instead of a holding: "
                f"{cleaned_name}"
            )

        holding["rank"] = expected_rank

        holding["name"] = cleaned_name

        holding["weightPercent"] = numeric_weight

        expected_rank += 1


# ============================================================================
# FUND NAME
# ============================================================================

def extract_fund_name_from_text(
    text: str,
) -> str | None:

    lines = [
        normalize_pdf_line(line)
        for line in text.splitlines()
    ]

    candidates = []

    for index, line in enumerate(
        lines[:80]
    ):

        if not line:
            continue

        lowered = line.lower()

        if "prulink" in lowered:

            candidates.append(
                line
            )

            if index + 1 < len(lines):

                next_line = lines[
                    index + 1
                ]

                if (
                    next_line
                    and "fund" in next_line.lower()
                    and len(next_line) < 160
                ):

                    combined = normalize_whitespace(
                        f"{line} {next_line}"
                    )

                    candidates.append(
                        combined
                    )

    if not candidates:
        return None

    for candidate in candidates:

        lowered = candidate.lower()

        if (
            "prulink" in lowered
            and "fund" in lowered
        ):
            return candidate

    return candidates[0]


# ============================================================================
# FUND EXTRACTION
# ============================================================================

def extract_single_fund(
    browser,
    fund: dict,
) -> dict:

    excel_row = fund[
        "excelRow"
    ]

    prudential_url = fund[
        "prudentialUrl"
    ]

    pruaccess_name = fund[
        "pruaccessName"
    ]

    print("")
    print("=" * 60)
    print(
        f"PROCESSING FUND ROW {excel_row}"
    )
    print("=" * 60)
    print("")

    print(
        f"Prudential URL: {prudential_url}"
    )

    print(
        f"Excel PruAccess name: {pruaccess_name}"
    )

    ensure_official_url(
        prudential_url,
        "Prudential fund URL",
    )

    page = browser.new_page()

    try:

        page.set_default_timeout(
            PAGE_TIMEOUT_MS
        )

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
        print(
            "Factsheet link found:"
        )

        print(
            factsheet_url
        )

        print("")
        print(
            "Downloading factsheet..."
        )

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

        section_text, section_lines = (
            extract_holdings_section(
                pdf_text
            )
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
                "running robust fallback parser..."
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
        # FINAL VALIDATION
        # --------------------------------------------------------------

        validate_holdings(
            holdings
        )

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

        fund_name = extract_fund_name_from_text(
            pdf_text
        )

        if fund_name:
            result[
                "fundName"
            ] = fund_name

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
# OUTPUT
# ============================================================================

def get_output_identifier(
    result: dict,
) -> str:

    fund_name = result.get(
        "fundName"
    )

    if fund_name:

        return safe_filename(
            fund_name
        )

    url = result.get(
        "prudentialUrl",
        "",
    )

    parsed = urlparse(
        url
    )

    path = parsed.path.rstrip(
        "/"
    )

    if path:

        slug = path.split(
            "/"
        )[-1]

        if slug.lower().endswith(
            ".html"
        ):
            slug = slug[:-5]

        if slug:

            return safe_filename(
                slug
            )

    return "fund"


def save_success_result(
    result: dict,
    pdf_bytes: bytes,
    pdf_text: str,
    section_text: str,
) -> Path:

    excel_row = result[
        "excelRow"
    ]

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

    pdf_path = (
        fund_dir
        / "factsheet.pdf"
    )

    with pdf_path.open(
        "wb"
    ) as handle:

        handle.write(
            pdf_bytes
        )

    text_path = (
        fund_dir
        / "factsheet_text.txt"
    )

    text_path.write_text(
        pdf_text,
        encoding="utf-8",
    )

    section_path = (
        fund_dir
        / "top_holdings_section.txt"
    )

    section_path.write_text(
        section_text,
        encoding="utf-8",
    )

    holdings_path = (
        fund_dir
        / "top_holdings.json"
    )

    write_json(
        holdings_path,
        {
            "status": result[
                "status"
            ],
            "excelRow": result[
                "excelRow"
            ],
            "fundName": result.get(
                "fundName"
            ),
            "prudentialUrl": result[
                "prudentialUrl"
            ],
            "excelPruAccessName": result[
                "excelPruAccessName"
            ],
            "factsheetUrl": result[
                "factsheetUrl"
            ],
            "factsheetDataAsAt": result[
                "factsheetDataAsAt"
            ],
            "factsheetDocumentDate": result[
                "factsheetDocumentDate"
            ],
            "parser": result[
                "parser"
            ],
            "holdings": result[
                "holdings"
            ],
        },
    )

    metadata_path = (
        fund_dir
        / "metadata.json"
    )

    write_json(
        metadata_path,
        {
            "excelRow": result[
                "excelRow"
            ],
            "fundName": result.get(
                "fundName"
            ),
            "prudentialUrl": result[
                "prudentialUrl"
            ],
            "excelPruAccessName": result[
                "excelPruAccessName"
            ],
            "factsheetUrl": result[
                "factsheetUrl"
            ],
            "factsheetDataAsAt": result[
                "factsheetDataAsAt"
            ],
            "factsheetDocumentDate": result[
                "factsheetDocumentDate"
            ],
            "parser": result[
                "parser"
            ],
            "holdingsCount": len(
                result[
                    "holdings"
                ]
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

    excel_row = fund[
        "excelRow"
    ]

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
            "pruaccessName"
        ],
        "error": str(
            error
        ),
        "failedAtUtc": utc_now_iso(),
    }

    if attempt is not None:

        failure[
            "attempt"
        ] = attempt

    failure_path = (
        failure_dir
        / "failure.json"
    )

    write_json(
        failure_path,
        failure,
    )

    return failure_path


# ============================================================================
# CONSOLE OUTPUT
# ============================================================================

def print_holdings(
    result: dict,
) -> None:

    print("")
    print(
        "SUCCESS"
    )

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
        "Parser used: "
        f"{result.get('parser') or '-'}"
    )

    print(
        "Top holdings extracted: "
        f"{len(result['holdings'])}"
    )

    for holding in result[
        "holdings"
    ]:

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
    print(
        "=" * 70
    )

    print(
        "VGrat FMS - PRUDENTIAL TOP HOLDINGS "
        "ALL-FUND TEST EXTRACTOR"
    )

    print(
        "=" * 70
    )

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

    total = len(
        funds
    )

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

                        result = extracted[
                            "result"
                        ]

                        validate_holdings(
                            result[
                                "holdings"
                            ]
                        )

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
                                    "pruaccessName"
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

                    if (
                        attempt
                        < MAX_RETRIES
                    ):

                        time.sleep(
                            RETRY_DELAY_SECONDS
                        )

                if not success:

                    print("")
                    print(
                        "FAILED"
                    )

                    failure_path = save_failure(
                        fund,
                        (
                            last_error
                            if last_error is not None
                            else "Unknown extraction failure."
                        ),
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
                                "pruaccessName"
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
        "excelFundUniverse": len(
            funds
        ),
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
            "zeroPercentAllowed": True,
            "portfolioWeightUsesFinalRelevantPercentage": True,
            "multiLineHoldingsSupported": True,
            "multipleHoldingsPerPdfLineSupported": True,
            "maturityDateContinuationSupported": True,
            "fallbackUsesDateAwarePercentageClassification": True,
            "fallbackHandlesAdjacentPercentages": True,
            "fallbackFlattensPhysicalPdfLines": True,
            "noSyntheticHoldings": True,
            "noSyntheticWeights": True,
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
        "excelFundUniverse": len(
            funds
        ),
        "successfulFunds": len(
            successful_results
        ),
        "failedFunds": [
            failure[
                "excelRow"
            ]
            for failure in failed_results
        ],
        "successfulFundRows": [
            result[
                "excelRow"
            ]
            for result in successful_results
        ],
        "totalHoldingsExtracted": sum(
            len(
                result.get(
                    "holdings",
                    []
                )
            )
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
            "zeroPercentAllowed": True,
            "multipleHoldingsOnSamePdfLine": True,
            "multiLineSecurityNames": True,
            "maturityDateContinuation": True,
            "adjacentPercentageHandling": True,
            "fallbackUsesDateAwarePercentageClassification": True,
            "fallbackFlattensPhysicalPdfLines": True,
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
    print(
        "=" * 70
    )

    print(
        "RUN COMPLETE"
    )

    print(
        "=" * 70
    )

    print("")

    print(
        f"Excel fund universe : {len(funds)}"
    )

    print(
        f"Successful funds    : "
        f"{len(successful_results)}"
    )

    print(
        f"Failed funds        : "
        f"{len(failed_results)}"
    )

    print(
        "Total holdings      : "
        f"{run_summary['totalHoldingsExtracted']}"
    )

    print("")

    print(
        f"All holdings file   : "
        f"{ALL_HOLDINGS_FILE}"
    )

    print(
        f"Run summary file    : "
        f"{RUN_SUMMARY_FILE}"
    )

    print("")

    if failed_results:

        print(
            "FAILED ROWS:"
        )

        for failure in failed_results:

            print(
                f"  Row {failure['excelRow']}: "
                f"{failure['error']}"
            )

        print("")

        return 1

    print(
        "ALL FUNDS PASSED"
    )

    print("")

    return 0


if __name__ == "__main__":

    sys.exit(
        main()
    )
```

This version specifically preserves your hard rule that **the official published security name must remain intact**, including a `0.00%` rate, while using the later `5.8%` as the portfolio weight.

For your Row 23 example, the expected output is therefore:

```text
USA DL-Treasury Bills 2026(26) 0.00% - 5.8%
USA DL-Treasury Bills 2025(26) 0.00% - 5.8%
```

with `weightPercent` equal to `5.8`, **not `0.00`**.

Run this version against all 67 funds. If Row 23 then passes but another fund fails, send me the new `FAILED ROWS` output and we'll address that parser case without weakening the no-guessing rules.
