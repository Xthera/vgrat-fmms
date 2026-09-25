#!/usr/bin/env python3

"""
VGrat FMS - Prudential Top Holdings FAILED-FUND RECOVERY 3

PURPOSE
=======

Recover ONLY funds that still failed after Recovery 2:

    scripts/test_prudential_holdings_recovery2.py

Recovery 2 remains completely untouched.

RECOVERY SOURCE
===============

1. output_holdings_recovery_2/run_summary.json
       |
       v
   failedFundsDetail
       |
       v
2. Funds Links.xlsm
       |
       v
   authoritative Excel row + Prudential URL
       |
       v
3. Official Prudential Singapore fund page
       |
       v
4. Official Prudential factsheet PDF
       |
       v
5. Fixed-income-specific holdings extraction
       |
       v
6. Exact official-PDF re-download
       |
       v
7. Exact rank/name/weight signature verification


PURPOSE OF RECOVERY 3
=====================

Recovery 3 exists for factsheets where normal PDF extraction can expose
multiple percentages on a single logical holding line.

Typical example:

    Security Name 4.25% 3.10%

where:

    4.25% = coupon/rate contained in the security name
    3.10% = published portfolio weight

Recovery 3 therefore uses the LAST percentage on the logical holding
line as the published portfolio weight.

Earlier percentages remain part of the holding name.


HARD RULES
==========

- Recovery 2 failedFundsDetail controls the Recovery 3 universe.
- Funds Links.xlsm controls the authoritative Excel row and URL.
- No hardcoded fund rows.
- Only official Prudential Singapore URLs are accepted.
- Only official Prudential Singapore factsheets are accepted.
- No third-party holdings sources.
- No PruAccess holdings discovery.
- No fuzzy matching.
- No coordinate proximity matching.
- No fabricated holdings.
- No fabricated percentages.
- No forced ten holdings.
- Fewer than ten published holdings is valid.
- Published order must be preserved.
- Duplicate names are allowed.
- Duplicate percentages are allowed.
- Earlier percentages in fixed-income names are preserved.
- LAST percentage on a logical fixed-income line is the portfolio weight.
- Hyphenated PDF line breaks are rejoined.
- A new rank cannot appear before the previous holding has a weight.
- At least one logical holding line must contain multiple percentages.
- This multiple-percentage condition is required so that Recovery 3
  does not accidentally become a duplicate of Recovery 2.
- Exact verification downloads the official factsheet again.
- Exact verification reruns the same Recovery 3 parser.
- Exact rank/name/weight signature must match.
- Recovery 3 never modifies Recovery 2 output.
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

RECOVERY_2_RUN_SUMMARY_FILE = Path(
    "output_holdings_recovery_2/run_summary.json"
)

RECOVERY_OUTPUT_DIR = Path(
    "output_holdings_recovery_3"
)

RECOVERY_FUNDS_OUTPUT_DIR = (
    RECOVERY_OUTPUT_DIR / "funds"
)

RECOVERY_RUN_SUMMARY_FILE = (
    RECOVERY_OUTPUT_DIR / "run_summary.json"
)


# =============================================================================
# BROWSER
# =============================================================================

BROWSER_HEADLESS = True

PAGE_TIMEOUT_MS = 120000

FACTSHEET_DOWNLOAD_TIMEOUT_MS = 120000

POST_PAGE_WAIT_MS = 1500

RETRY_COUNT = 3

RETRY_DELAY_SECONDS = 3.0


# =============================================================================
# HOLDINGS
# =============================================================================

MAX_HOLDINGS = 10


# =============================================================================
# OFFICIAL PRUDENTIAL HOSTS
# =============================================================================

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


class HoldingsParseFailure(RuntimeError):

    def __init__(
        self,
        message: str,
        section_text: str = "",
        full_text: str = "",
        diagnostics: dict | None = None,
    ) -> None:

        super().__init__(
            message
        )

        self.section_text = section_text

        self.full_text = full_text

        self.diagnostics = (
            diagnostics
            or {}
        )


# =============================================================================
# LOAD RECOVERY 2 RESULTS
# =============================================================================

def load_recovery_2_run_summary() -> dict:

    if not RECOVERY_2_RUN_SUMMARY_FILE.exists():

        raise FileNotFoundError(
            "Recovery 2 run summary not found: "
            f"{RECOVERY_2_RUN_SUMMARY_FILE}"
        )

    data = json.loads(
        RECOVERY_2_RUN_SUMMARY_FILE.read_text(
            encoding="utf-8"
        )
    )

    if not isinstance(
        data,
        dict,
    ):

        raise RuntimeError(
            "Recovery 2 run summary is not a JSON object."
        )

    failed = data.get(
        "failedFundsDetail"
    )

    if not isinstance(
        failed,
        list,
    ):

        raise RuntimeError(
            "Recovery 2 run summary does not contain "
            "a valid failedFundsDetail list."
        )

    return data


# =============================================================================
# EXCEL
# =============================================================================

def read_excel_funds() -> dict[int, dict]:

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

    funds = {}

    try:

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

            funds[row_number] = {
                "excelRow":
                    row_number,

                "prudentialUrl":
                    prudential_url,

                "pruAccessName":
                    pruaccess_name,
            }

    finally:

        workbook.close()

    if not funds:

        raise RuntimeError(
            "No populated Prudential URLs found in Excel."
        )

    return funds


def build_recovery_universe(
    recovery_2: dict,
    excel_funds: dict[int, dict],
) -> list[dict]:

    failed_entries = recovery_2.get(
        "failedFundsDetail"
    )

    if not isinstance(
        failed_entries,
        list,
    ):

        raise RuntimeError(
            "Recovery 2 failedFundsDetail is invalid."
        )

    universe = []

    seen_rows = set()

    for failure in failed_entries:

        if not isinstance(
            failure,
            dict,
        ):

            raise RuntimeError(
                "A Recovery 2 failedFundsDetail entry "
                "is not an object."
            )

        if "excelRow" not in failure:

            raise RuntimeError(
                "A Recovery 2 failed entry has no excelRow."
            )

        try:

            excel_row = int(
                failure[
                    "excelRow"
                ]
            )

        except Exception as error:

            raise RuntimeError(
                "Invalid Recovery 2 excelRow: "
                f"{failure.get('excelRow')}"
            ) from error

        if excel_row in seen_rows:

            raise RuntimeError(
                f"Duplicate Recovery 2 failed excelRow: "
                f"{excel_row}"
            )

        seen_rows.add(
            excel_row
        )

        if excel_row not in excel_funds:

            raise RuntimeError(
                f"Recovery 2 failed row {excel_row} "
                "does not exist in Funds Links.xlsm."
            )

        excel_fund = excel_funds[
            excel_row
        ]

        recovery_url = clean_text(
            failure.get(
                "prudentialUrl"
            )
        )

        excel_url = clean_text(
            excel_fund[
                "prudentialUrl"
            ]
        )

        if (
            recovery_url
            and
            recovery_url
            !=
            excel_url
        ):

            raise RuntimeError(
                "Recovery 2 URL does not match Excel "
                f"for row {excel_row}."
            )

        recovery_name = clean_text(
            failure.get(
                "pruAccessName"
            )
            or
            failure.get(
                "excelPruAccessName"
            )
        )

        excel_name = clean_text(
            excel_fund.get(
                "pruAccessName"
            )
        )

        if (
            recovery_name
            and
            excel_name
            and
            recovery_name
            !=
            excel_name
        ):

            raise RuntimeError(
                "Recovery 2 PruAccess name does not match "
                f"Excel for row {excel_row}."
            )

        universe.append(
            {
                "excelRow":
                    excel_row,

                "prudentialUrl":
                    excel_url,

                "pruAccessName":
                    excel_name,

                "recovery2Failure":
                    failure,
            }
        )

    return universe


# =============================================================================
# FACTSHEET DISCOVERY
# =============================================================================

def find_factsheet_url(
    page,
) -> str:

    anchors = page.locator(
        "a"
    )

    candidates = []

    for index in range(
        anchors.count()
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

        if "fund factsheet" in text_lower:
            score += 200

        elif "factsheet" in text_lower:
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
            "Could not find an official Prudential factsheet."
        )

    candidates.sort(
        key=lambda item: (
            -item["score"],
            item["url"],
        )
    )

    return ensure_prudential_url(
        candidates[0]["url"]
    )


# =============================================================================
# PDF DOWNLOAD
# =============================================================================

def download_factsheet(
    page,
    factsheet_url: str,
) -> bytes:

    factsheet_url = ensure_prudential_url(
        factsheet_url
    )

    response = page.request.get(
        factsheet_url,
        timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
    )

    if response.status != 200:

        raise RuntimeError(
            "Factsheet HTTP status: "
            f"{response.status}"
        )

    body = response.body()

    if not body.startswith(
        b"%PDF"
    ):

        raise RuntimeError(
            "Factsheet response is not a PDF."
        )

    return body


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
            "PDF has zero pages."
        )

    pages = []

    for page_number, pdf_page in enumerate(
        reader.pages,
        start=1,
    ):

        try:

            text = (
                pdf_page.extract_text()
                or ""
            )

        except Exception as error:

            raise RuntimeError(
                f"Failed to extract PDF page {page_number}: "
                f"{error}"
            ) from error

        pages.append(
            text
        )

    full_text = "\n".join(
        pages
    )

    if not clean_text(
        full_text
    ):

        raise RuntimeError(
            "PDF contains no extractable text."
        )

    return (
        full_text,
        page_count,
    )


# =============================================================================
# HOLDINGS SECTION
# =============================================================================

def pdf_lines(
    text: str,
) -> list[str]:

    lines = []

    for raw_line in text.splitlines():

        line = clean_text(
            raw_line
        )

        if line:
            lines.append(
                line
            )

    return lines


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

    if normalized in {
        "source",
        "source:",
        "inception date",
        "important information",
        "important information:",
        "disclaimer",
        "past performance",
        "portfolio characteristics",
        "asset allocation",
    }:

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
) -> tuple[str, str]:

    lines = pdf_lines(
        text
    )

    start = find_holdings_start(
        lines
    )

    if start is None:

        return (
            "",
            "not_published",
        )

    section = []

    for line in lines[
        start + 1:
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
# HOLDING NAME HELPERS
# =============================================================================

def clean_holding_name(
    value: str,
) -> str:

    name = clean_text(
        value
    )

    if not name:
        return ""

    name = re.sub(
        r"^[•·▪■□*]+",
        "",
        name,
    )

    name = re.sub(
        r"^\d{1,2}[.)]\s+",
        "",
        name,
    )

    name = name.replace(
        "|",
        " ",
    )

    name = re.sub(
        r"\bNone\b",
        " ",
        name,
        flags=re.IGNORECASE,
    )

    name = clean_text(
        name
    )

    return name.strip(
        " -|"
    )


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


def combine_holding_name_fragments(
    fragments: list[str],
) -> str:

    cleaned = []

    for fragment in fragments:

        fragment = clean_holding_fragment(
            fragment
        )

        if not fragment:
            continue

        if re.fullmatch(
            r"\d{1,2}",
            fragment,
        ):

            continue

        cleaned.append(
            fragment
        )

    if not cleaned:
        return ""

    return clean_holding_name(
        " ".join(
            cleaned
        )
    )


# =============================================================================
# RANK / PERCENTAGE
# =============================================================================

def percentage_matches(
    line: str,
) -> list[re.Match]:

    return list(
        re.finditer(
            r"(?<![\d.])"
            r"([0-9]+(?:\.[0-9]+)?)"
            r"\s*%",
            line,
        )
    )


def extract_leading_rank(
    line: str,
) -> tuple[int | None, str]:

    text = clean_text(
        line
    )

    match = re.match(
        r"^\s*(\d{1,2})[.)\-:]\s*(.*)$",
        text,
    )

    if match:

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

    match = re.match(
        r"^\s*(\d{1,2})\s+(.+)$",
        text,
    )

    if match:

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

    return (
        None,
        text,
    )


def find_last_percentage_in_line(
    line: str,
) -> tuple[float, str, int, int] | None:

    matches = percentage_matches(
        line
    )

    if not matches:
        return None

    match = matches[-1]

    value = float(
        match.group(1)
    )

    if not 0 <= value <= 100:
        return None

    return (
        value,
        clean_text(
            match.group(0)
        ),
        match.start(),
        match.end(),
    )


# =============================================================================
# HYPHENATED LINE REJOIN
# =============================================================================

def merge_hyphenated_line_breaks(
    lines: list[str],
) -> list[str]:

    merged = []

    index = 0

    while index < len(lines):

        current = clean_text(
            lines[index]
        )

        while (
            current.endswith("-")
            and
            index + 1 < len(lines)
        ):

            index += 1

            next_line = clean_text(
                lines[index]
            )

            current = (
                current
                +
                next_line
            )

        merged.append(
            current
        )

        index += 1

    return merged


def build_logical_holding_lines(
    section_text: str,
) -> list[str]:

    return merge_hyphenated_line_breaks(
        pdf_lines(
            section_text
        )
    )


# =============================================================================
# VALIDATION
# =============================================================================

def validate_holdings(
    holdings: list[dict],
) -> list[dict]:

    if not holdings:

        raise RuntimeError(
            "Recovery 3 parser produced no holdings."
        )

    if len(holdings) > MAX_HOLDINGS:

        raise RuntimeError(
            "Recovery 3 parser produced more than "
            f"{MAX_HOLDINGS} holdings."
        )

    expected_ranks = list(
        range(
            1,
            len(holdings) + 1,
        )
    )

    actual_ranks = [
        holding[
            "rank"
        ]
        for holding
        in holdings
    ]

    if actual_ranks != expected_ranks:

        raise RuntimeError(
            "Invalid sequential ranks. "
            f"Actual={actual_ranks}; "
            f"Expected={expected_ranks}"
        )

    for holding in holdings:

        name = clean_holding_name(
            holding[
                "name"
            ]
        )

        if not name:

            raise RuntimeError(
                "Recovery 3 parser produced an empty holding name."
            )

        weight = holding[
            "weightPercent"
        ]

        if not isinstance(
            weight,
            (int, float),
        ):

            raise RuntimeError(
                "Recovery 3 parser produced an invalid weight."
            )

        if not 0 <= weight <= 100:

            raise RuntimeError(
                "Recovery 3 parser produced a weight "
                "outside 0-100."
            )

        holding[
            "name"
        ] = name

    return holdings


# =============================================================================
# RECOVERY 3 FIXED-INCOME PARSER
# =============================================================================

def parse_holdings_fixed_income_recovery(
    section_text: str,
) -> list[dict]:

    """
    Fixed-income-specific Recovery 3 parser.

    IMPORTANT:

    A logical line may contain:

        SECURITY NAME 4.25% 3.10%

    The earlier percentage may belong to the security name
    (coupon/rate), while the LAST percentage is the published
    portfolio weight.

    Therefore:

        name       = everything before the LAST %
        weight     = LAST %

    The parser also requires at least one logical line containing
    multiple percentages. This prevents Recovery 3 from simply
    duplicating Recovery 2's normal fallback parser.
    """

    lines = build_logical_holding_lines(
        section_text
    )

    if not lines:

        raise RuntimeError(
            "Fixed-income Recovery 3 received an empty section."
        )

    holdings = []

    pending_fragments = []

    pending_rank = None

    multiple_percentage_lines = 0

    logical_lines_with_weights = 0

    for raw_line in lines:

        line = clean_text(
            raw_line
        )

        if not line:
            continue

        normalized = normalize_text(
            line
        )

        if normalized in {
            "holding",
            "holdings",
            "name",
            "names",
            "weight",
            "weights",
            "allocation",
            "allocations",
            "%",
        }:

            continue

        if (
            normalized.startswith(
                "top 10 holdings"
            )
            or
            normalized.startswith(
                "top ten holdings"
            )
        ):

            continue

        detected_rank, remainder = (
            extract_leading_rank(
                line
            )
        )

        if detected_rank is not None:

            if pending_fragments:

                raise RuntimeError(
                    "A new rank appeared before the previous "
                    "holding received its weight."
                )

            pending_rank = detected_rank

            line = remainder

            if not line:
                continue

        matches = percentage_matches(
            line
        )

        if len(matches) > 1:

            multiple_percentage_lines += 1

        percentage_info = (
            find_last_percentage_in_line(
                line
            )
        )

        if percentage_info:

            (
                weight,
                weight_text,
                start,
                _end,
            ) = percentage_info

            logical_lines_with_weights += 1

            name_fragment = clean_text(
                line[:start]
            )

            if name_fragment:

                pending_fragments.append(
                    name_fragment
                )

            name = combine_holding_name_fragments(
                pending_fragments
            )

            if not name:

                raise RuntimeError(
                    "A percentage was found without "
                    "a holding name."
                )

            rank = (
                pending_rank
                if pending_rank is not None
                else len(holdings) + 1
            )

            if rank != len(holdings) + 1:

                raise RuntimeError(
                    "Unexpected holding rank. "
                    f"Expected={len(holdings) + 1}; "
                    f"Found={rank}"
                )

            holdings.append(
                {
                    "rank":
                        rank,

                    "name":
                        name,

                    "weightPercent":
                        weight,

                    "weightText":
                        weight_text,

                    "percentageCountOnLogicalLine":
                        len(matches),

                    "recoveryParser":
                        "fixed_income_table",
                }
            )

            pending_fragments = []

            pending_rank = None

            if len(holdings) >= MAX_HOLDINGS:

                break

            continue

        fragment = clean_holding_fragment(
            line
        )

        if fragment:

            pending_fragments.append(
                fragment
            )

    if multiple_percentage_lines <= 0:

        raise RuntimeError(
            "Recovery 3 fixed-income parser did not find "
            "any logical holding line containing multiple "
            "percentages. Recovery 3 evidence condition failed."
        )

    if logical_lines_with_weights != len(
        holdings
    ):

        raise RuntimeError(
            "Recovery 3 logical holding/weight count mismatch."
        )

    return validate_holdings(
        holdings
    )


# =============================================================================
# RECOVERY 3 EXTRACTION ENGINE
# =============================================================================

def extract_fixed_income_recovery_engine(
    pdf_bytes: bytes,
) -> dict:

    (
        full_text,
        page_count,
    ) = extract_pdf_text(
        pdf_bytes
    )

    (
        section_text,
        section_status,
    ) = extract_holdings_section(
        full_text
    )

    if section_status == "not_published":

        raise HoldingsParseFailure(
            "No published Top 10 Holdings section found.",
            section_text="",
            full_text=full_text,
            diagnostics={
                "status":
                    "no_holdings_section",
            },
        )

    try:

        holdings = (
            parse_holdings_fixed_income_recovery(
                section_text
            )
        )

    except Exception as error:

        raise HoldingsParseFailure(
            "Recovery 3 fixed-income parser failed: "
            f"{error}",
            section_text=section_text,
            full_text=full_text,
            diagnostics={
                "parser":
                    "fixed_income_table",

                "error":
                    clean_text(
                        str(error)
                    ),
            },
        ) from error

    return {
        "status":
            "success",

        "holdings":
            holdings,

        "parser":
            "fixed_income_table",

        "fullText":
            full_text,

        "sectionText":
            section_text,

        "pageCount":
            page_count,
    }


# =============================================================================
# FUND PAGE NAME
# =============================================================================

def extract_fund_page_name(
    page,
) -> str:

    try:

        h1 = page.locator(
            "h1"
        ).first

        if h1.count():

            name = clean_text(
                h1.inner_text()
            )

            if name:
                return name

    except Exception:

        pass

    try:

        body = clean_text(
            page.locator(
                "body"
            ).inner_text()
        )

        match = re.search(
            r"\bPRU(?:Link|Prime)\s+[^\n]+",
            body,
            re.IGNORECASE,
        )

        if match:

            return clean_text(
                match.group(0)
            )

    except Exception:

        pass

    return ""


# =============================================================================
# SINGLE FUND RECOVERY
# =============================================================================

def recover_single_fund(
    page,
    excel_fund: dict,
) -> dict:

    excel_row = int(
        excel_fund[
            "excelRow"
        ]
    )

    prudential_url = ensure_prudential_url(
        excel_fund[
            "prudentialUrl"
        ]
    )

    print(
        "\n"
        + "=" * 72
    )

    print(
        f"RECOVERY 3 - FUND ROW {excel_row}"
    )

    print(
        "=" * 72
    )

    print(
        f"URL: {prudential_url}"
    )

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

    if not is_prudential_url(
        final_url
    ):

        raise RuntimeError(
            "Fund page redirected outside Prudential Singapore: "
            f"{final_url}"
        )

    page_title = clean_text(
        page.title()
    )

    fund_name = extract_fund_page_name(
        page
    )

    factsheet_url = find_factsheet_url(
        page
    )

    print(
        f"Fund: {fund_name or '-'}"
    )

    print(
        f"Factsheet: {factsheet_url}"
    )

    factsheet_bytes = download_factsheet(
        page,
        factsheet_url,
    )

    engine = extract_fixed_income_recovery_engine(
        factsheet_bytes
    )

    return {
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
            excel_fund.get(
                "pruAccessName"
            ),

        "factsheetUrl":
            factsheet_url,

        "factsheetPageCount":
            engine[
                "pageCount"
            ],

        "topHoldingsCount":
            len(
                engine[
                    "holdings"
                ]
            ),

        "topHoldings":
            engine[
                "holdings"
            ],

        "holdingsParser":
            engine[
                "parser"
            ],

        "recoveryStage":
            "recovery3",

        "recoveryEngine":
            "fixed_income_table",

        "_factsheetBytes":
            factsheet_bytes,

        "_fullText":
            engine[
                "fullText"
            ],

        "_sectionText":
            engine[
                "sectionText"
            ],
    }


# =============================================================================
# EXACT VERIFICATION
# =============================================================================

def holding_signature(
    holdings: list[dict],
) -> list[tuple]:

    return [
        (
            int(
                holding[
                    "rank"
                ]
            ),

            normalize_text(
                holding[
                    "name"
                ]
            ),

            float(
                holding[
                    "weightPercent"
                ]
            ),
        )

        for holding
        in holdings
    ]


def verify_exact_against_official_pdf(
    page,
    result: dict,
) -> dict:

    print(
        "FINAL VERIFICATION: "
        "downloading official PDF again..."
    )

    factsheet_url = ensure_prudential_url(
        result[
            "factsheetUrl"
        ]
    )

    response = page.request.get(
        factsheet_url,
        timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
    )

    if response.status != 200:

        raise RuntimeError(
            "Final verification returned HTTP "
            f"{response.status}."
        )

    verification_bytes = response.body()

    if not verification_bytes.startswith(
        b"%PDF"
    ):

        raise RuntimeError(
            "Final verification response was not a PDF."
        )

    verification_engine = (
        extract_fixed_income_recovery_engine(
            verification_bytes
        )
    )

    original_holdings = result[
        "topHoldings"
    ]

    verified_holdings = (
        verification_engine[
            "holdings"
        ]
    )

    original_signature = holding_signature(
        original_holdings
    )

    verified_signature = holding_signature(
        verified_holdings
    )

    if (
        original_signature
        !=
        verified_signature
    ):

        raise HoldingsParseFailure(
            "EXACT VERIFICATION FAILED: "
            "re-downloaded official Prudential PDF "
            "produced a different rank/name/weight signature.",
            section_text=verification_engine[
                "sectionText"
            ],
            full_text=verification_engine[
                "fullText"
            ],
            diagnostics={
                "originalSignature":
                    original_signature,

                "verifiedSignature":
                    verified_signature,

                "originalHoldings":
                    original_holdings,

                "verifiedHoldings":
                    verified_holdings,
            },
        )

    return {
        "verified":
            True,

        "verifiedParser":
            verification_engine[
                "parser"
            ],

        "verifiedHoldingCount":
            len(
                verified_holdings
            ),

        "verifiedSignature":
            verified_signature,

        "verificationFactsheetBytes":
            verification_bytes,

        "verificationFullText":
            verification_engine[
                "fullText"
            ],

        "verificationSectionText":
            verification_engine[
                "sectionText"
            ],
    }


# =============================================================================
# OUTPUT
# =============================================================================

def recovery_directory(
    result: dict,
) -> Path:

    excel_row = int(
        result[
            "excelRow"
        ]
    )

    identifier = safe_filename(
        result.get(
            "fundName"
        )
        or
        f"fund_{excel_row}"
    )

    return (
        RECOVERY_FUNDS_OUTPUT_DIR
        /
        f"{excel_row}_{identifier}"
    )


def save_success(
    result: dict,
    verification: dict,
) -> Path:

    directory = recovery_directory(
        result
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    (
        directory
        /
        "factsheet.pdf"
    ).write_bytes(
        verification[
            "verificationFactsheetBytes"
        ]
    )

    (
        directory
        /
        "factsheet_text.txt"
    ).write_text(
        verification[
            "verificationFullText"
        ],
        encoding="utf-8",
    )

    (
        directory
        /
        "top_holdings_section.txt"
    ).write_text(
        verification[
            "verificationSectionText"
        ],
        encoding="utf-8",
    )

    clean_result = {
        key:
            value
        for key, value
        in result.items()
        if not key.startswith("_")
    }

    clean_result[
        "verification"
    ] = {
        "verified":
            True,

        "verifiedParser":
            verification[
                "verifiedParser"
            ],

        "verifiedHoldingCount":
            verification[
                "verifiedHoldingCount"
            ],

        "exactSignatureMatch":
            True,

        "verifiedAtUtc":
            utc_now_iso(),
    }

    save_json(
        directory
        /
        "top_holdings.json",
        clean_result,
    )

    save_json(
        directory
        /
        "recovery_diagnostics.json",
        {
            "recoveryStage":
                "recovery3",

            "recoveryEngine":
                "fixed_income_table",

            "holdingsParser":
                "fixed_income_table",

            "verificationParser":
                verification[
                    "verifiedParser"
                ],

            "exactSignatureMatch":
                True,

            "savedAtUtc":
                utc_now_iso(),
        },
    )

    save_json(
        directory
        /
        "metadata.json",
        {
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

            "recoveryStage":
                "recovery3",

            "recoveryEngine":
                "fixed_income_table",

            "topHoldingsCount":
                result.get(
                    "topHoldingsCount"
                ),

            "exactVerification":
                True,

            "savedAtUtc":
                utc_now_iso(),
        },
    )

    return directory


def save_failure(
    excel_fund: dict,
    error_text: str,
    attempts: list[dict],
    section_text: str | None = None,
    full_text: str | None = None,
    diagnostics: dict | None = None,
) -> Path:

    excel_row = int(
        excel_fund[
            "excelRow"
        ]
    )

    directory = (
        RECOVERY_FUNDS_OUTPUT_DIR
        /
        f"{excel_row}_failed"
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    save_json(
        directory
        /
        "recovery3_failure.json",
        {
            "status":
                "failed",

            "recoveryStage":
                "recovery3",

            "excelRow":
                excel_row,

            "prudentialUrl":
                excel_fund[
                    "prudentialUrl"
                ],

            "pruAccessName":
                excel_fund.get(
                    "pruAccessName"
                ),

            "error":
                clean_text(
                    error_text
                ),

            "attempts":
                attempts,

            "failedAtUtc":
                utc_now_iso(),
        },
    )

    if section_text:

        (
            directory
            /
            "recovery3_top_holdings_section.txt"
        ).write_text(
            section_text,
            encoding="utf-8",
        )

    if full_text:

        (
            directory
            /
            "recovery3_factsheet_text.txt"
        ).write_text(
            full_text,
            encoding="utf-8",
        )

    if diagnostics:

        save_json(
            directory
            /
            "recovery3_diagnostics.json",
            diagnostics,
        )

    return directory


# =============================================================================
# MAIN
# =============================================================================

def main() -> int:

    RECOVERY_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    RECOVERY_FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    started_at = utc_now_iso()

    print(
        "\n"
        + "#" * 78
    )

    print(
        "VGRAT FMS - PRUDENTIAL TOP HOLDINGS RECOVERY 3"
    )

    print(
        "#" * 78
    )

    print(
        f"Started UTC: {started_at}"
    )

    # -------------------------------------------------------------------------
    # Recovery 2 source
    # -------------------------------------------------------------------------

    recovery_2 = load_recovery_2_run_summary()

    print(
        "\nRECOVERY 2 SOURCE"
    )

    print(
        f"Recovery 2 status: "
        f"{recovery_2.get('status')}"
    )

    print(
        f"Recovery 2 recovery universe: "
        f"{recovery_2.get('recoveryUniverse')}"
    )

    print(
        f"Recovery 2 failed funds: "
        f"{recovery_2.get('failedFunds')}"
    )

    # -------------------------------------------------------------------------
    # Excel
    # -------------------------------------------------------------------------

    excel_funds = read_excel_funds()

    recovery_universe = build_recovery_universe(
        recovery_2,
        excel_funds,
    )

    print(
        f"\nRecovery 3 universe: "
        f"{len(recovery_universe)}"
    )

    # -------------------------------------------------------------------------
    # Nothing to recover
    # -------------------------------------------------------------------------

    if not recovery_universe:

        completed_at = utc_now_iso()

        summary = {
            "status":
                "nothing_to_recover",

            "startedAtUtc":
                started_at,

            "completedAtUtc":
                completed_at,

            "recovery2RunSummary":
                str(
                    RECOVERY_2_RUN_SUMMARY_FILE
                ),

            "excelFile":
                str(
                    EXCEL_FILE
                ),

            "recovery2FailedFunds":
                recovery_2.get(
                    "failedFunds",
                    0,
                ),

            "recovery3Universe":
                0,

            "attemptedFunds":
                0,

            "recoveredFunds":
                0,

            "failedFunds":
                0,

            "rules": {
                "recovery2FailedFundsOnly":
                    True,

                "hardcodedRows":
                    False,

                "officialPrudentialOnly":
                    True,

                "fixedIncomeParserOnly":
                    True,

                "lastPercentageIsWeight":
                    True,

                "multiplePercentageEvidenceRequired":
                    True,

                "exactFinalVerification":
                    True,
            },
        }

        save_json(
            RECOVERY_RUN_SUMMARY_FILE,
            summary,
        )

        print(
            "\nNo Recovery 2 failures remain."
        )

        return 0

    recovered = []

    failed = []

    attempts_total = 0

    # -------------------------------------------------------------------------
    # Process Recovery 2 failures
    # -------------------------------------------------------------------------

    for index, excel_fund in enumerate(
        recovery_universe,
        start=1,
    ):

        excel_row = int(
            excel_fund[
                "excelRow"
            ]
        )

        print(
            "\n"
            + "=" * 78
        )

        print(
            f"RECOVERY 3 FUND {index}/{len(recovery_universe)}"
        )

        print(
            f"Excel row: {excel_row}"
        )

        print(
            "=" * 78
        )

        attempts = []

        final_result = None

        last_exception = None

        for attempt in range(
            1,
            RETRY_COUNT + 1,
        ):

            attempts_total += 1

            attempt_record = {
                "attempt":
                    attempt,

                "startedAtUtc":
                    utc_now_iso(),
            }

            print(
                f"\nAttempt {attempt}/{RETRY_COUNT}"
            )

            try:

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

                        result = recover_single_fund(
                            page,
                            excel_fund,
                        )

                        verification = (
                            verify_exact_against_official_pdf(
                                page,
                                result,
                            )
                        )

                        directory = save_success(
                            result,
                            verification,
                        )

                        result[
                            "outputDirectory"
                        ] = str(
                            directory
                        )

                        result[
                            "exactVerification"
                        ] = True

                        result[
                            "verificationParser"
                        ] = verification[
                            "verifiedParser"
                        ]

                        final_result = result

                    finally:

                        context.close()

                        browser.close()

                attempt_record[
                    "status"
                ] = "success"

                attempt_record[
                    "completedAtUtc"
                ] = utc_now_iso()

                attempt_record[
                    "parser"
                ] = (
                    final_result.get(
                        "holdingsParser"
                    )
                    if final_result
                    else
                    None
                )

                attempts.append(
                    attempt_record
                )

                last_exception = None

                break

            except Exception as error:

                last_exception = error

                attempt_record[
                    "status"
                ] = "failed"

                attempt_record[
                    "error"
                ] = clean_text(
                    str(error)
                )

                attempt_record[
                    "completedAtUtc"
                ] = utc_now_iso()

                if isinstance(
                    error,
                    HoldingsParseFailure,
                ):

                    attempt_record[
                        "diagnostics"
                    ] = error.diagnostics

                attempts.append(
                    attempt_record
                )

                print(
                    f"Attempt failed: "
                    f"{error}"
                )

                if attempt < RETRY_COUNT:

                    print(
                        f"Retrying in "
                        f"{RETRY_DELAY_SECONDS} seconds..."
                    )

                    time.sleep(
                        RETRY_DELAY_SECONDS
                    )

        # ---------------------------------------------------------------------
        # Failed
        # ---------------------------------------------------------------------

        if final_result is None:

            section_text = getattr(
                last_exception,
                "section_text",
                None,
            )

            full_text = getattr(
                last_exception,
                "full_text",
                None,
            )

            diagnostics = getattr(
                last_exception,
                "diagnostics",
                None,
            )

            directory = save_failure(
                excel_fund,
                clean_text(
                    str(
                        last_exception
                    )
                    if last_exception
                    else
                    "Unknown Recovery 3 failure."
                ),
                attempts,
                section_text=section_text,
                full_text=full_text,
                diagnostics=diagnostics,
            )

            failed.append(
                {
                    "excelRow":
                        excel_row,

                    "prudentialUrl":
                        excel_fund[
                            "prudentialUrl"
                        ],

                    "pruAccessName":
                        excel_fund.get(
                            "pruAccessName"
                        ),

                    "error":
                        clean_text(
                            str(
                                last_exception
                            )
                            if last_exception
                            else
                            "Unknown Recovery 3 failure."
                        ),

                    "outputDirectory":
                        str(
                            directory
                        ),

                    "attempts":
                        attempts,
                }
            )

            continue

        # ---------------------------------------------------------------------
        # Recovered
        # ---------------------------------------------------------------------

        recovered.append(
            final_result
        )

        print(
            "\nRECOVERY 3 SUCCESS"
        )

        print(
            f"Fund: "
            f"{final_result.get('fundName') or '-'}"
        )

        print(
            "Parser: fixed_income_table"
        )

        print(
            f"Holdings: "
            f"{final_result.get('topHoldingsCount')}"
        )

        print(
            "Exact final verification: PASS"
        )

        for holding in final_result[
            "topHoldings"
        ]:

            print(
                f"  {holding['rank']}. "
                f"{holding['name']} - "
                f"{holding['weightText']}"
            )

    # =========================================================================
    # SUMMARY
    # =========================================================================

    completed_at = utc_now_iso()

    total_holdings = sum(
        int(
            result.get(
                "topHoldingsCount",
                0,
            )
            or
            0
        )
        for result
        in recovered
    )

    summary = {
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

        "recovery2RunSummary":
            str(
                RECOVERY_2_RUN_SUMMARY_FILE
            ),

        "excelFile":
            str(
                EXCEL_FILE
            ),

        "recovery2FailedFunds":
            recovery_2.get(
                "failedFunds"
            ),

        "recovery3Universe":
            len(
                recovery_universe
            ),

        "attemptedFunds":
            len(
                recovery_universe
            ),

        "recovery3AttemptCount":
            attempts_total,

        "recoveredFunds":
            len(
                recovered
            ),

        "failedFunds":
            len(
                failed
            ),

        "totalPublishedTopHoldingsRecovered":
            total_holdings,

        "recoveredFundsDetail":
            [
                {
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

                    "topHoldingsCount":
                        result.get(
                            "topHoldingsCount"
                        ),

                    "holdingsParser":
                        result.get(
                            "holdingsParser"
                        ),

                    "recoveryStage":
                        result.get(
                            "recoveryStage"
                        ),

                    "recoveryEngine":
                        result.get(
                            "recoveryEngine"
                        ),

                    "exactVerification":
                        result.get(
                            "exactVerification"
                        ),

                    "verificationParser":
                        result.get(
                            "verificationParser"
                        ),

                    "outputDirectory":
                        result.get(
                            "outputDirectory"
                        ),
                }

                for result
                in recovered
            ],

        "failedFundsDetail":
            failed,

        "rules": {
            "recovery2FailedFundsOnly":
                True,

            "hardcodedRows":
                False,

            "officialPrudentialSingaporeOnly":
                True,

            "officialFactsheetOnly":
                True,

            "thirdPartyHoldings":
                False,

            "pruAccessHoldingsDiscovery":
                False,

            "fuzzyMatching":
                False,

            "coordinateProximityPairing":
                False,

            "inferredHoldings":
                False,

            "fabricatedPercentages":
                False,

            "forcedTenEntries":
                False,

            "fewerThanTenPublishedHoldingsAllowed":
                True,

            "publishedOrderPreserved":
                True,

            "hyphenatedLineWrapReconstruction":
                True,

            "fixedIncomeRecovery":
                True,

            "lastPercentageIsPortfolioWeight":
                True,

            "earlierPercentagesRemainInHoldingName":
                True,

            "multiplePercentageEvidenceRequired":
                True,

            "exactFinalPdfVerification":
                True,

            "exactRankNameWeightSignatureRequired":
                True,

            "recovery2Modified":
                False,
        },
    }

    save_json(
        RECOVERY_RUN_SUMMARY_FILE,
        summary,
    )

    # =========================================================================
    # CONSOLE SUMMARY
    # =========================================================================

    print(
        "\n\n"
        + "=" * 78
    )

    print(
        "PRUDENTIAL TOP HOLDINGS RECOVERY 3 COMPLETE"
    )

    print(
        "=" * 78
    )

    print(
        f"Recovery 2 failed funds: "
        f"{recovery_2.get('failedFunds')}"
    )

    print(
        f"Recovery 3 universe: "
        f"{len(recovery_universe)}"
    )

    print(
        f"Recovered: "
        f"{len(recovered)}"
    )

    print(
        f"Still failed: "
        f"{len(failed)}"
    )

    print(
        f"Published holdings recovered: "
        f"{total_holdings}"
    )

    print(
        "\nRecovery 3 parser:"
    )

    print(
        "  fixed_income_table"
    )

    print(
        "  LAST percentage = portfolio weight"
    )

    print(
        "  Earlier percentages remain in holding name"
    )

    print(
        "  Multiple-percentage evidence required"
    )

    print(
        "  Exact second PDF verification required"
    )

    print(
        "\nRecovery 2 was not modified."
    )

    print(
        "\nOutput:"
    )

    print(
        f" - {RECOVERY_RUN_SUMMARY_FILE}"
    )

    print(
        f" - {RECOVERY_FUNDS_OUTPUT_DIR}"
    )

    if failed:

        print(
            "\nSTILL FAILED AFTER RECOVERY 3:"
        )

        for item in failed:

            print(
                f" - Row "
                f"{item['excelRow']}: "
                f"{item['error']}"
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

    except Exception as error:

        print(
            "\nFATAL RECOVERY 3 ERROR:",
            file=sys.stderr,
        )

        print(
            str(error),
            file=sys.stderr,
        )

        raise SystemExit(
            1
        )
